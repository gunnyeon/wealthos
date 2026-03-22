// ============================================================
// WEALTHOS Family — Google Apps Script 백엔드
// 가족 공유 실시간 동기화 버전
// ============================================================

const CFG = {
  VERSION: '2.0.0-family',
  SHEETS: {
    LEDGER:    '가계부',
    ASSETS:    '자산',
    GOALS:     '목표',
    SETTINGS:  '설정',
    MEMBERS:   '가족구성원',
    ACTIVITY:  '활동로그',
  },
  EXPENSE_CATS: [
    '식비','카페','외식','교통','주유','쇼핑',
    '구독','통신','의료','문화','교육','경조사','보험','기타',
  ],
  INCOME_CATS: ['급여','부업','투자수익','임대','기타수입'],
  PAY_METHODS: ['체크카드','신용카드','현금','계좌이체','간편결제'],
  ASSET_TYPES: ['예금/현금','주식/ETF','부동산','자동차','적금/보험','기타'],
  KEYWORDS: {
    '식비':  'GS25,CU,세븐,이마트,홈플러스,롯데마트,마트,슈퍼,편의점',
    '카페':  '스타벅스,카페,커피,메가커피,투썸,빽다방,할리스,이디야',
    '외식':  '배민,쿠팡이츠,요기요,맥도날드,버거킹,피자,치킨,음식점,식당,라멘',
    '교통':  'T-money,지하철,버스,택시,카카오택시,티머니,KTX,기차',
    '주유':  '주유소,GS칼텍스,SK에너지,현대오일',
    '쇼핑':  '쿠팡,네이버,11번가,올리브영,무신사,G마켓,옥션,위메프',
    '구독':  '넷플릭스,유튜브프리미엄,멜론,스포티파이,왓챠,ChatGPT,Adobe',
    '통신':  'SKT,KT,LG유플러스,알뜰폰',
    '의료':  '병원,약국,의원,내과,치과,한의원,약국',
    '급여':  '급여,월급,임금,봉급',
    '부업':  '프리랜서,강의료,외주,용돈',
  },
};

// ============================================================
// 웹 앱 진입점
// ============================================================
function doGet(e) {
  const output = d => ContentService
    .createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON);

  try {
    const action = e.parameter.action || '';
    const since  = e.parameter.since  || '';   // ISO 타임스탬프 (폴링용)

    switch (action) {
      case 'getDashboard': return output(_getDashboard());
      case 'getLedger':    return output(_getLedger(e.parameter));
      case 'getAssets':    return output(_getAssets());
      case 'getGoals':     return output(_getGoals());
      case 'getMembers':   return output(_getMembers());
      case 'getActivity':  return output(_getActivity());
      case 'poll':         return output(_poll(since));   // ← 실시간 폴링 핵심
      case 'getConfig':    return output(_getConfig());
      default:             return output({ ok: true, version: CFG.VERSION });
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const output = d => ContentService
    .createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON);

  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || '';

    switch (action) {
      case 'addTransaction':  return output(_addTransaction(body));
      case 'deleteTransaction': return output(_deleteTransaction(body));
      case 'addAsset':        return output(_addAsset(body));
      case 'updateAsset':     return output(_updateAsset(body));
      case 'addGoal':         return output(_addGoal(body));
      case 'updateGoal':      return output(_updateGoal(body));
      case 'addMember':       return output(_addMember(body));
      case 'saveSettings':    return output(_saveSettings(body));
      default: return output({ ok: false, error: '알 수 없는 액션: ' + action });
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// 핵심: 폴링 API — 특정 시각 이후 변경된 데이터만 반환
// ============================================================
function _poll(since) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sinceDate = since ? new Date(since) : new Date(0);
  const now = new Date();

  // 가계부: since 이후 추가된 거래
  const newTxns = _getLedgerSince(ss, sinceDate);

  // 활동 로그: since 이후 활동
  const newActivity = _getActivitySince(ss, sinceDate);

  // 집계 항상 포함 (가벼운 연산)
  const summary = _getLedgerMonth(ss);

  return {
    ok: true,
    timestamp: now.toISOString(),
    hasChanges: newTxns.length > 0 || newActivity.length > 0,
    newTransactions: newTxns,
    newActivity,
    summary,    // 이번 달 수입/지출 요약 (항상 최신)
  };
}

function _getLedgerSince(ss, sinceDate) {
  const sheet = ss.getSheetByName(CFG.SHEETS.LEDGER);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  return rows
    .filter(r => r[0] && r[9] && new Date(r[9]) > sinceDate)
    .map(r => _rowToTxn(r))
    .reverse();
}

function _getActivitySince(ss, sinceDate) {
  const sheet = ss.getSheetByName(CFG.SHEETS.ACTIVITY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, Math.min(sheet.getLastRow() - 1, 50), 4).getValues();
  return rows
    .filter(r => r[0] && new Date(r[0]) > sinceDate)
    .map(r => ({ time: r[0], member: r[1], action: r[2], detail: r[3] }))
    .reverse()
    .slice(0, 20);
}

// ============================================================
// 대시보드 데이터
// ============================================================
function _getDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = _getLedgerMonth(ss);
  const totalAssets = _sumCol(ss, CFG.SHEETS.ASSETS, 4);
  const totalDebts  = Number(_getSettingVal(ss, '총부채') || 0);
  const goals = (_getGoals().data || []);
  const members = (_getMembers().data || []);
  const activity = _getActivity().data || [];

  // 구성원별 이번 달 지출
  const memberExpense = {};
  members.forEach(m => { memberExpense[m.name] = 0; });
  const ledger = ss.getSheetByName(CFG.SHEETS.LEDGER);
  if (ledger && ledger.getLastRow() > 1) {
    const now = new Date();
    const rows = ledger.getRange(2, 1, ledger.getLastRow() - 1, 10).getValues();
    rows.forEach(r => {
      if (!r[0]) return;
      const d = new Date(r[0]);
      if (d.getFullYear() !== now.getFullYear() || d.getMonth() !== now.getMonth()) return;
      if (r[1] === '지출') {
        const who = r[7] || '미지정';
        memberExpense[who] = (memberExpense[who] || 0) + (Number(r[4]) || 0);
      }
    });
  }

  return {
    ok: true,
    data: {
      totalAssets, totalDebts,
      netWorth: totalAssets - totalDebts,
      monthlyIncome:  summary.income,
      monthlyExpense: summary.expense,
      savingsRate: summary.income > 0
        ? Math.round((summary.income - summary.expense) / summary.income * 100) : 0,
      categories:    summary.categories,
      budgets:       _getBudgets(ss),
      goals,
      members,
      memberExpense,
      recentActivity: activity.slice(0, 10),
      updatedAt: new Date().toISOString(),
    },
  };
}

// ============================================================
// 가계부
// ============================================================
function _getLedger(params) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.LEDGER);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };

  const rows       = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  const filterMonth  = params.month  || '';
  const filterMember = params.member || '';

  const result = rows
    .filter(r => {
      if (!r[0] || !r[4]) return false;
      const d  = new Date(r[0]);
      const ym = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
      if (filterMonth  && ym   !== filterMonth)  return false;
      if (filterMember && r[7] !== filterMember) return false;
      return true;
    })
    .map(r => _rowToTxn(r))
    .reverse();

  return { ok: true, data: result };
}

function _rowToTxn(r) {
  const d = new Date(r[0]);
  return {
    id:        r[8] || '',
    date:      `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`,
    type:      r[1],
    category:  r[2],
    desc:      r[3],
    amount:    Number(r[4]) || 0,
    payMethod: r[5],
    memo:      r[6],
    member:    r[7],
    createdAt: r[9] ? new Date(r[9]).toISOString() : '',
  };
}

// ============================================================
// 거래 추가 (작성자 기록 포함)
// ============================================================
function _addTransaction(body) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.LEDGER);
  if (!sheet) return { ok: false, error: '가계부 시트 없음' };

  const { date, type, category, desc, amount, payMethod, memo, member } = body;
  if (!date || !type || !amount) return { ok: false, error: '필수 항목 누락' };

  let finalCat = category || _autoClassify(desc || '');
  if (!finalCat) finalCat = type === '수입' ? '기타수입' : '기타';

  const id        = Utilities.getUuid();
  const createdAt = new Date().toISOString();
  const lastRow   = Math.max(sheet.getLastRow(), 1);

  // 컬럼: 날짜|유형|카테고리|내용|금액|결제|메모|작성자|ID|생성시각
  sheet.getRange(lastRow + 1, 1, 1, 10).setValues([[
    date, type, finalCat, desc || '',
    Number(amount), payMethod || '', memo || '',
    member || '미지정', id, createdAt,
  ]]);
  sheet.getRange(lastRow + 1, 5).setNumberFormat('#,##0');

  // 활동 로그 기록
  _logActivity(ss, member, '거래추가', `${type} ${finalCat} ₩${Number(amount).toLocaleString()}`);

  return { ok: true, id, category: finalCat, createdAt };
}

// ============================================================
// 거래 삭제
// ============================================================
function _deleteTransaction(body) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.LEDGER);
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, error: '데이터 없음' };

  const rows = sheet.getRange(2, 9, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === body.id) {
      sheet.deleteRow(i + 2);
      _logActivity(ss, body.member, '거래삭제', body.desc || body.id.slice(0, 8));
      return { ok: true };
    }
  }
  return { ok: false, error: '거래를 찾을 수 없음' };
}

// ============================================================
// 자산
// ============================================================
function _getAssets() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.ASSETS);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return {
    ok: true,
    data: rows.filter(r => r[0]).map(r => ({
      id: r[6] || '', name: r[0], type: r[1], institution: r[2],
      currentValue: Number(r[3]) || 0, purchaseValue: Number(r[4]) || 0, memo: r[5],
    })),
  };
}

function _addAsset(body) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.ASSETS);
  if (!sheet) return { ok: false, error: '자산 시트 없음' };
  if (!body.name || !body.currentValue) return { ok: false, error: '필수 항목 누락' };
  const id = Utilities.getUuid();
  const lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(lastRow + 1, 1, 1, 7).setValues([[
    body.name, body.type || '기타', body.institution || '',
    Number(body.currentValue), Number(body.purchaseValue || body.currentValue),
    body.memo || '', id,
  ]]);
  sheet.getRange(lastRow + 1, 4, 1, 2).setNumberFormat('#,##0');
  _logActivity(ss, body.member, '자산추가', `${body.name} ₩${Number(body.currentValue).toLocaleString()}`);
  return { ok: true, id };
}

function _updateAsset(body) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.ASSETS);
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, error: '데이터 없음' };
  const rows = sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === body.id) {
      if (body.currentValue !== undefined)
        sheet.getRange(i + 2, 4).setValue(Number(body.currentValue)).setNumberFormat('#,##0');
      if (body.memo !== undefined) sheet.getRange(i + 2, 6).setValue(body.memo);
      _logActivity(ss, body.member, '자산수정', body.name || body.id.slice(0, 8));
      return { ok: true };
    }
  }
  return { ok: false, error: '자산을 찾을 수 없음' };
}

// ============================================================
// 목표
// ============================================================
function _getGoals() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.GOALS);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return {
    ok: true,
    data: rows.filter(r => r[0]).map(r => {
      const target = Number(r[2]) || 0, current = Number(r[3]) || 0;
      return {
        id: r[7] || '', name: r[0], type: r[1], target, current,
        targetDate: r[4] ? _fmtDate(new Date(r[4])) : '',
        monthly: Number(r[5]) || 0, status: r[6] || '진행중',
        rate: target > 0 ? Math.round(current / target * 100) : 0,
        remaining: Math.max(target - current, 0),
      };
    }),
  };
}

function _addGoal(body) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.GOALS);
  if (!sheet) return { ok: false, error: '목표 시트 없음' };
  if (!body.name || !body.target) return { ok: false, error: '필수 항목 누락' };
  const id = Utilities.getUuid();
  const lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(lastRow + 1, 1, 1, 8).setValues([[
    body.name, body.type || '기타', Number(body.target), Number(body.current || 0),
    body.targetDate || '', Number(body.monthly || 0), '진행중', id,
  ]]);
  _logActivity(ss, body.member, '목표추가', `${body.name} ₩${Number(body.target).toLocaleString()}`);
  return { ok: true, id };
}

function _updateGoal(body) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.GOALS);
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, error: '데이터 없음' };
  const rows = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === body.id) {
      if (body.current  !== undefined) sheet.getRange(i + 2, 4).setValue(Number(body.current)).setNumberFormat('#,##0');
      if (body.monthly  !== undefined) sheet.getRange(i + 2, 6).setValue(Number(body.monthly)).setNumberFormat('#,##0');
      if (body.status   !== undefined) sheet.getRange(i + 2, 7).setValue(body.status);
      _logActivity(ss, body.member, '목표수정', body.name || body.id.slice(0, 8));
      return { ok: true };
    }
  }
  return { ok: false, error: '목표를 찾을 수 없음' };
}

// ============================================================
// 가족 구성원
// ============================================================
function _getMembers() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.MEMBERS);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return {
    ok: true,
    data: rows.filter(r => r[0]).map(r => ({
      name: r[0], role: r[1], color: r[2], emoji: r[3],
    })),
  };
}

function _addMember(body) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.MEMBERS);
  if (!sheet) return { ok: false, error: '구성원 시트 없음' };
  if (!body.name) return { ok: false, error: '이름 필수' };
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const colors  = ['#4f8ef7','#34d399','#f87171','#fbbf24','#a78bfa','#fb923c'];
  const emojis  = ['👨','👩','👧','👦','👴','👵'];
  const idx     = lastRow - 1;
  sheet.getRange(lastRow + 1, 1, 1, 4).setValues([[
    body.name, body.role || '가족',
    body.color  || colors[idx % colors.length],
    body.emoji  || emojis[idx % emojis.length],
  ]]);
  _logActivity(ss, '시스템', '구성원추가', body.name);
  return { ok: true };
}

// ============================================================
// 활동 로그
// ============================================================
function _getActivity() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.ACTIVITY);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };
  const cnt  = Math.min(sheet.getLastRow() - 1, 100);
  const rows = sheet.getRange(sheet.getLastRow() - cnt + 1, 1, cnt, 4).getValues();
  return {
    ok: true,
    data: rows.filter(r => r[0]).map(r => ({
      time: new Date(r[0]).toISOString(), member: r[1], action: r[2], detail: r[3],
    })).reverse(),
  };
}

function _logActivity(ss, member, action, detail) {
  const sheet = ss.getSheetByName(CFG.SHEETS.ACTIVITY);
  if (!sheet) return;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(lastRow + 1, 1, 1, 4).setValues([[
    new Date().toISOString(), member || '미지정', action, detail || '',
  ]]);
  // 1000행 이상이면 오래된 것 삭제
  if (sheet.getLastRow() > 1001) sheet.deleteRow(2);
}

// ============================================================
// 설정
// ============================================================
function _getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    ok: true,
    data: {
      budgets:    _getBudgets(ss),
      members:    (_getMembers().data || []),
      expenseCats: CFG.EXPENSE_CATS,
      incomeCats:  CFG.INCOME_CATS,
      payMethods:  CFG.PAY_METHODS,
      assetTypes:  CFG.ASSET_TYPES,
    },
  };
}

function _saveSettings(body) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.SETTINGS);
  if (!sheet) return { ok: false, error: '설정 시트 없음' };
  const rows = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  if (body.budgets) {
    rows.forEach((r, i) => {
      if (body.budgets[r[0]] !== undefined)
        sheet.getRange(i + 1, 2).setValue(Number(body.budgets[r[0]]));
    });
  }
  _logActivity(ss, body.member, '설정변경', '예산 업데이트');
  return { ok: true };
}

// ============================================================
// 내부 유틸
// ============================================================
function _getLedgerMonth(ss) {
  const sheet  = ss.getSheetByName(CFG.SHEETS.LEDGER);
  const result = { income: 0, expense: 0, categories: {} };
  if (!sheet || sheet.getLastRow() < 2) return result;
  const now  = new Date();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  rows.forEach(r => {
    if (!r[0]) return;
    const d = new Date(r[0]);
    if (d.getFullYear() !== now.getFullYear() || d.getMonth() !== now.getMonth()) return;
    const amt = Number(r[4]) || 0;
    if (r[1] === '수입') { result.income += amt; }
    else { result.expense += amt; result.categories[r[2]] = (result.categories[r[2]] || 0) + amt; }
  });
  return result;
}

function _sumCol(ss, sheetName, col) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  return sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues()
    .reduce((s, r) => s + (Number(r[0]) || 0), 0);
}

function _getBudgets(ss) {
  const sheet = ss.getSheetByName(CFG.SHEETS.SETTINGS);
  const budgets = {};
  if (!sheet) return budgets;
  const rows = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  rows.forEach(r => { if (r[0] && r[1] && CFG.EXPENSE_CATS.includes(r[0])) budgets[r[0]] = Number(r[1]); });
  return budgets;
}

function _getSettingVal(ss, key) {
  const sheet = ss.getSheetByName(CFG.SHEETS.SETTINGS);
  if (!sheet) return null;
  const rows = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  for (const r of rows) { if (r[0] === key) return r[1]; }
  return null;
}

function _autoClassify(desc) {
  const d = desc.toLowerCase();
  for (const [cat, kwStr] of Object.entries(CFG.KEYWORDS)) {
    for (const kw of kwStr.split(',')) {
      if (d.includes(kw.trim().toLowerCase())) return cat;
    }
  }
  return '';
}

function _fmtDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

// ============================================================
// 스프레드시트 메뉴 & 초기화
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 WEALTHOS Family')
    .addItem('🚀 시스템 초기화 (최초 1회)', 'initializeSystem')
    .addSeparator()
    .addItem('🔗 웹앱 URL 확인', 'showWebAppUrl')
    .addSeparator()
    .addItem('👨‍👩‍👧 구성원 추가', 'showAddMember')
    .addItem('➕ 거래 빠른 입력', 'showQuickEntry')
    .addItem('🔄 카테고리 자동 분류', 'runAutoClassify')
    .addSeparator()
    .addItem('📋 이번 달 리포트', 'generateReport')
    .addItem('❓ 사용법', 'showHelp')
    .addToUi();
}

function initializeSystem() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.alert('🚀 WEALTHOS Family 초기화',
    '가족 공유 재산관리 시스템을 설치합니다.\n계속하시겠습니까?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getActiveSpreadsheet().toast('⏳ 초기화 중...', 'WEALTHOS Family', 30);

  _buildSettings(ss);
  _buildMembers(ss);
  _buildActivity(ss);
  _buildAssets(ss);
  _buildGoals(ss);
  _buildLedger(ss);

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 완료!', 'WEALTHOS Family', 3);
  ui.alert('✅ 초기화 완료!',
    '다음 단계를 진행해주세요:\n\n' +
    '1. 메뉴 > 👨‍👩‍👧 구성원 추가 (가족 이름 입력)\n' +
    '2. 메뉴 > 🔗 웹앱 URL 확인\n' +
    '3. URL을 복사해 GitHub Pages 웹앱 설정에 입력\n\n' +
    '자세한 내용은 README.md를 참고하세요.',
    ui.ButtonSet.OK);
}

function showWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  try {
    const url = ScriptApp.getService().getUrl();
    ui.alert('🔗 웹앱 URL',
      '아래 URL을 복사해 GitHub Pages 웹앱 설정에 붙여넣으세요:\n\n' + url +
      '\n\n⚠️ 비어있으면: Apps Script > 배포 > 새 배포 > 웹 앱\n' +
      '  실행 주체: 나 / 액세스: 모든 사용자 (익명 포함)',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('배포 필요', 'Apps Script 편집기에서\n배포 > 새 배포 > 웹 앱을 먼저 실행하세요.', ui.ButtonSet.OK);
  }
}

function showAddMember() {
  const html = HtmlService.createHtmlOutput(`
<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Google Sans',sans-serif;padding:20px;background:#f8fafc}
h3{font-size:16px;margin-bottom:16px;color:#1e293b}label{display:block;font-size:12px;color:#64748b;margin:10px 0 4px}
input,select{width:100%;padding:9px 12px;border:1px solid #e2e8f0;border-radius:8px;font-size:13px}
input:focus,select:focus{outline:none;border-color:#4f8ef7}
button{width:100%;margin-top:16px;padding:11px;background:#4f8ef7;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer}
.msg{text-align:center;padding:10px;color:#16a34a;display:none;margin-top:8px}</style>
<h3>👨‍👩‍👧 구성원 추가</h3>
<label>이름</label><input id="name" placeholder="예: 아빠, 엄마, 첫째">
<label>역할</label>
<select id="role"><option>가족</option><option>부모</option><option>자녀</option></select>
<label>이모지</label>
<select id="emoji"><option>👨</option><option>👩</option><option>👧</option><option>👦</option><option>👴</option><option>👵</option></select>
<button onclick="add()">추가</button>
<div class="msg" id="msg">✅ 추가됐습니다!</div>
<script>
function add(){
  const name=document.getElementById('name').value.trim();
  if(!name){alert('이름을 입력하세요');return;}
  google.script.run.withSuccessHandler(()=>{
    document.getElementById('msg').style.display='block';
    document.getElementById('name').value='';
  })._addMemberFromUI({name,role:document.getElementById('role').value,emoji:document.getElementById('emoji').value});
}
</script>`).setWidth(300).setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, '구성원 추가');
}

function _addMemberFromUI(data) { _addMember({ ...data, action: 'addMember' }); }

function showQuickEntry() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const members = _getMembers().data || [];
  const memberOpts = members.map(m => `<option value="${m.name}">${m.emoji} ${m.name}</option>`).join('');

  const html = HtmlService.createHtmlOutput(`
<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Google Sans',sans-serif;padding:16px;background:#f8fafc;font-size:14px}
h3{color:#1e293b;font-size:15px;margin-bottom:14px}label{display:block;font-size:11px;color:#64748b;margin:9px 0 3px;text-transform:uppercase;letter-spacing:.5px}
input,select{width:100%;padding:8px 10px;border:1px solid #e2e8f0;border-radius:8px;font-size:13px;background:#fff}
input:focus,select:focus{outline:none;border-color:#4f8ef7}
.toggle{display:flex;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden}
.tbtn{flex:1;padding:8px;border:none;background:#fff;cursor:pointer;font-size:13px;color:#64748b}
.tbtn.exp{background:#fee2e2;color:#991b1b;font-weight:600}
.tbtn.inc{background:#d1fae5;color:#065f46;font-weight:600}
.row{display:flex;gap:8px}.row>div{flex:1}
button.save{width:100%;margin-top:14px;padding:11px;background:#4f8ef7;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer}
.msg{text-align:center;padding:6px;color:#16a34a;display:none;font-size:13px;margin-top:6px}</style>
<h3>➕ 거래 입력</h3>
<label>누가</label>
<select id="member">${memberOpts || '<option>미지정</option>'}</select>
<label>유형</label>
<div class="toggle">
  <button class="tbtn exp" onclick="setType('지출',this)">지출</button>
  <button class="tbtn" onclick="setType('수입',this)">수입</button>
</div>
<input type="hidden" id="type" value="지출">
<div class="row">
  <div><label>날짜</label><input type="date" id="date"></div>
  <div><label>금액(원)</label><input type="number" id="amt" placeholder="6400" min="1"></div>
</div>
<div class="row">
  <div><label>카테고리</label><select id="cat">
    <optgroup label="지출"><option>식비</option><option>카페</option><option>외식</option><option>교통</option><option>쇼핑</option><option>구독</option><option>통신</option><option>의료</option><option>문화</option><option>교육</option><option>보험</option><option>기타</option></optgroup>
    <optgroup label="수입"><option>급여</option><option>부업</option><option>투자수익</option><option>기타수입</option></optgroup>
  </select></div>
  <div><label>결제</label><select id="pay"><option>체크카드</option><option>신용카드</option><option>현금</option><option>계좌이체</option><option>간편결제</option></select></div>
</div>
<label>내용/가맹점</label><input id="desc" placeholder="예: 스타벅스">
<label>메모</label><input id="memo" placeholder="선택사항">
<button class="save" onclick="save()">저장</button>
<div class="msg" id="msg">✅ 저장됐습니다!</div>
<script>
document.getElementById('date').value=new Date().toISOString().slice(0,10);
function setType(v,el){document.getElementById('type').value=v;document.querySelectorAll('.tbtn').forEach(b=>{b.className='tbtn'});el.className='tbtn '+(v==='지출'?'exp':'inc')}
function save(){
  const d={date:document.getElementById('date').value,type:document.getElementById('type').value,
    category:document.getElementById('cat').value,desc:document.getElementById('desc').value,
    amount:document.getElementById('amt').value,payMethod:document.getElementById('pay').value,
    memo:document.getElementById('memo').value,member:document.getElementById('member').value};
  if(!d.amount||!d.desc){alert('내용과 금액을 입력하세요');return;}
  google.script.run.withSuccessHandler(()=>{
    document.getElementById('msg').style.display='block';
    document.getElementById('desc').value='';document.getElementById('amt').value='';
    setTimeout(()=>document.getElementById('msg').style.display='none',2000);
  })._addTransactionFromUI(d);
}
</script>`).setWidth(360).setHeight(570);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ 거래 입력');
}

function _addTransactionFromUI(data) { _addTransaction({ ...data, action: 'addTransaction' }); }

function runAutoClassify() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.LEDGER);
  if (!sheet || sheet.getLastRow() < 2) return;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  let count = 0;
  rows.forEach((r, i) => {
    if (!r[3]) return;
    const cat = _autoClassify(String(r[3]));
    if (cat && !r[2]) {
      sheet.getRange(i + 2, 3).setValue(cat);
      count++;
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ${count}개 자동 분류 완료`, 'WEALTHOS', 3);
}

function generateReport() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const data = _getDashboard().data;
  const now  = new Date();
  const lbl  = `${now.getFullYear()}년 ${now.getMonth() + 1}월`;
  const name = `📋 ${lbl} 리포트`;

  let sheet = ss.getSheetByName(name);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(name);
  sheet.setTabColor('#f59e0b');

  const H = (t, r) => {
    sheet.getRange(r, 1, 1, 8).merge().setValue(t)
      .setFontWeight('bold').setFontSize(13).setBackground('#1e293b').setFontColor('#fff');
    sheet.setRowHeight(r, 32);
  };
  const R = (k, v, r, bg) => {
    sheet.getRange(r, 1, 1, 4).merge().setValue(k).setBackground(bg || '#f8fafc');
    sheet.getRange(r, 5, 1, 4).merge().setValue(v).setHorizontalAlignment('right')
      .setFontWeight('bold').setBackground(bg || '#fff');
  };
  const fmt = n => Math.round(n).toLocaleString();
  let r = 1;

  sheet.getRange(r, 1, 1, 8).merge().setValue(`WEALTHOS Family — ${lbl} 재무 리포트`)
    .setFontSize(18).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff').setHorizontalAlignment('center');
  sheet.setRowHeight(r++, 48); r++;

  H('💰 순자산', r++);
  R('총 자산', `₩ ${fmt(data.totalAssets)}`, r++);
  R('총 부채', `₩ ${fmt(data.totalDebts)}`, r++);
  R('순 자산', `₩ ${fmt(data.netWorth)}`, r++, '#dbeafe'); r++;

  H('📊 이번 달 수입/지출', r++);
  R('수입', `₩ ${fmt(data.monthlyIncome)}`, r++);
  R('지출', `₩ ${fmt(data.monthlyExpense)}`, r++);
  R('저축', `₩ ${fmt(data.monthlyIncome - data.monthlyExpense)}`, r++, '#d1fae5');
  R('저축률', `${data.savingsRate}%`, r++); r++;

  H('👨‍👩‍👧 구성원별 지출', r++);
  Object.entries(data.memberExpense || {}).forEach(([m, amt]) => { R(m, `₩ ${fmt(amt)}`, r++); });
  r++;

  H('🏷️ 카테고리별 지출', r++);
  Object.entries(data.categories || {}).sort((a,b) => b[1]-a[1]).forEach(([cat, amt]) => {
    const pct = data.monthlyExpense > 0 ? Math.round(amt / data.monthlyExpense * 100) : 0;
    R(cat, `₩ ${fmt(amt)} (${pct}%)`, r++);
  });
  r++;

  H('🎯 목표 달성 현황', r++);
  (data.goals || []).forEach(g => {
    R(g.name, `${g.rate}% (₩${fmt(g.current)} / ₩${fmt(g.target)})`, r++,
      g.rate >= 80 ? '#d1fae5' : g.rate >= 50 ? '#eff6ff' : '#fff');
  });

  for (let c = 1; c <= 8; c++) sheet.setColumnWidth(c, 110);
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 리포트 생성 완료!', 'WEALTHOS', 3);
}

function showHelp() {
  SpreadsheetApp.getUi().alert('📖 WEALTHOS Family 사용법',
    '1. 메뉴 > 👨‍👩‍👧 구성원 추가로 가족 이름 입력\n' +
    '2. 메뉴 > 🔗 웹앱 URL 확인 후 GitHub Pages에 입력\n' +
    '3. 가족 모두가 같은 웹앱 URL 접속\n' +
    '4. 각자 이름 선택 후 거래 입력\n' +
    '5. 실시간으로 서로의 기록 확인!\n\n' +
    '💡 Tips:\n' +
    '• 거래 입력: 웹앱 또는 메뉴 > 거래 빠른 입력\n' +
    '• 30초마다 자동 동기화 (변경 있을 때 알림)\n' +
    '• 자산/목표는 가족 공동으로 관리',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
// 시트 빌더
// ============================================================
function _buildSettings(ss) {
  let s = ss.getSheetByName(CFG.SHEETS.SETTINGS);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(CFG.SHEETS.SETTINGS);
  s.setTabColor('#6b7280');
  const rows = [
    ['총부채', 0], ['', ''],
    ...CFG.EXPENSE_CATS.map(c => [c, 300000]),
  ];
  s.getRange(1, 1, rows.length, 2).setValues(rows);
  s.setColumnWidths(1, 2, [160, 140]);
}

function _buildMembers(ss) {
  let s = ss.getSheetByName(CFG.SHEETS.MEMBERS);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(CFG.SHEETS.MEMBERS);
  s.setTabColor('#4f8ef7');
  s.getRange(1, 1, 1, 4).setValues([['이름', '역할', '색상', '이모지']]);
  _hdr(s.getRange(1, 1, 1, 4));
  // 기본 구성원 샘플
  s.getRange(2, 1, 2, 4).setValues([
    ['아빠', '부모', '#4f8ef7', '👨'],
    ['엄마', '부모', '#f87171', '👩'],
  ]);
  s.setColumnWidths(1, 4, [120, 80, 100, 80]);
}

function _buildActivity(ss) {
  let s = ss.getSheetByName(CFG.SHEETS.ACTIVITY);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(CFG.SHEETS.ACTIVITY);
  s.setTabColor('#6b7280');
  s.getRange(1, 1, 1, 4).setValues([['시각', '구성원', '액션', '상세']]);
  _hdr(s.getRange(1, 1, 1, 4));
  s.setColumnWidths(1, 4, [200, 100, 120, 300]);
  _logActivity(ss, '시스템', '초기화', 'WEALTHOS Family 설치 완료');
}

function _buildAssets(ss) {
  let s = ss.getSheetByName(CFG.SHEETS.ASSETS);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(CFG.SHEETS.ASSETS);
  s.setTabColor('#3b82f6');
  const hdrs = ['자산명', '유형', '기관', '평가금액(원)', '취득금액(원)', '메모', 'ID'];
  s.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
  _hdr(s.getRange(1, 1, 1, hdrs.length));
  const sample = [
    ['국민은행 통장', '예금/현금', '국민은행', 8500000, 8500000, '', Utilities.getUuid()],
    ['카카오뱅크', '예금/현금', '카카오뱅크', 5200000, 5200000, '', Utilities.getUuid()],
    ['삼성전자 주식', '주식/ETF', '키움증권', 4800000, 3500000, '100주', Utilities.getUuid()],
    ['아파트', '부동산', '서울 마포', 320000000, 280000000, '34평', Utilities.getUuid()],
  ];
  s.getRange(2, 1, sample.length, hdrs.length).setValues(sample);
  s.getRange(2, 4, sample.length, 2).setNumberFormat('#,##0');
  s.setColumnWidths(1, 7, [160, 110, 130, 140, 130, 180, 0]);
  s.setColumnWidth(7, 1); // ID 컬럼 숨김
  s.setFrozenRows(1);
}

function _buildGoals(ss) {
  let s = ss.getSheetByName(CFG.SHEETS.GOALS);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(CFG.SHEETS.GOALS);
  s.setTabColor('#f59e0b');
  const hdrs = ['목표명', '유형', '목표금액(원)', '현재금액(원)', '목표일', '월저축(원)', '상태', 'ID'];
  s.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
  _hdr(s.getRange(1, 1, 1, hdrs.length));
  const sample = [
    ['내집마련', '부동산', 500000000, 310000000, '2027-06-01', 1200000, '진행중', Utilities.getUuid()],
    ['노후자금', '은퇴', 300000000, 84000000, '2038-01-01', 800000, '진행중', Utilities.getUuid()],
    ['차량구매', '자동차', 50000000, 42000000, '2025-08-01', 500000, '진행중', Utilities.getUuid()],
    ['유럽여행', '여행', 5000000, 1200000, '2025-12-01', 380000, '진행중', Utilities.getUuid()],
  ];
  s.getRange(2, 1, sample.length, hdrs.length).setValues(sample);
  s.getRange(2, 3, sample.length, 4).setNumberFormat('#,##0');
  s.setColumnWidths(1, 8, [140, 90, 140, 140, 110, 120, 70, 0]);
  s.setColumnWidth(8, 1);
  s.setFrozenRows(1);
}

function _buildLedger(ss) {
  let s = ss.getSheetByName(CFG.SHEETS.LEDGER);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(CFG.SHEETS.LEDGER);
  s.setTabColor('#8b5cf6');
  const hdrs = ['날짜', '유형', '카테고리', '내용/가맹점', '금액(원)', '결제수단', '메모', '작성자', 'ID', '생성시각'];
  s.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
  _hdr(s.getRange(1, 1, 1, hdrs.length));

  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['수입', '지출'], true).build();
  s.getRange('B2:B5000').setDataValidation(typeRule);
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([...CFG.INCOME_CATS, ...CFG.EXPENSE_CATS], true).build();
  s.getRange('C2:C5000').setDataValidation(catRule);
  const payRule = SpreadsheetApp.newDataValidation().requireValueInList(CFG.PAY_METHODS, true).build();
  s.getRange('F2:F5000').setDataValidation(payRule);

  const now = new Date();
  const ym  = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const sample = [
    [`${ym}-01`, '수입', '급여', '3월 급여', 4200000, '계좌이체', '', '아빠', Utilities.getUuid(), new Date().toISOString()],
    [`${ym}-02`, '지출', '교통', 'T-money 충전', 50000, '체크카드', '', '아빠', Utilities.getUuid(), new Date().toISOString()],
    [`${ym}-03`, '지출', '카페', '스타벅스', 6400, '신용카드', '', '엄마', Utilities.getUuid(), new Date().toISOString()],
    [`${ym}-05`, '지출', '식비', 'GS25', 8500, '체크카드', '', '아빠', Utilities.getUuid(), new Date().toISOString()],
    [`${ym}-07`, '지출', '외식', '점심 외식', 45000, '현금', '가족 식사', '엄마', Utilities.getUuid(), new Date().toISOString()],
    [`${ym}-10`, '지출', '쇼핑', '쿠팡', 89000, '신용카드', '생활용품', '엄마', Utilities.getUuid(), new Date().toISOString()],
    [`${ym}-12`, '지출', '구독', '넷플릭스', 17000, '신용카드', '', '아빠', Utilities.getUuid(), new Date().toISOString()],
    [`${ym}-15`, '수입', '급여', '3월 급여', 3800000, '계좌이체', '', '엄마', Utilities.getUuid(), new Date().toISOString()],
    [`${ym}-18`, '지출', '의료', '소아과 진료', 15000, '체크카드', '', '엄마', Utilities.getUuid(), new Date().toISOString()],
    [`${ym}-22`, '지출', '교육', '학원비', 300000, '계좌이체', '', '아빠', Utilities.getUuid(), new Date().toISOString()],
  ];
  s.getRange(2, 1, sample.length, hdrs.length).setValues(sample);
  s.getRange(2, 5, sample.length, 1).setNumberFormat('#,##0');
  s.setColumnWidths(1, 10, [110, 65, 110, 200, 120, 100, 160, 80, 0, 0]);
  s.setColumnWidth(9, 1);  // ID 숨김
  s.setColumnWidth(10, 1); // 생성시각 숨김
  s.setFrozenRows(1);

  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="수입"')
      .setBackground('#d1fae5').setRanges([s.getRange('A2:J5000')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="지출"')
      .setBackground('#fff1f2').setRanges([s.getRange('A2:J5000')]).build(),
  ]);
}

function _hdr(range) {
  range.setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
}
