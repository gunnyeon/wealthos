// ============================================================
// WEALTHOS Family — Apps Script 백엔드
// CORS 완전 해결: ContentService + Access-Control-Allow-Origin
// ============================================================

const SHEET = {
  LEDGER:   '가계부',
  ASSETS:   '자산',
  GOALS:    '목표',
  SETTINGS: '설정',
  MEMBERS:  '가족구성원',
  LOG:      '활동로그',
};

const KEYWORDS = {
  '식비':  ['GS25','CU','세븐일레븐','이마트','홈플러스','롯데마트','마트','슈퍼','편의점','다이소'],
  '카페':  ['스타벅스','카페','커피','메가커피','투썸','빽다방','할리스','이디야','폴바셋'],
  '외식':  ['배달의민족','배민','쿠팡이츠','요기요','맥도날드','버거킹','피자','치킨','음식점','식당','라멘','파스타'],
  '교통':  ['T-money','티머니','지하철','버스','택시','카카오택시','KTX','SRT','주차'],
  '주유':  ['주유소','GS칼텍스','SK에너지','현대오일뱅크','S-OIL'],
  '쇼핑':  ['쿠팡','네이버쇼핑','11번가','올리브영','무신사','G마켓','옥션','위메프','티몬'],
  '구독':  ['넷플릭스','유튜브프리미엄','멜론','스포티파이','왓챠','디즈니','ChatGPT','Adobe','Apple'],
  '통신':  ['SKT','KT','LG유플러스','알뜰폰','통신'],
  '의료':  ['병원','약국','의원','내과','치과','한의원','의료'],
  '교육':  ['학원','교육','학습','온라인강의','인강','책','도서'],
  '급여':  ['급여','월급','임금','봉급','연봉'],
  '부업':  ['프리랜서','강의료','외주','용돈','알바'],
};

// ── 응답 헬퍼 ─────────────────────────────────────────
function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 진입점 ────────────────────────────────────────────
function doGet(e) {
  try {
    const p = e.parameter || {};
    switch (p.action) {
      case 'getDashboard':      return ok(getDashboard());
      case 'getLedger':         return ok(getLedger(p));
      case 'getAssets':         return ok(getAssets());
      case 'getGoals':          return ok(getGoals());
      case 'getMembers':        return ok(getMembers());
      case 'getActivity':       return ok(getActivity(20));
      case 'poll':              return ok(poll(p.since || ''));
      case 'addTransaction':    return ok(addTransaction(p));
      case 'deleteTransaction': return ok(deleteTransaction(p));
      case 'addAsset':          return ok(addAsset(p));
      case 'updateAsset':       return ok(updateAsset(p));
      case 'addGoal':           return ok(addGoal(p));
      case 'updateGoal':        return ok(updateGoal(p));
      case 'addMember':         return ok(addMember(p));
      case 'saveSettings':      return ok(saveSettings(p));
      default:                  return ok({ version: '4.0', status: 'ok' });
    }
  } catch(e) {
    return err(e.message);
  }
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents || '{}');
    return doGet({ parameter: p });
  } catch(e) {
    return err(e.message);
  }
}

// ── 폴링 (실시간 핵심) ────────────────────────────────
function poll(since) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sinceDate = since ? new Date(since) : new Date(0);
  const now       = new Date();

  const newTxns = getLedgerSince(ss, sinceDate);
  const newLogs = getLogSince(ss, sinceDate);
  const summary = getLedgerMonth(ss);

  return {
    timestamp:       now.toISOString(),
    hasChanges:      newTxns.length > 0 || newLogs.length > 0,
    newTransactions: newTxns,
    newActivity:     newLogs,
    summary:         summary,
  };
}

function getLedgerSince(ss, sinceDate) {
  const sheet = ss.getSheetByName(SHEET.LEDGER);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  return rows
    .filter(r => r[0] && r[9] && new Date(r[9]) > sinceDate)
    .map(rowToTxn)
    .reverse();
}

function getLogSince(ss, sinceDate) {
  const sheet = ss.getSheetByName(SHEET.LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return rows
    .filter(r => r[0] && new Date(r[0]) > sinceDate)
    .map(r => ({ time: new Date(r[0]).toISOString(), member: r[1], action: r[2], detail: r[3] }))
    .reverse();
}

// ── 대시보드 ──────────────────────────────────────────
function getDashboard() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const totalAssets = sumCol(ss, SHEET.ASSETS, 4);
  const totalDebts  = Number(getSettingVal(ss, '총부채') || 0);
  const summary     = getLedgerMonth(ss);
  const goals       = getGoals();
  const members     = getMembers();
  const activity    = getActivity(15);

  // 구성원별 이번 달 지출
  const memberExpense = {};
  members.forEach(m => { memberExpense[m.name] = 0; });

  const ledger = ss.getSheetByName(SHEET.LEDGER);
  if (ledger && ledger.getLastRow() > 1) {
    const now  = new Date();
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

  // 최근 거래 10건
  const recentTxns = getLedger({ limit: '10' });

  return {
    totalAssets, totalDebts,
    netWorth:       totalAssets - totalDebts,
    monthlyIncome:  summary.income,
    monthlyExpense: summary.expense,
    savingsRate:    summary.income > 0
      ? Math.round((summary.income - summary.expense) / summary.income * 100) : 0,
    categories:     summary.categories,
    budgets:        getBudgets(ss),
    goals, members, memberExpense,
    recentTxns,
    recentActivity: activity,
    updatedAt:      new Date().toISOString(),
  };
}

// ── 가계부 ────────────────────────────────────────────
function getLedger(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.LEDGER);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows   = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  const month  = p.month  || '';
  const member = p.member || '';
  const limit  = Number(p.limit) || 0;

  const result = rows
    .filter(r => {
      if (!r[0] || !r[4]) return false;
      const d  = new Date(r[0]);
      const ym = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
      if (month  && ym   !== month)  return false;
      if (member && r[7] !== member) return false;
      return true;
    })
    .map(rowToTxn)
    .reverse();

  return limit > 0 ? result.slice(0, limit) : result;
}

function rowToTxn(r) {
  const d = new Date(r[0]);
  return {
    id:        r[8] || '',
    date:      `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`,
    type:      r[1], category: r[2], desc: r[3],
    amount:    Number(r[4]) || 0,
    payMethod: r[5], memo: r[6], member: r[7],
    createdAt: r[9] ? new Date(r[9]).toISOString() : '',
  };
}

// ── 거래 추가 ─────────────────────────────────────────
function addTransaction(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.LEDGER);
  if (!sheet) throw new Error('가계부 시트 없음 — 메뉴 > 초기화를 먼저 실행해주세요');

  if (!p.date || !p.amount) throw new Error('날짜와 금액은 필수입니다');

  let cat = p.category || autoClassify(p.desc || '');
  if (!cat) cat = (p.type === '수입') ? '기타수입' : '기타';

  const id        = Utilities.getUuid();
  const createdAt = new Date().toISOString();
  const lastRow   = Math.max(sheet.getLastRow(), 1);

  sheet.getRange(lastRow + 1, 1, 1, 10).setValues([[
    p.date, p.type || '지출', cat, p.desc || '',
    Number(p.amount), p.payMethod || '', p.memo || '',
    p.member || '미지정', id, createdAt,
  ]]);
  sheet.getRange(lastRow + 1, 5).setNumberFormat('#,##0');

  log(ss, p.member, '거래추가', `${p.type} ${cat} ₩${Number(p.amount).toLocaleString()}`);
  return { id, category: cat, createdAt };
}

// ── 거래 삭제 ─────────────────────────────────────────
function deleteTransaction(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.LEDGER);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('데이터 없음');

  const rows = sheet.getRange(2, 9, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === p.id) {
      sheet.deleteRow(i + 2);
      log(ss, p.member, '거래삭제', p.desc || p.id.slice(0, 8));
      return { deleted: true };
    }
  }
  throw new Error('거래를 찾을 수 없습니다');
}

// ── 자산 ──────────────────────────────────────────────
function getAssets() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ASSETS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues()
    .filter(r => r[0])
    .map(r => ({
      id: r[6] || '', name: r[0], type: r[1], institution: r[2],
      currentValue: Number(r[3]) || 0, purchaseValue: Number(r[4]) || 0, memo: r[5],
    }));
}

function addAsset(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ASSETS);
  if (!sheet) throw new Error('자산 시트 없음');
  if (!p.name || !p.currentValue) throw new Error('자산명과 금액은 필수입니다');

  const id = Utilities.getUuid();
  const lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(lastRow + 1, 1, 1, 7).setValues([[
    p.name, p.type || '기타', p.institution || '',
    Number(p.currentValue), Number(p.purchaseValue || p.currentValue),
    p.memo || '', id,
  ]]);
  sheet.getRange(lastRow + 1, 4, 1, 2).setNumberFormat('#,##0');
  log(ss, p.member, '자산추가', `${p.name} ₩${Number(p.currentValue).toLocaleString()}`);
  return { id };
}

function updateAsset(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ASSETS);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('데이터 없음');

  const rows = sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === p.id) {
      if (p.currentValue) sheet.getRange(i + 2, 4).setValue(Number(p.currentValue)).setNumberFormat('#,##0');
      if (p.memo !== undefined) sheet.getRange(i + 2, 6).setValue(p.memo);
      log(ss, p.member, '자산수정', p.name || p.id.slice(0, 8));
      return { updated: true };
    }
  }
  throw new Error('자산을 찾을 수 없습니다');
}

// ── 목표 ──────────────────────────────────────────────
function getGoals() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.GOALS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
    .filter(r => r[0])
    .map(r => {
      const target = Number(r[2]) || 0;
      const current = Number(r[3]) || 0;
      return {
        id: r[7] || '', name: r[0], type: r[1],
        target, current,
        targetDate: r[4] ? fmtDate(new Date(r[4])) : '',
        monthly: Number(r[5]) || 0,
        status: r[6] || '진행중',
        rate: target > 0 ? Math.round(current / target * 100) : 0,
        remaining: Math.max(target - current, 0),
      };
    });
}

function addGoal(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.GOALS);
  if (!sheet) throw new Error('목표 시트 없음');
  if (!p.name || !p.target) throw new Error('목표명과 금액은 필수입니다');

  const id = Utilities.getUuid();
  const lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(lastRow + 1, 1, 1, 8).setValues([[
    p.name, p.type || '기타', Number(p.target), Number(p.current || 0),
    p.targetDate || '', Number(p.monthly || 0), '진행중', id,
  ]]);
  sheet.getRange(lastRow + 1, 3, 1, 4).setNumberFormat('#,##0');
  log(ss, p.member, '목표추가', `${p.name} ₩${Number(p.target).toLocaleString()}`);
  return { id };
}

function updateGoal(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.GOALS);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('데이터 없음');

  const rows = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === p.id) {
      if (p.current  !== undefined) sheet.getRange(i + 2, 4).setValue(Number(p.current)).setNumberFormat('#,##0');
      if (p.monthly  !== undefined) sheet.getRange(i + 2, 6).setValue(Number(p.monthly)).setNumberFormat('#,##0');
      if (p.status   !== undefined) sheet.getRange(i + 2, 7).setValue(p.status);
      log(ss, p.member, '목표수정', p.name || p.id.slice(0, 8));
      return { updated: true };
    }
  }
  throw new Error('목표를 찾을 수 없습니다');
}

// ── 가족 구성원 ───────────────────────────────────────
function getMembers() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.MEMBERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(r => r[0])
    .map(r => ({ name: r[0], role: r[1], color: r[2], emoji: r[3] }));
}

function addMember(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.MEMBERS);
  if (!sheet) throw new Error('구성원 시트 없음');
  if (!p.name) throw new Error('이름은 필수입니다');

  const colors = ['#4f8ef7','#f87171','#34d399','#fbbf24','#a78bfa','#fb923c'];
  const emojis = ['👨','👩','👧','👦','👴','👵'];
  const idx    = Math.max(sheet.getLastRow() - 1, 0);

  sheet.getRange(Math.max(sheet.getLastRow(), 1) + 1, 1, 1, 4).setValues([[
    p.name, p.role || '가족',
    p.color || colors[idx % colors.length],
    p.emoji || emojis[idx % emojis.length],
  ]]);
  log(ss, '시스템', '구성원추가', p.name);
  return { added: true };
}

// ── 활동 로그 ─────────────────────────────────────────
function getActivity(limit) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const cnt  = Math.min(sheet.getLastRow() - 1, 200);
  const rows = sheet.getRange(sheet.getLastRow() - cnt + 1, 1, cnt, 4).getValues();
  return rows
    .filter(r => r[0])
    .map(r => ({ time: new Date(r[0]).toISOString(), member: r[1], action: r[2], detail: r[3] }))
    .reverse()
    .slice(0, limit || 50);
}

function log(ss, member, action, detail) {
  const sheet = ss.getSheetByName(SHEET.LOG);
  if (!sheet) return;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(lastRow + 1, 1, 1, 4).setValues([[
    new Date().toISOString(), member || '미지정', action, detail || '',
  ]]);
  // 1000행 초과시 오래된 것 삭제
  if (sheet.getLastRow() > 1001) sheet.deleteRow(2);
}

// ── 설정 ──────────────────────────────────────────────
function saveSettings(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.SETTINGS);
  if (!sheet) throw new Error('설정 시트 없음');

  const rows = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  // budget_식비, budget_카페 등 파라미터 처리
  rows.forEach((r, i) => {
    const key = `budget_${r[0]}`;
    if (p[key] !== undefined) {
      sheet.getRange(i + 1, 2).setValue(Number(p[key]));
    }
    if (r[0] === '총부채' && p.totalDebts !== undefined) {
      sheet.getRange(i + 1, 2).setValue(Number(p.totalDebts));
    }
  });
  log(ss, p.member, '설정변경', '예산 업데이트');
  return { saved: true };
}

// ── 유틸 ──────────────────────────────────────────────
function getLedgerMonth(ss) {
  const sheet  = ss.getSheetByName(SHEET.LEDGER);
  const result = { income: 0, expense: 0, categories: {} };
  if (!sheet || sheet.getLastRow() < 2) return result;

  const now  = new Date();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  rows.forEach(r => {
    if (!r[0]) return;
    const d = new Date(r[0]);
    if (d.getFullYear() !== now.getFullYear() || d.getMonth() !== now.getMonth()) return;
    const amt = Number(r[4]) || 0;
    if (r[1] === '수입') {
      result.income += amt;
    } else {
      result.expense += amt;
      const cat = r[2] || '기타';
      result.categories[cat] = (result.categories[cat] || 0) + amt;
    }
  });
  return result;
}

function sumCol(ss, sheetName, col) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  return sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues()
    .reduce((s, r) => s + (Number(r[0]) || 0), 0);
}

function getBudgets(ss) {
  const sheet   = ss.getSheetByName(SHEET.SETTINGS);
  const budgets = {};
  if (!sheet) return budgets;
  const cats = ['식비','카페','외식','교통','주유','쇼핑','구독','통신','의료','문화','교육','경조사','보험','기타'];
  const rows = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  rows.forEach(r => { if (r[0] && r[1] && cats.includes(r[0])) budgets[r[0]] = Number(r[1]); });
  return budgets;
}

function getSettingVal(ss, key) {
  const sheet = ss.getSheetByName(SHEET.SETTINGS);
  if (!sheet) return null;
  const rows = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  for (const r of rows) { if (r[0] === key) return r[1]; }
  return null;
}

function autoClassify(desc) {
  const d = desc.toLowerCase();
  for (const [cat, kws] of Object.entries(KEYWORDS)) {
    for (const kw of kws) {
      if (d.includes(kw.toLowerCase())) return cat;
    }
  }
  return '';
}

function fmtDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

// ── 스프레드시트 초기화 ───────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 WEALTHOS')
    .addItem('🚀 초기화 (최초 1회)', 'initSystem')
    .addSeparator()
    .addItem('🔗 웹앱 URL 확인', 'showUrl')
    .addItem('👨‍👩‍👧 구성원 추가', 'showAddMember')
    .addItem('➕ 거래 입력', 'showAddTxn')
    .addSeparator()
    .addItem('📋 이번 달 리포트', 'makeReport')
    .addToUi();
}

function initSystem() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.alert('🚀 초기화', '시트를 생성합니다. 계속하시겠습니까?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getActiveSpreadsheet().toast('⏳ 초기화 중...', 'WEALTHOS', 30);

  buildSettings(ss);
  buildMembers(ss);
  buildLog(ss);
  buildAssets(ss);
  buildGoals(ss);
  buildLedger(ss);

  ui.alert('✅ 완료',
    '초기화가 완료됐습니다!\n\n' +
    '다음 단계:\n' +
    '1. 배포 → 새 배포 → 웹 앱\n' +
    '   실행 주체: 나 / 액세스: 모든 사용자\n' +
    '2. URL 복사 → GitHub Pages index.html에 입력\n' +
    '3. 메뉴 > 구성원 추가로 가족 이름 등록',
    ui.ButtonSet.OK);
}

function showUrl() {
  const ui = SpreadsheetApp.getUi();
  try {
    const url = ScriptApp.getService().getUrl();
    ui.alert('🔗 웹앱 URL', url + '\n\n이 URL을 GitHub Pages index.html에 입력하세요.', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('배포 필요', '배포 → 새 배포 → 웹 앱을 먼저 실행하세요.', ui.ButtonSet.OK);
  }
}

function showAddMember() {
  const html = HtmlService.createHtmlOutput(`
<style>*{box-sizing:border-box}body{font-family:sans-serif;padding:16px;background:#f8fafc}
label{display:block;font-size:12px;color:#64748b;margin:8px 0 3px}
input,select{width:100%;padding:8px;border:1px solid #e2e8f0;border-radius:6px;font-size:13px}
button{width:100%;margin-top:14px;padding:10px;background:#4f8ef7;color:#fff;border:none;border-radius:6px;font-size:14px;cursor:pointer}
.ok{color:#16a34a;text-align:center;margin-top:8px;display:none}</style>
<label>이름</label><input id="n" placeholder="예: 아빠">
<label>역할</label><select id="r"><option>가족</option><option>부모</option><option>자녀</option></select>
<label>이모지</label><select id="e"><option>👨</option><option>👩</option><option>👧</option><option>👦</option><option>👴</option><option>👵</option></select>
<button onclick="add()">추가</button>
<div class="ok" id="ok">✅ 추가됐습니다!</div>
<script>
function add(){
  const name=document.getElementById('n').value.trim();
  if(!name){alert('이름을 입력하세요');return;}
  google.script.run.withSuccessHandler(()=>{
    document.getElementById('ok').style.display='block';
    document.getElementById('n').value='';
  }).addMemberUI({name,role:document.getElementById('r').value,emoji:document.getElementById('e').value});
}
</script>`).setWidth(280).setHeight(310);
  SpreadsheetApp.getUi().showModalDialog(html, '구성원 추가');
}
function addMemberUI(p) { addMember(p); }

function showAddTxn() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const members = getMembers();
  const mOpts   = members.map(m => `<option>${m.emoji} ${m.name}</option>`).join('') || '<option>미지정</option>';
  const html = HtmlService.createHtmlOutput(`
<style>*{box-sizing:border-box}body{font-family:sans-serif;padding:16px;background:#f8fafc;font-size:13px}
label{display:block;font-size:11px;color:#64748b;margin:8px 0 3px;text-transform:uppercase}
input,select{width:100%;padding:8px;border:1px solid #e2e8f0;border-radius:6px;font-size:13px}
.row{display:flex;gap:8px}.row>div{flex:1}
button{width:100%;margin-top:12px;padding:10px;background:#4f8ef7;color:#fff;border:none;border-radius:6px;font-size:14px;cursor:pointer}
.ok{color:#16a34a;text-align:center;margin-top:6px;display:none}
.tt{display:flex;border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;margin-top:3px}
.tb{flex:1;padding:8px;border:none;background:#fff;cursor:pointer;font-size:13px;color:#64748b}
.tb.e{background:#fee2e2;color:#991b1b}.tb.i{background:#d1fae5;color:#065f46}</style>
<label>누가</label><select id="who">${mOpts}</select>
<label>유형</label>
<div class="tt">
  <button class="tb e" id="be" onclick="st('지출',this)">지출</button>
  <button class="tb" id="bi" onclick="st('수입',this)">수입</button>
</div><input type="hidden" id="tp" value="지출">
<div class="row">
  <div><label>날짜</label><input type="date" id="dt"></div>
  <div><label>금액(원)</label><input type="number" id="am" placeholder="6400"></div>
</div>
<div class="row">
  <div><label>카테고리</label><select id="ct">
    <option>식비</option><option>카페</option><option>외식</option><option>교통</option>
    <option>쇼핑</option><option>구독</option><option>통신</option><option>의료</option>
    <option>문화</option><option>교육</option><option>보험</option><option>기타</option>
    <option>급여</option><option>부업</option><option>투자수익</option><option>기타수입</option>
  </select></div>
  <div><label>결제수단</label><select id="pm">
    <option>체크카드</option><option>신용카드</option><option>현금</option><option>계좌이체</option><option>간편결제</option>
  </select></div>
</div>
<label>내용/가맹점</label><input id="dc" placeholder="예: 스타벅스 강남점">
<label>메모</label><input id="mm" placeholder="선택사항">
<button onclick="save()">저장</button>
<div class="ok" id="ok">✅ 저장됐습니다!</div>
<script>
document.getElementById('dt').value=new Date().toISOString().slice(0,10);
function st(v,el){document.getElementById('tp').value=v;document.querySelectorAll('.tb').forEach(b=>b.className='tb');el.className='tb '+(v==='지출'?'e':'i');}
function save(){
  const d={type:document.getElementById('tp').value,date:document.getElementById('dt').value,
    amount:document.getElementById('am').value,category:document.getElementById('ct').value,
    payMethod:document.getElementById('pm').value,desc:document.getElementById('dc').value,
    memo:document.getElementById('mm').value,
    member:document.getElementById('who').value.replace(/^[^ ]+ /,'')};
  if(!d.amount||!d.desc){alert('내용과 금액을 입력하세요');return;}
  google.script.run.withSuccessHandler(()=>{
    document.getElementById('ok').style.display='block';
    document.getElementById('am').value='';document.getElementById('dc').value='';
    setTimeout(()=>document.getElementById('ok').style.display='none',2000);
  }).addTxnUI(d);
}
</script>`).setWidth(360).setHeight(530);
  SpreadsheetApp.getUi().showModalDialog(html, '거래 입력');
}
function addTxnUI(p) { addTransaction(p); }

function makeReport() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const d    = getDashboard();
  const now  = new Date();
  const lbl  = `${now.getFullYear()}년 ${now.getMonth()+1}월`;
  const name = `📋 ${lbl} 리포트`;

  let s = ss.getSheetByName(name);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(name);
  s.setTabColor('#f59e0b');

  const f = n => Math.round(n).toLocaleString();
  const H = (t, r) => {
    s.getRange(r, 1, 1, 6).merge().setValue(t)
      .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold').setFontSize(13);
    s.setRowHeight(r, 30);
  };
  const R = (k, v, r, bg) => {
    s.getRange(r, 1, 1, 3).merge().setValue(k).setBackground(bg||'#f8fafc');
    s.getRange(r, 4, 1, 3).merge().setValue(v).setHorizontalAlignment('right').setFontWeight('bold').setBackground(bg||'#fff');
  };

  let r = 1;
  s.getRange(r, 1, 1, 6).merge().setValue(`WEALTHOS Family — ${lbl} 재무 리포트`)
    .setFontSize(17).setFontWeight('bold').setBackground('#0f172a').setFontColor('#fff').setHorizontalAlignment('center');
  s.setRowHeight(r++, 44); r++;

  H('💰 순자산', r++);
  R('총 자산', `₩ ${f(d.totalAssets)}`, r++);
  R('총 부채', `₩ ${f(d.totalDebts)}`, r++);
  R('순 자산', `₩ ${f(d.netWorth)}`, r++, '#dbeafe'); r++;

  H('📊 이번 달', r++);
  R('수입', `₩ ${f(d.monthlyIncome)}`, r++, '#d1fae5');
  R('지출', `₩ ${f(d.monthlyExpense)}`, r++, '#fee2e2');
  R('저축', `₩ ${f(d.monthlyIncome-d.monthlyExpense)}`, r++);
  R('저축률', `${d.savingsRate}%`, r++); r++;

  H('👨‍👩‍👧 구성원별', r++);
  Object.entries(d.memberExpense||{}).forEach(([m,a]) => R(m, `₩ ${f(a)}`, r++));
  r++;

  H('🏷️ 카테고리', r++);
  Object.entries(d.categories||{}).sort((a,b)=>b[1]-a[1]).forEach(([c,a]) => {
    const pct = d.monthlyExpense > 0 ? Math.round(a/d.monthlyExpense*100) : 0;
    R(c, `₩ ${f(a)} (${pct}%)`, r++);
  });
  r++;

  H('🎯 목표', r++);
  (d.goals||[]).forEach(g => R(g.name, `${g.rate}% (₩${f(g.current)}/₩${f(g.target)})`, r++,
    g.rate>=80?'#d1fae5':g.rate>=50?'#eff6ff':'#fff'));

  for (let c=1; c<=6; c++) s.setColumnWidth(c, 120);
  ss.setActiveSheet(s);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 리포트 생성 완료!', 'WEALTHOS', 3);
}

// ── 시트 빌더 ─────────────────────────────────────────
function hdr(range) {
  range.setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
}
function alt(sheet, start, n, cols) {
  for (let i=0; i<n; i++)
    sheet.getRange(start+i, 1, 1, cols).setBackground(i%2===0?'#fff':'#f8fafc');
}

function buildSettings(ss) {
  let s = ss.getSheetByName(SHEET.SETTINGS);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(SHEET.SETTINGS);
  s.setTabColor('#6b7280');
  const rows = [
    ['총부채', 0], ['', ''],
    ['식비',300000],['카페',100000],['외식',250000],['교통',150000],
    ['주유',100000],['쇼핑',300000],['구독',50000],['통신',80000],
    ['의료',80000],['문화',150000],['교육',300000],['경조사',100000],
    ['보험',100000],['기타',200000],
  ];
  s.getRange(1, 1, rows.length, 2).setValues(rows);
  s.setColumnWidths(1, 2, [150, 140]);
}

function buildMembers(ss) {
  let s = ss.getSheetByName(SHEET.MEMBERS);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(SHEET.MEMBERS);
  s.setTabColor('#4f8ef7');
  s.getRange(1, 1, 1, 4).setValues([['이름','역할','색상','이모지']]);
  hdr(s.getRange(1, 1, 1, 4));
  s.getRange(2, 1, 3, 4).setValues([
    ['아빠','부모','#4f8ef7','👨'],
    ['엄마','부모','#f87171','👩'],
    ['자녀','자녀','#34d399','👧'],
  ]);
  s.setColumnWidths(1, 4, [120,80,100,80]);
  s.setFrozenRows(1);
}

function buildLog(ss) {
  let s = ss.getSheetByName(SHEET.LOG);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(SHEET.LOG);
  s.setTabColor('#6b7280');
  s.getRange(1, 1, 1, 4).setValues([['시각','구성원','액션','상세']]);
  hdr(s.getRange(1, 1, 1, 4));
  s.setColumnWidths(1, 4, [200,100,120,300]);
  s.setFrozenRows(1);
  log(ss, '시스템', '초기화', 'WEALTHOS Family 설치 완료');
}

function buildAssets(ss) {
  let s = ss.getSheetByName(SHEET.ASSETS);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(SHEET.ASSETS);
  s.setTabColor('#3b82f6');
  const h = ['자산명','유형','기관','평가금액(원)','취득금액(원)','메모','ID'];
  s.getRange(1, 1, 1, h.length).setValues([h]);
  hdr(s.getRange(1, 1, 1, h.length));

  const data = [
    ['국민은행 통장','예금/현금','국민은행',8500000,8500000,'',Utilities.getUuid()],
    ['카카오뱅크','예금/현금','카카오뱅크',5200000,5200000,'',Utilities.getUuid()],
    ['삼성전자 주식','주식/ETF','키움증권',4800000,3500000,'100주',Utilities.getUuid()],
    ['아파트','부동산','서울 마포',320000000,280000000,'34평',Utilities.getUuid()],
    ['청년도약계좌','적금/보험','신한은행',6000000,6000000,'만기2028',Utilities.getUuid()],
  ];
  s.getRange(2, 1, data.length, h.length).setValues(data);
  s.getRange(2, 4, data.length, 2).setNumberFormat('#,##0');
  s.setColumnWidths(1, 7, [160,110,130,140,130,180,1]);
  s.setFrozenRows(1);
  alt(s, 2, data.length, h.length);
}

function buildGoals(ss) {
  let s = ss.getSheetByName(SHEET.GOALS);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(SHEET.GOALS);
  s.setTabColor('#f59e0b');
  const h = ['목표명','유형','목표금액(원)','현재금액(원)','목표일','월저축(원)','상태','ID'];
  s.getRange(1, 1, 1, h.length).setValues([h]);
  hdr(s.getRange(1, 1, 1, h.length));

  const data = [
    ['내집마련','부동산',500000000,310000000,'2027-06-01',1200000,'진행중',Utilities.getUuid()],
    ['노후자금','은퇴',300000000,84000000,'2038-01-01',800000,'진행중',Utilities.getUuid()],
    ['차량구매','자동차',50000000,42000000,'2025-08-01',500000,'진행중',Utilities.getUuid()],
    ['유럽여행','여행',5000000,1200000,'2025-12-01',380000,'진행중',Utilities.getUuid()],
  ];
  s.getRange(2, 1, data.length, h.length).setValues(data);
  s.getRange(2, 3, data.length, 4).setNumberFormat('#,##0');
  s.setColumnWidths(1, 8, [140,90,140,140,110,120,70,1]);
  s.setFrozenRows(1);
  alt(s, 2, data.length, h.length);
}

function buildLedger(ss) {
  let s = ss.getSheetByName(SHEET.LEDGER);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(SHEET.LEDGER);
  s.setTabColor('#8b5cf6');
  const h = ['날짜','유형','카테고리','내용/가맹점','금액(원)','결제수단','메모','작성자','ID','생성시각'];
  s.getRange(1, 1, 1, h.length).setValues([h]);
  hdr(s.getRange(1, 1, 1, h.length));

  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['수입','지출'],true).build();
  s.getRange('B2:B5000').setDataValidation(typeRule);
  const payRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['체크카드','신용카드','현금','계좌이체','간편결제'],true).build();
  s.getRange('F2:F5000').setDataValidation(payRule);

  const now = new Date();
  const ym  = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const sample = [
    [`${ym}-01`,'수입','급여','이번 달 급여',4200000,'계좌이체','','아빠',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-03`,'지출','카페','스타벅스',6400,'신용카드','','엄마',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-05`,'지출','식비','GS25 편의점',8500,'체크카드','','아빠',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-07`,'지출','외식','가족 외식',85000,'신용카드','주말 점심','엄마',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-10`,'지출','쇼핑','쿠팡',45000,'신용카드','생활용품','엄마',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-12`,'지출','구독','넷플릭스',17000,'신용카드','','아빠',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-15`,'수입','급여','이번 달 급여',3800000,'계좌이체','','엄마',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-18`,'지출','의료','소아과 진료',15000,'체크카드','','엄마',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-20`,'지출','교육','학원비',300000,'계좌이체','','아빠',Utilities.getUuid(),new Date().toISOString()],
  ];
  s.getRange(2, 1, sample.length, h.length).setValues(sample);
  s.getRange(2, 5, sample.length, 1).setNumberFormat('#,##0');
  s.setColumnWidths(1, 10, [110,65,110,200,120,100,160,80,1,1]);
  s.setFrozenRows(1);

  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="수입"')
      .setBackground('#d1fae5').setRanges([s.getRange('A2:J5000')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="지출"')
      .setBackground('#fff1f2').setRanges([s.getRange('A2:J5000')]).build(),
  ]);
}
