// ============================================================
// WEALTHOS Family — Apps Script 통합 버전
// Code.gs 하나로 웹앱 + 백엔드 완전 통합
// GitHub Pages 불필요 · CORS 문제 없음
// ============================================================

// ── 시트 이름 상수 ────────────────────────────────────
const SH = {
  LEDGER:  '가계부',
  ASSETS:  '자산',
  GOALS:   '목표',
  SETTINGS:'설정',
  MEMBERS: '가족구성원',
  LOG:     '활동로그',
};

// ── 자동 분류 키워드 ──────────────────────────────────
const KW = {
  '식비':  ['GS25','CU','세븐일레븐','이마트','홈플러스','롯데마트','마트','슈퍼','편의점'],
  '카페':  ['스타벅스','카페','커피','메가커피','투썸','빽다방','할리스','이디야'],
  '외식':  ['배달의민족','배민','쿠팡이츠','요기요','맥도날드','버거킹','피자','치킨','음식점','식당'],
  '교통':  ['T-money','티머니','지하철','버스','택시','카카오택시','KTX','SRT'],
  '주유':  ['주유소','GS칼텍스','SK에너지','현대오일'],
  '쇼핑':  ['쿠팡','네이버쇼핑','11번가','올리브영','무신사','G마켓'],
  '구독':  ['넷플릭스','유튜브프리미엄','멜론','스포티파이','ChatGPT','Adobe'],
  '통신':  ['SKT','KT','LG유플러스','알뜰폰'],
  '의료':  ['병원','약국','의원','내과','치과','한의원'],
  '교육':  ['학원','교육','인강','도서','책'],
  '급여':  ['급여','월급','임금','봉급'],
  '부업':  ['프리랜서','강의료','외주','용돈'],
};

// ============================================================
// 웹앱 진입점 — HTML 페이지 반환
// ============================================================
function doGet(e) {
  const p = e.parameter || {};

  // API 요청 (action 파라미터 있을 때)
  if (p.action) {
    return apiHandler(p);
  }

  // 웹앱 HTML 반환
  const html = HtmlService.createHtmlOutput(getHtml())
    .setTitle('WEALTHOS Family')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  return html;
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents || '{}');
    return apiHandler(p);
  } catch(err) {
    return res({ ok: false, error: err.message });
  }
}

// ── API 응답 헬퍼 ─────────────────────────────────────
function res(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── API 라우터 ────────────────────────────────────────
function apiHandler(p) {
  try {
    switch (p.action) {
      case 'getDashboard':      return res({ ok:true, data: getDashboard() });
      case 'getLedger':         return res({ ok:true, data: getLedger(p) });
      case 'getAssets':         return res({ ok:true, data: getAssets() });
      case 'getGoals':          return res({ ok:true, data: getGoals() });
      case 'getMembers':        return res({ ok:true, data: getMembers() });
      case 'getActivity':       return res({ ok:true, data: getActivity(30) });
      case 'poll':              return res({ ok:true, data: poll(p.since||'') });
      case 'addTransaction':    return res({ ok:true, data: addTransaction(p) });
      case 'deleteTransaction': return res({ ok:true, data: deleteTransaction(p) });
      case 'addAsset':          return res({ ok:true, data: addAsset(p) });
      case 'updateAsset':       return res({ ok:true, data: updateAsset(p) });
      case 'addGoal':           return res({ ok:true, data: addGoal(p) });
      case 'updateGoal':        return res({ ok:true, data: updateGoal(p) });
      case 'addMember':         return res({ ok:true, data: addMember(p) });
      case 'saveSettings':      return res({ ok:true, data: saveSettings(p) });
      default:                  return res({ ok:true, version:'5.0' });
    }
  } catch(err) {
    return res({ ok:false, error: err.message });
  }
}

// ============================================================
// Apps Script 클라이언트 함수 (google.script.run용)
// ============================================================
function serverGetDashboard()       { return getDashboard(); }
function serverGetLedger(p)         { return getLedger(p); }
function serverGetAssets()          { return getAssets(); }
function serverGetGoals()           { return getGoals(); }
function serverGetMembers()         { return getMembers(); }
function serverGetActivity()        { return getActivity(30); }
function serverPoll(since)          { return poll(since); }
function serverAddTransaction(p)    { return addTransaction(p); }
function serverDeleteTransaction(p) { return deleteTransaction(p); }
function serverAddAsset(p)          { return addAsset(p); }
function serverUpdateAsset(p)       { return updateAsset(p); }
function serverAddGoal(p)           { return addGoal(p); }
function serverUpdateGoal(p)        { return updateGoal(p); }
function serverAddMember(p)         { return addMember(p); }
function serverSaveSettings(p)      { return saveSettings(p); }

// ============================================================
// 백엔드 로직
// ============================================================

// ── 폴링 ──────────────────────────────────────────────
function poll(since) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sd   = since ? new Date(since) : new Date(0);
  const now  = new Date();
  const txns = getLedgerSince(ss, sd);
  const logs = getLogSince(ss, sd);
  const sum  = getLedgerMonth(ss);
  return { timestamp: now.toISOString(), hasChanges: txns.length>0||logs.length>0,
    newTransactions: txns, newActivity: logs, summary: sum };
}

function getLedgerSince(ss, sd) {
  const s = ss.getSheetByName(SH.LEDGER);
  if (!s || s.getLastRow()<2) return [];
  return s.getRange(2,1,s.getLastRow()-1,10).getValues()
    .filter(r=>r[0]&&r[9]&&new Date(r[9])>sd).map(rowToTxn).reverse();
}

function getLogSince(ss, sd) {
  const s = ss.getSheetByName(SH.LOG);
  if (!s || s.getLastRow()<2) return [];
  return s.getRange(2,1,s.getLastRow()-1,4).getValues()
    .filter(r=>r[0]&&new Date(r[0])>sd)
    .map(r=>({time:new Date(r[0]).toISOString(),member:r[1],action:r[2],detail:r[3]}))
    .reverse();
}

// ── 대시보드 ──────────────────────────────────────────
function getDashboard() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const tot  = sumCol(ss, SH.ASSETS, 4);
  const debt = Number(getSettingVal(ss,'총부채')||0);
  const sum  = getLedgerMonth(ss);
  const goals   = getGoals();
  const members = getMembers();
  const activity= getActivity(15);

  const memExp = {};
  members.forEach(m => { memExp[m.name]=0; });
  const ls = ss.getSheetByName(SH.LEDGER);
  if (ls && ls.getLastRow()>1) {
    const now = new Date();
    ls.getRange(2,1,ls.getLastRow()-1,10).getValues().forEach(r=>{
      if (!r[0]) return;
      const d = new Date(r[0]);
      if (d.getFullYear()!==now.getFullYear()||d.getMonth()!==now.getMonth()) return;
      if (r[1]==='지출') memExp[r[7]||'미지정']=(memExp[r[7]]||0)+(Number(r[4])||0);
    });
  }

  return {
    totalAssets:tot, totalDebts:debt, netWorth:tot-debt,
    monthlyIncome:sum.income, monthlyExpense:sum.expense,
    savingsRate: sum.income>0 ? Math.round((sum.income-sum.expense)/sum.income*100) : 0,
    categories:sum.categories, budgets:getBudgets(ss),
    goals, members, memberExpense:memExp,
    recentTxns: getLedger({limit:'10'}),
    recentActivity: activity,
    updatedAt: new Date().toISOString(),
  };
}

// ── 가계부 ────────────────────────────────────────────
function getLedger(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.LEDGER);
  if (!s||s.getLastRow()<2) return [];
  const mon = p.month||'', mem = p.member||'', lim = Number(p.limit)||0;
  const res = s.getRange(2,1,s.getLastRow()-1,10).getValues()
    .filter(r=>{
      if (!r[0]||!r[4]) return false;
      const d = new Date(r[0]);
      const ym=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
      if (mon&&ym!==mon) return false;
      if (mem&&r[7]!==mem) return false;
      return true;
    }).map(rowToTxn).reverse();
  return lim>0 ? res.slice(0,lim) : res;
}

function rowToTxn(r) {
  const d = new Date(r[0]);
  return {
    id: r[8]||'',
    date:`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`,
    type:r[1], category:r[2], desc:r[3], amount:Number(r[4])||0,
    payMethod:r[5], memo:r[6], member:r[7],
    createdAt:r[9]?new Date(r[9]).toISOString():'',
  };
}

function addTransaction(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.LEDGER);
  if (!s) throw new Error('가계부 시트 없음 — 메뉴>초기화를 먼저 실행하세요');
  if (!p.date||!p.amount) throw new Error('날짜와 금액은 필수입니다');
  let cat = p.category||autoClassify(p.desc||'');
  if (!cat) cat = p.type==='수입'?'기타수입':'기타';
  const id = Utilities.getUuid();
  const cr = new Date().toISOString();
  const lr = Math.max(s.getLastRow(),1);
  s.getRange(lr+1,1,1,10).setValues([[
    p.date, p.type||'지출', cat, p.desc||'',
    Number(p.amount), p.payMethod||'', p.memo||'',
    p.member||'미지정', id, cr,
  ]]);
  s.getRange(lr+1,5).setNumberFormat('#,##0');
  log(ss, p.member,'거래추가',`${p.type} ${cat} ₩${Number(p.amount).toLocaleString()}`);
  return { id, category:cat, createdAt:cr };
}

function deleteTransaction(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.LEDGER);
  if (!s||s.getLastRow()<2) throw new Error('데이터 없음');
  const rows = s.getRange(2,9,s.getLastRow()-1,1).getValues();
  for (let i=0;i<rows.length;i++) {
    if (rows[i][0]===p.id) {
      s.deleteRow(i+2);
      log(ss,p.member,'거래삭제',p.desc||p.id.slice(0,8));
      return {deleted:true};
    }
  }
  throw new Error('거래를 찾을 수 없습니다');
}

// ── 자산 ──────────────────────────────────────────────
function getAssets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.ASSETS);
  if (!s||s.getLastRow()<2) return [];
  return s.getRange(2,1,s.getLastRow()-1,7).getValues().filter(r=>r[0]).map(r=>({
    id:r[6]||'', name:r[0], type:r[1], institution:r[2],
    currentValue:Number(r[3])||0, purchaseValue:Number(r[4])||0, memo:r[5],
  }));
}

function addAsset(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.ASSETS);
  if (!s) throw new Error('자산 시트 없음');
  if (!p.name||!p.currentValue) throw new Error('자산명과 금액은 필수입니다');
  const id = Utilities.getUuid();
  const lr = Math.max(s.getLastRow(),1);
  s.getRange(lr+1,1,1,7).setValues([[
    p.name, p.type||'기타', p.institution||'',
    Number(p.currentValue), Number(p.purchaseValue||p.currentValue), p.memo||'', id,
  ]]);
  s.getRange(lr+1,4,1,2).setNumberFormat('#,##0');
  log(ss,p.member,'자산추가',`${p.name} ₩${Number(p.currentValue).toLocaleString()}`);
  return {id};
}

function updateAsset(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.ASSETS);
  if (!s||s.getLastRow()<2) throw new Error('데이터 없음');
  const rows = s.getRange(2,7,s.getLastRow()-1,1).getValues();
  for (let i=0;i<rows.length;i++) {
    if (rows[i][0]===p.id) {
      if (p.currentValue) s.getRange(i+2,4).setValue(Number(p.currentValue)).setNumberFormat('#,##0');
      if (p.memo!==undefined) s.getRange(i+2,6).setValue(p.memo);
      log(ss,p.member,'자산수정',p.name||p.id.slice(0,8));
      return {updated:true};
    }
  }
  throw new Error('자산을 찾을 수 없습니다');
}

// ── 목표 ──────────────────────────────────────────────
function getGoals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.GOALS);
  if (!s||s.getLastRow()<2) return [];
  return s.getRange(2,1,s.getLastRow()-1,8).getValues().filter(r=>r[0]).map(r=>{
    const tgt=Number(r[2])||0, cur=Number(r[3])||0;
    return { id:r[7]||'', name:r[0], type:r[1], target:tgt, current:cur,
      targetDate:r[4]?fmtDate(new Date(r[4])):'',
      monthly:Number(r[5])||0, status:r[6]||'진행중',
      rate:tgt>0?Math.round(cur/tgt*100):0, remaining:Math.max(tgt-cur,0) };
  });
}

function addGoal(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.GOALS);
  if (!s) throw new Error('목표 시트 없음');
  if (!p.name||!p.target) throw new Error('목표명과 금액은 필수입니다');
  const id = Utilities.getUuid();
  const lr = Math.max(s.getLastRow(),1);
  s.getRange(lr+1,1,1,8).setValues([[
    p.name, p.type||'기타', Number(p.target), Number(p.current||0),
    p.targetDate||'', Number(p.monthly||0), '진행중', id,
  ]]);
  s.getRange(lr+1,3,1,4).setNumberFormat('#,##0');
  log(ss,p.member,'목표추가',`${p.name} ₩${Number(p.target).toLocaleString()}`);
  return {id};
}

function updateGoal(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.GOALS);
  if (!s||s.getLastRow()<2) throw new Error('데이터 없음');
  const rows = s.getRange(2,8,s.getLastRow()-1,1).getValues();
  for (let i=0;i<rows.length;i++) {
    if (rows[i][0]===p.id) {
      if (p.current!==undefined) s.getRange(i+2,4).setValue(Number(p.current)).setNumberFormat('#,##0');
      if (p.monthly!==undefined) s.getRange(i+2,6).setValue(Number(p.monthly)).setNumberFormat('#,##0');
      if (p.status!==undefined)  s.getRange(i+2,7).setValue(p.status);
      log(ss,p.member,'목표수정',p.name||p.id.slice(0,8));
      return {updated:true};
    }
  }
  throw new Error('목표를 찾을 수 없습니다');
}

// ── 구성원 ────────────────────────────────────────────
function getMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.MEMBERS);
  if (!s||s.getLastRow()<2) return [];
  return s.getRange(2,1,s.getLastRow()-1,4).getValues().filter(r=>r[0])
    .map(r=>({name:r[0],role:r[1],color:r[2],emoji:r[3]}));
}

function addMember(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.MEMBERS);
  if (!s) throw new Error('구성원 시트 없음');
  if (!p.name) throw new Error('이름은 필수입니다');
  const colors=['#4f8ef7','#f87171','#34d399','#fbbf24','#a78bfa','#fb923c'];
  const emojis=['👨','👩','👧','👦','👴','👵'];
  const idx=Math.max(s.getLastRow()-1,0);
  s.getRange(Math.max(s.getLastRow(),1)+1,1,1,4).setValues([[
    p.name, p.role||'가족',
    p.color||colors[idx%colors.length],
    p.emoji||emojis[idx%emojis.length],
  ]]);
  log(ss,'시스템','구성원추가',p.name);
  return {added:true};
}

// ── 활동 로그 ─────────────────────────────────────────
function getActivity(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.LOG);
  if (!s||s.getLastRow()<2) return [];
  const cnt = Math.min(s.getLastRow()-1,200);
  return s.getRange(s.getLastRow()-cnt+1,1,cnt,4).getValues()
    .filter(r=>r[0])
    .map(r=>({time:new Date(r[0]).toISOString(),member:r[1],action:r[2],detail:r[3]}))
    .reverse().slice(0,limit||50);
}

function log(ss, member, action, detail) {
  const s = ss.getSheetByName(SH.LOG);
  if (!s) return;
  const lr = Math.max(s.getLastRow(),1);
  s.getRange(lr+1,1,1,4).setValues([[new Date().toISOString(),member||'미지정',action,detail||'']]);
  if (s.getLastRow()>1001) s.deleteRow(2);
}

// ── 설정 ──────────────────────────────────────────────
function saveSettings(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SH.SETTINGS);
  if (!s) throw new Error('설정 시트 없음');
  const rows = s.getRange(1,1,s.getLastRow(),2).getValues();
  rows.forEach((r,i)=>{
    const key=`budget_${r[0]}`;
    if (p[key]!==undefined) s.getRange(i+1,2).setValue(Number(p[key]));
    if (r[0]==='총부채'&&p.totalDebts!==undefined) s.getRange(i+1,2).setValue(Number(p.totalDebts));
  });
  log(ss,p.member,'설정변경','예산 업데이트');
  return {saved:true};
}

// ── 내부 유틸 ─────────────────────────────────────────
function getLedgerMonth(ss) {
  const s = ss.getSheetByName(SH.LEDGER);
  const r = {income:0, expense:0, categories:{}};
  if (!s||s.getLastRow()<2) return r;
  const now = new Date();
  s.getRange(2,1,s.getLastRow()-1,5).getValues().forEach(row=>{
    if (!row[0]) return;
    const d = new Date(row[0]);
    if (d.getFullYear()!==now.getFullYear()||d.getMonth()!==now.getMonth()) return;
    const amt = Number(row[4])||0;
    if (row[1]==='수입') r.income+=amt;
    else { r.expense+=amt; r.categories[row[2]]=(r.categories[row[2]]||0)+amt; }
  });
  return r;
}

function sumCol(ss, sheetName, col) {
  const s = ss.getSheetByName(sheetName);
  if (!s||s.getLastRow()<2) return 0;
  return s.getRange(2,col,s.getLastRow()-1,1).getValues()
    .reduce((sum,r)=>sum+(Number(r[0])||0),0);
}

function getBudgets(ss) {
  const s = ss.getSheetByName(SH.SETTINGS);
  const b = {};
  if (!s) return b;
  const cats=['식비','카페','외식','교통','주유','쇼핑','구독','통신','의료','문화','교육','경조사','보험','기타'];
  s.getRange(1,1,s.getLastRow(),2).getValues().forEach(r=>{
    if (r[0]&&r[1]&&cats.includes(r[0])) b[r[0]]=Number(r[1]);
  });
  return b;
}

function getSettingVal(ss, key) {
  const s = ss.getSheetByName(SH.SETTINGS);
  if (!s) return null;
  for (const r of s.getRange(1,1,s.getLastRow(),2).getValues()) {
    if (r[0]===key) return r[1];
  }
  return null;
}

function autoClassify(desc) {
  const d = desc.toLowerCase();
  for (const [cat,kws] of Object.entries(KW)) {
    for (const kw of kws) { if (d.includes(kw.toLowerCase())) return cat; }
  }
  return '';
}

function fmtDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

// ============================================================
// 스프레드시트 메뉴
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 WEALTHOS')
    .addItem('🚀 초기화 (최초 1회)', 'initSystem')
    .addSeparator()
    .addItem('🔗 웹앱 URL 열기', 'openWebApp')
    .addItem('👨‍👩‍👧 구성원 추가', 'showAddMember')
    .addSeparator()
    .addItem('📋 이번 달 리포트', 'makeReport')
    .addToUi();
}

function openWebApp() {
  const ui = SpreadsheetApp.getUi();
  try {
    const url = ScriptApp.getService().getUrl();
    ui.alert('🔗 웹앱 URL',
      '아래 URL을 브라우저에서 열거나 가족에게 공유하세요:\n\n' + url +
      '\n\n이 URL 하나로 모든 가족이 접속할 수 있습니다!',
      ui.ButtonSet.OK);
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
  }).addMember({name,role:document.getElementById('r').value,emoji:document.getElementById('e').value});
}
</script>`).setWidth(280).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '구성원 추가');
}

function initSystem() {
  const ui  = SpreadsheetApp.getUi();
  const res2 = ui.alert('🚀 초기화', '시트를 생성합니다. 계속하시겠습니까?', ui.ButtonSet.YES_NO);
  if (res2 !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getActiveSpreadsheet().toast('⏳ 초기화 중...','WEALTHOS',30);
  buildSettings(ss); buildMembers(ss); buildLog(ss);
  buildAssets(ss); buildGoals(ss); buildLedger(ss);
  ui.alert('✅ 완료',
    '초기화 완료!\n\n다음 단계:\n' +
    '1. 배포 → 새 배포 → 웹 앱\n' +
    '   실행 주체: 나 / 액세스: 모든 사용자\n' +
    '2. 메뉴 > 🔗 웹앱 URL 열기\n' +
    '3. URL을 가족과 공유하면 끝!\n\n' +
    '메뉴 > 👨‍👩‍👧 구성원 추가로 가족 이름을 등록하세요.',
    ui.ButtonSet.OK);
}

function makeReport() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const d   = getDashboard();
  const now = new Date();
  const lbl = `${now.getFullYear()}년 ${now.getMonth()+1}월`;
  const nm  = `📋 ${lbl} 리포트`;
  let s = ss.getSheetByName(nm);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(nm);
  s.setTabColor('#f59e0b');
  const f = n => Math.round(n).toLocaleString();
  const H = (t,r) => { s.getRange(r,1,1,6).merge().setValue(t).setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold').setFontSize(13); s.setRowHeight(r,30); };
  const R = (k,v,r,bg) => { s.getRange(r,1,1,3).merge().setValue(k).setBackground(bg||'#f8fafc'); s.getRange(r,4,1,3).merge().setValue(v).setHorizontalAlignment('right').setFontWeight('bold').setBackground(bg||'#fff'); };
  let r=1;
  s.getRange(r,1,1,6).merge().setValue(`WEALTHOS Family — ${lbl} 재무 리포트`).setFontSize(17).setFontWeight('bold').setBackground('#0f172a').setFontColor('#fff').setHorizontalAlignment('center');
  s.setRowHeight(r++,44); r++;
  H('💰 순자산',r++); R('총 자산',`₩${f(d.totalAssets)}`,r++); R('총 부채',`₩${f(d.totalDebts)}`,r++); R('순 자산',`₩${f(d.netWorth)}`,r++,'#dbeafe'); r++;
  H('📊 이번 달',r++); R('수입',`₩${f(d.monthlyIncome)}`,r++,'#d1fae5'); R('지출',`₩${f(d.monthlyExpense)}`,r++,'#fee2e2'); R('저축',`₩${f(d.monthlyIncome-d.monthlyExpense)}`,r++); R('저축률',`${d.savingsRate}%`,r++); r++;
  H('👨‍👩‍👧 구성원별',r++);
  Object.entries(d.memberExpense||{}).forEach(([m,a])=>R(m,`₩${f(a)}`,r++));
  r++;
  H('🏷️ 카테고리',r++);
  Object.entries(d.categories||{}).sort((a,b)=>b[1]-a[1]).forEach(([c,a])=>{ const pct=d.monthlyExpense>0?Math.round(a/d.monthlyExpense*100):0; R(c,`₩${f(a)} (${pct}%)`,r++); });
  r++;
  H('🎯 목표',r++);
  (d.goals||[]).forEach(g=>R(g.name,`${g.rate}% (₩${f(g.current)}/₩${f(g.target)})`,r++,g.rate>=80?'#d1fae5':g.rate>=50?'#eff6ff':'#fff'));
  for(let c=1;c<=6;c++) s.setColumnWidth(c,120);
  ss.setActiveSheet(s);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 리포트 생성 완료!','WEALTHOS',3);
}

// ── 시트 빌더 ─────────────────────────────────────────
function hdr(range) { range.setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center'); }
function alt(s,start,n,cols) { for(let i=0;i<n;i++) s.getRange(start+i,1,1,cols).setBackground(i%2===0?'#fff':'#f8fafc'); }

function buildSettings(ss) {
  let s=ss.getSheetByName(SH.SETTINGS); if(s) ss.deleteSheet(s);
  s=ss.insertSheet(SH.SETTINGS); s.setTabColor('#6b7280');
  s.getRange(1,1,16,2).setValues([
    ['총부채',0],['',''],
    ['식비',300000],['카페',100000],['외식',250000],['교통',150000],
    ['주유',100000],['쇼핑',300000],['구독',50000],['통신',80000],
    ['의료',80000],['문화',150000],['교육',300000],['경조사',100000],
    ['보험',100000],['기타',200000],
  ]);
  s.setColumnWidths(1,2,[150,140]);
}

function buildMembers(ss) {
  let s=ss.getSheetByName(SH.MEMBERS); if(s) ss.deleteSheet(s);
  s=ss.insertSheet(SH.MEMBERS); s.setTabColor('#4f8ef7');
  s.getRange(1,1,1,4).setValues([['이름','역할','색상','이모지']]); hdr(s.getRange(1,1,1,4));
  s.getRange(2,1,3,4).setValues([['아빠','부모','#4f8ef7','👨'],['엄마','부모','#f87171','👩'],['자녀','자녀','#34d399','👧']]);
  s.setColumnWidths(1,4,[120,80,100,80]); s.setFrozenRows(1);
}

function buildLog(ss) {
  let s=ss.getSheetByName(SH.LOG); if(s) ss.deleteSheet(s);
  s=ss.insertSheet(SH.LOG); s.setTabColor('#6b7280');
  s.getRange(1,1,1,4).setValues([['시각','구성원','액션','상세']]); hdr(s.getRange(1,1,1,4));
  s.setColumnWidths(1,4,[200,100,120,300]); s.setFrozenRows(1);
  log(ss,'시스템','초기화','WEALTHOS Family 설치 완료');
}

function buildAssets(ss) {
  let s=ss.getSheetByName(SH.ASSETS); if(s) ss.deleteSheet(s);
  s=ss.insertSheet(SH.ASSETS); s.setTabColor('#3b82f6');
  const h=['자산명','유형','기관','평가금액(원)','취득금액(원)','메모','ID'];
  s.getRange(1,1,1,h.length).setValues([h]); hdr(s.getRange(1,1,1,h.length));
  const data=[
    ['국민은행 통장','예금/현금','국민은행',8500000,8500000,'',Utilities.getUuid()],
    ['카카오뱅크','예금/현금','카카오뱅크',5200000,5200000,'',Utilities.getUuid()],
    ['삼성전자 주식','주식/ETF','키움증권',4800000,3500000,'100주',Utilities.getUuid()],
    ['아파트','부동산','서울 마포',320000000,280000000,'34평',Utilities.getUuid()],
  ];
  s.getRange(2,1,data.length,h.length).setValues(data);
  s.getRange(2,4,data.length,2).setNumberFormat('#,##0');
  s.setColumnWidths(1,7,[160,110,130,140,130,180,1]); s.setFrozenRows(1); alt(s,2,data.length,h.length);
}

function buildGoals(ss) {
  let s=ss.getSheetByName(SH.GOALS); if(s) ss.deleteSheet(s);
  s=ss.insertSheet(SH.GOALS); s.setTabColor('#f59e0b');
  const h=['목표명','유형','목표금액(원)','현재금액(원)','목표일','월저축(원)','상태','ID'];
  s.getRange(1,1,1,h.length).setValues([h]); hdr(s.getRange(1,1,1,h.length));
  const data=[
    ['내집마련','부동산',500000000,310000000,'2027-06-01',1200000,'진행중',Utilities.getUuid()],
    ['노후자금','은퇴',300000000,84000000,'2038-01-01',800000,'진행중',Utilities.getUuid()],
    ['차량구매','자동차',50000000,42000000,'2025-08-01',500000,'진행중',Utilities.getUuid()],
  ];
  s.getRange(2,1,data.length,h.length).setValues(data);
  s.getRange(2,3,data.length,4).setNumberFormat('#,##0');
  s.setColumnWidths(1,8,[140,90,140,140,110,120,70,1]); s.setFrozenRows(1); alt(s,2,data.length,h.length);
}

function buildLedger(ss) {
  let s=ss.getSheetByName(SH.LEDGER); if(s) ss.deleteSheet(s);
  s=ss.insertSheet(SH.LEDGER); s.setTabColor('#8b5cf6');
  const h=['날짜','유형','카테고리','내용/가맹점','금액(원)','결제수단','메모','작성자','ID','생성시각'];
  s.getRange(1,1,1,h.length).setValues([h]); hdr(s.getRange(1,1,1,h.length));
  s.getRange('B2:B5000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['수입','지출'],true).build());
  s.getRange('F2:F5000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['체크카드','신용카드','현금','계좌이체','간편결제'],true).build());
  const now=new Date(), ym=`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const sample=[
    [`${ym}-01`,'수입','급여','이번 달 급여',4200000,'계좌이체','','아빠',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-03`,'지출','카페','스타벅스',6400,'신용카드','','엄마',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-05`,'지출','식비','GS25',8500,'체크카드','','아빠',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-07`,'지출','외식','가족 외식',85000,'신용카드','주말','엄마',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-10`,'지출','교육','학원비',300000,'계좌이체','','아빠',Utilities.getUuid(),new Date().toISOString()],
    [`${ym}-15`,'수입','급여','이번 달 급여',3800000,'계좌이체','','엄마',Utilities.getUuid(),new Date().toISOString()],
  ];
  s.getRange(2,1,sample.length,h.length).setValues(sample);
  s.getRange(2,5,sample.length,1).setNumberFormat('#,##0');
  s.setColumnWidths(1,10,[110,65,110,200,120,100,160,80,1,1]); s.setFrozenRows(1);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="수입"').setBackground('#d1fae5').setRanges([s.getRange('A2:J5000')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="지출"').setBackground('#fff1f2').setRanges([s.getRange('A2:J5000')]).build(),
  ]);
}

// ============================================================
// HTML 웹앱 (google.script.run으로 직접 통신 — CORS 없음)
// ============================================================
function getHtml() {
  return `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>WEALTHOS Family</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden}
:root{--bg:#0f172a;--bg2:#1e293b;--bg3:#334155;--border:#334155;--text:#f1f5f9;--muted:#94a3b8;--accent:#4f8ef7;--green:#34d399;--red:#f87171;--amber:#fbbf24;--r:12px;--rs:8px;--sb:220px}
body{font-family:-apple-system,'Apple SD Gothic Neo',sans-serif;background:var(--bg);color:var(--text)}
.app{display:flex;height:100vh;width:100vw;overflow:hidden}
.sb{width:var(--sb);flex-shrink:0;background:var(--bg2);border-right:1px solid var(--border);display:flex;flex-direction:column;height:100vh;overflow-y:auto}
.logo-w{padding:18px 16px 14px;border-bottom:1px solid var(--border)}
.logo{font-size:16px;font-weight:700;letter-spacing:-.5px}
.logo em{color:var(--accent);font-style:normal}
.logo-s{font-size:10px;color:var(--muted);margin-top:2px}
.mem-w{padding:10px 8px;border-bottom:1px solid var(--border)}
.mem-l{font-size:10px;color:var(--muted);margin-bottom:6px;padding:0 4px;text-transform:uppercase;letter-spacing:.5px;font-weight:600}
.mp{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:var(--rs);border:1px solid transparent;cursor:pointer;font-size:13px;color:var(--muted);background:transparent;width:100%;text-align:left;transition:all .15s}
.mp:hover{background:var(--bg3);color:var(--text)}
.mp.on{background:rgba(79,142,247,.15);color:var(--accent);border-color:rgba(79,142,247,.25);font-weight:600}
.mp .dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.nav-w{padding:8px;flex:1}
.ni{display:flex;align-items:center;gap:9px;padding:9px 11px;border-radius:var(--rs);font-size:13px;color:var(--muted);border:none;background:transparent;width:100%;text-align:left;transition:all .15s;margin-bottom:1px;cursor:pointer}
.ni:hover{background:var(--bg3);color:var(--text)}
.ni.on{background:rgba(79,142,247,.15);color:var(--accent);font-weight:600}
.ni .ic{font-size:15px;width:18px;text-align:center}
.sync-w{padding:8px 12px;border-top:1px solid var(--border);display:flex;align-items:center;gap:6px;font-size:11px;color:var(--muted)}
.sdot{width:7px;height:7px;border-radius:50%;background:var(--muted);flex-shrink:0;transition:background .3s}
.sdot.live{background:var(--green);animation:pulse 2s infinite}
.sdot.sync{background:var(--amber)}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}}
.main{flex:1;height:100vh;overflow-y:auto;display:flex;flex-direction:column}
.pg{display:none;flex:1;padding:20px 22px;flex-direction:column;gap:12px;min-height:100%}
.pg.on{display:flex}
.ph{display:flex;justify-content:space-between;align-items:center;flex-shrink:0}
.pt{font-size:19px;font-weight:700;letter-spacing:-.5px}
.ps{font-size:12px;color:var(--muted);margin-top:2px}
.g4{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
.g3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}
.g2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}
.g21{display:grid;grid-template-columns:1.4fr 1fr;gap:10px}
.g12{display:grid;grid-template-columns:1fr 1.4fr;gap:10px}
.card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r);padding:14px 16px}
.card.fl{flex:1;display:flex;flex-direction:column;min-height:0}
.ct{font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.7px;margin-bottom:10px}
.m{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r);padding:13px 15px}
.ml{font-size:10px;color:var(--muted);margin-bottom:5px;text-transform:uppercase;letter-spacing:.4px;font-weight:600}
.mv{font-size:19px;font-weight:700;letter-spacing:-.5px}
.mv.lg{font-size:24px}.mv.g{color:var(--green)}.mv.r{color:var(--red)}.mv.a{color:var(--accent)}
.cr{margin-bottom:9px}
.ch{display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px}
.cn{color:var(--muted)}.ca{font-weight:600}
.ctr{height:4px;background:var(--bg3);border-radius:2px}
.cf{height:100%;border-radius:2px;transition:width .5s}
.cp{font-size:10px;color:var(--muted);margin-top:2px;text-align:right}
.mbr{display:flex;align-items:center;gap:9px;padding:6px 0;border-bottom:1px solid var(--border)}
.mbr:last-child{border-bottom:none}
.mbi{flex:1}.mbn{font-size:12px;font-weight:600}
.mbtr{height:4px;background:var(--bg3);border-radius:2px;margin-top:4px}
.mbfl{height:100%;border-radius:2px}
.mba{font-size:13px;font-weight:700;flex-shrink:0}
.tl{display:flex;flex-direction:column;gap:1px}
.ti{display:flex;align-items:center;gap:10px;padding:8px 11px;border-radius:var(--rs);background:var(--bg2);transition:background .1s}
.ti:hover{background:var(--bg3)}
.ti.new{animation:sIn .35s ease}
@keyframes sIn{from{opacity:0;transform:translateY(-5px)}to{opacity:1;transform:none}}
.tic{width:32px;height:32px;border-radius:9px;background:var(--bg3);display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
.tif{flex:1;min-width:0}
.tin{font-size:13px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tic2{font-size:11px;color:var(--muted);margin-top:1px;display:flex;align-items:center;gap:5px}
.tmb{padding:1px 6px;border-radius:20px;font-size:10px;font-weight:700}
.tam{font-size:13px;font-weight:700;flex-shrink:0}
.tam.neg{color:var(--red)}.tam.pos{color:var(--green)}
.tdel{opacity:0;background:none;border:none;color:var(--red);cursor:pointer;font-size:13px;padding:3px 5px;transition:opacity .15s}
.ti:hover .tdel{opacity:1}
.feed{display:flex;flex-direction:column;gap:1px;overflow-y:auto}
.fi{display:flex;align-items:flex-start;gap:9px;padding:8px 10px;border-radius:var(--rs);background:var(--bg2);border-left:3px solid transparent}
.fi.new{animation:sIn .35s ease;border-left-color:var(--accent)}
.fav{width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
.fb{flex:1;min-width:0}.fw{font-size:12px;font-weight:700}
.fwt{font-size:11px;color:var(--muted);margin-top:1px}
.ft{font-size:10px;color:var(--muted);flex-shrink:0;margin-top:2px}
.fgr{display:flex;flex-direction:column;gap:4px}
.fl2{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.4px}
.fi2,.fs{padding:8px 11px;background:var(--bg3);border:1px solid var(--border);border-radius:var(--rs);color:var(--text);font-size:13px;font-family:inherit;transition:border-color .15s}
.fi2:focus,.fs:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px rgba(79,142,247,.12)}
.fs option{background:var(--bg2)}
.tt{display:flex;border:1px solid var(--border);border-radius:var(--rs);overflow:hidden}
.tb{flex:1;padding:8px;border:none;background:transparent;color:var(--muted);font-size:13px;font-weight:500;cursor:pointer;transition:all .15s}
.tb.ae{background:rgba(248,113,113,.15);color:var(--red)}
.tb.ai{background:rgba(52,211,153,.15);color:var(--green)}
.btn{padding:8px 16px;border-radius:var(--rs);font-size:13px;font-weight:600;border:none;cursor:pointer;transition:all .15s;font-family:inherit}
.bp{background:var(--accent);color:#fff}.bp:hover{background:#3b82f6}
.bs{background:var(--bg3);color:var(--text);border:1px solid var(--border)}.bs:hover{background:var(--bg2)}
.bsm{padding:5px 12px;font-size:12px}
.badge{font-size:10px;padding:2px 7px;border-radius:20px;font-weight:700}
.ba{background:rgba(79,142,247,.15);color:var(--accent)}
.ab{position:fixed;top:14px;left:50%;transform:translateX(-50%) translateY(-80px);background:var(--bg2);border:1px solid var(--accent);border-radius:24px;padding:9px 18px;font-size:13px;font-weight:500;display:flex;align-items:center;gap:8px;box-shadow:0 8px 32px rgba(0,0,0,.5);transition:transform .4s cubic-bezier(.34,1.56,.64,1);z-index:500;white-space:nowrap}
.ab.on{transform:translateX(-50%) translateY(0)}
.toast{position:fixed;bottom:20px;right:20px;background:var(--bg2);color:var(--text);border:1px solid var(--border);border-radius:10px;padding:10px 16px;font-size:13px;font-weight:500;box-shadow:0 8px 24px rgba(0,0,0,.4);transform:translateY(60px);opacity:0;transition:all .3s cubic-bezier(.34,1.56,.64,1);z-index:999}
.toast.on{transform:translateY(0);opacity:1}
.toast.ok{border-color:var(--green)}.toast.err{border-color:var(--red)}
.mod-bg{position:fixed;inset:0;background:rgba(0,0,0,.65);backdrop-filter:blur(4px);z-index:100;display:flex;align-items:center;justify-content:center;opacity:0;pointer-events:none;transition:opacity .2s}
.mod-bg.on{opacity:1;pointer-events:all}
.mod{background:var(--bg2);border:1px solid var(--border);border-radius:16px;padding:22px;width:460px;max-width:96vw;max-height:90vh;overflow-y:auto;transform:scale(.95);transition:transform .2s}
.mod-bg.on .mod{transform:scale(1)}
.mot{font-size:15px;font-weight:700;margin-bottom:16px}
.moa{display:flex;gap:8px;margin-top:16px;justify-content:flex-end}
.warn-b{background:rgba(251,191,36,.08);border:1px solid rgba(251,191,36,.3);border-radius:var(--rs);padding:10px 14px;font-size:12px;color:var(--amber);display:flex;align-items:center;gap:8px}
.empty{padding:24px;text-align:center;color:var(--muted);font-size:13px;line-height:1.8}
.empty .ei{font-size:26px;margin-bottom:8px}
.spin{display:inline-block;width:16px;height:16px;border:2px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin .8s linear infinite;vertical-align:middle;margin-right:6px}
@keyframes spin{to{transform:rotate(360deg)}}
@media(max-width:900px){:root{--sb:60px}.logo-s,.mem-l,.mp span,.ni span{display:none}.mp{justify-content:center;padding:10px}.ni{justify-content:center;padding:10px}.g4{grid-template-columns:1fr 1fr}.g21,.g12{grid-template-columns:1fr}.pg{padding:14px 12px}}
@media(max-width:600px){.g4,.g3{grid-template-columns:1fr 1fr}}
</style>
</head>
<body>
<div class="app">
<nav class="sb">
  <div class="logo-w"><div class="logo">WEALTH<em>OS</em></div><div class="logo-s">가족 공유 재산관리</div></div>
  <div class="mem-w"><div class="mem-l">지금 나는</div><div id="mem-pills"><div class="empty" style="padding:8px;font-size:12px">로딩 중...</div></div></div>
  <div class="nav-w">
    <button class="ni on" onclick="go('dash')"><span class="ic">📊</span><span>대시보드</span></button>
    <button class="ni" onclick="go('ledger')"><span class="ic">📒</span><span>가계부</span></button>
    <button class="ni" onclick="go('assets')"><span class="ic">🏦</span><span>자산</span></button>
    <button class="ni" onclick="go('goals')"><span class="ic">🎯</span><span>목표</span></button>
    <button class="ni" onclick="go('activity')"><span class="ic">🔔</span><span>활동</span></button>
    <button class="ni" onclick="go('settings')"><span class="ic">⚙️</span><span>설정</span></button>
  </div>
  <div class="sync-w"><div class="sdot sync" id="sdot"></div><span id="slbl">로딩 중...</span></div>
</nav>
<main class="main">
<section id="pg-dash" class="pg on">
  <div class="ph">
    <div><div class="pt">가족 대시보드</div><div class="ps" id="dash-time"><span class="spin"></span>데이터 로딩 중...</div></div>
    <button class="btn bp" onclick="openTxn()">➕ 거래 입력</button>
  </div>
  <div id="warn-area"></div>
  <div class="g4">
    <div class="m"><div class="ml">순 자산</div><div class="mv lg a" id="m-nw">—</div></div>
    <div class="m"><div class="ml">이번 달 수입</div><div class="mv g" id="m-inc">—</div></div>
    <div class="m"><div class="ml">이번 달 지출</div><div class="mv r" id="m-exp">—</div></div>
    <div class="m"><div class="ml">저축률</div><div class="mv" id="m-sv">—</div></div>
  </div>
  <div class="g21" style="flex:1;min-height:0">
    <div class="card fl"><div class="ct">이번 달 지출</div><div id="cat-bars" style="flex:1;overflow-y:auto"></div></div>
    <div style="display:flex;flex-direction:column;gap:10px;min-height:0">
      <div class="card" style="flex:1"><div class="ct">구성원별 지출</div><div id="mem-bars"></div></div>
      <div class="card" style="flex:1">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
          <div class="ct" style="margin-bottom:0">가족 목표</div>
          <button class="btn bs bsm" onclick="go('goals')">전체</button>
        </div>
        <div id="dash-goals"></div>
      </div>
    </div>
  </div>
  <div class="g12" style="flex:1;min-height:0">
    <div class="card fl">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
        <div class="ct" style="margin-bottom:0">최근 거래</div>
        <button class="btn bs bsm" onclick="go('ledger')">전체</button>
      </div>
      <div class="tl" id="dash-txns" style="flex:1;overflow-y:auto"></div>
    </div>
    <div class="card fl">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
        <div class="ct" style="margin-bottom:0">실시간 활동</div>
        <span class="badge ba" id="feed-badge">—</span>
      </div>
      <div class="feed" id="live-feed" style="flex:1"></div>
    </div>
  </div>
</section>
<section id="pg-ledger" class="pg">
  <div class="ph">
    <div><div class="pt">가계부</div><div class="ps">가족 전체 거래 내역</div></div>
    <div style="display:flex;gap:8px;align-items:center">
      <div style="display:flex;align-items:center;gap:4px;background:var(--bg2);border:1px solid var(--border);border-radius:var(--rs);padding:2px">
        <button class="btn bs" onclick="shiftMon(-1)" style="padding:5px 10px;font-size:16px;border:none;background:transparent">‹</button>
        <span id="mon-lbl" style="font-size:13px;font-weight:600;min-width:90px;text-align:center;color:var(--text)"></span>
        <button class="btn bs" id="mon-nxt" onclick="shiftMon(1)" style="padding:5px 10px;font-size:16px;border:none;background:transparent">›</button>
      </div>
      <button class="btn bp" onclick="openTxn()">➕ 거래 입력</button>
    </div>
  </div>
  <div class="g3">
    <div class="m"><div class="ml">수입</div><div class="mv g" id="ls-inc">—</div></div>
    <div class="m"><div class="ml">지출</div><div class="mv r" id="ls-exp">—</div></div>
    <div class="m"><div class="ml">잔액</div><div class="mv" id="ls-bal">—</div></div>
  </div>
  <div class="card fl">
    <div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap">
      <select class="fs" id="f-type" style="width:85px" onchange="filterLed()"><option value="">전체</option><option value="수입">수입</option><option value="지출">지출</option></select>
      <select class="fs" id="f-mem" style="width:105px" onchange="filterLed()"><option value="">전체 구성원</option></select>
      <select class="fs" id="f-cat" style="width:115px" onchange="filterLed()"><option value="">전체 카테고리</option></select>
      <input class="fi2" id="f-q" placeholder="검색..." style="width:130px" oninput="filterLed()">
    </div>
    <div class="tl" id="led-list" style="flex:1;overflow-y:auto"></div>
  </div>
</section>
<section id="pg-assets" class="pg">
  <div class="ph"><div><div class="pt">자산 현황</div><div class="ps">가족 공동 자산 관리</div></div><button class="btn bp" onclick="openAsset()">➕ 자산 추가</button></div>
  <div class="g3">
    <div class="m"><div class="ml">총 자산</div><div class="mv g" id="a-tot">—</div></div>
    <div class="m"><div class="ml">수익금</div><div class="mv" id="a-prf">—</div></div>
    <div class="m"><div class="ml">평균 수익률</div><div class="mv a" id="a-rt">—</div></div>
  </div>
  <div class="card fl"><div id="asset-list" style="flex:1;overflow-y:auto"></div></div>
</section>
<section id="pg-goals" class="pg">
  <div class="ph"><div><div class="pt">가족 목표</div><div class="ps">함께 이루는 재무 목표</div></div><button class="btn bp" onclick="openGoal()">➕ 목표 추가</button></div>
  <div class="g2" id="goals-grid" style="flex:1;overflow-y:auto;align-content:start"></div>
</section>
<section id="pg-activity" class="pg">
  <div class="ph"><div class="pt">활동 로그</div><div class="ps">가족 모두의 재무 활동 기록</div></div>
  <div class="card fl"><div class="feed" id="full-feed" style="flex:1"></div></div>
</section>
<section id="pg-settings" class="pg">
  <div class="ph"><div class="pt">설정</div><div class="ps">예산 및 구성원 관리</div></div>
  <div style="background:rgba(79,142,247,.08);border:1px solid rgba(79,142,247,.2);border-radius:var(--r);padding:14px 16px;font-size:13px;color:var(--muted);line-height:1.75">
    <strong style="color:var(--accent)">✅ Google Sheets 직접 연동</strong><br>
    이 웹앱은 Google Apps Script로 구동되어 CORS 문제 없이 Sheets와 직접 통신합니다.<br>
    별도 URL 설정이 필요 없습니다. 데이터는 스프레드시트에 자동 저장됩니다.
  </div>
  <div class="card fl">
    <div class="ct">월별 예산 설정 (원)</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px;align-content:start;overflow-y:auto;flex:1" id="bgt-form"></div>
    <div style="margin-top:12px;text-align:right;flex-shrink:0"><button class="btn bp" onclick="saveBudgets()">예산 저장</button></div>
  </div>
  <div class="card">
    <div class="ct">구성원 관리</div>
    <div id="mem-list" style="margin-bottom:12px"></div>
    <div style="display:flex;gap:8px;flex-wrap:wrap">
      <input class="fi2" id="new-mem-name" placeholder="이름 (예: 아빠)" style="flex:1;min-width:100px">
      <select class="fs" id="new-mem-emoji" style="width:70px"><option>👨</option><option>👩</option><option>👧</option><option>👦</option><option>👴</option><option>👵</option></select>
      <button class="btn bp" onclick="addMemUI()">추가</button>
    </div>
  </div>
</section>
</main>
</div>
<div class="ab" id="ab"><span id="ab-icon">🔔</span><span id="ab-msg"></span></div>
<!-- 거래 모달 -->
<div class="mod-bg" id="mod-txn">
  <div class="mod">
    <div class="mot">➕ 거래 입력</div>
    <div class="fgr" style="margin-bottom:11px"><label class="fl2">누가</label><select class="fs" id="t-who"></select></div>
    <div class="fgr" style="margin-bottom:11px">
      <label class="fl2">유형</label>
      <div class="tt"><button class="tb ae" id="tb-e" onclick="setType('지출')">지출</button><button class="tb" id="tb-i" onclick="setType('수입')">수입</button></div>
      <input type="hidden" id="t-type" value="지출">
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
      <div class="fgr"><label class="fl2">날짜</label><input class="fi2" type="date" id="t-date"></div>
      <div class="fgr"><label class="fl2">금액(원)</label><input class="fi2" type="number" id="t-amt" placeholder="6400" min="1"></div>
      <div class="fgr"><label class="fl2">카테고리</label><select class="fs" id="t-cat"></select></div>
      <div class="fgr"><label class="fl2">결제수단</label><select class="fs" id="t-pay"><option>체크카드</option><option>신용카드</option><option>현금</option><option>계좌이체</option><option>간편결제</option></select></div>
      <div class="fgr" style="grid-column:1/-1"><label class="fl2">내용/가맹점</label><input class="fi2" id="t-desc" placeholder="예: 스타벅스 강남점"></div>
      <div class="fgr" style="grid-column:1/-1"><label class="fl2">메모 (선택)</label><input class="fi2" id="t-memo" placeholder="선택사항"></div>
    </div>
    <div class="moa"><button class="btn bs" onclick="closeM('mod-txn')">취소</button><button class="btn bp" id="t-btn" onclick="submitTxn()">저장</button></div>
  </div>
</div>
<!-- 자산 모달 -->
<div class="mod-bg" id="mod-asset">
  <div class="mod">
    <div class="mot">🏦 자산 추가/수정</div>
    <div style="display:flex;flex-direction:column;gap:12px">
      <div class="fgr"><label class="fl2">자산명</label><input class="fi2" id="a-nm" placeholder="예: 국민은행 통장"></div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div class="fgr"><label class="fl2">유형</label><select class="fs" id="a-type"><option>예금/현금</option><option>주식/ETF</option><option>부동산</option><option>자동차</option><option>적금/보험</option><option>기타</option></select></div>
        <div class="fgr"><label class="fl2">기관</label><input class="fi2" id="a-inst" placeholder="예: 국민은행"></div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div class="fgr"><label class="fl2">현재 평가금액 (원)</label><input class="fi2" type="number" id="a-val" placeholder="8500000"></div>
        <div class="fgr"><label class="fl2">취득금액 (원)</label><input class="fi2" type="number" id="a-cost" placeholder="8500000"></div>
      </div>
      <div class="fgr"><label class="fl2">메모 (선택)</label><input class="fi2" id="a-memo" placeholder="선택사항"></div>
      <input type="hidden" id="a-id">
    </div>
    <div class="moa"><button class="btn bs" onclick="closeM('mod-asset')">취소</button><button class="btn bp" onclick="submitAsset()">저장</button></div>
  </div>
</div>
<!-- 목표 모달 -->
<div class="mod-bg" id="mod-goal">
  <div class="mod">
    <div class="mot">🎯 목표 추가/수정</div>
    <div style="display:flex;flex-direction:column;gap:12px">
      <div class="fgr"><label class="fl2">목표명</label><input class="fi2" id="g-nm" placeholder="예: 내집마련"></div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div class="fgr"><label class="fl2">유형</label><select class="fs" id="g-type"><option>부동산</option><option>은퇴</option><option>자동차</option><option>여행</option><option>비상금</option><option>교육</option><option>기타</option></select></div>
        <div class="fgr"><label class="fl2">목표일</label><input class="fi2" type="date" id="g-date"></div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div class="fgr"><label class="fl2">목표금액 (원)</label><input class="fi2" type="number" id="g-tgt" placeholder="500000000"></div>
        <div class="fgr"><label class="fl2">현재금액 (원)</label><input class="fi2" type="number" id="g-cur" placeholder="100000000"></div>
      </div>
      <div class="fgr"><label class="fl2">월 저축 계획 (원)</label><input class="fi2" type="number" id="g-mon" placeholder="1000000"></div>
      <input type="hidden" id="g-id">
    </div>
    <div class="moa"><button class="btn bs" onclick="closeM('mod-goal')">취소</button><button class="btn bp" onclick="submitGoal()">저장</button></div>
  </div>
</div>
<div class="toast" id="toast"></div>
<script>
const ICONS={식비:'🛒',카페:'☕',외식:'🍜',교통:'🚇',주유:'⛽',쇼핑:'🛍️',구독:'📺',통신:'📱',의료:'💊',문화:'🎬',교육:'📚',경조사:'🎁',보험:'🛡️',기타:'📌',급여:'💰',부업:'💼',투자수익:'📈',임대:'🏠',기타수입:'💵'};
const ECATS=['식비','카페','외식','교통','주유','쇼핑','구독','통신','의료','문화','교육','경조사','보험','기타'];
const ICATS=['급여','부업','투자수익','임대','기타수입'];
const CLRS=['#4f8ef7','#f87171','#34d399','#fbbf24','#a78bfa','#fb923c'];
const MK='wo_me', BK='wo_budgets';
const S={me:localStorage.getItem(MK)||'아빠',members:[],ledger:[],assets:[],goals:[],budgets:{},pts:new Date(0).toISOString(),timer:null};

// ── google.script.run 래퍼 ──────────────────────────
function srv(fnName, arg) {
  return new Promise((resolve, reject) => {
    const runner = google.script.run
      .withSuccessHandler(resolve)
      .withFailureHandler(e => reject(new Error(e.message||String(e))));
    if (arg !== undefined) runner[fnName](arg);
    else runner[fnName]();
  });
}

// ── 폴링 ──────────────────────────────────────────
function startPoll() {
  if (S.timer) clearInterval(S.timer);
  doPoll();
  S.timer = setInterval(doPoll, 30000);
}
async function doPoll() {
  sync('sync');
  try {
    const d = await srv('serverPoll', S.pts);
    S.pts = d.timestamp;
    sync('live');
    if (d.hasChanges) {
      (d.newTransactions||[]).forEach(t=>{
        if(t.member!==S.me){alert2(t.member+' · '+t.category+' '+fmt(t.amount)+'원','🔔');addTxnRow(t,'dash-txns');if(active('ledger'))loadLedger();}
      });
      (d.newActivity||[]).forEach(a=>addFeedRow(a,'live-feed'));
      if(d.summary){$('m-inc').textContent=fmt(d.summary.income);$('m-exp').textContent=fmt(d.summary.expense);$('m-sv').textContent=d.summary.income>0?Math.round((d.summary.income-d.summary.expense)/d.summary.income*100)+'%':'—';}
    }
  } catch(e) { sync('off'); }
}
function sync(s){
  $('sdot').className='sdot'+(s==='live'?' live':s==='sync'?' sync':'');
  $('slbl').textContent={live:'실시간 동기화',sync:'동기화 중...',off:'오류'}[s]||'대기중';
}

// ── 네비 ──────────────────────────────────────────
function go(pg){
  document.querySelectorAll('.pg').forEach(e=>e.classList.remove('on'));
  document.querySelectorAll('.ni').forEach(e=>e.classList.remove('on'));
  $('pg-'+pg).classList.add('on');
  document.querySelectorAll('.ni').forEach(e=>{if(e.getAttribute('onclick')?.includes("'"+pg+"'"))e.classList.add('on');});
  ({dash:loadDash,ledger:loadLedger,assets:loadAssets,goals:loadGoals,activity:loadActivity,settings:loadSettings})[pg]?.();
}
function active(pg){return !!$('pg-'+pg)?.classList.contains('on');}

// ── 구성원 ────────────────────────────────────────
function initMems(list){
  S.members=(list&&list.length)?list:[{name:'아빠',color:'#4f8ef7',emoji:'👨'},{name:'엄마',color:'#f87171',emoji:'👩'},{name:'자녀',color:'#34d399',emoji:'👧'}];
  if(!S.me||!S.members.find(m=>m.name===S.me)){S.me=S.members[0].name;localStorage.setItem(MK,S.me);}
  $('mem-pills').innerHTML=S.members.map(m=>'<button class="mp '+(m.name===S.me?'on':'')+'" onclick="pickMe(\''+esc(m.name)+'\',this)"><div class="dot" style="background:'+(m.color||'#94a3b8')+'"></div><span>'+(m.emoji||'👤')+' '+m.name+'</span></button>').join('');
  syncSels();
}
function pickMe(name,el){S.me=name;localStorage.setItem(MK,name);document.querySelectorAll('.mp').forEach(b=>b.classList.remove('on'));el.classList.add('on');syncSels();}
function syncSels(){
  const sel=$('t-who');
  if(sel)sel.innerHTML=S.members.map(m=>'<option value="'+m.name+'" '+(m.name===S.me?'selected':'')+'>'+( m.emoji||'')+' '+m.name+'</option>').join('');
  const fs=$('f-mem');
  if(fs){const cur=fs.value;fs.innerHTML='<option value="">전체 구성원</option>'+S.members.map(m=>'<option value="'+m.name+'">'+(m.emoji||'')+' '+m.name+'</option>').join('');if(cur)fs.value=cur;}
}

// ── 대시보드 ──────────────────────────────────────
async function loadDash(){
  $('dash-time').innerHTML='<span class="spin"></span>데이터 로딩 중...';
  try{
    const d=await srv('serverGetDashboard');
    initMems(d.members||null);
    $('m-nw').textContent=fmt(d.netWorth);$('m-inc').textContent=fmt(d.monthlyIncome);$('m-exp').textContent=fmt(d.monthlyExpense);$('m-sv').textContent=(d.savingsRate||0)+'%';
    $('dash-time').textContent='업데이트: '+rtime(d.updatedAt);
    $('warn-area').innerHTML='';
    // 카테고리 바
    const cats=Object.entries(d.categories||{}).sort((a,b)=>b[1]-a[1]).slice(0,6);
    const mx=cats[0]?.[1]||1;
    $('cat-bars').innerHTML=cats.length?cats.map(([c,a],i)=>{const bgt=(d.budgets||{})[c]||0,ov=bgt&&a>bgt;return '<div class="cr"><div class="ch"><span class="cn">'+(ICONS[c]||'•')+' '+c+'</span><span class="ca" style="color:'+(ov?'var(--red)':'var(--text)')+'">'+fmt(a)+'</span></div><div class="ctr"><div class="cf" style="width:'+(a/mx*100)+'%;background:'+(ov?'var(--red)':CLRS[i%6])+'"></div></div>'+(bgt?'<div class="cp">예산 '+Math.round(a/bgt*100)+'%</div>':'')+'</div>';}).join(''):'<div class="empty"><div class="ei">📊</div>이번 달 지출 없음</div>';
    // 구성원 바
    const exp=d.memberExpense||{},tot=Math.max(Object.values(exp).reduce((s,v)=>s+v,0),1);
    $('mem-bars').innerHTML=S.members.map(m=>{const a=exp[m.name]||0,c=m.color||'#94a3b8';return '<div class="mbr"><span style="font-size:16px;flex-shrink:0">'+(m.emoji||'👤')+'</span><div class="mbi"><div class="mbn">'+m.name+'</div><div class="mbtr"><div class="mbfl" style="width:'+(a/tot*100)+'%;background:'+c+'"></div></div></div><div class="mba">'+fmt(a)+'</div></div>';}).join('');
    // 목표
    $('dash-goals').innerHTML=(d.goals||[]).slice(0,3).map(g=>{const c=g.rate>=80?'var(--green)':g.rate>=50?'var(--accent)':'var(--amber)';return '<div style="margin-bottom:10px"><div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px"><span style="font-weight:600">'+g.name+'</span><span style="color:'+c+';font-weight:700">'+g.rate+'%</span></div><div style="height:4px;background:var(--bg3);border-radius:2px"><div style="height:100%;width:'+g.rate+'%;background:'+c+';border-radius:2px"></div></div></div>';}).join('')||'<div style="color:var(--muted);font-size:12px">목표를 추가해보세요</div>';
    // 최근 거래
    const txns=d.recentTxns||[];
    $('dash-txns').innerHTML=txns.length?txns.map(t=>txnRow(t,false)).join(''):'<div class="empty"><div class="ei">📒</div>아직 거래 내역이 없습니다<br><small>➕ 거래 입력 버튼으로 추가해보세요</small></div>';
    // 피드
    renderFeed(d.recentActivity||[],'live-feed');
  }catch(e){$('dash-time').textContent='';$('warn-area').innerHTML='<div class="warn-b">⚠️ 로딩 오류: '+e.message+'</div>';}
}

// ── 가계부 ────────────────────────────────────────
let CUR_MON=(()=>{const n=new Date();return n.getFullYear()+'-'+String(n.getMonth()+1).padStart(2,'0');})();
function initMonLbl(){
  const [y,m]=CUR_MON.split('-').map(Number);
  $('mon-lbl').textContent=y+'년 '+m+'월';
  const now=new Date(),ym=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');
  $('mon-nxt').style.opacity=CUR_MON>=ym?'0.3':'1';
}
function shiftMon(d){
  const [y,m]=CUR_MON.split('-').map(Number);
  const nd=new Date(y,m-1+d,1);
  if(nd>new Date()&&d>0)return;
  CUR_MON=nd.getFullYear()+'-'+String(nd.getMonth()+1).padStart(2,'0');
  initMonLbl();loadLedger();
}
async function loadLedger(){
  initMonLbl();
  $('led-list').innerHTML='<div class="empty"><span class="spin"></span>로딩 중...</div>';
  try{
    const d=await srv('serverGetLedger',{month:CUR_MON});
    S.ledger=Array.isArray(d)?d:[];
    const inc=S.ledger.filter(t=>t.type==='수입').reduce((s,t)=>s+t.amount,0);
    const exp=S.ledger.filter(t=>t.type!=='수입').reduce((s,t)=>s+t.amount,0);
    const bal=inc-exp;
    $('ls-inc').textContent=fmt(inc);$('ls-exp').textContent=fmt(exp);
    $('ls-bal').textContent=fmt(bal);$('ls-bal').className='mv '+(bal>=0?'g':'r');
    const cats=[...new Set(S.ledger.map(t=>t.category).filter(Boolean))];
    $('f-cat').innerHTML='<option value="">전체 카테고리</option>'+cats.map(c=>'<option>'+c+'</option>').join('');
    syncSels();filterLed();
  }catch(e){$('led-list').innerHTML='<div class="empty"><div class="ei">⚠️</div>'+e.message+'</div>';}
}
function filterLed(){
  const type=gv('f-type'),mem=gv('f-mem'),cat=gv('f-cat'),q=gv('f-q').toLowerCase();
  const res=S.ledger.filter(t=>(!type||t.type===type)&&(!mem||t.member===mem)&&(!cat||t.category===cat)&&(!q||t.desc?.toLowerCase().includes(q)||t.category?.toLowerCase().includes(q)));
  $('led-list').innerHTML=res.length?res.map(t=>txnRow(t,true)).join(''):'<div class="empty">조건에 맞는 거래가 없습니다</div>';
}

// ── 자산 ──────────────────────────────────────────
async function loadAssets(){
  $('asset-list').innerHTML='<div class="empty"><span class="spin"></span>로딩 중...</div>';
  try{
    const d=await srv('serverGetAssets');
    S.assets=Array.isArray(d)?d:[];
    const tot=S.assets.reduce((s,a)=>s+a.currentValue,0);
    const cost=S.assets.reduce((s,a)=>s+(a.purchaseValue||a.currentValue),0);
    const prf=tot-cost;
    $('a-tot').textContent=fmt(tot);$('a-prf').textContent=(prf>=0?'+':'')+fmt(prf);$('a-prf').className='mv '+(prf>=0?'g':'r');$('a-rt').textContent=cost>0?((prf/cost)*100).toFixed(1)+'%':'—';
    const byT={};S.assets.forEach(a=>(byT[a.type]=byT[a.type]||[]).push(a));
    $('asset-list').innerHTML=S.assets.length?Object.entries(byT).map(([t,items])=>'<div style="margin-bottom:16px"><div style="font-size:10px;font-weight:700;color:var(--muted);margin-bottom:7px;text-transform:uppercase;letter-spacing:.5px">'+t+'</div>'+items.map(a=>{const r=a.purchaseValue>0?((a.currentValue-a.purchaseValue)/a.purchaseValue*100).toFixed(1):null;return '<div class="ti" onclick="openEditAsset(\''+esc(a.id)+'\',\''+esc(a.name)+'\',\''+esc(a.type)+'\',\''+esc(a.institution||'')+'\','+a.currentValue+','+a.purchaseValue+',\''+esc(a.memo||'\'')+'\')" ><div class="tic">'+aIcon(a.type)+'</div><div class="tif"><div class="tin">'+a.name+'</div><div class="tic2">'+(a.institution||'-')+(a.memo?' · '+a.memo:'')+'</div></div><div style="text-align:right"><div style="font-weight:700;font-size:13px">'+fmt(a.currentValue)+'</div>'+(r!==null?'<div style="font-size:11px;color:'+(Number(r)>=0?'var(--green)':'var(--red)')+'>'+(Number(r)>=0?'+':'')+r+'%</div>':'')+'</div></div>';}).join('')+'</div>').join(''):'<div class="empty"><div class="ei">🏦</div>자산을 추가해보세요</div>';
  }catch(e){$('asset-list').innerHTML='<div class="empty"><div class="ei">⚠️</div>'+e.message+'</div>';}
}

// ── 목표 ──────────────────────────────────────────
async function loadGoals(){
  $('goals-grid').innerHTML='<div class="empty"><span class="spin"></span>로딩 중...</div>';
  try{
    const d=await srv('serverGetGoals');
    S.goals=Array.isArray(d)?d:[];
    $('goals-grid').innerHTML=S.goals.length?S.goals.map(g=>{
      const c=g.rate>=80?'var(--green)':g.rate>=50?'var(--accent)':'var(--amber)';
      const ml=g.targetDate?Math.max(0,Math.round((new Date(g.targetDate)-new Date())/(1000*60*60*24*30))):null;
      return '<div style="background:var(--bg2);border:1px solid var(--border);border-radius:var(--r);padding:14px 15px;cursor:pointer;transition:border-color .15s" onmouseenter="this.style.borderColor=\'var(--accent)\'" onmouseleave="this.style.borderColor=\'var(--border)\'" onclick="openEditGoal(\''+esc(g.id)+'\',\''+esc(g.name)+'\',\''+esc(g.type)+'\','+g.target+','+g.current+',\''+(g.targetDate||'')+'\','+g.monthly+')"><div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px"><div style="font-size:13px;font-weight:600">'+g.name+'</div><span class="badge ba">'+g.type+'</span></div><div style="display:flex;justify-content:space-between;font-size:11px;color:var(--muted);margin-bottom:4px"><span>'+fmt(g.current)+'</span><span>'+fmt(g.target)+'</span></div><div style="height:5px;background:var(--bg3);border-radius:3px;margin-bottom:6px"><div style="height:100%;width:'+Math.min(g.rate,100)+'%;background:'+c+';border-radius:3px"></div></div><div style="display:flex;justify-content:space-between;font-size:11px;color:var(--muted)"><span>남은 '+fmt(g.remaining)+'</span>'+(ml!==null?'<span>목표까지 '+ml+'개월</span>':'')+'</div><div style="font-size:22px;font-weight:700;color:'+c+';text-align:right;margin-top:6px">'+g.rate+'%</div></div>';
    }).join(''):'<div class="empty" style="grid-column:1/-1"><div class="ei">🎯</div>목표를 추가해보세요</div>';
  }catch(e){$('goals-grid').innerHTML='<div class="empty" style="grid-column:1/-1"><div class="ei">⚠️</div>'+e.message+'</div>';}
}

// ── 활동 ──────────────────────────────────────────
async function loadActivity(){
  try{const d=await srv('serverGetActivity');renderFeed(Array.isArray(d)?d:[],'full-feed');}
  catch(e){$('full-feed').innerHTML='<div class="empty"><div class="ei">⚠️</div>'+e.message+'</div>';}
}

function renderFeed(items,id){
  const el=$(id);if(!el)return;
  el.innerHTML=items.length?items.map(a=>feedRow(a)).join(''):'<div class="empty">활동 없음</div>';
  const b=$('feed-badge');if(b&&items.length)b.textContent=items.length+'건';
}
function addFeedRow(a,id){const el=$(id);if(!el)return;const div=document.createElement('div');div.innerHTML=feedRow(a);const ch=div.firstElementChild;if(ch){ch.classList.add('new');el.prepend(ch);}while(el.children.length>15)el.removeChild(el.lastChild);}
function feedRow(a){
  const m=S.members.find(mb=>mb.name===a.member);const c=m?.color||'#94a3b8';
  const L={거래추가:'거래 기록',거래삭제:'거래 삭제',자산추가:'자산 추가',자산수정:'자산 수정',목표추가:'목표 설정',목표수정:'목표 수정',설정변경:'설정 변경',초기화:'시스템 초기화',구성원추가:'구성원 추가'};
  return '<div class="fi"><div class="fav" style="background:'+c+'20;color:'+c+'">'+(m?.emoji||'👤')+'</div><div class="fb"><div class="fw" style="color:'+c+'">'+a.member+'</div><div class="fwt">'+(L[a.action]||a.action)+' · '+(a.detail||'')+'</div></div><div class="ft">'+rtime(a.time)+'</div></div>';
}

// ── 설정 ──────────────────────────────────────────
function loadSettings(){
  try{const b=localStorage.getItem(BK);if(b)S.budgets=JSON.parse(b);}catch(e){}
  $('bgt-form').innerHTML=ECATS.map(c=>'<div class="fgr"><label class="fl2">'+(ICONS[c]||'•')+' '+c+'</label><input class="fi2" type="number" id="bgt-'+c+'" value="'+(S.budgets[c]||300000)+'"></div>').join('');
  renderMemList();
}
function renderMemList(){
  $('mem-list').innerHTML=S.members.map(m=>'<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid var(--border)"><span style="font-size:16px">'+(m.emoji||'👤')+'</span><span style="flex:1;font-size:13px;font-weight:500">'+m.name+'</span><span class="badge ba">'+m.role+'</span></div>').join('');
}
async function saveBudgets(){
  ECATS.forEach(c=>{const v=document.getElementById('bgt-'+c)?.value;if(v!==undefined&&v!=='')S.budgets[c]=Number(v);});
  localStorage.setItem(BK,JSON.stringify(S.budgets));
  try{const p={action:'saveSettings',member:S.me};ECATS.forEach(c=>{if(S.budgets[c])p['budget_'+c]=S.budgets[c];});await srv('serverSaveSettings',p);toast('✅ 예산 저장됐습니다!');}
  catch(e){toast('✅ 로컬 저장됐습니다!');}
}
async function addMemUI(){
  const name=$('new-mem-name').value.trim();
  if(!name){toast('이름을 입력하세요','err');return;}
  try{
    await srv('serverAddMember',{name,emoji:gv('new-mem-emoji'),role:'가족'});
    $('new-mem-name').value='';
    toast('✅ '+name+' 추가됐습니다!');
    const mems=await srv('serverGetMembers');
    initMems(mems);renderMemList();
  }catch(e){toast(e.message,'err');}
}

// ── 모달: 거래 ────────────────────────────────────
function openTxn(){$('t-date').value=today();['t-amt','t-desc','t-memo'].forEach(id=>$(id).value='');syncSels();fillCats();openM('mod-txn');}
function setType(t){$('t-type').value=t;$('tb-e').className='tb '+(t==='지출'?'ae':'');$('tb-i').className='tb '+(t==='수입'?'ai':'');fillCats();}
function fillCats(){$('t-cat').innerHTML=($('t-type').value==='수입'?ICATS:ECATS).map(c=>'<option>'+c+'</option>').join('');}
async function submitTxn(){
  const p={date:gv('t-date'),type:gv('t-type'),category:gv('t-cat'),desc:gv('t-desc'),amount:gv('t-amt'),payMethod:gv('t-pay'),memo:gv('t-memo'),member:gv('t-who')||S.me};
  if(!p.amount||!p.desc){toast('내용과 금액을 입력하세요','err');return;}
  const btn=$('t-btn');btn.textContent='저장 중...';btn.disabled=true;
  try{
    const res=await srv('serverAddTransaction',p);
    closeM('mod-txn');toast('✅ 저장됐습니다!');
    const txn={...p,id:res.id||'',amount:Number(p.amount),createdAt:new Date().toISOString()};
    addTxnRow(txn,'dash-txns');
    addFeedRow({time:new Date().toISOString(),member:p.member,action:'거래추가',detail:p.type+' '+p.category+' '+fmt(Number(p.amount))+'원'},'live-feed');
    if(active('ledger'))loadLedger();
    setTimeout(loadDash,800);
  }catch(e){toast('저장 실패: '+e.message,'err');}
  finally{btn.textContent='저장';btn.disabled=false;}
}
function addTxnRow(t,id){
  const el=$(id);if(!el)return;
  const em=el.querySelector('.empty');if(em)em.remove();
  const div=document.createElement('div');div.innerHTML=txnRow(t,false);
  const ch=div.firstElementChild;if(ch){ch.classList.add('new');el.prepend(ch);}
  while(el.children.length>8)el.removeChild(el.lastChild);
}
async function delTxn(id,desc,e){
  e.stopPropagation();if(!confirm('"'+desc+'" 거래를 삭제할까요?'))return;
  try{await srv('serverDeleteTransaction',{id,desc,member:S.me});toast('🗑️ 삭제됐습니다');loadLedger();setTimeout(loadDash,500);}
  catch(e){toast(e.message,'err');}
}

// ── 모달: 자산 ────────────────────────────────────
function openAsset(){['a-nm','a-inst','a-val','a-cost','a-memo'].forEach(id=>$(id).value='');$('a-id').value='';openM('mod-asset');}
function openEditAsset(id,name,type,inst,val,cost,memo){$('a-id').value=id;$('a-nm').value=name;$('a-type').value=type;$('a-inst').value=inst;$('a-val').value=val;$('a-cost').value=cost;$('a-memo').value=memo;openM('mod-asset');}
async function submitAsset(){
  const id=gv('a-id');
  const p={name:gv('a-nm'),type:gv('a-type'),institution:gv('a-inst'),currentValue:gv('a-val'),purchaseValue:gv('a-cost'),memo:gv('a-memo'),member:S.me};
  if(!p.name||!p.currentValue){toast('자산명과 금액을 입력하세요','err');return;}
  try{id?await srv('serverUpdateAsset',{id,...p}):await srv('serverAddAsset',p);closeM('mod-asset');toast('✅ 저장됐습니다!');loadAssets();setTimeout(loadDash,500);}
  catch(e){toast(e.message,'err');}
}

// ── 모달: 목표 ────────────────────────────────────
function openGoal(){['g-nm','g-date','g-tgt','g-cur','g-mon'].forEach(id=>$(id).value='');$('g-id').value='';openM('mod-goal');}
function openEditGoal(id,name,type,tgt,cur,date,mon){$('g-id').value=id;$('g-nm').value=name;$('g-type').value=type;$('g-tgt').value=tgt;$('g-cur').value=cur;$('g-date').value=date;$('g-mon').value=mon;openM('mod-goal');}
async function submitGoal(){
  const id=gv('g-id');
  const p={name:gv('g-nm'),type:gv('g-type'),target:gv('g-tgt'),current:gv('g-cur'),targetDate:gv('g-date'),monthly:gv('g-mon'),member:S.me};
  if(!p.name||!p.target){toast('목표명과 금액을 입력하세요','err');return;}
  try{id?await srv('serverUpdateGoal',{id,current:p.current,monthly:p.monthly,name:p.name,member:S.me}):await srv('serverAddGoal',p);closeM('mod-goal');toast('✅ 저장됐습니다!');loadGoals();setTimeout(loadDash,500);}
  catch(e){toast(e.message,'err');}
}

// ── 유틸 ──────────────────────────────────────────
const $=id=>document.getElementById(id);
const gv=id=>$(id)?.value||'';
const esc=s=>String(s).replace(/\\\\/g,'').replace(/'/g,"\\\\'").replace(/"/g,'&quot;');
function fmt(n){if(n==null)return '—';const a=Math.abs(Math.round(n));if(a>=100000000)return (n/100000000).toFixed(1)+'억';if(a>=10000)return Math.round(n/10000).toLocaleString()+'만';return Math.round(n).toLocaleString();}
function today(){return new Date().toISOString().slice(0,10);}
function aIcon(t){return {'예금/현금':'💰','주식/ETF':'📈','부동산':'🏠','자동차':'🚗','적금/보험':'🛡️'}[t]||'📦';}
function rtime(iso){if(!iso)return '';const d=(Date.now()-new Date(iso))/1000;if(d<60)return '방금';if(d<3600)return Math.floor(d/60)+'분 전';if(d<86400)return Math.floor(d/3600)+'시간 전';return new Date(iso).toLocaleDateString('ko-KR',{month:'short',day:'numeric'});}
function txnRow(t,showDel){
  const isInc=t.type==='수입';const m=S.members.find(mb=>mb.name===t.member);const mc=m?.color||'#94a3b8';
  const sid=String(t.id||'').replace(/'/g,'');const sdsc=String(t.desc||t.category||'').replace(/'/g,'');
  return '<div class="ti"><div class="tic">'+(ICONS[t.category]||'•')+'</div><div class="tif"><div class="tin">'+(t.desc||t.category||'')+'</div><div class="tic2"><span class="tmb" style="background:'+mc+'20;color:'+mc+'">'+(t.member||'—')+'</span><span>'+(t.category||'')+'</span><span>'+(t.date||'')+'</span><span>'+(t.payMethod||'')+'</span></div></div><div class="tam '+(isInc?'pos':'neg')+'">'+(isInc?'+':'-')+fmt(t.amount)+'</div>'+(showDel&&sid?'<button class="tdel" onclick="delTxn(\''+sid+'\',\''+sdsc+'\',event)">✕</button>':'')+'</div>';
}
function openM(id){$(id).classList.add('on');}
function closeM(id){$(id).classList.remove('on');}
document.querySelectorAll('.mod-bg').forEach(el=>el.addEventListener('click',e=>{if(e.target===el)el.classList.remove('on');}));
let toastT;
function toast(msg,type='ok'){const el=$('toast');el.textContent=msg;el.className='toast on '+type;clearTimeout(toastT);toastT=setTimeout(()=>el.classList.remove('on'),4000);}
let alertT;
function alert2(msg,icon='🔔'){$('ab-msg').textContent=msg;$('ab-icon').textContent=icon;$('ab').classList.add('on');clearTimeout(alertT);alertT=setTimeout(()=>$('ab').classList.remove('on'),4000);}

// ── 초기화 ────────────────────────────────────────
(async function init(){
  try{const b=localStorage.getItem(BK);if(b)S.budgets=JSON.parse(b);}catch(e){}
  initMonLbl();
  // 구성원 먼저 로드
  try{const mems=await srv('serverGetMembers');initMems(mems);}catch(e){initMems(null);}
  // 대시보드 로드
  loadDash();
  // 폴링 시작
  startPoll();
})();
</script>
</body>
</html>`;
}
