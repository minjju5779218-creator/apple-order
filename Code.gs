/**
 * 🍎 사과 주문 관리 (구글 시트 + 앱스 스크립트)  — 2판
 *
 * - 주문서에서 들어온 주문을 [주문] 시트에 쌓습니다. 여러 곳에 보내는 주문은 배송지마다 한 줄씩 들어가고
 *   주문번호가 같습니다. [입금할 금액]은 그 주문의 첫 줄에만 적혀요(합계).
 * - [입금확인]은 같은 주문 중 아무 줄이나 체크하면 나머지 줄도 같이 체크됩니다.
 * - 손님은 주문서의 [내 주문 확인]에서 이름 + 휴대폰 번호로 주문 상태(입금 확인 중 / 입금 확인 완료 / 발송 준비 중 / 취소됨)를 보고,
 *   주문 후 정해진 시간(기본 2시간) 안에, 입금 확인 전이면 직접 취소할 수 있어요.
 * - 가격·품절·계좌·공지·안내사항은 [가격표], [설정] 시트에서 바꾸면 바로 반영됩니다.
 * - 메뉴 [🍎 주문관리 > 로젠 송장 파일 만들기] : 입금확인된(취소 안 된) 주문만 로젠 업로드용 엑셀로 만듭니다.
 *
 * 1판에서 바꿔 붙여넣은 경우에도 기존 시트는 그대로 쓰고, 새 칸만 자동으로 추가돼요.
 */

const SHEET_ORDERS = '주문';
const SHEET_PRICES = '가격표';
const SHEET_SETTINGS = '설정';
const TZ = 'Asia/Seoul';
const MAX_DEST = 20;

const ORDER_HEADERS = [
  '주문번호', '주문일시', '입금확인', '송장출력',
  '주문자', '주문자 연락처', '받는 분', '받는 분 연락처',
  '우편번호', '주소', '상세주소', '주문 내역', '박스 수',
  '상품 금액', '추가 택배비', '입금할 금액', '입금자명', '배송 메모',
  '상품코드(수정금지)',
  // ↓ 2판에서 추가
  '보내는 분', '보내는 분 연락처', '취소일시', '배송지',
  // ↓ 3판에서 추가 (손님이 지운 날짜 — 손님 화면에서만 숨김, 시트 기록은 남음)
  '손님 내역삭제'
];
const COL = {}; ORDER_HEADERS.forEach((h, i) => COL[h] = i + 1);

const DEFAULT_PRICES = [
  ['사과 10kg', '20과', 130000], ['사과 10kg', '22과', 120000], ['사과 10kg', '24과', 110000],
  ['사과 10kg', '26과', 100000], ['사과 10kg', '30과', 90000],  ['사과 10kg', '32과', 85000],
  ['사과 10kg', '36과', 75000],  ['사과 10kg', '40과', 65000],  ['사과 10kg', '42과', 60000],
  ['사과 5kg', '10과', 70000], ['사과 5kg', '11과', 65000], ['사과 5kg', '12과', 60000],
  ['사과 5kg', '13과', 55000], ['사과 5kg', '15과', 50000], ['사과 5kg', '16과', 45000],
  ['사과 5kg', '18과', 40000], ['사과 5kg', '20과', 35000],
  ['사과즙', '1박스 (50포)', 30000], ['사과즙', '2박스 (100포)', 50000]
];

const OLD_GUIDES = [
  '· 입금 확인 후 2~3일 안에 발송해요.\n· 여러 곳에 보내실 때는 [다른 곳에도 보내기]를 눌러 배송지를 추가하세요.\n· 주문 후 2시간 안에는 [내 주문 확인]에서 직접 취소할 수 있어요.\n· 받으신 사과에 문제가 있으면 사진과 함께 문자 주세요.'
];

const DEFAULT_SETTINGS = [
  ['농장 이름', '○○농원', '주문서 맨 위에 나옵니다'],
  ['입금 계좌', '○○은행 000-0000-0000 (예금주 ○○○)', '주문 완료 화면에 나옵니다'],
  ['문의 전화', '010-0000-0000', ''],
  ['공지', '기본 택배비 무료 · 입금 확인 후 순서대로 발송', '주문서 맨 위 빨간 띠에 나옵니다 (비워도 됨)'],
  ['안내사항', '· 입금 확인 후 2~3일 안에 발송해요.\n· 여러 곳에 보내실 때는 [다른 곳에도 보내기]를 눌러 배송지를 추가하세요.\n· 주문 후 2시간 안에는 [내 주문 확인]에서 직접 취소할 수 있어요.\n· 입금 확인된 주문은 주문일로부터 2주가 지나면 [내 주문 확인]에서 내역을 직접 지울 수 있어요.\n· 주문 내역은 주문일로부터 1달이 지나면 [내 주문 확인]에서 자동으로 사라져요.', '주문서의 안내사항 상자에 나옵니다. 칸 안에서 줄바꿈은 Ctrl+Enter (비워도 됨)'],
  ['주문 받기', '켜기', '켜기 / 끄기 — 끄면 주문서가 닫힙니다'],
  ['기본 보내는 분', '', '송장에 찍힐 보내는 분 이름. 비우면 농장 이름'],
  ['기본 보내는 분 연락처', '', '송장에 찍힐 보내는 분 연락처. 비우면 문의 전화'],
  ['내역 삭제 가능 일수', 14, '입금 확인된 주문을 손님이 [내 주문 확인]에서 지울 수 있게 되는 날 (주문일로부터 며칠 뒤)'],
  ['내역 보관 일수', 30, '주문일로부터 며칠이 지나면 [내 주문 확인]에서 자동으로 사라지는지 (시트 기록은 그대로 남아요)'],
  ['취소 가능 시간', 2, '주문 후 몇 시간 안에 손님이 직접 취소할 수 있는지 (입금 확인 전만, 0이면 취소 버튼 없음)'],
  ['추가 택배비', 3000, '제주·도서산간 추가 금액'],
  ['추가 택배비 방식', '박스당', '박스당 / 주문당(배송지당)'],
  ['추가 택배비 지역', '제주, 울릉군, 옹진군', '쉼표로 구분. 주소에 이 글자가 들어가면 추가 (예: 신안군 흑산면)'],
  ['로젠 운임구분', '신용', '로젠과 계약한 방식: 신용 / 선불'],
  ['새 주문 알림 메일', '', '메일 주소를 적으면 주문·취소가 생길 때마다 메일로 알려줍니다 (비워도 됨)']
];

/* ===================== 메뉴 · 시트 준비 ===================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🍎 주문관리')
    .addItem('로젠 송장 파일 만들기 (입금확인된 주문)', 'makeLogenFile')
    .addSeparator()
    .addItem('처음 설정 / 새 기능 칸 추가', 'setup')
    .addToUi();
}

// 같은 주문번호의 [입금확인]을 함께 체크/해제
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== SHEET_ORDERS || e.range.getColumn() !== COL['입금확인'] || e.range.getNumRows() !== 1 || e.range.getRow() < 2) return;
    const id = String(sh.getRange(e.range.getRow(), COL['주문번호']).getValue());
    if (!id) return;
    const val = e.range.getValue() === true;
    const ids = sh.getRange(2, COL['주문번호'], sh.getLastRow() - 1, 1).getValues();
    ids.forEach((r, i) => {
      const row = i + 2;
      if (row !== e.range.getRow() && String(r[0]) === id) sh.getRange(row, COL['입금확인']).setValue(val);
    });
  } catch (x) {}
}

function setup() {
  const ss = SpreadsheetApp.getActive();

  let sh = ss.getSheetByName(SHEET_ORDERS) || ss.insertSheet(SHEET_ORDERS, 0);
  if (sh.getLastRow() === 0) {
    sh.appendRow(ORDER_HEADERS);
    sh.setFrozenRows(1);
    sh.hideColumns(COL['상품코드(수정금지)']);
    [COL['상품 금액'], COL['추가 택배비'], COL['입금할 금액']].forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('#,##0'));
    sh.setColumnWidth(COL['주소'], 280); sh.setColumnWidth(COL['주문 내역'], 240);
  }
  ensureOrderHeaders(sh);

  sh = ss.getSheetByName(SHEET_PRICES) || ss.insertSheet(SHEET_PRICES);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['구분', '상품', '가격', '품절']);
    sh.getRange(2, 1, DEFAULT_PRICES.length, 3).setValues(DEFAULT_PRICES);
    sh.getRange(2, 4, DEFAULT_PRICES.length, 1).insertCheckboxes();
    sh.getRange('C:C').setNumberFormat('#,##0');
    sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#fdeceb');
    sh.setFrozenRows(1);
    sh.getRange(DEFAULT_PRICES.length + 3, 1).setValue('※ 줄을 추가하면 주문서에도 생깁니다. 품절에 체크하면 주문서에서 "품절"로 바뀝니다. 구분 이름이 같으면 한 묶음으로 보입니다.');
  }

  sh = ss.getSheetByName(SHEET_SETTINGS) || ss.insertSheet(SHEET_SETTINGS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['항목', '값', '설명']);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#fdeceb');
    sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 360); sh.setColumnWidth(3, 420);
  }
  ensureSettings(sh);

  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);

  SpreadsheetApp.getUi().alert('설정 완료! [설정] 시트에서 농장 이름, 입금 계좌, 안내사항을 확인해주세요.');
}

// 주문 시트 맨 윗줄에 빠진 제목 칸을 채움 (1판 → 2판)
function ensureOrderHeaders(sh) {
  const width = Math.max(sh.getLastColumn(), 1);
  const cur = sh.getRange(1, 1, 1, width).getValues()[0];
  if (cur.length >= ORDER_HEADERS.length && cur[ORDER_HEADERS.length - 1] === ORDER_HEADERS[ORDER_HEADERS.length - 1]) return;
  sh.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS])
    .setFontWeight('bold').setBackground('#fdeceb');
}

// 설정 시트에 빠진 항목을 추가 (기존 값은 건드리지 않음)
function ensureSettings(sh) {
  const have = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => have[String(r[0]).trim()] = true);
  const missing = DEFAULT_SETTINGS.filter(r => !have[r[0]]);
  if (!missing.length) return;
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, missing.length, 3).setValues(missing);
  sh.getRange(start, 3, missing.length, 1).setFontColor('#888');
  sh.getRange(start, 2, missing.length, 1).setWrap(true);
}

/* ===================== 읽기 ===================== */
function readSettings() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  const m = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) m[String(r[0]).trim()] = r[1]; });
  // 1판 시트라면 새 설정 항목(안내사항 등)을 알아서 추가
  if (sh && (!('안내사항' in m) || !('내역 보관 일수' in m))) {
    ensureSettings(sh);
    DEFAULT_SETTINGS.forEach(r => { if (!(r[0] in m)) m[r[0]] = r[1]; });
  }
  // 예전 기본 안내문을 그대로 쓰고 있으면 새 안내문으로 바꿈 (직접 고친 안내문은 건드리지 않음)
  if (sh && OLD_GUIDES.indexOf(String(m['안내사항'] || '').replace(/\r/g, '')) !== -1) {
    const newGuide = DEFAULT_SETTINGS.find(r => r[0] === '안내사항')[1];
    const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    keys.forEach((r, i) => { if (String(r[0]).trim() === '안내사항') sh.getRange(i + 2, 2).setValue(newGuide); });
    m['안내사항'] = newGuide;
  }
  const num = (k, def) => (m[k] === '' || m[k] == null || isNaN(Number(m[k]))) ? def : Number(m[k]);
  const farmName = String(m['농장 이름'] || '');
  const phone = String(m['문의 전화'] || '');
  return {
    farmName: farmName,
    bank: String(m['입금 계좌'] || ''),
    phone: phone,
    notice: String(m['공지'] || ''),
    guide: String(m['안내사항'] || ''),
    open: String(m['주문 받기'] || '켜기').trim() !== '끄기',
    senderName: String(m['기본 보내는 분'] || '').trim() || farmName,
    senderPhone: String(m['기본 보내는 분 연락처'] || '').trim() || phone,
    cancelHours: num('취소 가능 시간', 2),
    deleteDays: num('내역 삭제 가능 일수', 14),
    keepDays: num('내역 보관 일수', 30),
    islandFee: Number(m['추가 택배비']) || 0,
    islandFeeMode: String(m['추가 택배비 방식'] || '박스당').trim().indexOf('주문당') === 0 ? '주문당' : '박스당',
    islandAreas: String(m['추가 택배비 지역'] || '').split(',').map(s => s.trim()).filter(Boolean),
    logenFare: String(m['로젠 운임구분'] || '신용').trim(),
    notifyEmail: String(m['새 주문 알림 메일'] || '').trim()
  };
}

function readProducts() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PRICES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(r => r[0] && r[1] && Number(r[2]) > 0)
    .map(r => ({ group: String(r[0]).trim(), name: String(r[1]).trim(), price: Number(r[2]), soldOut: r[3] === true }));
}

/* ===================== 주문서 ↔ 시트 ===================== */
function doGet() {
  try {
    const s = readSettings();
    return json({
      ok: true,
      settings: { farmName: s.farmName, bank: s.bank, phone: s.phone, notice: s.notice, guide: s.guide, open: s.open,
                  senderName: s.senderName, cancelHours: s.cancelHours,
                  islandFee: s.islandFee, islandFeeMode: s.islandFeeMode, islandAreas: s.islandAreas },
      products: readProducts()
    });
  } catch (err) {
    return json({ ok: false, error: '가격표를 읽지 못했어요.' });
  }
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (x) { return json({ ok: false, error: '잘못된 요청이에요.' }); }
  if (d.action === 'lookup') return lookupOrders(d);
  if (d.action === 'cancel') return cancelOrder(d);
  if (d.action === 'hide') return hideOrder(d);
  return placeOrder(d);
}

/* ---------- 공통 도우미 ---------- */
const str = (v, max) => String(v == null ? '' : v).trim().slice(0, max || 100);
const digits = v => String(v == null ? '' : v).replace(/\D/g, '');
function fmtPhone(v) {
  const n = digits(v).slice(0, 11);
  if (n.length === 11) return n.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
  if (n.length === 10) return n.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
  return n;
}
const mobileOk = p => /^01[016789]\d{7,8}$/.test(digits(p));
const anyPhoneOk = p => /^0\d{8,10}$/.test(digits(p));   // 보내는 분은 집 전화도 허용

// '주문일시' 칸 값을 날짜로 (시트가 날짜로 바꿔 저장해도, 글자로 저장돼도 처리)
function orderTime(v) {
  if (v instanceof Date) return v;
  const m = String(v).match(/(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})/);
  return m ? new Date(m[1] + '-' + m[2] + '-' + m[3] + 'T' + m[4] + ':' + m[5] + ':00+09:00') : null;
}

// 한 주문(여러 줄)의 상태
function statusOfRows(rows) {
  if (rows.some(x => String(x.g('취소일시')).trim())) return '취소됨';
  if (rows.some(x => String(x.g('송장출력')).trim())) return '발송 준비 중';
  if (rows.some(x => x.g('입금확인') === true)) return '입금 확인 완료';
  return '입금 확인 중';
}

/* ---------- 주문 접수 ---------- */
function placeOrder(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const s = readSettings();
    if (!s.open) return json({ ok: false, error: '지금은 주문을 받지 않아요.' });

    const buyerName = str(d.buyerName, 30), buyerPhone = fmtPhone(d.buyerPhone), payer = str(d.payer, 30);
    const custom = d.senderMode === 'custom';
    const senderName = custom ? str(d.senderName, 30) : '', senderPhone = custom ? fmtPhone(d.senderPhone) : '';
    if (!buyerName || !payer) return json({ ok: false, error: '이름과 입금자명을 확인해주세요.' });
    if (!mobileOk(buyerPhone)) return json({ ok: false, error: '주문하시는 분 휴대폰 번호를 확인해주세요.' });
    if (custom && (!senderName || !anyPhoneOk(senderPhone))) return json({ ok: false, error: '보내는 분 이름과 연락처를 확인해주세요.' });

    // 배송지 목록 (1판 주문서 형식도 받음)
    let rawDests = Array.isArray(d.destinations) ? d.destinations
      : [{ recvName: d.recvName, recvPhone: d.recvPhone, zip: d.zip, addr: d.addr, addr2: d.addr2, memo: d.memo, items: d.items }];
    if (!rawDests.length || rawDests.length > MAX_DEST) return json({ ok: false, error: '배송지는 1~' + MAX_DEST + '곳까지 가능해요.' });

    const prices = {};
    readProducts().forEach(p => prices[p.group + '|' + p.name] = p);

    const dests = [];
    for (let n = 0; n < rawDests.length; n++) {
      const r = rawDests[n] || {}, label = rawDests.length > 1 ? (n + 1) + '번째 배송지의 ' : '';
      const o = { recvName: str(r.recvName, 30), recvPhone: fmtPhone(r.recvPhone), zip: str(r.zip, 10),
                  addr: str(r.addr, 200), addr2: str(r.addr2, 100), memo: str(r.memo, 200), items: [] };
      if (!o.recvName || !mobileOk(o.recvPhone)) return json({ ok: false, error: label + '받는 분 이름·휴대폰 번호를 확인해주세요.' });
      if (!/^\d{5}$/.test(o.zip) || !o.addr) return json({ ok: false, error: label + '주소를 다시 검색해주세요.' });
      for (const it of (Array.isArray(r.items) ? r.items : [])) {
        const p = prices[str(it.group) + '|' + str(it.name)];
        const q = Math.floor(Number(it.qty));
        if (!p || !(q > 0) || q > 50) continue;
        if (p.soldOut) return json({ ok: false, error: p.group + ' ' + p.name + ' 상품이 방금 품절됐어요. 다른 상품을 골라주세요.' });
        o.items.push({ group: p.group, name: p.name, price: p.price, qty: q });
      }
      if (!o.items.length) return json({ ok: false, error: label + '상품을 골라주세요.' });
      o.boxes = o.items.reduce((a, i) => a + i.qty, 0);
      o.sub = o.items.reduce((a, i) => a + i.qty * i.price, 0);
      const island = s.islandAreas.some(a => o.addr.indexOf(a) !== -1);
      o.fee = island ? (s.islandFeeMode === '주문당' ? s.islandFee : s.islandFee * o.boxes) : 0;
      o.itemText = o.items.map(i => i.group + ' ' + i.name + ' ×' + i.qty).join(', ');
      dests.push(o);
    }
    const total = dests.reduce((a, o) => a + o.sub + o.fee, 0);

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
    ensureOrderHeaders(sh);
    const now = new Date();
    const orderId = makeOrderId(sh, now);
    const when = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm');
    const rows = dests.map((o, n) => {
      const row = new Array(ORDER_HEADERS.length).fill('');
      const put = (h, v) => row[COL[h] - 1] = (typeof v === 'string' ? safe(v) : v);
      put('주문번호', orderId); put('주문일시', when); put('입금확인', false);
      put('주문자', buyerName); put('주문자 연락처', buyerPhone);
      put('받는 분', o.recvName); put('받는 분 연락처', o.recvPhone);
      put('우편번호', o.zip); put('주소', o.addr); put('상세주소', o.addr2);
      put('주문 내역', o.itemText); put('박스 수', o.boxes); put('상품 금액', o.sub); put('추가 택배비', o.fee);
      put('입금할 금액', n === 0 ? total : '');
      put('입금자명', payer); put('배송 메모', o.memo);
      put('상품코드(수정금지)', JSON.stringify(o.items.map(i => [i.group, i.name, i.qty])));
      put('보내는 분', senderName || '(농장 기본)'); put('보내는 분 연락처', senderPhone);
      put('배송지', (n + 1) + '/' + dests.length);
      return row;
    });
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, ORDER_HEADERS.length).setValues(rows);
    sh.getRange(start, COL['입금확인'], rows.length, 1).insertCheckboxes();
    sh.getRange(start, COL['우편번호'], rows.length, 1).setNumberFormat('@').setValues(dests.map(o => [o.zip]));

    notify(s, '🍎 새 주문 ' + orderId + ' · ' + buyerName + ' · ' + total.toLocaleString() + '원' + (dests.length > 1 ? ' · ' + dests.length + '곳' : ''),
      ['주문번호: ' + orderId, '주문자: ' + buyerName + ' (' + buyerPhone + ')',
       '보내는 분: ' + (senderName ? senderName + ' (' + senderPhone + ')' : '농장 기본'),
       '입금할 금액: ' + total.toLocaleString() + '원 (입금자명 ' + payer + ')', '']
      .concat(dests.map((o, n) => '[' + (n + 1) + '] ' + o.recvName + ' ' + o.recvPhone + ' / (' + o.zip + ') ' + o.addr + ' ' + o.addr2 + ' / ' + o.itemText + (o.memo ? ' / 메모: ' + o.memo : ''))));

    return json({ ok: true, orderId: orderId, total: total, bank: s.bank, cancelHours: s.cancelHours });
  } catch (err) {
    return json({ ok: false, error: '주문 저장 중 문제가 생겼어요.' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- 내 주문 확인 ---------- */
// 이름 + 휴대폰이 맞는 주문을 주문번호별로 묶어서 돌려줌 (최신순)
function findOrders(name, phone) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  if (!sh || sh.getLastRow() < 2) return { sh: sh, groups: [] };
  ensureOrderHeaders(sh);
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, ORDER_HEADERS.length).getValues();
  const byId = {}, order = [];
  data.forEach((r, i) => {
    const g = h => r[COL[h] - 1];
    if (digits(g('주문자 연락처')) !== phone || String(g('주문자')).replace(/\s/g, '') !== name) return;
    const id = String(g('주문번호'));
    if (!byId[id]) { byId[id] = []; order.push(id); }
    byId[id].push({ row: i + 2, g: g });
  });
  return { sh: sh, groups: order.reverse().map(id => ({ id: id, rows: byId[id] })) };
}

function lookupOrders(d) {
  const name = str(d.name, 30).replace(/\s/g, ''), phone = digits(d.phone);
  if (!name || !mobileOk(phone)) return json({ ok: false, error: '이름과 휴대폰 번호를 확인해주세요.' });
  const s = readSettings();
  const now = Date.now();
  const { groups } = findOrders(name, phone);
  const DAY = 24 * 3600 * 1000;
  const visible = groups.filter(({ rows }) => {
    if (rows.some(x => String(x.g('손님 내역삭제')).trim())) return false;      // 손님이 지운 주문
    const t = orderTime(rows[0].g('주문일시'));
    return !(t && s.keepDays > 0 && now > t.getTime() + s.keepDays * DAY);    // 보관 기간 지난 주문
  });
  const fmtDay = ms => Utilities.formatDate(new Date(ms), TZ, 'M월 d일');
  const orders = visible.slice(0, 20).map(({ id, rows }) => {
    const g0 = rows[0].g;
    const t = orderTime(g0('주문일시'));
    const status = statusOfRows(rows);
    const deadline = t ? t.getTime() + s.cancelHours * 3600 * 1000 : 0;
    const delFrom = t ? t.getTime() + s.deleteDays * DAY : 0;
    const sender = String(g0('보내는 분') || '');
    return {
      orderId: id,
      date: t ? Utilities.formatDate(t, TZ, 'yyyy-MM-dd HH:mm') : String(g0('주문일시')),
      total: rows.reduce((a, x) => a + (Number(x.g('상품 금액')) || 0) + (Number(x.g('추가 택배비')) || 0), 0),
      payer: String(g0('입금자명')),
      sender: (!sender || sender === '(농장 기본)') ? s.senderName : sender,
      dests: rows.map(x => ({
        recv: String(x.g('받는 분')) + ' · ' + String(x.g('받는 분 연락처')),
        addr: '(' + String(x.g('우편번호')) + ') ' + String(x.g('주소')) + ' ' + String(x.g('상세주소')),
        items: String(x.g('주문 내역')),
        memo: String(x.g('배송 메모'))
      })),
      status: status,
      canCancel: status === '입금 확인 중' && s.cancelHours > 0 && now < deadline,
      cancelUntil: deadline ? Utilities.formatDate(new Date(deadline), TZ, 'MM월 dd일 HH:mm') : '',
      canDelete: canHide(status, t, s),
      deleteFrom: (status === '입금 확인 완료' || status === '발송 준비 중') && delFrom > now ? fmtDay(delFrom) : '',
      autoGone: t && s.keepDays > 0 ? fmtDay(t.getTime() + s.keepDays * DAY) : ''
    };
  });
  return json({ ok: true, orders: orders, bank: s.bank, phone: s.phone });
}

/* ---------- 손님 취소 ---------- */
function cancelOrder(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const name = str(d.name, 30).replace(/\s/g, ''), phone = digits(d.phone), id = str(d.orderId, 20);
    const s = readSettings();
    const { sh, groups } = findOrders(name, phone);
    const hit = groups.find(x => x.id === id);
    if (!hit) return json({ ok: false, error: '주문을 찾지 못했어요.' });
    const st = statusOfRows(hit.rows);
    const call = s.phone ? '\n' + s.phone + '로 연락 주세요.' : '';
    if (st === '취소됨') return json({ ok: false, error: '이미 취소된 주문이에요.' });
    if (st !== '입금 확인 중') return json({ ok: false, error: '입금이 확인된 주문은 직접 취소할 수 없어요.' + call });
    const t = orderTime(hit.rows[0].g('주문일시'));
    if (!t || s.cancelHours <= 0 || Date.now() > t.getTime() + s.cancelHours * 3600 * 1000)
      return json({ ok: false, error: '취소 가능 시간(' + s.cancelHours + '시간)이 지났어요.' + call });

    const when = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    hit.rows.forEach(x => {
      sh.getRange(x.row, COL['취소일시']).setValue(when);
      sh.getRange(x.row, 1, 1, ORDER_HEADERS.length).setFontColor('#aaaaaa').setFontLine('line-through');
    });
    notify(s, '❌ 주문 취소 ' + id + ' · ' + hit.rows[0].g('주문자'),
      ['손님이 주문을 취소했어요.', '주문번호: ' + id].concat(hit.rows.map(x => '- ' + x.g('받는 분') + ': ' + x.g('주문 내역'))));
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: '취소 중 문제가 생겼어요.' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// 손님이 내역을 지울 수 있는지: 입금 확인된(또는 발송 준비 중) 주문은 주문일로부터 정한 일수가 지나야, 취소된 주문은 언제든
function canHide(status, t, s) {
  if (status === '취소됨') return true;
  if (status !== '입금 확인 완료' && status !== '발송 준비 중') return false;
  return !!t && Date.now() >= t.getTime() + s.deleteDays * 24 * 3600 * 1000;
}

/* ---------- 손님 내역 지우기 (손님 화면에서만 숨김) ---------- */
function hideOrder(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const name = str(d.name, 30).replace(/\s/g, ''), phone = digits(d.phone), id = str(d.orderId, 20);
    const s = readSettings();
    const { sh, groups } = findOrders(name, phone);
    const hit = groups.find(x => x.id === id);
    if (!hit) return json({ ok: false, error: '주문을 찾지 못했어요.' });
    const st = statusOfRows(hit.rows);
    if (!canHide(st, orderTime(hit.rows[0].g('주문일시')), s))
      return json({ ok: false, error: st === '입금 확인 중' ? '입금 확인 전인 주문은 지울 수 없어요.' : '주문일로부터 ' + s.deleteDays + '일이 지나야 지울 수 있어요.' });
    const when = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    hit.rows.forEach(x => sh.getRange(x.row, COL['손님 내역삭제']).setValue(when));
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: '지우는 중 문제가 생겼어요.' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function makeOrderId(sh, now) {
  const prefix = Utilities.formatDate(now, TZ, 'yyMMdd') + '-';
  let n = 0;
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
      const v = String(r[0]);
      if (v.indexOf(prefix) === 0) n = Math.max(n, Number(v.slice(prefix.length)) || 0);
    });
  }
  return prefix + String(n + 1).padStart(3, '0');
}

function notify(s, subject, lines) {
  if (!s.notifyEmail) return;
  try { MailApp.sendEmail(s.notifyEmail, subject, lines.join('\n')); } catch (x) { /* 메일 실패해도 계속 */ }
}

/* ===================== 로젠 송장 파일 ===================== */
const LOGEN_HEADERS = ['수하인명', '우편번호', '수하인주소', '수하인전화번호', '수하인핸드폰번호',
                       '택배수량', '운임구분', '물품명', '배송메세지', '송하인명', '송하인전화번호', '주문번호'];

function makeLogenFile() {
  const ui = SpreadsheetApp.getUi();
  const s = readSettings();
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  ensureOrderHeaders(sh);
  const last = sh.getLastRow();
  if (last < 2) { ui.alert('주문이 아직 없어요.'); return; }

  const data = sh.getRange(2, 1, last - 1, ORDER_HEADERS.length).getValues();
  const out = [];
  const doneRows = [];
  const orderIds = {};
  data.forEach((r, i) => {
    const g = h => r[COL[h] - 1];
    if (g('입금확인') !== true || String(g('송장출력')).trim() || String(g('취소일시')).trim()) return;
    let items = [];
    try { items = JSON.parse(g('상품코드(수정금지)')); } catch (x) {}
    if (!items.length) items = [['사과', '', Number(g('박스 수')) || 1]];
    const address = (String(g('주소')) + ' ' + String(g('상세주소'))).trim();
    const sender = String(g('보내는 분') || '').trim();
    const custom = sender && sender !== '(농장 기본)';
    const sName = custom ? sender : s.senderName;
    const sPhone = custom ? (String(g('보내는 분 연락처')) || s.senderPhone) : s.senderPhone;
    // 박스 1개당 송장 1장 → 박스마다 물품명이 찍혀 포장할 때 헷갈리지 않음
    items.forEach(([grp, name, qty]) => {
      for (let k = 0; k < qty; k++) {
        out.push([g('받는 분'), String(g('우편번호')).padStart(5, '0'), address, g('받는 분 연락처'), g('받는 분 연락처'),
                  1, s.logenFare, (grp + ' ' + name).trim(), String(g('배송 메모')), sName, sPhone, g('주문번호')]);
      }
    });
    doneRows.push(i + 2);
    orderIds[g('주문번호')] = true;
  });

  if (!out.length) { ui.alert('보낼 주문이 없어요.\n\n[입금확인]에 체크되어 있고, [송장출력]·[취소일시]가 비어 있는 주문만 파일로 만들어요.'); return; }

  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
  const file = SpreadsheetApp.create('로젠송장_' + stamp);
  const fsh = file.getSheets()[0];
  fsh.getRange(1, 1, 1, LOGEN_HEADERS.length).setValues([LOGEN_HEADERS]).setFontWeight('bold');
  fsh.getRange(2, 1, out.length, LOGEN_HEADERS.length).setNumberFormat('@').setValues(out.map(r => r.map(v => safe(String(v)))));
  fsh.autoResizeColumns(1, LOGEN_HEADERS.length);
  SpreadsheetApp.flush();

  const mark = Utilities.formatDate(new Date(), TZ, 'MM/dd HH:mm');
  doneRows.forEach(r => sh.getRange(r, COL['송장출력']).setValue(mark));

  const url = 'https://docs.google.com/spreadsheets/d/' + file.getId() + '/export?format=xlsx';
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;font-size:15px;line-height:1.6">' +
    '<p>주문 <b>' + Object.keys(orderIds).length + '건</b> (배송지 ' + doneRows.length + '곳), 송장 <b>' + out.length + '장</b> 파일을 만들었어요.</p>' +
    '<p><a href="' + url + '" target="_blank" style="display:inline-block;background:#c8322b;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:bold">엑셀 파일 내려받기</a></p>' +
    '<p style="color:#666">내려받은 파일을 로젠 프로그램의 엑셀 일괄 등록에 올리세요.<br>' +
    '주문 시트의 [송장출력] 칸에 오늘 시간이 적혔어요. 다시 뽑으려면 그 칸을 지우고 메뉴를 다시 누르면 돼요.</p></div>'
  ).setWidth(440).setHeight(250);
  ui.showModalDialog(html, '로젠 송장 파일');
}

/* ===================== 도우미 ===================== */
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
// 시트에서 수식으로 실행되지 않게 막기
function safe(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }
