/**
 * 이감 모의고사 주문 취합 시스템 — Apps Script 백엔드
 * 잇올 독학관리학원(수만휘, 지점 40곳) 주문을 구글 시트에 저장하는 API.
 *
 * 주문 흐름:
 *  1) 학생 주문: student.html?b=지점주문코드 링크로 학생이 개별 주문 (상태: 접수)
 *  2) 지점 승인: branch.html?k=지점관리코드 에서 학원이 자기 지점 주문을 취합·승인/반려
 *  3) 본사 취합: admin.html 에서 전 지점 승인분을 취합해 이감으로 전달 (확정→결제완료→출고)
 *  별도로 지점 담당자 일괄 주문(index.html)도 지원 (상태: 접수부터 동일 흐름)
 *
 * 설치 순서 (설치가이드.md 참고):
 * 1) 새 구글 스프레드시트 생성
 * 2) 확장 프로그램 > Apps Script > 이 코드 전체 붙여넣기
 * 3) 함수 선택에서 setup 선택 후 실행 (권한 승인)
 * 4) 배포 > 새 배포 > 웹 앱: 실행 계정 "나", 액세스 권한 "모든 사용자"
 * 5) /exec URL을 config.js 또는 admin.html 설정 탭에 입력
 * 6) 관리 토큰은 스프레드시트의 "설정" 시트에 표시됨
 */

var SHEET_ORDERS = "주문";
var SHEET_BRANCHES = "지점";
var SHEET_PRODUCTS = "상품";
var SHEET_CONFIG = "설정";

var ORDER_HEADER = ["주문ID", "접수일시", "지점", "주문자", "연락처", "상품", "수량", "단가", "금액", "메모", "상태", "구분"];
var BRANCH_HEADER = ["지점명", "사용", "주문코드", "관리코드"];
var STATUSES = ["접수", "지점승인", "반려", "확정", "결제완료", "출고", "취소"];
var BRANCH_STATUSES = ["접수", "지점승인", "반려"]; // 지점이 바꿀 수 있는 상태

/** 최초 1회 실행: 시트·헤더·기본 데이터·관리 토큰 생성 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var orders = getOrCreateSheet_(ss, SHEET_ORDERS);
  if (orders.getLastRow() === 0) orders.appendRow(ORDER_HEADER);

  var branches = getOrCreateSheet_(ss, SHEET_BRANCHES);
  if (branches.getLastRow() === 0) {
    branches.appendRow(BRANCH_HEADER);
    branches.appendRow(["본점(예시)", true, genCode_(), genCode_()]);
  }

  var products = getOrCreateSheet_(ss, SHEET_PRODUCTS);
  if (products.getLastRow() === 0) {
    products.appendRow(["상품명", "단가", "판매"]);
    products.appendRow(["시즌4 최상위 실전대비 (모의고사 2회)", 0, true]);
    products.appendRow(["시즌5 파이널1 9모대비 (모의고사+간쓸개 8회)", 0, true]);
  }

  var config = getOrCreateSheet_(ss, SHEET_CONFIG);
  var props = PropertiesService.getScriptProperties();
  var tokenValue = props.getProperty("ADMIN_TOKEN");
  if (!tokenValue) {
    tokenValue = Utilities.getUuid().replace(/-/g, "").slice(0, 20);
    props.setProperty("ADMIN_TOKEN", tokenValue);
  }
  config.clear();
  config.appendRow(["항목", "값"]);
  config.appendRow(["관리 토큰 (admin.html 설정 탭에 입력)", tokenValue]);
  config.appendRow(["생성일", new Date()]);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function genCode_() {
  return Utilities.getUuid().replace(/-/g, "").slice(0, 8);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function checkToken_(t) {
  var saved = PropertiesService.getScriptProperties().getProperty("ADMIN_TOKEN");
  return saved && t === saved;
}

/** GET: 데이터 조회 */
function doGet(e) {
  try {
    var p = e.parameter || {};
    var action = p.action || "public";

    // 지점 주문 페이지용: 지점명 목록 + 상품 (코드류는 내보내지 않음)
    if (action === "public") {
      return json_({
        ok: true,
        branches: readBranches_().map(function (b) { return { name: b.name, active: b.active }; }),
        products: readProducts_()
      });
    }
    // 학생 주문 페이지용: 주문코드로 지점 확인 + 상품
    if (action === "student") {
      var br = findBranchByCode_(p.b, "code");
      if (!br || !br.active) return json_({ ok: false, error: "유효하지 않은 주문 링크입니다. 학원에 문의해 주세요." });
      return json_({ ok: true, branchName: br.name, products: readProducts_() });
    }
    // 지점 취합·승인 페이지용: 관리코드로 자기 지점 주문만
    if (action === "branch") {
      var brv = findBranchByCode_(p.k, "viewKey");
      if (!brv) return json_({ ok: false, error: "유효하지 않은 관리 링크입니다." });
      var mine = readOrders_().filter(function (o) { return o.branch === brv.name; });
      return json_({ ok: true, branchName: brv.name, orders: mine });
    }
    // 본사 관리자용 전체
    if (action === "all") {
      if (!checkToken_(p.token)) return json_({ ok: false, error: "토큰이 올바르지 않습니다." });
      return json_({ ok: true, branches: readBranches_(), products: readProducts_(), orders: readOrders_() });
    }
    return json_({ ok: false, error: "알 수 없는 action" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** POST: 주문 접수·승인·상태 변경·지점/상품 저장 (본문은 text/plain JSON) */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    // 지점 담당자 일괄 주문
    if (action === "order") return addOrder_(body.data, "지점");

    // 학생 개별 주문: 주문코드로 지점을 서버에서 확정 (지점 위조 방지)
    if (action === "studentOrder") {
      var br = findBranchByCode_(body.data && body.data.code, "code");
      if (!br || !br.active) return json_({ ok: false, error: "유효하지 않은 주문 링크입니다." });
      body.data.branch = br.name;
      return addOrder_(body.data, "학생");
    }

    // 지점의 승인/반려: 관리코드로 인증, 자기 지점 주문만, 허용 상태만
    if (action === "branchSetStatus") {
      var brv = findBranchByCode_(body.k, "viewKey");
      if (!brv) return json_({ ok: false, error: "유효하지 않은 관리 링크입니다." });
      if (BRANCH_STATUSES.indexOf(body.status) < 0) return json_({ ok: false, error: "지점이 변경할 수 없는 상태입니다." });
      return setStatus_(body.orderId, body.product, body.status, brv.name);
    }

    // 이하 본사 관리자 전용
    if (!checkToken_(body.token)) return json_({ ok: false, error: "토큰이 올바르지 않습니다." });
    if (action === "setStatus") return setStatus_(body.orderId, body.product, body.status, null);
    if (action === "saveBranches") return saveBranches_(body.data);
    if (action === "saveProducts") {
      return saveList_(SHEET_PRODUCTS, ["상품명", "단가", "판매"], body.data.map(function (p) {
        return [p.name, Number(p.price) || 0, !!p.active];
      }));
    }
    return json_({ ok: false, error: "알 수 없는 action" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function addOrder_(data, kind) {
  if (!data || !data.branch || !data.items || !data.items.length) {
    return json_({ ok: false, error: "주문 데이터가 비어 있습니다." });
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  var now = new Date();
  var stamp = Utilities.formatDate(now, "Asia/Seoul", "yyMMdd-HHmmss");
  var orderId = (kind === "학생" ? "S" : "B") + stamp + "-" + Math.floor(Math.random() * 900 + 100);
  var iso = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd'T'HH:mm:ss");
  var products = {};
  readProducts_().forEach(function (p) { products[p.name] = p.price; });

  data.items.forEach(function (it) {
    var qty = Math.max(0, parseInt(it.qty, 10) || 0);
    if (!qty) return;
    var unit = products.hasOwnProperty(it.product) ? products[it.product] : (Number(it.price) || 0);
    sheet.appendRow([
      orderId, iso, String(data.branch), String(data.manager || ""), String(data.phone || ""),
      String(it.product), qty, unit, qty * unit, String(data.memo || ""), "접수", kind
    ]);
  });
  return json_({ ok: true, orderId: orderId });
}

/** 상태 변경. branchName이 주어지면 해당 지점 주문만 변경 가능 */
function setStatus_(orderId, product, status, branchName) {
  if (STATUSES.indexOf(status) < 0) return json_({ ok: false, error: "허용되지 않은 상태값" });
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  var values = sheet.getDataRange().getValues();
  var updated = 0;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) !== String(orderId)) continue;
    if (product && String(values[i][5]) !== String(product)) continue;
    if (branchName && String(values[i][2]) !== String(branchName)) continue;
    sheet.getRange(i + 1, 11).setValue(status);
    updated++;
  }
  return updated ? json_({ ok: true, updated: updated }) : json_({ ok: false, error: "해당 주문을 찾지 못했습니다." });
}

/** 지점 저장: 기존 지점의 주문코드·관리코드는 유지, 새 지점은 생성 */
function saveBranches_(list) {
  var existing = {};
  readBranches_().forEach(function (b) { existing[b.code] = b; });
  var rows = (list || []).filter(function (b) { return b.name && String(b.name).trim(); }).map(function (b) {
    var keep = b.code && existing[b.code];
    return [String(b.name).trim(), !!b.active, keep ? b.code : genCode_(), keep ? existing[b.code].viewKey : genCode_()];
  });
  return saveList_(SHEET_BRANCHES, BRANCH_HEADER, rows);
}

function saveList_(sheetName, header, rows) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.clear();
  sheet.appendRow(header);
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  return json_({ ok: true, count: rows.length });
}

function findBranchByCode_(code, field) {
  if (!code) return null;
  var hit = null;
  readBranches_().forEach(function (b) { if (b[field] === String(code)) hit = b; });
  return hit;
}

function readBranches_() {
  return readRows_(SHEET_BRANCHES).map(function (r) {
    return {
      name: String(r[0]),
      active: r[1] === true || String(r[1]).toUpperCase() === "TRUE",
      code: String(r[2] || ""),
      viewKey: String(r[3] || "")
    };
  }).filter(function (b) { return b.name; });
}

function readProducts_() {
  return readRows_(SHEET_PRODUCTS).map(function (r) {
    return { name: String(r[0]), price: Number(r[1]) || 0, active: r[2] === true || String(r[2]).toUpperCase() === "TRUE" };
  }).filter(function (p) { return p.name; });
}

function readOrders_() {
  return readRows_(SHEET_ORDERS).map(function (r) {
    return {
      orderId: String(r[0]),
      timestamp: r[1] instanceof Date ? Utilities.formatDate(r[1], "Asia/Seoul", "yyyy-MM-dd'T'HH:mm:ss") : String(r[1]),
      branch: String(r[2]), manager: String(r[3]), phone: String(r[4]),
      product: String(r[5]), qty: Number(r[6]) || 0, unitPrice: Number(r[7]) || 0,
      amount: Number(r[8]) || 0, memo: String(r[9]), status: String(r[10]) || "접수",
      kind: String(r[11] || "지점")
    };
  }).filter(function (o) { return o.orderId; });
}

function readRows_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}
