/** @OnlyCurrentDoc */
// Dòng trên giới hạn quyền của script vào ĐÚNG file Sheet này (không xin quyền mọi Sheet của bạn). Đừng xoá.

/**
 * NHẬN DỮ LIỆU TỪ FORM TRÊN TRANG -> GOOGLE SHEET + TAB DASHBOARD
 *
 * Cài một lần:
 *  1. Mở Google Sheet sẽ chứa dữ liệu > Extensions > Apps Script > xoá code mẫu, dán toàn bộ file này > Save.
 *  2. Ô chọn hàm ở thanh trên để `setup` > Run > cho phép quyền truy cập.
 *     (Hiện "Google hasn't verified this app": bấm Advanced > Go to ... - đây là script của chính bạn.)
 *  3. Deploy > New deployment > Select type: Web app
 *     Execute as: Me  /  Who has access: Anyone  > Deploy > copy URL đuôi /exec.
 *  4. Dán URL đó vào biến FORM_ENDPOINT trên trang.
 *
 * Sửa code này sau khi đã deploy: Deploy > Manage deployments > bút chì > Version: New version > Deploy.
 * Thiếu bước đó thì URL cũ vẫn chạy code cũ.
 * Đổi phần dashboard trong CẤU HÌNH: xoá tab Dashboard rồi chạy lại `setup`.
 */

/* ======================= CẤU HÌNH - sửa theo form của bạn ======================= */
var FORM = {
  dataTab: 'Lead RecruitPlan', // tên tab chứa các lượt gửi
  fields: [
    { key: 'ho_ten', label: 'Họ tên' },
    { key: 'sdt', label: 'Số điện thoại' },
    { key: 'email', label: 'Email' },
    { key: 'cong_ty', label: 'Công ty' },
    { key: 'quy_mo', label: 'Quy mô nhân sự' },
    { key: 'quan_tam', label: 'Quan tâm' },
    { key: 'dong_y', label: 'Đồng ý' }
  ],
  dashboard: [
    { type: 'count', label: 'Tổng lead' },
    { type: 'today', label: 'Hôm nay' },
    { type: 'days', days: 7, label: '7 ngày gần nhất' },
    { type: 'by', field: 'utm_source', label: 'Theo nguồn', chart: 'bar' },
    { type: 'by', field: 'quan_tam', label: 'Theo mối quan tâm', chart: 'pie' },
    { type: 'by', field: 'quy_mo', label: 'Theo quy mô nhân sự', chart: 'bar' },
    { type: 'per_day', label: 'Theo ngày', chart: 'column' },
    { type: 'latest', n: 10, fields: ['ho_ten', 'sdt', 'quan_tam'], label: '10 lead mới nhất' }
  ]
};
/* ============================ HẾT PHẦN CẤU HÌNH ============================ */

var DASH = 'Dashboard';
var HONEYPOT = 'website'; // ô ẩn chống bot - trùng name="website" trên form
// Cột hệ thống: thời điểm đứng đầu, nguồn truy cập + trang đứng cuối.
var COLS = [['Thời điểm', 'thoi_diem']]
  .concat(FORM.fields.map(function (f) { return [f.label, f.key]; }))
  .concat([['Nguồn', 'utm_source'], ['Kênh', 'utm_medium'], ['Chiến dịch', 'utm_campaign'], ['Trang', 'trang']]);
var KEY_OF = {}, IS_NUMBER = {};
COLS.forEach(function (c) { KEY_OF[c[0]] = c[1]; });
FORM.fields.forEach(function (f) { if (f.type === 'number') IS_NUMBER[f.key] = true; });

// Hàm đầu tiên trong file = hàm hiện sẵn ở ô chọn hàm của Apps Script, nên setup đứng đầu.
// Chạy tay một lần lúc cài (để cấp quyền + tạo sẵn 2 tab). Gọi lại không sao, không tạo trùng.
function setup() {
  ensureSetup_(true);
  SpreadsheetApp.getActiveSpreadsheet().toast('Đã có tab ' + FORM.dataTab + ' và Dashboard. Bước tiếp: Deploy > New deployment.');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // chặn 2 lượt gửi cùng lúc ghi đè nhau
    var b = JSON.parse(e.postData.contents); // trang gửi text/plain nhưng nội dung vẫn là JSON
    if (!b || typeof b !== 'object' || Array.isArray(b)) throw new Error('bad body');
    if (b[HONEYPOT]) return out_({ ok: true }); // bot điền ô ẩn -> bỏ, không lưu
    delete b[HONEYPOT];
    var sh = ensureSetup_();
    var keys = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return KEY_OF[h] || h; });
    Object.keys(b).forEach(function (k) {
      // Link nhận dữ liệu để công khai: chỉ nhận name đúng quy ước (chữ không dấu, số, gạch dưới), tối đa 60 cột,
      // để không ai chèn được công thức hay hàng nghìn cột rác vào dòng tiêu đề.
      if (!/^[A-Za-z0-9_]{1,40}$/.test(k)) return;
      if (keys.indexOf(k) === -1 && keys.length < 60) {
        keys.push(k);
        sh.getRange(1, keys.length).setValue(k);
      }
    });
    var row = keys.map(function (k) {
      if (k === 'thoi_diem') return new Date();
      if (k === 'utm_source' && !b.utm_source) return '(trực tiếp)';
      return cell_(k, b[k]);
    });
    sh.appendRow(row);
    return out_({ ok: true });
  } catch (err) {
    console.error(err);
    return out_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

// Mở URL /exec trên trình duyệt, thấy dòng này = đường nhận dữ liệu đang sống. Không bao giờ trả dữ liệu ra đây.
function doGet() {
  return out_({ ok: true, ping: 'form-endpoint' });
}

function ensureSetup_(forSetup) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FORM.dataTab);
  if (!sh) {
    var first = ss.getSheets()[0];
    sh = (ss.getSheets().length === 1 && first.getLastRow() === 0) ? first.setName(FORM.dataTab) : ss.insertSheet(FORM.dataTab);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLS.map(function (c) { return c[0]; }));
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  } else {
    // Ô mới thêm vào FORM.fields sau khi đã có dữ liệu: thêm cột ở CUỐI (không chèn giữa, không đụng cột cũ).
    var have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return KEY_OF[h] || h; });
    COLS.forEach(function (c) {
      if (have.indexOf(c[1]) === -1) { have.push(c[1]); sh.getRange(1, have.length).setValue(c[0]); }
    });
  }
  if (!ss.getSheetByName(DASH)) {
    // Lỗi cấu hình Dashboard: chạy tay setup thì báo lỗi; lúc nhận lượt gửi thì vẫn lưu lượt đó, chỉ ghi lỗi vào log.
    try { buildDashboard_(ss); } catch (err) { if (forSetup) throw err; console.error(err); }
  }
  return sh;
}

// Ô tick không đặt value gửi "on" -> ghi "Có". Ô số -> ghi số. Còn lại ghi dạng chữ: giữ số 0 đầu SĐT,
// và chặn chuỗi mở đầu bằng = + - @ bị Sheet hiểu thành công thức.
function cell_(k, v) {
  if (v === undefined || v === null || v === '') return '';
  if (v === 'on') return 'Có';
  if (IS_NUMBER[k]) {
    var n = Number(String(v).replace(',', '.'));
    if (isFinite(n)) return n;
  }
  return "'" + String(v).slice(0, 2000);
}

function buildDashboard_(ss) {
  // Sheet tiếng Việt (và mọi locale dùng dấu phẩy thập phân) ngăn tham số công thức bằng ";" chứ không phải ",".
  // setFormula đọc theo locale của Sheet, viết sai dấu là ra #ERROR! - nên đổi dấu theo locale trước khi ghi.
  var SEP = Number(1.5).toLocaleString(ss.getSpreadsheetLocale().replace('_', '-')).indexOf(',') > -1 ? ';' : ',';
  var f = function (formula) { return fx_(formula, SEP); };
  var T = "'" + FORM.dataTab + "'!";
  // Đọc vị trí cột từ dòng tiêu đề THẬT của tab dữ liệu (cột thêm sau nằm ở cuối, không theo thứ tự cấu hình).
  var head = ss.getSheetByName(FORM.dataTab).getRange(1, 1, 1, ss.getSheetByName(FORM.dataTab).getLastColumn()).getValues()[0];
  var keys = head.map(function (h) { return KEY_OF[h] || h; });
  var col = function (key) {
    var i = keys.indexOf(key);
    if (i < 0) throw new Error('Dashboard trỏ tới ô "' + key + '" không có trong FORM.fields');
    return colLetter_(i + 1);
  };
  var label = function (key) { return String(head[keys.indexOf(key)]).replace(/'/g, ''); };
  var LAST = colLetter_(head.length);
  var ALL = T + 'A2:' + LAST;
  // Kiểm cấu hình TRƯỚC khi tạo tab, để lỗi không để lại tab Dashboard dở dang.
  FORM.dashboard.forEach(function (x) { [x.field, x.sum].concat(x.fields || []).forEach(function (k) { if (k) col(k); }); });
  var d = ss.insertSheet(DASH, 0);

  d.getRange('A1').setValue('Dashboard - ' + FORM.dataTab).setFontSize(16).setFontWeight('bold');
  d.getRange('A2').setValue('Tự cập nhật mỗi khi có lượt gửi mới. Đừng sửa ô có công thức.').setFontColor('#666666');

  var tiles = FORM.dashboard.filter(function (x) { return ['count', 'today', 'days', 'sum', 'avg', 'progress'].indexOf(x.type) > -1; });
  tiles.forEach(function (x, i) {
    var c = colLetter_(i + 1);
    var v;
    if (x.type === 'count') v = '=COUNTA(' + T + 'A2:A)';
    if (x.type === 'today') v = '=COUNTIFS(' + T + 'A2:A,">="&TODAY())';
    if (x.type === 'days') v = '=COUNTIFS(' + T + 'A2:A,">="&(TODAY()-' + ((x.days || 7) - 1) + '))';
    if (x.type === 'sum') v = '=SUM(' + T + col(x.field) + '2:' + col(x.field) + ')';
    if (x.type === 'avg') v = '=IFERROR(ROUND(AVERAGE(' + T + col(x.field) + '2:' + col(x.field) + '),1),"-")';
    if (x.type === 'progress') v = (x.field ? '=SUM(' + T + col(x.field) + '2:' + col(x.field) + ')' : '=COUNTA(' + T + 'A2:A)') + '&" / ' + x.target + '"';
    d.getRange(c + '4').setValue(x.label).setFontColor('#666666').setWrap(true);
    d.getRange(c + '5').setFormula(f(v)).setFontSize(20).setFontWeight('bold').setHorizontalAlignment('left');
    d.setColumnWidth(i + 1, 150);
  });

  var row = tiles.length ? 8 : 4;
  FORM.dashboard.forEach(function (x) {
    if (['by', 'per_day', 'latest'].indexOf(x.type) < 0) return;
    var used, q;
    if (x.type === 'by') {
      var X = col(x.field), lim = x.limit || 15;
      var order = x.order === 'value' ? X : (x.sum ? 'sum(' + col(x.sum) + ')' : 'count(A)');
      var dir = x.order === 'value' ? 'asc' : 'desc';
      var agg = x.sum ? 'sum(' + col(x.sum) + ')' : 'count(A)';
      var aggLabel = x.sum ? 'Tổng ' + label(x.sum) : 'Số lượt';
      q = '=IFERROR(QUERY(' + ALL + ', "select ' + X + ', ' + agg + ' where A is not null group by ' + X +
        ' order by ' + order + ' ' + dir + ' limit ' + lim + " label " + X + " '" + label(x.field) + "', " + agg + " '" + aggLabel + "'\", 0), \"Chưa có dữ liệu\")";
      used = lim + 1;
    }
    if (x.type === 'per_day') {
      q = '=IFERROR(QUERY(QUERY(' + T + 'A2:A, "select toDate(A), count(A) where A is not null group by toDate(A)' +
        " order by toDate(A) desc limit 14 label toDate(A) 'Ngày', count(A) 'Số lượt'\", 0), \"select * order by Col1\", 1), \"Chưa có dữ liệu\")";
      used = 15;
      d.getRange('A' + (row + 2) + ':A' + (row + 15)).setNumberFormat('dd/MM');
    }
    if (x.type === 'latest') {
      var cs = ['A'].concat((x.fields || []).map(col));
      q = '=IFERROR(QUERY(' + T + 'A1:' + LAST + ', "select ' + cs.join(', ') + ' where A is not null order by A desc limit ' + (x.n || 10) + '", 1), "Chưa có dữ liệu")';
      used = (x.n || 10) + 1;
      d.getRange('A' + (row + 2) + ':A' + (row + 1 + (x.n || 10))).setNumberFormat('dd/MM/yyyy HH:mm');
    }
    d.getRange('A' + row).setValue(x.label).setFontWeight('bold');
    d.getRange('A' + (row + 1)).setFormula(f(q));
    if (x.chart && x.type !== 'latest') {
      var TYPES = { column: Charts.ChartType.COLUMN, bar: Charts.ChartType.BAR, pie: Charts.ChartType.PIE, line: Charts.ChartType.LINE };
      d.insertChart(d.newChart()
        .setChartType(TYPES[x.chart] || Charts.ChartType.COLUMN)
        .addRange(d.getRange('A' + (row + 1) + ':B' + (row + used)))
        .setNumHeaders(1)
        .setPosition(row, 5, 0, 0)
        .setOption('title', x.label)
        .setOption('legend', { position: x.chart === 'pie' ? 'right' : 'none' })
        .build());
      used = Math.max(used, 19); // biểu đồ cao khoảng 19 dòng - chừa đủ để khối sau không bị che
    }
    row += used + 3;
  });
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Đổi dấu "," ngăn tham số sang dấu của Sheet, giữ nguyên dấu phẩy nằm trong chuỗi "...".
function fx_(formula, sep) {
  if (sep === ',') return formula;
  var res = '', inStr = false;
  for (var i = 0; i < formula.length; i++) {
    var c = formula.charAt(i);
    if (c === '"') inStr = !inStr;
    res += (c === ',' && !inStr) ? sep : c;
  }
  return res;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* CÁCH VIẾT DASHBOARD (mỗi dòng trong FORM.dashboard)
 *  { type: 'count', label }                          tổng số lượt gửi
 *  { type: 'today', label }                          số lượt hôm nay
 *  { type: 'days', days: 7, label }                  số lượt N ngày gần nhất
 *  { type: 'sum', field, label }                     cộng một ô số (vd số lượng)
 *  { type: 'avg', field, label }                     trung bình một ô số (vd điểm 1-5)
 *  { type: 'progress', target: 50, label, field }    "đã có / chỉ tiêu"; field (ô số, vd so_nguoi) = cộng ô đó thay vì đếm lượt
 *  { type: 'by', field, label, chart, limit, sum, order }
 *        đếm theo từng giá trị của một ô (vd theo nguồn, theo buổi, theo gói); sum: cộng ô số thay vì đếm;
 *        order: 'value' = xếp theo giá trị (hợp với điểm 1-5); chart: 'bar' | 'column' | 'pie' | 'line'
 *  { type: 'per_day', label, chart }                 số lượt theo ngày (14 ngày có dữ liệu gần nhất)
 *  { type: 'latest', n: 10, fields: [...], label }   N lượt mới nhất, các cột chọn
 * field / fields dùng key trong FORM.fields, hoặc utm_source, utm_medium, utm_campaign, trang.
 */
