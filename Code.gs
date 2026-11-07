/**
 * How Computers Evolved · Hộp nhận bài (ITTR330336)
 * Dán toàn bộ file này vào Tiện ích → Apps Script của Google Sheet "How Computers Evolved · Bài nộp".
 * Chạy hàm setup() một lần, rồi Triển khai → Tùy chọn triển khai mới → Ứng dụng web
 * (Thực thi với tư cách: Tôi · Người có quyền truy cập: Bất kỳ ai).
 */
const CFG = {
  KEY: 'ittr-hce-76c8820e',                                   // mã lớp – phải trùng với config.json của app
  LESSON: 'hce-how-computers-evolved',
  CLASS_FOLDERS: {                                  // thư mục Drive cho từng lớp
    'Lớp thứ 3': '1PoQSPM4xuUrgiTF-UKeO-EATK3QW6Rnn',
    'Lớp thứ 5': '1qOVq2maJw5U9Lf-rxOAStdV-0K3JjCn8'
  },
  ROOT_FOLDER: '1OWnDqOvJtWzN_2I7WbklWdbuLVcBSav1', // dùng khi lớp không khớp
  TZ: 'Asia/Ho_Chi_Minh'
};
const ANS = {"G1": [["abacus"], ["subtraction"], ["Pascal"], ["roots"], ["1820"], ["commercially"], ["Difference"], ["Census"], ["Tabulating"], ["1924"], ["Watson"], ["765,000", "765000"], ["miles"], ["typewriter"], ["tubes"], ["1000", "1,000", "one thousand"], ["20", "twenty"], ["2400", "2,400"], ["1951"], ["Eisenhower", "Dwight Eisenhower"], ["11", "eleven"], ["1988"], ["1991"]], "G2": [["1950"], ["cores"], ["702"], ["1969"], ["semiconductor", "semiconductors"], ["tape"], ["Honeywell"], ["1977"], ["1958"], ["1108"], ["1980s", "1980"], ["VLSI"], ["transistors"], ["Apple"], ["intelligence"]], "TF": ["F", "F", "T", "F", "NG", "T", "F", "T", "NG", "F"], "MAP": ["the abacus", "a calculator", "From … to …", "go back in time", "…, thereby …", "Inspired by …", "commercially available", "It was … who …", "the Difference Engine", "the Analytical Engine", "levers and gears", "a decimal number", "the US Census Bureau", "punched cards", "With the help of …", "fully automatic", "Later, …", "was renamed", "according to the video", "miles / feet", "an electric typewriter", "is regarded as", "an electromechanical analytical machine", "fully digital", "nuclear weapons", "vacuum tubes", "… times faster than", "a trajectory", "was said to", "store a program", "one of the most …", "correctly predicted", "paved the way for", "ring-shaped ferrite cores", "bulky", "semiconductor digital elements", "magnetic tape", "integrated circuits (ICs)", "microprocessors", "As a result, …", "very-large-scale integration (VLSI)", "transistors", "circuit density", "…, which led to …", "personal computers", "one of the first …", "a graphical user interface (GUI)", "a luggable computer", "a clamshell design", "resolution", "By …", "artificial intelligence"]};
const NTR = 70, NFC = 8;

function setup() {
  const ss = SpreadsheetApp.getActive();
  const head = ['Thời gian nhận', 'Lớp', 'MSSV', 'Họ tên', 'Nhóm', 'Lần nộp', 'Mới nhất',
    'Task 1 (/' + ANS.G1.length + ')', 'Task 2 (/' + ANS.G2.length + ')', 'Task 3 (/' + ANS.TF.length + ')', 'Concept map (/' + ANS.MAP.length + ')',
    'Số câu dịch (/' + NTR + ')', 'Thời gian học (phút)', 'Đồng ý ghi quá trình', 'File bài làm', 'Thiết bị'];
  for (let i = 1; i <= NTR; i++) head.push('Câu ' + i);
  for (let i = 1; i <= NFC; i++) head.push('Nguồn ' + i);
  head.push('Task 1 – trả lời', 'Task 2 – trả lời', 'Task 3 – trả lời', 'Map – trả lời', 'Map – tiêu đề/nhánh');
  sheet_(ss, 'Bài nộp', head);
  sheet_(ss, 'Quá trình dịch', ['Thời gian nhận', 'Lớp', 'MSSV', 'Họ tên', 'Lần nộp', 'Câu', 'Thời gian soạn (giây)', 'Số lần sửa', 'Số lần nghe lại', 'Sửa sau khi nghe', 'Độ dài bản cuối', 'Các bản nháp']);
  sheet_(ss, 'Log', ['Thời gian nhận', 'Lớp', 'MSSV', 'Họ tên', 'Thiết bị', 'Thời điểm (máy SV)', 'Sự kiện', 'Chi tiết']);
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (s1 && ss.getSheets().length > 3) ss.deleteSheet(s1);
  SpreadsheetApp.getUi && toastSafe_('Đã tạo các trang tính. Tiếp theo: Triển khai → Ứng dụng web.');
}
function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#dce6f0');
  sh.setFrozenRows(1);
  return sh;
}
function toastSafe_(m) { try { SpreadsheetApp.getActive().toast(m); } catch (e) {} }

function doGet() { return ContentService.createTextOutput('How Computers Evolved – hộp nhận bài đang hoạt động.'); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const d = JSON.parse(e.postData.contents);
    if (d.key !== CFG.KEY) return out_({ ok: false, err: 'key' });
    if (d.action === 'submit') return out_(submit_(d));
    if (d.action === 'log') return out_(log_(d));
    return out_({ ok: false, err: 'action' });
  } catch (err) {
    return out_({ ok: false, err: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
const norm_ = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
function clean_(s, n) { return String(s == null ? '' : s).slice(0, n || 5000); }
function safe_(row) { return row.map(x => (typeof x === 'string' && /^[=+\-@]/.test(x)) ? "'" + x : x); }

function submit_(d) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Bài nộp') || (setup(), ss.getSheetByName('Bài nộp'));
  const p = d.profile || {}, id = clean_(p.id, 20).trim(), name = clean_(p.name, 80).trim(), cls = clean_(p.cls, 30);
  if (!/^\d{6,10}$/.test(id) || !name) return { ok: false, err: 'profile' };
  const now = new Date();
  // previous attempts by this student
  const last = sh.getLastRow();
  let attempt = 1;
  if (last > 1) {
    const ids = sh.getRange(2, 3, last - 1, 1).getValues();
    const latest = sh.getRange(2, 7, last - 1, 1).getValues();
    ids.forEach((r, i) => { if (String(r[0]) === id) { attempt++; if (latest[i][0] === '✓') sh.getRange(i + 2, 7).setValue(''); } });
  }
  const g1 = d.g1 || [], g2 = d.g2 || [], tf = d.tf || {}, map = (d.map && d.map.slots) || {};
  const okG = (A, arr) => A.map((alts, i) => alts.map(norm_).includes(norm_(arr[i])));
  const k1 = okG(ANS.G1, g1), k2 = okG(ANS.G2, g2);
  const kt = ANS.TF.map((k, i) => tf[i] === k), km = ANS.MAP.map((a, i) => map[i + 1] === a);
  const cnt = a => a.filter(Boolean).length;
  const tr = d.tr || {}, fc = d.fc || {};
  const nTr = Object.keys(tr).filter(k => String(tr[k]).trim()).length;
  // Google Doc in the class folder
  const folder = DriveApp.getFolderById(CFG.CLASS_FOLDERS[cls] || CFG.ROOT_FOLDER);
  const title = id + ' – ' + name + ' – ' + Utilities.formatDate(now, CFG.TZ, 'yyyy-MM-dd HH:mm') + (attempt > 1 ? ' (lần ' + attempt + ')' : '');
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  body.appendParagraph('How Computers Evolved · Bài nộp').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Họ tên: ' + name + '   MSSV: ' + id + '   Lớp: ' + cls + '   Nhóm: ' + clean_(p.group, 20));
  body.appendParagraph('Nhận lúc: ' + Utilities.formatDate(now, CFG.TZ, 'dd/MM/yyyy HH:mm') + ' · Lần nộp: ' + attempt + ' · Thời gian học: ' + Math.round((d.activeMs || 0) / 60000) + ' phút');
  body.appendParagraph('Điểm tự động – Task 1: ' + cnt(k1) + '/' + k1.length + ' · Task 2: ' + cnt(k2) + '/' + k2.length + ' · Task 3: ' + cnt(kt) + '/' + kt.length + ' · Concept map: ' + cnt(km) + '/' + km.length);
  body.appendParagraph('Listening').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const lrows = [['Task', '#', 'Sinh viên', 'Đáp án']];
  ANS.G1.forEach((a, i) => lrows.push(['1', String(i + 1), (g1[i] || '') + (k1[i] ? ' ✓' : ' ✗'), a.join(' / ')]));
  ANS.G2.forEach((a, i) => lrows.push(['2', String(i + 1), (g2[i] || '') + (k2[i] ? ' ✓' : ' ✗'), a.join(' / ')]));
  ANS.TF.forEach((a, i) => lrows.push(['3', String(i + 1), (tf[i] || '') + (kt[i] ? ' ✓' : ' ✗'), a]));
  body.appendTable(lrows);
  body.appendParagraph('Concept map').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Tiêu đề: ' + clean_(d.map && d.map.title, 200) + ' · Nhánh: ' + ((d.map && d.map.br) ? Object.values(d.map.br).join(' | ') : ''));
  const mrows = [['#', 'Sinh viên', 'Đáp án']];
  ANS.MAP.forEach((a, i) => mrows.push([String(i + 1), (map[i + 1] || '') + (km[i] ? ' ✓' : ' ✗'), a]));
  body.appendTable(mrows);
  body.appendParagraph('Bản dịch Việt → Anh (' + nTr + '/' + NTR + ' câu)').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const trows = [['#', 'Tiếng Việt', 'Bản dịch của sinh viên']];
  (d.vi || []).forEach((v, i) => trows.push([String(i + 1), String(v), String(tr[i + 1] || '')]));
  body.appendTable(trows);
  body.appendParagraph('Source check – translator’s notes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  for (let i = 0; i < NFC; i++) body.appendListItem((i + 1) + '. ' + clean_(fc[i], 3000));
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder);
  // summary row
  const row = [now, cls, id, name, clean_(p.group, 20), attempt, '✓', cnt(k1), cnt(k2), cnt(kt), cnt(km), nTr,
    Math.round((d.activeMs || 0) / 60000), d.consent ? 'Có' : 'Không', file.getUrl(), clean_(d.device, 40)];
  for (let i = 1; i <= NTR; i++) row.push(clean_(tr[i], 3000));
  for (let i = 0; i < NFC; i++) row.push(clean_(fc[i], 3000));
  row.push(g1.join(' | '), g2.join(' | '), ANS.TF.map((_, i) => (i + 1) + '-' + (tf[i] || '')).join(' '),
    ANS.MAP.map((_, i) => (i + 1) + '. ' + (map[i + 1] || '')).join(' | '), clean_(d.map && d.map.title, 200) + ' || ' + ((d.map && d.map.br) ? Object.values(d.map.br).join(' | ') : ''));
  sh.appendRow(safe_(row));
  // per-sentence process statistics (only with consent)
  if (d.consent && d.stats) {
    const ps = ss.getSheetByName('Quá trình dịch');
    const rows = Object.keys(d.stats).map(n => { const s = d.stats[n] || {};
      return [now, cls, id, name, attempt, Number(n), Math.round((s.focus || 0) / 1000), s.edits || 0, s.replays || 0, s.editsAfter || 0, String(tr[n] || '').length, clean_((s.drafts || []).join(' ⟶ '), 20000)]; });
    if (rows.length) ps.getRange(ps.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows.map(safe_));
  }
  return { ok: true, attempt: attempt, doc: file.getUrl() };
}

function log_(d) {
  if (!d.consent) return { ok: true, skipped: true };
  const sh = SpreadsheetApp.getActive().getSheetByName('Log');
  const p = d.profile || {}, now = new Date();
  const ev = (d.events || []).slice(0, 500).map(x => [now, clean_(p.cls, 30), clean_(p.id, 20), clean_(p.name, 80), clean_(d.device, 40),
    new Date(x.t || 0), clean_(x.ev, 40), clean_(JSON.stringify(x.d || {}), 2000)]);
  if (ev.length) sh.getRange(sh.getLastRow() + 1, 1, ev.length, ev[0].length).setValues(ev.map(safe_));
  return { ok: true, n: ev.length };
}
