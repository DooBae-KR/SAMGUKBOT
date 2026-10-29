// 구글 시트 > 확장 프로그램 > Apps Script 에 붙여넣기
// 프로젝트 설정(톱니바퀴) > 스크립트 속성 에 SAVE_KEY 를 추가하세요 (봇 환경변수 SAVE_KEY 와 같은 값)

function _key() { return PropertiesService.getScriptProperties().getProperty('SAVE_KEY'); }
function _out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function _sheet(n) { var ss = SpreadsheetApp.getActive(); return ss.getSheetByName(n) || ss.insertSheet(n); }

function doGet(e) {
  if (!_key() || e.parameter.key !== _key()) return _out({ ok: false, error: 'forbidden' });
  var rows = _sheet('봇_저장').getDataRange().getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === e.parameter.name) return _out({ ok: true, data: JSON.parse(rows[i][1]) });
  }
  return _out({ ok: true, data: null });
}

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return _out({ ok: false, error: 'bad json' }); }
  if (!_key() || b.key !== _key()) return _out({ ok: false, error: 'forbidden' });
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = _sheet('봇_저장'), rows = sh.getDataRange().getValues(), row = -1;
    for (var i = 0; i < rows.length; i++) if (rows[i][0] === b.name) { row = i + 1; break; }
    if (row < 0) row = sh.getLastRow() + 1;
    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    sh.getRange(row, 1, 1, 3).setValues([[b.name, JSON.stringify(b.data), now]]);
    try { _mirror(b.name, b.data, now); } catch (err) {}
    return _out({ ok: true });
  } finally { lock.releaseLock(); }
}

// ---- 봇_인증의 CID(A열)를 '인사' 탭의 CID(UID)와 매핑하는 함수 열(E~H) ----
var HR_GID = 818075343;   // 인사 탭 gid (다르면 수정)
var HR_UID = ['uid', '유저id', '아이디', 'id', 'cid', '캐릭터 id', '캐릭터id'];
var HR_NAME = ['닉네임', '이름', '닉'];
var HR_JOB = ['직업'];
var HR_TITLE = ['직위', '직책', '등급'];

function _colLetter(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function _hrSheet() {
  var ss = SpreadsheetApp.getActive(), all = ss.getSheets();
  for (var i = 0; i < all.length; i++) if (all[i].getSheetId() === HR_GID) return all[i];
  return ss.getSheetByName('인사');
}

// 인사 탭에서 헤더 행을 찾아 열 번호(1부터)를 돌려준다
function _hrCols(hr) {
  var rows = hr.getRange(1, 1, Math.min(hr.getLastRow(), 15), hr.getLastColumn()).getValues();
  function find(H, c) { for (var i = 0; i < H.length; i++) if (c.indexOf(String(H[i]).trim().toLowerCase()) >= 0) return i + 1; return 0; }
  for (var r = 0; r < rows.length; r++) {
    var H = rows[r], u = find(H, HR_UID);
    if (u) return { uid: u, name: find(H, HR_NAME), job: find(H, HR_JOB), title: find(H, HR_TITLE) };
  }
  return null;
}

function _hrLink(sh, n) {
  var hr = _hrSheet(), c = hr && _hrCols(hr);
  sh.getRange(1, 5, 1, 4).setValues([['인사탭 닉네임', '인사탭 직업', '인사탭 직위', '인사탭 매핑']]);
  if (!c || n < 2) return;
  var U = "ARRAYFORMULA('" + hr.getName() + "'!$" + _colLetter(c.uid) + ":$" + _colLetter(c.uid) + '&"")';
  function pick(col) { return col ? '=IFERROR(INDEX(\'' + hr.getName() + "'!$" + _colLetter(col) + ':$' + _colLetter(col) + ', MATCH($A{r}&"", ' + U + ', 0)), "")' : '=""'; }
  var f = [];
  for (var r = 2; r <= n; r++) {
    var row = [pick(c.name), pick(c.job), pick(c.title), '=IF(ISNUMBER(MATCH($A{r}&"", ' + U + ', 0)), "OK", "인사탭에 없음")'];
    f.push(row.map(function (x) { return x.replace(/\{r\}/g, r); }));
  }
  sh.getRange(2, 5, n - 1, 4).setFormulas(f);
}

// 사람이 보기 쉬운 탭
function _mirror(name, data, now) {
  if (name === 'verified') {
    var sh = _sheet('봇_인증'); sh.clear();
    var out = [['캐릭터 ID', '닉네임', '디스코드 ID(인증자)', '갱신']];
    Object.keys(data.uid || {}).forEach(function (k) { out.push([k, String((data.nick || {})[k] || ''), String(data.uid[k]), now]); });
    sh.getRange(1, 1, out.length, 1).setNumberFormat('@');
    sh.getRange(1, 1, out.length, 4).setValues(out);
    _hrLink(sh, out.length);
  } else if (name.indexOf('dex_') === 0) {
    var s3 = _sheet('도감선택');
    if (s3.getLastRow() === 0) s3.appendRow(['CID', '닉네임', '보유 장수 수', '보유 장수', '보유 전법 수', '보유 전법', '갱신']);
    var vals = s3.getRange(1, 1, s3.getLastRow(), 1).getValues(), r = -1;
    for (var j = 1; j < vals.length; j++) if (String(vals[j][0]) === String(data.cid)) { r = j + 1; break; }
    if (r < 0) r = s3.getLastRow() + 1;
    s3.getRange(r, 1).setNumberFormat('@');
    s3.getRange(r, 1, 1, 7).setValues([[String(data.cid), data.name, (data.own || []).length, (data.own || []).join(', '), (data.ownsk || []).length, (data.ownsk || []).join(', '), data.at]]);
  } else if (name === 'leave') {
    var s4 = _sheet('봇_이탈처리'); s4.clear();
    var lg = [['처리 시각', 'CID', '닉네임', '처리', '처리자', '디스코드 ID']];
    (data.log || []).slice().reverse().forEach(function (x) { lg.push([x.t, x.uid, x.nick, x.act, x.by, String(x.did || '')]); });
    s4.getRange(1, 2, lg.length, 1).setNumberFormat('@'); s4.getRange(1, 6, lg.length, 1).setNumberFormat('@');
    s4.getRange(1, 1, lg.length, 6).setValues(lg);
    var pn = Object.keys(data.pending || {}).length;
    s4.getRange(1, 8, 2, 1).setValues([['확인 대기 중'], [pn]]);
  } else if (name === 'jungchul') {
    var s2 = _sheet('봇_정철'); s2.clear();
    var t = (data.sieges || []).map(function (x, i) { return [i + 1, x.level, x.cost, String(x.user)]; });
    s2.getRange(1, 1, 3, 2).setValues([['보유량(기준)', data.stock], ['기준 시각', Utilities.formatDate(new Date(data.at * 1000), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')], ['등록된 공성 수', t.length]]);
    if (t.length) s2.getRange(5, 1, t.length + 1, 4).setValues([['순서', '레벨', '필요량', '등록자 디스코드ID']].concat(t));
  }
}
