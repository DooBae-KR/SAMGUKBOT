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

// 사람이 보기 쉬운 탭
function _mirror(name, data, now) {
  if (name === 'verified') {
    var sh = _sheet('봇_인증'); sh.clear();
    var out = [['캐릭터 ID', '디스코드 ID(인증자)', '갱신']];
    Object.keys(data.uid || {}).forEach(function (k) { out.push([k, String(data.uid[k]), now]); });
    sh.getRange(1, 1, out.length, 3).setValues(out);
  } else if (name === 'jungchul') {
    var s2 = _sheet('봇_정철'); s2.clear();
    var t = (data.timers || []).map(function (x) { return [x.level, x.need, String(x.user)]; });
    s2.getRange(1, 1, 3, 2).setValues([['보유량(기준)', data.stock], ['기준 시각', Utilities.formatDate(new Date(data.at * 1000), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')], ['대기 알림 수', t.length]]);
    if (t.length) s2.getRange(5, 1, t.length + 1, 3).setValues([['레벨', '필요량', '요청자 디스코드ID']].concat(t));
  }
}
