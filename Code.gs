// 27EV 심층 설문 저장용 Apps Script
// 배포 순서
// 1. regular0529 계정으로 새 구글 스프레드시트를 만든다 → 확장 프로그램 > Apps Script
// 2. 이 파일 내용을 Code.gs에 붙여 넣고 저장
// 3. 배포 > 새 배포 > 유형: 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자 → 배포 → 권한 승인
// 4. 나온 웹 앱 URL을 index.html의 ENDPOINT에 붙여 넣는다
// 5. 이 코드를 고치면 배포 > 배포 관리 > 수정 > 버전: 새 버전 (URL 유지)
//
// 시트
// - responses: 학번(sid)당 한 행. '제출'할 때만 갱신(덮어쓰기)
// - log: 제출·임시저장 모두 한 줄씩 계속 쌓임. 지우거나 고치지 않는 이력
// - saved_at은 서버 시각이라 응답자가 조작할 수 없다

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = Object.assign({saved_at: new Date()}, JSON.parse(e.postData.contents));
    const ss = SpreadsheetApp.getActive();
    put(sheet(ss, 'log'), d, false);
    if (d.status === '제출') put(sheet(ss, 'responses'), d, true);
    return out({ok: true, saved_at: d.saved_at.toISOString()});
  } catch (err) {
    return out({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

function sheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// 처음 보는 열은 오른쪽에 추가하므로 문항을 늘려도 코드를 고칠 필요가 없다
function put(sh, d, upsert) {
  let head = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  const add = Object.keys(d).filter(k => !head.includes(k));
  if (add.length) {
    sh.getRange(1, head.length + 1, 1, add.length).setValues([add.map(literal)]);
    head = head.concat(add);
  }
  const row = head.map(k => k in d ? literal(d[k]) : '');
  let n = sh.getLastRow() + 1;
  if (upsert && n > 2) {
    const ids = sh.getRange(2, head.indexOf('sid') + 1, n - 2, 1).getDisplayValues().flat();
    const i = ids.indexOf(String(d.sid));
    if (i >= 0) n = i + 2;
  }
  // 문자는 텍스트 형식으로 넣어 학번 앞자리 0을 지킨다. 0/1은 합산할 수 있게 숫자로 둔다
  const formats = row.map(x => x instanceof Date ? 'yyyy-mm-dd hh:mm:ss' : typeof x === 'number' ? '0' : '@');
  sh.getRange(n, 1, 1, head.length).setNumberFormats([formats]).setValues([row]);
}

// 텍스트 형식이어도 =로 시작하면 수식으로 실행되므로(실측) 글자 표시(')를 붙인다
function literal(x) {
  return typeof x === 'string' && /^\s*=/.test(x) ? "'" + x : x;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
