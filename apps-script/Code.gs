// 구글 시트 > 확장 프로그램 > Apps Script 에 붙여넣고
// 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자) → URL을 index.html CONFIG.applyEndpoint 에 입력

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var phone = String(d.phone || '').replace(/\D/g, '');
  var name = String(d.name || '').trim();
  var rows = sheet.getDataRange().getValues(); // 열: 신청일시, 이름, 생년월일, 핸드폰, 참여가족, 문자수신
  var found = rows.filter(function (r) { return r[1] === name && String(r[3]).replace(/\D/g, '') === phone; })[0];

  if (d.action === 'check') {
    return json({ found: !!found }); // 신청 여부만 알려줌
  }
  if (!name || !/^01\d{8,9}$/.test(phone)) return json({ ok: false, message: '이름과 핸드폰번호를 확인해주세요' });
  if (found) return json({ ok: false, message: '이미 신청되어 있어요' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  if (sheet.getLastRow() === 0) sheet.appendRow(['신청일시', '이름', '생년월일', '핸드폰', '참여가족', '문자수신']);
  // 앞에 ' 를 붙여 수식 주입 방지 + 0으로 시작하는 번호 유지
  sheet.appendRow([new Date(), "'" + name, "'" + d.birth, "'" + phone, "'" + d.family, d.sms === '네' ? '네' : '아니오']);
  lock.releaseLock();
  return json({ ok: true });
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
