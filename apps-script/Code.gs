// 구글 시트 > 확장 프로그램 > Apps Script 에 붙여넣고
// 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자) → URL을 index.html CONFIG.applyEndpoint 에 입력
//
// 오류 알림: 처리 중 오류가 나면 NOTIFY_EMAIL 로 메일 + 카카오톡 "나에게 보내기"
// 카카오톡 연결 (한 번만):
//   1) 프로젝트 설정 > 스크립트 속성에 KAKAO_REST_KEY (앱 키 > REST API 키) 추가
//      클라이언트 시크릿을 켜 두었다면 KAKAO_CLIENT_SECRET 도 추가
//   2) https://kauth.kakao.com/oauth/authorize?response_type=code&scope=talk_message&client_id=<REST API 키>&redirect_uri=<WEBAPP_URL>
//      에 접속해 로그인·동의하면 doGet 이 토큰을 저장하고 "연결 완료" 메시지를 보냄
var NOTIFY_EMAIL = 'xonic0506@gmail.com';
var WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbwXU4yrpxhV0qmFfu6bnivtkGD4TV73e6xW16GxJ01vdy_Z3rw9uKj93R1v44Hy--_c/exec'; // 카카오 Redirect URI 와 같아야 함
var PAGE_URL = 'https://sonic0506.github.io/sorisaem-2026-yearend/';

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  try {
    return handle(d);
  } catch (err) {
    notify((d.action === 'check' ? '신청확인' : '참가 신청') + ' 처리 중 오류: ' + err);
    return json({ ok: false, message: '처리 중 문제가 생겼어요. 02-824-1411로 연락주세요' });
  }
}

function handle(d) {
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

// 신청자 개인정보는 넣지 않고 오류 내용과 시각만 보냄
function notify(msg) {
  var text = '[송년행사 신청 오류] ' + msg + '\n' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MM/dd HH:mm:ss');
  try { MailApp.sendEmail(NOTIFY_EMAIL, '[송년행사 신청] 오류 발생', text); } catch (e) { console.error(e); }
  try { sendKakao(text); } catch (e) { console.error(e); }
}

function kakaoToken(params) {
  var props = PropertiesService.getScriptProperties();
  params.client_id = props.getProperty('KAKAO_REST_KEY');
  var secret = props.getProperty('KAKAO_CLIENT_SECRET');
  if (secret) params.client_secret = secret;
  return JSON.parse(UrlFetchApp.fetch('https://kauth.kakao.com/oauth/token', { method: 'post', payload: params, muteHttpExceptions: true }).getContentText());
}

function sendKakao(text) {
  var props = PropertiesService.getScriptProperties();
  var refresh = props.getProperty('KAKAO_REFRESH_TOKEN');
  if (!refresh) return; // 아직 연결 안 함
  var t = kakaoToken({ grant_type: 'refresh_token', refresh_token: refresh });
  if (t.refresh_token) props.setProperty('KAKAO_REFRESH_TOKEN', t.refresh_token); // 만료가 가까우면 새로 내려줌
  UrlFetchApp.fetch('https://kapi.kakao.com/v2/api/talk/memo/default/send', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + t.access_token },
    payload: { template_object: JSON.stringify({ object_type: 'text', text: text, link: { web_url: PAGE_URL, mobile_web_url: PAGE_URL } }) }
  });
}

// 카카오 로그인 후 이 웹 앱으로 돌아오면 토큰 저장. 이미 연결돼 있으면 무시 (남이 덮어쓰지 못하게)
function doGet(e) {
  var props = PropertiesService.getScriptProperties();
  if (!e.parameter.code || props.getProperty('KAKAO_REFRESH_TOKEN')) return HtmlService.createHtmlOutput('ok');
  var t = kakaoToken({ grant_type: 'authorization_code', redirect_uri: WEBAPP_URL, code: e.parameter.code });
  if (!t.refresh_token) return HtmlService.createHtmlOutput('연결 실패: ' + (t.error_description || t.error));
  if (String(t.scope).indexOf('talk_message') < 0) return HtmlService.createHtmlOutput('연결 실패: 동의 화면에서 "카카오톡 메시지 전송"에 체크해 주세요.');
  props.setProperty('KAKAO_REFRESH_TOKEN', t.refresh_token);
  sendKakao('카카오톡 오류 알림이 연결됐어요');
  return HtmlService.createHtmlOutput('카카오톡 알림이 연결됐어요. 이 창을 닫아도 됩니다.');
}

// 편집기에서 실행해 권한 승인 + 알림 테스트
function testNotify() {
  notify('테스트 알림입니다');
}
