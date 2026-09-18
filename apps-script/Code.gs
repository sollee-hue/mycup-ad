/**
 * LG myCup 사이니지 광고 문의 — 리드 수신 스크립트
 *
 * 설치 방법
 *  1. 리드를 쌓을 구글 스프레드시트를 엽니다.
 *  2. 확장 프로그램 → Apps Script 를 엽니다.
 *  3. 기본 코드를 지우고 이 파일 내용을 전부 붙여넣습니다.
 *  4. 아래 SECRET 을 원하는 값으로 바꾸고, index.html 의 SHEET.secret 에 같은 값을 넣습니다.
 *  5. 배포 → 새 배포 → 유형 "웹 앱"
 *       - 실행 계정: 나
 *       - 액세스 권한: 모든 사용자   ← 이게 아니면 페이지에서 전송이 안 됩니다
 *  6. 배포 후 나오는 /exec 로 끝나는 URL 을 index.html 의 SHEET.endpoint 에 넣습니다.
 *
 * 코드를 고친 뒤에는 반드시 "배포 관리 → 수정 → 새 버전"으로 다시 배포해야 반영됩니다.
 */

var SECRET     = 'mycup-2026';   // index.html 의 SHEET.secret 과 반드시 동일
var SHEET_NAME = '리드';
var NOTIFY_TO  = '';             // 새 리드 알림 메일을 받을 주소 (비우면 보내지 않음)

var HEADERS = ['접수 시각', '회사/브랜드', '담당자명', '휴대폰', '이메일',
               '희망 패키지', '문의 내용', '개인정보 동의', '유입 페이지'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json({ ok: false, error: 'busy' });
  }

  try {
    var raw = (e && e.postData && e.postData.contents) || '{}';
    var d;
    try {
      d = JSON.parse(raw);
    } catch (err) {
      d = (e && e.parameter) || {};   // 폼 인코딩으로 들어온 경우
    }

    if (String(d.secret || '') !== SECRET) {
      return json({ ok: false, error: 'unauthorized' });
    }

    // 네트워크 사정으로 같은 요청이 두 번 도착해도 한 줄만 쌓습니다.
    if (d.reqId) {
      var cache = CacheService.getScriptCache();
      if (cache.get('req_' + d.reqId)) {
        return json({ ok: true, duplicate: true });
      }
      cache.put('req_' + d.reqId, '1', 600);
    }

    var sh = getSheet_();
    var row = [
      new Date(),
      d.company || '',
      d.name || '',
      d.phone || '',
      d.email || '',
      d['package'] || '',
      d.message || '',
      d.agree || '',
      d.page || ''
    ];
    sh.appendRow(row);

    if (NOTIFY_TO) {
      MailApp.sendEmail({
        to: NOTIFY_TO,
        subject: '[myCup 사이니지] 새 광고 문의 — ' + (d.company || '(회사명 없음)'),
        body: [
          '회사/브랜드: ' + (d.company || ''),
          '담당자명: '   + (d.name || ''),
          '휴대폰: '     + (d.phone || ''),
          '이메일: '     + (d.email || ''),
          '희망 패키지: ' + (d['package'] || ''),
          '',
          '문의 내용:',
          (d.message || '(없음)'),
          '',
          '시트에서 보기: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
        ].join('\n')
      });
    }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 배포가 살아 있는지 브라우저로 확인할 때 씁니다. */
function doGet() {
  return json({ ok: true, service: 'mycup-signage-lead' });
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#E6F5F4');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(7, 320);
  }
  return sh;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Apps Script 편집기에서 직접 실행해 동작을 확인하는 테스트 함수입니다. */
function testAppend() {
  doPost({ postData: { contents: JSON.stringify({
    secret: SECRET, company: '테스트컴퍼니', name: '홍길동',
    phone: '010-0000-0000', email: 'test@example.com',
    'package': '오피스', message: '테스트 문의입니다', agree: '동의', page: 'local',
    reqId: 'test-' + Date.now()
  }) } });
}
