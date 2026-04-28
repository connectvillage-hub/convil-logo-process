/**
 * 컨빌디자인 - 교회 로고 질문지 수집 Apps Script
 *
 * 사용법:
 *  1) 구글 스프레드시트 → 확장 프로그램 → Apps Script
 *  2) 이 파일 내용을 Code.gs 에 붙여넣고 저장
 *  3) 배포 → 새 배포 → 유형: 웹 앱
 *     - 다음 사용자로 실행: 나
 *     - 액세스 권한: 모든 사용자 (익명 포함)
 *  4) 발급된 웹앱 URL을 convil-logo-process.html 의
 *     APPS_SCRIPT_URL 변수에 붙여넣기
 *
 * 주의:
 *  - 폼은 fetch(..., { mode: 'no-cors' }) 로 호출하므로
 *    응답 본문은 클라이언트가 읽을 수 없습니다. (정상)
 *  - postData.contents 에는 JSON 문자열이 들어옵니다.
 */

// ─── 설정 ─────────────────────────────────────────────
const SHEET_NAME = '로고 질문지';

// 알림을 받을 이메일 (사용 안 하려면 '' 로 두세요)
const NOTIFY_EMAIL = '';

const HEADERS = [
  '제출시간', '교회명', '담당자', '연락처', '이메일',
  '로고형태', '교회가치', '참고파일', '선호분위기', '피할분위기'
];

// ─── 엔드포인트 ──────────────────────────────────────
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ result: 'error', message: 'No payload' });
    }

    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet_();

    sheet.appendRow([
      data.timestamp    || nowKST_(),
      data.churchName   || '',
      data.contactName  || '',
      data.contactPhone || '',
      data.contactEmail || '',
      data.logoType     || '',
      data.churchValue  || '',
      data.fileNames    || '',
      data.likeMood     || '',
      data.avoidMood    || ''
    ]);

    sendNotification_(data);

    return jsonResponse_({ result: 'success' });
  } catch (err) {
    console.error(err);
    return jsonResponse_({ result: 'error', message: String(err) });
  }
}

function doGet() {
  return jsonResponse_({ status: 'ok', service: 'convil-logo-form' });
}

// ─── 시트 준비 ────────────────────────────────────────
function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#F0E4CC')
      .setFontColor('#7A5A20');
    sheet.setFrozenRows(1);

    const widths = [160, 140, 100, 130, 200, 180, 320, 220, 260, 260];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

// ─── 알림 메일 ────────────────────────────────────────
function sendNotification_(data) {
  if (!NOTIFY_EMAIL) return;

  const subject = '[컨빌디자인] 로고 질문지 제출 · ' + (data.churchName || '');
  const body = [
    '새로운 로고 질문지가 제출되었습니다.',
    '',
    '제출시간: '  + (data.timestamp    || ''),
    '교회명: '    + (data.churchName   || ''),
    '담당자: '    + (data.contactName  || ''),
    '연락처: '    + (data.contactPhone || ''),
    '이메일: '    + (data.contactEmail || ''),
    '로고형태: '  + (data.logoType     || ''),
    '교회가치: '  + (data.churchValue  || ''),
    '참고파일: '  + (data.fileNames    || ''),
    '선호분위기: ' + (data.likeMood    || ''),
    '피할분위기: ' + (data.avoidMood   || '')
  ].join('\n');

  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

// ─── 유틸 ─────────────────────────────────────────────
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function nowKST_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

// ─── 테스트용 ────────────────────────────────────────
// Apps Script 편집기에서 testAppend 를 실행해 시트에 샘플 행이 들어가는지 확인할 수 있습니다.
function testAppend() {
  doPost({
    postData: {
      contents: JSON.stringify({
        timestamp:    nowKST_(),
        churchName:   '테스트교회',
        contactName:  '홍길동',
        contactPhone: '010-1234-5678',
        contactEmail: 'test@example.com',
        logoType:     '타이포형, 심볼형',
        churchValue:  '회복, 공동체, 환대',
        fileNames:    '없음',
        likeMood:     '부드럽고 따뜻한, 깨끗하고 순수한',
        avoidMood:    '어둡고 무거운'
      })
    }
  });
}
