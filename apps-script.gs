/**
 * 컨빌디자인 - 교회 로고 질문지 수집 Apps Script
 *
 * 동작:
 *  - 폼이 보낸 JSON을 받아 시트에 한 행 추가
 *  - 첨부파일(base64)이 있으면 드라이브 폴더에 저장
 *  - 시트의 "참고파일" 셀에는 폴더 링크를 하이퍼링크로 기록
 *
 * 사용법:
 *  1) 구글 스프레드시트 → 확장 프로그램 → Apps Script
 *  2) 이 파일 내용을 Code.gs 에 붙여넣고 저장
 *  3) (최초 1회) testAppend 실행 → 권한 승인 (시트 + 드라이브 + 메일)
 *  4) 배포 → 새 배포 → 유형: 웹 앱
 *     - 다음 사용자로 실행: 나
 *     - 액세스 권한: 모든 사용자 (익명 포함)
 *  5) 발급된 웹앱 URL을 convil-logo-process.html 의
 *     APPS_SCRIPT_URL 변수에 붙여넣기
 *
 *  ※ Apps Script 코드를 수정한 뒤에는 반드시
 *    "배포 → 배포 관리 → 편집(연필) → 새 버전 → 배포" 로 새 버전을 배포해야
 *    웹앱이 새 코드를 사용합니다.
 */

// ─── 설정 ─────────────────────────────────────────────
const SHEET_NAME = '로고 질문지';

// 첨부파일을 저장할 드라이브 폴더
//   ROOT_FOLDER_ID 가 있으면 그 폴더 안에 교회별 하위폴더 생성
//   비어있으면 내 드라이브 루트에 ROOT_FOLDER_NAME 폴더를 자동 생성
const ROOT_FOLDER_ID   = '';
const ROOT_FOLDER_NAME = '컨빌 로고 질문지 첨부파일';

// 첨부파일 공유 설정
//   true  → 링크가 있는 누구나 보기 가능 (시트 보는 사람이 링크로 바로 열람)
//   false → 비공개 (스크립트 소유자만 열람)
const SHARE_ANYONE_WITH_LINK = true;

// 알림 받을 이메일 (사용 안 하려면 '' 로 두세요)
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

    // 1) 첨부파일 → 드라이브 저장
    const folderInfo = saveFiles_(data);

    // 2) 시트에 행 추가
    const sheet = getOrCreateSheet_();
    const filesCell = folderInfo
      ? '=HYPERLINK("' + folderInfo.url + '","' + folderInfo.count + '개 파일 보기")'
      : (data.fileNames || '없음');

    sheet.appendRow([
      data.timestamp    || nowKST_(),
      data.churchName   || '',
      data.contactName  || '',
      data.contactPhone || '',
      data.contactEmail || '',
      data.logoType     || '',
      data.churchValue  || '',
      filesCell,
      data.likeMood     || '',
      data.avoidMood    || ''
    ]);

    // 3) 알림 메일
    sendNotification_(data, folderInfo);

    return jsonResponse_({
      result: 'success',
      folderUrl: folderInfo ? folderInfo.url : null
    });
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

// ─── 드라이브 저장 ────────────────────────────────────
function getRootFolder_() {
  if (ROOT_FOLDER_ID) return DriveApp.getFolderById(ROOT_FOLDER_ID);

  const it = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
}

function saveFiles_(data) {
  if (!Array.isArray(data.files) || data.files.length === 0) return null;

  const safe = sanitizeName_(data.churchName || 'unknown');
  const stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd-HHmmss');
  const folder = getRootFolder_().createFolder(safe + '_' + stamp);

  let saved = 0;
  data.files.forEach(f => {
    if (!f || !f.data) return;
    try {
      const bytes = Utilities.base64Decode(f.data);
      const blob = Utilities.newBlob(
        bytes,
        f.mimeType || 'application/octet-stream',
        f.name || 'untitled'
      );
      const file = folder.createFile(blob);
      if (SHARE_ANYONE_WITH_LINK) {
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (_) { /* 일부 워크스페이스 정책에서 실패 가능 */ }
      }
      saved++;
    } catch (err) {
      console.error('File save failed:', f && f.name, err);
    }
  });

  if (SHARE_ANYONE_WITH_LINK) {
    try {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (_) {}
  }

  return saved > 0 ? { url: folder.getUrl(), count: saved } : null;
}

function sanitizeName_(name) {
  return String(name).replace(/[\\/:*?"<>|]/g, '_').trim().substring(0, 60) || 'unknown';
}

// ─── 알림 메일 ────────────────────────────────────────
function sendNotification_(data, folderInfo) {
  if (!NOTIFY_EMAIL) return;

  const subject = '[컨빌디자인] 로고 질문지 제출 · ' + (data.churchName || '');
  const body = [
    '새로운 로고 질문지가 제출되었습니다.',
    '',
    '제출시간: '   + (data.timestamp    || ''),
    '교회명: '     + (data.churchName   || ''),
    '담당자: '     + (data.contactName  || ''),
    '연락처: '     + (data.contactPhone || ''),
    '이메일: '     + (data.contactEmail || ''),
    '로고형태: '   + (data.logoType     || ''),
    '교회가치: '   + (data.churchValue  || ''),
    '선호분위기: ' + (data.likeMood     || ''),
    '피할분위기: ' + (data.avoidMood    || ''),
    '',
    folderInfo
      ? '첨부파일 (' + folderInfo.count + '개): ' + folderInfo.url
      : '첨부파일: 없음'
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
// Apps Script 편집기에서 testAppend 를 실행해 권한 승인 + 동작 확인.
// (1x1 투명 PNG 한 장이 드라이브에 저장됩니다)
function testAppend() {
  const TINY_PNG_BASE64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII=';

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
        fileNames:    'sample.png',
        likeMood:     '부드럽고 따뜻한, 깨끗하고 순수한',
        avoidMood:    '어둡고 무거운',
        files: [
          { name: 'sample.png', mimeType: 'image/png', size: 70, data: TINY_PNG_BASE64 }
        ]
      })
    }
  });
}
