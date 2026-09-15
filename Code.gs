/**
 * 62기 길라잡이 - 주차별 과제 파일 업로드 백엔드
 *
 * 사용 방법
 * 1) 이 파일을 저장할 구글 시트를 하나 만들거나 정합니다.
 * 2) 그 시트의 주소창 URL에서 시트 ID를 복사합니다.
 *    예) https://docs.google.com/spreadsheets/d/이 부분이_시트_ID/edit
 * 3) 아래 SPREADSHEET_ID 값에 복사한 ID를 붙여넣습니다. (★ 필수 ★
 *    독립형 Apps Script 프로젝트라 이 값이 없으면 "알 수 없는 오류"가 납니다.)
 * 4) 배포 > 새 배포 > 유형: 웹앱
 *      - 실행 대상: 나(본인 계정)
 *      - 액세스 권한: 전체 허용(익명 포함)
 * 5) 배포 후 나오는 "웹 앱 URL"을 복사해서
 *    62th_guidebook.html 안의 HW_GAS_URL 값에 붙여넣으세요.
 *
 * 동작 방식
 * - 파일은 구글 드라이브의 "62기_주차별과제 > N주" 폴더에 저장됩니다.
 * - 파일마다 "링크가 있는 모든 사용자에게 보기 허용"으로 공유 설정됩니다.
 * - 업로드 기록(업로드일시, 주차, 파일명, 파일링크)은
 *   위에서 지정한 스프레드시트의 "주차별과제" 시트에 한 줄씩 쌓입니다.
 */

var SPREADSHEET_ID = '1lDm_eP6Sf9Tq34akdC0voBRHh3_qmL5KTUjQLj-VMWU';
var SHEET_NAME = '주차별과제';
var DRIVE_ROOT_FOLDER_NAME = '62기_주차별과제';

function doGet(e) {
  try {
    var action = e && e.parameter ? e.parameter.action : null;
    if (action === 'listHomework') {
      return jsonOutput({ items: getHomeworkList() });
    }
    return jsonOutput({ error: 'invalid action' });
  } catch (err) {
    return jsonOutput({ error: err.message });
  }
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput({ success: false, message: '요청 형식이 올바르지 않습니다.' });
  }

  if (data.action === 'uploadHomework') {
    return jsonOutput(uploadHomeworkFile(data));
  }
  return jsonOutput({ success: false, message: 'invalid action' });
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['업로드일시', '주차', '파일명', '파일링크']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getHomeworkList() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return values
    .filter(function (row) { return row[2]; }) // 파일명이 있는 행만
    .map(function (row) {
      return {
        uploadedAt: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
        week: row[1],
        fileName: row[2],
        fileUrl: row[3]
      };
    });
}

function getOrCreateFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function uploadHomeworkFile(data) {
  try {
    if (!data.fileName || !data.base64 || !data.week) {
      return { success: false, message: '필수 값이 누락되었습니다.' };
    }

    var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), DRIVE_ROOT_FOLDER_NAME);
    var weekFolder = getOrCreateFolder(rootFolder, data.week + '주');

    var decoded = Utilities.base64Decode(data.base64);
    var blob = Utilities.newBlob(decoded, data.mimeType || 'application/octet-stream', data.fileName);
    var file = weekFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var url = file.getUrl();
    var now = new Date();

    var sheet = getSheet();
    sheet.appendRow([now, data.week, data.fileName, url]);

    return {
      success: true,
      week: data.week,
      fileName: data.fileName,
      fileUrl: url,
      uploadedAt: now.toISOString()
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
