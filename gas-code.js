/**
 * eSTACK 受付システム - 訪問履歴記録スクリプト
 * Google Apps Script に貼り付けて「Webアプリとしてデプロイ」してください。
 * アクセス権限：全員（匿名を含む）
 */

const SHEET_NAME = '訪問履歴';
const HEADERS = ['日時', '種別', '担当者', '会社名', '訪問者名'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // シートを取得（なければ作成）
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);

      // ヘッダーのスタイル設定
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setBackground('#2563EB');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160); // 日時
      sheet.setColumnWidth(2, 80);  // 種別
      sheet.setColumnWidth(3, 140); // 担当者
      sheet.setColumnWidth(4, 180); // 会社名
      sheet.setColumnWidth(5, 140); // 訪問者名
    }

    // データを追記
    sheet.appendRow([
      data.datetime || '',
      data.type     || '',
      data.staff    || '',
      data.company  || '',
      data.visitor  || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** テスト用（GASエディタから直接実行して動作確認できます） */
function testDoPost() {
  const mockEvent = {
    postData: {
      contents: JSON.stringify({
        datetime: '2026/02/17 17:35',
        type: '来客',
        staff: '竹田よりちか',
        company: '株式会社テスト',
        visitor: '山田 太郎'
      })
    }
  };
  const result = doPost(mockEvent);
  Logger.log(result.getContent());
}
