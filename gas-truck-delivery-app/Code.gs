/**
 * 家具配送トラック業務管理アプリ
 * Google Apps Script サーバーサイドコード
 */

// スプレッドシートのID（デプロイ時に実際のIDに置き換え）
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';
const SHEET_NAME = '配送履歴';

/**
 * Webアプリのエントリーポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('配送業務管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * スプレッドシートを取得または作成
 */
function getOrCreateSheet() {
  let ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    // スプレッドシートが存在しない場合は新規作成
    ss = SpreadsheetApp.create('配送業務管理データ');
    Logger.log('新しいスプレッドシートを作成しました: ' + ss.getId());
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // ヘッダー行を設定
    sheet.getRange(1, 1, 1, 9).setValues([[
      '記録ID',
      '号車',
      'ドライバー名',
      '記録種別',
      '日付',
      '時刻',
      '走行距離(km)',
      '備考',
      '登録日時'
    ]]);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 業務記録を保存
 * @param {Object} data - 送信データ
 * @returns {Object} 結果
 */
function saveRecord(data) {
  try {
    const sheet = getOrCreateSheet();
    const recordId = Utilities.getUuid().substring(0, 8);
    const now = new Date();
    const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    sheet.appendRow([
      recordId,
      data.truckNumber,
      data.driverName,
      data.recordType,
      data.date,
      data.time,
      data.mileage || '',
      data.notes || '',
      timestamp
    ]);

    return {
      success: true,
      message: '記録を保存しました',
      recordId: recordId
    };
  } catch (e) {
    return {
      success: false,
      message: 'エラー: ' + e.message
    };
  }
}

/**
 * 履歴を取得
 * @param {Object} filter - フィルター条件（オプション）
 * @returns {Object} 履歴データ
 */
function getHistory(filter) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        success: true,
        records: [],
        message: '履歴がありません'
      };
    }

    const headers = data[0];
    let records = [];

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const record = {
        recordId: row[0],
        truckNumber: row[1],
        driverName: row[2],
        recordType: row[3],
        date: row[4],
        time: row[5],
        mileage: row[6],
        notes: row[7],
        timestamp: row[8]
      };

      // フィルター適用
      if (filter) {
        if (filter.truckNumber && record.truckNumber !== filter.truckNumber) continue;
        if (filter.date && record.date !== filter.date) continue;
      }

      records.push(record);
    }

    // 最新50件に制限
    records = records.slice(0, 50);

    return {
      success: true,
      records: records
    };
  } catch (e) {
    return {
      success: false,
      message: 'エラー: ' + e.message,
      records: []
    };
  }
}

/**
 * 履歴をテキスト形式で取得（コピー用）
 * @param {Array} records - 履歴レコード
 * @returns {string} テキスト形式の履歴
 */
function formatHistoryForCopy(records) {
  if (!records || records.length === 0) {
    return 'データがありません';
  }

  let text = '=== 配送業務履歴 ===\n\n';

  records.forEach((record, index) => {
    text += `【${index + 1}】${record.date} ${record.time}\n`;
    text += `号車: ${record.truckNumber}\n`;
    text += `ドライバー: ${record.driverName}\n`;
    text += `種別: ${record.recordType}\n`;
    if (record.mileage) {
      text += `走行距離: ${record.mileage}km\n`;
    }
    if (record.notes) {
      text += `備考: ${record.notes}\n`;
    }
    text += '\n';
  });

  return text;
}

/**
 * 号車リストを取得
 * @returns {Array} 号車リスト
 */
function getTruckList() {
  // 必要に応じてスプレッドシートから取得するように変更可能
  return [
    { id: '1', name: '1号車' },
    { id: '2', name: '2号車' },
    { id: '3', name: '3号車' },
    { id: '4', name: '4号車' },
    { id: '5', name: '5号車' },
    { id: '6', name: '6号車' },
    { id: '7', name: '7号車' },
    { id: '8', name: '8号車' },
    { id: '9', name: '9号車' },
    { id: '10', name: '10号車' }
  ];
}
