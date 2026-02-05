/**
 * 家具配送トラック業務管理アプリ
 * Google Apps Script サーバーサイドコード
 */

// 管理者パスワード（デプロイ時に変更してください）
const ADMIN_PASSWORD = 'admin1234';

// シート名
const SHEET_NAMES = {
  HISTORY: '配送履歴',
  TRUCK_MASTER: '号車マスタ',
  TRUCK_SETTINGS: '号車設定'
};

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
 * スプレッドシートを取得（スクリプトプロパティでIDを永続化）
 */
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('SPREADSHEET_ID');

  // 既存のスプレッドシートIDがある場合
  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      // IDは保存されているがスプレッドシートが削除された場合
      Logger.log('保存されたスプレッドシートが見つかりません。新規作成します。');
    }
  }

  // 新しいスプレッドシートを作成
  const ss = SpreadsheetApp.create('配送業務管理データ');
  const newId = ss.getId();

  // スクリプトプロパティに保存
  props.setProperty('SPREADSHEET_ID', newId);
  Logger.log('新しいスプレッドシートを作成しました: ' + newId);
  Logger.log('スプレッドシートURL: ' + ss.getUrl());

  return ss;
}

/**
 * 現在使用中のスプレッドシートURLを取得（管理用）
 */
function getSpreadsheetUrl() {
  const ss = getSpreadsheet();
  return ss.getUrl();
}

/**
 * 履歴シートを取得または作成
 */
function getOrCreateHistorySheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.HISTORY);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.HISTORY);
    sheet.getRange(1, 1, 1, 22).setValues([[
      '記録ID',
      '日付',
      '号車',
      'ドライバー名',
      '車両No',
      '庸車区分①',
      '庸車区分②',
      '助手',
      '出庫メーター',
      '出庫時刻',
      '行き先',
      '帰庫メーター',
      '配送センター到着時刻',
      '営業所帰庫時刻',
      '代引き返金有無',
      '有料道路使用',
      '有料道路区間情報',
      '記録種別',
      '備考',
      '登録日時',
      '送信タイミング',
      '移動時間（分）'
    ]]);
    sheet.getRange(1, 1, 1, 22).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 号車マスタシートを取得または作成
 */
function getOrCreateTruckMasterSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.TRUCK_MASTER);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.TRUCK_MASTER);
    sheet.getRange(1, 1, 1, 4).setValues([[
      '号車',
      'ドライバー名',
      '車両No',
      '登録日時'
    ]]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 号車リストを取得（1〜30号車）
 */
function getTruckList() {
  const trucks = [];
  for (let i = 1; i <= 30; i++) {
    trucks.push({ id: String(i), name: i + '号車' });
  }
  return trucks;
}

/**
 * 号車に紐づくドライバーリストを取得
 */
function getDriversByTruck(truckNumber) {
  try {
    const sheet = getOrCreateTruckMasterSheet();
    const data = sheet.getDataRange().getValues();
    const drivers = [];
    const searchTruck = String(truckNumber).trim();

    for (let i = 1; i < data.length; i++) {
      const rowTruck = String(data[i][0]).trim();
      const driverName = data[i][1] ? String(data[i][1]).trim() : '';

      if (rowTruck === searchTruck && driverName) {
        // 重複チェック
        if (!drivers.find(d => d.name === driverName)) {
          drivers.push({
            id: driverName,
            name: driverName
          });
        }
      }
    }
    return drivers;
  } catch (e) {
    Logger.log('getDriversByTruck error: ' + e.message);
    return [];
  }
}

/**
 * 号車に紐づく車両Noリストを取得
 */
function getVehiclesByTruck(truckNumber) {
  try {
    const sheet = getOrCreateTruckMasterSheet();
    const data = sheet.getDataRange().getValues();
    const vehicles = [];
    const searchTruck = String(truckNumber).trim();

    for (let i = 1; i < data.length; i++) {
      const rowTruck = String(data[i][0]).trim();
      const vehicleNo = data[i][2] ? String(data[i][2]).trim() : '';

      if (rowTruck === searchTruck && vehicleNo) {
        // 重複チェック
        if (!vehicles.find(v => v.name === vehicleNo)) {
          vehicles.push({
            id: vehicleNo,
            name: vehicleNo
          });
        }
      }
    }
    return vehicles;
  } catch (e) {
    Logger.log('getVehiclesByTruck error: ' + e.message);
    return [];
  }
}

/**
 * ドライバーを号車に登録
 */
function registerDriver(truckNumber, name) {
  try {
    if (!truckNumber) {
      return { success: false, message: '号車を選択してください' };
    }
    if (!name || name.trim() === '') {
      return { success: false, message: 'ドライバー名を入力してください' };
    }

    const sheet = getOrCreateTruckMasterSheet();
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const truckStr = String(truckNumber).trim();
    const nameStr = String(name).trim();

    Logger.log('registerDriver: 号車=' + truckStr + ', ドライバー=' + nameStr);

    sheet.appendRow([truckStr, nameStr, '', timestamp]);

    // シートを即座に保存
    SpreadsheetApp.flush();

    Logger.log('registerDriver: 登録完了');

    return {
      success: true,
      message: 'ドライバーを登録しました',
      driver: { id: nameStr, name: nameStr }
    };
  } catch (e) {
    Logger.log('registerDriver error: ' + e.message);
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 車両Noを号車に登録
 */
function registerVehicle(truckNumber, vehicleNo) {
  try {
    if (!truckNumber) {
      return { success: false, message: '号車を選択してください' };
    }
    if (!vehicleNo || vehicleNo.trim() === '') {
      return { success: false, message: '車両Noを入力してください' };
    }

    // 半角数字4桁チェック
    const numOnly = String(vehicleNo).trim();
    if (!/^\d{1,4}$/.test(numOnly)) {
      return { success: false, message: '車両Noは半角数字4桁以内で入力してください' };
    }

    const sheet = getOrCreateTruckMasterSheet();
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const truckStr = String(truckNumber).trim();

    Logger.log('registerVehicle: 号車=' + truckStr + ', 車両No=' + numOnly);

    sheet.appendRow([truckStr, '', numOnly, timestamp]);

    // シートを即座に保存
    SpreadsheetApp.flush();

    Logger.log('registerVehicle: 登録完了');

    return {
      success: true,
      message: '車両Noを登録しました',
      vehicle: { id: numOnly, name: numOnly }
    };
  } catch (e) {
    Logger.log('registerVehicle error: ' + e.message);
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 号車設定シートを取得または作成
 */
function getOrCreateTruckSettingsSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.TRUCK_SETTINGS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.TRUCK_SETTINGS);
    sheet.getRange(1, 1, 1, 4).setValues([[
      '号車',
      '送信タイミング',
      '移動時間（分）',
      '更新日時'
    ]]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // 1〜30号車のデフォルト設定を追加
    const defaultData = [];
    for (let i = 1; i <= 30; i++) {
      defaultData.push([i + '号車', '営業所到着', 30, '']);
    }
    sheet.getRange(2, 1, 30, 4).setValues(defaultData);
  }

  return sheet;
}

/**
 * 号車設定を取得
 */
function getTruckSettings(truckNumber) {
  try {
    const sheet = getOrCreateTruckSettingsSheet();
    const data = sheet.getDataRange().getValues();
    const searchTruck = String(truckNumber).trim();

    for (let i = 1; i < data.length; i++) {
      const rowTruck = String(data[i][0]).trim();
      if (rowTruck === searchTruck) {
        return {
          truckNumber: rowTruck,
          submitTiming: data[i][1] || '営業所到着',
          travelTime: parseInt(data[i][2]) || 30
        };
      }
    }

    // デフォルト値を返す
    return {
      truckNumber: searchTruck,
      submitTiming: '営業所到着',
      travelTime: 30
    };
  } catch (e) {
    Logger.log('getTruckSettings error: ' + e.message);
    return {
      truckNumber: String(truckNumber),
      submitTiming: '営業所到着',
      travelTime: 30
    };
  }
}

/**
 * 号車設定を更新
 */
function updateTruckSettings(truckNumber, submitTiming, travelTime) {
  try {
    const sheet = getOrCreateTruckSettingsSheet();
    const data = sheet.getDataRange().getValues();
    const searchTruck = String(truckNumber).trim();
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    for (let i = 1; i < data.length; i++) {
      const rowTruck = String(data[i][0]).trim();
      if (rowTruck === searchTruck) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[submitTiming, parseInt(travelTime), timestamp]]);
        SpreadsheetApp.flush();
        return { success: true, message: '設定を保存しました' };
      }
    }

    // 見つからなければ新規追加
    sheet.appendRow([searchTruck, submitTiming, parseInt(travelTime), timestamp]);
    SpreadsheetApp.flush();
    return { success: true, message: '設定を保存しました' };
  } catch (e) {
    Logger.log('updateTruckSettings error: ' + e.message);
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 全号車の設定を取得
 */
function getAllTruckSettings() {
  try {
    const sheet = getOrCreateTruckSettingsSheet();
    const data = sheet.getDataRange().getValues();
    const settings = [];

    for (let i = 1; i < data.length; i++) {
      settings.push({
        truckNumber: String(data[i][0]).trim(),
        submitTiming: data[i][1] || '営業所到着',
        travelTime: parseInt(data[i][2]) || 30
      });
    }

    return settings;
  } catch (e) {
    Logger.log('getAllTruckSettings error: ' + e.message);
    return [];
  }
}

/**
 * 庸車区分①リストを取得
 */
function getYoshaCategory1List() {
  return [
    { id: '1man', name: '1マン' },
    { id: '2man', name: '2マン' },
    { id: '3man', name: '3マン' }
  ];
}

/**
 * 庸車区分②リストを取得
 */
function getYoshaCategory2List() {
  return [
    { id: 'monthly', name: '月間契約' },
    { id: 'temporary', name: '臨時' },
    { id: 'other', name: 'その他' }
  ];
}

/**
 * 助手リストを取得
 */
function getHelperList() {
  return [
    { id: 'vendor', name: '業者' },
    { id: 'employee', name: '社員' },
    { id: 'parttime', name: 'アルバイト' }
  ];
}

/**
 * 行き先リストを取得
 */
function getDestinationList() {
  return [
    { id: 'office', name: '営業所' },
    { id: 'center', name: '配送センター' },
    { id: 'direct', name: '直帰' }
  ];
}

/**
 * 全マスタデータを取得
 */
function getAllMasterData() {
  return {
    trucks: getTruckList(),
    yoshaCategory1: getYoshaCategory1List(),
    yoshaCategory2: getYoshaCategory2List(),
    helpers: getHelperList(),
    destinations: getDestinationList()
  };
}

/**
 * 号車変更時のデータ取得
 */
function getTruckData(truckNumber) {
  return {
    drivers: getDriversByTruck(truckNumber),
    vehicles: getVehiclesByTruck(truckNumber)
  };
}

/**
 * 日付を文字列に変換（YYYY-MM-DD形式）
 */
function formatDateValue(value) {
  if (!value) return '';

  // 既にDate型の場合
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  // 文字列の場合、そのまま返す（既にYYYY-MM-DD形式と仮定）
  const strValue = String(value).trim();

  // 日本語形式の日付（YYYY/MM/DD）をYYYY-MM-DDに変換
  if (/^\d{4}\/\d{1,2}\/\d{1,2}/.test(strValue)) {
    const parts = strValue.split('/');
    return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0').substring(0, 2)}`;
  }

  return strValue;
}

/**
 * 時刻を文字列に変換（HH:MM形式）
 */
function formatTimeValue(value) {
  if (!value) return '';

  // オブジェクト（Date含む）の場合、まずUtilities.formatDateを試す
  if (typeof value === 'object') {
    try {
      return Utilities.formatDate(value, 'Asia/Tokyo', 'HH:mm');
    } catch (e) {
      Logger.log('formatTimeValue object error: ' + e.message);
    }
  }

  const strValue = String(value).trim();

  // HH:MM:SS形式の場合
  if (/^\d{1,2}:\d{2}:\d{2}$/.test(strValue)) {
    const parts = strValue.split(':');
    return parts[0].padStart(2, '0') + ':' + parts[1];
  }

  // HH:MM形式の場合
  if (/^\d{1,2}:\d{2}$/.test(strValue)) {
    const parts = strValue.split(':');
    return parts[0].padStart(2, '0') + ':' + parts[1];
  }

  // 数値の場合（Excelシリアル時刻）
  const num = parseFloat(strValue);
  if (!isNaN(num) && num >= 0 && num < 1) {
    const totalMinutes = Math.round(num * 24 * 60);
    const hours = Math.floor(totalMinutes / 60) % 24;
    const mins = totalMinutes % 60;
    return String(hours).padStart(2, '0') + ':' + String(mins).padStart(2, '0');
  }

  // "GMT"や"1899"を含む日付文字列の場合
  if (strValue.includes('GMT') || strValue.includes('1899') || strValue.includes('Mon ') || strValue.includes('Tue ') || strValue.includes('Wed ') || strValue.includes('Thu ') || strValue.includes('Fri ') || strValue.includes('Sat ') || strValue.includes('Sun ')) {
    // 時刻部分を正規表現で抽出（HH:MM:SS）
    const timeMatch = strValue.match(/(\d{1,2}):(\d{2}):(\d{2})/);
    if (timeMatch) {
      return timeMatch[1].padStart(2, '0') + ':' + timeMatch[2];
    }
  }

  return strValue;
}

/**
 * 業務開始記録を保存
 */
function saveStartRecord(data) {
  try {
    const sheet = getOrCreateHistorySheet();
    const recordId = Utilities.getUuid().substring(0, 8);
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    // 日付は文字列として保存（自動変換を防ぐ）
    const dateStr = String(data.date || '').trim();

    sheet.appendRow([
      recordId,
      dateStr,
      String(data.truckNumber || '').trim(),
      String(data.driverName || '').trim(),
      String(data.vehicleNo || '').trim(),
      String(data.yoshaCategory1 || ''),
      String(data.yoshaCategory2 || ''),
      String(data.helper || ''),
      String(data.departureMeter || ''),
      String(data.departureTime || ''),
      '', // 行き先（終了時）
      '', // 帰庫メーター
      '', // 配送センター到着時刻
      '', // 営業所帰庫時刻
      '', // 代引き返金有無
      '', // 有料道路使用
      '', // 有料道路区間情報
      '業務開始',
      String(data.notes || ''),
      timestamp
    ]);

    // シートを即座に保存
    SpreadsheetApp.flush();

    return {
      success: true,
      message: '業務開始を記録しました',
      recordId: recordId
    };
  } catch (e) {
    Logger.log('saveStartRecord error: ' + e.message);
    return {
      success: false,
      message: 'エラー: ' + e.message
    };
  }
}

/**
 * 業務終了記録を保存
 */
function saveEndRecord(data) {
  try {
    const sheet = getOrCreateHistorySheet();
    const recordId = Utilities.getUuid().substring(0, 8);
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    // 有料道路区間情報をJSON形式で保存
    let tollRoadInfo = '';
    if (data.useTollRoad === 'はい' && data.tollRoadSections && data.tollRoadSections.length > 0) {
      tollRoadInfo = JSON.stringify(data.tollRoadSections);
    }

    // 直帰の場合、配送センター到着時刻は営業所帰庫時刻と同じ
    const centerArrivalTime = data.destination === '直帰' ? data.officeArrivalTime : data.centerArrivalTime;

    // 日付は文字列として保存（自動変換を防ぐ）
    const dateStr = String(data.date || '').trim();

    // 号車設定を取得（送信時の設定をレコードに紐づける）
    const truckSettings = getTruckSettings(data.truckNumber);
    const submitTiming = data.submitTiming || truckSettings.submitTiming;
    const travelTime = data.travelTime !== undefined ? data.travelTime : truckSettings.travelTime;

    sheet.appendRow([
      recordId,
      dateStr,
      String(data.truckNumber || '').trim(),
      String(data.driverName || '').trim(),
      '', // 車両No（開始時のみ）
      '', // 庸車区分①
      '', // 庸車区分②
      '', // 助手
      '', // 出庫メーター
      '', // 出庫時刻
      String(data.destination || ''),
      String(data.returnMeter || ''),
      String(centerArrivalTime || ''),
      String(data.officeArrivalTime || ''),
      String(data.hasCashOnDelivery || ''),
      String(data.useTollRoad || ''),
      tollRoadInfo,
      '業務終了',
      String(data.notes || ''),
      timestamp,
      submitTiming,
      travelTime
    ]);

    // シートを即座に保存
    SpreadsheetApp.flush();

    return {
      success: true,
      message: '業務終了を記録しました',
      recordId: recordId
    };
  } catch (e) {
    Logger.log('saveEndRecord error: ' + e.message);
    return {
      success: false,
      message: 'エラー: ' + e.message
    };
  }
}

/**
 * 履歴を取得
 */
function getHistory(filter) {
  try {
    const sheet = getOrCreateHistorySheet();
    const data = sheet.getDataRange().getValues();

    Logger.log('getHistory: データ行数 = ' + data.length);

    if (data.length <= 1) {
      return {
        success: true,
        records: [],
        message: '履歴がありません'
      };
    }

    let records = [];
    const filterTruck = filter && filter.truckNumber ? String(filter.truckNumber).trim() : null;
    const filterDate = filter && filter.date ? String(filter.date).trim() : null;

    Logger.log('getHistory: フィルタ - 号車=' + filterTruck + ', 日付=' + filterDate);

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];

      // 空行をスキップ
      if (!row[0]) continue;

      // 日付を適切に変換
      const recordDate = formatDateValue(row[1]);

      const record = {
        recordId: row[0] ? String(row[0]) : '',
        date: recordDate,
        truckNumber: row[2] ? String(row[2]).trim() : '',
        driverName: row[3] ? String(row[3]).trim() : '',
        vehicleNo: row[4] ? String(row[4]).trim() : '',
        yoshaCategory1: row[5] ? String(row[5]) : '',
        yoshaCategory2: row[6] ? String(row[6]) : '',
        helper: row[7] ? String(row[7]) : '',
        departureMeter: row[8] ? String(row[8]) : '',
        departureTime: formatTimeValue(row[9]),
        destination: row[10] ? String(row[10]) : '',
        returnMeter: row[11] ? String(row[11]) : '',
        centerArrivalTime: formatTimeValue(row[12]),
        officeArrivalTime: formatTimeValue(row[13]),
        hasCashOnDelivery: row[14] ? String(row[14]) : '',
        useTollRoad: row[15] ? String(row[15]) : '',
        tollRoadInfo: row[16] ? String(row[16]) : '',
        recordType: row[17] ? String(row[17]) : '',
        notes: row[18] ? String(row[18]) : '',
        timestamp: row[19] ? String(row[19]) : '',
        submitTiming: row[20] ? String(row[20]) : '営業所到着',
        travelTime: row[21] ? parseInt(row[21]) : 30
      };

      // フィルター適用
      if (filterTruck && record.truckNumber !== filterTruck) continue;
      if (filterDate && record.date !== filterDate) continue;

      records.push(record);
    }

    // 最新50件に制限
    records = records.slice(0, 50);

    Logger.log('getHistory: 取得件数 = ' + records.length);

    return {
      success: true,
      records: records
    };
  } catch (e) {
    Logger.log('getHistory error: ' + e.message);
    return {
      success: false,
      message: 'エラー: ' + e.message,
      records: []
    };
  }
}

/**
 * 既存レコードをチェック
 */
function checkExistingRecord(date, truckNumber, recordType) {
  try {
    const sheet = getOrCreateHistorySheet();
    const data = sheet.getDataRange().getValues();
    const searchDate = String(date).trim();
    const searchTruck = String(truckNumber).trim();

    for (let i = 1; i < data.length; i++) {
      const rowDate = formatDateValue(data[i][1]);
      const rowTruck = data[i][2] ? String(data[i][2]).trim() : '';
      const rowType = data[i][17] ? String(data[i][17]) : '';

      if (rowDate === searchDate && rowTruck === searchTruck && rowType === recordType) {
        return {
          exists: true,
          rowIndex: i + 1,
          recordId: data[i][0]
        };
      }
    }

    return { exists: false };
  } catch (e) {
    Logger.log('checkExistingRecord error: ' + e.message);
    return { exists: false };
  }
}

/**
 * 既存レコードを削除
 */
function deleteExistingRecord(date, truckNumber, recordType) {
  try {
    const check = checkExistingRecord(date, truckNumber, recordType);
    if (check.exists) {
      const sheet = getOrCreateHistorySheet();
      sheet.deleteRow(check.rowIndex);
      SpreadsheetApp.flush();
      return { success: true };
    }
    return { success: true };
  } catch (e) {
    Logger.log('deleteExistingRecord error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * 履歴をテキスト形式で取得（コピー用）
 */
function formatHistoryForCopy(records) {
  if (!records || records.length === 0) {
    return 'データがありません';
  }

  let text = '=== 配送業務履歴 ===\n\n';

  records.forEach((record, index) => {
    text += `【${index + 1}】${record.date}\n`;
    text += `種別: ${record.recordType}\n`;
    text += `号車: ${record.truckNumber}\n`;
    text += `ドライバー: ${record.driverName}\n`;

    if (record.recordType === '業務開始') {
      text += `車両No: ${record.vehicleNo}\n`;
      if (record.yoshaCategory1) text += `庸車区分①: ${record.yoshaCategory1}\n`;
      if (record.yoshaCategory2) text += `庸車区分②: ${record.yoshaCategory2}\n`;
      if (record.helper) text += `助手: ${record.helper}\n`;
      text += `出庫メーター: ${record.departureMeter}km\n`;
      text += `出庫時刻: ${record.departureTime}\n`;
    } else {
      if (record.destination) text += `行き先: ${record.destination}\n`;
      if (record.returnMeter) text += `帰庫メーター: ${record.returnMeter}km\n`;
      if (record.centerArrivalTime) text += `配送センター到着: ${record.centerArrivalTime}\n`;
      if (record.officeArrivalTime) text += `営業所帰庫: ${record.officeArrivalTime}\n`;
      if (record.hasCashOnDelivery) text += `代引き・返金: ${record.hasCashOnDelivery}\n`;
      if (record.useTollRoad === 'はい' && record.tollRoadInfo) {
        text += `有料道路: 使用あり\n`;
        try {
          const sections = JSON.parse(record.tollRoadInfo);
          sections.forEach((sec, idx) => {
            text += `  区間${idx + 1}: ${sec.start} → ${sec.end} (${sec.fee}円)\n`;
          });
        } catch (e) {}
      }
    }

    if (record.notes) text += `備考: ${record.notes}\n`;
    text += '\n';
  });

  return text;
}

/**
 * 管理者パスワードを検証
 */
function verifyAdminPassword(password) {
  return password === ADMIN_PASSWORD;
}

/**
 * 全ドライバー・車両マスタを取得（管理用）
 */
function getAllMasterEntries() {
  try {
    const sheet = getOrCreateTruckMasterSheet();
    const data = sheet.getDataRange().getValues();
    const entries = [];

    Logger.log('getAllMasterEntries: データ行数 = ' + data.length);

    for (let i = 1; i < data.length; i++) {
      // 空行をスキップ
      if (!data[i][0] && !data[i][1] && !data[i][2]) continue;

      const entry = {
        rowIndex: i + 1, // スプレッドシートの行番号（1始まり）
        truckNumber: data[i][0] ? String(data[i][0]).trim() : '',
        driverName: data[i][1] ? String(data[i][1]).trim() : '',
        vehicleNo: data[i][2] ? String(data[i][2]).trim() : '',
        timestamp: data[i][3] ? String(data[i][3]) : ''
      };
      entries.push(entry);
    }

    Logger.log('getAllMasterEntries: 取得件数 = ' + entries.length);

    return {
      success: true,
      entries: entries
    };
  } catch (e) {
    Logger.log('getAllMasterEntries error: ' + e.message);
    return {
      success: false,
      message: 'エラー: ' + e.message,
      entries: []
    };
  }
}

/**
 * マスタエントリを削除
 */
function deleteMasterEntry(rowIndex) {
  try {
    const sheet = getOrCreateTruckMasterSheet();
    sheet.deleteRow(rowIndex);
    return {
      success: true,
      message: '削除しました'
    };
  } catch (e) {
    return {
      success: false,
      message: 'エラー: ' + e.message
    };
  }
}

/**
 * 名前で号車マスタからエントリを削除
 */
function deleteByName(truckNumber, type, name) {
  try {
    const sheet = getOrCreateTruckMasterSheet();
    const data = sheet.getDataRange().getValues();
    const searchTruck = String(truckNumber).trim();
    const searchName = String(name).trim();

    // 下から上に削除（行番号がずれないように）
    for (let i = data.length - 1; i >= 1; i--) {
      const rowTruck = String(data[i][0]).trim();
      if (rowTruck !== searchTruck) continue;

      if (type === 'driver') {
        const driverName = data[i][1] ? String(data[i][1]).trim() : '';
        if (driverName === searchName) {
          sheet.deleteRow(i + 1);
          SpreadsheetApp.flush();
          return { success: true, message: '削除しました' };
        }
      } else if (type === 'vehicle') {
        const vehicleNo = data[i][2] ? String(data[i][2]).trim() : '';
        if (vehicleNo === searchName) {
          sheet.deleteRow(i + 1);
          SpreadsheetApp.flush();
          return { success: true, message: '削除しました' };
        }
      }
    }

    return { success: false, message: '該当するデータが見つかりません' };
  } catch (e) {
    Logger.log('deleteByName error: ' + e.message);
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 号車のデータを一括リセット（ドライバー・車両No・設定）
 */
function resetTruckData(truckNumber) {
  try {
    const searchTruck = String(truckNumber).trim();

    // 号車マスタからドライバー・車両を削除
    const masterSheet = getOrCreateTruckMasterSheet();
    const masterData = masterSheet.getDataRange().getValues();
    for (let i = masterData.length - 1; i >= 1; i--) {
      const rowTruck = String(masterData[i][0]).trim();
      if (rowTruck === searchTruck) {
        masterSheet.deleteRow(i + 1);
      }
    }

    // 号車設定をデフォルトに戻す
    const settingsSheet = getOrCreateTruckSettingsSheet();
    const settingsData = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < settingsData.length; i++) {
      const rowTruck = String(settingsData[i][0]).trim();
      if (rowTruck === searchTruck) {
        settingsSheet.getRange(i + 1, 2, 1, 3).setValues([['営業所到着', 30, '']]);
        break;
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: searchTruck + ' のデータをリセットしました' };
  } catch (e) {
    Logger.log('resetTruckData error: ' + e.message);
    return { success: false, message: 'エラー: ' + e.message };
  }
}
