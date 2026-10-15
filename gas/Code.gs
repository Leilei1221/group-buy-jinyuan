/**
 * 幾隻貓×金園排骨 補貨團 — GAS 訂單接收
 * 部署為網頁應用程式，執行身分：我，存取者：所有人
 */

const PRODUCT_LIST = [
  { id: 'p01', name: '綜合十果 2小包', price: 330 },
  { id: 'p02', name: '綜合十果 1盒', price: 1085 },
  { id: 'p03', name: '鹹酥雞腿肉 8包', price: 1998 },
  { id: 'p04', name: '鹽酥雞腿肉6包+起司披薩2片', price: 1948 },
  { id: 'p05', name: '厚切排骨12片＋金饌雞腿12片', price: 2990 },
  { id: 'p06', name: '排骨12+雞腿12+鹽酥雞2+披薩2', price: 3680 },
  { id: 'p07', name: '厚切排骨 12片', price: 1476 },
  { id: 'p08', name: '金饌雞腿 12片', price: 1668 },
  { id: 'p09', name: '厚切排骨6片+金饌雞腿6片', price: 1572 },
  { id: 'p10', name: '排骨5+雞腿5+鯖魚3', price: 1950 },
  { id: 'p11', name: '排骨4+雞腿4+鹽酥雞2', price: 1680 },
  { id: 'p12', name: '排骨3+雞腿3+鹽酥雞1+披薩1', price: 1340 },
  { id: 'p13', name: '鹹酥雞腿肉', price: 250 },
  { id: 'p14', name: '厚切里肌', price: 263 },
  { id: 'p15', name: '起司披薩1片+鹽酥雞1包', price: 440 },
  { id: 'p16', name: '挪威鯖魚片', price: 160 },
  { id: 'p17', name: '特級麻辣養生鍋底', price: 380 },
  { id: 'p18', name: '泰式酸辣鍋底', price: 380 },
  { id: 'p19', name: '義式蕃茄鍋底', price: 380 },
  { id: 'p20', name: '養生麻油酒香鍋底', price: 380 },
];

function getOrCreateSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch(e) { /* recreate */ }
  }
  const ss = SpreadsheetApp.create('金園排骨補貨團-訂單');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  const sheet = ss.getActiveSheet();
  sheet.setName('訂單');
  setupHeaders(sheet);
  return ss;
}

function setupHeaders(sheet) {
  // 欄位：時間戳記 | 姓名 | 各商品數量 | 訂購摘要 | 總金額（元）| 備註及回覆
  const headers = ['時間戳記', '姓名'];
  PRODUCT_LIST.forEach(p => headers.push(p.name + '\n$' + p.price));
  headers.push('訂購摘要', '總金額（元）', '備註及回覆');

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1a5c2a');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setRowHeight(1, 48);

  // 欄寬
  sheet.setColumnWidth(1, 155); // 時間
  sheet.setColumnWidth(2, 75);  // 姓名
  for (let i = 0; i < PRODUCT_LIST.length; i++) {
    sheet.setColumnWidth(3 + i, 78);
  }
  const summaryCol = 3 + PRODUCT_LIST.length;
  sheet.setColumnWidth(summaryCol, 220);     // 訂購摘要
  sheet.setColumnWidth(summaryCol + 1, 90);  // 總金額
  sheet.setColumnWidth(summaryCol + 2, 160); // 備註及回覆

  // 凍結第1列 + 前2欄
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

function doPost(e) {
  const params = e.parameter;
  const name = params.name || '';
  const summary = params.summary || '';
  const total = parseInt(params.total || '0', 10);
  const note = params.note || '';

  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('訂單') || ss.getActiveSheet();

  // 檢查標題列是否存在
  if (!sheet.getRange('C1').getValue()) {
    setupHeaders(sheet);
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  const row = [now, name];

  // 各商品數量
  const itemLines = [];
  PRODUCT_LIST.forEach(p => {
    const q = parseInt(params[p.id] || '0', 10);
    row.push(q > 0 ? q : '');
    if (q > 0) itemLines.push(p.name + ' ' + q + '份 單價' + p.price + '元');
  });

  // 備註及回覆
  const replyMsg = name + '你好，你所訂購的產品如下：' + itemLines.join('、') + '、總金額：' + total + '元' + (note ? '\n備註：' + note : '');

  row.push(summary, total, replyMsg);

  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);

  // 數量欄置中
  const qtyRange = sheet.getRange(newRow, 3, 1, PRODUCT_LIST.length);
  qtyRange.setHorizontalAlignment('center');

  // 交替底色
  const bgColor = (newRow % 2 === 0) ? '#e8f5ea' : '#ffffff';
  sheet.getRange(newRow, 1, 1, row.length).setBackground(bgColor);

  // 訂購摘要 wrap
  const summaryCol = 3 + PRODUCT_LIST.length;
  sheet.getRange(newRow, summaryCol).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('金園排骨補貨團 GAS 運作中');
}

function getSpreadsheetUrl() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID');
  if (ssId) {
    Logger.log('試算表網址：https://docs.google.com/spreadsheets/d/' + ssId);
  } else {
    Logger.log('尚未建立試算表，請先送出一筆訂單');
  }
}
