/**
 * フォームの回答受信時に実行する
 * トリガー設定(池田)
 */
function submitForm() {
  // どのフォームが更新されたかを判定
  var formIndex = koukokuFormSheet.getIndex();
  var data = koukokuFormSheet.getUpdateForm() || caFormSheet.getUpdateForm();
  if(data === null) return;
  
  // 受信したフォームを依頼待機シートに移す
  var lastRow = ordersSheet.sheet.getRange('B:B').getValues().filter(String).length + 1;
  Object.keys(formIndex).forEach(function (key) {
    if (key === 'task') return;
    ordersSheet.sheet.getRange(ordersSheet.getRowKey(key) + lastRow).setValue(data[formIndex[key]]);
  });
  botWhenRequestComes(data, formIndex);
}

/**
 * 申請が来たときに通知するBOT
 */
function botWhenRequestComes(data, index) { // PC依頼用BOTで送る
  WorkplaceApi.postBotForArms('# ' + data[index.requesterName] + 'さんからPC配布依頼が来ました。\n' +
    '```\n' + 
    '依頼者      　 ： ' + data[index.requesterName] + '\n' +
    '所属　       　： ' + data[index.affiliation] + '\n' + 
    'PCの交換依頼か　： ' + data[index.isExchange] + '\n' +
    '希望PC       　： ' + data[index.requestPc] + '\n' +
    '申請理由　　　　： ' + data[index.reason] + '\n' +
    '希望PC　　　　　： ' + data[index.requestPc] + '\n' +
    '特定期間の利用か： ' + data[index.limitedTime] + '\n' +
    '```\n担当者はこの投稿にリアクションした上、シートの担当者欄に自身の名前を入れてください。\n▶[シートを確認する](https://docs.google.com/spreadsheets/d/' + MY_SHEET_ID + '/edit#gid=1398613080)'
    , 'pc');
}

/**
 * 依頼者にPC提案のメールを送る
 */
function sendMailProposalByStock() {
  var rowNum = Browser.inputBox('依頼者にPC提案のメールを送ります', '対応する行数(orオーダーNO)を入力してください。', Browser.Buttons.OK);
  var target = ordersSheet.values[Number(rowNum) - 1];
  if (!target) { Browser.msgBox('データが見つかりません'); return; }
  
  var index = ordersSheet.getIndex();
  if (target[index.mailDate] != '') { Browser.msgBox(rowNum + '行目はすでにメールを送っています。'); return; }
  
  var popup = Browser.msgBox(target[index.requesterName] + 'さんの依頼に返事をします。', '実行してよろしいですか？ ', Browser.Buttons.OK_CANCEL);
  if (popup != 'ok') return;
  
  var data = mailSheet.createMailData(target, 'proposalByStock');
  if (data.text === '') return;
  // メールを送る(キャンセルの場合return)
  if (!mailSheet.sendMail(target[index.requesterMail], target[index.userMail], data.title, data.text, true)) return;
  
  ordersSheet.sheet.getRange(ordersSheet.getRowKey('mailDate') + rowNum).setValue(Utilities.formatDate(new Date(), 'JST', 'MM/dd(E) HH:mm'));
  ordersSheet.sheet.getRange(ordersSheet.getRowKey('mailText') + rowNum).setValue(data.text); 
}

/**
 * メニューを設定する
 * トリガー登録しています。(池田)
 */
function createMenu() {
  var ui = SpreadsheetApp.getUi();      
  var menu = ui.createMenu('▼スクリプト');
  // メニューにアイテムを追加する
  menu.addItem('在庫の案内メールを送る', 'sendMailProposalByStock');
  menu.addToUi(); // メニューをUiクラスに追加する
}

function getFormIndex(filterData) {
  return {
    timeStamp    : filterData.indexOf('タイムスタンプ'),
    requesterMail: filterData.indexOf('メールアドレス'),
    sameOrNot    : filterData.indexOf('依頼者と利用者は'),
    requesterNo  : filterData.indexOf('依頼者の社員番号'),
    requesterName: filterData.indexOf('依頼者の氏名'),
    userNo       : filterData.indexOf('利用者の社員番号'),
    userName     : filterData.indexOf('利用者の氏名'),
    userMail     : filterData.indexOf('利用者のメールアドレス'),
    superiorNo   : filterData.indexOf('利用者の上長の社員番号'),
    superiorName : filterData.indexOf('利用者の上長の氏名'),
    affiliation  : filterData.indexOf('利用者の所属'),
    userSection  : filterData.indexOf('PC利用者の区分'),
    isExchange   : filterData.indexOf('PCの交換依頼ですか？'),
    requestPc    : filterData.indexOf('希望PC'),
    reason       : filterData.indexOf('申請理由'),
    limitedTime  : filterData.indexOf('特定期間の利用ですか？'),
    place        : filterData.indexOf('希望受取拠点')
  };
}

function showTitleError(key) {
  Browser.msgBox('データが見つかりません', '表のタイトル名を変えていませんか？ : ' + key, Browser.Buttons.OK);
}