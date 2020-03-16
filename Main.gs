/**
 * フォームの回答受信時に実行する
 * トリガー設定(池田)
 */
function updateForm() {
  var formIndex = {};
  var data = null;
  
  // どのフォームが更新されたかでデータを変える
  data = koukokuFormSheet.getNewData();
  if (data != null) {
    formIndex = koukokuFormSheet.getIndex();
    formIndex.affiliation = 100; // 存在しない項目をセット
    data[formIndex.affiliation] = '広告本部';
  } else {
    data = caFormSheet.getNewData();
    if (data != null) {
      formIndex = caFormSheet.getIndex();
    } else {
      data = aiFormSheet.getNewData();
      if (data != null) {
        formIndex = aiFormSheet.getIndex();
        formIndex.affiliation = 100; // 存在しない項目をセット
        data[formIndex.affiliation] = 'AI事業本部';
      }
    }
  }
  if (data === null) return;
  
  // 受信したフォームを依頼待機シートに移す
  var lastRow = ordersSheet.sheet.getRange('B:B').getValues().filter(String).length + 1;
  Object.keys(formIndex).forEach(function (key) {
    if (key === 'task' || key === 'sameOrNot') return;
    ordersSheet.sheet.getRange(ordersSheet.getRowKey(key) + lastRow).setValue(data[formIndex[key]]);
  });
  // オーダーNoをセット
  ordersSheet.sheet.getRange(ordersSheet.getRowKey('orderNo') + lastRow).setValue(lastRow);
  
  botWhenRequestComes(data, formIndex);
}

/**
 * 申請が来たときに通知するBOT
 */
function botWhenRequestComes(data, index) { // PC依頼用BOTで送る
  //WorkplaceApi.postBotForArms(
  WorkplaceApi.postBotForTest(
    '# ' + data[index.requesterName] + 'さんからPC配布依頼が来ました。\n' +
    '```\n' + 
    '依頼者      　 ： ' + data[index.requesterName] + '\n' +
    '所属　       　： ' + data[index.affiliation] + '\n' + 
    'PCの交換依頼か　： ' + data[index.isExchange] + '\n' +
    '希望PC        ： ' + data[index.requestPc] + '\n' +
    '申請理由　　　　： ' + data[index.reason] + '\n' +
    'ご要望・ご連絡　　： ' + data[index.request] + '\n' +
    '特定期間の利用か： ' + data[index.limitedTime] + '\n' +
    '```\n担当者はこの投稿にリアクションした上、シートの担当者欄に自身の名前を入れてください。\n▶[シートを確認する](https://docs.google.com/spreadsheets/d/' + MY_SHEET_ID + '/edit#gid=1398613080)'
   );// , 'pc');
}

/**
 * 依頼者にPC提案のメールを送る
 */
function sendMail() {
  var rowNum = Browser.inputBox('依頼者にPC提案のメールを送ります', '対応する行数を半角数字で入力してください。', Browser.Buttons.OK);
  var data = createMailData(rowNum);
  if (!data || data.text === '') return;
  
  // メールを送る
  var index = ordersSheet.getIndex();
  var target = ordersSheet.values[Number(rowNum) - 1];
  var sendSuccess = mailSheet.sendMail(target[index.requesterMail], target[index.userMail], data.title, data.text, true);
  if (!sendSuccess) return;
  
  ordersSheet.sheet.getRange(ordersSheet.getRowKey('mailDate') + rowNum).setValue(Utilities.formatDate(new Date(), 'JST', 'MM/dd HH:mm'));
  ordersSheet.sheet.getRange(ordersSheet.getRowKey('mailText') + rowNum).setValue(data.text); 
}

/**
 * メール本文を作る
 */
function createMailText() {
  var rowNum = Browser.inputBox('メール本文を作成します。', '対応する行数を半角数字で入力してください。', Browser.Buttons.OK);
  var data = createMailData(rowNum);
  if (!data || data.text === '') return;
  ordersSheet.sheet.getRange(ordersSheet.getRowKey('mailText') + rowNum).setValue(data.text);
  Browser.msgBox(rowNum + '行目【メール文章】欄に入力されました。ご確認ください。');
}

/**
 * メールデータを作って返す
 */
function createMailData(rowNum) {
  var target = ordersSheet.values[Number(rowNum) - 1];
  if (!target) { Browser.msgBox('データが見つかりません'); return null; }
  
  var index = ordersSheet.getIndex();
  
  if (target[index.mailDate] != '') { Browser.msgBox(rowNum + '行目はすでにメールを送っています。'); return null; }
  if (target[index.checkPerson] === '') { Browser.msgBox(rowNum + '行目は担当者が書かれていません。記入してやり直してください。'); return null; }
  
  var candidate1 = target[index.candidate1];
  if (candidate1 === '') { Browser.msgBox(rowNum + '行目は提案するPC情報が書かれていません。記入してやり直してください。'); return null; }
  
  var popup = Browser.msgBox(target[index.requesterName] + 'さんの依頼に返事をします。', '実行してよろしいですか？ ' + TEXT_MAIL_SETTING, Browser.Buttons.OK_CANCEL);
  if (popup != 'ok') return null;
  
  var mailType = (candidate1.indexOf('CA-') === 0) ? 'proposalByStock' : 'newRental';
  return mailSheet.createMailData(target, mailType);
}

/**
 * メニューを設定する
 * トリガー登録しています。(池田)
 */
function createMenu() {
  var ui = SpreadsheetApp.getUi();      
  var menu = ui.createMenu('▼スクリプト');
  // メニューにアイテムを追加する
  menu.addItem('PC案内メールを送る', 'sendMail');
  menu.addItem('PC案内メール本文を作る', 'createMailText');
  menu.addToUi();
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
    request      : filterData.indexOf('ご要望・ご連絡'),
    limitedTime  : filterData.indexOf('特定期間の利用ですか？'),
    place        : filterData.indexOf('希望受取拠点')
  };
}

function showTitleError(key) {
  Browser.msgBox('データが見つかりません', '表のタイトル名を変えていませんか？ : ' + key, Browser.Buttons.OK);
}

/**
 * テストメールを送ります
 * トリガー：「はじめに」シート設置のボタンより
 */
function sendTestMail() {
  var adress = Browser.inputBox('テストメールを送信しますか？', 'テスト用のメールアドレスを入力してください。(個人のアドレスを推奨します)' + TEXT_MAIL_SETTING, Browser.Buttons.OK_CANCEL);
  if (adress === 'cancel') return;
  mailSheet.sendMail(adress, '', 'テストメール', 'infosusからテストメールを送っています。', true);
}