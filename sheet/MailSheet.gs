var MailSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('メールテンプレート');
  this.values = this.sheet.getDataRange().getValues();
  this.index = {};
  this.rowIndex = { title: 4, text: 5 };
  this.infosysAdress = this.values[0][3];
  this.PC_TEMPLATES = this.values[8];
  this.RENTAL_INFO = this.PC_TEMPLATES[4];
  
  this.createIndex = function() {
    const FIRST_TITLE = 'メール種類';
    var filterData = this.values.filter(function(value) {
      return value.indexOf(FIRST_TITLE) > -1;
    })[0];
    if(!filterData || filterData.length === 0) {
      showTitleError();
      return;
    }
    this.index = {
      proposalByStock : filterData.indexOf('提案(在庫)'),
      newRental       : filterData.indexOf('提案(新規レンタル)'),
      yokokawaExchange: filterData.indexOf('横河レンタルPC修理手配案内')
    };
    return this.index;
  }
}
  
MailSheet.prototype = {
  getRowKey: function(target) {
    var index = this.getIndex();
    var targetIndex = index[target];
    var returnKey = (targetIndex > -1) ? SHEET_ROWS[targetIndex] : '';
    if (!returnKey || returnKey === '') showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  },
  /**
   * メール件名とテキストを返します
   * @param array 依頼待機リストの１行分のデータ
   * @pram type string 以下のいずれか{proposalByStock or newRental or yokokawaExchange}
   * @return { title: string, text: string }
   */
  createMailData: function(data, type) {
    var index = this.getIndex();
    return {
      title: this.values[this.rowIndex.title][index[type]].replace('{userName}', data[ordersSheet.getIndex().requesterName]).replace('{orderNo}', data[ordersSheet.getIndex().orderNo]),
      text: this.replaceText(this.values[this.rowIndex.text][index[type]], data, type)
    };
  },
  replaceText: function(text, data, type) {
    var index = ordersSheet.getIndex();
    var candidates = [data[index.candidate1], data[index.candidate2], data[index.candidate3], data[index.candidate4], data[index.candidate5], data[index.candidate6]];
    var pcText = '';
    if(type === 'proposalByStock') pcText = this.createPcText(candidates);
    else if (type === 'newRental') pcText = rentalSheet.createRentalPcText(candidates);
    if ((type === 'proposalByStock' || type === 'newRental') && pcText === '') return '';
    
    return text
      .replace('{userName}', data[index.requesterName] + 'さん' + (data[index.userName] === '' ? '' : '\n  CC: ' + data[index.userName] + 'さん'))
      .replace('{checkPerson}', data[index.checkPerson])
      .replace('{message}', data[index.message])
      .replace('{pcInfo}', pcText)               // PC在庫のみにある項目
      .replace('{rentalPcInfo}', pcText)         // レンタル案内のみにある項目
      .replace('{rentalInfo}', this.RENTAL_INFO + '\n') // レンタル案内のみにある項目
      .replace('{date}', Utilities.formatDate(data[index.timeStamp], 'JST', 'yyyy/MM/dd HH:mm'))
      .replace('{orderNo}', data[index.orderNo]);
  },
  /**
   * 在庫PC情報
   * @param candidates [String] PC候補の番号たち
   */
  createPcText: function(candidates) {
    var pcData = KintonePCData.pcDataSheet.values;
    var pcTitles = KintonePCData.pcDataSheet.getTitles();
    var hasError = false;
    var hasRental = false;
    var me = this;
    candidates = candidates.filter(function(text) { return text != ''; });
    
    var returnText = candidates.map(function(candidate, i) {
      if (candidate === '') return '';
      // PC情報をひっぱってくる
      var data = pcData.filter(function(value) { return value[pcTitles.capc_id.index] === candidate });
      if (data.length === 0) {
        Browser.msgBox('台帳データが見つかりません。', 'PC番号' + candidate + 'の記載は正しいでしょうか？確認し、やり直してください。', Browser.Buttons.OK);
        hasError = true;
        return;
      }
      data = data[0];
      var num = (candidates.length === 1) ? '' : '【' + (i + 1) + '】';
      var ssd = (data[pcTitles.ssd.index] != '') ? 'SSD-' + data[pcTitles.ssd.index] + 'GB' : '';
      var hdd = (data[pcTitles.hdd.index] != '') ? 'HDD-' + data[pcTitles.hdd.index] + 'GB' : '';
      
      var text = '\n' + num + me.PC_TEMPLATES[1]
        .replace('{maker}', data[pcTitles.pc_maker.index])
        .replace('{product}', data[pcTitles.pc_product.index])
        .replace('{model}', data[pcTitles.pc_model.index])
        .replace('{pc_display}', data[pcTitles.pc_display.index] === 'なし' ? 'デスクトップ' : data[pcTitles.pc_display.index] + 'インチ')
        .replace('{keyboard}', data[pcTitles.keyboard.index])
        .replace('{cpu}', data[pcTitles.cpu.index])
        .replace('{memory}', data[pcTitles.memory.index])
        .replace('{capacity}', ssd + ((ssd != '' && hdd != '') ? ' / ' : '') + hdd);
        
      // 初回費用負担ありの場合
      if (data[pcTitles.paid_in_adv.index] === 'あり') {
        var money = data[pcTitles.purchase_amount.index];
        if(money === '' || money === '0') {
          Browser.msgBox('台帳の【購入金額】を記入してください', '費用負担PCのため金額を利用者に提示します。' + candidate + 'の台帳の購入金額を入れてからやり直してください。', Browser.Buttons.OK);
          hasError = true;
          return;
        }
        text += '\n' + me.PC_TEMPLATES[2].replace('{money}', money);
      }
        
      // レンタルの場合
      if (data[pcTitles.rentalid.index] != '') {
        var end = data[pcTitles.rental_end.index];
        var fee = data[pcTitles.rental_fee.index];
        if (end === '' || fee === '' || fee == 0) {
          Browser.msgBox('台帳の【レンタル終了日】【レンタル料金】を確認してください', 'どちらかが空欄になっていると実行できません。台帳を更新し、やり直してください。', Browser.Buttons.OK);
          hasError = true;
          return;
        }
        hasRental = true;
        text += '\n' + me.PC_TEMPLATES[3]
          .replace('{rental_end}', Utilities.formatDate(end, 'JST', 'yyyy/MM/dd')).replace('{rental_fee}', fee);
      }
      return text;
    });
    if (hasError) return '';
    if (hasRental) returnText.push('\n' + this.RENTAL_INFO + '\n');
    return returnText.join('\n') + '\n';
  },
  /**
   * メールを内容確認ポップアップ
   * @return boolean
   */
  openMailCheckPopup: function(text) {
    var popup = Browser.msgBox('以下の内容でメールを送信します。よろしいでしょうか？(実際には改行されます)', text, Browser.Buttons.OK_CANCEL);
    return popup === 'ok';
  },
  /**
   * メール送信
   * @param to string 送信先
   * @param cc string 送信先
   * @param title string 件名
   * @param text string 本文
   * @param showPopup boolean 確認popupを出すか否か
   */
  sendMail: function(to, cc, title, text, showPopup) {
    if (showPopup && !this.openMailCheckPopup(text)) return false;
    
    try {
      GmailApp.sendEmail(to, title, text, {
        from: this.infosysAdress,
        replyTo: this.infosysAdress,
        name: '全社システム本部 資産管理チーム',
        cc: cc
      });
    } catch(e) {
      return false;
    }
    return true;
  }
};
var mailSheet = new MailSheet();

/**
 * 依頼者にPC提案のメールを送る
 */
function sendMail() {
  var rowNum = Browser.inputBox('依頼者にPC提案のメールを送ります', '対応する行数を半角数字で入力してください。', Browser.Buttons.OK);
  var data = createMailData(rowNum);
  if (!data || data.text === '') return;
  
  sendMailForUser(rowNum, data); // メールを送る
}

/**
 * メール本文を作る
 */
function createMailText() {
  var rowNum = Browser.inputBox('メール本文を作成します。', '対応する行数を半角数字で入力してください。', Browser.Buttons.OK);
  var data = createMailData(rowNum);
  if (data && data.text != '') setMailText(rowNum, data.text);;
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
 * 横河の故障交換を案内するメールを送る
 */
function sendYokokawaExchangeMail() {
  var rowNum = Browser.inputBox('依頼者に横河の故障交換を案内するメールを送ります', '対応する行数を半角数字で入力してください。', Browser.Buttons.OK);
  var data = createRentalMailData(rowNum);
  if (!data || data.text === '') return;
  
  sendMailForUser(rowNum, data); // メールを送る
}

/**
 * 横河の故障交換を案内するメールを作る
 */
function createYokokawaExchangeMail() {
  var rowNum = Browser.inputBox('依頼者に横河の故障交換を案内するメール本文を作成します。', '対応する行数を半角数字で入力してください。', Browser.Buttons.OK);
  var data = createRentalMailData(rowNum);
  if (data && data.text != '') setMailText(rowNum, data.text);
}

/**
 * 横河故障交換のメールデータを作って返す
 */
function createRentalMailData(rowNum) {
  var target = ordersSheet.values[Number(rowNum) - 1];
  if (!target) { Browser.msgBox('データが見つかりません'); return null; }
  
  var index = ordersSheet.getIndex();
  
  if (target[index.mailDate] != '') { Browser.msgBox(rowNum + '行目はすでにメールを送っています。'); return null; }
  if (target[index.checkPerson] === '') { Browser.msgBox(rowNum + '行目は担当者が書かれていません。記入してやり直してください。'); return null; }
  
  var popup = Browser.msgBox(target[index.requesterName] + 'さんの依頼に横河の故障交換案内を返します。', '実行してよろしいですか？ ' + TEXT_MAIL_SETTING, Browser.Buttons.OK_CANCEL);
  if (popup != 'ok') return null;
  
  return mailSheet.createMailData(target, 'yokokawaExchange');
}

/**
 * 利用者にメールを送る。成功したらシートに書き込む
 */
function sendMailForUser(rowNum, data) {
  var index = ordersSheet.getIndex();
  var target = ordersSheet.values[Number(rowNum) - 1];
  var sendSuccess = mailSheet.sendMail(target[index.requesterMail], target[index.userMail], data.title, data.text, true);
  if (!sendSuccess) return;
  ordersSheet.sheet.getRange(ordersSheet.getRowKey('mailDate') + rowNum).setValue(Utilities.formatDate(new Date(), 'JST', 'MM/dd HH:mm'));
  ordersSheet.sheet.getRange(ordersSheet.getRowKey('mailText') + rowNum).setValue(data.text); 
}

function setMailText(rowNum, text) {
  ordersSheet.sheet.getRange(ordersSheet.getRowKey('mailText') + rowNum).setValue(text);
  Browser.msgBox(rowNum + '行目【メール文章】欄に入力されました。ご確認ください。');
}