var MailSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('メールテンプレート');
  this.values = this.sheet.getDataRange().getValues();
  this.index = {};
  this.rowIndex = { title: 4, text: 5 };
  this.infosysAdress = this.values[0][3];
  
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
      proposalByStock: filterData.indexOf('提案(在庫)')
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
   * @pram type string 以下のいずれか{proposalByStock}
   * @return { title: string, text: string }
   */
  createMailData: function(data, type) {
    var index = this.getIndex();
    return {
      title: this.values[this.rowIndex.title][index[type]].replace('{userName}', data[ordersSheet.getIndex().requesterName]),
      text: this.replaceText(this.values[this.rowIndex.text][index[type]], data)
    };
  },
  replaceText: function(text, data) {
    var index = ordersSheet.getIndex();
    var pcText = this.createPcText([data[index.candidate1], data[index.candidate2], data[index.candidate3], data[index.candidate4], data[index.candidate5], data[index.candidate6]]);
    if (pcText === '') return '';
    return text
      .replace('{userName}', data[index.requesterName] + 'さん' + (data[index.userName] === '' ? '' : '\n  CC: ' + data[index.userName]))
      .replace('{checkPerson}', data[index.checkPerson])
      .replace('{message}', data[index.message])
      .replace('{pcInfo}', pcText)
      .replace('{date}', data[index.timeStamp])
      .replace('{orderNo}', data[index.orderNo]);
  },
  /**
   * PC情報
   * @param candidates [String] PC候補の番号たち
   */
  createPcText: function(candidates) {
    var pcData = KintonePCData.pcDataSheet.values;
    var pcTitles = KintonePCData.pcDataSheet.getTitles();
    var templates = this.values[8];
    var hasError = false;
    var hasRental = false;
    
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
      var text = '\n' + num + templates[1]
        .replace('{maker}', data[pcTitles.pc_maker.index])
        .replace('{product}', data[pcTitles.pc_product.index])
        .replace('{model}', data[pcTitles.pc_model.index])
        .replace('{capc_id}', data[pcTitles.capc_id.index])
        .replace('{pc_display}', data[pcTitles.pc_display.index] + 'インチ')
        .replace('{keyboard}', data[pcTitles.keyboard.index])
        .replace('{cpu}', data[pcTitles.cpu.index])
        .replace('{memory}', data[pcTitles.memory.index] + 'GB')
        .replace('{ssd}', data[pcTitles.ssd.index]);
        
      // 初回費用負担ありの場合
      if (data[pcTitles.paid_in_adv.index] === 'あり') {
        var money = data[pcTitles.purchase_amount.index];
        if(money === '') {
          Browser.msgBox('台帳の【購入金額】を記入してください', '費用負担PCのため金額を利用者に提示します。' + candidate + 'の台帳の購入金額を入れてからやり直してください。', Browser.Buttons.OK);
          hasError = true;
          return;
        }
        text += '\n' + templates[2].replace('{money}', money);
      }
        
      // レンタルの場合
      if (data[pcTitles.rentalid.index] != '') {
        hasRental = true;
        text += '\n' + templates[3]
          .replace('{rental_end}', data[pcTitles.rental_end.index])
          .replace('{rental_fee}', data[pcTitles.rental_fee.index]);
      }
        
      return text;
    });
    if (hasError) return '';
    if (hasRental) returnText.push(templates[4])
    return returnText.join('\n');
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

function mailTest() {
  var data = mailSheet.createMailData(ordersSheet.values[5], 'proposalByStock');
  mailSheet.sheet.getRange('C6').setValue(data.text)
  var hoge = ''
}