var OrdersSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('依頼');
  this.values = this.sheet.getDataRange().getValues();
  this.index = {};
  this.titleRow = 0;
  
  this.createIndex = function() {
    const PERSON = '担当者';
    var me = this;
    var filterData = (function() {
      for(var i = 0; i < me.values.length; i++) {
        if (me.values[i].indexOf(PERSON) > -1) {
          me.titleRow = i + 1;
          return me.values[i];
        }
      }
    }());
    if(!filterData || filterData.length === 0) {
      showTitleError();
      return;
    }
    var myIndex = {
      checkPerson: filterData.indexOf(PERSON),
      candidate1 : filterData.indexOf('PC候補1'),
      candidate2 : filterData.indexOf('PC候補2'),
      candidate3 : filterData.indexOf('PC候補3'),
      candidate4 : filterData.indexOf('PC候補4'),
      candidate5 : filterData.indexOf('PC候補5'),
      candidate6 : filterData.indexOf('PC候補6'),
      message    : filterData.indexOf('依頼者へのメッセージ'),
      mailKind   : filterData.indexOf('メール種類'),
      memo       : filterData.indexOf('メモ'),
      mailDate   : filterData.indexOf('確認メール送信日時'),
      mailText   : filterData.indexOf('メール文章'),
      orderNo    : filterData.indexOf('オーダーNo')
    }
    this.index = Object.assign(getFormIndex(filterData), myIndex);
    return this.index;
  }
}
  
OrdersSheet.prototype = {
  getRowKey: function(target) {
    var index = this.getIndex();
    var targetIndex = index[target];
    var returnKey = (targetIndex > -1) ? SHEET_ROWS[targetIndex] : '';
    if (!returnKey || returnKey === '') showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  }
};

var ordersSheet = new OrdersSheet();

function ordersSheetTest() {
  var inde = ordersSheet.getIndex();
  var hoe = '';
}