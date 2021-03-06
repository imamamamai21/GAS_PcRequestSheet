var FormSheet = function(sheetName) {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('回答(' + sheetName + ')');
  this.values = this.sheet.getDataRange().getValues();
  this.index = {};
  this.titleRow = 2;
  
  this.createIndex = function() {
    var filterData = this.values.filter(function(value) {
      return value.indexOf('タイムスタンプ') > -1;
    })[0];
    if(!filterData || filterData.length === 0) {
      showTitleError();
      return;
    }
    this.index = Object.assign(getFormIndex(filterData), { task: filterData.indexOf('タスク化') } );
    return this.index;
  }
}
  
FormSheet.prototype = {
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
   * 「タスク化」欄が「済」になっていないデータを返す
   */
  getNewData: function() {
    var lastRow = this.sheet.getRange('A:A').getValues().filter(String).length;
    if (this.values[lastRow - 1][this.getIndex().task] != '') return null;

    // タスク化済にチェックを入れて対象データを返す
    this.sheet.getRange(this.getRowKey('task') + lastRow).setValue('済');
    return this.values[lastRow - 1];
  }
};

var koukokuFormSheet = new FormSheet('広告本部');
var caFormSheet      = new FormSheet('その他CA');
var aiFormSheet      = new FormSheet('AI事業本部');
//var companyFormSheet = new FormSheet('子会社');

function formTest() {
  var f = koukokuFormSheet.getNewData();
  var hoge = ''
}