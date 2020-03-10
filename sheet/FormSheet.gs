var FormSheet = function(className) {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('フォーム回答(' + className + ')');
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
    this.index = Object.assign(getFormIndex(filterData), { task: filterData.indexOf('タスク化済み') } );
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
  getUpdateForm: function() {
    var index = this.getIndex();
    var lastRow = this.sheet.getRange('A:A').getValues().filter(String).length;
    if (this.values[lastRow - 1][this.index.task] != '') return null;

    // タスク化済にチェックを入れて対象データを返す
    this.sheet.getRange(this.getRowKey('task') + lastRow).setValue(true);
    return this.values[lastRow - 1];
  }
};

var koukokuFormSheet = new FormSheet('広告本部');
var caFormSheet = new FormSheet('その他CA');

function formTest() {
}