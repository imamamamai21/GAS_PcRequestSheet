var RentalSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('レンタルカタログ');
  this.values = this.sheet.getDataRange().getValues();
  
  /**
   * @param key {string} 商品名(A列)
   * @return string レンタルPCの詳細(B列)
   */
  this.getDataByKey = function(key) {
    var targetData = this.values.slice(2).filter(function(value) { return value[0] === key; });
    if (targetData.length === 0) {
      Browser.msgBox('データが見つかりません',
        '指定は【' + key + '】で合っていますでしょうか？\nレンタルの場合【レンタルカタログの表タイトル】と、\n在庫の案内の場合【台帳のCAグループPC管理番号】と一致するか確認してください。',
        Browser.Buttons.OK);
      return null;
    }
    return targetData[0][1];
  }
  /**
   * レンタル情報を返します
   * @param candidates [商品名(A列)] PC候補の番号たち
   * @return string
   */
  this.createRentalPcText = function(candidates) {
    var returnText = [];
    var hasError = false;
    var me = this;
    // 空欄を除く
    candidates = candidates.filter(function(candidate) { return candidate != ''; });
    
    candidates.forEach(function(candidate, i) {
        var text = '\n　';
        var rentalText = me.getDataByKey(candidate);
        hasError = rentalText === null;
        text += (candidates.length === 1) ? '' : '【' + (i + 1) + '】';
        text += rentalText;
        returnText.push(text);
      });
      
    if (hasError) return '';
    return returnText.join('\n') + '\n';
  }
}
var rentalSheet = new RentalSheet();

function renTest() {
}