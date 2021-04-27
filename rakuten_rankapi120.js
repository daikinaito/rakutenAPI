function myFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ランキング");  // 商品シートを取得する（商品シートが無いと次の行でエラーになる）
  sheet.getRange(1, 1).setValue("日付");  // A1セル（1行目の1列目）に「画像」と入力する
  sheet.getRange(1, 2).setValue("ランク");  // A1セル（1行目の1列目）に「画像」と入力する
  sheet.getRange(1, 3).setValue("商品名");
  sheet.getRange(1, 4).setValue("商品URL");  // B1セル（1行目の2列目）に「商品名」と入力する

  // APIテストフォームで表示されたURL（applicationId=以降は自分のアプリIDに置き換えてください）
  for (var page = 1; page < 5; page++) {
    var pagest =String(page);
    var url = "https://app.rakuten.co.jp/services/api/IchibaItem/Ranking/20170628?format=json&genreId=214122&page="+pagest+"&applicationId=自分のアプリID";
    var response = UrlFetchApp.fetch(url);  // 指定したURLのページを取得する
    var json = JSON.parse(response.getContentText());  // JSON形式の文字列をプログラムから扱えるようパースする
    var formatDate = Utilities.formatDate(new Date(), "JST","MM/dd");
    var lastRow = sheet.getLastRow();

    // すべての商品について反復処理する
    json.Items.forEach(function(item, i) {
      // 個別の商品の処理を書く。iは0から29まで。
      var row = lastRow + 1 + i;  // 行番号
      sheet.setRowHeight(row, 100);  // 行の高さを100pxにする
      sheet.getRange(row, 1).setValue(formatDate); 
      sheet.getRange(row, 2).setValue(item.Item.rank);  // A列に画像を表示する
      sheet.getRange(row, 3).setValue(item.Item.itemName);  // B列に商品名を入力する
      sheet.getRange(row, 4).setValue(item.Item.itemUrl);  // B列に商品名を入力する
    });
  }

}
