function doGet() {
  return HtmlService.createHtmlOutputFromFile("index").setTitle("北薬データベース");
}

function saveData(e, t, n, a) {
  var r = Session.getActiveUser().getEmail();
  
  if (!r.endsWith("@elms.hokudai.ac.jp")) {
    return "エラー: 学内メールアドレスでログインしてください。";
  }

  var i = SpreadsheetApp.openById("1TbiLnLd5zbzDcFYL7PB8zACjdBi964gbUP1VqbUbQng");
  var d = i.getSheetByName(e);

  if (!d) {
    d = i.insertSheet(e);
    d.appendRow(["配属年度", "日時", "薬科学科", "薬学科", "メールアドレス"]);
  }

  var o = d.getDataRange().getValues();

  for (var p = 1; p < o.length; p++) {
    if (o[p][4] === r) {
      return "エラー: あなたはすでにデータを送信済みです。";
    }
  }

  var newData = [a, new Date(), "", "", r];
  
  if (t === "薬科学科") {
    newData[2] = n;
  } else if (t === "薬学科") {
    newData[3] = n;
  }

  d.appendRow(newData);

  return "データを保存しました！";
}
