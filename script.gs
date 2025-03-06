function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('北薬データベース');
}

function saveData(lab, gakka, gpa, year) {
  var userEmail = Session.getActiveUser().getEmail();

  if (!userEmail.endsWith("@elms.hokudai.ac.jp")) {
    return "エラー: 学内メールアドレスでログインしてください。";
  }

  var spreadsheet = SpreadsheetApp.openById('1TbiLnLd5zbzDcFYL7PB8zACjdBi964gbUP1VqbUbQng');
  var sheet = spreadsheet.getSheetByName(lab);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(lab);
    sheet.appendRow(["配属年度", "日時", "薬科学科", "薬学科", "メールアドレス"]);
  }

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === userEmail) {
      return "エラー: あなたはすでにデータを送信済みです。";
    }
  }

  var rowData = [year, new Date(), "", "", userEmail];
  if (gakka === "薬科学科") {
    rowData[2] = gpa;
  } else if (gakka === "薬学科") {
    rowData[3] = gpa;
  }

  sheet.appendRow(rowData);
  return "データを保存しました！";
}
