var SECRET_KEY = "hokyak";
var PRIVATE_SHEET_ID = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"; //非公開シート
var PUBLIC_SHEET_ID = "1TbiLnLd5zbzDcFYL7PB8zACjdBi964gbUP1VqbUbQng"; // 公開シート

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index").setTitle("北薬データベース");
}

function validateEmail(email) {
  var allowedDomain = "@elms.hokudai.ac.jp";
  return email.endsWith(allowedDomain);
}

function saveData(year, lab, gakka, gpa, password, email) {
  console.log("受け取ったパスワード: " + password);

  if (password !== SECRET_KEY) {
    return "エラー: 合言葉が違います。";
  }

  if (!validateEmail(email)) {
    return "エラー: 学内メールアドレスを入力してください。";
  }

  var privateSheet = SpreadsheetApp.openById(PRIVATE_SHEET_ID).getSheetByName("emails");
  var publicSheet = SpreadsheetApp.openById(PUBLIC_SHEET_ID).getSheetByName(lab);

  if (!privateSheet) {
    privateSheet = SpreadsheetApp.openById(PRIVATE_SHEET_ID).insertSheet("emails");
    privateSheet.appendRow(["日時", "メールアドレス"]);
  }

  if (!publicSheet) {
    publicSheet = SpreadsheetApp.openById(PUBLIC_SHEET_ID).insertSheet(lab);
    publicSheet.appendRow(["配属年度", "日時", "薬科学科", "薬学科"]);
  }

  var privateData = privateSheet.getDataRange().getValues();

  for (var i = 1; i < privateData.length; i++) {
    if (privateData[i][1] === email) {
      return "エラー: あなたはすでにデータを送信済みです。";
    }
  }

  var timestamp = new Date();

  privateSheet.appendRow([timestamp, email]);

  var publicData = [year, timestamp, "", ""]; 

  if (gakka === "薬科学科") {
    publicData[2] = gpa;
  } else if (gakka === "薬学科") {
    publicData[3] = gpa;
  }

  publicSheet.appendRow(publicData);

  return "データを保存しました！";
}
