var PRIVATE_SHEET_ID = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
var PUBLIC_SHEET_ID = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index").setTitle("xxxxxxxxxxx");
}

function validateEmail(email) {
  return email.endsWith("@xxxxxxxxxxxxx.com");
}

function saveData(year, lab, gakka, gpa, email) {
  console.log("受け取ったメール: " + email);

  if (!validateEmail(email)) {
    return "エラー: 学内メールアドレスを入力してください。";
  }

  var privateSheet = SpreadsheetApp.openById(PRIVATE_SHEET_ID).getSheetByName("emails");
  var publicSpreadsheet = SpreadsheetApp.openById(PUBLIC_SHEET_ID);
  var publicSheet = publicSpreadsheet.getSheetByName(lab);

  if (!privateSheet) {
    privateSheet = SpreadsheetApp.openById(PRIVATE_SHEET_ID).insertSheet("emails");
    privateSheet.appendRow(["日時", "メールアドレス"]);
  }
  if (!publicSheet) {
    publicSheet = publicSpreadsheet.insertSheet(lab);
    publicSheet.appendRow(["配属年度", "日時", "xxx科", "xxxx科"]);
  }

  var privateData = privateSheet.getDataRange().getValues();
  var emailSet = new Set(privateData.map(row => row[1]));
  if (emailSet.has(email)) {
    return "エラー: あなたはすでにデータを送信済みです。";
  }

  var timestamp = new Date();
  privateSheet.appendRow([timestamp, email]);

  var publicData = [year, timestamp, "", ""];
  if (!isNaN(parseFloat(gpa))) {
    if (gakka === "xxx科") {
      publicData[2] = parseFloat(gpa);
    } else if (gakka === "xxxx科") {
      publicData[3] = parseFloat(gpa);
    }
  }

  publicSheet.appendRow(publicData);

  return "データを保存しました！";
}
