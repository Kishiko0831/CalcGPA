var PRIVATE_SHEET_ID = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
var PUBLIC_SHEET_ID = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index").setTitle("北薬データベース");
}

function validateEmail(email) {
  return email.endsWith("@xxxxxxxxxxx.com");
}

function saveData(year, lab, gakka, gpa, email) {
  console.log("受け取ったメール: " + email);

  if (!validateEmail(email)) {
    return "エラー: 学内メールアドレスを入力してください。";
  }

  var privateSheet = SpreadsheetApp.openById(PRIVATE_SHEET_ID).getSheetByName("emails");
  var publicSpreadsheet = SpreadsheetApp.openById(PUBLIC_SHEET_ID);
  var publicSheet = publicSpreadsheet.getSheetByName(lab);
  var summarySheet = publicSpreadsheet.getSheetByName("Summary");

  if (!privateSheet) {
    privateSheet = SpreadsheetApp.openById(PRIVATE_SHEET_ID).insertSheet("emails");
    privateSheet.appendRow(["日時", "メールアドレス"]);
  }
  if (!publicSheet) {
    publicSheet = publicSpreadsheet.insertSheet(lab);
    publicSheet.appendRow(["配属年度", "日時", "薬科学科", "薬学科"]);
  }
  if (!summarySheet) {
    summarySheet = publicSpreadsheet.insertSheet("Summary");
    summarySheet.appendRow(["配属年度", "薬科学科 平均", "薬科学科 標準偏差", "薬学科 平均", "薬学科 標準偏差"]);
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
    if (gakka === "薬科学科") {
      publicData[2] = parseFloat(gpa);
    } else if (gakka === "薬学科") {
      publicData[3] = parseFloat(gpa);
    }
  }

  publicSheet.appendRow(publicData);

  var formulaCell = publicSheet.getRange("F1").getFormula();
  if (!formulaCell || formulaCell === "") {
    publicSheet.getRange("F1").setFormula('=QUERY(A:D, "SELECT A, AVG(C), AVG(D) WHERE A IS NOT NULL GROUP BY A ORDER BY A DESC", 1)');
  }

  updateSummarySheet();

  return "データを保存しました！";
}

function updateSummarySheet() {
  var publicSpreadsheet = SpreadsheetApp.openById(PUBLIC_SHEET_ID);
  var sheets = publicSpreadsheet.getSheets();
  var dataMap = {};

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.getName() !== "Summary") {
      var data = sheet.getDataRange().getValues();
      for (var j = 1; j < data.length; j++) {
        var year = data[j][0];
        if (!year) continue;

        if (!dataMap[year]) {
          dataMap[year] = { gpaSci: [], gpaPharm: [] };
        }

        if (data[j][2]) dataMap[year].gpaSci.push(parseFloat(data[j][2]));
        if (data[j][3]) dataMap[year].gpaPharm.push(parseFloat(data[j][3]));
      }
    }
  }

  var summarySheet = publicSpreadsheet.getSheetByName("Summary");
  summarySheet.clear();
  summarySheet.appendRow(["配属年度", "薬科学科 平均", "薬科学科 標準偏差", "薬学科 平均", "薬学科 標準偏差"]);

  var sortedYears = Object.keys(dataMap).sort((a, b) => b - a);

  sortedYears.forEach(year => {
    var sciGpa = dataMap[year].gpaSci;
    var pharmGpa = dataMap[year].gpaPharm;
    
    var sciAvg = sciGpa.length > 0 ? average(sciGpa) : "";
    var sciStd = sciGpa.length > 0 ? standardDeviation(sciGpa) : "";
    var pharmAvg = pharmGpa.length > 0 ? average(pharmGpa) : "";
    var pharmStd = pharmGpa.length > 0 ? standardDeviation(pharmGpa) : "";

    summarySheet.appendRow([year, sciAvg, sciStd, pharmAvg, pharmStd]);
  });
}

function average(arr) {
  var sum = arr.reduce((a, b) => a + b, 0);
  return sum / arr.length;
}

function standardDeviation(arr) {
  var avg = average(arr);
  var squareDiffs = arr.map(value => Math.pow(value - avg, 2));
  var avgSquareDiff = average(squareDiffs);
  return Math.sqrt(avgSquareDiff);
}
