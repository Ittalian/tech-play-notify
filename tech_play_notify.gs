const SHEET_URL = PropertiesService.getScriptProperties().getProperty('SHEET_URL');

function confirmInformation() {
  const threads = getThreads();
  var studyInfoList = [];

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (message of messages) {
      studyInfoList.push(...getStudyInfo(message.getPlainBody()));
    }
  }
  classifyStudyInfo(studyInfoList);
}

// プライベートメソッド
function getThreads() {
  var date = new Date();
  date.setDate(date.getDate() - 1);
  const yesterday = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  const query = 'from:noreply@techplay.jp subject:新着イベント ' + 'after:' + yesterday;
  const threads = GmailApp.search(query);

  return threads;
}

function getStudyInfo(message) {
  var reg = /\n.+ \n<.+>\n \n[0-9]{4}\/[0-9]{2}\/[0-9]{2} \(.\) [0-9]+:[0-9]{2} 開催/g;
  var studyInfo = message.match(reg);
  return studyInfo;
}

function classifyStudyInfo(studyInfoList) {
  for (const stufyInfo of studyInfoList) {
    const infoList = stufyInfo.split('\n');
    const title = infoList[1];
    const date = infoList[2];
    const url = infoList[4];
    setSheet(title, 1);
    setSheet(date, 2);
    setSheet(url, 3);
  }
}

function setSheet(value, col) {
  const spreadSheet = SpreadsheetApp.openByUrl(SHEET_URL);
  const firstSheet = spreadSheet.getSheets()[0];
  const lastRow = firstSheet.getLastRow();
  firstSheet.getRange(lastRow + 1, col).setValue(value);
}
