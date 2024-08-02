const SHEET_URL = PropertiesService.getScriptProperties().getProperty('SHEET_URL');
const SHEET = SpreadsheetApp.openByUrl(SHEET_URL);
const FIRSTSHEET = SHEET.getSheets()[0];

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
  for (let i = 0; i < studyInfoList.length; i++) {
    const infoList = studyInfoList[i].split('\n');
    const title = infoList[1];
    const url = infoList[2].substring(1, infoList[2].length - 1);
    const date = infoList[4];
    setSheet(title, i + 2, 1);
    setSheet(url, i + 2, 2);
    setSheet(date, i + 2, 3);
  }
}

function setSheet(value, row, col) {
  FIRSTSHEET.getRange(row, col).setValue(value);
}
