const compensatoryDayOffSheet = SpreadsheetApp.openById(COMPENSATORY_DAY_OFF_SHEET_ID);
const applyFormSheet = compensatoryDayOffSheet.getSheetByName("振替休日申請書");

/**
 * 振替休日とその原因の日を二次元配列をプロパティに持つオブジェクトで返却する
 */
const searchCompensationDaysOff = () => {
  // 呼出時のDateオブジェクトを取得し（毎月１日実行）、前月に設定する
  const dateObj = new Date();
  dateObj.setMonth(dateObj.getMonth() - 1);

  // 勤務表から前月の休日区分を取得し、振替休日発生原因の勤務日と紐付ける
  const workSchedule = SpreadsheetApp.openById(SPREADSHEET_ID);
  const latestWorkScheduleSheet = workSchedule.getSheets().splice(-1)[0];
  const searchTarget = latestWorkScheduleSheet.getRange(10, 2 , 31, 10);
  const filteredTarget = searchTarget.getValues().filter(values => values[2] === COMPENSATORY_DAY_OFF);
  const origin = [];
  const destination = [];
  for (let i = 0; i < filteredTarget.length; i++) {
    origin.push([filteredTarget[i][9].split(" ")[0].split("の")[0]]);
    destination.push([`${dateObj.getMonth() + 1}/${filteredTarget[i][0]}`]);
  }
  const result = {
    origin,
    destination,
  }
  return result;
}

/**
 * result: {origin: [][], destination: [][]} 
 * @return: ファイル名
 */
const updateApplyForm = () => {
  const result = searchCompensationDaysOff();
  const origin = applyFormSheet.getRange(11, 4, result.origin.length, 1);
  const destination = applyFormSheet.getRange(11, 14, result.destination.length, 1);
  const formCreatedDate = applyFormSheet.getRange(4, 15);
  const rangeList = applyFormSheet.getRangeList(["D11:D15", "N11:N15"]);
  rangeList.clearContent().setNumberFormat("yyyy年mm月dd日");
  const today = new Date();
  today.setMonth(today.getMonth() - 1);
  formCreatedDate.setValue(today);
  origin.setValues(result.origin);
  destination.setValues(result.destination);
  return today;
}

const main = () => {
  console.time("start");
  let fileName = updateApplyForm();
  fileName = `${PREFIX}${fileName.getFullYear()}年${fileName.getMonth() + 1}月`;
  Utilities.sleep(300000);

  const token = ScriptApp.getOAuthToken();
  const url = "https://docs.google.com/spreadsheets/d/" + COMPENSATORY_DAY_OFF_SHEET_ID + "/export?format=xlsx";
  const file = UrlFetchApp.fetch(url, {headers: {'Authorization': 'Bearer ' + token}}).getBlob().setName(fileName); 
  DriveApp.getFolderById("1C62boj9_5wPaWYN6QurS1ZlYOlRRUXTi").createFile(file);
  MailApp.sendEmail(TO, SUBJECT, MESSAGE, {attachments: file});
  console.timeEnd("start");
}
