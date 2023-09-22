var SpreadSheetID = "1JprPNRcr2hiQKyk1xn0OeZJ9KaTgORTMvvVAStB0a2I"
var SheetName = "Sheet2"
var EmailSheet = "Sheet3"

function projectTracker() {
  var ss = SpreadsheetApp.openById(SpreadSheetID);
  var time_point = ss.getSheetByName(SheetName);
  var data = getData(time_point)

  const now = new Date();
  const MILLS_PER_DAY = 1000 * 60 * 60 * 24;
  var plus_two_weeks = new Date(now.getTime() + 15*MILLS_PER_DAY);
  var plus_month = new Date(now.getTime() + 40*MILLS_PER_DAY);

  var two_remind = [];
  var month_remind = [];

  // one column for due dates, UPDATE: only if data[row]['project'] has a value (to exclude the row with series and date information)
  for (let row = 0; row < data.length; row++){
    // console.log(`dates to add to due column: ${Object.keys(data[row])}`);
    for (key of Object.keys(data[row])){
      if (data[row][key] instanceof Date){
        if (data[row][key] <= plus_two_weeks){
          two_remind.push(data[row]);
        }
        if (data[row][key] <= plus_month){
          month_remind.push(data[row]);
        }
      }
    }
  }

  // one column for due dates, UPDATE: only if data[row]['project'] has a value (to exclude the row with series and date information)
  for (let row = 0; row < data.length; row++){
    if (data[row]['project'] != ""){
      // console.log(`dates to add to due column: ${Object.keys(data[row])}`);
      for (key of Object.keys(data[row])){
        if (data[row][key] instanceof Date){
          data[row]['due'] = data[row][key];
        }
      }
    }
  }

  // delete duplicates
  two_remind = removeDuplicates(two_remind);
  month_remind = removeDuplicates(month_remind);

  // console.log(`2 weeks: ${plus_two_weeks} || month: ${plus_month}`);
  // for (let row = 0; row < two_remind.length; row++){
  //   console.log(two_remind[row]['project'], two_remind[row]['due'], two_remind[row]['product type']);
  // }

  // console.log(`2 weeks: ${plus_two_weeks} || month: ${plus_month}`);
  // for (let row = 0; row < month_remind.length; row++){
  //   console.log(month_remind[row]['project'], month_remind[row]['due']);
  // }

  // email part
  var emailInfo = ss.getSheetByName(EmailSheet);
  var emails = getEmails(emailInfo);

  var reminders = [];
  reminders["Two Week"] = two_remind
  reminders["Month"] = month_remind;

  for (r of Object.keys(reminders)){
    if (reminders[r].length != 0){
      for (var j=0; j<emails.length; j++){
        MailApp.sendEmail({to: emails[j].email,
                          subject: `TOX-011 Project ${r} Reminders`,
                          htmlBody: printStuff(reminders[r]),
                          noReply:true})
      }
    }
    else{
      for (var j=0; j<emails.length; j++){
        MailApp.sendEmail({to: emails[j].email,
                            subject: `No TOX-011 Projects in the next ${r}(s)`,
                            htmlBody: "",
                            noReply:true})
      }

    }
  }

}

function printStuff(reminders){
  string = "<html><body><br><table border=1><tr><th>Series</th><th>Project</th><th>Set</th><th>Client</th><th>Product Type</th></tr></br>";
  for (var i=0; i<reminders.length; i++){
    string = string + "<tr>";

    temp = `<td> ${reminders[i]['series']} </td><td> ${reminders[i]['project']}  </td><td> ${reminders[i]['set']} </td><td> ${reminders[i]['client']} </td><td> ${reminders[i]['product type']}</td>`;

    string = string.concat(temp);
    string = string + "</tr>";
  }
  string = string + "</table></body></html>";
  return string;
}

// https://blog.devgenius.io/send-mass-emails-using-google-apps-script-from-a-google-spreadsheet-fc2f79c9febd
function getData(project_data){
  var dataArray = [];
  // collecting data from 2nd Row , 1st column to last row and last    // column sheet.getLastRow()-1
  var rows = project_data.getRange(3,1,project_data.getLastRow()-2, project_data.getLastColumn()).getValues();

  for(var i = 0, l= rows.length; i<l ; i++){
    var dataRow = rows[i];
    var record = {};
    record['series'] = dataRow[0];
    record['project'] = dataRow[1];
    record['set'] = dataRow[2];
    record['client'] = dataRow[3];
    record['product type'] = dataRow[4];
    record['0 mo'] = dataRow[6];
    record['1 mo'] = dataRow[7];
    record['2 mo'] = dataRow[8];
    record['3 mo'] = dataRow[9];
    record['4 mo'] = dataRow[10];
    record['5 mo'] = dataRow[11];
    record['6 mo'] = dataRow[12];
    record['7 mo'] = dataRow[13];
    record['8 mo'] = dataRow[14];
    record['9 mo'] = dataRow[15];
    record['10 mo'] = dataRow[16];
    record['11 mo'] = dataRow[17];
    record['12 mo'] = dataRow[18];
    record['13 mo'] = dataRow[19];
    record['14 mo'] = dataRow[20];
    record['15 mo'] = dataRow[21];
    record['16 mo'] = dataRow[22];
    record['17 mo'] = dataRow[23];
    record['18 mo'] = dataRow[24];
    record['19 mo'] = dataRow[25];
    record['20 mo'] = dataRow[26];
    record['21 mo'] = dataRow[27];
    record['22 mo'] = dataRow[28];
    record['23 mo'] = dataRow[29];
    record['24 mo'] = dataRow[30];

    dataArray.push(record);
  }
  return dataArray;
}

function getEmails(email_sheet){
  var dataArray = [];
  // collecting data from 2nd Row , 1st column to last row and last    // column sheet.getLastRow()-1
  var rows = email_sheet.getRange(2,1,email_sheet.getLastRow()-1, email_sheet.getLastColumn()).getValues();
  for(var i = 0, l= rows.length; i<l ; i++){
    var dataRow = rows[i];
    var record = {};
    record['email'] = dataRow[0];
    dataArray.push(record);
  }
  return dataArray;
}

function removeDuplicates(with_duplicates) {
  jsonObject = with_duplicates.map(JSON.stringify);
  uniqueSet = new Set(jsonObject);
  uniqueArray = Array.from(uniqueSet).map(JSON.parse);

  return uniqueArray;
}
