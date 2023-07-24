var SpreadSheetID = "1cQAmwN-naQHs4BkgDM0Pu8DXxT3cVonDW-R7d1eMKNY"
var SheetName = "Copy of Time Point/Stability"


function projectTracker() {
  var ss = SpreadsheetApp.openById(SpreadSheetID);
  var time_point = ss.getSheetByName(SheetName);
  var data = getData(time_point)

  var series = "";
  var dates = [];

  // goes through rows
  for (let row = 0; row < data.length; row++){
    // selecting rows with series and date information
    if (data[row]['series'] != ""){
      dates = [];
      series = data[row]['series'];
      client = data[row]['client'];
      product_type = data[row]['product type'];

      // going through columns for months 1-24
      // if there is a date in the row add it to a variable to use for selecting a column
      // have date for 1 mo, 2 mo...24 mo
      for (let col = 0; col < 25; col++){
        month = col + " mo";
        // if there is a date in a month column add the date to an object: [ [month: date], [month: date], [month: date] ]
        if (data[row][month] != ""){
          dates[month] = data[row][month];

          // one_date = [];
          // one_date[month] = data[row][month];
          // dates.push(one_date);
          // dates.push(data[row][month]);
        }
      }
    } // checking for series if

    // after series and dates in the row have been added to variables, use the info
    if (data[row]['project'] != ""){
      data[row]['series'] = series;
      data[row]['client'] = client;
      data[row]['product type'] = product_type;

      // in this part use the variables to assign the date to the corresponding month
      // if column has an x in 1 mo, add date from above
      // if column has an x in 3 mo, add date from above
      for (let col = 0; col < 25; col++){ 
        month = col + " mo";
        // if there is a date in a month column add the date to an object: [ [month: date], [month: date], [month: date] ]
        // console.log(`series: ${data[row]["series"]} month: ${data[row][month]}`);
        if (data[row][month] == "x" || data[row][month] == "X"){
          // nice still has access to dates here
          // console.log(`this x (${data[row][month]}) is replaced by a date (${dates[month]})`);
          data[row][month] = dates[month];
        }
      }
    } // checking for project if

  } // rows for loop

  const now = new Date();
  const MILLS_PER_DAY = 1000 * 60 * 60 * 24;
  var plus_two_weeks = new Date(now.getTime() + 13*MILLS_PER_DAY);
  var plus_month = new Date(now.getTime() + 30*MILLS_PER_DAY);

  var two_remind = [];
  var month_remind = [];

  // one column for due dates, UPDATE: only if data[row]['project'] has a value (to exclude the row with series and date information)
  for (let row = 0; row < data.length; row++){
    if (data[row]['project'] != ""){
      // console.log(`dates to add to due column: ${Object.keys(data[row])}`);
      for (key of Object.keys(data[row])){
        if (data[row][key] instanceof Date){
          data[row]['due'] = data[row][key];
          // console.log(`this date ${data[row][key]} is added to due date column`)
        }
      }
    }
  }

  for (let row = 0; row < data.length; row++){
    if (data[row]['due'] <= plus_two_weeks){
      two_remind.push(data[row]);
    }
    if (data[row]['due'] <= plus_month){
      month_remind.push(data[row]);
    }
  }

  // console.log(`2 weeks: ${plus_two_weeks} || month: ${plus_month}`);
  // for (let row = 0; row < two_remind.length; row++){
  //   console.log(two_remind[row]['project'], two_remind[row]['due'], two_remind[row]['product type']);
  // }

  // console.log(`2 weeks: ${plus_two_weeks} || month: ${plus_month}`);
  // for (let row = 0; row < month_remind.length; row++){
  //   console.log(month_remind[row]['project'], month_remind[row]['due']);
  // }


  var reminders = [];
  reminders["Two Week"] = two_remind
  reminders["Month"] = month_remind;

  for (r of Object.keys(reminders)){
    console.log(r);
    console.log(reminders[r]);
    if (reminders[r].length != 0){
      MailApp.sendEmail({to: "EMAIL",
                          subject: `TOX-011 Project ${r} Reminders`,
                          htmlBody: printStuff(reminders[r]),
                          noReply:true});
    }
    else{
      MailApp.sendEmail({to: "EMAIL",
                          subject: `No TOX-011 Projects in the next ${r}`,
                          htmlBody: "",
                          noReply:true});
    }
  }
}

function printStuff(reminders){
  string = "<html><body><br><table border=1><tr><th>Series</th><th>Project</th><th>Client</th><th>Product Type</th><th>Date</th></tr></br>";
  for (var i=0; i<reminders.length; i++){
    string = string + "<tr>";

    temp = `<td> ${reminders[i]['series']} </td><td> ${reminders[i]['project']}  </td><td> ${reminders[i]['client']} </td><td> ${reminders[i]['product type']}</td><td> ${Utilities.formatDate(reminders[i]['due'], 'America/New_York', 'MMMM dd, yyyy')}</td>`;

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
