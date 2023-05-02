var attendsheet = ss.getSheetByName("Daily Attendance")

var intime = 'Sat Dec 30 1899 05:00:00 GMT+0521 (India Standard Time)'
var intime1 = 'Sat Dec 30 1899 07:00:00 GMT+0521 (India Standard Time)'
var intimeend = 'Sat Dec 30 1899 08:00:00 GMT+0521 (India Standard Time)'

var datetime1 = new Date(intime);
var datetime2 = new Date(intime1);
var datetime3 = new Date(intimeend);

var outtime = 'Sat Dec 30 1899 20:00:00 GMT+0521 (India Standard Time)'
var outtime1 = 'Sat Dec 30 1899 22:00:00 GMT+0521 (India Standard Time)'
var outtime2 = 'Sat Dec 30 1899 24:00:00 GMT+0521 (India Standard Time)'
var outtime3 = 'Sat Dec 30 1899 00:00:00 GMT+0521 (India Standard Time)'
var outtime4 = 'Sat Dec 30 1899 02:00:00 GMT+0521 (India Standard Time)'
var outtime5 = 'Sat Dec 30 1899 05:00:00 GMT+0521 (India Standard Time)'

var datetimeout = new Date(outtime);
var datetimeout1 = new Date(outtime1);
var datetimeout2 = new Date(outtime2);
var datetimeout3 = new Date(outtime3);
var datetimeout4 = new Date(outtime4);
var datetimeout5 = new Date(outtime5);

function onEdit(e) {
  var sheetName = "Daily Attendance";
  var sheetName1 = "Attendance Record";
  var range = e.range;
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var value = e.value;
  var oldValue = e.oldValue
  var time = attendsheet.getRange(row, 3).getValue()
  // console.log(time)

if(sheetName === e.source.getActiveSheet().getName() && col === 15 && row === 1){
  statusValues()
}


  var sheetdatetime = new Date(time);
  if (sheetName === e.source.getActiveSheet().getName() && col === 3) {
    if (datetime1 > sheetdatetime) {
      statusValues()
      //  SpreadsheetApp.getUi().alert('Done G'+row);
      SpreadsheetApp.flush();
      attendsheet.getRange(row, 7).setValue('P')
    } else if (datetime2 > sheetdatetime) {
      statusValues()
      //  SpreadsheetApp.getUi().alert('Done H'+row);
      SpreadsheetApp.flush();
      attendsheet.getRange(row, 8).setValue('P')
    } else if (datetime3 > sheetdatetime) {
      statusValues()
      //  SpreadsheetApp.getUi().alert('Done I'+row);
      SpreadsheetApp.flush();
      attendsheet.getRange(row, 9).setValue('P')
    }else{
      statusValues()
    }
  }


  if (sheetName === e.source.getActiveSheet().getName() && col === 4 && e.range.offset(0,-1).getValue() !== '') {
    var time1 = attendsheet.getRange(row, 4).getValue();
    if (time1) {
      time1 = new Date(time1);
      console.log(time1)
      if (time1 > datetimeout && time1 < datetimeout1) {
        // SpreadsheetApp.getUi().alert('Done J'+row);
        SpreadsheetApp.flush();
        attendsheet.getRange(row, 10).setValue('P')
      } else if (time1 > datetimeout1 && time1 < datetimeout2) {
        // SpreadsheetApp.getUi().alert('Done K'+row);
        SpreadsheetApp.flush();
        attendsheet.getRange(row, 11).setValue('P')
      } else if (datetimeout3 < time1 && time1 < datetimeout4) {
        // SpreadsheetApp.getUi().alert('Done L'+row);
        SpreadsheetApp.flush();
        attendsheet.getRange(row, 12).setValue('P')
      } else if(datetimeout5>time1 && datetimeout4<time1){
        // SpreadsheetApp.getUi().alert('Done M'+row);
        SpreadsheetApp.flush();
        attendsheet.getRange(row, 13).setValue('P')
      }
    }
  }

if(sheetName1 === e.source.getActiveSheet().getName() && col === 15 && row === 1 ){
  Logger.log('Run')
 recall()
}

}


function oT() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = attendsheet.getLastRow();
  var valuesC = attendsheet.getRange("C3:C" + lastRow).getValues();
  var valuesE = attendsheet.getRange("E3:E" + lastRow).getValues();
  var range = attendsheet.getRange("F3:F" + lastRow);
  var output = [];
  for (var i = 0; i < valuesC.length; i++) {
    if (valuesC[i][0] == "") {
      output.push([""]);
    } else if (valuesE[i][0] == "R") {
      output.push(["P"]);
    } else {
      output.push([""]);
    }
  }

  range.setValues(output);
}


var settings = ss.getSheetByName("Settings");
var attendstatus = attendsheet.getRange('E3:E').getValues().filter(f => f[0] != "")
const daysArray = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thrusday", "Friday", "Saturday"];
var date = attendsheet.getRange('O1').getValue()

function statusValues() {
  var todayDay = new Date(date).getDay();
  Logger.log(todayDay)
  var values1 = settings.getRange("A2:B").getDisplayValues().filter(f => f[0] != "");
  var values2 = attendsheet.getRange(3, 2, attendsheet.getLastRow() - 2, 1).getDisplayValues();
  var values3 = attendsheet.getRange(3, 3, attendsheet.getLastRow() - 2, 1).getDisplayValues();
  var values4 = attendsheet.getRange(3, 5, attendsheet.getLastRow() - 2, 1).getDisplayValues();
  // Logger.log(values1)
  // Logger.log(values2)

  for (var i = 0; i < values2.length; i++) {
    for (var j = 0; j < values1.length; j++) {


      if (values2[i][0] == values1[j][0] && daysArray.indexOf(values1[j][1] || "Sunday") == todayDay) {
        attendsheet.getRange(i + 3, 5).setValue("R");
        // Logger.log(daysArray.indexOf(values1[j]))
      }
      else
        if (values2[i][0] == values1[j][0] && daysArray.indexOf(values1[j][1]) !== todayDay) {
        attendsheet.getRange(i + 3, 5).setValue("A");
      }
      }
      if(values4[i][0] !='R'&& values3[i][0] !=""   ){
        attendsheet.getRange(i + 3, 5).setValue("P");
      }
      }
      oT()
    }



