var ss = SpreadsheetApp.getActive()
var dSheet = ss.getSheetByName("Data")
var attendRecord = ss.getSheetByName("Attendance Record")

var date = new Date()

// function doGet(e) {

  // var sheet = ss.getSheetByName("Daily Attendance")

//   return getUsers(sheet)
// }


function getUsers() {
  var sheet = ss.getSheetByName("Daily Attendance")
  var jo = {}
  var dataArray = []
  var rows = sheet.getRange(3, 1, sheet.getLastRow() - 2, 15).getValues().filter(f => f[0] != "")
  // Logger.log(rows)

  for (var i = 0, l = rows.length; i < l; i++) {

    var dataRow = rows[i];
    var record = {}
    record['SNo.'] = dataRow[0]
    record['Name'] = dataRow[1]
    record['Intime'] = dataRow[2]
    record['Outtime'] = dataRow[3]
    record['Status'] = dataRow[4]
    record['OT'] = dataRow[5]
    record['Before5'] = dataRow[6]
    record['Before7'] = dataRow[7]
    record['Before8'] = dataRow[8]
    record['b8_10'] = dataRow[9]
    record['b10_12'] = dataRow[10]
    record['b12_2'] = dataRow[11]
    record['b2_5'] = dataRow[12]
    record['Sign'] = dataRow[13]
    record['VehicleKM'] = dataRow[14]

    dataArray.push(record)

  }
  jo.user = dataArray;

  var result = JSON.stringify(jo);

  // Logger.log(result)

  // return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
  // var dSheet = ss.getSheetByName("Data")

  dSheet.getRange(dSheet.getLastRow() + 1, 1, 1, 1).setValue(date)
  dSheet.getRange(dSheet.getLastRow(), 2, 1, 1).setValue(result)
}


function clear(){
attendRecord.getRange(3,2,attendRecord.getLastRow()-2,14).clearContent()
}


function recall() {
  clear()
var aDate = attendRecord.getRange('O1').getValue().getDate()
// var js = dSheet.getRange('A2').getValue().getDate()
var jsData = dSheet.getRange(2,1,dSheet.getLastRow(),2).getValues().filter(f=>f[0]!="")

// Logger.log(jsData[1][0])

var js 

for(i=0;i<jsData.length;i++){
  if(jsData[i][0].getDate()  ==  aDate ){
        js = jsData[i][1]
      break;
  }
}
// Logger.log(js)
// Logger.log(jsData[i][1])

const obj = JSON.parse(js)
  Logger.log(obj.user.length)
  var l = obj.user.length
  var dArray = []
for(i=0;i<l;i++){
  const { Name } = obj.user[i]
  const { Intime } = obj.user[i]
  const { Outtime } = obj.user[i]
  const { Status } = obj.user[i]
  const { OT } = obj.user[i]
  const { Before5 } = obj.user[i]
  const { Before7 } = obj.user[i]
  const { Before8 } = obj.user[i]
  const { b8_10 } = obj.user[i]
  const { b10_12 } = obj.user[i]
  const { b12_2 } = obj.user[i]
  const { b2_5 } = obj.user[i]
  const { Sign } = obj.user[i]
  const { VehicleKM } = obj.user[i]

  var array = [
    Name,
    Intime,
    Outtime,
    Status,
    OT,
    Before5,
    Before7,
    Before8,
    b8_10,
    b10_12,
    b12_2,
    b2_5,
    Sign,
    VehicleKM,
  ]
dArray.push(array)
// Logger.log(array)
}
// Logger.log(dArray)
attendRecord.getRange(3,2,l,14).setValues(dArray)   
}



