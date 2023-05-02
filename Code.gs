/** Copyright © Zillion Analytics Pvt. Ltd.
*  his software is specially developed by Zillion Analytics Pvt. Ltd. for use by its clients
*  Unauthorized use & copying of this software will attract penalties.
*  For support contact connect@zillion.io
*/
/**
* The function triggered when installing the add-on.
*/

function onOpen() {
const menuBarUI = SpreadsheetApp.getUi();
menuBarUI.createMenu('AUTOMATION')
.addItem('Update', 'updateSheet')
.addItem('Create Sheet','createSheet')
.addSeparator()
.addItem('Reset Formula','reset')
.addToUi();
}



var ss = SpreadsheetApp.getActive();
const monthArray = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function createSheet(){
const ss = SpreadsheetApp.getActive();
const masterDataSheet = ss.getSheetByName('MasterData');
masterDataSheet.getRange(2,1,masterDataSheet.getLastRow()-1,3).getValues().forEach((el,index)=>{
  if(el[0] && el[1] && !el[2]){
    var templateSheet = DriveApp.getFileById('1ZsZlPUNQ0RsraWxp_ccMCcl5kfsrMQ7AObwMXyDBWBs').makeCopy();
    var sheetUrl = templateSheet.getUrl();
    templateSheet.setName(el[0] + " " + el[1]);
    masterDataSheet.getRange(index+2,3).setValue(sheetUrl);
    var dailySheet = SpreadsheetApp.openByUrl(sheetUrl).getSheetByName("Daily Attendance");
    dailySheet.getRange("C2").setValue(new Date(parseInt(el[1]),monthArray.indexOf(el[0]),1));
  }
})
}

function updateSheet(){
  
  var attendnaceSheet = ss.getSheetByName("Daily Attendance");
  var masterDataSheet = ss.getSheetByName("MasterData");
  var dateValue = attendnaceSheet.getRange("O1").getValue();
  if(dateValue==""){
    SpreadsheetApp.getUi().alert("Date cant be blank!");
    return;
  }
  var date = new Date(dateValue);
  var month = monthArray[date.getMonth()];
  var year = date.getFullYear();
  var url;
  masterDataSheet.getRange("A2:C").getValues().filter(f=>f[0]!="").forEach(r=>{
    if(r[0]==month && r[1]==parseInt(year)){
      url = r[2];
    }
  });
  console.log(month,year,url);
  if(url){
    var ms = SpreadsheetApp.openByUrl(url);
    attendnaceSheet.getRange("B3:O").getValues().filter(f=>f[0]!="" && f[1]!="" && f[2]!="").forEach(r=>{
      var targetSheet = ms.getSheetByName(r[0]);
      if(targetSheet){
        targetSheet.getRange("A4:A34").getValues().forEach((d,i)=>{
          if(new Date(d[0]).toDateString()==date.toDateString()){
            r.shift();
            targetSheet.getRange(i+4,3,1,13).setValues([r]);
            targetSheet.getRange(i+4,3,1,2).setNumberFormat('[h]:mm:ss');
          }
        });
      }
    });
  }else{
    SpreadsheetApp.getUi().alert("Data not submitted! Some error occured.");
  }
    getUsers()
  attendnaceSheet.getRange("C3:O").clearContent();
  attendnaceSheet.getRange("O1").clearContent();
}

/*
function reset(){
  var dailyAttendanceSheet = ss.getSheetByName("Daily Attendance");
  var settingSheet = ss.getSheetByName("Settings");
  settingSheet.getRange("A1").setValue(`=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1ZsZlPUNQ0RsraWxp_ccMCcl5kfsrMQ7AObwMXyDBWBs/edit#gid=1929578592", "Settings!A1:C")`);
  dailyAttendanceSheet.getRange("A2:Q2").setValues([[`=ArrayFormula(if(row(A2:A)=2,"Sl.No.",IF(B2:B="", "", Row(B2:B)-2)))`,
  `=Unique(Settings!A:A)`,
  `In Time`,
  `Out Time`,
  `=ArrayFormula(if(row(A2:A)=2,"Status", IF(B2:B="", "",if(vlookup(B2:B, Settings!A:D, 2, 0)=Text($O$1, "Dddd"),"R",IF(C2:C<>"", "P", IF(C2:C="", "A"))))))`,
  `=ArrayFormula(if(row(A2:A)=2,"OT", IF(C2:C="", "", IF(C2:C<>"", IF(E2:E="R", "P", "")))))`,
  `=ArrayFormula(if(row(A2:A)=2,Settings!E1, if(P2:P="", "", if(P2:P<Settings!$D$1, "P", ""))))`,
  `=ArrayFormula(if(row(A2:A)=2,Settings!E2, IF(P2:P="", "", IF(G2:G="P", "",IF(P2:P<Settings!D2, "P", "")))))`,
  `=ArrayFormula(if(row(A2:A)=2,Settings!E3, IF(P2:P="", "", IF(P2:P>Settings!$D$2,"P", ""))))`,
  `=ArrayFormula(if(row(A2:A)=2,Settings!E4, IF(Q2:Q="", "",  IF(Q2:Q>Settings!$D$5,"",IF(Q2:Q>Settings!$D$4,"P","")))))`,
  `=ArrayFormula(if(row(A2:A)=2,Settings!E5, IF(Q2:Q="", "", IF(J2:J="P", "", IF(Q2:Q>Settings!D5,"P","")))))`,
  `=ArrayFormula(if(row(A2:A)=2,Settings!E6, IF(Q2:Q="", "", IF(J2:J="P", "",IF(K2:K="P", "",IF(Q2:Q<Settings!D7,"P",""))))))`,
  `=ArrayFormula(if(row(A2:A)=2,"2 - 5 AM 600/-", IF(D2:D="", "", IF(J2:J="P", "",IF(K2:K="P", "",IF(L2:L="P", "",IF(Q2:Q>Settings!D7,"P","")))))))`,
  `Sign`,
  `Vehicle KM`,
  '=ARRAYFORMULA(IF(C2:C="", "", TEXT(C2:C, "HH:MM:SS")))',
  `=ARRAYFORMULA(IF(D2:D="", "", TEXT(D2:D, "HH:MM:SS")))`]]);
} */
