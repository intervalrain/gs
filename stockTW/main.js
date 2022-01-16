function execute() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadSheetName = ss.getName();
  let nowSheet = ss.getActiveSheet();
  var sheetName = "TodaySummaryReport";
  
  let newSheet = ss.getSheetByName(sheetName);
  if (newSheet == null){
    addSheet(sheetName);
  } else{
    newSheet.clear();
  }
  var data = nowSheet.getDataRange().getValues();
  
  var headers = ['名次', '股名', '股號', '成交價', '漲跌', '漲跌幅', '最高', '最低', '價差', '成交量(張)', '成交值(億)'];
  for (var i = 0; i < headers.length; i++){
    newSheet.getRange(1, i+1).setValue(headers[i]);
  }
  for (var i = 10; i < data.length; i++){
    var nowRow = (i - 10) / 11 + 2;
    var nowCol = (i - 10) % 11 + 1;
    newSheet.getRange(nowRow, nowCol).setValue(data[i]);
  }
  
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Manual")
    .addItem("Run", "execute").addToUi();
}

function hello(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  var guest = ss.getViewers();
  Logger.log("Hello" + guest);
}

function addSheet(sheetName){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet(sheetName)
}