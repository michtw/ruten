function indent()
{
    var s = SpreadsheetApp.getActiveSheet(); // Get spreadsheet name 
    var cell = s.getActiveCell(); // store active cell name in current spreadsheet
    var value = cell.getValue();
    var formula = '=CONCAT(REPT( CHAR( 160 ), 2), "' + value + '")';
  
    cell.setFormula([formula]);
    cell.setValue(cell.getValue());
}
  
function getFirstEmptyRow()
{
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    var column = spr.getRange('A:A');
    // Optimizing Spreadsheet Operations
    // http://googleappsscript.blogspot.tw/2010/06/optimizing-spreadsheet-operations.html
    var values = column.getValues(); // get all data in one call
    var ct = 0;
    while ( values[ct][0] != "" )
        ct++;
  
    return (ct);
}

function splitNum(str) 
{
   var num = str.match(/(\d{4}).*(\d{1,})/);
   return num;
}

function getCost2(title) 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var costSheet = ss.getSheetByName("成本資料庫");
  
  var range_cost = costSheet.getRange("C:C");
  var values_cost = range_cost.getValues();
  
  var range_mode = costSheet.getRange("B1:B288");
  var modes = range_mode.getValues();
  
  var i;
  for (i = 0; i < modes.length; i++) {
    var regex = new RegExp('\\b' + modes[i][0] + '\\b');
    if (title == modes[i][0]) {  
      return values_cost[i][0];
    }
  }
  
  if (i >= modes.length) {
    Browser.msgBox("\n\n \"" + title + "\"  找不到相對應的成本。\n\n");
    Logger.log("錯誤: \"" + title + "\" 找不到相對應的成本。");
  }
  return null;
}

function getCode(title)
{  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var costSheet = ss.getSheetByName("成本資料庫");
  
  var range_mode = costSheet.getRange("B1:B288");
  var modes = range_mode.getValues();
  
  var range_code = costSheet.getRange("A1:A288");
  var codes = range_code.getValues();

  for (var i = 0; i < modes.length; i++) {
    var regex = new RegExp('\\b' + modes[i][0] + '\\b');
    if (title == modes[i][0]) {  
      Logger.log(title);
      Logger.log("i: " + i + " code: " + codes[i][0]);
      return codes[i][0];
    }
  }
  return null;
}

function parseValues(str)
{
    var phoneNum;
    var name;
    var address;
    var tmp;
    var idx;
    var auction = {};    
  
    // var auction = [["項次", "日期", "型號", "數量", "金額", "淨利", "姓名", "電話", "地址"]];
    // var i = str.replace(/(?:\r\n|\r|\n)/g, '#');
  
    if (str.indexOf("露天拍賣") < 0 && str.indexOf("露天拍賣LOGO") < 0 && str.indexOf("露天拍賣Logo") < 0) {
        Logger.log("Cannot find the key word. \"露天拍賣\"");            
        return 'undefined';
    }

    var res = str.split("投訴");
   
    idx = str.indexOf("收件人姓名：");    
    tmp = str.substring(idx, idx + 30);
  
    data = tmp.match(/(：)\s*(.*)[#$]/);
    name = data[2].toString();    
      
    idx = str.indexOf("手機：", idx);   // search index of '手機：'   Start from "收件人姓名：". 
    tmp = str.substring(idx, idx + 50);
    data = tmp.match(/(：)\s*(.*)[#$]/);
    phoneNum = data[2].toString();
  
    idx = str.indexOf("收件地址：");
    // TODO 60個字的長度應該可以滿足大部分的需求
    tmp = str.substring(idx, idx + 60);
    data = tmp.match(/(：)\s*(\S+)/);
    address = data[2].toString();
    Logger.log("收件地址: " + address);
  
    var x = 0;
    var i = 1;
 
    // 最後一個 res 資料不是所要的資料。
    while (res[i] && res[i+1]) {
      var item = [];   
      item[0] = 0;  // 流水號
      // 得標時間      
      idx = str.indexOf("投訴");
      
      // pattern:  2015/01/05
      var data = res[i].match(/\d{4}\/\d{1,}\/\d{1,}/);
      var acceptedTime = data[0];
      item[1] = acceptedTime;
      Logger.log("得標時間: " + acceptedTime); 
      
      idx = res[i].indexOf(acceptedTime);
      
      var a = res[i].substring(idx, idx + 100);
      var model = a.split(/\(\d+\)/);
      var title = model[0].split(/\d{4}\/\d{2}\/\d{2}/)
      var cost = getCost2(title[1].trim());
      if (cost === null) {
        Loggger.log("Cannot find cost of: " + title[1]);
        return;
      }
      
      var code = getCode(title[1].trim());
      if (code === null) {
        Loggger.log("Cannot find code of: " + title[1]);
        return;
      }
      
      var data = a.match(/(\S+)/g);
      item[2] = data[2].toString();
      Logger.log("型號: " + item[2]); 
      item[3] = 0;
      
      data = res[i].match(/(\([0-9]{8,}\))\s+(\d+)/);
      var count = data[2].toString();
      item[3] = (item[3] == 0) ? count : item[3];
      Logger.log("數量: " + item[3]); 
      
      data = res[i].match(/(\$\d{2,})/);                            
      count = data[0].toString();
      item[4] = count.replace(/\$/g, "");
      Logger.log("商品總價: " + item[4]);
      
      item[5] = item[4] - (cost * item[3]);
      Logger.log("淨利: " + item[5]);
      
      item[6] = name;
      item[7] = phoneNum;
      item[8] = address;
      item[9] = code;      
      
      auction[x] = item;
      Logger.log("淨利-" + x + ": " + auction[x][5])
      ++i;
      ++x;
    }  // end while
  
    return auction;
}

function onEdit()
{  
    var s = SpreadsheetApp.getActiveSheet(); // Get spreadsheet name 
    var sheetName = s.getName();
  
    if ((sheetName != "2015銷售表") && (sheetName != "2016銷售表") && (sheetName != "銷售表"))
         return 'undefined';
  
    var ac = s.getActiveCell(); // store active cell name in current spreadsheet
    var rawStr = ac.getValue();
    var str = rawStr.replace(/(\r\n|\n|\r)/gm, " # ");
    var row = getFirstEmptyRow();
  
    // 不管目前的位置在何處，借由設定 setActiveCell() 讓最後一筆資料可以大約在螢幕好方便檢視。
    // (一開始的就要把 active cell 拉到螢幕中央才有辦法達到。)
    if (row > 10)
        s.setActiveCell(s.getDataRange().offset(row - 10, 0, 1, 1));

    var data = parseValues(str);    
       
    if (data === 'undefined') {
        Logger.log("ERROR: 文件解析錯誤。");
        return;
    }    
  
    var idxVal = s.getRange(row, 1).getValue() + 1;   
    var x = 0

    while (data[x])
        data[x++][0] = idxVal++;
    
    // Google spreadsheet row number is start from 1 not 0.
    row++;
  
    //     var auction = [["項次", "日期", "型號", "數量", "金額", "淨利", "姓名", "電話", "地址", "代碼"]];
    x = 0
    while (data[x]) {
        for (var i = 0; i < 10; i++) 
            s.getRange(row, i+1).setValue(data[x][i]);       
        row++;
        x++;
    }       
  
    // 清除cell 中複製的內容
    ac.clear();
}

function readRows()
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    //Logger.log(row);
  }
}

function gotoLastRow()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("2016銷售表");
  var lastrow = sheet.getLastRow();
  sheet.setActiveCell(sheet.getDataRange().offset(lastrow + 25, 0, 1, 1));
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() 
{ 
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();       
  var sheet = spreadsheet.getSheetByName("2015銷售表");
  var range_profit = sheet.getRange("F:F");
  var values_profit = range_profit.getValues();
  
  var range_date = sheet.getRange("B:B");
  var values_date = range_date.getValues();
  
  var range_model = sheet.getRange("C:C");
  var values_model = range_model.getValues();
}

