function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui=SpreadsheetApp.getUi();
  ui.createMenu('Admin')
      .addItem('Generate Timesheet','genTimeSheet')
      .addItem('Run Monthly Routine', 'monthly')
      .addSeparator()
      .addItem('Get Remaining Email Quota','getQuota')
      .addToUi();
}

function monthly(){
  var today = new Date();
  var formattedName = Utilities.formatDate(today, "GMT+8", "yyyy MMM")
  var generatedSheet = genTimeSheet(formattedName)
  var reportid = makepdf(generatedSheet.getSheetId(), "[RA Timesheet] " + formattedName + '.pdf')
  
  sendMail(reportid)
  
  generatedSheet.hideSheet()
}

function genTimeSheet(tsname="", sendMail=false){
  var ss = SpreadsheetApp.openById("1QM-abh0F7FVU-owd1MgBEAjX_3Iih-qOzB2TXGVi8d0")
  var master = ss.getSheetByName("master")
  var settings = ss.getSheetByName("Settings")
  var ui=SpreadsheetApp.getUi();
  var d = new Date();
  var maxHours = settings.getRange("B5").getValue()
    
  // Set Timesheet name
  if (tsname=="") {
    tsname = ui.prompt(
      "Timesheet Generator",
      "Please enter an identifier",
      ui.ButtonSet.OK_CANCEL
    ).getResponseText()
  }
  
  //copy from template
  var report = ss.insertSheet({template: ss.getSheetByName("template")});
  report.setName(tsname).showSheet()
  
  var data = master.getDataRange().getValues()
  var out = knapsack_wrapper(data, maxHours)
  
  report.getRange("B6").setHorizontalAlignment("left").setValue(Utilities.formatDate(d, "GMT+8", "MMMM, yyyy"))
  report.getRange("E32").setHorizontalAlignment("center").setValue(Utilities.formatDate(d, "GMT+8", "d MMM yyyy"))
  
  for(var i = 0; i < out.arr.length; i++){
    report.getRange(9+i, 1).setValue(data[out.arr[i]][0])
    report.getRange(9+i, 2).setValue(data[out.arr[i]][1])
    report.getRange(9+i, 6).setValue(data[out.arr[i]][2])
    master.getRange(out.arr[i]+1, 4).setValue(tsname)
  }
  
  //set last edit
  settings.getRange("F5").setValue(d);
  
  //edit contracts sheet
  var contracts = ss.getSheetByName("contracts")
  var last = contracts.getLastRow()
  contracts.getRange(last+1, 1).setRichTextValue(
    SpreadsheetApp.newRichTextValue()
      .setText(tsname)
      .setLinkUrl('#gid=' + report.getSheetId())
      .build()
  )
  contracts.getRange(last+1, 2).setValue(out.arr.length) 
  contracts.getRange(last+1, 3).setValue(out.total) 
  contracts.getRange(last+1, 4).setValue(maxHours) 
  contracts.getRange(last+1, 5).setValue(new Date()) 
  
  return report
}

function makepdf(sheetid, filename="Unamed Export.pdf"){
  SpreadsheetApp.flush()
  var url = 'https://docs.google.com/spreadsheets/d/' 
    + '1QM-abh0F7FVU-owd1MgBEAjX_3Iih-qOzB2TXGVi8d0'
    + '/export?exportFormat=pdf&format=pdf'
    + '&size=A4'
    + '&portrait=true'
    + '&fitw=true'                 
    + '&sheetnames=false&printtitle=false'
    + '&horizontal_alignment=CENTER'
    + '&vertical_alignment=TOP'
    + '&pagenum=false'
    + '&gridlines=false'
    + '&fzr=FALSE'      
    + '&gid='
    + sheetid.toString()
  
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  var response = UrlFetchApp.fetch(url, params).getBlob();
  var file = DriveApp.createFile(response).setName(filename); 
  
  //Move File to RA folder
  file.moveTo(DriveApp.getFolderById("1W-8ZlW56z5HWzgsm9KdTs8UematzDP7O"))
  
  return file.getId()
}

function knapsack_wrapper(data, limit=37){
  
  var items_list = []
  
  for(var i = data.length-1; i >=1; i--){
    if (data[i][3]) continue
    items_list[items_list.length] = {w: data[i][2], v: data[i][2], index: i}
  }
  
  Logger.log(items_list)
  var output = knapsack(items_list, limit)
  
  Logger.log(output)
  var ans = []
  
  for (i=0; i < output.subset.length; i++){
    ans[i] = Number(output.subset[i].index)
  }
  
  return {total: output.maxValue, arr: ans.sort((a,b) => a-b)}
}

function knapsack(items, capacity){
  // This implementation uses dynamic programming.
  // Variable 'memo' is a grid(2-dimentional array) to store optimal solution for sub-problems,
  // which will be later used as the code execution goes on.
  // This is called memoization in programming.
  // The cell will store best solution objects for different capacities and selectable items.
  var memo = []

  // Filling the sub-problem solutions grid.
  for (var i = 0; i < items.length; i++) {
    var row = []
    for (var cap = 1; cap <= capacity; cap++) {
      row.push(getSolution(i,cap))
    }
    memo.push(row)
  }

  // The right-bottom-corner cell of the grid contains the final solution for the whole problem.
  return(getLast());

  function getLast(){
    var lastRow = memo[memo.length - 1];
    return lastRow[lastRow.length - 1];
  }

  function getSolution(row,cap){
    const NO_SOLUTION = {maxValue:0, subset:[]};
    // the column number starts from zero.
    var col = cap - 1;
    var lastItem = items[row];
    // The remaining capacity for the sub-problem to solve.
    var remaining = cap - lastItem.w;

    // Refer to the last solution for this capacity,
    // which is in the cell of the previous row with the same column
    var lastSolution = row > 0 ? memo[row - 1][col] || NO_SOLUTION : NO_SOLUTION;
    // Refer to the last solution for the remaining capacity,
    // which is in the cell of the previous row with the corresponding column
    var lastSubSolution = row > 0 ? memo[row - 1][remaining - 1] || NO_SOLUTION : NO_SOLUTION;

    // If any one of the items weights greater than the 'cap', return the last solution
    if(remaining < 0){
      return lastSolution;
    }

    // Compare the current best solution for the sub-problem with a specific capacity
    // to a new solution trial with the lastItem(new item) added
    var lastValue = lastSolution.maxValue;
    var lastSubValue = lastSubSolution.maxValue;

    var newValue = lastSubValue + lastItem.v;
    if(newValue >= lastValue){
      // copy the subset of the last sub-problem solution
      var _lastSubSet = lastSubSolution.subset.slice();
      _lastSubSet.push(lastItem);
      return {maxValue: newValue, subset:_lastSubSet};
    }else{
      return lastSolution;
    }
  }
}

function sendMail(fileid){
  var file = DriveApp.getFileById(fileid)
  var settings = SpreadsheetApp.openById("1QM-abh0F7FVU-owd1MgBEAjX_3Iih-qOzB2TXGVi8d0").getSheetByName("Settings")
  var recipient = settings.getRange("B6").getValue()
  var subject = '[' + Utilities.formatDate(new Date(), "GMT+8", "yyyy MMM") + '] ' + settings.getRange("B7").getValue()
  var body = settings.getRange("B8").getValue()+ '\n\n' +
    "This timesheet is created at " + new Date()
  
  MailApp.sendEmail(recipient, subject, body, {
    name: "Automatic RA Timesheet Generator",
    attachments: [file.getAs(MimeType.PDF)]
  })
}

function getQuota(){
  var quota = MailApp.getRemainingDailyQuota();
  SpreadsheetApp.getUi().alert("Daily mailing quota remaining: "+quota);
}
