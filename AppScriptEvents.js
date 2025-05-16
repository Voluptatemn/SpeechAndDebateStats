function makeAttendanceSheet() {


    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("backend");
  
  
    sheet.appendRow(["count","date","form created"])
  
  
    var values = sheet.getDataRange().getValues();
  
  
    var today = sheet.getRange(values.length,2).setValue('=TEXTJOIN("",True,TODAY())').getValue()
  
  
    sheet.getRange(values.length,2).setValue(today)
  
  
    sheet.getRange(values.length,1).setValue(values.length - 1)
  
  
  
  
  
  
    var title = String("Event Attendance " + today);
  
  
    var form = FormApp.create(title);
    var response_sheet = SpreadsheetApp.create(title + " Responses")
  
  
    var form_file = DriveApp.getFileById(form.getId())
    var sheet_file = DriveApp.getFileById(response_sheet.getId())
    
    var folder = DriveApp.getFolderById("1NPeCx-ayp8eIh9D9HHNRIu-aU5jzdnme");
    form_file.moveTo(folder)
    sheet_file.moveTo(folder)
  
  
    var response_id = response_sheet.getId()
    var form_id = form.getId()
  
  
    form.setDestination(FormApp.DestinationType.SPREADSHEET,response_id)
   
    sheet.getRange(values.length,3).setValue(form_id)
    sheet.getRange(values.length,4).setValue(response_id)
  
  
    sheet.getRange(values.length,5).setValue(form.getEditUrl())
    sheet.getRange(values.length,6).setValue(response_sheet.getUrl())
  
  
    sheet.getRange(values.length,7).setValue(0)
  
  
    console.log("finished")
    form.setCollectEmail(true)
  
  
    form.setLimitOneResponsePerUser(true)
  
  
    form.setRequireLogin(true)
  
  
    var item = form.addMultipleChoiceItem();
  
  
    // Set some choices with go-to-page logic.
    var old_member = item.createChoice('Yes');
    var new_member = item.createChoice('No');
  
  
  // For GO_TO_PAGE, just pass in the page break item. For CONTINUE (normally the default), pass in
  // CONTINUE explicitly because page navigation cannot be mixed with non-navigation choices.
    item.setChoices([new_member,old_member]);
  
  
    item.setTitle("Are you a new member to this event?")
        .setHelpText("If you have never attended a practice for this event, select yes. If you have attended at least one prior practice for this event, select no")
    item.setRequired(true)
  
  
  
  
    Logger.log('Published URL: ' + form.getPublishedUrl());
    Logger.log('Editor URL: ' + form.getEditUrl());
  }
  