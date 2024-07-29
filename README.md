function setBaseline() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  for (var i = 1; i < data.length; i++) {
    // Skip the header row
    var plannedStartDate = data[i][1];
    var plannedEndDate = data[i][2];
    
    // Only set the baseline if it is currently empty
      if (!data[i][3] && !data[i][4]) {
      sheet.getRange(i + 1, 4).setValue(plannedStartDate);
      sheet.getRange(i + 1, 5).setValue(plannedEndDate);
    }
  }
  
  SpreadsheetApp.getUi().alert('Baseline has been set.');
}

function updateBaseline() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  for (var i = 1; i < data.length; i++) {
    // Skip the header row
    var plannedStartDate = data[i][1];
    var plannedEndDate = data[i][2];
    
    // Update the baseline dates with the current planned dates
    sheet.getRange(i + 1, 4).setValue(plannedStartDate);
    sheet.getRange(i + 1, 5).setValue(plannedEndDate);
  }
  
  SpreadsheetApp.getUi().alert('Baseline has been updated.');
}

function clearBaseline() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  for (var i = 1; i < data.length; i++) {
    // Skip the header row
    sheet.getRange(i + 1, 4).setValue('');
    sheet.getRange(i + 1, 5).setValue('');
  }
  
  SpreadsheetApp.getUi().alert('Baseline has been cleared.');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Project Management')
    .addItem('Set Baseline', 'setBaseline')
    .addItem('Update Baseline', 'updateBaseline')
    .addItem('Clear Baseline', 'clearBaseline')
    .addToUi();
}
