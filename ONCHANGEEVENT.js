function onChangeEvent(e) {
  if (e.changeType === 'INSERT_ROW') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Complaint Entry");

    // Make sure change happened in correct sheet
    if (ss.getActiveSheet().getName() !== "Complaint Entry") return;

    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());

    // Pass event object to your function
    handleComplaintForm({
      source: ss,
      range: range
    });
  }
}

