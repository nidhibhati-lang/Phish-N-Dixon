function doGet(e) {
  const token = e && e.parameter ? e.parameter.token : "";
  if (!token) {
    return ContentService.createTextOutput("No tracking token provided.");
  }

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    throw new Error("Set the SPREADSHEET_ID script property before deploying this web app.");
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const worksheetName = properties.getProperty("WORKSHEET_NAME");
  const worksheet = worksheetName
    ? spreadsheet.getSheetByName(worksheetName)
    : spreadsheet.getSheets()[0];
  if (!worksheet) {
    throw new Error("The configured tracking worksheet does not exist.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const values = worksheet.getDataRange().getDisplayValues();
    if (values.length < 2) {
      return ContentService.createTextOutput("Tracking token not found.");
    }

    const headers = values[0].map(value => value.trim().toLowerCase());
    const statusColumn = headers.indexOf("status");
    const tokenColumn = headers.indexOf("tracking token");
    if (statusColumn < 0 || tokenColumn < 0) {
      throw new Error('The worksheet must contain "Status" and "Tracking Token" columns.');
    }

    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      if (values[rowIndex][tokenColumn] !== token) continue;

      const statusCell = worksheet.getRange(rowIndex + 1, statusColumn + 1);
      if (statusCell.getDisplayValue().trim().toLowerCase() === "pending") {
        statusCell.setValue("Clicked");
      }
      return ContentService.createTextOutput("Recorded.");
    }
  } finally {
    lock.releaseLock();
  }

  return ContentService.createTextOutput("Tracking token not found.");
}
