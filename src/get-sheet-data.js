// Get sheet data with column mapping
function getSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SHEET_NAME
  );
  const range = sheet.getDataRange();
  const values = range.getValues();

  if (values.length < 2) {
    throw new Error("Sheet must have at least a header row and one data row");
  }

  const headers = values[0];
  const data = values.slice(1); // Skip header row

  // Find column positions
  const columnMap = {};
  Object.values(COLUMN_HEADERS).forEach((columnName) => {
    const index = findColumnIndex(headers, columnName);
    if (index >= 0) {
      columnMap[columnName] = index;
    }
  });

  // Check required columns exist
  const requiredColumns = [
    COLUMN_HEADERS.EMAIL,
    COLUMN_HEADERS.SUBJECT,
    COLUMN_HEADERS.CONTENT,
    COLUMN_HEADERS.START_DATE,
    COLUMN_HEADERS.DEADLINE,
  ];
  for (const columnName of requiredColumns) {
    if (!(columnName in columnMap)) {
      throw new Error(`Required column "${columnName}" not found in sheet`);
    }
  }

  return {headers, data, columnMap};
}
