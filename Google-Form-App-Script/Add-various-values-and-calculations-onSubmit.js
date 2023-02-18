const sheet_ID = "SHEET_ID";
const sheet_name = "SHEE_NAME";
const sheet = SpreadsheetApp.openById(sheet_ID).getSheetByName(sheet_name);
const lastRowIndex = sheet.getLastRow()

function onFormSubmit(e) {
  // Seller
  var cell = 5
  addDataToLastRowCell(cell, "Abir")

  // Paid ?
  cell = 6
  var ans = getDataToLastRowCell(cell)
  if (ans !== "YES")
    addDataToLastRowCell(cell, "NO")

  // Current Date
  cell = 1
  const d = new Date(getDataToLastRowCell(cell))
  addDataToLastRowCell(cell, d.toString().slice(0, 15))

  // Package
  cell = 3
  var p = getDataToLastRowCell(cell)
  var parts = p.toString().split(" - ")
  addDataToLastRowCell(cell, parts[0])
  
  // Amount
  if (p.toString().includes("TK")) {
    cell = 4
    var tk = parts[1].split(" ")[0]
    addDataToLastRowCell(cell, tk)
  }

  // Expires
  cell = 7
  var years = 1
  if (p.toString().includes("2 Year"))
    years = 2;
  else if (p.toString().includes("3 Year"))
    years = 3
  const ex = addYearsToDate(years, d)
  addDataToLastRowCell(cell, ex.toString().slice(0, 15))

  // Center Last Row
  var totalCell =8
  var r = sheet.getRange(1, 1, lastRowIndex, totalCell)
  r.setHorizontalAlignment("center")
}

function changeColorLastRowCell(color, cell) {
  sheet.getRange(lastRowIndex, cell).setBackground(color)
}

function addYearsToDate(years, d) {
  return new Date(d.getFullYear() + years, d.getMonth(), d.getDate() + 1)
}

function addDataToLastRowCell(cellIndex, data) {
  sheet.getRange(lastRowIndex, cellIndex).setValue(data)
}

function getDataToLastRowCell(cellIndex) {
  return sheet.getRange(lastRowIndex, cellIndex).getValue()
}
