var APP_NAME = 'Spreadsheet Utils';
var REFRESH_RANGE_PROPERTY_KEY = 'refreshRangePropertyKey';

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu(APP_NAME)
      .addItem('Show', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle(APP_NAME)
      .setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

function onSetRefreshRangeButtonClick(refreshRange) {
  var activeSheet = getActiveSheet();
  setProperty(REFRESH_RANGE_PROPERTY_KEY, refreshRange, activeSheet);
  
  refreshFormulas(activeSheet);
}

function onRefreshFormulas() {
  getAllSheets().forEach(refreshFormulas);
}

function refreshFormulas(sheet) {
  var refreshRange = getProperty(REFRESH_RANGE_PROPERTY_KEY, sheet);
  if (!refreshRange) {
    return;
  }
  
  var range = sheet.getRange(refreshRange);
  var formulas = range.getFormulas();
  var values = range.getValues();
  
  range.clearContent();
  SpreadsheetApp.flush();
  traverseCells(range, function(cell, row, column) {
    // ranges use 1 based indices, arrays use 0 based indices
    var rowIndex = row - 1;
    var columnIndex = column -1;
    var formula = formulas[rowIndex][columnIndex];
    var value = values[rowIndex][columnIndex];
    
    if (formula) {
      cell.setFormula(formula);
    } else {
      cell.setValue(value);
    }
  });
}

function onGetRefreshRangeButtonClick() {
  return getProperty(REFRESH_RANGE_PROPERTY_KEY, getActiveSheet());
}

function onClearRefreshRangeButtonClick() {
  clearProperty(REFRESH_RANGE_PROPERTY_KEY, getActiveSheet());
}

function onGenerateFormulasButtonClick(dataSheetName, 
                                       dataKeyRangeA1, 
                                       dataValueRangeA1, 
                                       targetSheetName, 
                                       targetKeyRangeA1, 
                                       targetValueRangeA1,
                                       targetOutputRangeA1,
                                       formulaTemplate) {
  
  if (!dataSheetName 
      || !dataKeyRangeA1 
      || !dataValueRangeA1 
      || !targetSheetName 
      || !targetKeyRangeA1 
      || !targetValueRangeA1
      || !targetOutputRangeA1
      || !formulaTemplate) {
    showErrorMessage('Missing Input');
    return;
  }
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName(dataSheetName);
  var targetSheet = spreadsheet.getSheetByName(targetSheetName);
  if (dataSheet == null || targetSheet == null) {
    showErrorMessage('Invalid Sheet Name');
    return;
  }
  
  var dataKeyRange = dataSheet.getRange(dataKeyRangeA1);
  var dataValueRange = dataSheet.getRange(dataValueRangeA1);
  var targetKeyRange = targetSheet.getRange(targetKeyRangeA1);
  var targetValueRange = targetSheet.getRange(targetValueRangeA1);
  var targetOutputRange = targetSheet.getRange(targetOutputRangeA1);
  if (!rangesHaveEqualDimensions(dataKeyRange, dataValueRange) 
      || !rangesHaveEqualDimensions(targetKeyRange, targetValueRange)
      || !rangesHaveEqualDimensions(targetValueRange, targetOutputRange)) {
    showErrorMessage('Incompatible Ranges');
    return;
  }
  
  var dataValueCellMap = mapRanges(dataKeyRange, dataValueRange);
  
  traverseCells(targetKeyRange, function(targetKeyCell, row, column) {
    if (!targetKeyCell.isBlank()) {
      var dataValueCell = dataValueCellMap[targetKeyCell.getValue()];
      var targetValueCell = targetValueRange.getCell(row, column);
      var targetOutputCell = targetOutputRange.getCell(row, column);
      
      var useGlobalNotationForData = dataSheetName != targetSheetName;
      var useStaticNotationForData = true;
      var dataValueCellA1 = getA1Notation(dataValueCell, useStaticNotationForData, useGlobalNotationForData);
      var targetValueCellA1 = getA1Notation(targetValueCell);
      
      var formula = generateFormula(formulaTemplate, dataValueCellA1, targetValueCellA1);
      targetOutputCell.setFormula(formula);
    }
  });
}

function generateFormula(formulaTemplate, dataValueCellA1, targetValueCellA1) {
  return formulaTemplate.replace('@data', dataValueCellA1)
                        .replace('@target', targetValueCellA1);
}

function getA1Notation(cell, useStaticNotation, useGlobalNotation) {
  var result = cell.getA1Notation();
  
  if (useStaticNotation) {
    result = "$" + result.charAt(0) + "$" + result.charAt(1);
  }
    
  if (useGlobalNotation) {
    result = "'" + cell.getSheet().getSheetName() + "'!" +  result;
  }
  
  return result;
}

function showErrorMessage(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Validation Error', message, ui.ButtonSet.OK);
}

function rangesHaveEqualDimensions(rangeA, rangeB) {
  return rangeA != null && rangeB != null
      && rangeA.getNumRows() == rangeB.getNumRows() 
      && rangeA.getNumColumns() == rangeB.getNumColumns();
}

function mapRanges(keyRange, valueRange) {
  var map = {};
  
  traverseCells(keyRange, function(keyCell, row, column) {
    if (!keyCell.isBlank()) {
      var valueCell = valueRange.getCell(row, column);
      map[keyCell.getValue()] = valueCell;
    }
  });
  
  return map;
}

function traverseCells(range, visitor) {
  for (var row = 1; row <= range.getNumRows(); row++) {
    for (var column = 1; column <= range.getNumColumns(); column++) {
      var cell = range.getCell(row, column);
      visitor(cell, row, column);
    }
  }
}

function getActiveCell() {
  return getActiveSheet().getActiveCell().getA1Notation();
}

function getActiveRange() {
  return getActiveSheet().getActiveRange().getA1Notation();
}

function getActiveSheetName() {
  return getActiveSheet().getSheetName();
}

function getActiveSheet() {
  return SpreadsheetApp.getActiveSheet();
}

function getAllSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets();
}

function getProperty(key, sheet) {
  return PropertiesService.getScriptProperties()
    .getProperty(prependSheetIdToPropertyKey(key, sheet));
}

function setProperty(key, value, sheet) {
  PropertiesService.getScriptProperties()
    .setProperty(prependSheetIdToPropertyKey(key, sheet), value);
}

function clearProperty(key, sheet) {
  PropertiesService.getScriptProperties()
    .deleteProperty(prependSheetIdToPropertyKey(key, sheet));
}

function prependSheetIdToPropertyKey(key, sheet) {
  if (!sheet) {
    return key;
  }
  
  return sheet.getSheetId() + ' ' + key;
}
