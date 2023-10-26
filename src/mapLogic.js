
function newBlockRow(projectMap) {
  var values = projectMap.getSheetValues(1, 1, projectMap.getLastRow(), 1);
  var rowIndexToAdd = -1;
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0]) {
      const cell = projectMap.getRange(i + 1, 1).getCell(1, 1)
      if (cell.isPartOfMerge()) {
        var mergedSize = cell.getMergedRanges()[0].getNumRows()
        rowIndexToAdd = i + mergedSize
      } else {
        rowIndexToAdd = i + 1;
      }
      break;
    }
  }
  if (rowIndexToAdd === -1) { rowIndexToAdd = values.length }
  rowIndexToAdd += 1
  if (rowIndexToAdd > projectMap.getLastRow()) {
    projectMap.appendRow([""])
  } else {
    projectMap.insertRowAfter(rowIndexToAdd);
  }
  return rowIndexToAdd
}


function createNewBlock(name, template = null, projectMap = null) {
  if (!template) { template = SSheet.getSheetByName(TEMPLATE_BLOCK) }
  if (!projectMap) { projectMap = SSheet.getSheetByName(TEMPLATE_MAP) }

  let newBlockIndex = newBlockRow(projectMap)

  var aboveCellValue = projectMap.getRange(newBlockIndex - 1, 1).getValue();
  var orderNumber = isNaN(aboveCellValue) ? 1 : aboveCellValue + 1;

  const newSheet = template.copyTo(SSheet);
  newSheet.setName(name);

  drawDefaultTimeline(newSheet)

  let replaces = replacePlaceholders(newSheet, {
    code: orderNumber,
    blockName: name
  })

  newSheet.showSheet()
  SSheet.setActiveSheet(newSheet);

  const codeCellAddress = replaces.code[0].getA1Notation();
  const nameCellAddress = replaces.blockName[0].getA1Notation();
  const dataRow = replaces.blockName[0].getRow()

  const sname = `'${name}'`
  const newBlockValues = [
    `=${sname}!${codeCellAddress}`, // Ячейка с кодом
    `=${sname}!${nameCellAddress}`, // Ячейка с именем
    `=COUNTIF(${sname}!C:C; TRUE)`,
    `=COUNTIFS(${sname}!C:C; TRUE; ${sname}!I:I; 1)`,
    `=${sname}!G${dataRow}`, // Дата начала
    `=${sname}!H${dataRow}`, // Дата завершения
    `=${sname}!I${dataRow}`, // Состояние
    // Первая версия, не самая оптимальная
    // `=ARRAYFORMULA(FILTER('${name}'!J${dataRow}:${dataRow}; ISNUMBER('${name}'!J${dataRow - 1}:${dataRow - 1})))`
    // Оптимальная версия, но с фиксацией на первую колонкку
    `=QUERY(${sname}!J${dataRow}:INDEX(${sname}!${dataRow}:${dataRow}; MATCH(1E+30; ${sname}!J${dataRow - 1}:${dataRow - 1})); "select *")`,
    // Версия с точными ссылками
    // `=QUERY(INDIRECT("${sname}!J${dataRow}:" & ADDRESS(${dataRow}; MATCH(1E+30; INDIRECT("${sname}!J${dataRow - 1}:${dataRow - 1}")))); "select *")`
  ]
  var newRange = projectMap.getRange(newBlockIndex, 1, 1, newBlockValues.length);
  newRange.setValues([newBlockValues])
  setBorder(setRowStyleMain(newRange, [2]))

  onBlocksChange()
}


function getBlockNames() {
  var sheet = SSheet.getSheetByName(TEMPLATE_MAP);
  var data = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();

  var blockNames = data
    .filter(row => row[0] && !isNaN(Number(row[0])))
    .map(row => row[1]);
  return blockNames;
}


function extractSheetName(formula) {
  if (formula.startsWith('=')) {
    formula = formula.slice(1);
  }
  let exclamIndex = formula.indexOf('!');
  if (exclamIndex === -1) {
    return null; 
  }
  let potentialSheetName = formula.slice(0, exclamIndex);
  if (potentialSheetName.startsWith("'") && potentialSheetName.endsWith("'")) {
    return potentialSheetName.slice(1, -1);
  }
  return potentialSheetName;
}



function deleteSheetsAndRows(selectedNames = ['asfasdsdvafadfgadfg']) {
  var namesSet = new Set(selectedNames);

  var sheet = SSheet.getSheetByName(TEMPLATE_MAP);
  var dataRange = sheet.getRange(1, 2, sheet.getLastRow(), 1)
  var dataColumn = dataRange.getValues().flat()
  var dataFormulas = dataRange.getFormulas().flat()

  var failedToDelete = [];

  for (var i = dataColumn.length - 1; i >= 0; i--) {
    let formula = dataFormulas[i]
    if (formula && namesSet.has(dataColumn[i])) {
      var sheetName =  extractSheetName(formula)
      if (sheetName) {
        var sheetToDelete = SSheet.getSheetByName(sheetName);
        if (sheetToDelete) {
          Logger.log(`deleting ${sheetName}`)
          SSheet.deleteSheet(sheetToDelete);
          sheet.deleteRow(i + 1);
          continue
        }
      }
      failedToDelete.push(dataColumn[i])
    }
  }

  if (failedToDelete.length > 0) {
    Logger.log(`failedToDelete ${failedToDelete}`)
    var message = 'Пропущены блоки:\n- ' + failedToDelete.join('\n- ') + "\n\nУдалите лист блока, затем удалите строку блока и проверьте работоспособность элементов таблицы.";
    SSheet.toast(message, 'Не получилось удалить', 10);
  }
  onBlocksChange()
}


