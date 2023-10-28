function readInitData(initSheet = null) {
  if (!initSheet) {
    initSheet = SSheet.getSheetByName(INIT_PAGE);
  }
  const rows = initSheet.getSheetValues(2, 2, initSheet.getLastRow() - 1, 2);
  let dataSet = {};
  for (let [value, key] of rows) {
    if (!value) continue;
    if (!dataSet.hasOwnProperty(key)) {
      dataSet[key] = value;
    } else if (Array.isArray(dataSet[key])) {
      dataSet[key].push(value);
    } else {
      dataSet[key] = [dataSet[key], value];
    }
  }
  return dataSet
}


function projectInit() {
  const init = SSheet.getSheetByName(INIT_PAGE)
  let dataSet = readInitData(init)
  storeTimelineStartDate(dataSet.date_start)
  storeTimelineEndDate(dataSet.date_end)


  let projectMap = SSheet.getSheetByName(TEMPLATE_MAP)
  let template = SSheet.getSheetByName(TEMPLATE_BLOCK)

  if (RELEASE) SSheet.deleteSheet(init);
  if (RELEASE) template.hideSheet()

  drawDefaultTimeline(projectMap)

  replacePlaceholders(projectMap, dataSet)
  projectMap.showSheet()

  if (!dataSet.block) {
    Browser.msgBox('На странице инициализации должны быть строки помеченные переменной "block" (3-я колонка)');
    return
  }

  if (!Array.isArray(dataSet.block)) {
    dataSet.block = [dataSet.block]
  }
  for (let block of dataSet.block) {
    createNewBlock(block, template, projectMap)
  }

  getHided().forEach(e => SSheet.getSheetByName(e).showSheet())
}


/**
 * Заменяет плейсхолдеры в ячейках листа на значения из предоставленного объекта данных.
 * Возвращает объект, содержащий информацию о произведенных заменах.
 *
 * @param {string|Object} sheetInput - Имя листа или объект листа, на котором нужно произвести замену.
 * @param {Object} dataSet - Объект данных для замены плейсхолдеров. Ключи объекта соответствуют плейсхолдерам без {{}}.
 * @return {Object} replacements - Объект, содержащий информацию о произведенных заменах. Ключи объекта соответствуют использованным плейсхолдерам, а значения - массивы объектов Range, где были произведены замены.
 *
 * @example
 * // Предположим, что на листе 'Sheet1' в ячейке A1 находится текст "{{placeholder}}".
 * // Этот код заменит текст в A1 на "replacement value" и вернет объект, содержащий информацию о произведенной замене.
 * var replacements = replacePlaceholders('Sheet1', {placeholder: 'replacement value'});
 *
 * @customfunction
 */
function replacePlaceholders(sheetInput, dataSet) {
  let sheet = typeof sheetInput === 'string' ? SSheet.getSheetByName(sheetInput) : sheetInput
  let sheetRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
  const values = sheetRange.getValues();
  const formulas = sheetRange.getFormulas()
  const replacements = {}

  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      const cellValue = values[row][col];
      if (typeof cellValue === 'string' && cellValue.startsWith("{{") && cellValue.endsWith("}}") && formulas[row][col] == "") {
        const key = cellValue.slice(2, -2).trim();
        const newValue = dataSet[key] ? dataSet[key] : DEFAULT_VALUE;
        let cell = sheet.getRange(row + 1, col + 1)
        cell.setValue(newValue);
        if (replacements[key]) {
          replacements[key].push(cell);
        } else {
          replacements[key] = [cell];
        }
      }
    }
  }
  return replacements
}
