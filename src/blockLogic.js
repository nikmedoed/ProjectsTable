const COLUMN_LEVEL = 2
const COLUMN_REAL_TASK = 3
const COLUMN_CODE = 4
const COLUMN_TASK = 5
const COLUMN_DURATION = 6
const COLUMN_START = 7
const COLUMN_END = 8
const COLUMN_PROGRESS = 9

const ROW_START = 9

function onBlocksChange() {
  temploraryTasksFormula()
  reloadGraphs()
}

function temploraryTasksFormula() {
  const sheet = SSheet.getSheetByName(TEMPLORARY);
  const sheetsList = getBlockSheets();

  const rangeString = sheetsList.map(sheet => `'${sheet.getName()}'!C${ROW_START}:I`).join(";")

  const [adjRealTask, adjCode, adjTask, adjStart, adjEnd, adjProgress] = [COLUMN_REAL_TASK, COLUMN_CODE, COLUMN_TASK, COLUMN_START, COLUMN_END, COLUMN_PROGRESS].map(value => value - 2);

  const todayStr = `date '"& TEXT(TODAY(); "yyyy-MM-dd") &"'`;
  const dateDiff = `dateDiff(Col${adjEnd}, ${todayStr})`;
  const selectColumns = `Col${adjCode}, Col${adjTask}, Col${adjProgress}, ${dateDiff}`;
  const whereCondition = `Col${adjRealTask}=True AND Col${adjStart} <= ${todayStr} AND Col${adjProgress} < 1`;

  const labelClear = `label ${dateDiff} ''`;

  const formula = `=ARRAYFORMULA(TRIM(QUERY({${rangeString}}; "SELECT ${selectColumns} WHERE ${whereCondition} ${labelClear}"; 0)))`;
  sheet.getRange("B3").setFormula(formula);
}


function onBlockEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const column = range.getColumn();
  var values = range.getValues();
  const row = range.getRow();
  const numRows = range.getNumRows();
  const lastRow = row + numRows - 1;
  const DATES_COLUMNS = [COLUMN_DURATION, COLUMN_START, COLUMN_END]

  if (!RELEASE) {
    Logger.log(column, row, values)
    Logger.log(row, values)
    Logger.log(values)
  }

  if (row < ROW_START) return;

  for (var i = 0; i < values.length; i++) {
    var currentRow = row + i;
    [COLUMN_REAL_TASK, COLUMN_PROGRESS].forEach(e => checkAndCopyFormula(sheet, currentRow, e, ROW_START - 1))
    SpreadsheetApp.flush()
    let isReal = sheet.getRange(currentRow, COLUMN_REAL_TASK).getValue()

    values[i].forEach((currentValue, offset) => {
      var currentColumn = column + offset;

      let taskrange = sheet.getRange(currentRow, COLUMN_TASK)
      if (currentColumn === COLUMN_LEVEL && typeof currentValue === "number") {
        let task = taskrange.getValue();
        if (task !== "") {
          taskrange.setValue(fixLevel(task, currentValue));
          if (isReal) {
            DATES_COLUMNS.forEach(e => clearErrorFormula(sheet, currentRow, e))
          }
        }
      }
      if (currentColumn === COLUMN_TASK) {
        let level = sheet.getRange(currentRow, COLUMN_LEVEL).getValue();
        if (level !== "") {
          taskrange.setValue(fixLevel(currentValue, parseInt(level)));
        }
      }
      taskrange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

      if (isReal && DATES_COLUMNS.includes(currentColumn)) {
        processDateColumns(sheet, currentRow, currentColumn, currentValue);
      }

    })
  }
  processDateAgregate(sheet, lastRow)
}

function clearErrorFormula(sheet, row, column) {
  var cell = sheet.getRange(row, column);
  var formula = cell.getFormula();
  if (formula) {
    var value = cell.getValue();
    if (typeof value === 'string' && value.startsWith('#')) {
      cell.clearContent();
    }
    console.log(value)
    console.log(formula)
  }
}


function processDateColumns(sheet, row, column, value) {
  let mmm = Math.min(COLUMN_DURATION, COLUMN_START, COLUMN_END);
  let mmmax = Math.max(COLUMN_DURATION, COLUMN_START, COLUMN_END);
  let datesRange = sheet.getRange(row, mmm, 1, mmmax - mmm + 1)

  let dates = datesRange.getValues()[0];
  let duration = dates[COLUMN_DURATION - mmm];
  let start = dates[COLUMN_START - mmm];
  let end = dates[COLUMN_END - mmm];

  if ([duration, start, end].filter(v => v !== "").length >= 2) {
    let targetCellInfo;
    switch (column) {
      case COLUMN_DURATION:
        targetCellInfo = start !== "" ? calculateEnd(start, duration) : calculateStart(end, duration);
        break;
      case COLUMN_START:
        targetCellInfo = duration !== "" ? calculateEnd(start, duration) : calculateDuration(start, end);
        break;
      case COLUMN_END:
        targetCellInfo = duration !== "" ? calculateStart(end, duration) : calculateDuration(start, end);
        break;
    }
    if (targetCellInfo) {
      sheet.getRange(row, targetCellInfo[0]).setValue(targetCellInfo[1]);
    }
  }
}


function bloksDataFix() {
  for (let sheet of getBlockSheets()) {
    let last = sheet.getLastRow() - ROW_START + 1
    if (last <= 0) { continue }
    let len = COLUMN_TASK - COLUMN_LEVEL
    var values = sheet.getSheetValues(ROW_START, COLUMN_LEVEL, last, len + 1)
    var i
    for (i = values.length - 1; i >= 0; i--) {
      if (values[i][0] !== "" && values[i][0] !== null) { break }
    }
    processDateAgregate(sheet, i + ROW_START)
    values = values.map(row => [fixLevel(row[len], row[0])])
    sheet.getRange(ROW_START, COLUMN_TASK, last, 1).setValues(values)
  }
}


function processDateAgregate(sheet, row) {
  let mmmax = Math.max(COLUMN_DURATION, COLUMN_START, COLUMN_END);

  if (row == ROW_START) return
  var valuesRange = sheet.getRange(ROW_START, COLUMN_REAL_TASK, row - ROW_START + 1, mmmax - COLUMN_REAL_TASK + 1)
  var values = valuesRange.getValues();
  let formulas = valuesRange.getFormulas();

  let [shifted_duration, shifted_start, shifted_end] = [COLUMN_DURATION, COLUMN_START, COLUMN_END].map(e => e - COLUMN_REAL_TASK)

  for (var i = values.length - 1; i >= 0; i--) {
    if (!values[i][0] && (!formulas[i][shifted_duration] || !formulas[i][shifted_start] || !formulas[i][shifted_end])) {
      var currentRow = ROW_START + i;
      let fmls = [
        '=INDIRECT("R[0]C[2]"; FALSE) - INDIRECT("R[0]C[1]"; FALSE)',
        `=QUERY(INDIRECT("R[1]C2:R[" & IFERROR(MATCH(TRUE; ARRAYFORMULA(INDIRECT("B"&ROW(B${currentRow})+1&":B") <= B${currentRow}); 0) - 1; ROWS(B:B) - ROW(B${currentRow}))&"]C[0]"; FALSE); "SELECT MIN(G) WHERE C=True label MIN(G) '' "; 0)`,
        `=QUERY(INDIRECT("R[1]C2:R[" & IFERROR(MATCH(TRUE; ARRAYFORMULA(INDIRECT("B"&ROW(B${currentRow})+1&":B") <= B${currentRow}); 0) - 1 ;ROWS(B:B) - ROW(B${currentRow}))&"]C[0]"; FALSE); "SELECT MAX(H) WHERE C=True label MAX(H) '' "; 0)`
      ]
      sheet.getRange(currentRow, COLUMN_DURATION, 1, 3).setFormulas([fmls]);
    }
  }
}


function checkAndCopyFormula(sheet, row, column, formulaRow) {
  const target = sheet.getRange(row, column)
  if (!target.getFormula()) {
    let donor = sheet.getRange(formulaRow, column)
    donor.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false)
    donor.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false)
  }
}


function calculateEnd(start, duration) {
  return [COLUMN_END, new Date(start.getTime() + duration * 86400000)];
}

function calculateStart(end, duration) {
  return [COLUMN_START, new Date(end.getTime() - duration * 86400000)];
}

function calculateDuration(start, end) {
  return [COLUMN_DURATION, (end.getTime() - start.getTime()) / 86400000];
}

// function fixLevel(text, level = 0) {
//   return "   ".repeat(level) + text.trim()
// }

function fixLevel(text, level = 0) {
  const words = text.trim().split(/\s+/);
  const indent = "   ".repeat(level);
  let currentLine = indent;
  let result = [];

  for (const word of words) {
    if ((currentLine + word).length > 60) {
      result.push(currentLine);
      currentLine = indent + word + ' ';
    } else {
      currentLine += word + ' ';
    }
  }
  if (currentLine.length > 0) {
    result.push(currentLine);
  }
  return result.join('\n');
}

function checkClear(cell) {
  return cell == null || String(cell).trim() == ""
}

/**
 * Вычисляет код задачи
 *
 * @param {number} currentLevel - Уровень целевой задачи
 * @param {number} upperLevel - Уровень задачи выше
 * @param {string} upperTaskCode - Код задачи вышее
 * @return {number|string} Результат вычислений или пустая строка, если один из параметров отсутствует.
 * @customfunction
 */
function GETCODE(currentLevel, upperLevel, upperTaskCode) {
  if (checkClear(currentLevel) || checkClear(upperLevel) || checkClear(upperTaskCode)) {
    return "";
  }

  if (currentLevel > upperLevel) {
    return upperTaskCode + ".1".repeat(currentLevel - upperLevel)
  } else {
    let code = upperTaskCode.split(".")
    code = code.slice(0, code.length - (upperLevel - currentLevel));
    code.push((parseInt(code.pop()) + 1).toString())
    return code.join(".")
  }
}

/**
 * Вычисляет коды задач для диапазона
 *
 * @param {Range} levelsRange - Диапазон уровней задач
 * @param {Range} codesRange - Диапазон, куда будут записаны коды задач
 * @param {string} initialCode - Начальный код для первой задачи
 * @param {boolean} full - Добавить начальный код в возвращаемых список
 * @customfunction
 */
function GETCODERANGE(levelsRange, initialCode, full = false) {
  var levels = levelsRange;
  var codes = []

  if (full) codes.push([initialCode])
  var previousCode = initialCode;

  for (var i = 1; i < levels.length; i++) {
    previousLevel = levels[i - 1][0];
    const newVal = GETCODE(levels[i][0], previousLevel, previousCode)
    previousCode = newVal
    codes.push([newVal])
  }
  return codes
}
