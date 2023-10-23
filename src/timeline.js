const TAG_TIMELINE = "{{timeline}}"


function drawDefaultTimeline(sheet = SSheet.getSheetByName("Карта проекта (копия)")) {
  let borders = findTimelineBorders(sheet)
  drawTimeline(sheet, borders.row, borders.col)
}


function extendTimeline(startDate, endDate) {
  startDate = new Date(startDate);
  endDate = new Date(endDate);
  Logger.log(startDate)
  Logger.log(endDate)
  // for (let s of getTimelineSheets()) {
  //   let borders = findTimelineBorders(s)
  //   drawTimeline(s, borders.row, borders.col, startDate, dateObj)
  // }
}


/**
 * Функция для поиска границ таймлайна на указанном листе Google Таблицы. (правый верхний угол)
 * 
 * @function findTimelineBorders
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Объект листа, на котором необходимо найти таймлайн.
 * 
 * @returns {Object} - Объект, содержащий информацию о расположении таймлайна.
 * @returns {number} Object.row - Номер строки, в которой начинается таймлайн.
 * @returns {number} Object.col - Номер столбца, в котором заканчивается таймлайн.
 * @returns {boolean} Object.isNew - Флаг, указывающий, является ли таймлайн новым (определяется по наличию маркера "{{timeline}}").
 * 
 * @throws {Error} Если таймлайн не удалось найти на листе.
 * 
 * @example
 * // Нахождение границ таймлайна на активном листе.
 * let borders = findTimelineBorders(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
 * Logger.log(`Таймлайн начинается с строки ${borders.row} и заканчивается в столбце ${borders.col}. Новый: ${borders.isNew}`);
 * 
 * @description
 * Функция сканирует указанный лист, начиная с верхнего левого угла (A1), и ищет границы таймлайна по следующим критериям:
 * - Если в ячейке встречается строка "{{timeline}}", то считается, что найден новый таймлайн.
 * - Если подряд идут две строки, в которых присутствуют объекты Date, определённым образом расположенные, то считается, что найден существующий таймлайн.
 *   Детальные условия расположения объектов Date определяются кодом функции.
 * 
 * Если таймлайн обнаружен, функция возвращает объект с информацией о его расположении.
 * Если таймлайн не обнаружен, функция генерирует исключение с сообщением об ошибке.
 */
function findTimelineBorders(sheet) {
  let data = sheet.getSheetValues(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

  for (let i = 0; i < data.length - 1; i++) {
    let firstLine = data[i];
    let secondLine = data[i + 1];

    let timelineStartCol = firstLine.indexOf(TAG_TIMELINE);
    if (timelineStartCol !== -1) {
      return { row: i + 1, col: timelineStartCol + 1, isNew: true };
    }

    if (firstLine.some(cell => cell instanceof Date) && secondLine.some(cell => cell instanceof Date)) {
      let startDateCol, endDateCol;
      for (let j = 0; j < firstLine.length; j++) {
        let secIsDate = secondLine[j] instanceof Date
        let fisIsDate = firstLine[j] instanceof Date
        if (!startDateCol && secIsDate && fisIsDate) {
          startDateCol = j;
        }
        if (startDateCol && !(secIsDate && (fisIsDate || !firstLine[j]))) {
          break
        }
        endDateCol = j;
      }
      if (endDateCol > startDateCol) {
        return {
          row: i + 1,
          endRow: i + 2,
          startCol: startDateCol + 1,
          col: endDateCol + 1,
          isNew: false
        };
      }
    }
  }
  throw new Error(`Не удалось найти таймлайн на листе ${sheet.getName()}`);
}


function drawTimLeft() {
  let sheet = SSheet.getSheetByName('fff (копия)')
  let borders = findTimelineBorders(sheet)
  const date = new Date(2023, 6, 6);
  drawTimeline(sheet, borders.row, borders.startCol, date)
}


/**
 * Функция для отрисовки таймлайна на Google Таблице с использованием Google Apps Script.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|string} sheet - Объект листа таблицы или имя листа,
 *  на котором должен быть отрисован таймлайн.
 * @param {number} timelineRow - Номер строки, начиная с которой будет отрисовываться таймлайн.
 * @param {number} timelineColumn - Номер столбца, с которого начнется отрисовка таймлайна.
 * @param {Date} [startDate=new Date()] - Начальная дата таймлайна. Если не указана, используется текущая дата.
 * @param {Date} [endDate=new Date(Date.now() + 120*24*60*60*1000)] - Конечная дата таймлайна.
 *  Если не указана, используется текущая дата плюс 120 дней.
 * @param {number} [height_higlight=2] - Высота подсветки заголовка таймлайна в строках.
 * 
 * @example
 * // Отрисовка таймлайна на активном листе начиная с 2 строки и 1 колонки,
 * // с начальной датой 2023-01-01 и конечной 2023-04-30.
 * drawTimeline(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), 2, 1, new Date('2023-01-01'), new Date('2023-04-30'));
 *
 * @throws {Error} Если переданное имя листа не существует.
 * @throws {TypeError} Если типы переданных параметров не соответствуют ожидаемым.
 * 
 * @returns {void}
 */
function drawTimeline(sheet, timelineRow, timelineColumn, startDate, endDate, height_higlight = 2) {
  sheet = typeof sheet === 'string' ? SSheet.getSheetByName(sheet) : sheet

  const initTimelineColumn = timelineColumn
  let undelete = 0
  let cell = sheet.getRange(timelineRow, timelineColumn).getCell(1, 1)
  if (cell.isPartOfMerge()) {
    timelineColumn = cell.getMergedRanges()[0].getColumn()
    undelete = initTimelineColumn - timelineColumn
    startDate = sheet.getRange(timelineRow + 1, timelineColumn, 1, 1).getValue()
  }

  if (!startDate) {
    let cellValue = sheet.getSheetValues(timelineRow + 1, initTimelineColumn, 1, 1)[0][0];
    if (isValidDate(cellValue)) {
      startDate = new Date(cellValue);
    }
  }
  startDate = startDate || getTimelineStartDate() || new Date()

  endDate = endDate || getTimelineEndDate() || new Date(Date.now() + 120 * 24 * 60 * 60 * 1000);
  endDate = addSeven(endDate)

  let dates = []
  let count = 0
  while (startDate < endDate) {
    dates.push(startDate)
    count++
    startDate = addSeven(startDate)
  }

  if (dates.length - 1 - undelete <= 0) {
    return
  }

  let formulas = [dates]
  if (sheet.getName() != TEMPLATE_MAP) {
    formulas.push(new Array(count).fill('=IF(INDIRECT("R[-1]C[0]"; FALSE)>TODAY();""; (SUM(INDIRECT("R[1]C[0]:R[900048576]C[0]"; FALSE)) + COUNTIF( INDIRECT("R[1]C[" & 10 - COLUMN() & "]:R[900048576]C[-1]"; FALSE);1 )) / $J$5)'))

    if (cell.getValue() == TAG_TIMELINE) {
      formulas[1][0] = '=IF(J$7>TODAY();"";(SUM(J9:J) / COUNTIF($C9:$C;TRUE)))'
    }
  }

  sheet.insertColumnsAfter(initTimelineColumn, dates.length - 1 - undelete)
  let dRange = sheet.getRange(timelineRow + 1, timelineColumn, formulas.length, dates.length)
    .setValues(formulas)
  setBackHeader(setBorder(dRange))

  sheet.getRange(timelineRow, timelineColumn, 1, dates.length).getMergedRanges().forEach(mr => mr.breakApart())
  let currentMonth = dates[0].getMonth();
  let monthStartIndex = 0;

  for (let i = 0; i < dates.length; i++) {
    const tempMonth = dates[i].getMonth()
    if (i === dates.length - 1) { i++ }
    if (i === dates.length || tempMonth !== currentMonth) {
      let monthRangeLength = i - monthStartIndex;
      let range = sheet.getRange(timelineRow, timelineColumn + monthStartIndex, 1, monthRangeLength);
      range.setValue(new Date(dates[monthStartIndex].getFullYear(), currentMonth, 1));
      if (monthRangeLength > 1) {
        range.mergeAcross();
      }
      setRowStyleMain(range)
      range.setNumberFormat("mmm yy");

      setBlockBorder(sheet.getRange(timelineRow, timelineColumn + monthStartIndex, height_higlight, monthRangeLength))

      if (i !== dates.length) {
        currentMonth = tempMonth
        monthStartIndex = i;
      }
    }
  }

  storeTimelineEndDate(endDate)
}

function addSeven(date) {
  let newDate = new Date(date);
  newDate.setDate(date.getDate() + 7);
  return newDate
}