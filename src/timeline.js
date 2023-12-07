function extendTimeline(startDate, endDate) {
  startDate = new Date(startDate);
  endDate = new Date(endDate);
  for (let s of getTimelineSheets()) {
    generateTimeline(s, startDate, endDate)
  }
  fixBorders()
}


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
      let startDateCol, endDateCol, startDate, endDate;
      for (let j = 0; j < firstLine.length; j++) {
        let secIsDate = secondLine[j] instanceof Date
        let fisIsDate = firstLine[j] instanceof Date
        if (!startDateCol && secIsDate && fisIsDate) {
          startDateCol = j;
          startDate = secondLine[j]
        }
        if (startDateCol && !(secIsDate && (fisIsDate || !firstLine[j]))) {
          break
        }
        endDateCol = j;
        endDate = secondLine[j]
      }
      if (endDateCol > startDateCol) {
        return {
          row: i + 1,
          endRow: i + 2,
          startCol: startDateCol + 1,
          col: endDateCol + 1,
          isNew: false,
          startDate,
          endDate
        };
      }
    }
  }
  throw new Error(`Не удалось найти таймлайн на листе ${sheet.getName()}`);
}


function generateTimeline(sheet, startDate, endDate) {
  let borders = findTimelineBorders(sheet)
  sheet = typeof sheet === 'string' ? SSheet.getSheetByName(sheet) : sheet
  if (borders.isNew || endDate > borders.endDate) {
    insertRight(sheet, startDate, endDate, borders)
  }
  if (!borders.isNew && startDate < borders.startDate) {
    insertLeft(sheet, startDate, borders)
  }
}

function insertLeft(sheet, startDate, borders) {
  let timelineRow = borders.row
  let timelineColumn = borders.startCol
  let endDate = borders.startDate
  let undelete = 0
  let cell = sheet.getRange(timelineRow, timelineColumn).getCell(1, 1)
  if (cell.isPartOfMerge()) {
    let mergedRange = cell.getMergedRanges()[0];
    let endColumn = mergedRange.getColumn() + mergedRange.getNumColumns() - 1;
    undelete = mergedRange.getNumColumns() - 1
    endDate = sheet.getRange(timelineRow + 1, endColumn, 1, 1).getValue()
  }

  const day = 24 * 60 * 60 * 1000;
  const week = 7 * day;
  let diffWeeks = Math.ceil((endDate - startDate) / week);
  startDate = new Date(endDate - diffWeeks * week);
  const columnsInsert = diffWeeks - undelete;

  if (columnsInsert <= 0) { return }
  sheet.insertColumnsAfter(timelineColumn, columnsInsert)

  if (sheet.getName() != TEMPLATE_MAP) {
    const dataStartRow = borders.endRow + 2
    const length = sheet.getLastRow() - dataStartRow + 1
    let source = sheet.getRange(dataStartRow, timelineColumn, length, 1)
    sheet.getRange(dataStartRow, timelineColumn + columnsInsert, length, 1).setValues(source.getValues());
    source.clearContent();
    cell.setValue(TAG_TIMELINE)
  }

  drawTimeline(sheet, timelineRow, timelineColumn, startDate, endDate)
  storeTimelineStartDate(startDate)
}


function insertRight(sheet, startDate, endDate, borders) {
  let timelineRow = borders.row
  let timelineColumn = borders.col
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
  const day = 24 * 60 * 60 * 1000
  startDate = startDate || getTimelineStartDate() || new Date()
  endDate = endDate || getTimelineEndDate() || new Date(Date.now() + 120 * day);

  const columnsInsert = Math.ceil((endDate - startDate) / (7 * day)) - undelete
  if (columnsInsert <= 0) { return }
  sheet.insertColumnsAfter(initTimelineColumn, columnsInsert)
  drawTimeline(sheet, timelineRow, timelineColumn, startDate, endDate)
  storeTimelineEndDate(endDate)

}


function drawTimeline(sheet, timelineRow, timelineColumn, startDate, endDate) {
  let dates = []

  endDate = addSeven(endDate)
  while (startDate < endDate) {
    dates.push(startDate)
    startDate = addSeven(startDate)
  }

  let cell = sheet.getRange(timelineRow, timelineColumn).getCell(1, 1)
  let formulas = [dates]
  let borderHeight = 2
  if (sheet.getName() != TEMPLATE_MAP) {
    borderHeight = sheet.getLastColumn() - timelineRow + 1
    formulas.push(new Array(dates.length).fill('=IF(INDIRECT("R[-1]C[0]"; FALSE)>TODAY();""; (SUM(INDIRECT("R[1]C[0]:R[900048576]C[0]"; FALSE)) + COUNTIF( INDIRECT("R[1]C[" & 10 - COLUMN() & "]:R[900048576]C[-1]"; FALSE);1 )) / $J$5)'))

    if (cell.getValue() == TAG_TIMELINE) {
      formulas[1][0] = '=IF(J$7>TODAY();"";(SUM(J9:J) / COUNTIF($C9:$C;TRUE)))'
    }
  }

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

      setBlockBorder(sheet.getRange(timelineRow, timelineColumn + monthStartIndex, borderHeight, monthRangeLength))

      if (i !== dates.length) {
        currentMonth = tempMonth
        monthStartIndex = i;
      }
    }
  }
}

function addSeven(date) {
  let newDate = new Date(date);
  newDate.setDate(date.getDate() + 7);
  return newDate
}