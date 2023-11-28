const SOLID_HEADER_COLOR = "#B7B7B7"
const HEADER_BACK = "#D9D9D9"

const DARK_GRAY = "#333333";


function setBorder(range) {
  range.setBorder(true, true, true, true, true, true, SOLID_HEADER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  return range
}

function setTopBorder(range) {
  range.setBorder(true, null, null, null, null, null, SOLID_HEADER_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  return range;
}

function setBlockBorder(range) {
  range.setBorder(true, true, null, true, null, null, SOLID_HEADER_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  return range;
}


function setRowStyleMain(range, leftAlignedCells = []) {
  range.setHorizontalAlignment('center');

  leftAlignedCells.forEach(cell => {
    range.getCell(1, cell).setHorizontalAlignment('left');
  });

  return range
}

function rotateAndAlign(range) {
  range.setRotation(-90);
  range.setVerticalAlignment("bottom");
  range.setHorizontalAlignment("center");
  return range
}

function setBackHeader(range) {
  range.setBackground(HEADER_BACK);
  return range
}


function fixBorders() {
  let projectMap = SSheet.getSheetByName(TEMPLATE_MAP)
  var mapValues = projectMap.getDataRange().getValues();
  var lastRow = 0;
  for (var i = mapValues.length - 1; i >= 0; i--) {
    if (mapValues[i][0]) {
      lastRow = i + 1;
      let timeline = findTimelineBorders(projectMap)
      fixBordersOnRange(projectMap.getRange(timeline.endRow + 1, 1, lastRow - timeline.endRow, projectMap.getLastColumn()), timeline)
      break;
    }
  }

  getBlockSheets().forEach(sheet => {
    let = timeline = findTimelineBorders(sheet)
    let range = sheet.getRange(ROW_START, 1, sheet.getLastRow() - ROW_START + 1, sheet.getLastColumn())
    fixBordersOnRange(range, timeline)
  })
}

function fixBordersOnRange(range, timeline) {
  let sheet = range.getSheet()
  setBorder(range)
  if (!timeline) {
    timeline = findTimelineBorders(sheet)
  }
  let colHei = range.getLastRow() + 1 - timeline.row
  let merged = sheet.getRange(timeline.row, timeline.startCol, 1, timeline.col - timeline.startCol + 1).getMergedRanges()
  merged.forEach(mer => { 
    setBlockBorder(sheet.getRange(timeline.row,  mer.getColumn(), colHei, mer.getWidth()))  
  })
}