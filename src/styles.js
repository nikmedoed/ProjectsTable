const SOLID_HEADER_COLOR = "#B7B7B7"
const HEADER_BACK = "#D9D9D9"

const DARK_GRAY = "#333333";

function setBorder(range) {
  var color = SOLID_HEADER_COLOR
  range.setBorder(true, true, true, true, false, false, color, SpreadsheetApp.BorderStyle.SOLID);
  return range
}

function setTopBorder(range) {
  range.setBorder(true, false, false, false, false, false, SOLID_HEADER_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  return range;
}

function setBlockBorder(range) {
  range.setBorder(true, true, false, true, false, false, SOLID_HEADER_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
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

function setBackHeader(range){
  range.setBackground(HEADER_BACK);
  return range
}