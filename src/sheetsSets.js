function getHided() {
  return [DYNAMIC_GRAPH, STATUS_GRAPH, TEMPLORARY]
}

function getBlockSheets() {
  let sheets = SSheet.getSheets();
  const excludedSheetNames = new Set([
    INIT_PAGE, TEMPLATE_MAP, TEMPLATE_BLOCK, DYNAMIC_GRAPH, STATUS_GRAPH, COMMANDO, TEMPLORARY
  ]);

  if (!RELEASE) {
    sheets = sheets.filter(sheet => !sheet.getName().startsWith("!"));
  }

  return sheets.filter(sheet => !excludedSheetNames.has(sheet.getName()));
}

function getBlockSheetNames() {
  const blockSheets = getBlockSheets();
  return new Set(blockSheets.map(sheet => sheet.getName()));
}

function getTimelineSheets() {
  let sheets = getBlockSheets()
  if (RELEASE) {
    sheets.push(SSheet.getSheetByName(TEMPLATE_MAP))
  }
  return sheets
}