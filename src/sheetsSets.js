const INIT_PAGE = "Инициация";
const TEMPLATE_MAP = "Карта проекта"
const TEMPLATE_BLOCK = "Шаблон блока";

const DYNAMIC_GRAPH = "Динамика"
const STATUS_GRAPH = "Динамика"

const COMMANDO = "Команда"
const TEMPLORARY = "Текущие задачи"

function getBlockSheets() {
  let sheets = SSheet.getSheets();
  const excludedSheetNames = new Set([INIT_PAGE, TEMPLATE_MAP, TEMPLATE_BLOCK, DYNAMIC_DATA, DYNAMIC_GRAPH, COMMANDO, TEMPLORARY]);

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