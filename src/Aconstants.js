
// base values
const DEFAULT_VALUE = "";

const SSheet = SpreadsheetApp.getActiveSpreadsheet();
const RELEASE = SSheet.getId() != "1WbZsXlvklQqtlzMroRSMNnj75bis2DiAmjz3cnYspnk"
const SLIDES_BASE_TEMPLATE = "1Q7UJHX0h_dZZQRc9Kmo_eMDW_GFTx8_MeQVNCO1JDCs"
const REPORT_EMOJI = "🔀"

// blockLogic numbers of columns
const COLUMN_LEVEL = 2
const COLUMN_REAL_TASK = 3
const COLUMN_CODE = 4
const COLUMN_TASK = 5
const COLUMN_DURATION = 6
const COLUMN_START = 7
const COLUMN_END = 8
const COLUMN_PROGRESS = 9

const ROW_START = 9

// sheetssets sheet names 
const INIT_PAGE = "Инициация";
const TEMPLATE_MAP = "Карта проекта"
const TEMPLATE_BLOCK = "Шаблон блока";

const DYNAMIC_GRAPH = "Динамика"
const STATUS_GRAPH = "Статус"

const COMMANDO = "Команда"
const TEMPLORARY = "Текущие задачи"

// timeline tag
const TAG_TIMELINE = "{{timeline}}"

// slides tags

const SLIDES_SHEET_ZONE_TAG = "{{sheetZone}}"
const SLIDES_CONTENT_TABLE_TAG = "{{contentTable}}"
const SLIDES_CONTENT_FILED_TAG = "{{contentField}}"
const SLIDES_CONTENT_NAME_TAG = "Короткий заголовок:"
