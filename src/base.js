
const DEFAULT_VALUE = "";

const SSheet = SpreadsheetApp.getActiveSpreadsheet();
const RELEASE = SSheet.getId() != "1WbZsXlvklQqtlzMroRSMNnj75bis2DiAmjz3cnYspnk"


function addListAtNow() {
  var currentDate = new Date();
  var dateString = currentDate.toLocaleString();
  createNewBlock(dateString)
}

// function resetProgress() {
//   var sheets = SSheet.getSheets();
//   for (var i = 3; i < sheets.length; i++) {
//     SSheet.deleteSheet(sheets[i]);
//   }
//   PropertiesService.getScriptProperties().deleteAllProperties()
// }

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Управление проектом')
    .addItem('Добавить новый блок', 'createNewBlockPrompt')
    .addItem('Расширить таймлайн', 'showExtendTimeline')
    .addItem('Удалить блок(и)', 'showDeleteBlocks')

    .addItem('Быстрый лист', 'addListAtNow')
    // .addItem('Удалить сгенерированные страницы', 'resetProgress')

    .addToUi();
}


function createNewBlockPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Добавление блока', 'Введите название нового блока работ:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var blockName = response.getResponseText();
    var sheet = SSheet.getSheetByName(blockName);
    if (sheet) {
      ui.alert('Невозможно создать', 'Страница с таким именем уже существует. Повторите и введите другое имя.', ui.ButtonSet.OK);
    } else {
      createNewBlock(blockName);
    }
  }
}


function showExtendTimeline() {
  var startDate = getTimelineStartDate();
  var endDate = getTimelineEndDate();

  if (!endDate) {
    SpreadsheetApp.getUi().alert('Пока нельзя расширить таймлайн, таблица не инициализирована');
    return;
  }
  var template = HtmlService.createTemplateFromFile('DatePicker');
  template.endDate = endDate.toISOString().split('T')[0];
  template.startDate = startDate.toISOString().split('T')[0];

  var htmlOutput = template.evaluate()
    .setWidth(350)
    .setHeight(250)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Расширение таймлайна');
}



function showDeleteBlocks() {
  var blockNames = getBlockNames();
  var template = HtmlService.createTemplateFromFile('DeleteBlockPicker');
  template.blockNames = blockNames;

  var maxWidth = Math.max(...blockNames.map(name => name.length));
  var calculatedWidth = Math.min(Math.max(maxWidth * 10, 400), 800);
  var htmlOutput = template.evaluate()
    .setWidth(calculatedWidth)
    .setHeight(Math.min(100 + 60 * blockNames.length, 700));
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Выберите блоки для удаления');
}


function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();

  var blockSheets = getBlockSheetNames();

  if (blockSheets.has(sheetName)) {
    onBlockEdit(e)
  }
}