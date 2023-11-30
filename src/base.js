function toRelease() {
  let FROM = "Карта проекта";
  let TO = "!КПШ";
  var sheets = SSheet.getSheets();

  var regex = new RegExp("'" + FROM + "'", "g");
  let tor = "'" + TO + "'";

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var range = sheet.getDataRange();
    var formulas = range.getFormulas();

    for (var j = 0; j < formulas.length; j++) {
      for (var k = 0; k < formulas[j].length; k++) {
        if (formulas[j][k]) {
          if (formulas[j][k].indexOf("INDIRECT") === -1) {
            formulas[j][k] = formulas[j][k].replace(regex, tor);
            sheet.getRange(j + 1, k + 1).setFormula(formulas[j][k]);
          }
        }
      }
    }
  }

  SSheet.deleteSheet(SSheet.getSheetByName(FROM));
  SSheet.getSheetByName(TO).setName(FROM);

  getBlockSheets().forEach(e => SSheet.deleteSheet(e))

  getHided().forEach(e => SSheet.getSheetByName(e).hideSheet())
}


function onOpen() {
  let ui = SpreadsheetApp.getUi()
  ui
    .createMenu('💎 Управление проектом ')
    .addItem('📄 Сгенерировать отчёт', 'generateReport')
    .addItem('🔀 Переключить режим отчёта', 'switchReport')
    .addItem('➕ Добавить новый блок', 'createNewBlockPrompt')
    .addItem('📈 Расширить таймлайн', 'showExtendTimeline')
    .addItem('❌ Удалить блок(и)', 'showDeleteBlocks')
    .addSubMenu(ui
      .createMenu('🔧 Починить')
      .addItem('Пересобрать графики принудительно', 'reloadGraphsForce')
      .addItem('Обновить формулу текущих задач', 'temploraryTasksFormula')
      .addItem('Поправить описание задач и формулы', 'bloksDataFix')
      .addItem('Поправить границы ячеек блоков', 'fixBorders')
    )
    .addItem('🔗 Ссылка на шаблон отчёта', 'slidesTemplateLink')
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

  if (!startDate) {
    let timeline = findTimelineBorders(SSheet.getSheetByName(TEMPLATE_MAP))
    if (!timeline.isNew) {
      startDate = timeline.startDate
      endDate = timeline.endDate
    }
  }

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