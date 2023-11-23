function generateReport() {
  const template = HtmlService.createTemplateFromFile('Report');
  template.steps = JSON.stringify([
    { name: 'Подготовка шаблона', func: 'prepareTemplate' },
    // {name: 'Подготовка шаблона', func: 'prepareTamplate'},
    //   {name: 'Подготовка шаблона', func: 'prepareTamplate'},
    { name: 'Генерация оглавления', func: 'contentTableGenerator' },
  ])

  SpreadsheetApp.getUi().showModalDialog(
    template
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.NATIVE)
      .setHeight(200),
    'Генерация отчёта'
  );
}


function mapToSlide() {

  var sheet = SSheet.getSheetByName(TEMPLATE_MAP);
  var rangeA1Notation = "A6:Q11";
  var range = sheet.getRange(rangeA1Notation);

  rangeToShape(range, shape)
}

function prepareTemplate() {
  let data = collectTemplateValues()

  // Сделали копию шаблона отчёта
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  name = `Отчёт - ${today} - ${data.sheetName}`
  var copiedFile = getSlidesCopy(getPresentationId(), name)
  data.reportUrl = copiedFile.getUrl()
  data.presentationId = copiedFile.getId()
  data.slidesName = name

  // Находим слайд шаблон для вставки данных из таблицы
  const presentation = SlidesApp.openById(data.presentationId);
  const slideWithShape = findSlideWithShape(presentation, SLIDES_SHEET_ZONE_TAG);
  data.sheetTemplateSlideId = slideWithShape.getObjectId()

  // Заменяем теги на значения
  replaceTemplateKeys(presentation, data);
  return data
}


function replaceTemplateKeys(presentation, data) {
  for (let key in data) {
    const placeholder = `{{${key}}}`;
    const replacement = data[key];
    presentation.replaceAllText(placeholder, replacement);
  }
}


function findShape(page, searchText) {
  let shapes = page.getShapes();
  for (let shape of shapes) {
    if (shape.getText().asString().includes(searchText)) {
      return shape;
    }
  }
  return null;
}

function findSlideWithShape(presentation, searchText) {
  const slides = presentation.getSlides();
  for (let slide of slides) {
    let shapes = slide.getShapes();
    for (let shape of shapes) {
      if (shape.getText().asString().includes(searchText)) {
        return slide;
      }
    }
  }
  throw (`Не найден блок с текстом ${searchText}`)
}



function collectTemplateValues() {
  var values = SSheet.getSheetByName(TEMPLATE_MAP).getRange("B:C").getValues()
  var projectName = ""
  for (var i = 0; i < values.length; i++) {
    if (values[i][0].toString().includes("Проект")) {
      projectName = values[i][1];
      break
    }
  }
  return {
    sheetName: SSheet.getName(),
    reportDate: formatDate(new Date()),
    projectName: projectName
  }
}


function formatDate(date) {
  const months = [
    'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
    'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'
  ];

  const day = date.getDate();
  const monthIndex = date.getMonth();
  const year = date.getFullYear();

  return `${day} ${months[monthIndex]} ${year} г.`;
}


function slidesTemplateLink() {
  let url = getPresentationTemplateLink()
  var htmlTemplate = HtmlService.createTemplateFromFile('SlidesTemplateLink');
  htmlTemplate.url = url;
  var userInterface = htmlTemplate.evaluate()
    .setWidth(580)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(userInterface, "Ссылка на шаблон отчёта");
}


function getSlidesCopy(slidesId, name) {
  var templateFile = DriveApp.getFileById(slidesId);
  var folder = DriveApp.getFileById(SSheet.getId()).getParents().next();
  return templateFile.makeCopy(name, folder);
}


function makeTemplateSlidesCopy(name) {
  name = `Шаблон отчёта по проекту "${name}"`
  var copiedFile = getSlidesCopy(SLIDES_BASE_TEMPLATE, name)
  storePresentationIdOrLink(copiedFile.getId())
  slidesTemplateLink();
}