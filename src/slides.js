function generateReport() {
  function anotation(sheetName) {
    let sheet = SSheet.getSheetByName(sheetName);
    let lastRow = sheet.getLastRow();
    let lastColumn = sheet.getLastColumn();
    let data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    while (lastRow > 0 && data[lastRow - 1].every(cell => cell.toString().trim() === '')) {
      lastRow--;
    }
    return sheet.getRange(1, 1, lastRow, lastColumn).getA1Notation();
  }

  function getBlockDataRange(sheetName) {
    sheet = SSheet.getSheetByName(sheetName)
    let borders = findTimelineBorders(sheet)
    let lastVisible = getLastVisibleTimelineColumn(sheet, borders.startCol, borders.col)
    let last = sheet.getLastRow()
    fixBordersOnRange(sheet.getRange(ROW_START, 1, last - ROW_START + 1, sheet.getLastColumn()), borders)
    const columnData = sheet.getRange(1, COLUMN_TASK, last, 1).getValues();
    last--
    while (!columnData[last][0]) {
      last--
    }
    return sheet.getRange(borders.row, 2, last - borders.row + 2, lastVisible - 1).getA1Notation()
  }

  var [mapRange, risksRange] = projectMapSplit()
  let blocks = getBlockSheets().map(sheet => sheet.getName()).map(e => [`Обрабработка страницы "${e}"`, { title: e, sheet: e, range: getBlockDataRange(e) }])
  let sheets = [
    ["Сбор команды", { title: "Команда проекта", sheet: COMMANDO, range: anotation(COMMANDO) }],
    ["Отрисовка динамики", { title: DYNAMIC_GRAPH, sheet: DYNAMIC_GRAPH }],
    ["Визуализация статуса", { title: STATUS_GRAPH, sheet: STATUS_GRAPH }],
    ["Сбор текущих задач", { title: TEMPLORARY, sheet: TEMPLORARY, range: anotation(TEMPLORARY) }],
    ["Анализ карты проекта", { title: TEMPLATE_MAP, sheet: TEMPLATE_MAP, range: mapRange }],
    ...blocks,
  ]
  if (risksRange) {
    sheets.push(["Комплектация вопросов и рисков", { title: "Вопросы и риски", sheet: TEMPLATE_MAP, range: risksRange }])
  }
  sheets = sheets.map(e => {
    return e[1].range
      ? {
        name: e[0],
        func: 'imgURIsToShapes',
        sheetPayload: e[1]
      }
      : {
        name: e[0],
        func: 'chartToSlides',
        payload: e[1]
      }
  })
  let data = [
    { name: 'Подготовка шаблона', func: 'prepareTemplate' },
    ...sheets,
    { name: 'Генерация оглавления', func: 'contentTableGenerator' },
  ]

  // Logger.log(data)
  // dataToPlay(data)
  // return 

  const template = HtmlService.createTemplateFromFile('Report');
  template.steps = JSON.stringify(data)

  SpreadsheetApp.getUi().showModalDialog(
    template
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.NATIVE)
      .setHeight(200),
    'Генерация отчёта'
  );
}


function getLastVisibleTimelineColumn(sheet, startColumn, endColumn) {
  let visibleCount = 0;
  let lastVisibleCol = 0;

  for (let col = startColumn; col <= endColumn; col++) {
    if (!sheet.isColumnHiddenByUser(col)) {
      visibleCount++;
      lastVisibleCol = col;
    }
  }
  if (visibleCount > 25) {
    let message = 'На листе "' + sheet.getName() + '" отображаются более 25 колонок таймлайна';
    SpreadsheetApp.getUi().alert(message + ", что запрещено. Скройте неактуальные колонки таймлайна.");
    throw new Error(message);
  }
  return lastVisibleCol;
}

function projectMapSplit() {
  let projectMap = SSheet.getSheetByName(TEMPLATE_MAP)
  var mapValues = projectMap.getDataRange().getValues();
  var lastRow = 0;
  var mapRange = ""
  for (var i = mapValues.length - 1; i >= 0; i--) {
    if (mapValues[i][0]) {
      lastRow = i + 1;
      let timeline = findTimelineBorders(projectMap)
      let lastVisible = getLastVisibleTimelineColumn(projectMap, timeline.startCol, timeline.col)
      mapRange = projectMap.getRange(1, 1, lastRow, lastVisible).getA1Notation()
      fixBordersOnRange(projectMap.getRange(timeline.endRow + 1, 1, lastRow - timeline.endRow, projectMap.getLastColumn()), timeline)
      break;
    }
  }

  var top = 0, bot = 0;
  for (var i = lastRow; i < mapValues.length; i++) {
    if (mapValues[i].some(cell => cell !== "")) {
      if (!top) top = i + 1;
      bot = i + 1;
    }
  }

  var left = mapValues[0].length, rig = 0;
  for (var col = 0; col < mapValues[0].length; col++) {
    var columnData = mapValues.slice(top - 1, bot).map(row => row[col]);
    if (columnData.some(cell => cell !== "")) {
      left = Math.min(left, col + 1);
      rig = Math.max(rig, col + 1);
    }
  }
  var risksRange = ""
  if (bot > top) {
    rig += Math.max(...projectMap.getRange(top, rig, bot - top + 1, 1).getMergedRanges().map(e => e.getValue() ? e.getWidth() - 1 : 0))
    bot += Math.max(...projectMap.getRange(bot, left, 1, rig - left + 1).getMergedRanges().map(e => e.getValue() ? e.getHeight() - 1 : 0))

    var risksRange = projectMap.getRange(top, left, bot - top + 1, rig - left + 1).getA1Notation();
  }

  return [mapRange, risksRange]
}


function chartToSlides(data) {
  // Logger.log('chartToSlides')
  // Logger.log(data)

  const presentation = SlidesApp.openById(data.presentationId);
  let slide = duplicateSlideById(presentation, data.sheetTemplateSlideId)
  let title = `${SLIDES_CONTENT_NAME_TAG} ${data.title}`
  findShape(slide, SLIDES_CONTENT_NAME_TAG).getText().setText(title)

  const shape = findShape(slide, SLIDES_SHEET_ZONE_TAG);
  let sheet = SSheet.getSheetByName(data.sheet)
  let chart = sheet.getCharts()[0]

  const originalHeight = chart.getOptions().get('height')
  const originalWidth = chart.getOptions().get('width')
  const width = shape.getWidth();
  const height = shape.getHeight();
  const left = shape.getLeft();
  const top = shape.getTop();

  chart = chart.modify()
    .setOption('width', width)
    .setOption('height', height)
    .build();
  sheet.updateChart(chart);

  slide.insertSheetsChartAsImage(chart, left, top, width, height);

  chart = sheet.getCharts()[0]
  chart = chart.modify()
    .setOption('width', originalWidth)
    .setOption('height', originalHeight)
    .build();
  sheet.updateChart(chart);

  shape.remove();
}


function duplicateSlideById(presentation, slideId) {
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; i++) {
    if (slides[i].getObjectId() === slideId) {
      var duplicatedSlide = slides[i].duplicate();
      duplicatedSlide.move(i);
      return duplicatedSlide
    }
  }
  throw "Слайд с ID '" + slideId + "' (шаблон) не найден"
}


function prepareTemplate() {
  let data = collectTemplateValues()

  // Сделали копию шаблона отчёта
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  let name = `Отчёт - ${today} - ${data.sheetName}`
  var copiedFile = getSlidesCopy(getPresentationId(), name)
  data.reportUrl = copiedFile.getUrl()
  data.presentationId = copiedFile.getId()
  data.slidesName = name

  // Находим слайд шаблон для вставки данных из таблицы
  const presentation = SlidesApp.openById(data.presentationId);
  const slideWithShape = findSlideWithShape(presentation, SLIDES_SHEET_ZONE_TAG);
  data.sheetTemplateSlideId = slideWithShape.getObjectId()

  const shape = findShape(slideWithShape, SLIDES_SHEET_ZONE_TAG);
  data.shapeSize = [shape.getHeight(), shape.getWidth()]

  // Заменяем теги на значения
  replaceTemplateKeys(presentation, data);
  switchReport(true)
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