function generateReport() {
  function anotation(sheetName) {
    let sheet = SSheet.getSheetByName(sheetName)
    return sheet.getDataRange().getA1Notation()
  }

  function getBlockDataRange(sheet) {
    sheet = SSheet.getSheetByName(sheet)
    let borders = findTimelineBorders(sheet)
    let last = sheet.getLastRow()
    const columnData = sheet.getRange(1, COLUMN_TASK, last, 1).getValues();
    last--
    while (!columnData[last][0]) {
      last--
    }
    return sheet.getRange(borders.row, 1, last - borders.row + 2, borders.col).getA1Notation()
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
    ["Комплектация вопросов и рисков", { title: "Вопросы и риски", sheet: TEMPLATE_MAP, range: risksRange }]
  ].map(e => {
    return e[1].range
      ? {
        name: e[0],
        func: 'imgURIsToShapes',
        sheetPayload: e[1]
      }
      : {
        name: e[0],
        func: 'chartToSlides',
        sheetPayload: e[1]
      }
  })
  let data = [
    { name: 'Подготовка шаблона', func: 'prepareTemplate' },
    ...sheets,
    { name: 'Генерация оглавления', func: 'contentTableGenerator' },
  ]

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

function dataToPlay(data) {
  const presentation = SlidesApp.openById(SLIDES_BASE_TEMPLATE);
  const slideWithShape = findSlideWithShape(presentation, SLIDES_SHEET_ZONE_TAG);
  const shape = findShape(slideWithShape, SLIDES_SHEET_ZONE_TAG);

  rangeToPDFblob({
    ...data[5].sheetPayload,
    sheetTemplateSlideId: slideWithShape.getObjectId(),
    shapeSize: [shape.getHeight(), shape.getWidth()]
  })
}

function rangeToPDFblob(data) {
  let [height, width] = data.shapeSize
  let sheet = SSheet.getSheetByName(data.sheet)
  let range = sheet.getRange(data.range)
  let frozenRows = sheet.getFrozenRows();

  var columnnum = range.getColumn();
  var columnnum2 = range.getLastColumn();
  var rownum = range.getRow();
  var rownum2 = range.getLastRow();

  let i
  var sheetWidth = 0;
  for (i = columnnum; i <= columnnum2; i++) {
    if (!sheet.isColumnHiddenByUser(i)) {
      sheetWidth += sheet.getColumnWidth(i);
    }
  }

  var frozenRowsHeight = 0;
  for (i = rownum; i <= frozenRows; i++) {
    if (!sheet.isRowHiddenByUser(i) && !sheet.isRowHiddenByFilter(i)) {
      frozenRowsHeight += sheet.getRowHeight(i);
    }
  }
  let slideHeight = ((height * sheetWidth) / width)
  let pdfBlobs = [];

  var currentHeight = frozenRowsHeight;
  var startRow = frozenRows + 1;
  var rowPack = [];
  for (i = startRow; i <= rownum2; i++) {
    if (!sheet.isRowHiddenByUser(i) && !sheet.isRowHiddenByFilter(i)) {
      let rowHeight = sheet.getRowHeight(i);
      if (currentHeight + rowHeight > slideHeight) {
        let blob = createPDFBlob(sheet.getSheetId(), rownum, columnnum, i - 1, columnnum2, sheetWidth, currentHeight);
        pdfBlobs.push(blob);
        rowPack.forEach(rowIndex => sheet.hideRows(rowIndex));
        rowPack = [];
        currentHeight = frozenRowsHeight;
        startRow = i;
      }
      rowPack.push(i);
      currentHeight += rowHeight;
    }
  }

  if (rowPack.length > 0) {
    let blob = createPDFBlob(sheet.getSheetId(), rownum, columnnum, rownum2, columnnum2, sheetWidth, currentHeight);
    pdfBlobs.push(blob);
  }

  return pdfBlobs
}


function createPDFBlob(sheetId, l, t, b, r, width, height) {
  var fileurl = SSheet.getUrl();
  var ratio = 96; // get inch from pixel 

  var exportUrl = fileurl.replace(/\/edit.*$/, '')
    + '/export?exportFormat=pdf&format=pdf'
    + '&size=' + [width, height].map(e => Number((e / ratio).toFixed(2))).join('x')
    //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
    // + '&portrait=true' //false= Landscape
    + '&scale=' + 2
    //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page     
    + '&top_margin=' + 0      //All four margins must be set!       
    + '&bottom_margin=' + 0
    + '&left_margin=' + 0
    + '&right_margin=' + 0
    + '&sheetnames=false&printtitle=false'
    + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
    + 'horizontal_alignment=LEFT' // //LEFT/CENTER/RIGHT
    + '&gridlines=false'
    + "&fmcmd=12"
    + '&fzr=FALSE'
    + '&gid=' + sheetId
    + `&r1=${t - 1}&r2=${b}&c1=${l - 1}&c2=${r}`

  Logger.log(exportUrl)

  var blob = UrlFetchApp.fetch(
    exportUrl,
    { headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() } }
  ).getBlob();
  return blob;
  // var blobBytes = blob.getBytes();
  // var blobString = Utilities.base64Encode(blobBytes);
  // return blobString
}


function projectMapSplit() {
  let projectMap = SSheet.getSheetByName(TEMPLATE_MAP)
  var mapValues = projectMap.getDataRange().getValues();
  var lastRow = 0;
  for (var i = mapValues.length - 1; i >= 0; i--) {
    if (mapValues[i][0]) {
      lastRow = i + 1;
      var mapRange = projectMap.getRange(1, 1, lastRow, projectMap.getLastColumn()).getA1Notation()
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

  rig += Math.max(...projectMap.getRange(top, rig, bot - top + 1, 1).getMergedRanges().map(e => e.getValue() ? e.getWidth() - 1 : 0))
  bot += Math.max(...projectMap.getRange(bot, left, 1, rig - left + 1).getMergedRanges().map(e => e.getValue() ? e.getHeight() - 1 : 0))

  var risksRange = projectMap.getRange(top, left, bot - top + 1, rig - left + 1).getA1Notation();
  return [mapRange, risksRange]
}


function chartToSlides(data) {
  const presentation = SlidesApp.openById(data.presentationId);
  let slide = duplicateSlideById(presentation, data.sheetTemplateSlideId)
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
  name = `Отчёт - ${today} - ${data.sheetName}`
  var copiedFile = getSlidesCopy(getPresentationId(), name)
  data.reportUrl = copiedFile.getUrl()
  data.presentationId = copiedFile.getId()
  data.slidesName = name

  // Находим слайд шаблон для вставки данных из таблицы
  const presentation = SlidesApp.openById(data.presentationId);
  const slideWithShape = findSlideWithShape(presentation, SLIDES_SHEET_ZONE_TAG);
  data.sheetTemplateSlideId = slideWithShape.getObjectId()

  const shape = findShape(slide, SLIDES_SHEET_ZONE_TAG);
  data.shapeSize = [shape.getHeight(), shape.getWidth()]

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