function getSlidesTemplateId() {
  // ID будет хранится в сторадже, будет инстефс для пересохранения. 
  // Т.е. нужно сделать копию шаблона, выдать ссылку, а потом использовать компию или предложить указать новую.

  return SLIDES_BASE_TEMPLATE
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


function getTemplateCopyFprReport(name) {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  name = `Отчёт - ${today} - ${name}`
  var copiedFile = getSlidesCopy(getPresentationId(), name)

  var fileUrl = copiedFile.getUrl();
  var fileId = copiedFile.getId();
  return
}

function mapToSlide() {
  const presentationId = getPresentationId()
  const presentation = SlidesApp.openById(presentationId);
  const slide = presentation.getSlides()[0].duplicate();

  const shapes = slide.getShapes();
  let shape
  for (shape of shapes) {
    if (shape.getText().asString().includes('{{sheet}}')) {
      break
    }
  }

  var sheet = SSheet.getSheetByName(TEMPLATE_MAP);
  var rangeA1Notation = "A6:Q11";
  var range = sheet.getRange(rangeA1Notation);

  rangeToShape(range, shape)
}

