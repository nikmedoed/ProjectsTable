const SLIDES_BASE_TEMPLATE = "1Q7UJHX0h_dZZQRc9Kmo_eMDW_GFTx8_MeQVNCO1JDCs"


function getSlidesTemplateId() {
  // ID будет хранится в сторадже, будет инстефс для пересохранения. 
  // Т.е. нужно сделать копию шаблона, выдать ссылку, а потом использовать компию или предложить указать новую.

  return SLIDES_BASE_TEMPLATE
}

function getTemplateCopy(name) {
  var templateFile = DriveApp.getFileById(SLIDES_BASE_TEMPLATE);

  if (!name) {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
    name = templateFile.getName() + " – " + today;
  }
  var folder = DriveApp.getFileById(SSheet.getId()).getParents().next();
  var copiedFile = templateFile.makeCopy(name, folder);

  var fileUrl = copiedFile.getUrl();
  console.log(fileUrl);

  return copiedFile.getId();
}

function mapToSlide() {
  const presentationId = getSlidesTemplateId()
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

