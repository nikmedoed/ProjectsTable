function generateReport() {
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
  var rangeA1Notation = "A1:AB16";
  var range = sheet.getRange(rangeA1Notation);

  let imageBlob = rangeToBlob(range)
  let shapeId = shape.getObjectId()

  openRender(presentationId, shapeId, imageBlob)
}


function openRender(presentationId, shapeId, imageBlobBase64) {
  var template = HtmlService.createTemplateFromFile('Report');
  template.presentationId = presentationId;
  template.shapeId = shapeId;
  template.imageBlobBase64 = imageBlobBase64;
  var html = template.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.NATIVE)
  SpreadsheetApp.getUi().showModalDialog(html, 'Converter');
}

function imageToShape(imageURI, presentationId, shapeId) {
  const presentation = SlidesApp.openById(presentationId);
  const page = presentation.getPageElementById(shapeId).getParentPage();
  const shape = page.getPageElementById(shapeId).asShape();

  var type = (imageURI.split(";")[0]).replace('data:', '');
  var imageUpload = Utilities.base64Decode(imageURI.split(",")[1]);
  var imageBlob = Utilities.newBlob(imageUpload, type, `${presentationId}${shapeId}.png`);

  const shapeWidth = shape.getWidth();
  const shapeHeight = shape.getHeight();
  const shapeLeft = shape.getLeft();
  const shapeTop = shape.getTop();

  const img = page.insertImage(imageBlob);
  const imgWidth = img.getWidth();
  const imgHeight = img.getHeight();
  const imgRatio = imgWidth / imgHeight;
  let newImgWidth, newImgHeight, newImgLeft, newImgTop;
  newImgWidth = shapeWidth;
  newImgHeight = newImgWidth / imgRatio;
  if (newImgHeight > shapeHeight) {
    newImgHeight = shapeHeight;
    newImgWidth = newImgHeight * imgRatio;
  }
  newImgLeft = shapeLeft + (shapeWidth - newImgWidth) / 2;
  newImgTop = shapeTop;

  img.setWidth(newImgWidth).setHeight(newImgHeight).setLeft(newImgLeft).setTop(newImgTop);
  shape.remove();
  return "OK, next";
}

function rangeToBlob(range) {
  let exportUrl = rangeToPDFurl(range)
  var blob = UrlFetchApp.fetch(
    exportUrl,
    { headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() } }
  ).getBlob();
  var blobBytes = blob.getBytes();
  var blobString = Utilities.base64Encode(blobBytes);
  return blobString
}


// TODO не учитывать скрытые строки и колонки в расчёте размера
// TODO точнее поработать с размерами, а то не попала
function rangeToPDFurl(range) {
  var ratio = 96; // get inch from pixel

  range = range || SpreadsheetApp.getActiveRange();
  var sheet = range.getSheet();
  var file = SpreadsheetApp.getActive();

  var fileurl = file.getUrl();
  var sheetid = sheet.getSheetId();
  var rownum = range.getRow();
  var columnnum = range.getColumn();
  var rownum2 = range.getLastRow();
  var columnnum2 = range.getLastColumn();

  var w = 0;
  for (var i = columnnum; i <= columnnum2; i++) {
    if (!sheet.isColumnHiddenByUser(i)) {
      w += sheet.getColumnWidth(i);
    }
  }

  var h = 0;
  for (var i = rownum; i <= rownum2; i++) {
    if (!sheet.isRowHiddenByUser(i) && !sheet.isRowHiddenByFilter(i)) {
      h += sheet.getRowHeight(i);
    }
  }

  hh = Number((h / ratio).toFixed(2));
  ww = Number((w / ratio).toFixed(2));

  var sets = {
    url: fileurl,
    sheetId: sheetid,
    r1: rownum - 1,
    r2: rownum2,
    c1: columnnum - 1,
    c2: columnnum2,
    size: ww + 'x' + hh,
    //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
    // portrait: true, //true= Potrait / false= Landscape
    scale: 2,          //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page
    top_margin: 0,     //All four margins must be set!        
    bottom_margin: 0,  //All four margins must be set!       
    left_margin: 0,    //All four margins must be set!         
    right_margin: 0,   //All four margins must be set!
  }
  var rangeParam =
    '&r1=' + sets.r1 +
    '&r2=' + sets.r2 +
    '&c1=' + sets.c1 +
    '&c2=' + sets.c2;
  var sheetParam = '&gid=' + sets.sheetId;
  var isPortrait = '';
  if (sets.portrait) {
    //true= Potrait / false= Landscape
    isPortrait = '&portrait=' + sets.portrait;
  }
  var exportUrl = sets.url.replace(/\/edit.*$/, '')
    + '/export?exportFormat=pdf&format=pdf'
    + '&size=' + sets.size             //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
    + isPortrait
    + '&scale=' + sets.scale            //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page     
    + '&top_margin=' + sets.top_margin       //All four margins must be set!       
    + '&bottom_margin=' + sets.bottom_margin    //All four margins must be set!     
    + '&left_margin=' + sets.left_margin      //All four margins must be set! 
    + '&right_margin=' + sets.right_margin     //All four margins must be set!     
    + '&sheetnames=false&printtitle=false'
    + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
    + 'horizontal_alignment=LEFT' // //LEFT/CENTER/RIGHT
    + '&gridlines=false'
    + "&fmcmd=12"
    + '&fzr=FALSE'
    + sheetParam
    + rangeParam;
  return exportUrl;
}