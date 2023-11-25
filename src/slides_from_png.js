function dataToPlay(data) {
  const presentation = SlidesApp.openById(SLIDES_BASE_TEMPLATE);
  const slideWithShape = findSlideWithShape(presentation, SLIDES_SHEET_ZONE_TAG);
  const shape = findShape(slideWithShape, SLIDES_SHEET_ZONE_TAG);

  rangeToPDFblob({
    ...data[6].sheetPayload,
    sheetTemplateSlideId: slideWithShape.getObjectId(),
    shapeSize: [shape.getHeight(), shape.getWidth()]
  })
}


function rangeToPDFblob(data) {
  // Logger.log('rangeToPDFblob')
  // Logger.log(data)
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
  let sheetOnSlideHeight = ((height * sheetWidth) / width)
  let pdfBlobs = [];

  var currentHeight = frozenRowsHeight;
  var startRow = rownum > frozenRows ? rownum : frozenRows + 1;
  var rowPack = [];
  let hiddenRows = [];
  for (i = startRow; i <= rownum2; i++) {
    if (!sheet.isRowHiddenByUser(i) && !sheet.isRowHiddenByFilter(i)) {
      let rowHeight = sheet.getRowHeight(i);
      if (currentHeight + rowHeight > sheetOnSlideHeight) {
        // let blob = createPDFBlob(sheet.getSheetId(), rownum, columnnum, i - 1, columnnum2, sheetWidth, currentHeight);
        let blob = createPDFBlob(sheet.getSheetId(), rownum, columnnum, i - 1, columnnum2, sheetWidth, sheetOnSlideHeight);
        pdfBlobs.push(blob);
        rowPack.forEach(rowIndex => sheet.hideRows(rowIndex));
        hiddenRows.push(...rowPack);
        rowPack = [];
        currentHeight = frozenRowsHeight;
        startRow = i;
      }
      rowPack.push(i);
      currentHeight += rowHeight;
    }
  }

  if (rowPack.length > 0) {
    // let blob = createPDFBlob(sheet.getSheetId(), rownum, columnnum, rownum2, columnnum2, sheetWidth, currentHeight);
    let blob = createPDFBlob(sheet.getSheetId(), rownum, columnnum, rownum2, columnnum2, sheetWidth, sheetOnSlideHeight);
    pdfBlobs.push(blob);
  }
  hiddenRows.forEach(rowIndex => sheet.showRows(rowIndex));
  // Logger.log('rangeToPDFblob blobs: '+ pdfBlobs.length)

  return pdfBlobs
}


function createPDFBlob(sheetId, t, l, b, r, width, height) {
  var fileurl = SSheet.getUrl();
  var ratio = 96; // get inch from pixel 

  var exportUrl = fileurl.replace(/\/edit.*$/, '')
    + '/export?exportFormat=pdf&format=pdf'
    + '&size=' + [width, height].map(e => Number((e / ratio).toFixed(2))).join('x')
    //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
    // + '&portrait=true' //false= Landscape
    + '&scale=' + 4
    //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page     
    + '&top_margin=' + 0      //All four margins must be set!       
    + '&bottom_margin=' + 0
    + '&left_margin=' + 0
    + '&right_margin=' + 0
    + '&sheetnames=false&printtitle=false'
    + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
    + '&horizontal_alignment=CENTER' // //LEFT/CENTER/RIGHT
    + '&gridlines=false'
    + "&fmcmd=12"
    + '&fzr=FALSE'
    + '&gid=' + sheetId
    + `&r1=${t - 1}&r2=${b}&c1=${l - 1}&c2=${r}`

  Logger.log(exportUrl)

  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  })
  // Logger.log(response.getContentText());
  var blob = response.getBlob()
  var blobBytes = blob.getBytes();
  var blobString = Utilities.base64Encode(blobBytes);
  return blobString
}


function imgURIsToShapes(data) {
  let presentationId = data.presentationId
  const presentation = SlidesApp.openById(presentationId);
  let title = `${SLIDES_CONTENT_NAME_TAG} ${data.title}`
  for (let imageURI of data.imageURIs) {
    let slide = duplicateSlideById(presentation, data.sheetTemplateSlideId)
    findShape(slide, SLIDES_CONTENT_NAME_TAG).getText().setText(title)
    const shape = findShape(slide, SLIDES_SHEET_ZONE_TAG);

    var type = (imageURI.split(";")[0]).replace('data:', '');
    var imageUpload = Utilities.base64Decode(imageURI.split(",")[1]);
    var imageBlob = Utilities.newBlob(imageUpload, type, `${presentationId}${shape.getObjectId()}.png`);

    const shapeWidth = shape.getWidth();
    const shapeHeight = shape.getHeight();
    const shapeLeft = shape.getLeft();
    const shapeTop = shape.getTop();

    const img = slide.insertImage(imageBlob);
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
  }
}