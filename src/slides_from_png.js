


function imgURIsToShapes(data) {
  const presentation = SlidesApp.openById(data.presentationId);  
  let title = `${SLIDES_CONTENT_NAME_TAG} ${data.title}`

  for (let imageBlob of data.imageURIs) {
    let slide = duplicateSlideById(presentation, data.sheetTemplateSlideId)
    findShape(slide, SLIDES_CONTENT_NAME_TAG).getText().setText(title)
    const shape = findShape(slide, SLIDES_SHEET_ZONE_TAG);

  // var type = (imageBlob.split(";")[0]).replace('data:', '');
  // var imageUpload = Utilities.base64Decode(imageURI.split(",")[1]);
  // var imageBlob = Utilities.newBlob(imageUpload, type, `${presentationId}${shapeId}.png`);

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