function contentTableGenerator(data) {
  switchReport(false)
  let presentation = SlidesApp.openById(data.presentationId);
  presentation.getSlideById(data.sheetTemplateSlideId).remove()

  let contents = collectSlidesTitles(presentation)
  const contnentMatrix = createMatrix(contents);

  const slides = presentation.getSlides();
  slides.forEach(slide => {
    const layout = slide.getLayout();
    const contentTable = findShape(layout, SLIDES_CONTENT_TABLE_TAG);
    if (contentTable) {
      const group = slide.getGroups().find(g => g.getChildren()[0].asShape().getText().asString().includes(SLIDES_CONTENT_FILED_TAG));
      if (group) {
        const shapes = collectShapesInMatrix(group)
        distributeWordsInShapes(contnentMatrix, shapes, slide)
        groupToShape(group, contentTable)
      }
    }
  });

  presentation.getLayouts().forEach(layout => {
    const shape = findShape(layout, SLIDES_CONTENT_TABLE_TAG);
    if (shape) { shape.remove(); }
  });
  return data
}


function distributeWordsInShapes(wordsMatrix, shapesMatrix, currentSlide) {
  for (let i = 0; i < shapesMatrix.length; i++) {
    for (let j = 0; j < shapesMatrix[i].length; j++) {
      const shape = shapesMatrix[i][j];
      if (i < wordsMatrix.length && j < wordsMatrix[i].length) {
        let contentItem = wordsMatrix[i][j]
        shape.getText().setText(contentItem.title);
        const textStyle = shape.getText().getTextStyle();
        let targetSlide = contentItem.slideGroupList[0]
        shape.setLinkSlide(targetSlide)
        textStyle.setLinkSlide(targetSlide)
        if (contentItem.slideIdSet.has(currentSlide.getObjectId())) {
          textStyle.setForegroundColor(SlidesApp.ThemeColorType.DARK1);
          shape.getBorder().setWeight(1).getLineFill().setSolidFill(SlidesApp.ThemeColorType.DARK1);
          shape.getFill().setSolidFill(SlidesApp.ThemeColorType.ACCENT2);
        } else {
          textStyle.setForegroundColor(SlidesApp.ThemeColorType.DARK2);
          shape.getBorder().setWeight(1).getLineFill().setSolidFill(SlidesApp.ThemeColorType.DARK2);
        }
      } else {
        shape.remove();
        shapesMatrix[i][j] = null;
      }
    }
  }
  if (wordsMatrix.length > 1) {
    const lastRowIndex = wordsMatrix.length - 1;
    const lastRow = shapesMatrix[lastRowIndex];
    const upperRow = shapesMatrix[lastRowIndex - 1];

    if (wordsMatrix[lastRowIndex].length < wordsMatrix[lastRowIndex - 1].length) {
      const totalWidthUpperRow = upperRow.reduce((sum, shape) => shape ? sum + shape.getWidth() : sum, 0);
      const totalWidthLastRow = lastRow.reduce((sum, shape) => shape ? sum + shape.getWidth() : sum, 0);
      const stretchFactor = totalWidthUpperRow / totalWidthLastRow;

      let currentLeft = lastRow[0].getLeft();
      for (const shape of lastRow) {
        if (shape) {
          const newWidth = shape.getWidth() * stretchFactor;
          shape.setWidth(newWidth);
          shape.setLeft(currentLeft);
          currentLeft += newWidth;
        }
      }
    }
  }
}


function groupToShape(group, shape) {
  const tableTop = shape.getTop();
  const tableLeft = shape.getLeft();
  const tableWidth = shape.getWidth();
  const tableHeight = shape.getHeight();
  group.setTop(tableTop);
  group.setLeft(tableLeft);
  group.setWidth(tableWidth);
  group.setHeight(tableHeight);
}


function collectSlidesTitles(presentation) {
  let slidesInfo = [];
  const slides = presentation.getSlides();
  slides.forEach(slide => {
    slide.getShapes().forEach(shape => {
      if (shape.getText) {
        const text = shape.getText().asString();
        if (text.startsWith(SLIDES_CONTENT_NAME_TAG)) {
          shape.remove()
          const title = text.replace(SLIDES_CONTENT_NAME_TAG, "").trim();
          if (!title) return
          const slideId = slide.getObjectId()
          if (slidesInfo.length > 0 && slidesInfo[slidesInfo.length - 1].title === title) {
            let item = slidesInfo[slidesInfo.length - 1]
            item.slideGroupList.push(slide);
            item.slideIdSet.add(slideId)
          } else {
            slidesInfo.push({
              title: title,
              slideGroupList: [slide],
              slideIdSet: new Set([slideId])
            });
          }
        }
      }
    });
  });
  return slidesInfo;
}


function collectShapesInMatrix(group) {
  const tolerance = 10;
  let shapes = group.getChildren().map(c => c.asShape())
  shapes.sort((a, b) => a.getTop() - b.getTop());

  const rows = [];
  let currentRow = [];
  let lastY = shapes[0].getTop();

  for (const shape of shapes) {
    if (Math.abs(shape.getTop() - lastY) > tolerance) {
      rows.push(currentRow);
      currentRow = [];
      lastY = shape.getTop();
    }
    currentRow.push(shape);
  }
  if (currentRow.length > 0) {
    rows.push(currentRow);
  }
  for (const row of rows) {
    row.sort((a, b) => a.getLeft() - b.getLeft());
  }
  return rows;
}


function createMatrix(words) {
  const maxRows = 3;
  const maxCols = words.length > 24 ? 10 : 8;
  let numRows = 1;
  let numCols = words.length;
  while (numRows <= maxRows && numCols > maxCols) {
    numRows++;
    numCols = Math.ceil(words.length / numRows);
  }
  while (numRows > 1 && numCols * (numRows - 1) >= words.length) {
    numRows--;
    numCols = Math.ceil(words.length / numRows);
  }
  const matrix = [];
  for (let i = 0; i < numRows; i++) {
    matrix.push(words.slice(i * numCols, (i + 1) * numCols));
  }
  return matrix;
}