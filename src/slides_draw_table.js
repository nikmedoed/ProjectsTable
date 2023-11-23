function testSlideDrawing() {
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
  var rangeA1Notation = "A6:AB11";
  var range = sheet.getRange(rangeA1Notation);
  rangeToShape(range, shape)
}


const DEFAULT_FONT_SIZE = 14;
const DEFAULT_FONT = 'Arial';
const CELL_HEIGHT = DEFAULT_FONT_SIZE * 1.5
const DEFAULT_GLIPH_RATIO = DEFAULT_FONT_SIZE / 1.4


// Проблемы
// 1. Нет автофита текста
//  Сейчас никак не установить автоматический размер текста под шейп. При копирвоании настройка сбрасывается. Остаётся только сотни полей создавать, но такое себе.
// let cell = slide.insertTextBox(cellText, currentLeftOffset, currentTopOffset, cellWidth, CELL_HEIGHT)
// textStyle.setAutoSize(SlidesApp.AutoSize.TEXT_AUTOFIT);
// 2. Управление полями
// При создании новых жлементов невозможно установить настройки полей, не получается, но вот при копировании уже вышло
// Код, который как-то позволяет управлять полями, но не даёт изменить базовые поля
// let paragraphStyle = textRange.getParagraphStyle()
//   .setIndentStart(0)
//   .setIndentEnd(0)
//   .setIndentFirstLine(0);
// 3. Размеры стррок и колонок можно брать прямо из таблицы
// Но проблема в том, что это придётся соотносить со шрифтом, или и размеры шрифта переносить, но ещё надо учитывать масштабирование таблицы под слайд. В общем, проблемы надо порешать.

const horizontalAlignmentMap = {
  'left-0': SlidesApp.ParagraphAlignment.START,
  'center-0': SlidesApp.ParagraphAlignment.CENTER,
  'right-0': SlidesApp.ParagraphAlignment.END,
  // Поворот на 90 градусов
  'top-90': SlidesApp.ParagraphAlignment.END,
  'middle-90': SlidesApp.ParagraphAlignment.CENTER,
  'bottom-90': SlidesApp.ParagraphAlignment.START,
  // Поворот на 180 градусов
  'left-180': SlidesApp.ParagraphAlignment.END,
  'center-180': SlidesApp.ParagraphAlignment.CENTER,
  'right-180': SlidesApp.ParagraphAlignment.START,
  // Поворот на 270 градусов
  'top-270': SlidesApp.ParagraphAlignment.START,
  'middle-270': SlidesApp.ParagraphAlignment.CENTER,
  'bottom-270': SlidesApp.ParagraphAlignment.END,
};

const verticalAlignmentMap = {
  'top-0': SlidesApp.ContentAlignment.TOP,
  'middle-0': SlidesApp.ContentAlignment.MIDDLE,
  'bottom-0': SlidesApp.ContentAlignment.BOTTOM,
  // Поворот на 90 градусов
  'left-90': SlidesApp.ContentAlignment.BOTTOM,
  'center-90': SlidesApp.ContentAlignment.MIDDLE,
  'right-90': SlidesApp.ContentAlignment.TOP,
  // Поворот на 180 градусов
  'top-180': SlidesApp.ContentAlignment.BOTTOM,
  'middle-180': SlidesApp.ContentAlignment.MIDDLE,
  'bottom-180': SlidesApp.ContentAlignment.TOP,
  // Поворот на 270 градусов
  'left-270': SlidesApp.ContentAlignment.TOP,
  'center-270': SlidesApp.ContentAlignment.MIDDLE,
  'right-270': SlidesApp.ContentAlignment.BOTTOM,
};


function rangeToShape(range, shape) {
  const slide = shape.getParentPage().asSlide();
  var values = range.getDisplayValues();
  var backgrounds = range.getBackgrounds();
  var fontWeights = range.getFontWeights();
  var fontColors = range.getFontColors();
  var textRotations = range.getTextRotations().map(e => e.map(t => t.getDegrees()))
  var fontSizes = range.getFontSizes();
  var horizontalAlignments = range.getHorizontalAlignments();
  var verticalAlignments = range.getVerticalAlignments();

  var mergedRanges = range.getMergedRanges();
  var mergedRangeSizes = {};
  for (let mergedRange of mergedRanges) {
    let key = `${mergedRange.getRow() - range.getRow()}-${mergedRange.getColumn() - range.getColumn()}`;
    mergedRangeSizes[key] = {
      width: mergedRange.getLastColumn() - mergedRange.getColumn() + 1,
      height: mergedRange.getLastRow() - mergedRange.getRow() + 1
    };
  }

  // var widths = calculateColumnWidths(values, fontSizes, textRotations, mergedRangeSizes);
  // var heights = calculateRowHeight(values, fontSizes, textRotations, mergedRangeSizes)

  var sheet = range.getSheet();
  var widths = [];
  var heights = [];
  var rownum = range.getRow();
  var columnnum = range.getColumn();
  var rownum2 = range.getLastRow();
  var columnnum2 = range.getLastColumn();

  for (var i = columnnum; i <= columnnum2; i++) {
    widths.push(sheet.getColumnWidth(i)); // Заполнение массива ширинами колонок
  }
  for (var j = rownum; j <= rownum2; j++) {
    heights.push(sheet.getRowHeight(j)); // Заполнение массива высотами строк
  }
  var templateField = slide.getPageElements().find(element => {
    return element.asShape().getText().asString().includes("{{fieldTemplate}}");
  })

  var totalWidth = widths.reduce((accumulator, currentValue) => accumulator + currentValue, 0);
  const ratio = 1.2 * shape.getWidth() / totalWidth;

  var cells = []
  values.forEach((row, rowIndex) => {
    row.forEach((cellText, columnIndex) => {
      if (!cellText) { return }

      const rotation = textRotations[rowIndex][columnIndex]
      let hAlignment = horizontalAlignmentMap[`${horizontalAlignments[rowIndex][columnIndex]}-${rotation}`] || SlidesApp.ParagraphAlignment.START;
      let vAlignment = verticalAlignmentMap[`${verticalAlignments[rowIndex][columnIndex]}-${rotation}`] || SlidesApp.ContentAlignment.MIDDLE;

      let cellWidth, cellHeight
      let mergedSize = mergedRangeSizes[`${rowIndex}-${columnIndex}`]
      if (mergedSize) {
        cellWidth = widths.slice(columnIndex, columnIndex + mergedSize.width).reduce((a, b) => a + b, 0);
        cellHeight = heights.slice(rowIndex, rowIndex + mergedSize.height).reduce((a, b) => a + b, 0);
      } else {
        cellWidth = widths[columnIndex];
        cellHeight = heights[rowIndex];
      }
      let cellLeft = widths.slice(0, columnIndex).reduce((a, b) => a + b, 0);
      let cellTop = heights.slice(0, rowIndex).reduce((a, b) => a + b, 0)

      let cell = templateField.asShape().duplicate().asShape()
      if (rotation) {
        if (rotation == 90) {
          cellTop += cellHeight / 2 - cellWidth / 2
          cellLeft += cellWidth / 2 - cellHeight / 2;
        }
        if (rotation == 270) {
          cellLeft += cellWidth / 2 - cellHeight / 2;
        }
        [cellWidth, cellHeight] = [cellHeight, cellWidth];

        cell.setRotation(360 - rotation);
      }
      cell
        .setWidth(cellWidth)
        .setHeight(cellHeight)
        .setLeft(cellLeft)
        .setTop(cellTop)
        .setContentAlignment(vAlignment)
        .getFill().setSolidFill(backgrounds[rowIndex][columnIndex])

      cells.push(cell)

      let textRange = cell.getText().setText(cellText)
      let textStyle = textRange.getTextStyle()
        // .setFontSize(DEFAULT_FONT_SIZE)
        .setFontSize(fontSizes[rowIndex][columnIndex] * ratio)
        .setFontFamily(DEFAULT_FONT)
        .setBold(fontWeights[rowIndex][columnIndex] === 'bold')
        .setForegroundColor(fontColors[rowIndex][columnIndex])

      let paragraphStyle = textRange.getParagraphStyle()
      paragraphStyle.setParagraphAlignment(hAlignment)
    })
  })
  let group = slide.group(cells);
  groupToShapeProportional(group, shape)
  templateField.remove()
}


function groupToShapeProportional(group, shape) {
  const shapeWidth = shape.getWidth();
  const shapeTop = shape.getTop();
  const shapeLeft = shape.getLeft();

  const originalGroupHeight = group.getHeight();
  const originalGroupWidth = group.getWidth();
  const newGroupHeight = (shapeWidth / originalGroupWidth) * originalGroupHeight;

  group.setWidth(shapeWidth);
  group.setHeight(newGroupHeight);
  group.setLeft(shapeLeft);
  group.setTop(shapeTop);

  shape.remove();
}


function calculateCellWidth(cellText, fontSize, isRotated = False) {
  const padding = 8;
  if (isRotated) {
    return CELL_HEIGHT * cellText.split('\n').length + padding;
  } else {
    return cellText.length * DEFAULT_GLIPH_RATIO + padding;
  }
}


function calculateColumnWidths(values, fontSizes, textRotations, mergedRangeSizes) {
  let maxWidthPerColumn = new Array(values[0].length).fill(0);
  values.forEach((row, rowIndex) => {
    row.forEach((cellText, columnIndex) => {
      let cellWidth = 0
      if (!mergedRangeSizes[`${rowIndex}-${columnIndex}`]) {
        cellWidth = calculateCellWidth(cellText, fontSizes[rowIndex][columnIndex], textRotations[rowIndex][columnIndex]);
      }
      if (cellWidth > maxWidthPerColumn[columnIndex]) {
        maxWidthPerColumn[columnIndex] = cellWidth;
      }
    })
  })
  return maxWidthPerColumn;
}


function calculateRowHeight(values, fontSizes, textRotations, mergedRangeSizes) {
  return values.map((row, rowIndex) => Math.max(...row.map((cellText, columnIndex) => {
    let cellHeight = 0
    if (!mergedRangeSizes[`${rowIndex}-${columnIndex}`]) {
      cellHeight = calculateCellWidth(cellText, fontSizes[rowIndex][columnIndex], !textRotations[rowIndex][columnIndex]);
    }
    return cellHeight
  }))
  )
}
