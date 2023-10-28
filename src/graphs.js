// Более гибкая формула, чтобы захватить все данные на листе
// =FILTER(
//     INDIRECT("'Карта проекта'!H8:" & ADDRESS(ROWS('Карта проекта'!A:A), COLUMNS('Карта проекта'!1:1))),
//     NOT(ISBLANK('Карта проекта'!A8:A))
// )

function createDynamic(force = false) {
  const sheet = SSheet.getSheetByName(DYNAMIC_GRAPH);
  var dataRange = getActualDataRange(sheet);
  var charts = sheet.getCharts();

  var chartBuilder;
  if (charts.length > 0) {
    chartBuilder = charts[0].modify();
    for (var c = 1; c < charts.length; c++) {
      sheet.removeChart(charts[c]);
    }
    try {
      chartBuilder.removeRange(charts[0].getRanges()[0]);
    }
    catch { }
    chartBuilder.addRange(dataRange);
  } else {
    chartBuilder = sheet.newChart();
    force = true
  }

  if (force) {
    chartBuilder
      .asLineChart()
      .addRange(dataRange)
      .setTransposeRowsAndColumns(true)
      .setNumHeaders(1)
      .setOption('title', 'Динамика')
      .setOption('titlePosition', 'center')
      .setOption('vAxis.viewWindowMode', 'explicit')
      .setOption('vAxis.viewWindow.min', 0)
      .setOption('vAxis.viewWindow.max', 1)
      .setOption('vAxis.gridlines', { count: 11 })
      .setOption('legend', { position: 'bottom' })
      .setOption('treatLabelsAsText', true)
      .setOption('hAxis.gridlines', { count: dataRange.getNumColumns() - 1 })
      // .setOption('hAxis.slantedTextAngle', 90) // Поворот подписей на 90 градусов
      // .setOption('theme', 'maximized')
      .setOption('width', 1800)
      .setOption('height', 700)
  }

  chartBuilder
    .setOption('pointSize', 10)
    .setOption('useFirstColumnAsDomain', true)
    .setPosition(1, 1, 0, 0)

  // Не работает
  // chartBuilder = annotations(chartBuilder)

  if (charts.length > 0) {
    sheet.updateChart(chartBuilder.build());
  } else {
    sheet.insertChart(chartBuilder.build());
  }
}


function reloadGraphsForce() {
  reloadGraphs(true)
}

function reloadGraphs(force) {
  createDynamic(force)
  // drawStatus(force)
}

function getActualDataRange(sheet) {
  const values = sheet.getDataRange().getValues();

  let lastColumn = values[0].length;
  while (lastColumn > 0 && !values.some(row => row[lastColumn - 1])) {
    lastColumn--;
  }

  let lastRow = values.length;
  while (lastRow > 0 && !values[lastRow - 1].some(cell => cell)) {
    lastRow--;
  }

  return sheet.getRange(1, 1, lastRow, lastColumn);
}
