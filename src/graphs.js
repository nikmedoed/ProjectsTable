function createCharts() {
  const sheet = SSheet.getSheetByName(DYNAMIC_GRAPH);
  var dataRange = getActualDataRange(sheet);
  var charts = sheet.getCharts();
  for (var c = 0; c < charts.length; c++) {
    sheet.removeChart(charts[c]);
  }

  var chartBuilder = sheet.newChart();
  chartBuilder
    // .setChartType(Charts.ChartType.LINE)
    .asLineChart()
    .addRange(dataRange)
    .setTransposeRowsAndColumns(true)
    .setNumHeaders(1)
    .setOption('title', 'Динамика')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('vAxis.viewWindowMode', 'explicit')
    .setOption('vAxis.viewWindow.min', 0)
    .setOption('vAxis.viewWindow.max', 1)
    .setOption('vAxis.gridlines', { count: 11 })
    .setOption('legend', { position: 'bottom' })
    .setOption('pointSize', 10)

    // .setOption('treatLabelsAsText', true)   
    .setOption('hAxis.gridlines', { count: dataRange.getNumColumns() - 1 })

    // .setOption('theme', 'maximized')
    .setPosition(1, 1, 0, 0)
    .setOption('width', 1800)
    .setOption('height', 700);

  chartBuilder = annotations(chartBuilder)

  sheet.insertChart(chartBuilder.build());
}

function annotations(chart) {

  return chart
    .setOption('annotations', {
      alwaysOutside: true,
      textStyle: {
        fontSize: 16,
        color: 'red'
      }
    })
    .setOption('annotations.domain', true)
    .setOption('series', {
      0: {
        annotations: {
          textStyle: { fontSize: 12, color: 'red' }
        }
      },
      1: {
        annotations: {
          textStyle: { fontSize: 12, color: 'blue' }
        }
      }
    })
    .setOption('series', {
      0: {
        annotations: {
          textStyle: { fontSize: 12, color: 'red' },
          alwaysOutside: true,
          boxStyle: {
            stroke: '#999',
            strokeWidth: 1,
            gradient: {
              color1: '#fbf6a7',
              color2: '#33b679',
              x1: '0%', y1: '0%',
              x2: '100%', y2: '100%',
              useObjectBoundingBoxCoordinates: true
            }
          }
        }
      }
    })
}

function test() {
  var graph = SSheet.getSheetByName(DYNAMIC_GRAPH)
  var chart = graph.getCharts()[0];

  var chartOptions = chart.getOptions();
  Logger.log(chartOptions.get('annotations'));
  Logger.log(chartOptions.get('series'));
  Logger.log(chartOptions.get('data'));

  var chartOptions = chart.getOptions();
  Logger.log(chartOptions.annotations);
  Logger.log(chartOptions.series);
  Logger.log(chartOptions.data);


}


function drawChart() {
  var sheet = SSheet.getSheetByName(STATUS_GRAPH)
  var dataRange = sheet.getRange("A1:B9");

  var chart = sheet.newChart()
      .asColumnChart()
      .setStacked()
      .setColors(['#a6cee3', '#1f78b4', '#b2df8a'])
      .setXAxisTitle('Наименования')
      .setYAxisTitle('Проценты')
      .setYAxisFormatPattern('#%')
      .setDimensions(800, 400)
      .setRange(dataRange)
      .setPosition(5, 5, 0, 0)
      .build();

  sheet.insertChart(chart);
}

function updateChart() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var charts = sheet.getCharts();

  if (charts.length > 0) {
    for (var i = 0; i < charts.length; i++) {
      var chart = charts[i];
      sheet.removeChart(chart);
    }
  }

  drawChart();
}


function updateChartDataRange() {
  const sheet = SSheet.getSheetByName(DYNAMIC_GRAPH);
  var newRange = getActualDataRange(sheet)

  var graph = SSheet.getSheetByName('Диаграмма 3')
  var chart = graph.getCharts()[0];

  Logger.log(chart.getRanges()[0].getA1Notation())
  chart = chart.modify()
    .removeRange(chart.getRanges()[0])
    .addRange(newRange)
    .build();
  graph.updateChart(chart);
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
