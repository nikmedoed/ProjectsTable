function insertCustomChartImage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Статус (копия)");

  // Получить данные из колонок A и B
  const labels = sheet.getRange("A:A").getValues().flat().filter(label => label);
  const data = sheet.getRange("B:B").getValues().flat().filter(value => value);

  // Цвета для колонок
  const backgroundColors = [
    'rgba(255, 105, 180, 0.5)', 'rgba(255, 165, 0, 0.5)', 'rgba(173, 216, 230, 0.5)',
    'rgba(238, 130, 238, 0.5)', 'rgba(250, 250, 210, 0.5)', 'rgba(176, 224, 230, 0.5)',
    'rgba(240, 128, 128, 0.5)', 'rgba(152, 251, 152, 0.5)'
  ];

  // Конфигурация графика
  const chartConfig = {
    type: "bar",
    data: {
      datasets: [{
        data: [0.34, 0.23, 0.23, 0.52],
        backgroundColor: ['rgba(255, 105, 180, 0.5)', 'rgba(255, 165, 0, 0.5)', 'rgba(173, 216, 230, 0.5)', 'rgba(238, 130, 238, 0.5)'],
        barPercentage: 0.9,
        categoryPercentage: 0.85,
        type: "bar",
        borderColor: "rgba(54, 162, 235, 0.5)",
        borderWidth: 1,
      }],
      labels: ["укп", "укппук", "уйкпуйкпуйкп", "пукпукп кsdfasdfas s dfasfas dfasdfasdfsad fasd fsadf asdf asdf asd fasdf asdf asdfупук"]
    },
    options: {
      title: {
        display: true,
        position: "top",
        fontSize: 14,
        text: "Фактический статус проекта"
      },
      legend: { display: false },
      scales: {
        x: {
          ticks: {
            autoSkip: false,
            maxRotation: 90,
            minRotation: 0
          },
          gridLines: {
            display: false,
            drawOnChartArea: false,
          },
        },
        y: {
          beginAtZero: true,
          max: 1,
          ticks: {
            stepSize: 0.1,
            callback: (value) => (value * 100) + '%'
          }
        },

      },
      plugins: {
        datalabels: {
          anchor: 'end',
          align: 'top',
          color: '#000',
          font: {
            weight: 'bold',
            size: 14
          },
          formatter: (value) => (value * 100) + '%',
        },
      }
    },
    backgroundColor: "white",
    width: 1800,
    height: 700,
    version: 4
  }

  // Генерация URL для графика
  const chartUrl = `https://quickchart.io/chart?c=${encodeURIComponent(JSON.stringify(chartConfig))}`;

  // Вставить изображение в лист
  sheet.insertImage(chartUrl, 1, 4);
}



// ================== EXPERIMENTS ========================
// function annotations(chart) {

//   return chart
//     .setOption('annotations', {
//       alwaysOutside: true,
//       textStyle: {
//         fontSize: 16,
//         color: 'red'
//       }
//     })
//     .setOption('annotations.domain', true)
//     .setOption('series', {
//       0: {
//         annotations: {
//           textStyle: { fontSize: 12, color: 'red' } // Цвет аннотаций меняется, размер меняется, они сами не появляются
//         }
//       },
//       1: {
//         annotations: {
//           textStyle: { fontSize: 12, color: 'blue' }
//         }
//       }
//     })
//     .setOption('series', {
//       0: {
//         annotations: {
//           textStyle: { fontSize: 12, color: 'red' },
//           alwaysOutside: true,
//           boxStyle: {
//             stroke: '#999',
//             strokeWidth: 1,
//             gradient: {
//               color1: '#fbf6a7',
//               color2: '#33b679',
//               x1: '0%', y1: '0%',
//               x2: '100%', y2: '100%',
//               useObjectBoundingBoxCoordinates: true
//             }
//           }
//         }
//       }
//     })
// }

// function test() {
//   var graph = SSheet.getSheetByName(DYNAMIC_GRAPH)
//   var chart = graph.getCharts()[0];

//   var chartOptions = chart.getOptions();
//   Logger.log(chartOptions.get('annotations'));
//   Logger.log(chartOptions.get('series'));
//   Logger.log(chartOptions.get('data'));

//   var chartOptions = chart.getOptions();
//   Logger.log(chartOptions.annotations);
//   Logger.log(chartOptions.series);
//   Logger.log(chartOptions.data);
// }