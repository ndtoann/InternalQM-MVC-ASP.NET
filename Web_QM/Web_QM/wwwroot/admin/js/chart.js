$(function() {
  /* ChartJS
   * -------
   * Data and config for chartjs
   */
  'use strict';
  var barData = {
    labels: ["T1","T2","T3","T4","T5","T6","T7","T8","T9","T10","T11","T12"],
    datasets: [{
      label: 'Số lỗi 5S',
        data: dataBar,
        backgroundColor: [
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)',
            'rgba(255, 105, 97, 0.7)'
        ],
        borderColor: [
            'rgba(255, 105, 97, 1)',
            'rgba(255, 105, 97, 1)',
            'rgba(255, 105, 97, 1)',
            'rgba(255, 105, 97, 1)',
            'rgba(255, 105, 97, 1)',
            'rgba(255, 105, 97, 1)',
            'rgba(255, 105, 97, 1)',
            'rgba(255, 105, 97, 1)',
            'rgba(255, 105, 97, 1)',
            'rgba(255, 105, 97, 1)',
            'rgba(255, 105, 97, 1)'
        ],
      borderWidth: 1,
      fill: false
    }]
  };

  var lineData = {
    labels: ["T1","T2","T3","T4","T5","T6","T7","T8","T9","T10","T11","T12"],
    datasets: [{
      label: 'Số kaizen',
        data: dataLine,
        backgroundColor: [
            'rgba(175, 222, 175, 0.7)'
        ],
        borderColor: [
            'rgba(175, 222, 175, 1)'
        ],
      borderWidth: 1,
      fill: false
    }]
  };

  var multiLineData = {
    labels: ["Red", "Blue", "Yellow", "Green", "Purple", "Orange"],
    datasets: [{
        label: 'Dataset 1',
        data: [12, 19, 3, 5, 2, 3],
        borderColor: [
          '#587ce4'
        ],
        borderWidth: 2,
        fill: false
      },
      {
        label: 'Dataset 2',
        data: [5, 23, 7, 12, 42, 23],
        borderColor: [
          '#ede190'
        ],
        borderWidth: 2,
        fill: false
      },
      {
        label: 'Dataset 3',
        data: [15, 10, 21, 32, 12, 33],
        borderColor: [
          '#f44252'
        ],
        borderWidth: 2,
        fill: false
      }
    ]
  };
  var options = {
    scales: {
      yAxes: [{
        ticks: {
          beginAtZero: true
        }
      }]
    },
    legend: {
      display: false
    },
    elements: {
      point: {
        radius: 0
      }
    }

    };

  var doughnutPieData = {
    datasets: [{
        data: dataPie,
        backgroundColor: [
            'rgba(255, 105, 97, 0.7)',   // Đỏ san hô
            'rgba(175, 222, 175, 0.7)',  // Xanh lá cây nhạt
            'rgba(106, 90, 205, 0.7)',   // Xanh tím
            'rgba(255, 165, 0, 0.7)',    // Cam
            'rgba(255, 230, 0, 0.7)',    // Vàng sáng
            'rgba(64, 224, 208, 0.7)',   // Xanh lam ngọc
            'rgba(255, 105, 180, 0.7)',  // Hồng
            'rgba(0, 191, 255, 0.7)',    // Xanh da trời
            'rgba(128, 0, 128, 0.7)',    // Tím
            'rgba(255, 192, 203, 0.7)',  // Hồng nhạt
            'rgba(189, 183, 107, 0.7)'   // Xanh ô liu đậm
        ],
        borderColor: [
            'rgba(255, 105, 97, 1)',
            'rgba(175, 222, 175, 1)',
            'rgba(106, 90, 205, 1)',
            'rgba(255, 165, 0, 1)',
            'rgba(255, 230, 0, 1)',
            'rgba(64, 224, 208, 1)',
            'rgba(255, 105, 180, 1)',
            'rgba(0, 191, 255, 1)',
            'rgba(128, 0, 128, 1)',
            'rgba(255, 192, 203, 1)',
            'rgba(189, 183, 107, 1)'
        ],
    }],
      labels: labelPie
  };
  var doughnutPieOptions = {
    responsive: true,
    animation: {
      animateScale: true,
      animateRotate: true
    }
  };
  var areaData = {
    labels: ["2013", "2014", "2015", "2016", "2017"],
    datasets: [{
      label: '# of Votes',
      data: [12, 19, 3, 5, 2, 3],
      backgroundColor: [
        'rgba(255, 99, 132, 0.2)',
        'rgba(54, 162, 235, 0.2)',
        'rgba(255, 206, 86, 0.2)',
        'rgba(75, 192, 192, 0.2)',
        'rgba(153, 102, 255, 0.2)',
        'rgba(255, 159, 64, 0.2)'
      ],
      borderColor: [
        'rgba(255,99,132,1)',
        'rgba(54, 162, 235, 1)',
        'rgba(255, 206, 86, 1)',
        'rgba(75, 192, 192, 1)',
        'rgba(153, 102, 255, 1)',
        'rgba(255, 159, 64, 1)'
      ],
      borderWidth: 1,
      fill: true, // 3: no fill
    }]
  };

  var areaOptions = {
    plugins: {
      filler: {
        propagate: true
      }
    }
  }

  var multiAreaData = {
    labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    datasets: [{
        label: 'Facebook',
        data: [8, 11, 13, 15, 12, 13, 16, 15, 13, 19, 11, 14],
        borderColor: ['rgba(255, 99, 132, 0.5)'],
        backgroundColor: ['rgba(255, 99, 132, 0.5)'],
        borderWidth: 1,
        fill: true
      },
      {
        label: 'Twitter',
        data: [7, 17, 12, 16, 14, 18, 16, 12, 15, 11, 13, 9],
        borderColor: ['rgba(54, 162, 235, 0.5)'],
        backgroundColor: ['rgba(54, 162, 235, 0.5)'],
        borderWidth: 1,
        fill: true
      },
      {
        label: 'Linkedin',
        data: [6, 14, 16, 20, 12, 18, 15, 12, 17, 19, 15, 11],
        borderColor: ['rgba(255, 206, 86, 0.5)'],
        backgroundColor: ['rgba(255, 206, 86, 0.5)'],
        borderWidth: 1,
        fill: true
      }
    ]
  };

  var multiAreaOptions = {
    plugins: {
      filler: {
        propagate: true
      }
    },
    elements: {
      point: {
        radius: 0
      }
    },
    scales: {
      xAxes: [{
        gridLines: {
          display: false
        }
      }],
      yAxes: [{
        gridLines: {
          display: false
        }
      }]
    }
  }

  var scatterChartData = {
    datasets: [{
        label: 'First Dataset',
        data: [{
            x: -10,
            y: 0
          },
          {
            x: 0,
            y: 3
          },
          {
            x: -25,
            y: 5
          },
          {
            x: 40,
            y: 5
          }
        ],
        backgroundColor: [
          'rgba(255, 99, 132, 0.2)'
        ],
        borderColor: [
          'rgba(255,99,132,1)'
        ],
        borderWidth: 1
      },
      {
        label: 'Second Dataset',
        data: [{
            x: 10,
            y: 5
          },
          {
            x: 20,
            y: -30
          },
          {
            x: -25,
            y: 15
          },
          {
            x: -10,
            y: 5
          }
        ],
        backgroundColor: [
          'rgba(54, 162, 235, 0.2)',
        ],
        borderColor: [
          'rgba(54, 162, 235, 1)',
        ],
        borderWidth: 1
      }
    ]
  }

  var scatterChartOptions = {
    scales: {
      xAxes: [{
        type: 'linear',
        position: 'bottom'
      }]
    }
  }
  // Get context with jQuery - using jQuery's .get() method.
  if ($("#barChart").length) {
    var barChartCanvas = $("#barChart").get(0).getContext("2d");
    // This will get the first returned node in the jQuery collection.
    var barChart = new Chart(barChartCanvas, {
      type: 'bar',
      data: barData,
      options: options
    });
  }

  if ($("#lineChart").length) {
    var lineChartCanvas = $("#lineChart").get(0).getContext("2d");
    var lineChart = new Chart(lineChartCanvas, {
      type: 'line',
      data: lineData,
      options: options
    });
  }

  if ($("#linechart-multi").length) {
    var multiLineCanvas = $("#linechart-multi").get(0).getContext("2d");
    var lineChart = new Chart(multiLineCanvas, {
      type: 'line',
      data: multiLineData,
      options: options
    });
  }

  if ($("#areachart-multi").length) {
    var multiAreaCanvas = $("#areachart-multi").get(0).getContext("2d");
    var multiAreaChart = new Chart(multiAreaCanvas, {
      type: 'line',
      data: multiAreaData,
      options: multiAreaOptions
    });
  }

  if ($("#doughnutChart").length) {
    var doughnutChartCanvas = $("#doughnutChart").get(0).getContext("2d");
    var doughnutChart = new Chart(doughnutChartCanvas, {
      type: 'doughnut',
      data: doughnutPieData,
      options: doughnutPieOptions
    });
  }

  if ($("#pieChart").length) {
    var pieChartCanvas = $("#pieChart").get(0).getContext("2d");
    var pieChart = new Chart(pieChartCanvas, {
      type: 'pie',
      data: doughnutPieData,
      options: doughnutPieOptions
    });
  }

  if ($("#areaChart").length) {
    var areaChartCanvas = $("#areaChart").get(0).getContext("2d");
    var areaChart = new Chart(areaChartCanvas, {
      type: 'line',
      data: areaData,
      options: areaOptions
    });
  }

  if ($("#scatterChart").length) {
    var scatterChartCanvas = $("#scatterChart").get(0).getContext("2d");
    var scatterChart = new Chart(scatterChartCanvas, {
      type: 'scatter',
      data: scatterChartData,
      options: scatterChartOptions
    });
  }

  if ($("#browserTrafficChart").length) {
    var doughnutChartCanvas = $("#browserTrafficChart").get(0).getContext("2d");
    var doughnutChart = new Chart(doughnutChartCanvas, {
      type: 'doughnut',
      data: browserTrafficData,
      options: doughnutPieOptions
    });
  }
});