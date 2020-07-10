//######################################################################
//HORIZONTAL STACKED BAR CHARTS

function addChartHorizontalBarStacked1(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        yAxes: [{stacked: true}],
				xAxes: [{
          stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}



function addChartHorizontalBarStacked2(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        yAxes: [{stacked: true}],
				xAxes: [{
          stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}




function addChartHorizontalBarStacked3(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        yAxes: [{stacked: true}],
				xAxes: [{
          stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0],
        DataPointsArray[1]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBarStacked4(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        yAxes: [{stacked: true}],
        xAxes: [{
          stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0],
        DataPointsArray[1],
        DataPointsArray[2]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBarStacked5(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        yAxes: [{stacked: true}],
				xAxes: [{
          stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBarStacked6(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        yAxes: [{stacked: true}],
				xAxes: [{
          stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}
function addChartHorizontalBarStacked7(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        yAxes: [{stacked: true}],
				xAxes: [{
          stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBarStacked8(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        yAxes: [{stacked: true}],
				xAxes: [{
          stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBarStacked9(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        yAxes: [{stacked: true}],
				xAxes: [{
          stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
      DataPointsArray[7],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBarStacked10(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        yAxes: [{stacked: true}],
				xAxes: [{
          stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
      DataPointsArray[7],
      DataPointsArray[8],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}



//######################################################################
//STACKED BAR CHARTS

function addChartBarStacked1(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
		responsive: true,
		scales: {
			xAxes: [{
				stacked: true
			}],
			yAxes: [{
			stacked: true,
			beginAtZero: true,
			}]
		}
	},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}



function addChartBarStacked2(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
		responsive: true,
		scales: {
			xAxes: [{
				stacked: true
			}],
			yAxes: [{
			stacked: true,
			beginAtZero: true,
			}]
		}
	},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}




function addChartBarStacked3(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
		responsive: true,
		scales: {
			xAxes: [{
				stacked: true
			}],
			yAxes: [{
			stacked: true,
			beginAtZero: true,
			}]
		}
	},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0],
        DataPointsArray[1]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBarStacked4(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
		responsive: true,
		scales: {
			xAxes: [{
				stacked: true
			}],
			yAxes: [{
			stacked: true,
			beginAtZero: true,
			}]
		}
	},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0],
        DataPointsArray[1],
        DataPointsArray[2]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBarStacked5(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {

		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
		responsive: true,
		scales: {
			xAxes: [{
				stacked: true
			}],
			yAxes: [{
			stacked: true,
			beginAtZero: true,
			}]
		}
	},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBarStacked6(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
		responsive: true,
		scales: {
			xAxes: [{
				stacked: true
			}],
			yAxes: [{
			stacked: true,
			beginAtZero: true,
			}]
		}
	},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}
function addChartBarStacked7(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
		responsive: true,
		scales: {
			xAxes: [{
				stacked: true
			}],
			yAxes: [{
			stacked: true,
			beginAtZero: true,
			}]
		}
	},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBarStacked8(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
		responsive: true,
		scales: {
			xAxes: [{
				stacked: true
			}],
			yAxes: [{
			stacked: true,
			beginAtZero: true,
			}]
		}
	},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBarStacked9(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
		responsive: true,
		scales: {
			xAxes: [{
				stacked: true
			}],
			yAxes: [{
			stacked: true,
			beginAtZero: true,
			}]
		}
	},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
      DataPointsArray[7],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBarStacked10(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
		responsive: true,
		scales: {
			xAxes: [{
				stacked: true
			}],
			yAxes: [{
			stacked: true,
			beginAtZero: true,
			}]
		}
	},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
      DataPointsArray[7],
      DataPointsArray[8],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

//######################################################################
//LINE CHARTS

function addChartLine1(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
				yAxes: [{
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
        fill: false,
        backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}



function addChartLine2(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
				yAxes: [{
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
        fill: false,
        backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}




function addChartLine3(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
				yAxes: [{
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
        fill: false,
        backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0],
        DataPointsArray[1]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartLine4(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,

			},
			scales: {
				yAxes: [{
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
        fill: false,
        backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,

			},

				DataPointsArray[0],
        DataPointsArray[1],
        DataPointsArray[2]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartLine5(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
				yAxes: [{
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
        fill: false,
        backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartLine6(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
				yAxes: [{
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
        fill: false,
        backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}
function addChartLine7(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
				yAxes: [{
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
        fill: false,
        backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartLine8(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
				yAxes: [{
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
        fill: false,
        backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartLine9(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
				yAxes: [{
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
        fill: false,
        backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
      DataPointsArray[7],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartLine10(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
				yAxes: [{
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
        fill: false,
        backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
      DataPointsArray[7],
      DataPointsArray[8],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}



//######################################################################
//HORIZONTAL BAR CHARTS

function addChartHorizontalBar1(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //yAxes: [{stacked: true}],
				xAxes: [{
          //stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}



function addChartHorizontalBar2(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //yAxes: [{stacked: true}],
				xAxes: [{
          //stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}




function addChartHorizontalBar3(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //yAxes: [{stacked: true}],
				xAxes: [{
          //stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0],
        DataPointsArray[1]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBar4(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //yAxes: [{stacked: true}],
        xAxes: [{
          //stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0],
        DataPointsArray[1],
        DataPointsArray[2]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBar5(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //yAxes: [{stacked: true}],
				xAxes: [{
          //stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBar6(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //yAxes: [{stacked: true}],
				xAxes: [{
          //stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}
function addChartHorizontalBar7(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //yAxes: [{stacked: true}],
				xAxes: [{
          //stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBar8(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //yAxes: [{stacked: true}],
				xAxes: [{
          //stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBar9(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //yAxes: [{stacked: true}],
				xAxes: [{
          //stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
      DataPointsArray[7],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartHorizontalBar10(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //yAxes: [{stacked: true}],
				xAxes: [{
          //stacked: true,
					ticks: {
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
      DataPointsArray[7],
      DataPointsArray[8],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}


//######################################################################
//BAR CHARTS

function addChartBar1(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
			scales: {
        //xAxes: [{stacked: true}],
				yAxes: [{
					ticks: {
            //stacked: true,
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}



function addChartBar2(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
      scales: {
        //xAxes: [{stacked: true}],
				yAxes: [{
					ticks: {
            //stacked: true,
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}




function addChartBar3(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
      scales: {
        //xAxes: [{stacked: true}],
				yAxes: [{
					ticks: {
            //stacked: true,
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0],
        DataPointsArray[1]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBar4(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
      scales: {
        //xAxes: [{stacked: true}],
				yAxes: [{
					ticks: {
            //stacked: true,
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

				DataPointsArray[0],
        DataPointsArray[1],
        DataPointsArray[2]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBar5(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
      scales: {
        //xAxes: [{stacked: true}],
				yAxes: [{
					ticks: {
            //stacked: true,
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3]
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBar6(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
      scales: {
        //xAxes: [{stacked: true}],
				yAxes: [{
					ticks: {
            //stacked: true,
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}
function addChartBar7(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
      scales: {
        //xAxes: [{stacked: true}],
				yAxes: [{
					ticks: {
            //stacked: true,
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBar8(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
      scales: {
        //xAxes: [{stacked: true}],
				yAxes: [{
					ticks: {
            //stacked: true,
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBar9(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
      scales: {
        //xAxes: [{stacked: true}],
				yAxes: [{
					ticks: {
            //stacked: true,
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
      DataPointsArray[7],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

function addChartBar10(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		options: {
			legend: {
				display: true,
			},
      scales: {
        //xAxes: [{stacked: true}],
				yAxes: [{
					ticks: {
            //stacked: true,
						beginAtZero: true,
					},
				}]
			},
		},
		data: {

			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

      DataPointsArray[0],
      DataPointsArray[1],
      DataPointsArray[2],
      DataPointsArray[3],
      DataPointsArray[4],
      DataPointsArray[5],
      DataPointsArray[6],
      DataPointsArray[7],
      DataPointsArray[8],
			],
			labels: webDevObjects[0].charts[0][9],
		},
	});
}

//######################################################################
//PIE CHARTS

function addChartPie(i){
	x = document.getElementById(webDevObjects[0].charts[i][0])
	CID = `Chart${[i]}`
	x.innerHTML = `
	 <div align='center' style='padding:5px;'>
		<div style='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px; height:${webDevObjects[0].charts[i][4]};'>
			<canvas id='${CID}' width='width:${screenWidthPercent(webDevObjects[0].charts[i][3])}px' height='${webDevObjects[0].charts[i][4]}px'></canvas>
		</div>
	</div>
	`
	var ctx = document.getElementById(CID).getContext('2d');
	var myChart = new Chart(ctx, {
		type: webDevObjects[0].charts[i][2],
		data: {
			labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
			datasets: [{
				label: webDevObjects[0].charts[i][6][0],
				data: webDevObjects[0].charts[i][5][0],
				backgroundColor: webDevObjects[0].charts[i][7][0],
				borderColor: webDevObjects[0].charts[i][8][0],
				borderWidth: 1,
				order: 1,
			},

			],
			labels: webDevObjects[0].charts[0][9]
		},

	});
}
