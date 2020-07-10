//#################################################################################################
var ary
var unique
var x
var z = Array[5]
var v
var cartDatapointsArray
var DataPointsArray = new Array()
//#################################################################################################
//FUNCTIONS
/*
function aryBlank(dt){
  for (i = 0; i < dt.length; i++){
    for (j = 0; j < 1; i++){
      ary[i][j] = ""
    };
  };
}

function uniqueArray(dt,dim){
  var tmp
  for (i = 0; i < dt.length; i++){
    tmp[i] = dt[i][dim]
  };
  unique = jQuery.unique(tmp);
}
*/

function screenWidthPercent(p){
	return screen.width / 100 * p
}
//#################################################################################################
/*
	webDevObjects[0].charts[i][5].length
*/





function chartBarDataset(i,j){
	try {
		z = [{
			label: webDevObjects[0].charts[i][6][j],
			data: webDevObjects[0].charts[i][5][j],
			backgroundColor: webDevObjects[0].charts[i][7][j],
			borderColor: webDevObjects[0].charts[i][8][j],
			borderWidth: 1,
			fill: false,
			order: j,


		}]

		return z
	}
	catch(err) {}
}

function DP(k) {
return [{	label:DataPointsArray[k-1].label,
					data: DataPointsArray[k-1].data,
					backgroundColor: DataPointsArray[k-1].backgroundColor,
					borderColor: DataPointsArray[k-1].borderColor,
					borderWidth: 1,
					order: 1, //use to be k

				}]
}

function addCharts(){
	for (i = 0; i < webDevObjects[0].charts.length; i++){


		for (j = 1; j < webDevObjects[0].charts[i][5].length; j++){
			chartBarDataset(i,j)
			DataPointsArray[j-1] = z[0]
		}

		if(webDevObjects[0].charts[i][2] == 'pie'){addChartPie(i)}

		//webDevObjects[0].charts[1][11]
		if(webDevObjects[0].charts[i][2] == 'bar'  && webDevObjects[0].charts[i][11] == false){
			if (DataPointsArray.length == 0) {addChartBar1(i)}
			if (DataPointsArray.length == 1) {addChartBar2(i)}
			if (DataPointsArray.length == 2) {addChartBar3(i)}
			if (DataPointsArray.length == 3) {addChartBar4(i)}
			if (DataPointsArray.length == 4) {addChartBar5(i)}
			if (DataPointsArray.length == 5) {addChartBar6(i)}
			if (DataPointsArray.length == 6) {addChartBar7(i)}
			if (DataPointsArray.length == 7) {addChartBar8(i)}
			if (DataPointsArray.length == 8) {addChartBar9(i)}
			if (DataPointsArray.length == 9) {addChartBar10(i)}
		}

		if(webDevObjects[0].charts[i][2] == 'horizontalBar'  && webDevObjects[0].charts[i][11] == false){
			if (DataPointsArray.length == 0) {addChartHorizontalBar1(i)}
			if (DataPointsArray.length == 1) {addChartHorizontalBar2(i)}
			if (DataPointsArray.length == 2) {addChartHorizontalBar3(i)}
			if (DataPointsArray.length == 3) {addChartHorizontalBar4(i)}
			if (DataPointsArray.length == 4) {addChartHorizontalBar5(i)}
			if (DataPointsArray.length == 5) {addChartHorizontalBar6(i)}
			if (DataPointsArray.length == 6) {addChartHorizontalBar7(i)}
			if (DataPointsArray.length == 7) {addChartHorizontalBar8(i)}
			if (DataPointsArray.length == 8) {addChartHorizontalBar9(i)}
			if (DataPointsArray.length == 9) {addChartHorizontalBar10(i)}
		}

		if(webDevObjects[0].charts[i][2] == 'bar'  && webDevObjects[0].charts[i][11] == true){
			if (DataPointsArray.length == 0) {addChartBarStacked1(i)}
			if (DataPointsArray.length == 1) {addChartBarStacked2(i)}
			if (DataPointsArray.length == 2) {addChartBarStacked3(i)}
			if (DataPointsArray.length == 3) {addChartBarStacked4(i)}
			if (DataPointsArray.length == 4) {addChartBarStacked5(i)}
			if (DataPointsArray.length == 5) {addChartBarStacked6(i)}
			if (DataPointsArray.length == 6) {addChartBarStacked7(i)}
			if (DataPointsArray.length == 7) {addChartBarStacked8(i)}
			if (DataPointsArray.length == 8) {addChartBarStacked9(i)}
			if (DataPointsArray.length == 9) {addChartBarStacked10(i)}
		}

		if(webDevObjects[0].charts[i][2] == 'horizontalBar'  && webDevObjects[0].charts[i][11] == true){
			if (DataPointsArray.length == 0) {addChartHorizontalBarStacked1(i)}
			if (DataPointsArray.length == 1) {addChartHorizontalBarStacked2(i)}
			if (DataPointsArray.length == 2) {addChartHorizontalBarStacked3(i)}
			if (DataPointsArray.length == 3) {addChartHorizontalBarStacked4(i)}
			if (DataPointsArray.length == 4) {addChartHorizontalBarStacked5(i)}
			if (DataPointsArray.length == 5) {addChartHorizontalBarStacked6(i)}
			if (DataPointsArray.length == 6) {addChartHorizontalBarStacked7(i)}
			if (DataPointsArray.length == 7) {addChartHorizontalBarStacked8(i)}
			if (DataPointsArray.length == 8) {addChartHorizontalBarStacked9(i)}
			if (DataPointsArray.length == 9) {addChartHorizontalBarStacked10(i)}
		}

		if(webDevObjects[0].charts[i][2] == 'line'  && webDevObjects[0].charts[i][11] == false){
			if (DataPointsArray.length == 0) {addChartLine1(i)}
			if (DataPointsArray.length == 1) {addChartLine2(i)}
			if (DataPointsArray.length == 2) {addChartLine3(i)}
			if (DataPointsArray.length == 3) {addChartLine4(i)}
			if (DataPointsArray.length == 4) {addChartLine5(i)}
			if (DataPointsArray.length == 5) {addChartLine6(i)}
			if (DataPointsArray.length == 6) {addChartLine7(i)}
			if (DataPointsArray.length == 7) {addChartLine8(i)}
			if (DataPointsArray.length == 8) {addChartLine9(i)}
			if (DataPointsArray.length == 9) {addChartLine10(i)}
		}
	}
}
function addMeasure() {
	for (i = 0;i < webDevObjects[0].measures.length ;i++){
		x = document.getElementById(webDevObjects[0].measures[i][0])
		x.innerHTML += `
		<div align='center'>
			<button type='button' class='btn ${webDevObjects[0].measures[i][1]}' style='${webDevObjects[0].measures[i][3]}</span>${webDevObjects[0].measures[i][6]} ; width:${webDevObjects[0].measures[i][5]}px;margin:5px ;'>
				${webDevObjects[0].measures[i][2]} <span class='badge badge-light'>${webDevObjects[0].measures[i][3]}</span>${webDevObjects[0].measures[i][4]}
			</button>
		</div>
		`
	}
}
function createInnerTable(k,i){
	var rc = ""
	//Does the table have headers
	if (i == 2) {
		//First Row
		if (webDevObjects[0].tables[k][1][1] == "yes") {

			//Table Headers
			for (j = 0; j < webDevObjects[0].tables[k][2].length; j++){
				rc += `
					<th> ${webDevObjects[0].tables[k][2][j]}</th>
				`
			};
		} else {

			//Table Body
			for (j = 0; j < webDevObjects[0].tables[k][2].length; j++){

				rc += `
					<td>${webDevObjects[0].tables[k][i][j]}</td>
				`
			};
		};
	} else {
	//Not First Row
		for (j = 0; j < webDevObjects[0].tables[k][2].length; j++){
				rc += `
					<td>${webDevObjects[0].tables[k][i][j]}</td>
				`
			};
	}

	return rc
}
function createTableRow(k){
	//Create Rows
	var rw = ""
	for (i = 2; i < webDevObjects[0].tables[k].length; i++){
		rw += `
			<tr>
				${createInnerTable(k,i)}
			</tr>
		`
	};
	return rw
}
function createTable(k){
	//Create Table
	tableWidth = screenWidthPercent(webDevObjects[0].tables[k][1][2])
	tableStyles = (webDevObjects[0].tables[k][1][3])
	DivClass = (webDevObjects[0].tables[k][1][4])
	txt = `
	<div style='width:${tableWidth}px;padding:5px; ${tableStyles}' class='${DivClass}'>
		<table class='${webDevObjects[0].tables[k][1][0]}' align='center'>
			${createTableRow(k)}
		</table>
	</div>
	`
	x = document.getElementById(webDevObjects[0].tables[k][0])
	x.innerHTML += txt
	//document.getElementById("main").innerHTML = txt
}

function selectTables(){
	//Select Table
	for (k = 0; k < webDevObjects[0].tables.length; k++){
		createTable(k)
	};
}
//#################################################################################################
function navBarLinks(i){
	let txt = ""
	i = i + 1
	for (j = 0; j < webDevObjects[0].navBar[1].length; j++){
		txt += `
			<a class="dropdown-item" href='${webDevObjects[0].navBar[i][j][1]}' target="none">${webDevObjects[0].navBar[i][j][0]}</a>
		`
	}
	return txt
}
function navBarDropdown(){
	let txt = ""
	for (i = 0; i < webDevObjects[0].navBar[0][0].length; i++){
		txt += `
		<li class="nav-item dropdown">
			<a class="nav-link dropdown-toggle" href="#" id="navbardrop" data-toggle="dropdown">
                ${webDevObjects[0].navBar[0][0][i]}
                </a>
			<div class="dropdown-menu">
				${navBarLinks(i)}
			</div>
		`
	}
	return txt
}
function navBar(){

	return `
			<nav class='navbar navbar-expand-sm bg-dark navbar-dark' style="margin-top:-20px;position:fixed;width:100%;color:white;font-weight:bold;'>
			<a class='navbar-brand' href='https://www.allform.tech' target="none" id='navBarTitle'>${webDevObjects[0].pageDetails.navBarTitle} &nbsp; &nbsp; </a>
			<ul class="navbar-nav" id="navBar">
                    <li class="nav-item dropdown">
                        ${navBarDropdown()}
                      </li>
                </ul>
            </nav>
			</nav>
		`
}
//#################################################################################################
function innerMasterTable(i) {
	let txt = ""
	for (j = 0; j < 20; j++){
		txt += `
			<td id='r${i}-c${j}'></td>
		`
	}
	return txt
}
function createMasterTable(){
	let txt = ""
		for (i = 0; i < 50; i++){
			txt +=  `
		<table style='width:100%;' id='r${i}' align='center'>
					<tr>
						${innerMasterTable(i)}
					</tr>
				</table>
		`
		}
  return txt
}
//webDevObjects[0].text[0][i]
function addText(){
	for (i = 0; i < webDevObjects[0].text.length; i++){
		x = document.getElementById(webDevObjects[0].text[i][0])
		let txt = `
			<div align='${webDevObjects[0].text[i][4]}' style='width: ${screenWidthPercent(webDevObjects[0].text[i][1])}px;height: ${webDevObjects[0].text[i][2]}px;${webDevObjects[0].text[i][3]}'>
				${webDevObjects[0].text[i][5]}
			</div>
		`
	x.innerHTML += txt
	}
}

//<iframe src="http://buythecity.com"  scrolling="no" style=" width: 550px; height: 500px;  overflow: hidden;" ></iframe>
//webDevObjects[0].iframe.length
function addIframe(){
	for (i = 0; i < webDevObjects[0].iframe.length; i++){
		x = document.getElementById(webDevObjects[0].iframe[i][0])
		let txt = `
				<iframe src='${webDevObjects[0].iframe[i][6]}'  scrolling='${webDevObjects[0].iframe[i][5]}' style=' width: ${screenWidthPercent(webDevObjects[0].iframe[i][1])}px; height: ${webDevObjects[0].iframe[i][2]}px; overflow: ${webDevObjects[0].iframe[i][7]};${webDevObjects[0].iframe[i][3]}'frameBorder=' ${webDevObjects[0].iframe[i][8]}'></iframe>
				`
	x.innerHTML += txt
	}
}

function addImage(){
	for (i = 0; i < webDevObjects[0].images.length; i++){
		x = document.getElementById(webDevObjects[0].images[i][0])
		let txt = `
				<div align='${webDevObjects[0].images[i][4]}'>
					<img src='${webDevObjects[0].images[i][5]}' style=' width: ${screenWidthPercent(webDevObjects[0].images[i][1])}px; height: ${webDevObjects[0].images[i][2]}px; ${webDevObjects[0].images[i][3]}'></img>
				</div>
				`
	x.innerHTML += txt
	}
}


//#################################################################################################
function main() {
		x = document.getElementById("main")
		x.innerHTML += navBar()
		x.innerHTML += "<br><br>"
		x.innerHTML += createMasterTable()
		selectTables()
		addMeasure()
		addCharts()
		addText()
		addIframe()
		addImage()
	}
//#################################################################################################
$(document).ready(main());
