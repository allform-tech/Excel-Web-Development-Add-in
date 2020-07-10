webDevObjects = [
	{
		pageDetails :
		  {
		   pageTitle : "My page",							//webDevObjects[0].pageDetails.pageTitle = "My page"
		   navBarTitle : "Allform Software Solutions",
		   headingText : "My heading",
		   headingClass : "alert alert-info",
		  },

		charts :
		[
			//webDevObjects[0].charts[0][6][1][0] = 'orange'
			//webDevObjects[0].charts[0][5][0][0] = 10
			//webDevObjects[0].charts[0][3] = 400
			//webDevObjects[0].charts[0][2] = 'pie'
			['r5-c1', 'My Title','pie',30,150,
			//     			Data Points
				[[10,50,20,39,54,44,45],[14,44,57,99,54,57,88],],
				//Labels
				['Dataset 1'],
				//												Background Color
				[['green','red','blue','orange','pink','brown','purple'],],
				//													Border Color
				[['gray','gray','gray','gray','gray','gray','gray'],],
				//Labels
				['Kris','Ben','John','Michle','James','Tim','Ryan']
        ,false
			],

			['r5-c2', 'My Title','horizontalBar',30,150,
			//     			Data Points
				[[10,50,20,39,54,44,45],[14,44,57,99,54,57,88],[11,22,57,88,99,55,44],[45,54,74,44,12,12,45],],
				//Labels
				['Dataset 1','Dataset 2','Dataset 3','Dataset 4'],
				//												Background Color
				[['green','green','green','green','green','green','green'],['orange','orange','orange','orange','orange','orange','orange'],['pink','pink','pink','pink','pink','pink','pink'],['red','red','red','red','red','red','red']],
				//													Border Color
				[['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray']],
				//Labels
				['Kris','Ben','John','Michle','James','Tim','Ryan'],
        ,false
			],

      ['r5-c3', 'My Title','line',30,150,
			//     			Data Points
				[[10,50,20,39,54,44,45],[14,44,57,99,54,57,88],[11,22,57,88,99,55,44],[45,54,74,44,12,12,45],],
				//Labels
				['Dataset 1','Dataset 2','Dataset 3','Dataset 4'],
				//												Background Color
        [['green','green','green','green','green','green','green'],['orange','orange','orange','orange','orange','orange','orange'],['pink','pink','pink','pink','pink','pink','pink'],['red','red','red','red','red','red','red']],
				//													Border Color
				[['green'],['orange'],['pink'],['red']],
        //Labels
				['Kris','Ben','John','Michle','James','Tim','Ryan'],
        ,false
			],

      ['r6-c1', 'My Title','bar',30,150,
			//     			Data Points
				[[10,50,20,39,54,44,45],[14,44,57,99,54,57,88],[11,22,57,88,99,55,44],[45,54,74,44,12,12,45],],
				//Labels
				['Dataset 1','Dataset 2','Dataset 3','Dataset 4'],
				//												Background Color
				[['green','green','green','green','green','green','green'],['orange','orange','orange','orange','orange','orange','orange'],['pink','pink','pink','pink','pink','pink','pink'],['red','red','red','red','red','red','red']],
				//													Border Color
				[['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray']],
				//Labels
				['Kris','Ben','John','Michle','James','Tim','Ryan'],
        ,false
			],
      ['r6-c2', 'My Title','bar',30,150,
			//     			Data Points
				[[10,50,20,39,54,44,45],[14,44,57,99,54,57,88],[11,22,57,88,99,55,44],[45,54,74,44,12,12,45],],
				//Labels
				['Dataset 1','Dataset 2','Dataset 3','Dataset 4'],
				//												Background Color
				[['green','green','green','green','green','green','green'],['orange','orange','orange','orange','orange','orange','orange'],['pink','pink','pink','pink','pink','pink','pink'],['red','red','red','red','red','red','red']],
				//													Border Color
				[['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray']],
				//Labels
				['Kris','Ben','John','Michle','James','Tim','Ryan'],
        ,true
			],
      ['r6-c3', 'My Title','horizontalBar',30,150,
			//     			Data Points
				[[10,50,20,39,54,44,45],[14,44,57,99,54,57,88],[11,22,57,88,99,55,44],[45,54,74,44,12,12,45],],
				//Labels
				['Dataset 1','Dataset 2','Dataset 3','Dataset 4'],
				//												Background Color
				[['green','green','green','green','green','green','green'],['orange','orange','orange','orange','orange','orange','orange'],['pink','pink','pink','pink','pink','pink','pink'],['red','red','red','red','red','red','red']],
				//													Border Color
				[['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray'],['gray','gray','gray','gray','gray','gray','gray']],
				//Labels
				['Kris','Ben','John','Michle','James','Tim','Ryan'],
        ,true
			],



		],

		measures :
		[

			["r1-c5","","Marked Q1: ",500," m","220",";background-color:purple;border:purple;color:white;"],  //Measure Details
			["r1-c5","btn-dark","Marked Q2: ",500," m","220",""],  //Measure Details
			["r1-c6","","Marked Q3: ",200," cm","220",";background-color:crimson;border:crimson;color:white;"],  //Measure Details				webDevObjects[0].measures[0][1] = "Qty"
			["r1-c6","btn-success","Marked Q4: ",500," m","220",""],  //Measure Details
			["r1-c7","","Qty: ",500," m","220",";background-color:purple;border:purple;color:white;"],  //Measure Details
			["r1-c7","btn-primary","Qty: ",500," m","220",""],  //Measure Details
			["r1-c8","","Qty: ",500," m","220",";background-color:orange;border:orange;color:white;"],  //Measure Details
			["r1-c8","btn-secondary","Qty: ",500," m","220",""],  //Measure Details
			["r1-c9","","Qty: ",200," cm","220",";background-color:pink;border:pinkcolor:white;color:white;"],  //Measure Details				webDevObjects[0].measures[0][1] = "Qty"
			["r1-c9","btn-warning","Qty: ",500," m","220",";color:white;"],  //Measure Details
			["r1-c10","","Qty: ",500," m","220",";background-color:black;border:black;color:white;"],  //Measure Details
			["r1-c10","btn-info","Qty: ",500," m","220",""],  //Measure Details
			["r1-c11","","Qty: ",500," m","220",";background-color:teal;border:teal;color:white;"],  //Measure Details
			["r1-c11","btn-danger","Qty: ",500," m","220",""],  //Measure Details
		],

		navBar :
		[

			[
				["Heading 1","Heading 2"], //NavBar Dropdown Headings	webDevObjects[0].navBar[0][0][0] = "Heading 1"
			],
			[
				["Heading 1 Link 1","#"], //NavBar Dropdown Link		webDevObjects[0].navBar[1][0][0] = "Heading 1 Link 1"
				["Heading 1 Link 2","#"], //NavBar Dropdown Link		webDevObjects[0].navBar[1][1][1] = "#"
			],
			[
				["Heading 2 Link 1","#"], //NavBar Dropdown Link
				["Heading 2 Link 2","#"], //NavBar Dropdown Link
			],

		],

		tables :
		[
			["r3-c1",								//webDevObjects[0].tables[0][0] = "r11-c0"
					["table table-hover", "yes", 49.8, "padding:10px", "item "],		//webDevObjects[0].tables[0][1][0] = "table table-dark-hover"

					["Name","Age","Gender","DoB"], 	//Table Header 			webDevObjects[0].tables[0][2][0] = "Name"
					["Kris","36","M","12/3/1990"],	//Table body Row 1		webDevObjects[0].tables[0][3][2] = "M"
					["Tim","28","M","4/12/1987"],		//Table body Row 2		webDevObjects[0].tables[0][4][1] = 28
					["Jane","32","F","3/10/200"],		//Table body Row 3
			],
			["r3-c2",								//webDevObjects[0].tables[1][0] = "r11-c1"
					["table table-dark table-hover", "no", 49.8,  "padding:10px", "item "], //webDevObjects[0].tables[1][1][1] = "no"

					["Name","Age","Gender","DoB"], 	//Table Header 			webDevObjects[0].tables[1][2][0] = "Name"
					["Kris","36","M","12/3/1990"],	//Table body Row 1		webDevObjects[0].tables[1][3][0] = "Kris"
					["Tim","28","M","4/12/1987"],		//Table body Row 2		webDevObjects[0].tables[1][4][1]  = 28
					["Jane","32","F","3/10/200"],		//Table body Row 3
			],
		],

		text :
		[
			['r7-c1',100,400,"margin-left:0%; text-align:left;","center",
			`<div class='item'>
				<h1 class="">Excel Web Development Add-In</h1>
				<p>
				  Below are numerious tools to assist with the Excel Web Development Add-in.
				  <br>
				  <a href="">Download</a> - <a href="#">Documentation</a>
				</p>
				<table>
				  <tbody><tr>
					<td>
					  <p>
						</p><h4>Examples:</h4>
						<ul>
						  <li>
							<a href="#">Resume</a>
						  </li>
						  <li>
							<a href="#">Artical</a>
						  </li>
						  <li>
							<a href="#">Dashboard</a>
						  </li>
						  <li>
							<a href="#">Profile</a>
						  </li>
						  <li>
							<a href="#">Navigation</a>
						  </li>
						  <li>
							<a href="#">Invite</a>
						  </li>
						  <li>
							<a href="#">Information</a>
						  </li>
						</ul>
					  <p></p>
					</td>
					<td style="padding-left:80px; padding-top: 0px; vertical-align:top;width:300px;">
					  <p>
						</p><h4>Documentation:</h4>
						<ul>
						  <li>
							<a href="#">Overview</a>
						  </li>
						  <li>
							<a href="#">Text</a>
						  </li>
						  <li>
							<a href="#">Tables</a>
						  </li>
						  <li>
							<a href="#">Blur Filters</a>
						  </li>
						  <li>
							<a href="#">Images</a>
						  </li>
						  <li>
							<a href="#">Charts</a>
						  </li>
						  <li>
							<a href="#">Metrics</a>
						  </li>
						</ul>
					  <p></p>
					</td>
				  </tr>
				</tbody></table>
				</div>
			`
		],

		],
		iframe : [
			['r8-c1',100,500,'margin-top:50px; text-align:left;','center','no','https://kjtakke.github.io/','hidden',0],
		],
		images : [
			['r9-c1',30,,'margin-top:50px; text-align:left;','center','https://kjtakke.github.io/images/background/Blue-Clouds.jpg'],

		]

	}
];
