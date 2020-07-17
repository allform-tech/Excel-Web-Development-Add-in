Attribute VB_Name = "ChartFunctions"
'Chart Functions
    'These Function create JSON/JavaScript formatted for Charts.JS
    
    'TEST_SUB###############################################################################################
        Sub TestChartFunctions()
            Dim colorsAry As Variant
            Dim colorStr As String
            colorStr = "Red, Blue"
            colorsAry = chartColors(colorStr)
            chartTableToJSON tableName:="Table1", address:="r5-c1", title:="My Title", chartType:="horizontalBar", chartWidth:="30", chartHeight:="150", colors:=colorsAry, WS:=Worksheets("Sheet1"), XaxiTitle:="Names", YaxixTitle:="Numbers"
        End Sub
    '#######################################################################################################
    
    'Chart Color String into an Array
        Function chartColors(colorStr As String) As Variant
            Dim colorsSplit As Variant
            colorStr = Replace(colorStr, " ", "")
            chartColors = Split(colorStr, ",")
        End Function
     
    'Creats a JSON/JavaScript Array
        Function chartTableToJSON(tableName As String, address As String, title As String, chartType As String, chartWidth As String, chartHeight As String, colors As Variant, WS As Worksheet, XaxiTitle As String, YaxixTitle As String) As String
            Dim Data As Variant
            Dim tb As Object
            Dim ColCount As Integer
            Dim i, j As Long
            Dim dataPoints As String
            Dim Lables As String
            Dim Color As String
        
            tbn = "Table1"
            ColCount = WS.ListObjects(tbn).ListColumns.Count
            Set tb = WS.ListObjects(tbn)
            Data = tb.Range.Value2
        
            'Data Points
            dataPoints = "["
            For j = 2 To ColCount
                dataPoints = dataPoints & "["
                For i = 2 To UBound(Data)
                    If i = UBound(Data) Then
                        dataPoints = dataPoints & Data(i, j)
                    Else
                        dataPoints = dataPoints & Data(i, j) & ","
                    End If
                Next i
                If j = ColCount Then
                    dataPoints = dataPoints & "]"
                Else
                    dataPoints = dataPoints & "],"
                End If
            Next j
            dataPoints = dataPoints & "]"
            'Debug.Print (dataPoints)
        
            'Lables
            Lables = "["
            For i = 1 To ColCount
                If i = ColCount Then
                    Lables = Lables & "'" & Data(1, i) & "'"
                Else
                    Lables = Lables & "'" & Data(1, i) & "',"
                End If
            Next i
            Lables = Lables & "]"
            'Debug.Print (Lables)
        
            'Colors
            Color = "["
            For i = 0 To UBound(colors)
                Color = Color & "["
                For j = 2 To UBound(Data)
                    If j = UBound(Data) Then
                        Color = Color & "'" & colors(i) & "'"
                    Else
                        Color = Color & "'" & colors(i) & "',"
                    End If
                Next j
                If i = UBound(colors) Then
                    Color = Color & "]"
                Else
                    Color = Color & "],"
                End If
                    
            Next i
            Color = Color & "],"
            
            Color = Color & "["
            For i = 0 To UBound(colors)
                Color = Color & "["
                For j = 2 To UBound(Data)
                    If j = UBound(Data) Then
                        Color = Color & "'gray'"
                    Else
                        Color = Color & "'gray',"
                    End If
                Next j
                If i = UBound(colors) Then
                    Color = Color & "]"
                Else
                    Color = Color & "],"
                End If
                    
            Next i
            Color = Color & "]"
            'Debug.Print (Color)
            
            chartTableToJSON = "['" & address & "'," & "'" & title & "'," & "'" & chartType & "'," & "'" & chartWidth & "'," & "'" & chartHeight & "'," & dataPoints & "," & Lables & "," & Color & "," & Lables & ",'" & XaxiTitle & "','" & YaxixTitle & "'],"
            'Debug.Print (chartTableToJSON)
        End Function
        
