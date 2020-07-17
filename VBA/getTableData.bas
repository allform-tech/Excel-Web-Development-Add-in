Attribute VB_Name = "getTableData"
'Get Table Metadata and Data
    'List of Tables within the current Workbook/Project
        Function listTables() As Variant
            Dim tbl As ListObject
            Dim WS As Worksheet
            Dim i As Single, j As Single
            Dim tableNameArray As Variant
            Dim tempArray As Variant
            ReDim tableNameArray(1 To 1000)
            
            On Error GoTo en:
            
            i = 1
            For Each WS In Worksheets
                For Each tbl In WS.ListObjects
                    tableNameArray(i) = tbl.Name
                    i = i + 1
                Next tbl
            Next WS
            For i = 1 To UBound(tableNameArray)
                If tableNameArray(i) = "" Then Exit For
            Next i
            i = i - 1
            ReDim tempArray(1 To i)
            For j = 1 To i
                tempArray(i) = tableNameArray(i)
            Next j
            listTables = tempArray
            GoTo enn:
en:
            ReDim tempArray(1 To 1)
            tempArray(1) = ""
            tempTableArray = tempArray
enn:
        End Function
    
    'List of Table Columns of a selected Table  within the current Workbook/Project
        Function listTablesCColumns(tableName As String) As Variant
            Dim tbl As ListObject
            Dim WS As Worksheet
            Dim i As Single, j As Single
            Dim tempArray, cellNames As Variant
            ReDim tempArray(1 To 1000)
            
            On Error GoTo en:
            
            i = 1
            For Each WS In Worksheets
                For Each tbl In WS.ListObjects
                    If tableName = tbl.Name Then
                        For j = 1 To tbl.Range.Columns.Count
                            tempArray(j) = tbl.Range.Cells(1, j)
                        Next j
                        i = i + 1
                    End If
                Next tbl
            Next WS
            For i = 1 To UBound(tempArray)
                If tempArray(i) = "" Then Exit For
            Next i
            i = i - 1
            ReDim cellNames(1 To i)
            For j = 1 To i
                cellNames(j) = tempArray(j)
            Next j
            GoTo enn:
en:
            ReDim tempArray(1 To i)
            listTablesCColumns = tempArray
enn:
            listTablesCColumns = cellNames
        End Function
    
    'Table Column Values into an Array from a selected Columns within a Selected within the current Workbook/Project
        Function getTableColumnValues(tableName, ColumnName As String) As Variant
            Dim tbl As ListObject
            Dim WS As Worksheet
            Dim i As Single, j As Single
            Dim tempArray As Variant
            On Error GoTo en:
            i = 1
            For Each WS In Worksheets
                For Each tbl In WS.ListObjects
                    If tableName = tbl.Name Then
                        For j = 1 To tbl.Range.Columns.Count
                            
                            If ColumnName = tbl.Range.Cells(1, j) Then
                                  tempArray = Worksheets(WS.Name).ListObjects(tableName).ListColumns(j).Range
                            End If
                            
                        Next j
                        i = i + 1
                    End If
                Next tbl
            Next WS
            getTableColumnValues = tempArray
            GoTo enn:
en:
            ReDim tempArray(1 To 1)
            tempArray(1) = ""
            getTableColumnValues = tempArray
enn:
        End Function
    
    'Get Worksheet name for a Table within the current Workbook/Project
        Function getWorksheetNameFromTable(tableName As String) As String
            Dim tbl As ListObject
            Dim WS As Worksheet
            Dim i As Single, j As Single
            Dim tempArray As Variant
            
            On Error GoTo en:
            
            i = 1
            For Each WS In Worksheets
                For Each tbl In WS.ListObjects
                    If tableName = tbl.Name Then
                        getWorksheetNameFromTable = WS.Name
                    End If
                Next tbl
            Next WS
            GoTo enn:
en:
            getWorksheetNameFromTable = ""
enn:
        End Function
