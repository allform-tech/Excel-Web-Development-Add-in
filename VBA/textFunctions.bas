Attribute VB_Name = "textFunctions"
Sub tstTextFunction1()
    Dim Text As String
    Dim tmpAry, output As Variant

    Text = "Hello my name is {% Sheet1!A1:A2 %} and it is nice to meet you." & vbCr
    Text = Text & "My last neme is {% Sheet1!A2 %}."
    tmpAry = Split(Text, vbCr)
    output = INSERT_TEXT_FROM_CELL_REF_INTO_ARRAY(tmpAry)
End Sub

Sub tstTextFunction2()
    Dim Text As String
    Dim output As String
    Dim tmpAry As Variant

    Text = "Hello my name is {% Sheet1!A1:A2 %} and it is nice to meet you."
    output = INSERT_TEXT_FROM_CELL_REF(Text)
End Sub

Function INSERT_TEXT_FROM_CELL_REF_INTO_ARRAY(ary As Variant) As Variant
    ' This Function takes a Cell Reference and Concatinates it to/within a String
    ' Similar to the Concatenate function, however, is more dynamic using Cell Ranges
    ' Format: {% SheetName!CelRange %}
    ' Main Points:
    '   - Must have a space after {% and before %}
    '   - Must have an ! between the Sheet and Cell Range
    '   - Must use the standard Cell Range Format:
    '       - Single Cell:  A1 or $A1 or A$1 or $A$1
    '       - Cell Range:   A1:A2
    '
    ' Examples: (Sheet1 - A1 = "John", A2 = "Smith")
    '   - Formula: =INSERT_TEXT_FROM_CELL_REF("Hello my name is {% Sheet1!A1 %} {% Sheet1!A2 %}.")
    '   - Returns: Hello my name is John Smith.
    '
    '   - Formula: =INSERT_TEXT_FROM_CELL_REF("Hello my name is {% Sheet1!A1:A2 %}.")
    '   - Returns: Hello my name is John Smith.

    Dim i, j, k, L, R As Long
    Dim line As String
    Dim lineArray, celRefArray, refArray, trackCellRef, cellRangeText As Variant
    ReDim trackCellRef(0 To 100, 0 To 4)
    Dim tracRefCounter As Long
    tracRefCounter = 0
    On Error Resume Next
    
    For i = 0 To UBound(ary)
        lineArray = Split(ary(i), "{%")
        For j = 0 To UBound(lineArray)
            If InStr(1, lineArray(j), "%}", 0) Then
                celRefArray = Split(lineArray(j), "%}")
                For k = 0 To UBound(celRefArray)
                    If InStr(1, celRefArray(k), "!", 0) Then
                        refArray = Split(celRefArray(k), "!")
                        For L = 0 To UBound(refArray)
                           refArray(L) = Trim(refArray(L))
                        Next L
                        cellRangeText = Worksheets(refArray(0)).Range(refArray(1)).Value
                        
                        If VarType(cellRangeText) <> vbString Then
                            line = ""
                            For R = 1 To UBound(cellRangeText)
                                If R = UBound(cellRangeText) Then
                                    line = line & cellRangeText(R, 1)
                                Else
                                    line = line & cellRangeText(R, 1) & " "
                                End If
                            Next R
                        Else
                            line = cellRangeText
                        End If
                        trackCellRef(tracRefCounter, 0) = "{% " & refArray(0) & "!" & refArray(1) & " %}"
                        trackCellRef(tracRefCounter, 1) = i
                        trackCellRef(tracRefCounter, 2) = j
                        trackCellRef(tracRefCounter, 3) = k
                        trackCellRef(tracRefCounter, 4) = line
                        tracRefCounter = tracRefCounter + 1
                    End If
                Next k
            End If
        Next j
    Next i
    
    For i = 0 To UBound(ary)
        For j = 0 To UBound(trackCellRef)
            If trackCellRef(j, 0) = "" Then Exit For
            ary(i) = Replace(ary(i), trackCellRef(j, 0), trackCellRef(j, 4))
        Next j
    Next i
    INSERTCELLREFFROMARRAY = ary
End Function

Function INSERT_TEXT_FROM_CELL_REF(str As String) As String
    ' This Function takes a Cell Reference and Concatinates it to/within a String
    ' Similar to the Concatenate function, however, is more dynamic using Cell Ranges
    ' Format: {% SheetName!CelRange %}
    ' Main Points:
    '   - Must have a space after {% and before %}
    '   - Must have an ! between the Sheet and Cell Range
    '   - Must use the standard Cell Range Format:
    '       - Single Cell:  A1 or $A1 or A$1 or $A$1
    '       - Cell Range:   A1:A2
    '
    ' Examples: (Sheet1 - A1 = "John", A2 = "Smith")
    '   - Formula: =INSERT_TEXT_FROM_CELL_REF("Hello my name is {% Sheet1!A1 %} {% Sheet1!A2 %}.")
    '   - Returns: Hello my name is John Smith.
    '
    '   - Formula: =INSERT_TEXT_FROM_CELL_REF("Hello my name is {% Sheet1!A1:A2 %}.")
    '   - Returns: Hello my name is John Smith.
    
    Dim i, j, k, L, R As Long
    Dim line As String
    Dim lineArray, celRefArray, refArray, trackCellRef, cellRangeText As Variant
    ReDim trackCellRef(0 To 100, 0 To 4)
    Dim tracRefCounter As Long
    tracRefCounter = 0
    On Error Resume Next
    lineArray = Split(str, "{%")
    For j = 0 To UBound(lineArray)
        If InStr(1, lineArray(j), "%}", 0) Then
            celRefArray = Split(lineArray(j), "%}")
            For k = 0 To UBound(celRefArray)
                If InStr(1, celRefArray(k), "!", 0) Then
                    refArray = Split(celRefArray(k), "!")
                    For L = 0 To UBound(refArray)
                       refArray(L) = Trim(refArray(L))
                    Next L
                    cellRangeText = Worksheets(refArray(0)).Range(refArray(1)).Value
                    
                    If VarType(cellRangeText) <> vbString Then
                        line = ""
                        For R = 1 To UBound(cellRangeText)
                                If R = UBound(cellRangeText) Then
                                    line = line & cellRangeText(R, 1)
                                Else
                                    line = line & cellRangeText(R, 1) & " "
                                End If
                        Next R
                    Else
                        line = cellRangeText
                    End If
                    trackCellRef(tracRefCounter, 0) = "{% " & refArray(0) & "!" & refArray(1) & " %}"
                    trackCellRef(tracRefCounter, 1) = i
                    trackCellRef(tracRefCounter, 2) = j
                    trackCellRef(tracRefCounter, 3) = k
                    trackCellRef(tracRefCounter, 4) = line
                    tracRefCounter = tracRefCounter + 1
                End If
            Next k
        End If
    Next j
    For j = 0 To UBound(trackCellRef)
        If trackCellRef(j, 0) = "" Then Exit For
        str = Replace(str, trackCellRef(j, 0), trackCellRef(j, 4))
    Next j
    INSERT_TEXT_FROM_CELL_REF = str
End Function





