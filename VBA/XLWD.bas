Attribute VB_Name = "XLWD"
'.XLWD File Controls
Public TempArray1 As Variant
Public i As Single

    'Array Dementions
        Function arrayDimentionCounter(index As Variant) As Integer
        'This Function Counts the Columns/Dimentions in an Array
        'index is the input array
        
            On Error GoTo LC:
            For L = 1 To 100
                TempVar = index(1, L)
            Next L
LC:
            L = L - 1
            On Error GoTo 0
            arrayDimentionCounter = L
        End Function
        
    'File Functions/Properties
        Public Sub clWDWriteToFile(Str As String)
            On Error Resume Next
            If Str = "" Then
                WDFileName = InputBox("File Name")
                With Application.FileDialog(msoFileDialogFolderPicker)
                    If .Show = -1 Then ' if OK is pressed
                        WDFilePath = .SelectedItems(1)
                    End If
                End With
                WDFileName = WDFilePath & "\" & Str & ".xlwd"
            End If
            
            On Error GoTo en:
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set A = fs.CreateTextFile(WDFileName, True)
            A.WriteLine WDXLWDStr
            A.Close
en:
        End Sub
        
    'Open ane import an .XLWD File
        Public Sub clWDOpenFile()
            Dim textline, myTempFile As String
            On Error GoTo en:
            WDFileName = Application.GetOpenFilename(FileFilter:="XLWD File (*.xlwd), *.xlwd")
            Open WDFileName For Input As #1
            Application.Wait (Now + TimeValue("0:00:05"))
            Do Until EOF(1)
                Line Input #1, textline
                WDXLWDStr = WDXLWDStr & textline
            Loop
en:
            Close #1
        End Sub
            
    'Read String from an Array
        'This uses Deliminated Strings (.init.SD,ED, OD) to (init.WD...'arrays')
        Public Sub clWDRead(Str As String, Del As String, ary As Variant)
            On Error Resume Next
            tmpArray1 = Split(Str, Del)
            For i = 1 To UBound(tmpArray) + 1
                ary(i + 1) = tmpArray1(i)
            Next i
        End Sub
    
    'Get String from an Array
        'This uses Deliminated Strings (.init.SD,ED, OD) to (init.WD...'arrays')
        Public Sub clWDGet(ary As Variant, cnt As Integer, Str As String)
            On Error Resume Next
            Str = ary(cnt)
        End Sub

    'Write a String to an Array
        'This uses Deliminated Strings (.init.SD,ED, OD) to (init.WD...'arrays')
        Public Sub clWDWrite(arry As Variant, Str As String)
            On Error Resume Next
            Str = ""
            For i = 1 To UBound(arry)
                If i = 1 Then
                    Str = arry(i)
                Else
                    Str = Str & arry(i)
                End If
            Next i
        End Sub
