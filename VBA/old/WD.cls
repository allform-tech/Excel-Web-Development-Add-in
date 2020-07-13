VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module
Public MD, ED, SD, OD As String
Public WDPages, WDPage, WDNavBar, WDNavBarHeadings, WDNavBarLinks, WDMeasures, WDMeasure, WDTables, WDTable, WDCharts, WDChart, WDTexts, WDText, WDImages, WDImage, WDiFrames, WDiFrame As Variant
Public WDXLWDStr, WDPagesStr, WDPageStr, WDNavBarStr, WDNavBarHeadingsStr, WDNavBarLinksStr, WDMeasuresStr, WDMeasureStr, WDTablesStr, WDTableStr, WDChartsStr, WDChartStr, WDTextsStr, WDTextStr, WDImagesStr, WDImageStr, WDiFramesStr, WDiFrameStr As String
Public WDCSS, WDJavaScript As String
Public WDPagesCnt, WDPageCnt, WDNavBarCnt, WDNavBarHeadingsCnt, WDNavBarLinksCnt, WDMeasureCnt, WDMeasuresCnt, WDTablesCnt, WDTableCnt, WDChartsCnt, WDChartCnt, WDTextsCnt, WDTextCnt, WDImagesCnt, WDImageCnt, WDiFramesCnt, WDiFrameCnt As Integer
Public tmpArray1, tmpArray2, tmpArray3 As Variant
Public tmpStr1, tmpStr2, tmpStr3 As String
Public Counter1, Counter2, Counter3 As Integer
Public Long1, Long2, Long3 As Long
Public WDFilePath, WDFileName As String

Private Sub Class_Initialize()
    MD = "C:2s3ZpnC8A,S{T/H)SWZ'24\mmuv3Egb%M/QDA86AUer`zn=Z'u@8;tTry{gqYa5VK`.(y9LvR~&PTs\=RQW2<}A@s:#Lr>V(W2;-s4~$Wbq9~NT'},Q.bm*Rj'7Nve" 'Main Data
    ED = "^cv8CR(U<3wbvh2>*ee.bK'6b:ZQqwj@s#?EQLhU:U>4Q:^[pALeg,/a+/]R$ZuG48_rTuC9)kQyKUZUe:#jv_.DK$3fm}g%*]~/,`A&$V;5;[yAz$BPw}TV`yXqB~G%" 'Element Data
    SD = "C3`:j~52,`/Bt:b:y~y[^PRtznp8^XE-vSA:93=#LjLR>M~8%%$jB<x<G;5)*cB4sPFV9#}/Rd5E8^)<@NazNjEX8S~ND&Qk/Mt_n&3?Y5Dbxx[GNG#En,GZ&k-3RhD:" 'Sub Data
    OD = "2@Nf<S>GH3VQEvZY+GSw:*-@(?%DV_h{#6AZp'6{DL`~w.cM<U$;8e'BqhyCpSZ2WQ'%}]N+6]xf`pT@,_b@a-g2[]*Hh!8}U4ngnYFVWgyV$y?::]D&bBw[fWD}Y~GF" 'Option Data
    
    WDPagesCnt = 100
    WDPageCnt = 4
    WDNavBarCnt = 3
    WDNavBarHeadingsCnt = 20
    WDNavBarLinksCnt = 50
    WDMeasuresCnt = 100
    WDMeasureCnt = 8
    WDTablesCnt = 100
    WDTableCnt = 10
    WDChartsCnt = 100
    WDChartCnt = 11
    WDTextsCnt = 100
    WDTextCnt = 7
    WDImagesCnt = 100
    WDImageCnt = 7
    WDiFramesCnt = 100
    WDiFrameCnt = 7

    ReDim WDPages(1 To WDPagesCnt)
    ReDim WDPage(1 To WDPageCnt)
    ReDim WDNavBar(1 To WDNavBarCnt)
    ReDim WDNavBarHeadings(1 To WDNavBarHeadingsCnt)
    ReDim WDNavBarLinks(1 To WDNavBarHeadingsCnt)
    ReDim WDMeasures(1 To WDMeasuresCnt)
    ReDim WDMeasure(1 To WDMeasureCnt)
    ReDim WDTables(1 To WDTablesCnt)
    ReDim WDTable(1 To WDTableCnt)
    ReDim WDCharts(1 To WDChartsCnt)
    ReDim WDChart(1 To WDChartCnt)
    ReDim WDTexts(1 To WDTextsCnt)
    ReDim WDText(1 To WDTextCnt)
    ReDim WDImages(1 To WDImagesCnt)
    ReDim WDImage(1 To WDImageCnt)
    ReDim WDiFrames(1 To WDiFramesCnt)
    ReDim WDiFrame(1 To WDiFrameCnt)

End Sub


'Get Array from String
Public Sub clWDXLWDRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, MD)
    For i = 1 To UBound(TmpArray) + 1
        WDPages(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDPagesRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, ED)
    For i = 1 To UBound(TmpArray) + 1
        WDPages(i + 1) = tmpArray1(i)
    Next i
    WDCSS = WDPages(UBound(WDPages) - 1)
    WDJavaScript = WDPages(UBound(WDPages))
End Sub

Public Sub clWDPageRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, SD)
    For i = 1 To UBound(TmpArray) + 1
        WDPage(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDNavBarRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, ED)
    For i = 1 To UBound(TmpArray) + 1
        WDNavBar(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDNavBarHeadingRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, SD)
    For i = 1 To UBound(TmpArray) + 1
        WDNavBarHeadings(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDNavBarLinksRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, OD)
    For i = 1 To UBound(TmpArray) + 1
        WDNavBarLinks(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDMeasuresRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, ED)
    For i = 1 To UBound(TmpArray) + 1
        WDMeasures(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDMeasureRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, SD)
    For i = 1 To UBound(TmpArray) + 1
        WDMeasure(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDTablesRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, ED)
    For i = 1 To UBound(TmpArray) + 1
        WDTables(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDTableRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, SD)
    For i = 1 To UBound(TmpArray) + 1
        WDTable(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDTChartsRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, ED)
    For i = 1 To UBound(TmpArray) + 1
        WDTCharts(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDChartRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, SD)
    For i = 1 To UBound(TmpArray) + 1
        WDChart(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDTTextsRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, ED)
    For i = 1 To UBound(TmpArray) + 1
        WDTTexts(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDTextRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, SD)
    For i = 1 To UBound(TmpArray) + 1
        WDText(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDImagesRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, ED)
    For i = 1 To UBound(TmpArray) + 1
        WDImages(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDImageRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, SD)
    For i = 1 To UBound(TmpArray) + 1
        WDImage(i + 1) = tmpArray1(i)
    Next i
End Sub

Public Sub clWDiFramesRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, ED)
    For i = 1 To UBound(TmpArray) + 1
        WDiFrames(i + 1) = tmpArray1(i)
    Next i
End Sub


Public Sub clWDiFrameRead(Str As String)
    On Error Resume Next
    tmpArray1 = Split(Str, SD)
    For i = 1 To UBound(TmpArray) + 1
        WDiFrame(i + 1) = tmpArray1(i)
    Next i
End Sub


'Get String from Array
Public Sub clWDPagesGet(cnt As Integer)
    On Error Resume Next
    WDPagesStr = WDPages(cnt)
End Sub

Public Sub clWDPageGet(cnt As Integer)
    On Error Resume Next
    WDPageStr = WDPage(cnt)
End Sub

Public Sub clWDNavBarGet(cnt As Integer)
    On Error Resume Next
    WDNavBarStr = WDNavBar(cnt)
End Sub

Public Sub clWDNavBarHeadingGet(cnt As Integer)
    On Error Resume Next
    WDNavBarHeadingsStr = WDNavBarHeading(cnt)
End Sub

Public Sub clWDNavBarLinksGet(cnt As Integer)
    On Error Resume Next
    WDNavBarLinksStr = WDNavBarLinks(cnt)
End Sub

Public Sub clWDMeasuresGet(cnt As Integer)
    On Error Resume Next
    WDMeasuresStr = WDMeasures(cnt)
End Sub

Public Sub clWDMeasureGet(cnt As Integer)
    On Error Resume Next
    WDMeasureStr = WDMeasure(cnt)
End Sub

Public Sub clWDTablesGet(cnt As Integer)
    On Error Resume Next
    WDTablesStr = WDTables(cnt)
End Sub

Public Sub clWDTableGet(cnt As Integer)
    On Error Resume Next
    WDTableStr = WDTable(cnt)
End Sub

Public Sub clWDTChartsGet(cnt As Integer)
    On Error Resume Next
    WDChartsStr = WDCharts(cnt)
End Sub

Public Sub clWDChartGet(cnt As Integer)
    On Error Resume Next
    WDChartStr = WDChart(cnt)
End Sub

Public Sub clWDTTextsGet(cnt As Integer)
    On Error Resume Next
    WDTextsStr = WDTexts(cnt)
End Sub

Public Sub clWDTextGet(cnt As Integer)
    On Error Resume Next
    WDTextStr = WDText(cnt)
End Sub

Public Sub clWDImagesGet(cnt As Integer)
    On Error Resume Next
    WDImagesStr = WDImages(cnt)
End Sub

Public Sub clWDImageGet(cnt As Integer)
    On Error Resume Next
    WDImageStr = WDImage(cnt)
End Sub

Public Sub clWDiFramesGet(cnt As Integer)
    On Error Resume Next
    WDiFramesStr = WDiFrames(cnt)
End Sub


Public Sub clWDiFrameGet(cnt As Integer)
    On Error Resume Next
    WDiFrameStr = WDiFrame(cnt)
End Sub


'Write String From Array
Public Sub clWDXLWDWrite(arry As Variant)
    On Error Resume Next
    WDXLWDStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDXLWDStr = arry(i)
        Else
            WDXLWDStr = WDXLWDStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDPagesWrite(arry As Variant)
    On Error Resume Next
    WDPagesStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDPagesStr = arry(i)
        Else
            WDPagesStr = WDPagesStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDPageWrite(arry As Variant)
    On Error Resume Next
    WDPageStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDPageStr = arry(i)
        Else
            WDPageStr = WDPageStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDNavBarWrite(arry As Variant)
    On Error Resume Next
    WDNavBarStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDNavBarStr = arry(i)
        Else
            WDNavBarStr = WDNavBarStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDNavBarHeadingWrite(arry As Variant)
    On Error Resume Next
    WDNavBarHeadingsStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDNavBarHeadingsStr = arry(i)
        Else
            WDNavBarHeadingsStr = WDNavBarHeadingsStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDNavBarLinksWrite(arry As Variant)
    On Error Resume Next
    WDNavBarLinksStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDNavBarLinksStr = arry(i)
        Else
            WDNavBarLinksStr = WDNavBarLinksStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDMeasuresWrite(arry As Variant)
    On Error Resume Next
    WDMeasuresStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDMeasuresStr = arry(i)
        Else
            WDMeasuresStr = WDMeasuresStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDMeasureWrite(arry As Variant)
    On Error Resume Next
    WDMeasureStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDMeasureStr = arry(i)
        Else
            WDMeasureStr = WDMeasureStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDTablesWrite(arry As Variant)
    On Error Resume Next
    WDTablesStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDTablesStr = arry(i)
        Else
            WDTablesStr = WDTablesStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDTableWrite(arry As Variant)
    On Error Resume Next
    WDTableStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDTableStr = arry(i)
        Else
            WDTableStr = WDTableStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDTChartsWrite(arry As Variant)
    On Error Resume Next
    WDChartsStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDChartsStr = arry(i)
        Else
            WDChartsStr = WDChartsStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDChartWrite(arry As Variant)
    On Error Resume Next
    WDChartStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDChartStr = arry(i)
        Else
            WDChartStr = WDChartStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDTTextsWrite(arry As Variant)
    On Error Resume Next
    WDTextsStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDTextsStr = arry(i)
        Else
            WDTextsStr = WDTextsStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDTextWrite(arry As Variant)
    On Error Resume Next
    WDTextStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDTextStr = arry(i)
        Else
            WDTextStr = WDTextStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDImagesWrite(arry As Variant)
    On Error Resume Next
    WDImagesStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDImagesStr = arry(i)
        Else
            WDImagesStr = WDImagesStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDImageWrite(arry As Variant)
    On Error Resume Next
    WDImageStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDImageStr = arry(i)
        Else
            WDImageStr = WDImageStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDiFramesWrite(arry As Variant)
    On Error Resume Next
    WDiFramesStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDiFramesStr = arry(i)
        Else
            WDiFramesStr = WDiFramesStr & arry(i)
        End If
    Next i
End Sub

Public Sub clWDiFrameWrite(arry As Variant)
    On Error Resume Next
    WDiFrameStr = ""
    For i = 1 To UBound(arry)
        If i = 1 Then
            WDiFrameStr = arry(i)
        Else
            WDiFrameStr = WDiFrameStr & arry(i)
        End If
    Next i
End Sub


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
    Set a = fs.CreateTextFile(WDFileName, True)
    a.WriteLine WDXLWDStr
    a.Close

en:

End Sub

Public Sub clWDWriteToNewFile()
    On Error Resume Next

    WDFileName = InputBox("File Name")
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            WDFilePath = .SelectedItems(1)
        End If
    End With
    WDFileName = WDFilePath & "\" & Str & ".xlwd"
    
    On Error GoTo en:
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(WDFileName, True)
    a.WriteLine WDXLWDStr
    a.Close

en:

End Sub

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



