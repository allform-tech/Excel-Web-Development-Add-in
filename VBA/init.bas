Attribute VB_Name = "init"
'Standard Modlue

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
Public LoadedChart, LoadedTable, LoadedText, LoadedMeasure, LoadediFrame, LoadedImage As Integer
Public ElementListBoxItemCount As Integer
Public SelectedElement
Public newElement As Boolean
Public tableNameArray As Variant

'Initilise Form and Public Variables
    Sub init()
        Main.Show
    End Sub
    
    Sub InitiateVariables()
        'MD = "C:2s3ZpnC8A,S{T/H)SWZ'24\mmuv3Egb%M/QDA86AUer`zn=Z'u@8;tTry{gqYa5VK`.(y9LvR~&PTs\=RQW2<}A@s:#Lr>V(W2;-s4~$Wbq9~NT'},Q.bm*Rj'7Nve" 'Main Data 'Not uesd
        ED = "^cv8CR(U<3wbvh2>*ee.bK'6b:ZQqwj@s#?EQLhU:U>4Q:^[pALeg,/a+/]R$ZuG48_rTuC9)kQyKUZUe:#jv_.DK$3fm}g%*]~/,`A&$V;5;[yAz$BPw}TV`yXqB~G%" 'Element Data
        SD = "C3`:j~52,`/Bt:b:y~y[^PRtznp8^XE-vSA:93=#LjLR>M~8%%$jB<x<G;5)*cB4sPFV9#}/Rd5E8^)<@NazNjEX8S~ND&Qk/Mt_n&3?Y5Dbxx[GNG#En,GZ&k-3RhD:" 'Sub Data
        OD = "2@Nf<S>GH3VQEvZY+GSw:*-@(?%DV_h{#6AZp'6{DL`~w.cM<U$;8e'BqhyCpSZ2WQ'%}]N+6]xf`pT@,_b@a-g2[]*Hh!8}U4ngnYFVWgyV$y?::]D&bBw[fWD}Y~GF" 'Option Data
        
        'WDPagesCnt = 100                                    'Not uesd
        WDPageCnt = 5
        WDNavBarCnt = 2
        WDNavBarHeadingsCnt = 20
        WDNavBarLinksCnt = 50
        WDMeasuresCnt = 100
        WDMeasureCnt = 7
        WDTablesCnt = 100
        WDTableCnt = 8
        WDChartsCnt = 100
        WDChartCnt = 11
        WDTextsCnt = 100
        WDTextCnt = 6
        WDImagesCnt = 100
        WDImageCnt = 6
        WDiFramesCnt = 100
        WDiFrameCnt = 6
        
        ElementListBoxItemCount = 0
    
        'ReDim WDPages(1 To WDPagesCnt)                      'Not uesd
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
        
        'Is the user ammending or adding new data
        '-1 = New >=0 = relevent Array(x)
        LoadedForms = -1
        LoadedChart = -1
        LoadedTable = -1
        LoadedText = -1
        LoadedMeasure = -1
        LoadediFrame = -1
        LoadedImage = -1
        
        'Selected List Box Element in the Main Form
        '-1 = Null >=0 = relevent Array(x)
        SelectedElement = -1
    End Sub

'Form Controls
    'Add Element Locations
        Sub AddLocations(obj)
        'This Sub creates a list of Element locations and adds them to a form chouce box/cell/control
            Dim lb As Object
            Dim locationArray(1 To 200) As Variant '200 Locations chouce box/cell/controls
            Dim listboxname As String
            Dim cnt As Long
            For i = 1 To 1
                If i = 1 Then Set lb = obj
                cnt = 0
                For k = 1 To 20 'Number of Rows
                    For p = 1 To 10 ' Number of columns
                        locationArray(cnt + p) = "r" & k & "-c" & p 'Naming convention: r1-c1
                    Next p
                    cnt = cnt + 10 'Advance to next row
                Next k
                For j = 1 To UBound(locationArray)
                    lb.AddItem locationArray(j) 'Add each location to the chouce box/cell/control
                Next j
            Next i
        End Sub

    'Change color of element digram cells/controls if an element is assigned
        Sub GUI()
            Dim ob As Control
            Dim Str, str2 As String
            Dim i, j, k, L As Integer
            Dim output, tempArray, wholeArray As Variant
            Dim hasElementColor
            Dim baseColor

            
            hasElementColor = &HA7B596
            baseColor = &HF5EAE5
            
            For Each ob In Main.Controls
               If Left(ob.Name, 3) = "tbR" Then ob.BackColor = baseColor
            Next ob
            
            For Each ob In Main.Controls
                If Left(ob.Name, 3) = "tbR" Then
                    str2 = ob.Name
                    str2 = Replace(str2, "tb", "")
                    str2 = LCase(str2)
                    str2 = Replace(str2, "c", "-c")
 
                    Call GUIColorControls(WDCharts, hasElementColor, str2, ob)
                    Call GUIColorControls(WDTexts, hasElementColor, str2, ob)
                    Call GUIColorControls(WDMeasures, hasElementColor, str2, ob)
                    Call GUIColorControls(WDiFrames, hasElementColor, str2, ob)
                    Call GUIColorControls(WDImages, hasElementColor, str2, ob)
                    Call GUIColorControls(WDTables, hasElementColor, str2, ob)
                End If
            Next ob
        End Sub
        
        'Suttporting Functio to GUI
            Sub GUIColorControls(ary As Variant, col, str2, ob)
                For i = 1 To UBound(ary)
                    If IsEmpty(ary(i)) Then GoTo nxt2:
                    If ary(i) = "" Then GoTo nxt2:
                    TempArray1 = Split(ary(i), SD)
                    If str2 = TempArray1(0) Then ob.BackColor = col
nxt2:
                    TempArray1 = Empty
                Next i
            End Sub

'Element ListBox Functions
    'Element Button
        Sub ElementButtonSelectd(ary As Variant)
             Dim i, j As Single

            For j = 1 To UBound(ary)
              If IsEmpty(ary(j)) Then GoTo nxt:
              If ary(j) = "" Then GoTo nxt:
              TempArray1 = Split(ary(j), SD)
                  
              Main.WDElements.AddItem
              Main.WDElements.List(ElementListBoxItemCount, 0) = j & "-Chart"
              Main.WDElements.List(ElementListBoxItemCount, 1) = "Chart"
              Main.WDElements.List(ElementListBoxItemCount, 2) = TempArray1(0)
            
              ElementListBoxItemCount = ElementListBoxItemCount + 1
nxt:
              TempArray1 = Empty
            Next j
        End Sub

   'Selected Element
        Sub GUIElementSelect(Str As String)
            'This Sub requires to have teh select Element to be passed into it.
            'It will then pass all relevant Elements into the Ellement ListBox
        
            Dim WDArray, ElementArray, SingleArray As Variant
            Dim i, j As Single
            
            Str = Replace(Str, "tb", "")
            Str = LCase(Str)
            Str = Replace(Str, "c", "-c")
            init.ElementListBoxItemCount = 0
            
            ReDim ElementArray(1 To 6)
            ElementArray(1) = WDCharts
            ElementArray(2) = WDTables
            ElementArray(3) = WDMeasures
            ElementArray(4) = WDImages
            ElementArray(5) = WDiFrames
            ElementArray(6) = WDTexts
        
            For i = 1 To UBound(ElementArray)
                SingleArray = ElementArray(i)
                For j = 1 To UBound(SingleArray)
                    If IsEmpty(SingleArray(j)) Then GoTo nxt:
                    If SingleArray(j) = "" Then GoTo nxt:
                    TempArray1 = Split(SingleArray(j), SD)
                        
                    Main.WDElements.AddItem
                    Main.WDElements.List(ElementListBoxItemCount, 0) = j & "-Chart"
                    Main.WDElements.List(ElementListBoxItemCount, 1) = "Chart"
                    Main.WDElements.List(ElementListBoxItemCount, 2) = TempArray1(0)
                  
                    init.ElementListBoxItemCount = init.ElementListBoxItemCount + 1
nxt:
                    TempArray1 = Empty
                Next j
            Next i
        End Sub
