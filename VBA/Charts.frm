VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Charts 
   Caption         =   "Charts.JS"
   ClientHeight    =   2760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8040
   OleObjectBlob   =   "Charts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Charts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
    Dim ArrayLocation As Integer
    Dim i As Integer
    Dim txt As String

    ArrayLocation = ChartFunctions.nextFreeCellCharts()

    txt = ""
    txt = txt & Me.WDChartLocation & SD
    txt = txt & Me.WDChartType & SD
    txt = txt & Me.WDChartTitle & SD
    txt = txt & Me.WDChartTable & SD
    txt = txt & Me.WDChartXaxisLable & SD
    txt = txt & Me.WDChartYaxixLabel & SD
    txt = txt & Me.WDChartColors & SD
    txt = txt & Me.WDChartStyles & SD
    txt = txt & Me.WDChartClass & SD
    txt = txt & Me.WDChartWidth & SD
    txt = txt & Me.WDChartLocation & SD
    
    init.WDCharts(ArrayLocation) = txt
    
    Me.Hide
    Main.Show
    init.LoadedChart = 0
    init.LoadedChart = ArrayLocation
    Main.WDElements.Clear
    Call init.loadChartsListBox
    ElementListBoxItemCount = 0
End Sub



Private Sub UserForm_Activate()
    If init.SelectedElement > -1 Then
    tempArray = Split(WDCharts(init.SelectedElement), SD)
    
        For i = 0 To UBound(tempArray)
            Me.WDChartLocation = tempArray(0)
            Me.WDChartType = tempArray(1)
            Me.WDChartTitle = tempArray(2)
            Me.WDChartTable = tempArray(3)
            Me.WDChartXaxisLable = tempArray(4)
            Me.WDChartYaxixLabel = tempArray(5)
            Me.WDChartColors = tempArray(6)
            Me.WDChartStyles = tempArray(7)
            Me.WDChartClass = tempArray(8)
            Me.WDChartWidth = tempArray(9)
            Me.WDChartLocation = tempArray(10)
        Next i
        
    End If



End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.Hide
    Main.Show
    LoadedChart = 0
    LoadedChart = -1
End Sub


Private Sub UserForm_Initialize()
    Call AddLocations
    
    Dim tempArray As Variant
    Dim i As Integer
    
    If init.SelectedElement > -1 Then
    tempArray = Split(WDCharts(init.SelectedElement), SD)
    
        For i = 0 To UBound(tempArray)
            Me.WDChartLocation = tempArray(0)
            Me.WDChartType = tempArray(1)
            Me.WDChartTitle = tempArray(2)
            Me.WDChartTable = tempArray(3)
            Me.WDChartXaxisLable = tempArray(4)
            Me.WDChartYaxixLabel = tempArray(5)
            Me.WDChartColors = tempArray(6)
            Me.WDChartStyles = tempArray(7)
            Me.WDChartClass = tempArray(8)
            Me.WDChartWidth = tempArray(9)
            Me.WDChartLocation = tempArray(10)
        Next i
        
    End If
    
End Sub



Sub AddLocations()
    Dim lb As Object
    Dim locationArray(1 To 200) As Variant
    Dim listboxname As String
    Dim cnt As Long
    
    For i = 1 To 1
        If i = 1 Then Set lb = Me.WDChartLocation
        
        cnt = 0
        For k = 1 To 20
            For p = 1 To 10
                locationArray(cnt + p) = "r" & k & "-c" & p
            Next p
            cnt = cnt + 10
        Next k
        
        For j = 1 To UBound(locationArray)
            lb.AddItem locationArray(j)
        Next j
    Next i
End Sub


