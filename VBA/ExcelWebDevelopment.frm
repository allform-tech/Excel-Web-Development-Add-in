VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelWebDevelopment 
   Caption         =   "Excel Web Development Software - Allform Software Solutions"
   ClientHeight    =   13260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23175
   OleObjectBlob   =   "ExcelWebDevelopment.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcelWebDevelopment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'User Form
Public WD As WD
Public WDPageNo, WDNavBarHeadingNo, WDNavBarLinkNo, WDChartNo, WDTextNo, WDMeasureNo, WDiFrameNo, WDImageNo, WDTableNo As Integer

Private Sub UserForm_Initialize()
    'Set Variables
    Set WD = New WD
    WDPageNo = 1
    WDNavBarHeadingNo = 1
    WDNavBarLinkNo = 1
    WDChartNo = 1
    WDTextNo = 1
    WDMeasureNo = 1
    WDiFrameNo = 1
    WDImageNo = 1
    WDTableNo = 1

    'ListBox Load
    Call AddLocations
    
End Sub

'Number Only Functions






'ListBox Load Functions
Sub AddLocations()
    Dim lb As Object
    Dim locationArray(0 To 1070) As Variant
    Dim listboxname As String
    Dim cnt As Long
    
    For i = 1 To 6
        If i = 1 Then Set lb = Me.WDTableLocation
        If i = 2 Then Set lb = Me.WDMeasureLocation
        If i = 3 Then Set lb = Me.WDiFrameLocation
        If i = 4 Then Set lb = Me.WDChartLocation
        If i = 5 Then Set lb = Me.WDTextLocation
        If i = 6 Then Set lb = Me.WDImageLocation
        
        cnt = 0
        For k = 0 To 50
            For p = 0 To 20
                locationArray(cnt + p) = "r" & k & "-c" & p
            Next p
            cnt = cnt + 21
        Next k
        
        For j = 0 To UBound(locationArray)
            lb.AddItem locationArray(j)
        Next j
    Next i
End Sub



Private Sub WDiFrameInclude_Click()

End Sub
