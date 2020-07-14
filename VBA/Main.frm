VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Main 
   Caption         =   "Excel Web Development Software - Allform Software Solutions"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15930
   OleObjectBlob   =   "Main.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public formWidth As Boolean
Public WD As WD



'Init
Private Sub UserForm_Initialize()

    'Load Class Module (WD)
    Call init.InitiateVariables
    
    'Load Elements to Lost Box (Me.WDElements)
    Me.Width = 122.25
    Me.WDElements.ColumnCount = 3
    Me.WDElements.ColumnWidths = "50, 50, 50"
    'Me.WDElements.AddItem "asd"

    
'    Call init.loadChartsListBox
'    Call init.loadTablesListBox
'    Call init.loadMeasuresListBox
'    Call init.loadiFramesListBox
'    Call init.loadImagesListBox
'    Call init.loadiTextsListBox
    
    
    
End Sub

'Load Element Forms
Private Sub btnAddChart_Click()
    Me.Hide
    Charts.Show
End Sub

Private Sub btnAddiFrame_Click()
    Me.Hide
    iframes.Show
End Sub

Private Sub btnAddImage_Click()
    Me.Hide
    Images.Show
End Sub

Private Sub btnAddMeasure_Click()
    Me.Hide
    Tables.Show
End Sub

Private Sub btnAddTable_Click()
    Me.Hide
    Measures.Show
End Sub

Private Sub btnClose_Click()
Me.Hide
End Sub

Private Sub btnCSS_Click()
    Me.Hide
    CSS.Show
End Sub

Private Sub btnJavaScript_Click()
    Me.Hide
    JavaScript.Show
End Sub

Private Sub btnNavBar_Click()
    Me.Hide
    NavBar.Show
End Sub

Private Sub btnPageInfo_Click()
    Me.Hide
    Page.Show
End Sub


'Elements Change View
Private Sub btnViewAll_Click()
    Call init.loadChartsListBox
    Call init.loadTablesListBox
    Call init.loadMeasuresListBox
    Call init.loadiFramesListBox
    Call init.loadImagesListBox
    Call init.loadTextsListBox
End Sub

Private Sub btnViewCharts_Click()
    Me.WDElements.Value = ""
    Call init.loadChartsListBox
End Sub

Private Sub btnViewiFrames_Click()
    Me.WDElements.Value = ""
    Call init.loadiFramesListBox
End Sub

Private Sub btnViewImages_Click()
    Me.WDElements.Value = ""
    Call init.loadImagesListBox
End Sub

Private Sub btnViewTables_Click()
    Me.WDElements.Value = ""
    Call init.loadTablesListBox
End Sub

Private Sub viewMeasures_Click()
    Me.WDElements.Value = ""
    Call init.loadMeasuresListBox
End Sub

Private Sub btnViewText_Click()
    Call init.loadTextsListBox
End Sub


Private Sub btnExpand_Click()
    If Me.btnExpand.Caption = "Expand" Then
        Me.btnExpand.Caption = "Minimise"
        Me.Width = 808.5
    Else
        Me.btnExpand.Caption = "Expand"
        Me.Width = 122.25
    End If
End Sub

Private Sub WDElements_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim str As String
    Dim tempArray As Variant
    str = WDElements.Value
    Debug.Print (init.SelectedElement)

    tempArray = Split(str, "-")
    init.SelectedElement = tempArray(0)
    
    Select Case True
        Case tempArray(1) = "Chart"
            Me.Hide
            Charts.Show
        Case tempArray(1) = "Table"
        
        Case tempArray(1) = "Text"
        
        Case tempArray(1) = "Measure"
        
        Case tempArray(1) = "Image"
        
        Case tempArray(1) = "iFrame"
        
    End Select
    
    
End Sub
