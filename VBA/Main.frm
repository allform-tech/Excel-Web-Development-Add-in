VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Main 
   Caption         =   "Excel Web Development Software - Allform Software Solutions"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15930
   OleObjectBlob   =   "Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public formWidth As Boolean
Public WD As WD

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

Private Sub UserForm_Initialize()
Set WD = New WD
formWidth = True
Me.Width = 122.25
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
