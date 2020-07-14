VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CSS 
   Caption         =   "Style Script"
   ClientHeight    =   11520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   OleObjectBlob   =   "CSS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Me.Hide
Main.Show
End Sub

