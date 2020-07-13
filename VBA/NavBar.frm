VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NavBar 
   Caption         =   "Navigation Bar"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14175
   OleObjectBlob   =   "NavBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NavBar"
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
