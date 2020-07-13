VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} iframes 
   Caption         =   "iFrame"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   OleObjectBlob   =   "iframes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "iframes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
Me.Hide
Main.Show
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Me.Hide
Main.Show
End Sub

Private Sub UserForm_Initialize()
Call AddLocations
End Sub
Sub AddLocations()
    Dim lb As Object
    Dim locationArray(1 To 200) As Variant
    Dim listboxname As String
    Dim cnt As Long
    
    For i = 1 To 1
        If i = 1 Then Set lb = Me.WDiFrameLocation
        
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
    Next
End Sub

