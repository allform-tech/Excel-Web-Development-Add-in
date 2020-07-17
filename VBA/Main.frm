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

'Initilise/Activate
    'Initilise
        Private Sub UserForm_Initialize()
            Call init.InitiateVariables
            Me.Width = 122.25
            Me.WDElements.ColumnCount = 3
            Me.WDElements.ColumnWidths = "50, 50, 50"
            Call init.GUI
        End Sub
    'Activate
        Private Sub UserForm_Activate()
            ElementListBoxItemCount = 0
            Call init.ElementButtonSelectd(init.WDCharts)
            Call init.ElementButtonSelectd(init.WDTables)
            Call init.ElementButtonSelectd(init.WDMeasures)
            Call init.ElementButtonSelectd(init.WDImages)
            Call init.ElementButtonSelectd(init.WDiFrames)
            Call init.ElementButtonSelectd(init.WDTexts)
            ElementListBoxItemCount = 0
            Call init.GUI
        End Sub


'Buttons
    'Add New Element
        'Add Chart
            Private Sub btnAddChart_Click()
                init.SelectedElement = -1
                Me.Hide
                init.newElement = True
                Charts.Show
            End Sub
        
        'Add iFrame
            Private Sub btnAddiFrame_Click()
                init.SelectedElement = -1
                Me.Hide
                init.newElement = True
                iframes.Show
            End Sub
         
        'Add Image
            Private Sub btnAddImage_Click()
                init.SelectedElement = -1
                Me.Hide
                init.newElement = True
                Images.Show
            End Sub
        
        'Add Measure
            Private Sub btnAddMeasure_Click()
                init.SelectedElement = -1
                Me.Hide
                init.newElement = True
                Tables.Show
            End Sub
        
        'Add Table
            Private Sub btnAddTable_Click()
                init.SelectedElement = -1
                Me.Hide
                init.newElement = True
                Measures.Show
            End Sub

    'Static Elements

        'CSS
            Private Sub btnCSS_Click()
                Me.Hide
                CSS.Show
            End Sub
        
        'JavaScript
            Private Sub btnJavaScript_Click()
                Me.Hide
                JavaScript.Show
            End Sub
            
        'Navigation Bar
            Private Sub btnNavBar_Click()
                Me.Hide
                NavBar.Show
            End Sub
            
        'Page Information
            Private Sub btnPageInfo_Click()
                Me.Hide
                Page.Show
            End Sub
            
    'Menu Buttons
        'Close
            Private Sub btnClose_Click()
                Me.Hide
            End Sub
            
        'Expand/Minimise Form
            Private Sub btnExpand_Click()
                If Me.btnExpand.Caption = "Expand" Then
                    Me.btnExpand.Caption = "Minimise"
                    Me.Width = 808.5
                Else
                    Me.btnExpand.Caption = "Expand"
                    Me.Width = 122.25
                End If
            End Sub
        
'Element List Box Controls
    'View all elements
        Private Sub btnViewAll_Click()
            Main.WDElements.Clear
            Call init.ElementButtonSelectd(init.WDCharts)
            Call init.ElementButtonSelectd(init.WDTables)
            Call init.ElementButtonSelectd(init.WDMeasures)
            Call init.ElementButtonSelectd(init.WDImages)
            Call init.ElementButtonSelectd(init.WDiFrames)
            Call init.ElementButtonSelectd(init.WDTexts)
            ElementListBoxItemCount = 0
        End Sub
    
    'Chatrs
        Private Sub btnViewCharts_Click()
            Main.WDElements.Clear
            init.ElementButtonSelectd (init.WDCharts)
            ElementListBoxItemCount = 0
        End Sub
    
    'iFrames
        Private Sub btnViewiFrames_Click()
            Main.WDElements.Clear
            Call init.ElementButtonSelectd(init.WDiFrames)
            ElementListBoxItemCount = 0
        End Sub
    
    'Images
        Private Sub btnViewImages_Click()
            Main.WDElements.Clear
            Call init.ElementButtonSelectd(init.WDImages)
            ElementListBoxItemCount = 0
        End Sub
    
    'Tables
        Private Sub btnViewTables_Click()
            Main.WDElements.Clear
            Call init.ElementButtonSelectd(init.WDTables)
            ElementListBoxItemCount = 0
        End Sub
    
    'Measures
        Private Sub viewMeasures_Click()
            Main.WDElements.Clear
            Call init.ElementButtonSelectd(init.WDMeasures)
            ElementListBoxItemCount = 0
        End Sub
    
    'Text
        Private Sub btnViewText_Click()
            Main.WDElements.Clear
            Call init.ElementButtonSelectd(init.WDTexts)
            ElementListBoxItemCount = 0
        End Sub
        
    'On Double Click an existing element
        Private Sub WDElements_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            If IsNull(WDElements) = True Then Exit Sub
            Dim Str As String
            Dim tempArray As Variant
            Str = WDElements.Value
            'Debug.Print (init.SelectedElement)
        
            tempArray = Split(Str, "-")
            init.SelectedElement = tempArray(0)
            init.newElement = False
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

    'GUI Controls
        
