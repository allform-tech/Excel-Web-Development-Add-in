VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Charts 
   Caption         =   "Charts.JS"
   ClientHeight    =   2760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8040
   OleObjectBlob   =   "Charts.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Charts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varPlaceHolder
Public txt As String

'Initilise, Activete and Close
        'Initilise
        Private Sub UserForm_Initialize()
            Dim tempArray As Variant
            Dim tempTableArray As Variant
            Dim i As Integer
            
            Call newElement 'Deturmine if a new or existing element is being created/ammended
            Call init.AddLocations(Me.WDChartLocation) 'Add element locations to the location choice/dropdown cell/control
            Call addChartTyprs 'Add chart types
            Call cellValues 'Add cell values based on the newElement Sub
        
            tempTableArray = getTableData.listTables() 'Get a list of all table names
            
            On Error Resume Next
            For i = 1 To UBound(tempTableArray) 'Add table name results to the Table hoice/dropdown cell/control
                Me.WDChartTable.AddItem tempTableArray(i)
            Next i
        End Sub
    'Activate
        Private Sub UserForm_Activate()
            Call newElement 'Deturmine if a new or existing element is being created/ammended
            Call cellValues 'Add cell values based on the newElement Sub
        End Sub
    'Close
        Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
            Me.Hide
            Main.Show
            LoadedChart = -1 'Reset Array(x) back to 0 Array is set to (1 to 100). a negitive -1 result is used to deturmine if a user selected this form via the
        End Sub
'Button Actions
    'Update/Add
        Private Sub btnAdd_Click()
            Dim ArrayLocation As Integer
            Dim i As Integer
            txt = ""
            
            'Are all controls filled in?
                If Me.WDChartLocation.Value = "" Then txt = txt & "- Location" & vbNewLine  'Location
                If Me.WDChartType.Value = "" Then txt = txt & "- Type" & vbNewLine          'Type
                If Me.WDChartTable.Value = "" Then txt = txt & "- Table" & vbNewLine        'Table
                If Me.WDChartColors.Value = "" Then txt = txt & "- Colors" & vbNewLine      'Colors
                If Me.WDChartHeight.Value = "" Then txt = txt & "- Height" & vbNewLine      'Height
                If Me.WDChartWidth.Value = "" Then txt = txt & "- Width" & vbNewLine        'Width
                If txt <> "" Then 'If required Cell are not filled out then display a message and exit sub
                    MsgBox prompt:="The Following require input:" & vbNewLine & txt, Buttons:=vbCritical, title:="Error!"
                    Exit Sub
                End If
                
            'Update Current Element Object
                If Me.btnAdd.Caption = "Update" Then
                    Dim tempSelectedElement As Integer
                    tempSelectedElement = init.SelectedElement
                    Call txtToDelimitedString(tempSelectedElement) 'Add each form control to a string
                    GoTo en: 'Finilise Sub
                End If
            
            'Add New Element Object
                ArrayLocation = ArrayFunctionsAndSubs.nextAvaliableArrayLocation(init.WDCharts) 'Find the next available/free/empty array location
                Call txtToDelimitedString(ArrayLocation) 'Add each form control to a string
en:
            'Finilise Sub
                Me.Hide 'Hide Current Form
                Main.Show 'Display Main Form
                init.LoadedChart = 0 'Reset Array(x) back to 0 Array is set to (1 to 100).
                init.LoadedChart = ArrayLocation 'Set Array(x) to current element. This is done in order to update the Elements List Box on the Main Form.
                Main.WDElements.Clear 'This Clears the Elements List Box on the Main Form.
        End Sub
        
        Sub txtToDelimitedString(ArrayLocation As Integer)
            'init.SD is the second level delimiter used when storing page data in a .xlwd file type
            txt = txt & Me.WDChartLocation & init.SD         'Location
            txt = txt & Me.WDChartType & init.SD             'Type
            txt = txt & Me.WDChartTitle & init.SD            'Title
            txt = txt & Me.WDChartTable & init.SD            'Table
            txt = txt & Me.WDChartXaxisLable & init.SD       'X Label
            txt = txt & Me.WDChartYaxixLabel & init.SD       'Y Label
            txt = txt & Me.WDChartColors & init.SD           'Colors
            txt = txt & Me.WDChartStyles & init.SD           'Styles
            txt = txt & Me.WDChartClass & init.SD            'Class
            txt = txt & Me.WDChartHeight & init.SD           'Height
            txt = txt & Me.WDChartWidth & init.SD            'Width
            txt = txt & Me.WDChartLocation & init.SD         'Location
            init.WDCharts(ArrayLocation) = txt 'Add txt to public array '################
        End Sub
    'Delete Element
        Private Sub btnDelete_Click()
            init.WDCharts(init.SelectedElement) = "" 'Reset the selectedElement variable ready for the next user command
            Me.Hide 'Hide current form
            Main.Show 'Show Main Form
            init.LoadedChart = ArrayLocation 'Set loaded element to the deleted element array(x)
            Main.WDElements.Clear 'Clear the Main Forms Element ListBox
        End Sub
'Control Validation
    'Main Subs
        'Save Variable
            Sub variableHold(var)
                varPlaceHolder = var 'Set placeholder variable in order to revert teh cell/control value if the value does not meet the validation requirements
            End Sub
        'IsNumeric
            Sub isaNumber(obj As Control, capt As String, var)
                If isNumeric(obj.Value) = False And obj.Value <> "" Then 'Is teh cell/control value a number?
                    isNoMessage (capt)  'Display a message
                    obj.Value = varPlaceHolder  'Revert the value
                End If
            End Sub
        'Message Box
            Sub isNoMessage(content As String)
                MsgBox prompt:=content & " must be a number", Buttons:=vbCritical, title:="Incorrect input type!" 'Display a must be numeric alert
            End Sub
    'Controls
        'Height
            Private Sub WDChartHeight_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
                Call variableHold(Me.WDChartHeight.Value) 'Add cell/control value to a place holder variable
            End Sub
            
            Private Sub WDChartHeight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
                Call isaNumber(Me.WDChartHeight, "Element 'height'", Me.WDChartHeight.Value) 'Deturmin if the entered cell/control value meets the validation requirements
            End Sub
        'Width
            Private Sub WDChartWidth_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
                Call variableHold(Me.WDChartWidth.Value) 'Add cell/control value to a place holder variable
            End Sub
            
            Private Sub WDChartWidth_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
                Call isaNumber(Me.WDChartWidth, "Element 'width'", Me.WDChartWidth.Value) 'Deturmin if the entered cell/control value meets the validation requirements
            End Sub
'Supporting Functions and Sub
    'Load cell/control values to form
        Sub cellValues()
            If init.SelectedElement >= 0 Then 'If the user is ammending an existing element then
                tempArray = Split(WDCharts(init.SelectedElement), SD) 'Get the array(x)/element and Split by a delimiter into an array
            
                For i = 0 To UBound(tempArray)
                    Me.WDChartLocation = tempArray(0)   'Load Location
                    Me.WDChartType = tempArray(1)       'Load Type
                    Me.WDChartTitle = tempArray(2)      'Load Title
                    Me.WDChartTable = tempArray(3)      'Table
                    Me.WDChartXaxisLable = tempArray(4) 'Load X Label
                    Me.WDChartYaxixLabel = tempArray(5) 'Load Y Label
                    Me.WDChartColors = tempArray(6)     'Load Colors
                    Me.WDChartStyles = tempArray(7)     'Load Styles
                    Me.WDChartClass = tempArray(8)      'Load Class
                    Me.WDChartHeight = tempArray(9)     'Load Height
                    Me.WDChartWidth = tempArray(10)     'Load Width
                Next i
            Else 'If the user has created a new element then
                Me.WDChartLocation = ""     'Set Location value to null
                Me.WDChartType = ""         'Set Type value to null
                Me.WDChartTitle = ""        'Set Title value to null
                Me.WDChartTable = ""        'Set Table value to null
                Me.WDChartXaxisLable = ""   'Set X Lavbel value to null
                Me.WDChartYaxixLabel = ""   'Set Y Label value to null
                Me.WDChartColors = ""       'Set Colors value to null
                Me.WDChartStyles = ""       'Set Styles value to null
                Me.WDChartClass = ""        'Set Class value to null
                Me.WDChartHeight = ""       'Set Height value to null
                Me.WDChartWidth = ""        'Set Width value to null
            End If
        End Sub
    'Add chart Types
        Sub addChartTyprs()
            With Me.WDChartType
                .AddItem "pie"                      'Pie
                .AddItem "bar"                      'Bar
                .AddItem "stackedBar"               'Stacked Bar
                .AddItem "horizontalBar"            'Horizontal Bar
                .AddItem "horizontalStackedBar"     'Horizontal Stacked Bar
                .AddItem "line"                     'Line
                .AddItem "bubble"                   'Bubble
            End With
        End Sub
    'Deturmine if a new or existing element is being created/ammended
        Sub newElement()
            If init.newElement = False Then     'If the user selects an existing Element then
                Me.btnAdd.Caption = "Update"    'Change the Add Button to display: Update
                Me.btnDelete.Enabled = True     'Disable the Delete Button
            Else
                Me.btnDelete.Enabled = False    'Enable the Delete Button
                Me.btnAdd.Caption = "Add"
            End If
        End Sub

