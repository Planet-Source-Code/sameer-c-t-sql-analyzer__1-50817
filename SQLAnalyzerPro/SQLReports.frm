VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reports"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar prgReport 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4170
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame fraTables 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5130
      Begin VB.CheckBox chkPrefix 
         Caption         =   "Add"
         Height          =   195
         Left            =   1305
         TabIndex        =   4
         ToolTipText     =   "Check this to add to the present selection"
         Top             =   345
         Width           =   585
      End
      Begin VB.CommandButton cmdPrefix 
         Caption         =   "Select"
         Height          =   300
         Left            =   1905
         TabIndex        =   5
         ToolTipText     =   "Click to select the tables with typed prefix"
         Top             =   285
         Width           =   660
      End
      Begin VB.TextBox txtPrefix 
         Height          =   285
         Left            =   615
         TabIndex        =   3
         ToolTipText     =   "Type prefix letters and click Select button to select tables with that prefix"
         Top             =   300
         Width           =   600
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "A&ll"
         Height          =   300
         Left            =   2595
         TabIndex        =   6
         ToolTipText     =   "Click to select all tables"
         Top             =   285
         Width           =   660
      End
      Begin VB.CheckBox chkObjects 
         Caption         =   "&Functions"
         Height          =   270
         Index           =   3
         Left            =   3360
         TabIndex        =   13
         ToolTipText     =   "Report will include list of all user defined functions"
         Top             =   2940
         Width           =   1680
      End
      Begin VB.CheckBox chkObjects 
         Caption         =   "&Stored Procedures"
         Height          =   270
         Index           =   2
         Left            =   3360
         TabIndex        =   12
         ToolTipText     =   "Report will include list of all stored procedures"
         Top             =   2415
         Width           =   1680
      End
      Begin VB.CheckBox chkObjects 
         Caption         =   "&Views"
         Height          =   270
         Index           =   1
         Left            =   3360
         TabIndex        =   11
         ToolTipText     =   "Report will include list of all views"
         Top             =   1890
         Width           =   1680
      End
      Begin VB.CheckBox chkObjects 
         Caption         =   "T&ables"
         Height          =   270
         Index           =   0
         Left            =   3360
         TabIndex        =   10
         ToolTipText     =   "Report will include list of all user tables"
         Top             =   1365
         Width           =   1680
      End
      Begin VB.OptionButton optType 
         Caption         =   "O&bject Listing"
         Height          =   315
         Index           =   1
         Left            =   3375
         TabIndex        =   9
         ToolTipText     =   "Generates report on checked database objects"
         Top             =   885
         Width           =   1290
      End
      Begin VB.OptionButton optType 
         Caption         =   "&Table Details"
         Height          =   315
         Index           =   0
         Left            =   3375
         TabIndex        =   8
         ToolTipText     =   "Generates report on selected tables"
         Top             =   315
         Width           =   1260
      End
      Begin VB.ListBox lstTables 
         Height          =   2535
         Left            =   150
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   7
         ToolTipText     =   "List of available tables"
         Top             =   630
         Width           =   3105
      End
      Begin VB.Label lblTotal 
         Caption         =   "Total number of selected tables: 0"
         Height          =   300
         Left            =   150
         TabIndex        =   1
         Top             =   3210
         Width           =   3105
      End
      Begin VB.Label lblPrefix 
         AutoSize        =   -1  'True
         Caption         =   "Prefix"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   345
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   315
      Left            =   4275
      TabIndex        =   15
      ToolTipText     =   "Click to close this window"
      Top             =   3765
      Width           =   975
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   315
      Left            =   3165
      TabIndex        =   14
      ToolTipText     =   "Click to generate the report in Excel"
      Top             =   3765
      Width           =   975
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Reports
'Description    :   Generates Excel Reports on database objects
'Developed By   :   Sameer C T
'Started On     :   2003 April 12
'Last Modified  :   2003 May 10
'-------------------------------------------------------------------------------------------

Option Explicit
'-------------------------------------------------------------------------------------------
'Module Level Constants
'-------------------------------------------------------------------------------------------
Private Const CONST_REPORTTYPE_TABLES = 0
Private Const CONST_REPORTTYPE_OBJECTS = 1

Private Const CONST_OBJ_TABLE = 0
Private Const CONST_OBJ_VIEW = 1
Private Const CONST_OBJ_SP = 2
Private Const CONST_OBJ_FUNCTION = 3

'-------------------------------------------------------------------------------------------
'Module Level Variables
'-------------------------------------------------------------------------------------------
Private objDatabase2 As Database2 '2 is necessary to get SQL Server 2000 feature (Functions)
Private objTable As Table
Private objColumn As Column
Private objView As View
Private objSP As StoredProcedure
Private objFunction As UserDefinedFunction

Private strTableName As String
    
Private blnQuit As Boolean
Private blnToggle As Boolean
Private intCount As Integer

Private exlApln As Excel.Application
Private exlBook As Excel.Workbook
Private exlSheet As Excel.Worksheet
Private intRow As Integer, intCol As Integer
Private intJ As Integer, intK As Integer
'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Load()
    On Error GoTo ErrorTrap
    
    Call ShowStatusMsg("Loading... Please Wait...", True)
    Me.Icon = frmLogin.Icon
    Me.Height = 4545
    
    'Default values of check boxes and option buttons
    optType(CONST_REPORTTYPE_TABLES).Value = True
    chkObjects(CONST_OBJ_TABLE).Value = 1
    blnToggle = True
    blnQuit = False
    
    'Creating Database object using SQL DMO
    Set objDatabase2 = objSQLServer.Databases(mstrDatabase)
    
    'Populating all tables in the database
    For Each objTable In objDatabase2.Tables
      If objTable.SystemObject = False Then
         strTableName = objTable.Name
         lstTables.AddItem (strTableName)
         lstTables.ListIndex = lstTables.NewIndex
      End If
    Next
    
    Call ShowStatusMsg("Ready", False)
    Exit Sub
ErrorTrap:
    MsgBox "Error while loading... " & vbCrLf & _
         "Technical Details:" & vbCrLf & Err.Description, vbInformation
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set objDatabase2 = Nothing
End Sub

Private Sub cmdToggle_Click()
   'Toggles the selection of tables
   If blnToggle = True Then
      cmdToggle.Caption = "&None"
      cmdToggle.ToolTipText = "Click to deselect all tables"
   Else
      cmdToggle.Caption = "A&ll"
      cmdToggle.ToolTipText = "Click to select all tables"
   End If
   
   For intCount = 0 To lstTables.ListCount - 1
      lstTables.Selected(intCount) = blnToggle
   Next intCount
   
   blnToggle = Not blnToggle
   lblTotal.Caption = "Total number of selected tables:" & lstTables.SelCount
End Sub

Private Sub cmdPrefix_Click()
   'Select only the tables with typed prefix; If the Add Check box is ticked, it adds to the present selection
   For intCount = 0 To lstTables.ListCount - 1
      If UCase(Left(lstTables.List(intCount), Len(Trim(txtPrefix.Text)))) = UCase(Trim(txtPrefix.Text)) Then
         lstTables.Selected(intCount) = True
      Else
         If chkPrefix.Value = 0 Then lstTables.Selected(intCount) = False
      End If
   Next intCount
   lblTotal.Caption = "Total number of selected tables:" & lstTables.SelCount
End Sub

Private Sub lstTables_ItemCheck(Item As Integer)
   lblTotal.Caption = "Total number of selected tables:" & lstTables.SelCount
End Sub

Private Sub optType_Click(Index As Integer)
   Call mprEnableCheckBoxes(Index)
End Sub

Private Sub cmdGenerate_Click()
    'Generates the report
    Dim intJ As Integer, intK As Integer
    
    On Error GoTo ErrorTrap
    
    Call ShowStatusMsg("Generating Report... Please Wait...", True)
    
    Me.Height = 4935 'Showing Progress bar
    prgReport.Value = 0
    
    'If Excel is not installed; exit; else create the object
    If gfnExlInitialise(exlApln, exlBook, exlSheet) = False Then
      Call ShowStatusMsg("Ready", False)
      Exit Sub
    End If
    
    'The basic Excel settings
    'Call gprExlDisableMenus(exlApln, False, False)
    'Call gprExlDisplayToolbars(exlApln, False)
    'Call gprExlDisplaySettings(exlApln, False, True, False)
    Call gprExlSetCaptions(exlApln, "SQL Analyzer", "Report")
    
    'Report starting at...
    intRow = 3: intCol = 1
    
    'If the Table Details option is selected...
    If optType(CONST_REPORTTYPE_TABLES).Value = True Then
      'Set the report title...
      Call gprExlSetTitles(exlSheet, "Server: " & mstrServer & " - Database: " & mstrDatabase, 12, "Table Details", 10)
      '... and arrange the selected table details in excel
      Call mprListTableDetails
      
      'If user closes window, while generating report
      If blnQuit = True Then
         Call ShowStatusMsg("Ready", False)
         Exit Sub
      End If
         
    Else 'Object listing section
    
      'Set the report title...
      Call gprExlSetTitles(exlSheet, "Server: " & mstrServer & " - Database: " & mstrDatabase, 12, "List of Objects", 10)
      
      'Listing Tables
      If chkObjects(CONST_OBJ_TABLE).Value = 1 Then
         Call mprListTables
      End If
      
      'If user closes window, while generating report
      If blnQuit = True Then
         Call ShowStatusMsg("Ready", False)
         Exit Sub
      End If
      
      'Listing Views
      If chkObjects(CONST_OBJ_VIEW).Value = 1 Then
         Call mprListViews
      End If
      
      'If user closes window, while generating report
      If blnQuit = True Then
         Call ShowStatusMsg("Ready", False)
         Exit Sub
      End If
      
      'Listing SPs
      If chkObjects(CONST_OBJ_SP).Value = 1 Then
         Call mprListSPs
      End If
      
      'If user closes window, while generating report
      If blnQuit = True Then
         Call ShowStatusMsg("Ready", False)
         Exit Sub
      End If
      
      'Listing Functions
      If chkObjects(CONST_OBJ_FUNCTION).Value = 1 Then
         Call mprListFunctions
      End If
      
      'If user closes window, while generating report
      If blnQuit = True Then
         Call ShowStatusMsg("Ready", False)
         Exit Sub
      End If
    End If
    
    'Report Footer
    exlSheet.Cells(intRow + 1, intCol) = "[End of the Report Generated by SQL Analyzer On " & Format(Date, "yyyy MMM dd") & " At " & Time & "]"
    
    'Showing the report after page setup and then disposing
    Call gprExlPageSetup(exlSheet, 3)
    Call gprExlShow(exlApln, exlBook, exlSheet, False)
    Call gprExlDispose(exlApln, exlBook, exlSheet)
    
    Me.Height = 4545 'Hiding Progress bar
    Call ShowStatusMsg("Ready", False)
    Exit Sub
ErrorTrap:
    MsgBox "Error while generating report... " & vbCrLf & _
         "Technical Details:" & vbCrLf & Err.Description, vbInformation
    Call gprExlDispose(exlApln, exlBook, exlSheet)
    Me.Height = 4545 'Hiding Progress bar
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub cmdClose_Click()
   'If user closes window, while generating report; make the quit flag to true to check this state
   blnQuit = True
   Call gprExlDispose(exlApln, exlBook, exlSheet)
   Unload Me
End Sub

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------
Private Sub mprEnableCheckBoxes(CheckType As Integer)
   'Enables or disables controls
   If CheckType = CONST_REPORTTYPE_TABLES Then
      lstTables.Enabled = True
      For intCount = 0 To 3
         chkObjects(intCount).Enabled = False
      Next intCount
   ElseIf CheckType = CONST_REPORTTYPE_OBJECTS Then
      lstTables.Enabled = False
      For intCount = 0 To 3
         chkObjects(intCount).Enabled = True
      Next intCount
   End If
End Sub

Private Sub mprListTableDetails()
   'Arrange the selected table's details in Excel sheet
   prgReport.Min = 0: prgReport.Max = lstTables.SelCount + 1
   intK = 1
   For intCount = 0 To lstTables.ListCount - 1
      'Allowing focus to other controls, so that user can interrupt report generation and close
      DoEvents
      If blnQuit = True Then
        Exit Sub
      End If
      If lstTables.Selected(intCount) = True Then
         'The heading section (No. and Table Name)
         intRow = intRow + 1
         exlSheet.Cells(intRow, intCol) = intK 'Serial Number
         exlSheet.Cells(intRow, intCol + 1) = lstTables.List(intCount) 'Table Name
         exlSheet.Range(exlSheet.Cells(intRow, intCol), exlSheet.Cells(intRow, intCol + 1)).Font.Bold = True
         exlSheet.Range(exlSheet.Cells(intRow, intCol), exlSheet.Cells(intRow, intCol + 1)).Font.Size = 9
         
         Set objTable = objDatabase2.Tables(lstTables.List(intCount))
         
         'Column Headers
         intRow = intRow + 1
         intK = intK + 1
         exlSheet.Cells(intRow, intCol + 1) = "No"
         exlSheet.Cells(intRow, intCol + 1).HorizontalAlignment = xlRight
         exlSheet.Cells(intRow, intCol + 2) = "Column Name"
         exlSheet.Cells(intRow, intCol + 3) = "PK"
         exlSheet.Cells(intRow, intCol + 4) = "Data Type"
         exlSheet.Cells(intRow, intCol + 5) = "Length"
         exlSheet.Cells(intRow, intCol + 6) = "Precision"
         exlSheet.Cells(intRow, intCol + 7) = "Scale"
         exlSheet.Cells(intRow, intCol + 8) = "Nullable"
         exlSheet.Cells(intRow, intCol + 9) = "Identity"
         exlSheet.Cells(intRow, intCol + 10) = "Default"
         exlSheet.Cells(intRow, intCol + 11) = "Computed"
         exlSheet.Range(exlSheet.Cells(intRow, intCol + 1), exlSheet.Cells(intRow, intCol + 11)).Font.Bold = True
         exlSheet.Range(exlSheet.Cells(intRow, intCol + 1), exlSheet.Cells(intRow, intCol + 11)).Font.Size = 8
         
         'And the details of the table columns
         intRow = intRow + 1
         intJ = 1
         For Each objColumn In objTable.Columns
            exlSheet.Cells(intRow, intCol + 1) = intJ
            exlSheet.Cells(intRow, intCol + 2) = objColumn.Name
            exlSheet.Columns(intCol + 2).AutoFit
            exlSheet.Cells(intRow, intCol + 3) = IIf(objColumn.InPrimaryKey, " * ", "")
            exlSheet.Columns(intCol + 3).ColumnWidth = 3
            exlSheet.Cells(intRow, intCol + 4) = objColumn.Datatype
            exlSheet.Columns(intCol + 4).ColumnWidth = 10
            exlSheet.Cells(intRow, intCol + 5) = objColumn.Length
            exlSheet.Columns(intCol + 5).ColumnWidth = 7
            exlSheet.Cells(intRow, intCol + 6) = objColumn.NumericPrecision
            exlSheet.Columns(intCol + 6).ColumnWidth = 8
            exlSheet.Cells(intRow, intCol + 7) = objColumn.NumericScale
            exlSheet.Columns(intCol + 7).ColumnWidth = 7
            exlSheet.Cells(intRow, intCol + 8) = IIf(objColumn.AllowNulls, "Yes", "")
            exlSheet.Columns(intCol + 8).ColumnWidth = 7
            exlSheet.Cells(intRow, intCol + 9) = IIf(objColumn.Identity, "Yes", "")
            exlSheet.Columns(intCol + 9).ColumnWidth = 7
            exlSheet.Cells(intRow, intCol + 10) = IIf(Trim(objColumn.DRIDefault.Text) = -1, "1", objColumn.DRIDefault.Text)
            exlSheet.Columns(intCol + 10).ColumnWidth = 7
            exlSheet.Cells(intRow, intCol + 11) = IIf(objColumn.IsComputed, objColumn.ComputedText, "")
            exlSheet.Columns(intCol + 11).ColumnWidth = 10
 
            intJ = intJ + 1
            intRow = intRow + 1
         Next
         If prgReport.Value < prgReport.Max Then prgReport.Value = prgReport.Value + 1
      End If 'Selected Tables
   Next intCount 'All Tables
End Sub

Private Sub mprListTables()
   'Lists all the user tables and their details
   prgReport.Min = 0: prgReport.Max = objDatabase2.Tables.Count + 1: prgReport.Value = 0
   
   'The heading section
   intK = 1
   intRow = intRow + 1
   exlSheet.Cells(intRow, intCol) = "Tables"
   exlSheet.Cells(intRow, intCol).Font.Bold = True
   exlSheet.Cells(intRow, intCol).Font.Size = 9
   
   'Column Headers
   intRow = intRow + 1
   exlSheet.Cells(intRow, intCol) = "No"
   exlSheet.Cells(intRow, intCol).HorizontalAlignment = xlRight
   exlSheet.Cells(intRow, intCol + 1) = "Table Name"
   exlSheet.Cells(intRow, intCol + 2) = "Owner"
   exlSheet.Cells(intRow, intCol + 3) = "Created On"
   exlSheet.Cells(intRow, intCol + 4) = "Columns"
   exlSheet.Cells(intRow, intCol + 5) = "Rows"
   
   exlSheet.Range(exlSheet.Cells(intRow, intCol), exlSheet.Cells(intRow, intCol + 5)).Font.Bold = True
   exlSheet.Range(exlSheet.Cells(intRow, intCol), exlSheet.Cells(intRow, intCol + 5)).Font.Size = 8
   
   'Details of Tables
   For Each objTable In objDatabase2.Tables
     'Allowing focus to other controls, so that user can interrupt report generation and close
     DoEvents
     If blnQuit = True Then
        Exit Sub
     End If
     
     If objTable.SystemObject = False Then
        intRow = intRow + 1
        exlSheet.Cells(intRow, intCol) = intK 'Serial Number
        exlSheet.Cells(intRow, intCol + 1) = objTable.Name
        exlSheet.Cells(intRow, intCol + 2) = objTable.Owner
        exlSheet.Cells(intRow, intCol + 3) = objTable.CreateDate
        exlSheet.Cells(intRow, intCol + 3).NumberFormat = "dd/mm/yyyy hh:mm"
        exlSheet.Cells(intRow, intCol + 4) = objTable.Columns.Count
        exlSheet.Cells(intRow, intCol + 5) = objTable.Rows
        intK = intK + 1
     End If
     If prgReport.Value < prgReport.Max Then prgReport.Value = prgReport.Value + 1
   Next
   
   exlSheet.Columns(intCol + 1).AutoFit
   exlSheet.Columns(intCol + 2).AutoFit
   exlSheet.Columns(intCol + 3).AutoFit
   exlSheet.Columns(intCol + 4).AutoFit
   exlSheet.Columns(intCol + 5).AutoFit
End Sub

Private Sub mprListViews()
   'Lists all the user views and their details
   prgReport.Min = 0: prgReport.Max = objDatabase2.Views.Count + 1: prgReport.Value = 0
   
   'The heading section
   intK = 1
   intRow = intRow + 1
   exlSheet.Cells(intRow, intCol) = "Views"
   exlSheet.Cells(intRow, intCol).Font.Bold = True
   exlSheet.Cells(intRow, intCol).Font.Size = 9
   
   'Column Headers
   intRow = intRow + 1
   exlSheet.Cells(intRow, intCol) = "No"
   exlSheet.Cells(intRow, intCol).HorizontalAlignment = xlRight
   exlSheet.Cells(intRow, intCol + 1) = "View Name"
   exlSheet.Cells(intRow, intCol + 2) = "Owner"
   exlSheet.Cells(intRow, intCol + 3) = "Created On"
   exlSheet.Cells(intRow, intCol + 4) = "Columns"
   
   exlSheet.Range(exlSheet.Cells(intRow, intCol), exlSheet.Cells(intRow, intCol + 5)).Font.Bold = True
   exlSheet.Range(exlSheet.Cells(intRow, intCol), exlSheet.Cells(intRow, intCol + 5)).Font.Size = 8
   
   'Details of Views
   For Each objView In objDatabase2.Views
     'Allowing focus to other controls, so that user can interrupt report generation and close
     DoEvents
     If blnQuit = True Then
        Exit Sub
     End If
     
     If objView.SystemObject = False Then
        intRow = intRow + 1
        exlSheet.Cells(intRow, intCol) = intK 'Serial Number
        exlSheet.Cells(intRow, intCol + 1) = objView.Name
        exlSheet.Cells(intRow, intCol + 2) = objView.Owner
        exlSheet.Cells(intRow, intCol + 3) = objView.CreateDate
        exlSheet.Cells(intRow, intCol + 3).NumberFormat = "dd/mm/yyyy hh:mm"
        exlSheet.Cells(intRow, intCol + 4) = objView.ListColumns.Count
        intK = intK + 1
     End If
     If prgReport.Value < prgReport.Max Then prgReport.Value = prgReport.Value + 1
   Next
   
   exlSheet.Columns(intCol + 1).AutoFit
   exlSheet.Columns(intCol + 2).AutoFit
   exlSheet.Columns(intCol + 3).AutoFit
   exlSheet.Columns(intCol + 4).AutoFit
End Sub

Private Sub mprListSPs()
   'Lists all the user stored procedures and their details
   prgReport.Min = 0: prgReport.Max = objDatabase2.StoredProcedures.Count + 1: prgReport.Value = 0
   
   'The heading section
   intK = 1
   intRow = intRow + 1
   exlSheet.Cells(intRow, intCol) = "Stored Procedures"
   exlSheet.Cells(intRow, intCol).Font.Bold = True
   exlSheet.Cells(intRow, intCol).Font.Size = 9
   
   'Column Headers
   intRow = intRow + 1
   exlSheet.Cells(intRow, intCol) = "No"
   exlSheet.Cells(intRow, intCol).HorizontalAlignment = xlRight
   exlSheet.Cells(intRow, intCol + 1) = "Stored Procedure Name"
   exlSheet.Cells(intRow, intCol + 2) = "Owner"
   exlSheet.Cells(intRow, intCol + 3) = "Created On"
   
   exlSheet.Range(exlSheet.Cells(intRow, intCol), exlSheet.Cells(intRow, intCol + 5)).Font.Bold = True
   exlSheet.Range(exlSheet.Cells(intRow, intCol), exlSheet.Cells(intRow, intCol + 5)).Font.Size = 8
   
   'Details of Stored Procedures
   For Each objSP In objDatabase2.StoredProcedures
     'Allowing focus to other controls, so that user can interrupt report generation and close
     DoEvents
     If blnQuit = True Then
        Exit Sub
     End If
     
     If objSP.SystemObject = False Then
        intRow = intRow + 1
        exlSheet.Cells(intRow, intCol) = intK 'Serial Number
        exlSheet.Cells(intRow, intCol + 1) = objSP.Name
        exlSheet.Cells(intRow, intCol + 2) = objSP.Owner
        exlSheet.Cells(intRow, intCol + 3) = objSP.CreateDate
        exlSheet.Cells(intRow, intCol + 3).NumberFormat = "dd/mm/yyyy hh:mm"
        intK = intK + 1
     End If
     If prgReport.Value < prgReport.Max Then prgReport.Value = prgReport.Value + 1
   Next
   
   exlSheet.Columns(intCol + 1).AutoFit
   exlSheet.Columns(intCol + 2).AutoFit
   exlSheet.Columns(intCol + 3).AutoFit
End Sub

Private Sub mprListFunctions()
   'Lists all the User Defined Functions and their details
   prgReport.Min = 0: prgReport.Max = objDatabase2.UserDefinedFunctions.Count + 1: prgReport.Value = 0
   
   'The heading section
   intK = 1
   intRow = intRow + 1
   exlSheet.Cells(intRow, intCol) = "User Defined Functions"
   exlSheet.Cells(intRow, intCol).Font.Bold = True
   exlSheet.Cells(intRow, intCol).Font.Size = 9
   
   'Column Headers
   intRow = intRow + 1
   exlSheet.Cells(intRow, intCol) = "No"
   exlSheet.Cells(intRow, intCol).HorizontalAlignment = xlRight
   exlSheet.Cells(intRow, intCol + 1) = "Function Name"
   exlSheet.Cells(intRow, intCol + 2) = "Owner"
   exlSheet.Cells(intRow, intCol + 3) = "Created On"

   exlSheet.Range(exlSheet.Cells(intRow, intCol), exlSheet.Cells(intRow, intCol + 5)).Font.Bold = True
   exlSheet.Range(exlSheet.Cells(intRow, intCol), exlSheet.Cells(intRow, intCol + 5)).Font.Size = 8
   
   'Details of User Defined Functions
   For Each objFunction In objDatabase2.UserDefinedFunctions
     'Allowing focus to other controls, so that user can interrupt report generation and close
     DoEvents
     If blnQuit = True Then
        Exit Sub
     End If
     
     If objFunction.SystemObject = False Then
        intRow = intRow + 1
        exlSheet.Cells(intRow, intCol) = intK 'Serial Number
        exlSheet.Cells(intRow, intCol + 1) = objFunction.Name
        exlSheet.Cells(intRow, intCol + 2) = objFunction.Owner
        exlSheet.Cells(intRow, intCol + 3) = objFunction.CreateDate
        exlSheet.Cells(intRow, intCol + 3).NumberFormat = "dd/mm/yyyy hh:mm"
        intK = intK + 1
     End If
     If prgReport.Value < prgReport.Max Then prgReport.Value = prgReport.Value + 1
   Next
   
   exlSheet.Columns(intCol + 1).AutoFit
   exlSheet.Columns(intCol + 2).AutoFit
   exlSheet.Columns(intCol + 3).AutoFit
End Sub
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------


