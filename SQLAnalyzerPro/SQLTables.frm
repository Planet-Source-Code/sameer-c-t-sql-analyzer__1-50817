VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmTables 
   Caption         =   "Tables"
   ClientHeight    =   7695
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11400
   ClipControls    =   0   'False
   Icon            =   "SQLTables.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtbTABDummy 
      Height          =   315
      Left            =   7320
      TabIndex        =   30
      Top             =   6450
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"SQLTables.frx":1D12
   End
   Begin VB.ListBox lstRight 
      Height          =   1230
      Index           =   3
      Left            =   8610
      TabIndex        =   27
      Top             =   4869
      Width           =   3180
   End
   Begin VB.ListBox lstRight 
      Height          =   1230
      Index           =   2
      Left            =   8610
      TabIndex        =   25
      Top             =   3351
      Width           =   3180
   End
   Begin VB.TextBox txtTABColName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4245
      TabIndex        =   11
      Top             =   2205
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.Frame fraTABStatements 
      Caption         =   "State&ments"
      Height          =   900
      Left            =   2730
      TabIndex        =   12
      Top             =   6735
      Width           =   5775
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate"
         Height          =   315
         Left            =   4575
         TabIndex        =   19
         ToolTipText     =   "Generates the selected statement"
         Top             =   345
         Width           =   975
      End
      Begin VB.OptionButton optStmt 
         Caption         =   "Data in the table"
         Height          =   255
         Index           =   5
         Left            =   2655
         TabIndex        =   18
         Top             =   570
         Width           =   1470
      End
      Begin VB.OptionButton optStmt 
         Caption         =   "CREATE"
         Height          =   255
         Index           =   4
         Left            =   2655
         TabIndex        =   17
         Top             =   285
         Width           =   1350
      End
      Begin VB.OptionButton optStmt 
         Caption         =   "DELETE"
         Height          =   255
         Index           =   3
         Left            =   1365
         TabIndex        =   16
         Top             =   570
         Width           =   930
      End
      Begin VB.OptionButton optStmt 
         Caption         =   "UPDATE"
         Height          =   255
         Index           =   2
         Left            =   1365
         TabIndex        =   15
         Top             =   285
         Width           =   960
      End
      Begin VB.OptionButton optStmt 
         Caption         =   "INSERT"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   14
         ToolTipText     =   "Select any one and click Generate button"
         Top             =   570
         Width           =   900
      End
      Begin VB.OptionButton optStmt 
         Caption         =   "SELECT"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   285
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.ListBox lstRight 
      Height          =   1230
      Index           =   4
      Left            =   8610
      TabIndex        =   29
      Top             =   6390
      Width           =   3180
   End
   Begin VB.ListBox lstTABTables 
      Height          =   6690
      Left            =   105
      TabIndex        =   3
      ToolTipText     =   "List of present Tables"
      Top             =   945
      Width           =   2535
   End
   Begin VB.OptionButton optTABMode 
      Caption         =   "&All Rows"
      Height          =   210
      Index           =   0
      Left            =   2730
      TabIndex        =   4
      ToolTipText     =   "Select this to show all data in the table"
      Top             =   90
      Width           =   975
   End
   Begin VB.OptionButton optTABMode 
      Caption         =   "T&op 10"
      Height          =   210
      Index           =   1
      Left            =   3780
      TabIndex        =   5
      ToolTipText     =   "Select this to show only top 10 rows in the table"
      Top             =   90
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.ListBox lstRight 
      Height          =   1230
      Index           =   0
      Left            =   8610
      TabIndex        =   21
      Top             =   315
      Width           =   3180
   End
   Begin VB.ListBox lstRight 
      Height          =   1230
      Index           =   1
      Left            =   8610
      TabIndex        =   23
      Top             =   1833
      Width           =   3180
   End
   Begin VB.TextBox txtTABSearch 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Type here to find tables"
      Top             =   315
      Width           =   2520
   End
   Begin MSComctlLib.ListView lsvTABRows 
      Height          =   1830
      Left            =   2730
      TabIndex        =   8
      ToolTipText     =   "Data in the selected Table"
      Top             =   315
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   3228
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxTABColumns 
      Height          =   3825
      Left            =   2730
      TabIndex        =   10
      ToolTipText     =   "Column details of the table"
      Top             =   2565
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6747
      _Version        =   393216
   End
   Begin VB.Label lblTABTriggers 
      Caption         =   "Tr&iggers:"
      Height          =   195
      Left            =   8610
      TabIndex        =   26
      Top             =   4635
      Width           =   2595
   End
   Begin VB.Label lblTABConstraints 
      Caption         =   "C&heck Constraints:"
      Height          =   195
      Left            =   8610
      TabIndex        =   24
      Top             =   3135
      Width           =   2595
   End
   Begin VB.Label lblTABDepend 
      Caption         =   "Dependent Objects:"
      Height          =   195
      Left            =   8625
      TabIndex        =   28
      Top             =   6165
      Width           =   2700
   End
   Begin VB.Label lblTABSearch 
      AutoSize        =   -1  'True
      Caption         =   "&Find:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   345
   End
   Begin VB.Label lblTABTotalTables 
      Caption         =   "Total &Tables:"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   735
      Width           =   2220
   End
   Begin VB.Label lblTABRows 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Showed &Rows:"
      Height          =   195
      Left            =   6225
      TabIndex        =   7
      Top             =   105
      Width           =   2220
   End
   Begin VB.Label lblTABColumns 
      Caption         =   "Total &Columns:"
      Height          =   195
      Left            =   2730
      TabIndex        =   9
      Top             =   2355
      Width           =   1515
   End
   Begin VB.Label lblTABReferring 
      Caption         =   "Referrin&g Tables:"
      Height          =   195
      Left            =   8610
      TabIndex        =   20
      Top             =   105
      Width           =   2670
   End
   Begin VB.Label lblTABReferred 
      Caption         =   "Re&ferred Tables:"
      Height          =   195
      Left            =   8610
      TabIndex        =   22
      Top             =   1590
      Width           =   2595
   End
   Begin VB.Label lblTABTableName 
      Caption         =   "Table Name:"
      Height          =   195
      Left            =   2730
      TabIndex        =   6
      Top             =   6480
      Width           =   5745
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Tables
'Description    :   To search the Tables and to view its details and data
'Developed By   :   Sameer C T
'Started On     :   2001 November 27
'Last Modified  :   2003 September 25
'-------------------------------------------------------------------------------------------

Option Explicit

'Constants for the Statements to be generates
Private Const CONST_SELECT = 0
Private Const CONST_INSERT = 1
Private Const CONST_UPDATE = 2
Private Const CONST_DELETE = 3
Private Const CONST_CREATE = 4
Private Const CONST_TABDATA = 5

'Constants for the list boxes on right side
Private Const LIST_REFERRING = 0
Private Const LIST_REFERRED = 1
Private Const LIST_CHECK = 2
Private Const LIST_TRIGGER = 3
Private Const LIST_DEPENDENT = 4

Private objResize As New clsResize

Private objTable As Table
Private objDatabase As Database
Private strTableName As String
Private strQualifiedTableName As String
Private strTableOwnerName As String

Private blnFaster As Boolean 'If true will not use DMO and so will be faster
'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Load()
    lsvTABRows.ColumnHeaderIcons = frmMain.imlSQL
    Call ShowStatusMsg("Populating Tables... Please wait...", True)
    Call PopulateTables(lstTABTables)
    lblTABTotalTables.Caption = "Total Tables : " & Str(lstTABTables.ListCount)
    Call ShowStatusMsg("Ready", False)
    Me.Height = 8100
    Me.Width = 12000
    objResize.Init Me
    objResize.FormResize Me
    Set objDatabase = objSQLServer.Databases(mstrDatabase)
    blnFaster = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Resize()
   objResize.FormResize Me
End Sub

Private Sub txtTABSearch_GotFocus()
    txtTABSearch.SelLength = Len(txtTABSearch.Text)
End Sub

Private Sub txtTABSearch_KeyPress(KeyAscii As Integer)
    'If Enter is hit then execute
    If KeyAscii = 13 Then
        Call lstTABTables_DblClick
    End If
End Sub

Private Sub txtTABSearch_Change()
    If mblnConnected Then Call ListSearch(lstTABTables, txtTABSearch.Text)
End Sub

Private Sub optTABMode_Click(Index As Integer)
    Call lstTABTables_DblClick
End Sub

Private Sub lstTABTables_KeyPress(KeyAscii As Integer)
    'If Enter is hit then execute
    If KeyAscii = 13 Then
        Call lstTABTables_DblClick
    End If
End Sub

Private Sub lstTABTables_Click()
    lstTABTables.ToolTipText = lstTABTables.Text
End Sub

Private Sub lstTABTables_DblClick()
    'Populates the Table specific details
    
    Dim CompleteData As Boolean
    
    If mblnConnected = False Then Exit Sub
    
    On Error GoTo ErrorTrap
    
    
    strTableName = Trim(lstTABTables.Text)
    If strTableName = "" Then Exit Sub
    
    Call ShowStatusMsg("Populating Table rows... Please wait...", True)
    
    txtTABSearch.Text = lstTABTables.Text
    strQualifiedTableName = gfnGetQualifiedName(strTableName)
    strTableOwnerName = gfnGetOwnerName(strTableName)
    If blnFaster = False Then
        Set objTable = objDatabase.Tables(strTableName, strTableOwnerName)
    End If
    lblTABTableName.Caption = "Table Name : " & strQualifiedTableName
    
    If optTABMode(0).Value = True Then
        CompleteData = True
    Else
        CompleteData = False
    End If
    
    Call PopulateDataList(lsvTABRows, strQualifiedTableName, CompleteData)
    lblTABRows.Caption = "Total Showed Rows : " & Str(lsvTABRows.ListItems.Count)
    
    Call ShowStatusMsg("Populating Table Columns... Please wait...", True)
    Call PopulateTableColumns(flxTABColumns, strQualifiedTableName)
    lblTABColumns.Caption = "Total Columns : " & Str(flxTABColumns.Rows - 1)
    
    Call ShowStatusMsg("Populating Referring Tables... Please wait...", True)
    Call PopulateReferringTables(lstRight(LIST_REFERRING), strQualifiedTableName)
    lblTABReferring.Caption = "Referring Tables : " & Str(lstRight(LIST_REFERRING).ListCount)
     
    Call ShowStatusMsg("Populating Referred Tables... Please wait...", True)
    Call PopulateReferredTables(lstRight(LIST_REFERRED), strQualifiedTableName)
    lblTABReferred.Caption = "Referred Tables : " & Str(lstRight(LIST_REFERRED).ListCount)
    
    Call ShowStatusMsg("Populating Check Constraints... Please wait...", True)
    Call PopulateCheckConstraints(lstRight(LIST_CHECK), strQualifiedTableName)
    lblTABConstraints.Caption = "Check Constraints : " & Str(lstRight(LIST_CHECK).ListCount)
    
    Call ShowStatusMsg("Populating Triggers... Please wait...", True)
    Call PopulateTriggers(lstRight(LIST_TRIGGER), strQualifiedTableName)
    lblTABTriggers.Caption = "Triggers : " & Str(lstRight(LIST_TRIGGER).ListCount)
    
    Call ShowStatusMsg("Populating Dependent Objects... Please wait...", True)
    Call PopulateDependentObjects(lstRight(LIST_DEPENDENT), strQualifiedTableName)
    lblTABDepend.Caption = "Dependent Objects : " & Str(lstRight(LIST_DEPENDENT).ListCount)

  
    Call ShowStatusMsg("Ready", False)
    Exit Sub
ErrorTrap:
    MsgBox "Error...Problems encountered while populating table details!" & vbCrLf & _
            "Details:" & Err.Description, vbOKOnly + vbCritical, App.Title
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub lsvTABRows_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'Sorting the data according to the column
    lsvTABRows.Sorted = True
    Call SortListView(lsvTABRows, ColumnHeader)
    'Setting the Column Header Icon
    If lsvTABRows.SortOrder = lvwAscending Then
        ColumnHeader.Icon = "UpArrow"
    Else
        ColumnHeader.Icon = "DownArrow"
    End If
    lsvTABRows.Sorted = False
End Sub

Private Sub flxTABColumns_Click()
    'If the clicked cell is a column name, then
    'displays its content to a text box and highlight it(Useful for copying)
    If flxTABColumns.Col = 1 Then
        If txtTABColName.Visible = False Then txtTABColName.Visible = True
        txtTABColName.Text = flxTABColumns.Text
        txtTABColName.SetFocus
        txtTABColName.SelLength = Len(txtTABColName.Text)
    End If
End Sub

Private Sub txtTABColName_LostFocus()
    txtTABColName.Visible = False
End Sub

Private Sub cmdGenerate_Click()
    'Displays the corresponding SQL Statement in an input box
    Dim intCount As Integer
    Dim strCols As String
    Dim strVals As String
    Dim strStmt As String
    Dim strSQL As String
    
    Dim lngScriptType   As Long
    
    On Error GoTo ErrorTrap
    
    If strTableName = "" Then
        MsgBox "Please select a table from the list", vbInformation, App.Title
        Exit Sub
    End If
    If Trim(txtTABSearch.Text) <> strTableName Then
        Call lstTABTables_DblClick
    End If
    
    Call ShowStatusMsg("Generating statement... Please wait...", True)

    For intCount = 0 To CONST_TABDATA
        If optStmt(intCount).Value = True Then Exit For
    Next intCount
    
    Select Case intCount
        Case CONST_SELECT
            strSQL = "SELECT"
            For intCount = 1 To flxTABColumns.Rows - 1
                strCols = strCols & "," & flxTABColumns.TextMatrix(intCount, 1)
            Next intCount
            strCols = Right(strCols, Len(strCols) - 1) 'Removing First Comma
            strStmt = "SELECT " & strCols & " FROM " & strTableName
        Case CONST_INSERT
            strSQL = "INSERT"
            For intCount = 1 To flxTABColumns.Rows - 1
                strCols = strCols & ", " & flxTABColumns.TextMatrix(intCount, 1)
                strVals = strVals & ", @" & flxTABColumns.TextMatrix(intCount, 1)
            Next intCount
            strCols = Right(strCols, Len(strCols) - 1) 'Removing First Comma
            strVals = Right(strVals, Len(strVals) - 1) 'Removing First Comma
            strStmt = "INSERT INTO " & strTableName & " (" & strCols & ") VALUES (" & strVals & " )"
        Case CONST_UPDATE
            strSQL = "UPDATE"
            For intCount = 1 To flxTABColumns.Rows - 1
                strCols = strCols & flxTABColumns.TextMatrix(intCount, 1) & " = @" & flxTABColumns.TextMatrix(intCount, 1) & ", "
            Next intCount
            'strCols = strCols & " = "
            strCols = Left(strCols, Len(strCols) - 2) 'Removing LAST Comma
            strStmt = "UPDATE " & strTableName & " SET " & strCols
        Case CONST_DELETE
            strSQL = "DELETE"
            strStmt = "DELETE FROM " & strTableName & " WHERE ..."
        Case CONST_CREATE
            strSQL = "CREATE"
            lngScriptType = SQLDMOScript_Drops + SQLDMOScript_Default + _
                  SQLDMOScript_DRI_AllConstraints
                  'SQLDMOScript_IncludeHeaders +
            Set objTable = objDatabase.Tables(strTableName, strTableOwnerName)
            strStmt = objTable.Script(ScriptType:=lngScriptType)
        Case CONST_TABDATA
            Dim intJ As Integer
            Dim strDummy As String
            Dim strInsert As String

            strSQL = "INSERT"
            For intCount = 1 To flxTABColumns.Rows - 1
                strCols = strCols & "," & flxTABColumns.TextMatrix(intCount, 1)
            Next intCount
            strCols = Right(strCols, Len(strCols) - 1) 'Removing First Comma
            strInsert = "INSERT INTO " & strTableName & " (" & strCols & ") VALUES ("
            Set objTable = objDatabase.Tables(strTableName, strTableOwnerName)
            For intCount = 1 To lsvTABRows.ListItems.Count
                'The first column of list view is treated seperately
                If objTable.Columns(lsvTABRows.ColumnHeaders(1).Text).Datatype = "char" Or _
                    objTable.Columns(lsvTABRows.ColumnHeaders(1).Text).Datatype = "varchar" Or _
                    objTable.Columns(lsvTABRows.ColumnHeaders(1).Text).Datatype = "text" Or _
                    objTable.Columns(lsvTABRows.ColumnHeaders(1).Text).Datatype = "nchar" Or _
                    objTable.Columns(lsvTABRows.ColumnHeaders(1).Text).Datatype = "nvarchar" Or _
                    objTable.Columns(lsvTABRows.ColumnHeaders(1).Text).Datatype = "ntext" Or _
                    objTable.Columns(lsvTABRows.ColumnHeaders(1).Text).Datatype = "datetime" Or _
                    objTable.Columns(lsvTABRows.ColumnHeaders(1).Text).Datatype = "smalldatetime" Then
                        'For the above data types, enclose the value in single quotes
                        strDummy = strInsert & "'" & lsvTABRows.ListItems(intCount).Text & "',"
                Else
                        strDummy = strInsert & lsvTABRows.ListItems(intCount).Text & ","
                End If
                
                'Then the remaining columns
                For intJ = 1 To lsvTABRows.ColumnHeaders.Count - 1
                    If objTable.Columns(lsvTABRows.ColumnHeaders(intJ + 1).Text).Datatype = "char" Or _
                     objTable.Columns(lsvTABRows.ColumnHeaders(intJ + 1).Text).Datatype = "varchar" Or _
                     objTable.Columns(lsvTABRows.ColumnHeaders(intJ + 1).Text).Datatype = "text" Or _
                     objTable.Columns(lsvTABRows.ColumnHeaders(intJ + 1).Text).Datatype = "nchar" Or _
                     objTable.Columns(lsvTABRows.ColumnHeaders(intJ + 1).Text).Datatype = "nvarchar" Or _
                     objTable.Columns(lsvTABRows.ColumnHeaders(intJ + 1).Text).Datatype = "ntext" Or _
                     objTable.Columns(lsvTABRows.ColumnHeaders(intJ + 1).Text).Datatype = "datetime" Or _
                     objTable.Columns(lsvTABRows.ColumnHeaders(intJ + 1).Text).Datatype = "smalldatetime" Then
                        'For the above data types, enclose the value in single quotes
                        strDummy = strDummy & "'" & Trim(lsvTABRows.ListItems(intCount).ListSubItems(intJ).Text) & "',"
                    Else
                        strDummy = strDummy & lsvTABRows.ListItems(intCount).ListSubItems(intJ).Text & ","
                    End If
                Next intJ
                strDummy = Left(strDummy, Len(strDummy) - 1) 'Removing last Comma
                strStmt = strStmt & strDummy & ")" & vbCrLf
            Next intCount
        Case Else
            Exit Sub
    End Select
    
    Call ShowStatusMsg("Ready", False)
    
    frmCodes.mstrTitle = "Following is the generated " & strSQL & " SQL statement for the " & _
                " table : " & strQualifiedTableName
    frmCodes.mstrCodes = strStmt
    frmCodes.mintDefaultFileType = FILETYPE_SQL
    frmCodes.mblnNoColoring = False
    frmCodes.mstrFileName = strTableName
    frmCodes.Show vbModal, frmMain
    
    Exit Sub
ErrorTrap:
    MsgBox "Error...Problems encountered while generating statements!" & vbCrLf & _
            "Details:" & Err.Description, vbOKOnly + vbCritical, App.Title
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub lstRight_KeyPress(Index As Integer, KeyAscii As Integer)
    'If Enter is hit then execute double click event
    If KeyAscii = 13 Then
        lstRight_DblClick (Index)
    End If
End Sub

Private Sub lstRight_Click(Index As Integer)
    'While Clicking, the text is shown as tool tip
    lstRight(Index).ToolTipText = lstRight(Index).Text
End Sub

Private Sub lstRight_DblClick(Index As Integer)
    'Populate accordingly
    Dim strFullText() As String
    Dim strText As String
    
    Dim objCheck As SQLDMO.Check
    Dim objTrigger As SQLDMO.Trigger
    
    If Trim(lstRight(Index).Text) = "" Then Exit Sub
    
    Call ShowStatusMsg("Populating Details... Please wait...", True)

    On Error Resume Next
    
    strFullText = Split(Trim(lstRight(Index).Text), "(")
    If UBound(strFullText) > 0 Then
        strText = Trim(strFullText(0))
    Else
        strText = Trim(lstRight(Index).Text)
    End If
    
    Select Case Index
        Case LIST_REFERRING, LIST_REFERRED
            txtTABSearch.Text = strText
            Call lstTABTables_DblClick
        Case LIST_CHECK
            If blnFaster = False Then
                Set objCheck = objTable.Checks(strText)
                frmCodes.mstrCodes = objCheck.Text
            Else
                frmCodes.mstrCodes = GetTextFromSysCom(strText)
            End If
            With frmCodes
                .mstrTitle = "Following is the source of the Check Constraint '" & strText & "' on the Table '" & strQualifiedTableName & "'"
                .mintDefaultFileType = FILETYPE_SQL
                .mblnNoColoring = False
                .mstrFileName = strText
                Call ShowStatusMsg("Ready", False)
                .Show vbModal, frmMain
            End With
        Case LIST_TRIGGER
            If blnFaster = False Then
                Set objTrigger = objTable.Triggers(strText)
                frmCodes.mstrCodes = objTrigger.Text
            Else
                frmCodes.mstrCodes = GetTextFromSysCom(strText)
            End If
            With frmCodes
                .mstrTitle = "Following is the source of the Trigger '" & strText & "' on the Table '" & strQualifiedTableName & "'"
                .mintDefaultFileType = FILETYPE_SQL
                .mblnNoColoring = False
                .mstrFileName = strText
                Call ShowStatusMsg("Ready", False)
                .Show vbModal, frmMain
            End With
        Case LIST_DEPENDENT
            Call PopulateObjectText(rtbTABDummy, strText)
            With frmCodes
                .mstrTitle = "Following is the source of the object '" & strText & "' which depends on the Table '" & strQualifiedTableName & "'"
                .mstrCodes = rtbTABDummy.Text
                .mintDefaultFileType = FILETYPE_SQL
                .mblnNoColoring = False
                .mstrFileName = strText
                Call ShowStatusMsg("Ready", False)
                .Show vbModal, frmMain
            End With
    End Select
    Call ShowStatusMsg("Ready", False)
End Sub

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------

Private Sub PopulateTables(ListBoxName As ListBox)
    'Populates the Tables in a databse into the Listbox
    
    Dim rsTables As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Distinct Name From SysObjects Where Xtype='U' Order By Name"
    rsTables.Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
    ListBoxName.Clear
    While Not rsTables.EOF
        ListBoxName.AddItem rsTables("Name")
        rsTables.MoveNext
    Wend
    
    rsTables.Close
    Set rsTables = Nothing
End Sub

Private Sub PopulateTableColumns(FlexGridName As MSFlexGrid, TableName As String)
    'Populates the Column details of passed Table into the Flex Grid
    
    Dim rsCols As ADODB.Recordset
    Dim rsPK As ADODB.Recordset
    Dim strSQL As String
    Dim strNullable As String
    Dim i As Variant
    Dim blnPK As Boolean
    Dim blnPKCheck As Boolean

    Dim intCounter As Integer
    Dim objColumn As SQLDMO.Column
    
    If blnFaster = False Then
        'Initialising the Flex Grid
        With FlexGridName
            .FixedCols = 1
            .FixedRows = 1
            .Rows = 1
            '.FormatString = "No.|Column Name" & Space(15) & "|Data Type|Length|Precision|Scale|IsNullable"
            .FormatString = "No.|Column Name" & Space(15) & "|Data Type|Length|Precision|Scale|IsNullable|Default|Identity|Computed" & Space(15)
        End With
        
        intCounter = 0
        For Each objColumn In objTable.Columns
            intCounter = intCounter + 1
        
            FlexGridName.AddItem IIf(objColumn.InPrimaryKey, "x  " & intCounter, intCounter) & Chr(9) & _
                                objColumn.Name & Chr(9) & _
                                objColumn.Datatype & Chr(9) & _
                                objColumn.Length & Chr(9) & _
                                objColumn.NumericPrecision & Chr(9) & _
                                objColumn.NumericScale & Chr(9) & _
                                IIf(objColumn.AllowNulls, "Yes", "") & Chr(9) & _
                                objColumn.DRIDefault.Text & Chr(9) & _
                                IIf(objColumn.Identity, "Yes", "") & Chr(9) & _
                                IIf(objColumn.IsComputed, objColumn.ComputedText, "") & Chr(9)
        Next
    Else
        'Initialising the Flex Grid
        With FlexGridName
            .FixedCols = 1
            .FixedRows = 1
            .Rows = 1
            .FormatString = "No.|Column Name" & Space(15) & "|Data Type|Length|Precision|Scale|IsNullable"
            '.FormatString = "No.|Column Name" & Space(15) & "|Data Type|Length|Precision|Scale|IsNullable|Default|Identity|Computed" & Space(15)
        End With
        
        'The conventional method, using sys tables; codes kept commented for reference
        Set rsPK = New ADODB.Recordset
        'rsPK.Open "sp_pkeys('" & TableName & "')", mconGeneral
        rsPK.Open "sp_pkeys @table_name = '" & strTableName & "', @table_owner ='" & strTableOwnerName & "' ,@table_qualifier = '" & mstrDatabase & "'", mconGeneral
        blnPKCheck = True
        If rsPK.EOF Then blnPKCheck = False
    
        Set rsCols = New ADODB.Recordset
        'Constructing SQL Query
        strSQL = "Select SysColumns.Name ColName,SysTypes.Name DataType,SysColumns.Length Length," & _
                 " SysColumns.XPrec Prec, SysColumns.XScale Scale,SysColumns.IsNullable Nullable " & _
                 " From SysColumns, SysTypes" & _
                 " Where   Id=Object_Id(N'" & TableName & "') And" & _
                 " SysColumns.XUserType = SysTypes.XUserType"
        With rsCols
            .Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
            While Not .EOF
                'Adding data to flex grid
                strNullable = IIf(.Fields("Nullable") = 1, "Yes", "")
                If blnPKCheck = True Then rsPK.MoveFirst
                blnPK = False
                While Not rsPK.EOF
                    If rsPK.Fields("COLUMN_NAME") = .Fields("ColName") Then
                        blnPK = True
                    End If
                    rsPK.MoveNext
                Wend
    
                If blnPK = True Then 'If Primary Key, add a mark
                    FlexGridName.AddItem "x  " & FlexGridName.Rows & Chr(9) & _
                                .Fields("ColName") & Chr(9) & _
                                .Fields("DataType") & Chr(9) & _
                                .Fields("Length") & Chr(9) & _
                                .Fields("Prec") & Chr(9) & _
                                .Fields("Scale") & Chr(9) & _
                                strNullable & Chr(9)
                Else
                    FlexGridName.AddItem FlexGridName.Rows & Chr(9) & _
                                .Fields("ColName") & Chr(9) & _
                                .Fields("DataType") & Chr(9) & _
                                .Fields("Length") & Chr(9) & _
                                .Fields("Prec") & Chr(9) & _
                                .Fields("Scale") & Chr(9) & _
                                strNullable & Chr(9)
                End If
                .MoveNext
            Wend
            .Close
        End With
        rsPK.Close
        Set rsPK = Nothing
        Set rsCols = Nothing
    End If
End Sub

Private Sub PopulateDependentObjects(ListBoxName As ListBox, TableName As String)
    'Populates the Dependent Objects of passed Table into the List Box
    Dim rsDep As ADODB.Recordset
    Dim strSQL As String
    Dim intCounter As Integer
    Dim i As Variant
    
    'Initialising the List Box
    ListBoxName.Clear
    
    Set rsDep = New ADODB.Recordset
    'Constructing SQL Query
    strSQL = "SELECT DISTINCT SysUsers2.Name As UserName,SysObjects2.Name As Dependent,SysObjects2.Type " & _
             " From SysObjects " & _
             " INNER JOIN SysUsers ON SysObjects.uid = SysUsers.uid " & _
             " INNER JOIN SysDepends ON SysObjects.id = SysDepends.depid " & _
             " INNER JOIN SysObjects SysObjects2 ON SysDepends.id = SysObjects2.id " & _
             " INNER JOIN SysUsers SysUsers2 ON SysObjects2.uid = SysUsers2.uid " & _
             " Where  SysObjects.xtype = 'U' " & _
             " AND SysObjects2.xtype <> 'C' " & _
             " AND SysObjects2.xtype <> 'TR' " & _
             " And SysObjects.Id = Object_Id(N'" & TableName & "')"
             
    With rsDep
        .Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            'Adding data to list box
            ListBoxName.AddItem .Fields("Dependent") & " ( " & .Fields("Type") & ")"
            .MoveNext
        Wend
        .Close
    End With

    Set rsDep = Nothing
End Sub

Private Sub PopulateReferringTables(ListBoxName As ListBox, TableName As String)
    'Populates the Referring Tables of passed Table into the List Box
    Dim rsReferring As ADODB.Recordset
    Dim strSQL As String
    Dim intCounter As Integer
    Dim i As Variant
    
    'Initialising the List Box
    ListBoxName.Clear
    
    Set rsReferring = New ADODB.Recordset
    'Constructing SQL Query
    strSQL = "Select  Distinct O.Name TableName, C.Name FieldName " & _
             " From    SysReferences R, SysObjects O, SysColumns C, SysConstraints T" & _
             " Where O.Id = R.RKeyId And R.ConstId = T.ConstId And C.Id = T.Id And C.ColId = T.ColId" & _
             " And R.FKeyId = Object_Id(N'" & TableName & "')"
    With rsReferring
        .Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            'Adding data to flex grid
            ListBoxName.AddItem .Fields("TableName") & " ( " & .Fields("FieldName") & " )"
            .MoveNext
        Wend
        .Close
    End With

    Set rsReferring = Nothing
End Sub

Private Sub PopulateReferredTables(ListBoxName As ListBox, TableName As String)
    'Populates the Referring Tables of passed Table into the List Box
    Dim rsReferred As ADODB.Recordset
    Dim strSQL As String
    Dim intCounter As Integer
    Dim i As Variant
    
    'Initialising the List Box
    ListBoxName.Clear
    
    Set rsReferred = New ADODB.Recordset
    'Constructing SQL Query
    strSQL = "Select Distinct O.Name TableName" & _
             " From    SysReferences R, SysObjects O" & _
             " Where   O.Id =R.FKeyId And R.RKeyId = object_id(N'" & TableName & "')"
    With rsReferred
        .Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            'Adding data to flex grid
            ListBoxName.AddItem .Fields("TableName")
            .MoveNext
        Wend
        .Close
    End With

    Set rsReferred = Nothing
End Sub

Private Sub PopulateCheckConstraints(ListBoxName As ListBox, TableName As String)
    'Populates the Check Constraints of passed Table into the List Box
    
    If blnFaster = False Then
        Dim objCheck As SQLDMO.Check
        'Initialising the List Box
        ListBoxName.Clear
        
        For Each objCheck In objTable.Checks
            ListBoxName.AddItem objCheck.Name
        Next
    Else
        Dim rsCheck As ADODB.Recordset
        Dim strSQL As String
        Dim intCounter As Integer
        Dim i As Variant
        
        'Initialising the List Box
        ListBoxName.Clear
        
        Set rsCheck = New ADODB.Recordset
        'Constructing SQL Query
        strSQL = "Select Distinct OCheck.Name CheckName,OParent.Name TableName, Com.Text" & _
                 " From   SysObjects OCheck , SysObjects OParent , SysComments Com" & _
                 " Where   OCheck.parent_obj = OParent.Id" & _
                 " And OCheck.xtype='C' And Com.id = OCheck.id" & _
                 " And OParent.Id = object_id(N'" & TableName & "')"
        With rsCheck
            .Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
            While Not .EOF
                'Adding data to flex grid
                ListBoxName.AddItem .Fields("CheckName")
                .MoveNext
            Wend
            .Close
        End With
    
        Set rsCheck = Nothing
    End If
End Sub

Private Sub PopulateTriggers(ListBoxName As ListBox, TableName As String)
    'Populates the Triggers of passed Table into the List Box
    If blnFaster = False Then
        Dim objTrigger As SQLDMO.Trigger
        'Initialising the List Box
        ListBoxName.Clear
        
        For Each objTrigger In objTable.Triggers
            ListBoxName.AddItem objTrigger.Name
        Next
    Else
        Dim rsCheck As ADODB.Recordset
        Dim strSQL As String
        Dim intCounter As Integer
        Dim i As Variant
        
        'Initialising the List Box
        ListBoxName.Clear
        
        Set rsCheck = New ADODB.Recordset
        'Constructing SQL Query
        strSQL = "Select Name From SysObjects where xtype='TR' and parent_obj = object_id(N'" & TableName & "')"
        With rsCheck
            .Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
            While Not .EOF
                'Adding data to flex grid
                ListBoxName.AddItem .Fields("Name")
                .MoveNext
            Wend
            .Close
        End With
    
        Set rsCheck = Nothing
    End If
End Sub

'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
