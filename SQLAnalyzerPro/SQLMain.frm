VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "SQL Analyzer"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   915
   ClientWidth     =   4680
   Icon            =   "SQLMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   1185
      Top             =   1005
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbSQL 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2940
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8819
            MinWidth        =   8819
            Object.ToolTipText     =   "Messages"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
            MinWidth        =   2999
            Object.ToolTipText     =   "Server"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
            Object.ToolTipText     =   "Database"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "User"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Text            =   "Ins"
            TextSave        =   "10/10/2003"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2293
            MinWidth        =   2293
            Text            =   "Caps"
            TextSave        =   "16:35"
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prbSQL 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   688
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList imlSQL 
      Left            =   0
      Top             =   1035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SQLMain.frx":1D12
            Key             =   "UpArrow"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SQLMain.frx":2164
            Key             =   "DownArrow"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SQLMain.frx":25B6
            Key             =   "Drive"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SQLMain.frx":2A08
            Key             =   "ClosedFolder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SQLMain.frx":2E5A
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SQLMain.frx":32AC
            Key             =   "File"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SQLMain.frx":36FE
            Key             =   "BakFile"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connect"
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "Re&fresh"
   End
   Begin VB.Menu mnuTables 
      Caption         =   "&Tables"
   End
   Begin VB.Menu mnuViews 
      Caption         =   "&Views"
   End
   Begin VB.Menu mnuSPs 
      Caption         =   "&SPs"
   End
   Begin VB.Menu mnuFunctions 
      Caption         =   "Fu&nctions"
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "&Query"
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "S&earch"
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
   End
   Begin VB.Menu mnuBackup 
      Caption         =   "Bac&kup"
   End
   Begin VB.Menu mnuRestore 
      Caption         =   "Rest&ore"
   End
   Begin VB.Menu mnuTips 
      Caption         =   "Ti&ps"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "A&bout"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   5.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Main
'Description    :   The main MDI forms containing other forms
'Developed By   :   Sameer C T
'Started On     :   2001 November 27
'Last Modified  :   2003 May 09
'-------------------------------------------------------------------------------------------
'History:

'2001 November (Version 1)
'Started From Soft Sytems Cochin, to handle hundreds of database objects in an efficient way.
'It was time consuming to use SQL commands in MS Query Analyzer for the frequently needed operations.
'Moreover, Enterprise Manager was resource hungry and the existing tools developed by colleagues
'lacked the flexibility and good looks.
'So started developing this, using tabbed pane as main interface and having options just to display
'table column details and data, view text and stored procedure text and a poorly working Query box.

'2002 January (Version 2)
'While using this tool for working, and after getting feedbacks from colleagues, fixed many bugs
'and added some more features like showing dependent objects of table, adding functions tab,
'searching words in database objects, etc.

'2002 November (Version 3)
'After about one year since its development, this tool was found very useful in day to day work
'for Soft Systems, which demanded handling of numerous database objects in different databases.
'Moreover, the tool became popular among some colleagues also and I was encouraged to add more
'features. While I was working in Nigeria, for Okomu Oil Palm Company, added some more features
'like auto generation of SQL statements for tables, Parameter display of SPs, VB codes for SPs,etc.

'2003 April (Version 4)
'When I was working in Kenya, for Sasini Tea and Coffe Limited, I got enough free time again and
'decided to stop spending time on this tool after making it completely professional. So enhanced
'the autogeneration of codes, fixed the bugs in cloring of keywords, added Reports, Backup, Restore
'and Refresh options, provided find button in the search window and beutified the whole windows
'including the tips and about windows.
'The final version is named as SQL Analyzer Pro (Professional Edition), under the banner of
'Sameeriya Soft (the latest name I decided, after Samtech and Sams World).

'2003 May 09
'In the table section, displayed triggers and check constraints
'In the table columns added Identity, Computed, Default

'2003 September 19-25 (Version 5)
'Added splash screen and made the loading of apln more faster by avoiding the use of SQL DMO for table dtls
'Used full qualified names for objects in all screens to avoid the bugs if owner is not dbo
'Backup and restore now shows a window with the server directory structure

'ToDo:
'more col dtls like default, etc

Option Explicit
   
Private WithEvents objBackup As SQLDMO.Backup
Attribute objBackup.VB_VarHelpID = -1
Private WithEvents objRestore As SQLDMO.Restore
Attribute objRestore.VB_VarHelpID = -1
'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub mnuConnect_Click()
    frmLogin.Show vbModal, Me
End Sub

Private Sub mnuRefresh_Click()
    'Refreshing contents
    Dim frmA As Form
    
    'Opening the connection after closing it to get refreshed data
    If mconGeneral.State = adStateOpen Then mconGeneral.Close
    mconGeneral.Open mstrConString
    
  'Disconnecting and connecting again the DMO object
   objSQLServer.Disconnect
   objSQLServer.Connect CStr(mstrServer), CStr(mstrUserName), CStr(mstrPassword)
    
    'Close all the opened windows and show the table window (with new data)
    For Each frmA In Forms
        If frmA.Name <> "frmMain" Then
            Unload frmA
        End If
    Next
    frmTables.Show
End Sub

Private Sub mnuTables_Click()
    frmTables.Show
    frmTables.ZOrder (0)
End Sub

Private Sub mnuViews_Click()
    frmViews.Show
    frmViews.ZOrder (0)
End Sub

Private Sub mnuSPs_Click()
    frmSPs.Show
    frmSPs.ZOrder (0)
End Sub

Private Sub mnuFunctions_Click()
    frmFunctions.Show
    frmFunctions.ZOrder (0)
End Sub

Private Sub mnuQuery_Click()
    frmQuery.Show
    frmQuery.ZOrder (0)
End Sub

Private Sub mnuSearch_Click()
    frmSearch.Show
    frmSearch.ZOrder (0)
End Sub

Private Sub mnuReports_Click()
   frmReports.Show vbModal, Me
End Sub

Private Sub mnuBackup_Click()
   'Backup
   On Error GoTo ErrorTrap
   
   'Showing Server Directory Form to choose the backup destination path
   frmServerDir.mblnFilesNeeded = False
   frmServerDir.mstrBakFileName = mstrDatabase & "_" & Format(Date, "yyyyMMMMdd") & "_" & Format(Time, "hhmmAMPM") & ".bak"
   frmServerDir.Show vbModal, Me
   If mblnCancelDirectory = True Then Exit Sub
   Call ShowStatusMsg("Backup in progress... Please wait...", True)
   Call mprBackupDB(frmServerDir.mstrBakFileName)
   Call ShowStatusMsg("Ready", False)
   Exit Sub
ErrorTrap:
    If Err.Number = -2147218303 Then
        MsgBox "Backup operation was not completed!" & vbCrLf & _
            "Details: Try again after selecting a file location in the database server " & objSQLServer.NetName, vbInformation, App.Title
    Else
        MsgBox "Backup operation was not completed!" & vbCrLf & _
                "Details:" & Err.Description, vbInformation, App.Title
    End If
   Unload frmProgress
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub mnuRestore_Click()
   'Restore
   On Error GoTo ErrorTrap
   
   'Showing Server Directory Form to choose the backup destination path
   frmServerDir.mblnFilesNeeded = True
   frmServerDir.mstrBakFileName = ""
   frmServerDir.Show vbModal, Me
   If mblnCancelDirectory = True Then Exit Sub
   
   Call ShowStatusMsg("Restore in progress... Please wait...", True)
   Call mprRestoreDB(frmServerDir.mstrBakFileName)
   Call ShowStatusMsg("Ready", False)
   Exit Sub
ErrorTrap:
    If Err.Number = -2147218303 Then
        MsgBox "Restore operation was not completed!" & vbCrLf & _
            "Details: Try again after selecting a file local to the database server " & objSQLServer.NetName, vbInformation, App.Title
    Else
        MsgBox "Restore operation was not completed!" & vbCrLf & _
                "Details:" & Err.Description, vbInformation, App.Title
    End If
    Unload frmProgress
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub mnuTips_Click()
    frmTip.Show vbModal, Me
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
    'frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'Confiramtion(I think its not required nowadays..)
    'If MsgBox("Do you really want to exit?", vbYesNo + vbQuestion) = vbNo Then Cancel = 1
    
    'Clean up
    On Error Resume Next
    objSQLServer.Disconnect
    If mconGeneral.State = adStateOpen Then
        mconGeneral.Close
        Set mconGeneral = Nothing
    End If
    End
End Sub
'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------
 Private Sub mprBackupDB(pstrFileName As String)
   'Backs up the present database
   

''   'Showing Save As Dialog box to get the file name
''   With dlgMain
''      .CancelError = True
''      .Filter = "Database Backup Files (*.bak)|*.bak|All Files (*.*)|*.*"
''      .FileName = mstrDatabase & "_" & Format(Date, "yyyyMMMMdd") & "_" & Format(Time, "hhmmAMPM")
''      .FilterIndex = 0
''      .ShowSave
''   End With
   
   'Showing the Progress screen
   frmProgress.prgProgress.Min = 0
   frmProgress.prgProgress.Max = 100
   frmProgress.Show
   
   'Backing up the db
   Set objBackup = New SQLDMO.Backup
   With objBackup
      .PercentCompleteNotification = 1
      .Action = SQLDMOBackup_Database
      .Initialize = True 'Overwrites exisitng media
      .BackupSetDescription = "Backup created by SQL Analyzer"
      .BackupSetName = mstrDatabase
      .MediaDescription = "Backup created by SQL Analyzer"
      .Database = mstrDatabase
      .Files = pstrFileName 'dlgMain.FileName
      .SQLBackup objSQLServer
   End With
   
   Unload frmProgress
End Sub

Private Sub objBackup_PercentComplete(ByVal Message As String, ByVal Percent As Long)
   frmProgress.prgProgress.Value = Percent
End Sub

Private Sub objBackup_Complete(ByVal Message As String)
   MsgBox "Backup of database '" & mstrDatabase & "' from server '" & mstrServer & "' completed successfully!", vbInformation
End Sub

Private Sub mprRestoreDB(pstrFileName As String)
   'Restoring the present database
   Dim intUsers As Integer
   Dim strUsers As String
   Dim rsKill As ADODB.Recordset
   Dim frmA As Form
   
''   'Showing Open Dialog box to get the file name
''   With dlgMain
''      .CancelError = True
''      .Filter = "Database Backup Files (*.bak)|*.bak|All Files (*.*)|*.*"
''      .FileName = ""
''      .FilterIndex = 0
''      .ShowOpen
''   End With
   
   'Disconnecting and connecting again the connections to server (safety)
   objSQLServer.Disconnect
   objSQLServer.Connect CStr(mstrServer), CStr(mstrUserName), CStr(mstrPassword)

   If mconGeneral.State = adStateOpen Then mconGeneral.Close
   mconGeneral.Open mstrConString
   
   
   'Finding out the computers connected to this database
   Set rsKill = New ADODB.Recordset
   
    Set rsKill = mconGeneral.Execute("sp_who '" & mstrUserName & "'")
    rsKill.Filter = "dbname='" & mstrDatabase & "'"
    intUsers = 0
    strUsers = ""
    Do While Not rsKill.EOF
        intUsers = intUsers + 1
        strUsers = strUsers & intUsers & ". " & Trim(rsKill("hostname")) & vbCrLf
        rsKill.MoveNext
    Loop
    
    'If its more than one (ie, other than this connection), give messsage
    If intUsers > 1 Then
        If (MsgBox("The following computer(s) are connected to this database: " & _
            vbCrLf & vbCrLf & strUsers & vbCrLf & _
            "Do you want to disconnect them and continue restoring?" & vbCrLf & _
            "(Click No to abort restoring and Yes to proceed with restoring)", vbYesNo + vbQuestion) = vbNo) Then
            rsKill.Close
            Set rsKill = Nothing
            Exit Sub
        End If
    End If
    
    'If the user wants to proceed, terminate those connections
    rsKill.MoveFirst
    Do While Not rsKill.EOF
        mconGeneral.Execute "KILL " & rsKill("spid")
        rsKill.MoveNext
        intUsers = intUsers + 1
    Loop
    rsKill.Close
    Set rsKill = Nothing
    
    'Destroy the connection object
    If mconGeneral.State = adStateOpen Then mconGeneral.Close
    Set mconGeneral = Nothing
    
   'Showing the Progress screen
   frmProgress.prgProgress.Min = 0
   frmProgress.prgProgress.Max = 100
   frmProgress.Show
    
   'Restoring the db
   Set objRestore = New SQLDMO.Restore
   With objRestore
      .PercentCompleteNotification = 1
      .Action = SQLDMORestore_Database
      .Database = mstrDatabase
      .Files = pstrFileName 'dlgMain.FileName
      objSQLServer.Databases(mstrDatabase).DBOption.SingleUser = True
      .SQLRestore objSQLServer
      objSQLServer.Databases(mstrDatabase).DBOption.SingleUser = False
   End With
   
   Unload frmProgress
       
    'Opening the connection for further use
    mconGeneral.Open mstrConString
    
    'Close all the opened windows and show the table window (with new restored db)
    For Each frmA In Forms
        If frmA.Name <> "frmMain" Then
            Unload frmA
        End If
    Next
    frmTables.Show
End Sub

Private Sub objRestore_PercentComplete(ByVal Message As String, ByVal Percent As Long)
   frmProgress.prgProgress.Value = Percent
End Sub

Private Sub objRestore_Complete(ByVal Message As String)
   MsgBox "Restore of database '" & mstrDatabase & "' to server '" & mstrServer & "' completed successfully!", vbInformation
End Sub
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------

