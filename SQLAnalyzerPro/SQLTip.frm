VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   2985
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "SQLTip.frx":0000
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   4
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Tips
'Description    :   Provides useful tips for the application
'Developed By   :   Sameer C T (Courtsey : Microsoft)
'Started On     :   2001 November 27
'Last Modified  :   2003 May 09
'-------------------------------------------------------------------------------------------

Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "SQLTips.dat"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long

Private blnFirstTime As Boolean
'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Load()
    Dim ShowAtStartup As Long
    
    blnFirstTime = True
    Me.Icon = frmLogin.Icon
    
    
''    'The following codes can be used to load the tips from a text file
    ' Seed Rnd
    Randomize
''
''    ' Read in the tips file and display a tip at random.
''    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
''        lblTipText.Caption = "File Not Found : " & TIP_FILE & vbCrLf & vbCrLf & _
''           "This file should be present in the same directory as the application. "
''    End If
    
    Call InitialiseTips
    Call DoNextTip
End Sub

Private Sub cmdNextTip_Click()
    Call DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------

Private Sub DoNextTip()
   If blnFirstTime = True Then
    'Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    blnFirstTime = False
   Else
    'Cycle through the Tips in order
    CurrentTip = CurrentTip + 1
   End If
    
    If Tips.Count < CurrentTip Then
        CurrentTip = 1
    End If
    
    ' Show it.
    frmTip.DisplayCurrentTip
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub
Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub InitialiseTips()
    With Tips
        .Add "General:" & vbCrLf & "The Find box helps you to quickly find listed database objects by typing each letter of them, and the bottom list will automatically scroll to the item if it is present"
        .Add "General:" & vbCrLf & "You can switch between various windows(Tables, Views,...) using Ctrl+Tab, if they are already loaded using menu"
        .Add "General:" & vbCrLf & "If an object name is too long, click on it to view it in full as tool tip"
        .Add "Connection:" & vbCrLf & "Specify Server option populates all available SQL servers in the network"
        .Add "Connection:" & vbCrLf & "When Server is specified, all available databases are populated"
        .Add "Connection:" & vbCrLf & "Use the Connect menu to use another database server or another database in the same server"
        .Add "Refresh:" & vbCrLf & "Clicking on the Refresh menu will refresh the database objects in SQL Analyzer after closing all the open windows and re-opening the Table window with latest objects"
        .Add "Tables:" & vbCrLf & "Double click on a listed Table name to view its contents, column details, references and dependent objects"
        .Add "Tables:" & vbCrLf & "The cross mark near the column number indicates that it is a Primary Key column"
        .Add "Tables:" & vbCrLf & "The column details grid provide more informations like computed columns, default values, identity columns, etc besides data type, size and others"
        .Add "Tables:" & vbCrLf & "Click on the column name to view it completely or to copy it for use while programming"
        .Add "Tables:" & vbCrLf & "The Statements section generates various SQL statements for the selected table, which can be used for programming"
        .Add "Tables:" & vbCrLf & "The generated statements appear in a code window which shows key words in different colors"
        .Add "Tables:" & vbCrLf & "The code window can be maximised to get better view of the contents. The contents can be copied to clipboard or saved to a file"
        .Add "Tables:" & vbCrLf & "The Data option in the Statements section generates Insert SQL statements to export data in a table"
        .Add "Tables:" & vbCrLf & "The Dependent Objects section, shows objects dependent on the selected table; (P) = Stored Procedure and (V) = View and (FN) = Function"
        .Add "Tables:" & vbCrLf & "Double click on the dependent object name to view its source text in the code window"
        .Add "Tables:" & vbCrLf & "Double click on a Check Constraint name or Trigger name to view its source text in the code window"
        .Add "Tables:" & vbCrLf & "Double click on the referred or referring table names to get their details populated"
        .Add "Tables and Views:" & vbCrLf & "Sort the rows in a Table or View by clicking on the Column Headings. Clicking the same heading again will change the sort order"
        .Add "Tables and Views:" & vbCrLf & "The All rows option poopulates all the data, but it may take more time if the table or view contains large number of records"
        .Add "Tables and Views:" & vbCrLf & "If there are more than 500 rows in table or view, SQL Analyzer warns you about this and you can view either 500 rows or all rows"
        .Add "SPs:" & vbCrLf & "The Code section generates VB and SQL codes for the selected Stored Procedure"
        .Add "Search:" & vbCrLf & "Use the Search window to search for a particular text like table name or author name in the database objects (SPs, Views, Functions)"
        .Add "Search:" & vbCrLf & "Use the Find button on the right corner to find the search text in the displayed object"
        .Add "Search:" & vbCrLf & "Click again on the Find button to find further occurrence of the search text"
        .Add "Reports:" & vbCrLf & "The Reports window helps to generate various reports on database objects"
        .Add "Reports:" & vbCrLf & "You can select tables with specified prefixes and either add them to the current selection or keep as a new set of selection for generating report"
        .Add "Reports:" & vbCrLf & "The Table Details option generates report on selected tables giving their column details like Name, Data Type, Size, etc."
        .Add "Reports:" & vbCrLf & "The Object Listing option generates a list of database objects with their details like Owner, Created Date, etc."
        .Add "Reports:" & vbCrLf & "Since the report is generated in Excel, you can further use its features to analyse reports. For eg., Sorting the list of tables by Rows will help to identify tables with more data"
        .Add "Reports:" & vbCrLf & "The progress bar will indicate the report generation progress and you are free to close the window during this process, if you want to quit"
        .Add "Backup:" & vbCrLf & "SQL Analyzer provides the most easiest way to take complete backup of a database into a disk file. It is just the matter of clicking Backup menu and selecting the destination"
        .Add "Backup:" & vbCrLf & "If an exisitng file name is given, SQL Analyzer overwrites the backup file. Path for the file should be in the same machine where the SQL server is running"
        .Add "Restore:" & vbCrLf & "Restore option allows you to restore a database, in a simplest way. It will list the computers connected to the database, giving you an option to forcefully disconnect them and proceed with restoring"
        .Add "Restore:" & vbCrLf & "The backup file should be picked from the same machine where the SQL server is running. SQL Analyzer refreshes itself after successful restoring of database"
    End With
End Sub

'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
