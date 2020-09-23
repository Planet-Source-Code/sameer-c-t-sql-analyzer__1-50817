VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmViews 
   Caption         =   "Views"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtVWSSearch 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Type here to find the views"
      Top             =   315
      Width           =   2535
   End
   Begin VB.ListBox lstVWSViews 
      Height          =   6690
      Left            =   105
      TabIndex        =   3
      ToolTipText     =   "List of available views"
      Top             =   945
      Width           =   2535
   End
   Begin VB.OptionButton optVWSMode 
      Caption         =   "Top 10"
      Height          =   180
      Index           =   1
      Left            =   3780
      TabIndex        =   5
      ToolTipText     =   "Select this to show only top 10 rows in the view"
      Top             =   105
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optVWSMode 
      Caption         =   "All Rows"
      Height          =   180
      Index           =   0
      Left            =   2730
      TabIndex        =   4
      ToolTipText     =   "Select this to show all data in the view"
      Top             =   105
      Width           =   1095
   End
   Begin MSComctlLib.ListView lsvVWSRows 
      Height          =   2310
      Left            =   2730
      TabIndex        =   8
      ToolTipText     =   "Data in the selected view"
      Top             =   315
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   4075
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
   Begin RichTextLib.RichTextBox rtbVWSText 
      Height          =   4905
      Left            =   2730
      TabIndex        =   9
      ToolTipText     =   "The source of view"
      Top             =   2730
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   8652
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"SQLViews.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblVWSRows 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Showed Rows :"
      Height          =   195
      Left            =   9660
      TabIndex        =   7
      Top             =   105
      Width           =   2055
   End
   Begin VB.Label lblVWSTotalViews 
      Caption         =   "TotalViews"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   735
      Width           =   2520
   End
   Begin VB.Label lblVWSViewName 
      Caption         =   "ViewName"
      Height          =   195
      Left            =   5040
      TabIndex        =   6
      Top             =   105
      Width           =   4470
   End
   Begin VB.Label lblVWSSearch 
      AutoSize        =   -1  'True
      Caption         =   "&Find:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   345
   End
End
Attribute VB_Name = "frmViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Views
'Description    :   To search the views and to see its content and data
'Developed By   :   Sameer C T
'Started On     :   2001 November 27
'Last Modified  :   2003 September 25
'-------------------------------------------------------------------------------------------

Option Explicit

Private objResize As New clsResize

'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Load()
    Me.Icon = frmLogin.Icon
    lsvVWSRows.ColumnHeaderIcons = frmMain.imlSQL
    lsvVWSRows.Sorted = True
    Call ShowStatusMsg("Populating Views... Please wait...", True)
    Call PopulateViews(lstVWSViews)
    lblVWSTotalViews.Caption = "Total Views : " & Str(lstVWSViews.ListCount)
    Me.Height = 8100
    Me.Width = 12000
    objResize.Init Me
    objResize.FormResize Me
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub Form_Resize()
   objResize.FormResize Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 27 Then Unload Me
End Sub

Private Sub txtVWSSearch_GotFocus()
    txtVWSSearch.SelLength = Len(txtVWSSearch.Text)
End Sub

Private Sub txtVWSSearch_KeyPress(KeyAscii As Integer)
    'If Enter is hit then execute
    If KeyAscii = 13 Then
        Call lstVWSViews_DblClick
    End If
End Sub

Private Sub txtVWSSearch_Change()
    If mblnConnected Then Call ListSearch(lstVWSViews, txtVWSSearch.Text)
End Sub

Private Sub optVWSMode_Click(Index As Integer)
    Call lstVWSViews_DblClick
End Sub

Private Sub lstVWSViews_KeyPress(KeyAscii As Integer)
    'If Enter is hit then execute
    If KeyAscii = 13 Then
        Call lstVWSViews_DblClick
    End If
End Sub

Private Sub lstVWSViews_Click()
    lstVWSViews.ToolTipText = lstVWSViews.Text
End Sub

Private Sub lstVWSViews_DblClick()
    'Populates the View specific details
    
    Dim CompleteData As Boolean
    Dim strViewName As String
    
    If mblnConnected = False Then Exit Sub
    
    On Error GoTo ErrorTrap
    
    If optVWSMode(0).Value = True Then
        CompleteData = True
    Else
        CompleteData = False
    End If
    
    strViewName = Trim(lstVWSViews.Text)
    If strViewName = "" Then Exit Sub

    txtVWSSearch.Text = strViewName
    
    strViewName = gfnGetQualifiedName(strViewName)
    
    Call ShowStatusMsg("Populating data... Please wait...", True)
    Call PopulateDataList(lsvVWSRows, strViewName, CompleteData)
    lblVWSRows.Caption = "Total Showed Rows : " & Str(lsvVWSRows.ListItems.Count)
    lblVWSViewName.Caption = "View Name : " & strViewName
    
    Call ShowStatusMsg("Filling View Text.. Please wait...", True)
    Call PopulateObjectText(rtbVWSText, strViewName)
    'Call gprColorKeyWords(rtbVWSText)
    
    Call ShowStatusMsg("Ready", False)
    Exit Sub
ErrorTrap:
    MsgBox "Error...Problems encountered while populating view details!" & vbCrLf & _
            "Details:" & Err.Description, vbOKOnly + vbCritical, App.Title
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub lsvVWSRows_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'Sorting the data according to the column
    Call SortListView(lsvVWSRows, ColumnHeader)
    'Setting the Column Header Icon
    If lsvVWSRows.SortOrder = lvwAscending Then
        ColumnHeader.Icon = "UpArrow"
    Else
        ColumnHeader.Icon = "DownArrow"
    End If
End Sub

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------

Private Sub PopulateViews(ListBoxName As ListBox)
    'Populates Views in a databse to the Listbox
    
    Dim rsTables As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Distinct Name From SysObjects Where Xtype='V' And Category = 0 Order By Name"
    rsTables.Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
    ListBoxName.Clear
    While Not rsTables.EOF
        ListBoxName.AddItem rsTables("Name")
        rsTables.MoveNext
    Wend
    
    rsTables.Close
    Set rsTables = Nothing
End Sub
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
