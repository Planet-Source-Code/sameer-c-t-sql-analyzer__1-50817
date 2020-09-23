VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmFunctions 
   Caption         =   "Functions"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   6870
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstFN 
      Height          =   6690
      Left            =   150
      TabIndex        =   1
      ToolTipText     =   "List of available functions"
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtFNSearch 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      ToolTipText     =   "Type here to find the functions"
      Top             =   330
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox rtbFNText 
      Height          =   4920
      Left            =   2775
      TabIndex        =   2
      ToolTipText     =   " The source of the function"
      Top             =   2730
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8678
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"SQLFunctions.frx":0000
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
   Begin MSFlexGridLib.MSFlexGrid flxFNSParams 
      Height          =   2085
      Left            =   2805
      TabIndex        =   6
      ToolTipText     =   "Parameters of the function"
      Top             =   315
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   3678
      _Version        =   393216
   End
   Begin VB.Label lblFNSParams 
      Caption         =   "Total &Parameters:"
      Height          =   195
      Left            =   2805
      TabIndex        =   7
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblFNName 
      Caption         =   "Function &Name"
      Height          =   195
      Left            =   2790
      TabIndex        =   5
      Top             =   2490
      Width           =   6210
   End
   Begin VB.Label lblFNSSearch 
      AutoSize        =   -1  'True
      Caption         =   "&Find:"
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   120
      Width           =   345
   End
   Begin VB.Label lblFNTotal 
      Caption         =   "&Total Functions"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   735
      Width           =   2535
   End
End
Attribute VB_Name = "frmFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Functions
'Description    :   To search the Functions and to view its content
'Developed By   :   Sameer C T
'Started On     :   2001 November 27
'Last Modified  :   2003 September 25
'-------------------------------------------------------------------------------------------

Option Explicit

Private objResize As New clsResize
Private strQualifiedFunctionName As String
'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Load()
    Me.Icon = frmLogin.Icon
    Call ShowStatusMsg("Populating Functions... Please wait...", True)
    Call PopulateFunctions(lstFN)
    
    lblFNTotal.Caption = "&Total Functions : " & Str(lstFN.ListCount)
    
    Me.Height = 8100
    Me.Width = 12000
    objResize.Init Me
    objResize.FormResize Me
    
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Resize()
   objResize.FormResize Me
End Sub

Private Sub txtFNSearch_GotFocus()
    txtFNSearch.SelLength = Len(txtFNSearch.Text)
End Sub

Private Sub txtFNSearch_KeyPress(KeyAscii As Integer)
    'If Enter is hit then execute
    If KeyAscii = 13 Then
        Call lstFN_DblClick
    End If
End Sub

Private Sub txtFNSearch_Change()
    If mblnConnected Then Call ListSearch(lstFN, txtFNSearch.Text)
End Sub


Private Sub lstFN_KeyPress(KeyAscii As Integer)
    'If Enter is hit then execute
    If KeyAscii = 13 Then
        Call lstFN_DblClick
    End If
End Sub

Private Sub lstFN_Click()
    lstFN.ToolTipText = lstFN.Text
End Sub

Private Sub lstFN_DblClick()
    If mblnConnected = False Then Exit Sub
    
    On Error GoTo ErrorTrap
    
    txtFNSearch.Text = Trim(lstFN.Text)
    strQualifiedFunctionName = gfnGetQualifiedName(Trim(lstFN.Text), True)
    lblFNName.Caption = "&Name:" & strQualifiedFunctionName
    
    Call ShowStatusMsg("Populating Parameters... Please wait...", True)
    Call frmSPs.PopulateSPParams(flxFNSParams, strQualifiedFunctionName)
    'Assumes that the first parameter is always the Return parameter for functions
    flxFNSParams.TextMatrix(1, 6) = "Return"
    lblFNSParams.Caption = "Total Parameters : " & Str(flxFNSParams.Rows - 1)
    If Val(flxFNSParams.Rows) = 1 Then flxFNSParams.Rows = 2

    
    Call ShowStatusMsg("Filling Object Text.. Please wait...", True)
    Call PopulateObjectText(rtbFNText, strQualifiedFunctionName)
    'Call gprColorKeyWords(rtbFNText)
    Call ShowStatusMsg("Ready", False)
    Exit Sub
ErrorTrap:
    MsgBox "Error...Problems encountered while populating Object details!" & vbCrLf & _
            "Details:" & Err.Description, vbOKOnly + vbCritical, App.Title
    Call ShowStatusMsg("Ready", False)
End Sub

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------

Private Sub PopulateFunctions(ListBoxName As ListBox)
    'Populates Functions in a databse to the list box
    
    Dim rsTables As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Distinct Name From SysObjects Where Xtype='FN' Or Xtype='TF' Order By Name"
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


