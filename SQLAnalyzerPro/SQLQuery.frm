VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmQuery 
   Caption         =   "Query Executer"
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
   Begin VB.CommandButton cmdQRYExecute 
      Caption         =   "&Run (F5)"
      Height          =   315
      Left            =   9765
      TabIndex        =   2
      Top             =   105
      Width           =   975
   End
   Begin VB.CommandButton cmdQRYClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   10815
      TabIndex        =   3
      Top             =   105
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid flxQRYOutput 
      Height          =   3015
      Left            =   105
      TabIndex        =   4
      ToolTipText     =   "Result set"
      Top             =   3570
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   5318
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbQRYInput 
      Height          =   3015
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Type the SQL statement here and hit Run button to execute"
      Top             =   420
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"SQLQuery.frx":0000
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
   Begin RichTextLib.RichTextBox rtbQRYMessages 
      Height          =   825
      Left            =   105
      TabIndex        =   5
      ToolTipText     =   "System messages"
      Top             =   6720
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   1455
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"SQLQuery.frx":0080
   End
   Begin VB.Label lblQRYBox 
      AutoSize        =   -1  'True
      Caption         =   "&Query Box:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   780
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Query Executer
'Description    :   To execute the typed SQL Queries
'Developed By   :   Sameer C T
'Started On     :   2001 November 27
'Last Modified  :   2003 April 16
'-------------------------------------------------------------------------------------------

Option Explicit

Private objResize As New clsResize

'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Load()
    Me.Icon = frmLogin.Icon
    Me.Height = 8100
    Me.Width = 12000
    objResize.Init Me
    objResize.FormResize Me
    rtbQRYInput.SelStart = 1
End Sub

Private Sub Form_Resize()
   objResize.FormResize Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 27 Then Unload Me
End Sub

Private Sub rtbQRYInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then Call cmdQRYExecute_Click
End Sub

Private Sub rtbQRYInput_Change()
    Dim lngPos As Long
    lngPos = rtbQRYInput.SelStart
    Call gprColorKeyWords(rtbQRYInput)
    rtbQRYInput.SelStart = lngPos
End Sub

Private Sub cmdQRYExecute_Click()
    'Executes the query
    Dim strSQL As String
    
    On Error GoTo ErrorTrap
    
    If Trim(rtbQRYInput.SelText) <> "" Then 'If text is marked as selected then take it
        strSQL = Trim(rtbQRYInput.SelText)
    Else
        strSQL = Trim(rtbQRYInput.Text) 'else take the whole text
    End If
    
    If strSQL = "" Then Exit Sub
    
    Call ShowStatusMsg("Query Executing... Please Wait...", True)
    rtbQRYMessages.Text = "Executing... Please Wait..."
    rtbQRYMessages.Text = FillFlexWithRecords(strSQL, flxQRYOutput)
    
    rtbQRYMessages.SelLength = Len(rtbQRYMessages.Text)
    If Trim(rtbQRYMessages.Text) = "Command Completed Successfully" Then
        rtbQRYMessages.SelColor = vbBlue
    Else
        rtbQRYMessages.SelColor = vbRed
    End If
    Call ShowStatusMsg("Ready", False)
    Exit Sub
ErrorTrap:
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub cmdQRYClear_Click()
    rtbQRYInput.Text = ""
    rtbQRYInput.SetFocus
End Sub

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------

Private Function FillFlexWithRecords(SQLQuery As String, FlexGridName As MSFlexGrid) As String
    'Fills the Flex with the output of passed Query
    
    On Error GoTo ErrorTrap
    Dim rsFill As ADODB.Recordset
    Dim intI As Integer
    Dim intJ As Integer
    Dim intColCount As Integer
    Dim strFormat As String
    
    Set rsFill = New ADODB.Recordset
    rsFill.Open SQLQuery, mconGeneral, adOpenKeyset, adLockOptimistic
    
    'Building the Formatstring for Flexgrid using Column names in the Recordset
    intColCount = rsFill.Fields.Count
    If intColCount > 0 Then
        strFormat = ""
        For intI = 0 To intColCount - 1
            If strFormat = "" Then
                strFormat = strFormat & Trim(rsFill.Fields.Item(intI).Name)
            Else
                strFormat = strFormat & " | " & Trim(rsFill.Fields.Item(intI).Name)
            End If
        Next intI
    
        With FlexGridName
            .Rows = 0
            .Rows = 2
            .FixedCols = 0
            .FixedRows = 1
            .FormatString = strFormat
            .Cols = intColCount
            If rsFill.RecordCount > 0 Then
                For intI = 0 To rsFill.RecordCount - 1
                    .Rows = .Rows + 1
                    For intJ = 0 To intColCount - 1
                        .TextMatrix(intI + 1, intJ) = IIf(IsNull(rsFill(intJ)) = True, "NULL", rsFill(intJ))
                    Next intJ
                    rsFill.MoveNext
                Next intI
            End If
        End With
        
        rsFill.Close
        Set rsFill = Nothing
    End If
    FillFlexWithRecords = "Command Completed Successfully"
    Exit Function
ErrorTrap:
    With FlexGridName
        .Rows = 0
        .Rows = 2
        .Cols = 1
        .FixedCols = 0
        .FixedRows = 1
    End With
    FillFlexWithRecords = Err.Description
End Function
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
