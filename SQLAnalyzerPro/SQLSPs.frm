VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSPs 
   Caption         =   "Stored Procedures"
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
   Begin VB.Frame fraCode 
      Caption         =   "&Code"
      Height          =   2130
      Left            =   10305
      TabIndex        =   8
      Top             =   210
      Width           =   1470
      Begin VB.OptionButton optCode 
         Caption         =   "&Visual Basic 2"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   10
         ToolTipText     =   "Select this to generate VB code for the SP"
         Top             =   765
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate"
         Height          =   315
         Left            =   210
         TabIndex        =   12
         ToolTipText     =   "Click to generate the selected code"
         Top             =   1545
         Width           =   975
      End
      Begin VB.OptionButton optCode 
         Caption         =   "S &Q L"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   11
         ToolTipText     =   "Select this to generate SQL code for the SP"
         Top             =   1140
         Width           =   1050
      End
      Begin VB.OptionButton optCode 
         Caption         =   "&Visual Basic 1"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   9
         ToolTipText     =   "Select this to generate VB code for the SP"
         Top             =   390
         Width           =   1320
      End
   End
   Begin VB.TextBox txtSPSSearch 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Type here to find the SPs"
      Top             =   315
      Width           =   2535
   End
   Begin VB.ListBox lstSPS 
      Height          =   6690
      Left            =   105
      TabIndex        =   3
      ToolTipText     =   "List of available SPs"
      Top             =   960
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox rtbSPSText 
      Height          =   4920
      Left            =   2730
      TabIndex        =   5
      ToolTipText     =   " The source of the SP"
      Top             =   2730
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8678
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"SQLSPs.frx":0000
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
   Begin MSFlexGridLib.MSFlexGrid flxSPSParams 
      Height          =   2085
      Left            =   2730
      TabIndex        =   6
      ToolTipText     =   "Parameters of the SP"
      Top             =   285
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   3678
      _Version        =   393216
   End
   Begin VB.Label lblSPSParams 
      Caption         =   "Total &Parameters:"
      Height          =   195
      Left            =   2730
      TabIndex        =   7
      Top             =   105
      Width           =   7335
   End
   Begin VB.Label lblSPSTotalSPs 
      Caption         =   "&Total SPs"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   720
      Width           =   2355
   End
   Begin VB.Label lblSPSSearch 
      AutoSize        =   -1  'True
      Caption         =   "&Find:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   345
   End
   Begin VB.Label lblSPSSPName 
      Caption         =   "SP &Name"
      Height          =   195
      Left            =   2760
      TabIndex        =   4
      Top             =   2520
      Width           =   6675
   End
End
Attribute VB_Name = "frmSPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Stored Procedures
'Description    :   To search the SPs and to view its content
'Developed By   :   Sameer C T
'Started On     :   2001 November 27
'Last Modified  :   2003 September 25
'-------------------------------------------------------------------------------------------

Option Explicit

Private Const CONST_VB = 0
Private Const CONST_VBPARAM = 1
Private Const CONST_SQL = 2
Private objResize As New clsResize

'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Load()
    Me.Icon = frmLogin.Icon
    Call ShowStatusMsg("Populating SPs... Please wait...", True)
    Call PopulateSPs(lstSPS)
    lblSPSTotalSPs.Caption = "&Total SPs : " & Str(lstSPS.ListCount)
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

Private Sub txtSPSSearch_GotFocus()
    txtSPSSearch.SelLength = Len(txtSPSSearch.Text)
End Sub

Private Sub txtSPSSearch_KeyPress(KeyAscii As Integer)
    'If Enter is hit then execute
    If KeyAscii = 13 Then
        Call lstSPS_DblClick
    End If
End Sub

Private Sub txtSPSSearch_Change()
    If mblnConnected Then Call ListSearch(lstSPS, txtSPSSearch.Text)
End Sub

Private Sub lstSPS_KeyPress(KeyAscii As Integer)
    'If Enter is hit then execute
    If KeyAscii = 13 Then
        Call lstSPS_DblClick
    End If
End Sub

Private Sub lstSPS_Click()
    lstSPS.ToolTipText = lstSPS.Text
End Sub

Private Sub lstSPS_DblClick()
    If mblnConnected = False Then Exit Sub
    Dim strSPName As String
    
    On Error GoTo ErrorTrap
    
    strSPName = Trim(lstSPS.Text)
    txtSPSSearch.Text = strSPName
    
    strSPName = gfnGetQualifiedName(strSPName, True)

    lblSPSSPName.Caption = "&Name:" & strSPName
    
    Call ShowStatusMsg("Populating Parameters... Please wait...", True)
    Call PopulateSPParams(flxSPSParams, strSPName)
    lblSPSParams.Caption = "Total Parameters : " & Str(flxSPSParams.Rows - 1)
    If Val(flxSPSParams.Rows) = 1 Then flxSPSParams.Rows = 2
 
    Call ShowStatusMsg("Filling Object Text.. Please wait...", True)
    Call PopulateObjectText(rtbSPSText, strSPName)
    'Call gprColorKeyWords(rtbSPSText)
    Call ShowStatusMsg("Ready", False)
    Exit Sub
ErrorTrap:
    MsgBox "Error...Problems encountered while populating Object details!" & vbCrLf & _
            "Details:" & Err.Description, vbOKOnly + vbCritical, App.Title
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub cmdGenerate_Click()
    'Generates the selected codes
    Dim intCount As Integer
    Dim strCols As String
    Dim strVals As String
    Dim strStmt As String
    Dim strSQL As String
    Dim strSPName As String
    Dim strOutput As String
    Dim strOPV As String
    Dim strOutputDics As String
    Dim strOutputVars As String
    Dim strOutputPrints As String
    
    On Error GoTo ErrorTrap
    
    For intCount = 0 To CONST_SQL
        If optCode(intCount).Value = True Then Exit For
    Next intCount
    
    strSPName = Trim(txtSPSSearch.Text)
    
    Select Case intCount
        Case CONST_VB
            strVals = "VB"
            strSQL = "Dim objCMD As ADODB.Command" & Chr(13) & _
                        "Set objCMD = New ADODB.Command" & Chr(13) & _
                        "With objCMD" & Chr(13) & vbTab & _
                        ".ActiveConnection = ?" & Chr(13) & vbTab & _
                        ".CommandType = adCmdStoredProc" & Chr(13) & vbTab & _
                        ".CommandText = """ & strSPName & """ " & Chr(13) & vbTab
                      
            For intCount = 1 To flxSPSParams.Rows - 1
               If Trim(flxSPSParams.TextMatrix(intCount, 1)) <> "" Then
                If UCase(Trim(flxSPSParams.TextMatrix(intCount, 6))) = "INPUT" Then
                    strCols = strCols & ".Parameters(""" & flxSPSParams.TextMatrix(intCount, 1) & """)= ?" & Chr(13) & vbTab
                Else 'Output Parameter
                    strOutput = strOutput & "? = " & ".Parameters(""" & flxSPSParams.TextMatrix(intCount, 1) & """)" & Chr(13) & vbTab
                End If
               End If
            Next intCount
            strStmt = strSQL & strCols
            strStmt = strStmt & ".Execute" & Chr(13)
            If Len(strOutput) > 0 Then
                strOutput = Left(strOutput, Len(strOutput) - 1)
                strStmt = strStmt & vbTab & strOutput
            End If
            strStmt = strStmt & "End With" & Chr(13) & "Set objCMD = Nothing"
            frmCodes.mintDefaultFileType = FILETYPE_TXT
            frmCodes.mblnNoColoring = True
        Case CONST_VBPARAM
            strVals = "VB"
            strSQL = "Dim objCMD As ADODB.Command" & Chr(13) & _
                        "Set objCMD = New ADODB.Command" & Chr(13) & _
                        "With objCMD" & Chr(13) & vbTab & _
                        ".ActiveConnection = ?" & Chr(13) & vbTab & _
                        ".CommandType = adCmdStoredProc" & Chr(13) & vbTab & _
                        ".CommandText = """ & strSPName & """ " & Chr(13) & vbTab
                      
            For intCount = 1 To flxSPSParams.Rows - 1
               If Trim(flxSPSParams.TextMatrix(intCount, 1)) <> "" Then
                If UCase(Trim(flxSPSParams.TextMatrix(intCount, 6))) = "INPUT" Then
                    strCols = strCols & ".Parameters.Append .CreateParameter(""" & flxSPSParams.TextMatrix(intCount, 1) & """, " & mfnGetDataTypeConstant(flxSPSParams.TextMatrix(intCount, 2)) & ", adParamInput, " & Trim(flxSPSParams.TextMatrix(intCount, 3)) & ", ?)" & Chr(13) & vbTab
                Else 'Output Parameter
                    strCols = strCols & ".Parameters.Append .CreateParameter(""" & flxSPSParams.TextMatrix(intCount, 1) & """, " & mfnGetDataTypeConstant(flxSPSParams.TextMatrix(intCount, 2)) & ", adParamOutput, " & Trim(flxSPSParams.TextMatrix(intCount, 3)) & ")" & Chr(13) & vbTab
                    strOutput = strOutput & "? = " & ".Parameters(""" & flxSPSParams.TextMatrix(intCount, 1) & """)" & Chr(13) & vbTab
                End If
                If UCase(Trim(flxSPSParams.TextMatrix(intCount, 2))) = "NUMERIC" Then
                     strCols = strCols & ".Parameters(""" & flxSPSParams.TextMatrix(intCount, 1) & """).Precision = " & Trim(flxSPSParams.TextMatrix(intCount, 4)) & Chr(13) & vbTab
                     strCols = strCols & ".Parameters(""" & flxSPSParams.TextMatrix(intCount, 1) & """).NumericScale = " & Trim(flxSPSParams.TextMatrix(intCount, 5)) & Chr(13) & vbTab
                End If
               End If
            Next intCount
            strStmt = strSQL & strCols
            strStmt = strStmt & ".Execute" & Chr(13)
            If Len(strOutput) > 0 Then
                strOutput = Left(strOutput, Len(strOutput) - 1)
                strStmt = strStmt & vbTab & strOutput
            End If
            strStmt = strStmt & "End With" & Chr(13) & "Set objCMD = Nothing"
            frmCodes.mintDefaultFileType = FILETYPE_TXT
            frmCodes.mblnNoColoring = True
        Case CONST_SQL
            strVals = "SQL"
            strSQL = "Execute " & strSPName & " "
            For intCount = 1 To flxSPSParams.Rows - 1
               If Trim(flxSPSParams.TextMatrix(intCount, 1)) <> "" Then
                If UCase(Trim(flxSPSParams.TextMatrix(intCount, 6))) = "INPUT" Then
                    strCols = strCols & ", " & flxSPSParams.TextMatrix(intCount, 1) & "= ?"
                Else 'Output Parameter
                    strOPV = Replace(flxSPSParams.TextMatrix(intCount, 1), "@", "@p")
                    
                     'In case of strings, specify their size in bracket
                     If UCase(Trim(flxSPSParams.TextMatrix(intCount, 2))) = "VARCHAR" _
                      Or UCase(Trim(flxSPSParams.TextMatrix(intCount, 2))) = "CHAR" Then
                        strOutputDics = strOutputDics & "DECLARE " & strOPV & " AS " & UCase(Trim(flxSPSParams.TextMatrix(intCount, 2))) & "(" & Trim(flxSPSParams.TextMatrix(intCount, 3)) & ")" & vbCrLf
                     ElseIf UCase(Trim(flxSPSParams.TextMatrix(intCount, 2))) = "NUMERIC" Then
                        strOutputDics = strOutputDics & "DECLARE " & strOPV & " AS " & UCase(Trim(flxSPSParams.TextMatrix(intCount, 2))) & "(" & Trim(flxSPSParams.TextMatrix(intCount, 4)) & "," & Trim(flxSPSParams.TextMatrix(intCount, 5)) & ")" & vbCrLf
                     Else
                        strOutputDics = strOutputDics & "DECLARE " & strOPV & " AS " & UCase(Trim(flxSPSParams.TextMatrix(intCount, 2))) & vbCrLf
                     End If
                    strOutputPrints = strOutputPrints & "PRINT " & strOPV & vbCrLf
                    strOutputVars = strOutputVars & ", " & flxSPSParams.TextMatrix(intCount, 1) & "= " & strOPV & " OUTPUT"
                End If
               End If
            Next intCount
            If Len(strCols) > 0 Then strCols = Right(strCols, Len(strCols) - 1) 'Removing First Comma
            'strStmt = strSQL & strCols
            strStmt = strOutputDics & strSQL & strCols & strOutputVars & vbCrLf & strOutputPrints
            frmCodes.mintDefaultFileType = FILETYPE_SQL
            frmCodes.mblnNoColoring = False
    End Select
    
    frmCodes.mstrTitle = "Following is the generated " & strVals & " code " & _
                " for the stored procedure : " & strSPName
    frmCodes.mstrFileName = strSPName
    frmCodes.mstrCodes = strStmt
    frmCodes.Show vbModal, frmMain
    
    Exit Sub
ErrorTrap:
    MsgBox "Error...Problems encountered while generating code!" & vbCrLf & _
            "Details:" & Err.Description, vbOKOnly + vbCritical, App.Title
    Call ShowStatusMsg("Ready", False)
End Sub

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------

Private Sub PopulateSPs(ListBoxName As ListBox)
    'Populates Stored Procedures in a databse to the list box
    
    Dim rsTables As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Distinct Name From SysObjects Where Xtype='P' Order By Name"
    rsTables.Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
    ListBoxName.Clear
    While Not rsTables.EOF
        ListBoxName.AddItem rsTables("Name")
        rsTables.MoveNext
    Wend
    
    rsTables.Close
    Set rsTables = Nothing
End Sub

Public Sub PopulateSPParams(FlexGridName As MSFlexGrid, SPName As String)
    'Populates the Column details of passed Table into the Flex Grid
    
    Dim rsCols As ADODB.Recordset
    Dim strSQL As String
    Dim strNullable As String
    Dim intCounter As Integer
    Dim i As Variant
    
    'Initialising the Flex Grid
    With FlexGridName
        .FixedCols = 1
        .FixedRows = 1
        .Rows = 1
        .FormatString = "No.|Parameter Name" & Space(35) & "|Data Type|Length|Precision|Scale|Parameter Type"
    End With
    
    Set rsCols = New ADODB.Recordset
    'Constructing SQL Query
    strSQL = "Select SysColumns.Name ColName,SysTypes.Name DataType,SysColumns.Length Length," & _
             " SysColumns.XPrec Prec, SysColumns.XScale Scale,SysColumns.IsOutParam OutParam " & _
             " From SysColumns, SysTypes" & _
             " Where   Id=Object_Id('" & SPName & "') And" & _
             " SysColumns.XUserType = SysTypes.XUserType"
    With rsCols
        .Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            'Adding data to flex grid
            strNullable = IIf(.Fields("OutParam") = 1, "Output", "Input")
            FlexGridName.AddItem FlexGridName.Rows & Chr(9) & _
                                .Fields("ColName") & Chr(9) & _
                                .Fields("DataType") & Chr(9) & _
                                .Fields("Length") & Chr(9) & _
                                .Fields("Prec") & Chr(9) & _
                                .Fields("Scale") & Chr(9) & _
                                strNullable & Chr(9)
            .MoveNext
        Wend
        .Close
    End With

    Set rsCols = Nothing
End Sub

Private Function mfnGetDataTypeConstant(strDataType As String) As String
   strDataType = UCase(Trim(strDataType))
   Select Case strDataType
      Case "BIGINT"
         mfnGetDataTypeConstant = "adBigInt"
      Case "INT"
         mfnGetDataTypeConstant = "adInteger"
      Case "SMALLINT"
         mfnGetDataTypeConstant = "adSmallInt"
      Case "TINYINT"
         mfnGetDataTypeConstant = "adTinyInt"
      Case "BIT"
         mfnGetDataTypeConstant = "adBoolean"
      Case "DECIMAL"
         mfnGetDataTypeConstant = "adDecimal"
      Case "NUMERIC"
         mfnGetDataTypeConstant = "adNumeric"
      Case "MONEY"
         mfnGetDataTypeConstant = "adCurrency"
       Case "SMALLMONEY"
         mfnGetDataTypeConstant = "adCurrency"
      Case "DATETIME"
         mfnGetDataTypeConstant = "adDate"
      Case "CHAR"
         mfnGetDataTypeConstant = "adChar"
      Case "VARCHAR"
         mfnGetDataTypeConstant = "adVarChar"
      Case Else
         mfnGetDataTypeConstant = "adIUnknown"
   End Select
End Function
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
