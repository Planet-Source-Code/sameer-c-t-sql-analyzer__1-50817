VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSearch 
   Caption         =   "Search"
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
   Begin VB.CommandButton cmdFind 
      Height          =   285
      Left            =   11415
      Picture         =   "SQLSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Finds the word in Search box"
      Top             =   60
      Width           =   360
   End
   Begin VB.ListBox lstSPS 
      Height          =   5520
      Left            =   135
      TabIndex        =   6
      ToolTipText     =   "List of objects matching the search"
      Top             =   2055
      Width           =   2535
   End
   Begin VB.TextBox txtSPSSearch 
      Height          =   315
      Left            =   135
      TabIndex        =   4
      ToolTipText     =   "Type here to find a listed object"
      Top             =   1380
      Width           =   2535
   End
   Begin VB.TextBox txtSPSTextSearch 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      ToolTipText     =   "Type the word to be searched here and click on the button below"
      Top             =   405
      Width           =   2535
   End
   Begin VB.CommandButton cmdSPSTextSearch 
      Caption         =   "&List Objects"
      Height          =   300
      Left            =   735
      TabIndex        =   2
      ToolTipText     =   "Click to start searching"
      Top             =   795
      Width           =   1170
   End
   Begin RichTextLib.RichTextBox rtbSPSText 
      Height          =   7185
      Left            =   2760
      TabIndex        =   9
      ToolTipText     =   "The source of the selected object"
      Top             =   390
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   12674
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"SQLSearch.frx":0102
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
   Begin VB.Label lblSPSSPName 
      Caption         =   "Object &Name"
      Height          =   195
      Left            =   2790
      TabIndex        =   8
      Top             =   150
      Width           =   6450
   End
   Begin VB.Label lblSPSSearch 
      AutoSize        =   -1  'True
      Caption         =   "&Find:"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   1170
      Width           =   345
   End
   Begin VB.Label lblSPSTotalSPs 
      Caption         =   "&Total Objects"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   1830
      Width           =   2445
   End
   Begin VB.Label lblSPSTextSearcha 
      AutoSize        =   -1  'True
      Caption         =   "Search for a &Text in Objects:"
      Height          =   195
      Left            =   330
      TabIndex        =   0
      Top             =   150
      Width           =   2025
   End
End
Attribute VB_Name = "frmSearch"
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
'Last Modified  :   2003 April 16
'-------------------------------------------------------------------------------------------

Option Explicit

Private objResize As New clsResize
Private strObjectName As String
Private strObjectType As String
Private strQualifiedObjectName As String

'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Load()
    Me.Icon = frmLogin.Icon
    Me.Height = 8100
    Me.Width = 12000
    objResize.Init Me
    objResize.FormResize Me
End Sub

Private Sub Form_Resize()
   objResize.FormResize Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 27 Then Unload Me
End Sub

Private Sub txtSPSTextSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdSPSTextSearch_Click
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

Private Sub cmdSPSTextSearch_Click()
    If Trim(txtSPSTextSearch.Text) = "" Then Exit Sub
    Call ShowStatusMsg("Searching... Please wait...", True)
    If mblnConnected Then Call PopulateSearchObjects(lstSPS, Trim(txtSPSTextSearch.Text))
    lblSPSTotalSPs.Caption = "&Total Objects : " & Str(lstSPS.ListCount)
    lblSPSSPName.Caption = ""
    txtSPSSearch.Text = ""
    rtbSPSText.Text = ""
    Call ShowStatusMsg("Ready", False)
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
    Dim strFullText() As String
    
    If mblnConnected = False Then Exit Sub

    On Error GoTo ErrorTrap
    
    strObjectName = Trim(lstSPS.Text)
    
    strFullText = Split(strObjectName, "(")
    If UBound(strFullText) > 0 Then
        strObjectName = Trim(strFullText(0))
        strObjectType = Trim(strFullText(1))
    End If
    txtSPSSearch.Text = strObjectName
    If InStr(strObjectType, "V") > 0 Then 'View
        strQualifiedObjectName = gfnGetQualifiedName(strObjectName, False)
    Else 'SP or Function
        strQualifiedObjectName = gfnGetQualifiedName(strObjectName, True)
    End If
    
    lblSPSSPName.Caption = "&Name:" & strQualifiedObjectName
    
    Call ShowStatusMsg("Filling Object Text... Please wait...", True)
    Call PopulateObjectText(rtbSPSText, strQualifiedObjectName)
    rtbSPSText.SelLength = 0
    'Call gprColorKeyWords(rtbSPSText)
    Call ShowStatusMsg("Ready", False)
    Exit Sub
ErrorTrap:
    MsgBox "Error...Problems encountered while populating Object details!" & vbCrLf & _
            "Details:" & Err.Description, vbOKOnly + vbCritical, App.Title
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub cmdFind_Click()
    'Finds a word in the object
    Static lngPos As Long
    Static intCount As Integer
    
    On Error Resume Next
    rtbSPSText.SetFocus
    lngPos = rtbSPSText.Find(Trim(txtSPSTextSearch.Text), lngPos + Len(Trim(txtSPSTextSearch.Text)))
    intCount = intCount + 1
    If lngPos = -1 Then
        MsgBox "Search Completed... The word '" & Trim(txtSPSTextSearch.Text) & "' not found!" & _
         vbCrLf & "Total number of occurrence : " & intCount - 1
        intCount = 0
    End If
End Sub

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------
Private Sub PopulateSearchObjects(ListBoxName As ListBox, SearchText As String)
    'Populates Searched objects in a databse to the list box
    
    Dim rsTables As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Distinct SysObjects.Name,SysObjects.Type From SysObjects,SysComments " & _
            " Where SysObjects.Id= SysComments.Id And SysObjects.parent_obj = 0 And " & _
            " SysComments.Text Like '%" & SearchText & "%' Order By Name"
    rsTables.Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
    ListBoxName.Clear
    While Not rsTables.EOF
        ListBoxName.AddItem rsTables("Name") & " ( " & rsTables("Type") & ")"
        rsTables.MoveNext
    Wend
    
    rsTables.Close
    Set rsTables = Nothing
End Sub
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
