VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCodes 
   Caption         =   "Code Window"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgCode 
      Left            =   315
      Top             =   5310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save..."
      Height          =   315
      Left            =   3885
      TabIndex        =   4
      ToolTipText     =   "Save the content to a file"
      Top             =   5340
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      ToolTipText     =   "Close this window"
      Top             =   5340
      Width           =   975
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   315
      Left            =   4942
      TabIndex        =   1
      ToolTipText     =   "Copy the content to clipboard"
      Top             =   5340
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtbCodes 
      Height          =   4485
      Left            =   75
      TabIndex        =   0
      Top             =   750
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   7911
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"SQLCodes.frx":0000
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
   Begin VB.Label lblTitle 
      Caption         =   "Title"
      Height          =   585
      Left            =   105
      TabIndex        =   3
      Top             =   120
      Width           =   6825
   End
End
Attribute VB_Name = "frmCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Code Window
'Description    :   Displays the generated codes
'Developed By   :   Sameer C T
'Started On     :   2003 March 10
'Last Modified  :   2003 April 10
'-------------------------------------------------------------------------------------------

Option Explicit

Public mstrCodes As String
Public mstrTitle As String
Public mstrFileName As String
Public mintDefaultFileType As Integer
Public mblnNoColoring As Boolean

Private objResize As New clsResize

Private Sub Form_Load()
    Call ShowStatusMsg("Loading Code Window... Please wait...", True)
    Me.Icon = frmLogin.Icon
    lblTitle.Caption = mstrTitle
    rtbCodes.Text = mstrCodes
    If mblnNoColoring = False Then
        Call gprColorKeyWords(rtbCodes)
    End If
    rtbCodes.SelLength = 0
    Me.Height = 6200
    Me.Width = 7185
    objResize.Init Me
    objResize.FormResize Me
    Call ShowStatusMsg("Ready", False)
End Sub

Private Sub Form_Resize()
   objResize.FormResize Me
End Sub

Private Sub cmdSave_Click()
   dlgCode.Filter = "SQL Script Files (*.sql)|*.sql|Text Files (*.txt)|*.txt"
   dlgCode.FileName = mstrFileName
   dlgCode.FilterIndex = mintDefaultFileType
   dlgCode.ShowSave
   rtbCodes.SaveFile dlgCode.FileName, rtfText
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText (rtbCodes.Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
