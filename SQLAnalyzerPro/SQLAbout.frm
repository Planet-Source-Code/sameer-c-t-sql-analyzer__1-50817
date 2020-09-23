VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3600
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2484.784
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   4785
      Top             =   2265
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4545
      TabIndex        =   5
      Top             =   2535
      Width           =   975
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Sys Info..."
      Height          =   315
      Left            =   4545
      TabIndex        =   0
      Top             =   2970
      Width           =   975
   End
   Begin VB.Label lblProduct 
      Caption         =   "This product is licensed to:"
      Height          =   285
      Left            =   1485
      TabIndex        =   8
      Top             =   1215
      Width           =   3885
   End
   Begin VB.Image imgAbout 
      BorderStyle     =   1  'Fixed Single
      Height          =   1830
      Left            =   150
      Picture         =   "SQLAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   195
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5175
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblLiscence 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Liscenced"
      Height          =   525
      Left            =   1485
      TabIndex        =   6
      Top             =   1500
      Width           =   4080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   140.858
      X2              =   5239.909
      Y1              =   1490.87
      Y2              =   1490.87
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright © 2001 - 2003 Sameeriya Soft"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1485
      TabIndex        =   1
      Top             =   825
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1470
      TabIndex        =   3
      Top             =   195
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   140.858
      X2              =   5225.823
      Y1              =   1501.224
      Y2              =   1490.87
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1485
      TabIndex        =   4
      Top             =   495
      Width           =   3180
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning"
      ForeColor       =   &H00000000&
      Height          =   1320
      Left            =   255
      TabIndex        =   2
      Top             =   2325
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   About
'Description    :   About this Application
'Developed By   :   Sameer C T (Curtsey : Microsoft)
'Started On     :   2001 November 27
'Last Modified  :   2003 April 29
'-------------------------------------------------------------------------------------------

Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

''For that blending effect while loading
'Const AW_HOR_NEGATIVE = &H2 'Animates the window from right to left. This flag can be used with roll or slide animation.
'Const AW_VER_POSITIVE = &H4 'Animates the window from top to bottom. This flag can be used with roll or slide animation.
'Const AW_VER_NEGATIVE = &H8 'Animates the window from bottom to top. This flag can be used with roll or slide animation.
'Const AW_CENTER = &H10 'Makes the window appear to collapse inward if AW_HIDE is used or expand outward if the AW_HIDE is not used.
'Const AW_HIDE = &H10000 'Hides the window. By default, the window is shown.
'Const AW_ACTIVATE = &H20000 'Activates the window.
'Const AW_SLIDE = &H40000 'Uses slide animation. By default, roll animation is used.
'Const AW_BLEND = &H80000 'Uses a fade effect. This flag can be used only if hwnd is a top-level window.
'Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean


Private intKeySum As Integer
'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Activate()
    Timer1.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    intKeySum = intKeySum + KeyAscii
    If intKeySum = 379 Then '(ie:asc("A")+asc("B")+asc("O")+asc("U")+asc("T"))
        Label1.Caption = Chr(66) + Chr(89) + ":" + Chr(83) + Chr(65) + Chr(77) + Chr(69) + Chr(69) + Chr(82) + " " + Chr(67) + " " + Chr(84)
        Label1.Visible = True
        Timer1.Enabled = True
        intKeySum = 0
    End If
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmLogin.Icon
    KeyPreview = True
    Label1.Visible = False
    Timer1.Interval = 50
    Me.Caption = "About SQL Analyzer "
    lblTitle.Caption = "SQL Analyzer™ Professional Edition"
    lblVersion.Caption = "Version 5.00.01"
    lblCopyright.Caption = "Copyright © 2001 - 2003 Sameeriya Soft"
    lblLiscence.Caption = "You!" & vbCrLf & "Release Number : 101020031637"
    'lblLiscence.Caption = "Soft Systems Limited, Cochin" & vbCrLf & "Release Number : 101020031637"
    lblDisclaimer.Caption = "Warning : This computer program is protected by copyright law and international" & _
    " treaties. Unauthorized reproduction or distribution of this program, or any portion of it, " & _
    "may result in severe civil and criminal penalties, and will be prosecuted to the maximum " & _
    "extent possible under law."
    
    'For that blending effect while loading
    'Me.AutoRedraw = True
    'AnimateWindow Me.hwnd, 3000, AW_BLEND
    'Me.AutoRedraw = False
    'Me.Refresh
End Sub

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    intKeySum = 0
End Sub

Private Sub Timer1_Timer()
    Label1.Move (Label1.Left - 30)
    If Label1.Left < -2200 Then
        Label1.Left = 5300
    End If
End Sub
