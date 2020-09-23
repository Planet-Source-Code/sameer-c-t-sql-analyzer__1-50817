VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQL Analyzer : Connection"
   ClientHeight    =   3000
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   3465
   Icon            =   "SQLLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSQLCancel 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   315
      Left            =   2385
      TabIndex        =   15
      Top             =   2580
      Width           =   975
   End
   Begin VB.CommandButton cmdCONConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   315
      Left            =   1335
      TabIndex        =   14
      ToolTipText     =   "Click to connect to the data source"
      Top             =   2580
      Width           =   975
   End
   Begin VB.Frame fraCON 
      Height          =   2460
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.CheckBox chkCONRemember 
         Caption         =   "&Remember Password"
         Height          =   195
         Left            =   1095
         TabIndex        =   5
         Top             =   1020
         Width           =   2010
      End
      Begin VB.TextBox txtCONUserName 
         Height          =   315
         Left            =   1095
         TabIndex        =   2
         ToolTipText     =   "Enter User Name"
         Top             =   240
         Width           =   2000
      End
      Begin VB.TextBox txtCONPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "Enter Password"
         Top             =   600
         Width           =   2000
      End
      Begin VB.OptionButton optCONMode 
         Caption         =   "Specify &Server"
         Height          =   195
         Index           =   1
         Left            =   1575
         TabIndex        =   7
         ToolTipText     =   "Click to specify a server"
         Top             =   1305
         Width           =   1455
      End
      Begin VB.OptionButton optCONMode 
         Caption         =   "Use &DSN"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         ToolTipText     =   "Click to use DSN"
         Top             =   1305
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame fraCONDSN 
         BorderStyle     =   0  'None
         Height          =   810
         Left            =   135
         TabIndex        =   17
         Top             =   1485
         Width           =   3015
         Begin VB.ComboBox cboCONDSN 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Select a DSN"
            Top             =   300
            Width           =   2000
         End
         Begin VB.Label lblDSN 
            AutoSize        =   -1  'True
            Caption         =   "DS&N:"
            Height          =   195
            Left            =   15
            TabIndex        =   8
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.Frame fraCONServer 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         TabIndex        =   16
         Top             =   1500
         Width           =   3015
         Begin VB.ComboBox cboCONServers 
            Height          =   315
            Left            =   975
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   90
            Width           =   1995
         End
         Begin VB.ComboBox cboCONDatabase 
            Height          =   315
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Database Name"
            Top             =   480
            Width           =   2000
         End
         Begin VB.Label lblCONDatabase 
            AutoSize        =   -1  'True
            Caption         =   "Data&base:"
            Height          =   195
            Left            =   30
            TabIndex        =   12
            Top             =   540
            Width           =   780
         End
         Begin VB.Label lblCONServerName 
            AutoSize        =   -1  'True
            Caption         =   "Se&rver:"
            Height          =   195
            Left            =   30
            TabIndex        =   10
            Top             =   150
            Width           =   555
         End
      End
      Begin VB.Label lblCONUserName 
         AutoSize        =   -1  'True
         Caption         =   "&User Name:"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblCONPassword 
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Login
'Description    :   Login window to connect to the database
'Developed By   :   Sameer C T
'Started On     :   2001 November 27
'Last Modified  :   2002 October 07
'-------------------------------------------------------------------------------------------

Option Explicit

Public OK As Boolean

'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Load()
    Dim strServer As String
    Dim strRegDSN As String
    Dim intCounter As Integer

    'Populating combos
    Call PopulateDSNsDrivers(cboCONDSN)
    Call PopulateSQLServers(cboCONServers)
    
    'Getting the default value of DSN from registry and setting it
    strRegDSN = GetSetting(App.Title, "UserInputs", "DSN", "")
    If strRegDSN <> "" Then
        For intCounter = 0 To cboCONDSN.ListCount
            If (cboCONDSN.List(intCounter) = strRegDSN) Then cboCONDSN.ListIndex = intCounter
        Next intCounter
    End If
    
    strServer = GetSetting(App.Title, "UserInputs", "Server", "(local)")
    If strServer <> "" Then
        For intCounter = 0 To cboCONServers.ListCount
            If (cboCONServers.List(intCounter) = strServer) Then cboCONServers.ListIndex = intCounter
        Next intCounter
    End If
    
    'Loading user inputs from registry as default
    Call GetFromRegistry
End Sub

Private Sub txtCONUserName_GotFocus()
    txtCONUserName.SelLength = Len(txtCONUserName.Text)
End Sub

Private Sub txtCONPassword_GotFocus()
    txtCONPassword.SelLength = Len(txtCONPassword.Text)
End Sub

Private Sub optCONMode_Click(Index As Integer)
    'Selection of mode (DSN/Specify Server)
    'Changes the visibility of frames according to selection
    Select Case Index
        Case 0  'DSN
            fraCONDSN.Visible = True
            fraCONServer.Visible = False
        Case 1  'Specify Server
            fraCONDSN.Visible = False
            fraCONServer.Visible = True
    End Select
End Sub

Private Sub cboCONDatabase_GotFocus()
    'Populates the Database combo
    Dim strDbase As String
    Dim intCounter As Integer
    
    On Error GoTo ErrorTrap
    
    Call GetInputs
    Call PopulateDatabases(cboCONDatabase)
    
    'Setting the default value of database name from registry
    strDbase = GetSetting(App.Title, "UserInputs", "Database", "Master")
    If strDbase <> "" Then
        For intCounter = 0 To cboCONDatabase.ListCount
            If (cboCONDatabase.List(intCounter) = strDbase) Then cboCONDatabase.ListIndex = intCounter
        Next intCounter
    End If
    
    Exit Sub
ErrorTrap:
    MsgBox "Error...Databases cannot be populated!" & vbCrLf & _
            "Details:" & Err.Description, vbOKOnly + vbCritical, App.Title
End Sub

Private Sub cmdCONConnect_Click()
    'Connects to the database and populates all objects
    Dim frmA As Form
    
    On Error GoTo ErrorTrap
    
    Screen.MousePointer = vbHourglass
    frmMain.stbSQL.Panels(2).Text = "Not Connected"
    
    'Call ShowStatusMsg("Validating inputs...", True)
    Call GetInputs
    If ValidateInputs = False Then Exit Sub
    
    'Branching according to the Connection mode(DSN/Server)
    If optCONMode(0) = True Then   'Use DSN
        'Constructing the Connection String
        mstrConString = "DSN=" & mstrDSN & ";uid=" & mstrUserName & ";"
        If mstrPassword <> "" Then mstrConString = mstrConString & "pwd=" & mstrPassword
    Else    'Specify Server
        'Constructing the Connection String
        mstrConString = "Provider=SQLOLEDB.1;User ID=" & mstrUserName & _
                        ";Pwd=" & mstrPassword & ";Initial Catalog=" & mstrDatabase & _
                        ";Data Source=" & mstrServer
    End If
    
    If mconGeneral.State = adStateOpen Then mconGeneral.Close
    mconGeneral.Open mstrConString
    
    'Setting the Statusbar info
    mstrServer = mconGeneral.Properties("Server Name")
    mstrDatabase = mconGeneral.Properties("Current Catalog")
    frmMain.stbSQL.Panels(2).Text = mstrServer
    frmMain.stbSQL.Panels(3).Text = mstrDatabase
    frmMain.stbSQL.Panels(4).Text = mstrUserName
    
    Set objSQLServer = New SQLDMO.SQLServer
    objSQLServer.Connect CStr(mstrServer), CStr(mstrUserName), CStr(mstrPassword)
    
    Call gprSetKeyWords
    OK = True
    Me.Hide
    
    If mblnIsLogged = True Then
        For Each frmA In Forms
            If frmA.Name <> "frmMain" Then
                Unload frmA
            End If
        Next
        frmTables.Show
    End If
    
    mblnIsLogged = True
    mblnConnected = True

    Screen.MousePointer = vbDefault
    Exit Sub
ErrorTrap:
    Screen.MousePointer = vbDefault
    MsgBox "Error...Problems encountered while connecting!" & vbCrLf & _
            "Details:" & Err.Description, vbOKOnly + vbCritical, App.Title
End Sub

Private Sub cmdSQLCancel_Click()
    If mblnIsLogged = False Then
        OK = False
        Me.Hide
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnIsLogged = False Then
        'End
    Else
        Call SetToRegistry
    End If
End Sub

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------

Private Sub GetFromRegistry()
    'Getting the saved userinputs(Connection Settings) from registry and setting them
    txtCONUserName.Text = GetSetting(App.Title, "UserInputs", "UserName", "sa")
    txtCONPassword.Text = gfnEncrypt(GetSetting(App.Title, "UserInputs", "Password", ""), 26)
    chkCONRemember.Value = GetSetting(App.Title, "UserInputs", "Remember", 0)
    optCONMode(0).Value = GetSetting(App.Title, "UserInputs", "ModeDSN", "TRUE")
    optCONMode(1).Value = GetSetting(App.Title, "UserInputs", "ModeServer", "FALSE")
End Sub

Private Sub SetToRegistry()
    'Saves the user inputs(Connection Settings) to registry
    SaveSetting App.Title, "UserInputs", "UserName", Trim(txtCONUserName.Text)
    If chkCONRemember.Value = 1 Then
        SaveSetting App.Title, "UserInputs", "Password", gfnEncrypt(Trim(txtCONPassword.Text), 26)
    Else
        SaveSetting App.Title, "UserInputs", "Password", ""
    End If
    SaveSetting App.Title, "UserInputs", "Remember", chkCONRemember.Value
    SaveSetting App.Title, "UserInputs", "DSN", Trim(cboCONDSN.Text)
    SaveSetting App.Title, "UserInputs", "Server", Trim(cboCONServers.Text)
    SaveSetting App.Title, "UserInputs", "Database", Trim(cboCONDatabase.Text)
    SaveSetting App.Title, "UserInputs", "ModeDSN", Str(optCONMode(0).Value)
    SaveSetting App.Title, "UserInputs", "ModeServer", Str(optCONMode(1).Value)
End Sub

Private Sub GetInputs()
    'Stores the user inputs into master variables
    mstrUserName = Trim(txtCONUserName.Text)
    mstrPassword = Trim(txtCONPassword.Text)
    mstrDSN = Trim(cboCONDSN.Text)
    mstrServer = Trim(cboCONServers.Text)
    mstrDatabase = Trim(cboCONDatabase.Text)
End Sub

Private Function ValidateInputs() As Boolean
    'Validates the user inputs
    
    ValidateInputs = False
    If mstrUserName = "" Then 'User Name
            MsgBox "User Name must be entered!", vbOKOnly + vbInformation, App.Title
            If txtCONUserName.Enabled = True Then txtCONUserName.SetFocus
            Exit Function
    End If
    
    If optCONMode(0) = True Then   'Use DSN
        If mstrDSN = "" Then
            MsgBox "DSN must be selected!", vbOKOnly + vbInformation, App.Title
            If cboCONDSN.Enabled = True Then cboCONDSN.SetFocus
            Exit Function
        End If
    Else    'Specify Server
        If mstrServer = "" Then
            MsgBox "Server name must be entered!", vbOKOnly + vbInformation, App.Title
            If cboCONServers.Enabled = True Then cboCONServers.SetFocus
            Exit Function
        End If
        
        If mstrDatabase = "" Then
            MsgBox "Database must be selected!", vbOKOnly + vbInformation, App.Title
            If cboCONDatabase.Enabled = True Then cboCONDatabase.SetFocus
            Exit Function
        End If
    End If
    ValidateInputs = True
End Function

Private Sub PopulateDatabases(ComboBoxName As ComboBox)
    'Populates the Database combo
    Dim conDbase As New ADODB.Connection
    Dim rsDbase As New ADODB.Recordset
    
    mstrDatabase = "Master"
    mstrConString = "Provider=SQLOLEDB.1;User ID=" & mstrUserName & _
                        ";Pwd=" & mstrPassword & ";Initial Catalog=" & mstrDatabase & _
                        ";Data Source=" & mstrServer
    conDbase.Open mstrConString
    Set rsDbase = conDbase.Execute("Select Name From SysDatabases")
    
    ComboBoxName.Clear
    While Not rsDbase.EOF
        ComboBoxName.AddItem rsDbase("Name")
        rsDbase.MoveNext
    Wend
    
    rsDbase.Close
    Set rsDbase = Nothing
    
    conDbase.Close
    Set conDbase = Nothing
End Sub

'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
