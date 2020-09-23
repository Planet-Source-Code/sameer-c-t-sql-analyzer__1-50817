VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServerDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Directory"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   315
      Left            =   2550
      TabIndex        =   5
      ToolTipText     =   "Click to start the operration"
      Top             =   4155
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3585
      TabIndex        =   6
      ToolTipText     =   "Click to close this window"
      Top             =   4155
      Width           =   975
   End
   Begin VB.TextBox txtFullPath 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Full path of file used for the operation; is editable"
      Top             =   3705
      Width           =   4470
   End
   Begin VB.TextBox txtFileName 
      Height          =   330
      Left            =   1425
      TabIndex        =   3
      Top             =   3225
      Width           =   3165
   End
   Begin MSComctlLib.TreeView tvwDir 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   315
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   4921
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "Folder/File &Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3300
      Width           =   1260
   End
   Begin VB.Label lblDir 
      AutoSize        =   -1  'True
      Caption         =   "Select &Folder/File:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   1290
   End
End
Attribute VB_Name = "frmServerDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   5.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   Server Directory
'Description    :   Displays the Server Directory Structure for Backup and Restore
'Developed By   :   Sameer C T (Courtsey : Manoj Dominic)
'Started On     :   2003 September 25
'Last Modified  :   2003 September 25
'-------------------------------------------------------------------------------------------

Option Explicit

Private ndChild As Node, ndChild1 As Node
Private ndKid As Node

Public mblnFilesNeeded As Boolean
Public mstrBakFileName As String

'-------------------------------------------------------------------------------------------
'Event Procedures
'-------------------------------------------------------------------------------------------

Private Sub Form_Load()
    Me.Icon = frmLogin.Icon
    tvwDir.ImageList = frmMain.imlSQL
    mblnCancelDirectory = False
    txtFileName = mstrBakFileName
    Call mprPopulateDirectory
End Sub

Private Sub tvwDir_Expand(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    If Node.FullPath <> tvwDir.Nodes(Node.Index).Root.FullPath Then
        tvwDir.Nodes.Remove (Node.Child.Index)
    End If
    Call mprExpandDir(Node)
End Sub

Private Sub tvwDir_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    Call mprExpandDir(Node)
    If mblnFilesNeeded = True Then
        txtFileName.Text = tvwDir.Nodes.Item(tvwDir.SelectedItem.Index).Text
        txtFullPath.Text = tvwDir.Nodes.Item(tvwDir.SelectedItem.Index).FullPath
    Else
        txtFullPath.Text = tvwDir.Nodes.Item(tvwDir.SelectedItem.Index).FullPath & "\" & Trim(txtFileName.Text)
    End If
End Sub

Private Sub txtFileName_Change()
    On Error Resume Next
    If mblnFilesNeeded = False Then
        txtFullPath.Text = tvwDir.Nodes.Item(tvwDir.SelectedItem.Index).FullPath & "\" & Trim(txtFileName.Text)
    End If
End Sub

Private Sub cmdStart_Click()
    If Trim(txtFullPath.Text) = "" Then
        MsgBox "Select/Enter the File Name", vbInformation
    End If
    
    mblnCancelDirectory = False
    mstrBakFileName = Trim(txtFullPath.Text)
    mconGeneral.Execute "Use " & mstrDatabase
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mblnCancelDirectory = True
    mconGeneral.Execute "Use " & mstrDatabase
    Unload Me
End Sub
'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------
Private Sub mprPopulateDirectory()
    Dim rsDir As New ADODB.Recordset
    Dim ndRoot As Node, ndTemp As Node
    
    'Set ndRoot = tvwDir.Nodes.Add(, , "Root", "Root")
    'ndRoot.Tag = 1
    mconGeneral.Execute "use Master"
    Set rsDir = mconGeneral.Execute("xp_availablemedia 2")
    With rsDir
        While Not .EOF
           Set ndTemp = tvwDir.Nodes.Add(, , !Name, !Name, "Drive")
            ndTemp.Tag = 1
            DoEvents
        .MoveNext
        Wend
    End With
    rsDir.Close
End Sub

Private Sub mprExpandDir(ParentDir As Node)
    Dim rsSubDir As New ADODB.Recordset, rsKid As New ADODB.Recordset
    
    If ParentDir.Children = 0 Then
        mconGeneral.Execute "Use Master"
        Set rsSubDir = mconGeneral.Execute("master.dbo.xp_dirtree N'" & Trim(ParentDir.FullPath) & "', " & IIf(ParentDir.Tag = "", 1, ParentDir.Tag) & " ,1")
            While Not rsSubDir.EOF
                If mblnFilesNeeded = True Then
                    If CBool(rsSubDir!File) = False Then 'folder
                        Set ndChild = tvwDir.Nodes.Add(ParentDir, tvwChild, , rsSubDir!SubDirectory, "ClosedFolder", "OpenFolder")
                        Set ndChild1 = tvwDir.Nodes.Add(ndChild, tvwChild, , "", "ClosedFolder", "OpenFolder")
                    Else 'its a file
                        If Right(rsSubDir!SubDirectory, 3) = "bak" Then 'If the extension is bak
                            Set ndChild = tvwDir.Nodes.Add(ParentDir, tvwChild, , rsSubDir!SubDirectory, "BakFile")
                        Else
                            Set ndChild = tvwDir.Nodes.Add(ParentDir, tvwChild, , rsSubDir!SubDirectory, "File")
                        End If
                    End If
                ElseIf CBool(rsSubDir!File) = False And mblnFilesNeeded = False Then  'its a folder
                    Set ndChild = tvwDir.Nodes.Add(ParentDir, tvwChild, , rsSubDir!SubDirectory, "ClosedFolder", "OpenFolder")
                    Set ndChild1 = tvwDir.Nodes.Add(ndChild, tvwChild, , "", "ClosedFolder", "OpenFolder")
                End If
                ndChild.Tag = rsSubDir!Depth
                
                If Not ParentDir.Selected Then
                    ParentDir.Image = "OpenFolder"
                End If
                rsSubDir.MoveNext
            Wend
            Set rsSubDir = Nothing
    End If
    ndChild.Expanded = False
End Sub
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
