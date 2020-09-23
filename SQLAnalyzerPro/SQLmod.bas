Attribute VB_Name = "modSQL"
'-------------------------------------------------------------------------------------------
'Project        :   SQL Analyzer
'Version        :   4.0
'Description    :   Easy to use tool for SQL Server Developers
'Module         :   SQL General Module
'Description    :   Public Variables and Procedures for this Application
'Developed By   :   Sameer C T
'Started On     :   2001 November 27
'Last Modified  :   2002 January 29
'-------------------------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------------------------
'Declarations
'-------------------------------------------------------------------------------------------

'API for Progress bar combination with status bar
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Const WM_USER = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)
Public lngRect As RECT
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'For delaying the process, now not used
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'API Declaration for List Search
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
  lParam As Any) As Long
Const LB_FINDSTRING = &H18F
  
Public mblnIsLogged As Boolean
Public mstrConString As String, mstrServer As String, mstrDatabase As String
Public mstrUserName As String, mstrPassword As String, mstrDSN As String
Public mblnConnected As Boolean



Public mconGeneral As New ADODB.Connection

Public objSQLServer As SQLDMO.SQLServer
Public mblnCancelDirectory As Boolean

Public Const FILETYPE_SQL = 1
Public Const FILETYPE_TXT = 2

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------
Sub Main()
    Dim fLogin As New frmLogin
    
    frmSplash.Show
    frmSplash.Refresh
    
    fLogin.Show vbModal
    If Not fLogin.OK Then
        'Login Failed so exit app
        End
    End If
    Unload fLogin

    Load frmMain
    Load frmTables
    
    Unload frmSplash

    frmMain.Show
    frmTables.Show
End Sub

Public Sub ShowStatusMsg(Messsage As String, IsBusy As Boolean)
    'Displays a message on status bar and sets the mouse pointer accordingly
    frmMain.stbSQL.Panels(1).Text = Messsage
    If IsBusy Then
        Screen.MousePointer = vbHourglass
    Else
        Screen.MousePointer = vbNormal
    End If
End Sub

Public Sub PopulateDataList(ListViewName As ListView, ObjectName As String, IsCompleteData As Boolean)
    'Populates the data in the table/view to the listview
    
    Dim rsTableData As ADODB.Recordset
    Dim strSQL As String
    Dim objList As ListItem
    Dim intCounter As Integer
    Dim intRecCount As Integer
    Dim blnTopF As Boolean
    Dim i As Integer
    Dim j As Integer
    
    'Populating the Column Names as List Headers
    Set rsTableData = New ADODB.Recordset
    strSQL = "Select Name From SysColumns Where Id=Object_Id(N'" & ObjectName & "') Order By ColId"
    rsTableData.Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
    With ListViewName
        .HideColumnHeaders = False
        .View = lvwReport
        .ColumnHeaders.Clear
    End With
    intCounter = 1
    While Not rsTableData.EOF
        ListViewName.ColumnHeaders.Add intCounter, , rsTableData("Name")
        intCounter = intCounter + 1
        rsTableData.MoveNext
    Wend
    rsTableData.Close
    
    'Populating the data as List Items
    If IsCompleteData = True Then 'Select all data
        strSQL = "Select * From " & ObjectName & ""
    Else 'Pick only 10 rows
        strSQL = "Select Top 10 * From " & ObjectName & ""
    End If
    intCounter = intCounter - 2
    ListViewName.ListItems.Clear
    
    rsTableData.Open strSQL, mconGeneral, adOpenKeyset
    intRecCount = rsTableData.RecordCount
    
    blnTopF = False
    If intRecCount > 500 Then
        If MsgBox("This contains " & Str(intRecCount) & " rows.." & vbCrLf & _
                "Do you want to Display them all?" & vbCrLf & _
                "[Click No to display only 500 rows]", vbQuestion + vbYesNo) = vbNo Then
            blnTopF = True
            intRecCount = 500
        End If
    End If
    
    'Call SetProgressBar(True, intRecCount + 1)
    'frmMain.prbSQL.Value = 0
    j = 0
    Do While Not rsTableData.EOF
        Set objList = ListViewName.ListItems.Add(, , IIf(IsNull(rsTableData(0)) = True, "NULL", rsTableData(0)))
        For i = 1 To intCounter
            If IsNull(rsTableData(i)) = True Then
                objList.ListSubItems.Add , , "NULL"
            Else
                If rsTableData(i) = True Then 'For the bit fields, it should display 1 instead of -1, if its true
                    objList.ListSubItems.Add , , "1"
                Else
                    objList.ListSubItems.Add , , rsTableData(i)
                End If
            End If
        Next i
        rsTableData.MoveNext
        j = j + 1
        'frmMain.prbSQL.Value = frmMain.prbSQL.Value + 1
        If blnTopF = True Then
            If j = 500 Then Exit Do
            'If frmMain.prbSQL.Value = 500 Then Exit Do
        End If
    Loop
    'Call SetProgressBar(False, 1)
    rsTableData.Close
    Set rsTableData = Nothing
End Sub

Public Sub PopulateObjectText(RichTextBoxName As RichTextBox, ObjectName As String)
    'Populates the Rich Text Box with the passed object's text
    Dim rsText As ADODB.Recordset
    Dim strSQL As String
    Dim strText As String
    
    RichTextBoxName.Text = ""
    strSQL = "SP_HelpText '" & ObjectName & "'"
    Set rsText = New ADODB.Recordset
    rsText.Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
    While Not rsText.EOF
        strText = Trim(strText) & Trim(rsText.Fields(0).Value)
        rsText.MoveNext
    Wend
    RichTextBoxName.Text = Trim(strText)
    
    rsText.Close
    Set rsText = Nothing
End Sub

Public Sub SetProgressBar(Visibility As Boolean, MaxValue As Integer)
    'Setting progress bar aligned to the status bar
    SendMessage frmMain.stbSQL.hwnd, SB_GETRECT, 0, lngRect
    SetParent frmMain.prbSQL.hwnd, frmMain.stbSQL.hwnd
    frmMain.prbSQL.Move lngRect.Left * Screen.TwipsPerPixelX, lngRect.Top * Screen.TwipsPerPixelY, (lngRect.Right - lngRect.Left) * Screen.TwipsPerPixelX, 300

    frmMain.prbSQL.Visible = Visibility
    frmMain.prbSQL.Min = 0
    frmMain.prbSQL.Max = MaxValue
End Sub

Public Function gfnGetQualifiedName(ObjectName As String, Optional SpOrFn As Boolean) As String
    'Returns the full(qualified) table name from INFORMATION_SCHEMA.TABLES
    Dim rsCom As ADODB.Recordset
    Dim strSQL As String
    Dim intCounter As Integer
    Dim i As Variant
    
   
    Set rsCom = New ADODB.Recordset
    'Constructing SQL Query
    If SpOrFn = True Then
        strSQL = "SELECT ROUTINE_CATALOG AS DBName, ROUTINE_SCHEMA As OwnerName FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_NAME = '" & ObjectName & "'"
    Else
        strSQL = "SELECT TABLE_CATALOG AS DBName, TABLE_SCHEMA As OwnerName FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & ObjectName & "'"
    End If
    
    With rsCom
        .Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
        If Not .EOF Then
            gfnGetQualifiedName = "[" & .Fields("DBName") & "].[" & .Fields("OwnerName") & "].[" & ObjectName & "]"
        End If
        .Close
    End With

    Set rsCom = Nothing
End Function

Public Function gfnGetOwnerName(ObjectName As String, Optional SpOrFn As Boolean) As String
    'Returns the full(qualified) table name from INFORMATION_SCHEMA.TABLES
    Dim rsCom As ADODB.Recordset
    Dim strSQL As String
    Dim intCounter As Integer
    Dim i As Variant
    
   
    Set rsCom = New ADODB.Recordset
    'Constructing SQL Query
    If SpOrFn = True Then
        strSQL = "SELECT ROUTINE_SCHEMA As OwnerName FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_NAME = '" & ObjectName & "'"
    Else
        strSQL = "SELECT TABLE_SCHEMA As OwnerName FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & ObjectName & "'"
    End If
    
    With rsCom
        .Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
        If Not .EOF Then
            gfnGetOwnerName = Trim(.Fields("OwnerName"))
        End If
        .Close
    End With

    Set rsCom = Nothing
End Function
Public Function GetTextFromSysCom(ObjectName As String) As String
    'Returns the source text from syscomments table
    Dim rsCom As ADODB.Recordset
    Dim strSQL As String
    Dim intCounter As Integer
    Dim i As Variant
    
   
    Set rsCom = New ADODB.Recordset
    'Constructing SQL Query
    strSQL = "Select Text From SysComments where Id = Object_Id(N'[" & ObjectName & "]')"
    With rsCom
        .Open strSQL, mconGeneral, adOpenForwardOnly, adLockReadOnly
        If Not .EOF Then
            GetTextFromSysCom = IIf(IsNull(.Fields("Text")), "", .Fields("Text"))
        End If
        .Close
    End With

    Set rsCom = Nothing
End Function
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
