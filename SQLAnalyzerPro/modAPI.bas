Attribute VB_Name = "modAPI"
'-------------------------------------------------------------------------------------------
'Module         :   API Code Collection
'Description    :   Useful API functions in a single module file
'                   (can be plugged into any project)
'Developed By   :   Sameer C T
'Started On     :   2002 January 25
'Last Modified  :   2002 January 30
'-------------------------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------------------------
'Declarations
'-------------------------------------------------------------------------------------------

'API Declarations for DSN Population
Public Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Public Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1

'API Declaration for List Search
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
  lParam As Any) As Long
Const LB_FINDSTRING = &H18F

'-------------------------------------------------------------------------------------------
'Procedures
'-------------------------------------------------------------------------------------------

Public Sub PopulateDSNsDrivers(Optional DSNComboName As ComboBox, Optional DriverComboName As ComboBox)
    '(c)Sameer C T on 2001 December
    Dim intCount As Integer
    Dim strDSNSpc As String * 1024
    Dim strDRVSpc As String * 1024
    Dim strDSN As String
    Dim strDRV As String
    Dim intDSNLn As Integer
    Dim intDRVLn As Integer
    Dim lngHenv As Long

    On Error Resume Next

    'Getting the DSNs using API and adding them to combos
    If SQLAllocEnv(lngHenv) <> -1 Then
        Do Until intCount <> SQL_SUCCESS
            strDSNSpc = Space$(1024)
            strDRVSpc = Space$(1024)
            intCount = SQLDataSources(lngHenv, SQL_FETCH_NEXT, strDSNSpc, 1024, intDSNLn, strDRVSpc, 1024, intDRVLn)
            strDSN = Left$(strDSNSpc, intDSNLn)
            strDRV = Left$(strDRVSpc, intDRVLn)
                
            If strDSN <> Space(intDSNLn) Then
                'If the DSN Combo is passed
                If Not DSNComboName Is Nothing Then
                    'Adding only SQL Server DSNs
                    If strDRV = "SQL Server" Then DSNComboName.AddItem strDSN
                End If
                'If the Driver Combo is passed
                If Not DriverComboName Is Nothing Then
                    DriverComboName.AddItem strDRV
                End If
            End If
        Loop
    End If
       
End Sub
  
Public Sub PopulateSQLServers(ComboName As ComboBox)
    'Populates the available SQL Servers
    'Reference : Microsoft SQLDMO Object Library
    
    Dim appSQL As SQLDMO.Application
    Dim objNames As SQLDMO.NameList
    Dim intCount As Integer
        
    Set appSQL = New SQLDMO.Application
    
    'Sets the available servers into the namelist
    Set objNames = appSQL.ListAvailableSQLServers
        
    'Adding them to combo
    For intCount = 1 To objNames.Count
        ComboName.AddItem objNames.Item(intCount)
    Next intCount
        
    Set objNames = Nothing
    Set appSQL = Nothing
End Sub
Public Sub ListSearch(ByVal ListName As ListBox, SearchText As String)
    'Searches for the passed text in a list box(Uses API call)
    With ListName
        .ListIndex = SendMessage(.hwnd, LB_FINDSTRING, -1, ByVal SearchText)
        If Not .ListIndex = -1 Then .TopIndex = .ListIndex
    End With
End Sub
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------
