Attribute VB_Name = "modGeneral"
'-------------------------------------------------------------------------------------------
'Module         :   General Code Collection
'Description    :   Useful General functions in a single module file
'                    (can be plugged into any project)
'Developed By   :   Sameer C T
'Started On     :   2002 January 25
'Last Modified  :   2002 January 30
'-------------------------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------------------------
'Declarations
'-------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------
'Common Procedures
'-------------------------------------------------------------------------------------------
Public Sub SortListView(ByVal ListViewName As ListView, ByVal ColumnHeaderName As MSComctlLib.ColumnHeader)
    'Sorts the content of a list view according to the Clicked Column
    'Also changes the Column Header Icon
    
    Dim objCol As ColumnHeader
    
    'Clearing all the Icons of Column Headers
    For Each objCol In ListViewName.ColumnHeaders
        objCol.Icon = 0
    Next
    
    With ListViewName
        If .SortKey = ColumnHeaderName.SubItemIndex Then 'If it was clicked earlier
            If .SortOrder = lvwAscending Then    'Reverse the Sort Order
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else    'Otherwise sort it using index
            .SortKey = ColumnHeaderName.SubItemIndex
        End If
    End With
End Sub

Public Function gfnEncrypt(Text As String, EncryptKey As Integer)

    'Dim Key As Integer
    'Key = (Int(Sqr(Len(SecretPassword) * 95)) + 21)

    Dim Temp As String, RR As Integer
    For RR = 1 To Len(Text)
        Temp$ = Temp$ + Chr$(Asc(Mid(Text, RR, 1)) Xor EncryptKey)
    Next RR
    gfnEncrypt = Temp
End Function
'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------

