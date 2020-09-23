Attribute VB_Name = "modKeyColor"
'-------------------------------------------------------------------------------------------
'Module         :   Key Color
'Description    :   Coloring of keywords in a rich text box
'Developed By   :   Sameer C T
'Started On     :   2003 April 03
'Last Modified  :   2003 April 11
'Usage:
'1.Plug this module into the required project
'2.Edit the value of CONST_TOTALKEYWORDS to required number of keywords
'3.Edit the procedure gprSetKeyWords to add the required key words and their colors
'4.Call the procedure gprSetKeyWords while loading the application
'5.call the procedure gprColorKeyWords wherever required
'Bug noted:
'1.When called in the change event of RTB, the content flashes
'-------------------------------------------------------------------------------------------

Option Explicit

'User defined data type for storing the Keywords and their colors
Public Type typColorKeyWords
    KeyWord As String
    ColorCode As ColorConstants
End Type

'Total Number of keywords
Private Const CONST_TOTALKEYWORDS = 106

'Array to store the keywords and associated colors
Public mstrKeyWords(CONST_TOTALKEYWORDS) As typColorKeyWords

Public Sub gprSetKeyWords()
    'Stores the key words that is to be coloured and the associated color, into the array
    'A space should be given at the beginning and end of words to avoid coloring parts of a word
    'The following variables are used instead of direct numbers to give flexibility while adding more keywords
    
    Dim intCount As Integer
    Dim intMark As Integer
    Dim intP As Integer
    
    
    intP = 0
    intMark = intP
    mstrKeyWords(intP).KeyWord = "ADD"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "ALL"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "ALTER"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "AND"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "ANY"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "AS"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "BEGIN"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "CLOSE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "COLLATE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "COLUMN"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "COMMIT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "CONSTRAINT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "CREATE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "CURSOR"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "DEALLOCATE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "DECLARE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "DEFAULT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "DELETE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "DISTINCT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "DISTRIBUTED"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "DOUBLE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "DROP"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "ELSE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "END"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "ESCAPE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "EXCEPT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "EXEC"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "EXECUTE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "EXISTS"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "EXIT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "FETCH"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "FILLFACTOR"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "FOREIGN"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "FROM"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "FUNCTION"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "GROUP"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "HAVING "
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "ID"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "IF"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "IN"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "INDEX"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "INNER"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "INSERT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "INTERSECT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "INTO"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "IS"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "JOIN"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "KEY"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "KILL"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "LEFT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "LIKE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "NOCHECK"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "NONCLUSTERED"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "NOT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "NULL"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "OF"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "OFF"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "ON"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "OPEN"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "OPTION"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "OR"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "ORDER"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "OUTER"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "OUTPUT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "OVER"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "PERCENT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "PRECISION"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "PRIMARY"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "PRINT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "PROC"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "PROCEDURE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "PUBLIC"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "RAISERROR"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "READ"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "REFERENCES"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "RESTORE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "RETURN"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "RIGHT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "ROLLBACK"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "SELECT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "SET"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "SOME"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "TABLE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "THEN"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "TO"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "TOP"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "TRAN"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "TRANSACTION"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "TRIGGER"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "TRUNCATE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "UNION"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "UNIQUE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "UPDATE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "USE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "VALUES"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "VIEW"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "WHEN"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "WHERE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "WHILE"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "WITH"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "WRITETEXT"
    
    For intCount = intMark To intP
        mstrKeyWords(intCount).ColorCode = vbBlue
    Next intCount
    
    intP = intP + 1
    intMark = intP
    mstrKeyWords(intP).KeyWord = "sysobjects"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "syscolumns"
    
    For intCount = intMark To intP
        mstrKeyWords(intCount).ColorCode = vbGreen
    Next intCount
    
    intP = intP + 1
    intMark = intP
    mstrKeyWords(intP).KeyWord = "COUNT"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "OBJECT_ID"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "OBJECTPROPERTY"
    intP = intP + 1
    mstrKeyWords(intP).KeyWord = "SUM"
    
    For intCount = intMark To intP
        mstrKeyWords(intCount).ColorCode = vbMagenta
    Next intCount
End Sub

Public Sub gprColorKeyWords(RichTextBoxName As RichTextBox)
    'Changes the text color of key words
    Dim intPos As Long
    Dim intCount As Integer
    
    'First making all text to black
    RichTextBoxName.SelStart = 0
    RichTextBoxName.SelLength = Len(RichTextBoxName.Text)
    RichTextBoxName.SelColor = vbBlack
    
    'Looping through the total number of set keywords
    For intCount = 0 To CONST_TOTALKEYWORDS - 1
    With RichTextBoxName
        'Check for the existance of that keyword
        intPos = .Find(mstrKeyWords(intCount).KeyWord, 0, Len(.Text), rtfWholeWord)
        If intPos <> -1 Then
            'If it exists, it is colored
            Call mprColorAgain(RichTextBoxName, intPos, mstrKeyWords(intCount).KeyWord, mstrKeyWords(intCount).ColorCode)
        End If
    End With
    Next intCount
End Sub

Private Sub mprColorAgain(RichTextBoxName As RichTextBox, Position As Long, KeyWord As String, KeyColor As ColorConstants)
    'Recursive procedure, that will keep on coloring a particular keyword,
    'after locating all its occurrences
    Dim intP As Long
    
    'Sets the color for the keyword
    With RichTextBoxName
        .SelStart = Position
        .SelLength = Len(KeyWord)
        .SelColor = KeyColor
    End With
    
    'Check for its occurrence again
    intP = RichTextBoxName.Find(KeyWord, Position + Len(KeyWord), Len(RichTextBoxName.Text), rtfWholeWord)
    
    'If the keyword is present again, call this procedure recursively to color it
    If intP <> -1 Then
        Call mprColorAgain(RichTextBoxName, intP, KeyWord, KeyColor)
    End If
End Sub

'-------------------------------------------------------------------------------------------
'The End
'-------------------------------------------------------------------------------------------

