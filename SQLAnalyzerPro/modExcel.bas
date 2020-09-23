Attribute VB_Name = "modExcel"
'-------------------------------------------------------------------------------------------
'Module         :   Excel Reports
'Description    :   Generates Excel Reports
'Developed By   :   Sameer C T
'Started On     :   2003 April 12
'Last Modified  :   2003 April 12

'Reference      :   Microsoft Excel 9.0 Object Library (EXCEL9.OLB)

'Usage:
'Create following variables in the form
'Private exlApln As Excel.Application
'Private exlBook As Excel.Workbook
'Private exlSheet As Excel.Worksheet
'-------------------------------------------------------------------------------------------

Option Explicit

Private intCount As Integer

Public Function gfnExlInitialise(objApln As Excel.Application, objBook As Excel.Workbook, objSheet As Excel.Worksheet) As Boolean
'   Create an Excel Application with a Workbook and Sheet; Returns True if Succees.
'   Usage:
'   If gfnExlInitialise(exlApln, exlBook, exlSheet) = False Then
'      Exit Sub
'   End If
    
   On Error GoTo ErrorTrap
   
   Set objSheet = Nothing
   Set objBook = Nothing
   Set objApln = Nothing
   
   Set objApln = CreateObject("Excel.Application")
   Set objBook = objApln.Workbooks.Add
   Set objSheet = objBook.Sheets(1)
   gfnExlInitialise = True
   Exit Function
ErrorTrap:
   MsgBox "Error encountered while creating the Excel sheet." & vbCrLf & _
      "Make sure that Microsoft Excel is installed", vbInformation
   gfnExlInitialise = False
End Function

Public Sub gprExlDisableMenus(objApln As Excel.Application, blnPrint As Boolean, blnSave As Boolean)
'   Disable or Enable Print and Save options in menu and tool bar
'   Usage: Call gprExlDisableMenus(exlApln, True, True) to disable both
   
   On Error Resume Next
   
   'Menu
   With objApln.CommandBars("File")
      For intCount = 1 To .Controls.Count
         If ((InStr(UCase(.Controls(intCount).Caption), "PRINT") <> 0) Or _
              (InStr(UCase(.Controls(intCount).Caption), "PAGE") <> 0)) Then
             .Controls(intCount).Enabled = Not blnPrint
         End If
         If ((InStr(UCase(.Controls(intCount).Caption), "SAVE") <> 0)) Then
             .Controls(intCount).Enabled = Not blnSave
         End If
      Next intCount
   End With
   
   'Toolbar
   With objApln.CommandBars("Standard")
      For intCount = 1 To .Controls.Count
         If ((InStr(UCase(.Controls(intCount).Caption), "PRINT") <> 0)) Then
              .Controls(intCount).Enabled = Not blnPrint
         End If
         If ((InStr(UCase(.Controls(intCount).Caption), "SAVE") <> 0)) Then
             .Controls(intCount).Enabled = Not blnSave
         End If
      Next intCount
   End With
End Sub

Public Sub gprExlDisplayToolbars(objApln As Excel.Application, blnDisplay As Boolean)
'  Sets the display of toolbars
   With objApln.Application
      .CommandBars("Standard").Enabled = blnDisplay
      .CommandBars("Formatting").Enabled = blnDisplay
      .CommandBars("Control Toolbox").Enabled = blnDisplay
      .CommandBars("Drawing").Enabled = blnDisplay
      .DisplayFormulaBar = blnDisplay
      .DisplayStatusBar = blnDisplay
   End With
End Sub

Public Sub gprExlDisplaySettings(objApln As Excel.Application, blnBookTab As Boolean, blnGrid As Boolean, blnZeros As Boolean)
'  Sets the display of work book tabs, grid lines and zeros
   With objApln.ActiveWindow
      .DisplayWorkbookTabs = blnBookTab
      .DisplayGridlines = blnGrid
      .DisplayZeros = blnZeros
   End With
End Sub

Public Sub gprExlSetCaptions(objApln As Excel.Application, strApln As String, strWindow As String)
'  Sets the captions
   With objApln
      .Caption = strApln
      .ActiveWindow.Caption = strWindow
   End With
End Sub

Public Sub gprExlSetTitles(objSheet As Excel.Worksheet, strMain As String, intMainSize As Integer, strSub As String, intSubSize As Integer)
'  Sets the Titles on first two lines
   With objSheet
      .Cells(1, 1) = strMain
      .Cells(1, 1).Font.Bold = True
      .Cells(1, 1).Font.Size = intMainSize
      .Cells(2, 1) = strSub
      .Cells(2, 1).Font.Bold = True
      .Cells(2, 1).Font.Size = intSubSize
   End With
End Sub

Public Sub gprExlPageSetup(objSheet As Excel.Worksheet, intTitleRows As Integer, Optional strPassword As String)
'  Page setup
   Dim strTitleRows As String
   
   strTitleRows = "$1:$" & intTitleRows
   With objSheet.PageSetup
      .PrintTitleRows = strTitleRows
      .LeftFooter = "Page No: &P / &N"
      .PaperSize = xlPaperA4
      .Orientation = xlPortrait
      .TopMargin = 50
      .LeftMargin = 50
      .RightMargin = 10
      .BottomMargin = 50
      .Zoom = 70
      '.CenterHorizontally = True
   End With
   objSheet.DisplayPageBreaks = False
   If Not IsMissing(strPassword) Then objSheet.Protect (strPassword)
End Sub

Public Sub gprExlShow(objApln As Excel.Application, objBook As Excel.Workbook, objSheet As Excel.Worksheet, blnPrint As Boolean)
'  Shows the Excel sheet or prints it
   objBook.Saved = True
   With objApln
      If blnPrint = False Then 'Show
         .WindowState = xlMaximized
         .ActiveWindow.WindowState = xlMaximized
         .Visible = True
      Else 'Print
         objSheet.PrintOut
         .Quit
      End If
   End With
End Sub
   
Public Sub gprExlDispose(objApln As Excel.Application, objBook As Excel.Workbook, objSheet As Excel.Worksheet)
'  Destroys an Excel Application with a Workbook and Sheet
   On Error Resume Next
   Set objSheet = Nothing
   Set objBook = Nothing
   Set objApln = Nothing
End Sub

