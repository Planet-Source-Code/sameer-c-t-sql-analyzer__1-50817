VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Private nFormHeight As Integer
Private nFormWidth As Integer
Private nNumOfControls As Integer
Private nTop() As Integer
Private nLeft() As Integer
Private nHeight() As Integer
Private nWidth() As Integer
Private nFontSize() As Integer
Private nRightMargin() As Integer
Private bFirstTime As Boolean

Public Sub Init(frm As Form, Optional nWindState As Variant)

    Dim i As Integer
    Dim bWinMax As Boolean

    bWinMax = Not IsMissing(nWindState)

    nFormHeight = frm.Height
    nFormWidth = frm.Width
    nNumOfControls = frm.Controls.Count - 1
    bFirstTime = True
    ReDim nTop(nNumOfControls)
    ReDim nLeft(nNumOfControls)
    ReDim nHeight(nNumOfControls)
    ReDim nWidth(nNumOfControls)
    ReDim nFontSize(nNumOfControls)

    ReDim nRightMargin(nNumOfControls)
    On Error Resume Next
    For i = 0 To nNumOfControls
        If TypeOf frm.Controls(i) Is Line Then
            nTop(i) = frm.Controls(i).Y1
            nLeft(i) = frm.Controls(i).X1
            nHeight(i) = frm.Controls(i).Y2
            nWidth(i) = frm.Controls(i).X2
        Else
            nTop(i) = frm.Controls(i).Top
            nLeft(i) = frm.Controls(i).Left
            nHeight(i) = frm.Controls(i).Height
            nWidth(i) = frm.Controls(i).Width
            nFontSize(i) = frm.FontSize
            nRightMargin(i) = frm.Controls(i).RightMargin
        End If
    Next

''    If bWinMax Or frm.WindowState = 2 Then ' maxim
''        frm.Height = Screen.Height
''        frm.Width = Screen.Width
''    Else
''        frm.Height = frm.Height * Screen.Height / 7290
''        frm.Width = frm.Width * Screen.Width / 9690
''    End If

    bFirstTime = True

End Sub

Public Sub FormResize(frm As Form)

    Dim i As Integer
    Dim nCaptionSize As Integer
    Dim dRatioX As Double
    Dim dRatioY As Double
    Dim nSaveRedraw As Long

    On Error Resume Next
    nSaveRedraw = frm.AutoRedraw

    frm.AutoRedraw = True

    If bFirstTime Then
        bFirstTime = False
        Exit Sub
    End If

    If frm.Height < nFormHeight / 2 Then frm.Height = nFormHeight / 2

    If frm.Width < nFormWidth / 2 Then frm.Width = nFormWidth / 2
    nCaptionSize = 400
    dRatioY = 1# * (nFormHeight - nCaptionSize) / (frm.Height - nCaptionSize)
    dRatioX = 1# * nFormWidth / frm.Width
    On Error Resume Next ' for comboboxes, timeres and other nonsizible controls

    For i = 0 To nNumOfControls
        If TypeOf frm.Controls(i) Is Line Then
            frm.Controls(i).Y1 = Int(nTop(i) / dRatioY)
            frm.Controls(i).X1 = Int(nLeft(i) / dRatioX)
            frm.Controls(i).Y2 = Int(nHeight(i) / dRatioY)
            frm.Controls(i).X2 = Int(nWidth(i) / dRatioX)
        Else
            frm.Controls(i).Top = Int(nTop(i) / dRatioY)
            frm.Controls(i).Left = Int(nLeft(i) / dRatioX)
            If Not (TypeOf frm.Controls(i) Is Image) Then
                frm.Controls(i).Height = Int(nHeight(i) / dRatioY)
                frm.Controls(i).Width = Int(nWidth(i) / dRatioX)
            End If
' frm.Controls(i).FontSize = Int(nFontSize(i) / dRatioX) + Int(nFontSize(i) / dRatioX) Mod 2
            frm.Controls(i).RightMargin = Int(nRightMargin(i) / dRatioY)
        End If
    Next
    frm.AutoRedraw = nSaveRedraw
End Sub

