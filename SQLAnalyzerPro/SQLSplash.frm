VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4140
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "SQLSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      Height          =   3945
      Left            =   105
      Top             =   90
      Width           =   7050
   End
   Begin VB.Image imgLogo 
      Height          =   3780
      Left            =   420
      Picture         =   "SQLSplash.frx":000C
      Top             =   180
      Width           =   2130
   End
   Begin VB.Label lblProduct 
      BackColor       =   &H00FFFFFF&
      Caption         =   "This product is licensed to:"
      Height          =   285
      Left            =   2895
      TabIndex        =   3
      Top             =   3030
      Width           =   3885
   End
   Begin VB.Label lblLiscence 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Liscenced"
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   2895
      TabIndex        =   2
      Top             =   3315
      Width           =   4080
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   2895
      TabIndex        =   1
      Top             =   285
      Width           =   3135
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2895
      TabIndex        =   0
      Top             =   1140
      Width           =   3180
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblTitle.Caption = "SQL Analyzer™ Professional Edition"
    lblVersion.Caption = "Version 5.00.01"
    lblLiscence.Caption = "You!" & vbCrLf & "Release Number : 101020031637"
    'lblLiscence.Caption = "Soft Systems Limited, Cochin" & vbCrLf & "Release Number : 101020031637"
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

