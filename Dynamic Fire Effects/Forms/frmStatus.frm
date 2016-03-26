VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importing Scene..."
   ClientHeight    =   840
   ClientLeft      =   8265
   ClientTop       =   7950
   ClientWidth     =   4335
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   56
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   Begin VB.Label FileName 
      AutoSize        =   -1  'True
      Caption         =   "No File."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   585
   End
   Begin VB.Label percent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "initializing..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   495
      Width           =   4155
   End
   Begin VB.Shape foreground 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   165
      Top             =   495
      Width           =   2055
   End
   Begin VB.Shape background 
      FillColor       =   &H8000000C&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   150
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Module:        frmStatus
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  None
' Last revision: 2006.06.23
'================================================

Option Explicit
Option Base 0

Private Sub Form_Load()
  Caption = "conversion"
  foreground.Width = 0
  percent.Caption = "initializing..."
  Show
  DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

