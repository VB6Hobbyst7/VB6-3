VERSION 5.00
Begin VB.Form frmStartup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Startup"
   ClientHeight    =   3285
   ClientLeft      =   11625
   ClientTop       =   10500
   ClientWidth     =   4095
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox chkSafe 
      Caption         =   "SafeMode (Low Quality)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.ComboBox cmbResolution 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmStartup.frx":014A
      Left            =   1320
      List            =   "frmStartup.frx":015A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Info 
      Caption         =   $"frmStartup.frx":018B
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label infoResolution 
      AutoSize        =   -1  'True
      Caption         =   "Resolution:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Module:        frmStartup
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  None
' Last revision: 2006.06.23
'================================================

Option Base 0
Option Explicit

Private Sub Form_Load()
  
  Show
  DoEvents
  
  cmbResolution.ListIndex = 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  End
End Sub

Private Sub cmdQuit_Click()
  Unload frmStartup
End Sub


Private Sub cmdRun_Click()
  'hide setup dialog
  frmStartup.Hide
  'choose resolution
  Select Case cmbResolution.ListIndex
    Case 0
      lngWidth = 640
      lngHeight = 480
    Case 1
      lngWidth = 800
      lngHeight = 600
    Case 2
      lngWidth = 1024
      lngHeight = 768
    Case 3
      lngWidth = Screen.Width / Screen.TwipsPerPixelX
      lngHeight = Screen.Height / Screen.TwipsPerPixelY
  End Select
  
  bSafeMode = CBool(chkSafe.Value)
 
  Initialize
End Sub

