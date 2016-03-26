VERSION 5.00
Begin VB.Form Output 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PixelForce - Output"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Output.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.Frame FrameStatistics 
      Caption         =   "Scene Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1800
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      Begin PixelForce.AlphaButton HideInfo 
         Height          =   420
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin VB.Label infoARGB 
         AutoSize        =   -1  'True
         Caption         =   "ARGB: 0 0 0 0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label infoLocked 
         AutoSize        =   -1  'True
         Caption         =   "Locked by: 0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   1260
      End
      Begin VB.Label infoZ 
         AutoSize        =   -1  'True
         Caption         =   "Depth: 0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label infoCursor 
         AutoSize        =   -1  'True
         Caption         =   "Cursor: 0 0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label infoDepth 
         AutoSize        =   -1  'True
         Caption         =   "Depth Buffer"
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
         TabIndex        =   6
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label infoTime 
         AutoSize        =   -1  'True
         Caption         =   "Rendering Time: 0.00 s"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label infoMaps 
         AutoSize        =   -1  'True
         Caption         =   "Texture Maps: unknown"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1740
      End
      Begin VB.Label infoLights 
         AutoSize        =   -1  'True
         Caption         =   "Area Lights: unknown"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1560
      End
      Begin VB.Label infoTriangles 
         AutoSize        =   -1  'True
         Caption         =   "Primitives: unknown"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1425
      End
   End
End
Attribute VB_Name = "Output"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'just to keep code clear
Option Explicit
Option Base 0

'x,y coords for grid
Private X As Long
Private Y As Long

'show the form first
Private Sub Form_Load()
  Show
  'get icon
  Icon = Workspace.Icon
  'set form size
  Width = BufX * Screen.TwipsPerPixelX + (Width - ScaleWidth * Screen.TwipsPerPixelX)
  Height = BufY * Screen.TwipsPerPixelY + (Height - ScaleHeight * Screen.TwipsPerPixelY)
  DoEvents
  'draw grid
  For Y = 0 To ScaleHeight / 8 Step 2
    For X = 0 To ScaleWidth / 8 Step 2
      Line (X * 8, Y * 8)-(X * 8 + 7, Y * 8 + 7), RGB(191, 191, 191), BF
      Line ((X + 1) * 8, (Y + 1) * 8)-((X + 1) * 8 + 7, (Y + 1) * 8 + 7), RGB(191, 191, 191), BF
      Line ((X + 1) * 8, Y * 8)-((X + 1) * 8 + 7, Y * 8 + 7), RGB(255, 255, 255), BF
      Line (X * 8, (Y + 1) * 8)-(X * 8 + 7, (Y + 1) * 8 + 7), RGB(255, 255, 255), BF
    Next X
  Next Y
  'load button
  HideInfo.AB_Tooltip "Hide Information Block"
  HideInfo.AB_LoadIcon fixPath & "ui\Button_Hide.tga"
End Sub

'refresh depth info
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  X = X + 1
  Y = Y + 1
  infoCursor.Caption = "Cursor: " & X & " " & Y
  If Depth(X, Y).Z - 1 = ZFar Then
    infoZ.Caption = "Depth: ZFar (" & ZFar & ") "
  Else
    infoZ.Caption = "Depth: " & Depth(X, Y).Z
  End If
  infoLocked.Caption = "Locked by: " & Depth(X, Y).i
  infoARGB.Caption = "ARGB: " & Out(X, Y).A & " " & Out(X, Y).R & " " & Out(X, Y).G & " " & Out(X, Y).B
End Sub

'cancel unload while rendering
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 And Workspace.ButtonRender.AB_Disabled Then Cancel = 1
End Sub

'move frame
Private Sub Form_Resize()
  If Not WindowState = 1 Then FrameStatistics.Move ScaleWidth - FrameStatistics.Width, ScaleHeight - FrameStatistics.Height
End Sub

'hide info box
Private Sub HideInfo_ABClick()
  FrameStatistics.Visible = False
End Sub
