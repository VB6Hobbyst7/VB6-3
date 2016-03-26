VERSION 5.00
Begin VB.Form frmOutput 
   BackColor       =   &H00000000&
   Caption         =   "Output"
   ClientHeight    =   5760
   ClientLeft      =   4755
   ClientTop       =   5670
   ClientWidth     =   7680
   Icon            =   "frmOutput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   Begin VB.Timer tmrRender 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer tmrFPSCount 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nothing To Render..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2985
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Module:        frmOutput
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  mSystem.bas
' Last revision: 2006.06.23
'================================================

Option Base 0
Option Explicit

Private lngBorderX As Long
Private lngBorderY As Long

Private Sub Form_Activate()
  If bOK Then frmSetup.SetFocus
End Sub


Private Sub Form_Load()
  
  bTexRotate = True
  bNormalMap = True
  bGCh = True
  bGTexture = True
  bRenderGrid = True
  bLighting = True
  bRenderStatic = True
  bReflections = True
  bTexture = True
  bWireframe = False
  bShade = True
  bWiregrid = False
  bAlphaBlend = True
  lngFPS = 0
 
  Show
  With Screen
    ' Resize window
    lngBorderX = Width - ScaleWidth * .TwipsPerPixelX
    lngBorderY = Height - ScaleHeight * .TwipsPerPixelY
    Width = lngWidth * .TwipsPerPixelX + lngBorderX
    Height = lngHeight * .TwipsPerPixelY + lngBorderY
    'center screen
    Left = .Width / 2 - Width / 2
    Top = .Height / 2 - Height / 2
  End With
  DoEvents
End Sub

Private Sub Form_Paint()
  If bOK Then RenderScene
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Shutdown
End Sub

Private Sub tmrFPSCount_Timer()
  Caption = "Output: " & lngWidth & "x" & lngHeight & " - " & lngFPS & " fps."
  If bSafeMode Then Caption = Caption & " (SafeMode)"
  ' Reset fps counter
  lngFPS = 0
End Sub

Private Sub tmrRender_Timer()
  RenderScene
End Sub

