VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   5535
   ClientLeft      =   12465
   ClientTop       =   6735
   ClientWidth     =   4935
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTopView 
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   28
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdStep 
      Caption         =   "Step"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   22
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fire Options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chkRotate 
         Caption         =   "Rot Texture"
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
         Left            =   840
         TabIndex        =   33
         ToolTipText     =   "Layer texture rotation feature"
         Top             =   3000
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkNormalMap 
         Caption         =   "Normal Mapping"
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
         Left            =   2160
         TabIndex        =   32
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkGCheck 
         Caption         =   "Checked"
         Enabled         =   0   'False
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
         Left            =   840
         TabIndex        =   31
         Top             =   3360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkGMap 
         Caption         =   "Allow Grid Mapping"
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
         Left            =   2160
         TabIndex        =   30
         Top             =   3360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkWg 
         Caption         =   "Wired grid"
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
         Left            =   3240
         TabIndex        =   27
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox cmbEnv 
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
         ItemData        =   "frmSetup.frx":0ECA
         Left            =   960
         List            =   "frmSetup.frx":0EE6
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3840
         Width           =   3495
      End
      Begin VB.CheckBox chkGShade 
         Caption         =   "Fine Shade"
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
         Left            =   840
         TabIndex        =   23
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkWire 
         Caption         =   "Wireframe"
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
         Left            =   3240
         TabIndex        =   18
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox chkTexture 
         Caption         =   "Texturing"
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
         Left            =   2160
         TabIndex        =   17
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkReflection 
         Caption         =   "Reflection"
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
         Left            =   840
         TabIndex        =   16
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkStatic 
         Caption         =   "Static Objects"
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
         Left            =   3240
         TabIndex        =   15
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkGrid 
         Caption         =   "Grid"
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
         Left            =   840
         TabIndex        =   14
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkLight 
         Caption         =   "Lightning"
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
         Left            =   2160
         TabIndex        =   13
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.HScrollBar scrCmp 
         Height          =   255
         LargeChange     =   20
         Left            =   840
         Max             =   100
         Min             =   -100
         TabIndex        =   11
         Top             =   1560
         Width           =   3135
      End
      Begin VB.HScrollBar scrFade 
         Height          =   255
         LargeChange     =   10
         Left            =   840
         Max             =   100
         TabIndex        =   8
         Top             =   1200
         Value           =   30
         Width           =   3135
      End
      Begin VB.HScrollBar scrHeight 
         Height          =   255
         LargeChange     =   10
         Left            =   840
         Max             =   100
         TabIndex        =   5
         Top             =   840
         Value           =   30
         Width           =   3135
      End
      Begin VB.HScrollBar scrWind 
         Height          =   255
         LargeChange     =   20
         Left            =   840
         Max             =   100
         Min             =   -100
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox chkAlpha 
         Caption         =   "Use Alpha"
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
         Left            =   2160
         TabIndex        =   29
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Label InfoPreset 
         AutoSize        =   -1  'True
         Caption         =   "Preset:"
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
         TabIndex        =   25
         Top             =   3840
         Width           =   600
      End
      Begin VB.Label Info 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You can also change fire radius, light && flame colors, amount of particles and other things, see source code..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   4200
         Width           =   4215
      End
      Begin VB.Label valCmp 
         AutoSize        =   -1  'True
         Caption         =   "1"
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
         Left            =   4080
         TabIndex        =   12
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label infoCmp 
         AutoSize        =   -1  'True
         Caption         =   "Cmp:"
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
         TabIndex        =   10
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label valFade 
         AutoSize        =   -1  'True
         Caption         =   "0,15"
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
         Left            =   4080
         TabIndex        =   9
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label infoFade 
         AutoSize        =   -1  'True
         Caption         =   "Fade:"
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
         TabIndex        =   7
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label valHeight 
         AutoSize        =   -1  'True
         Caption         =   "0,3"
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
         Left            =   4080
         TabIndex        =   6
         Top             =   840
         Width           =   240
      End
      Begin VB.Label infoHeight 
         AutoSize        =   -1  'True
         Caption         =   "Height:"
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
         Top             =   840
         Width           =   525
      End
      Begin VB.Label valWind 
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   4080
         TabIndex        =   3
         Top             =   480
         Width           =   90
      End
      Begin VB.Label infoWind 
         AutoSize        =   -1  'True
         Caption         =   "Wind:"
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
         TabIndex        =   1
         Top             =   480
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Module:        frmSetup
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  mEnvironment.bas, mSystem.bas
' Last revision: 2006.06.23
'================================================

Option Base 0
Option Explicit

' Camera height and distance
Private sngCamH As Single
Private sngCamD As Single

Private Sub chkchkGCheck_Click()

End Sub

Private Sub chkGMap_Click()
  bGTexture = CBool(chkGMap.Value)
  'render scene frame
  RenderScene
End Sub

Private Sub chkNormalMap_Click()
  bNormalMap = CBool(chkNormalMap.Value)
  chkGCheck.Enabled = Not bNormalMap
  'render scene frame
  RenderScene
End Sub

Private Sub Rotate_Click()

End Sub

Private Sub chkRotate_Click()
    bTexRotate = CBool(chkRotate.Value)
End Sub

Private Sub Form_Load()
  Show
 
  Left = frmOutput.Left + 15 * Screen.TwipsPerPixelX
  Top = frmOutput.Top + frmOutput.Height - frmSetup.Height - 15 * Screen.TwipsPerPixelY
  DoEvents
End Sub

' Toggle alphablending for static objects
Private Sub chkAlpha_Click()
  bAlphaBlend = CBool(chkAlpha.Value)
  RenderScene
End Sub

' Change scene & effect fire style
Private Sub cmbEnv_Click()
  
  ' Pause rendering, when loading
  frmOutput.tmrRender.Enabled = False
  DoEvents
  frmOutput.Refresh
  
  ' Load static scene geometry
  chkAlpha.Enabled = False
  Select Case cmbEnv.ListIndex
    Case 0: BootEnv "mesh\fxScene_01.ase"
    Case 1: BootEnv "mesh\fxScene_02.ase"
    Case 2: BootEnv "mesh\fxScene_03.ase": chkAlpha.Enabled = True
    Case 3: BootEnv "mesh\fxScene_04.ase"
    Case 4: BootEnv "mesh\fxScene_05.ase"
    Case 5: BootEnv "mesh\fxScene_06.ase"
    Case 6: BootEnv "mesh\fxScene_07.ase"
    Case 7: BootEnv "mesh\fxScene_08.ase"
  End Select
  
  RenderScene
  
  If Not cmdStep.Enabled Then
    frmOutput.tmrRender.Enabled = True
  End If
End Sub

' Camera top view (close)
Private Sub cmdTopView_Click()
  If cmdTopView.Caption = "Top" Then
    sngCamH = sngCamFinalHeight
    sngCamD = sngCamFinalDistance
    sngCamFinalHeight = -sngCamFinalDistance + 60
    sngCamFinalDistance = 1
    cmdTopView.Caption = "Back"
  Else
    sngCamFinalHeight = sngCamH
    sngCamFinalDistance = sngCamD
    cmdTopView.Caption = "Top"
  End If
End Sub

' Toggle shade mode
Private Sub chkGShade_Click()
  bShade = CBool(chkGShade.Value)
  RenderScene
End Sub

' Pause/play demo
Private Sub cmdPause_Click()
  If cmdPause.Caption = "Pause" Then
    ' Pause playing
    cmdPause.Caption = "Play"
    cmdStep.Enabled = True
    frmOutput.tmrRender.Enabled = False
    frmOutput.Refresh
  Else
    ' Start playing
    cmdPause.Caption = "Pause"
    cmdStep.Enabled = False
    frmOutput.tmrRender.Enabled = True
  End If
End Sub

Private Sub cmdQuit_Click()
  Shutdown
End Sub

' Toggle wire grid
Private Sub chkWg_Click()
  bWiregrid = CBool(chkWg.Value)
  RenderScene
End Sub

' Toggle wireframe/solid rendering
Private Sub chkWire_Click()
  bWireframe = CBool(chkWire.Value)
  RenderScene
End Sub

' Toggle texturing for static objects
Private Sub chkTexture_Click()
  bTexture = CBool(chkTexture.Value)
  RenderScene
End Sub

' Toggle specular lightning
Private Sub chkReflection_Click()
  bReflections = CBool(chkReflection.Value)
  RenderScene
End Sub

' Enable/disable static objects rendering
Private Sub chkStatic_Click()
  bRenderStatic = CBool(chkStatic.Value)
  chkTexture.Enabled = bRenderStatic
  RenderScene
End Sub

' Render frame by frame
Private Sub cmdStep_Click()
  RenderScene
  frmOutput.Refresh
End Sub

' Enable/disable grid rendering
Private Sub chkGrid_Click()
  bRenderGrid = CBool(chkGrid.Value)
  chkWg.Enabled = bRenderGrid
  RenderScene
End Sub

' Enable/disable lightning for grid
Private Sub chkLight_Click()
  bLighting = CBool(chkLight.Value)
  chkReflection.Enabled = bLighting
  RenderScene
End Sub

' Reset effect settings
Private Sub cmdReset_Click()
  GUI_SetScrolls
 
  chkGrid.Value = 1
  chkLight.Value = 1
  chkStatic.Value = 1
  chkReflection.Value = 1
  chkWire.Value = 0
  chkTexture.Value = 1
  chkGShade.Value = 1
  chkWg.Value = 0
  chkAlpha.Value = 1
  chkGrid_Click
  cmdReset.Enabled = False
  frmOutput.Refresh
End Sub

Private Sub chkGCheck_Click()
  frmOutput.tmrRender.Enabled = False
  bGCh = CBool(chkGCheck.Value)
  DOgrid
  'render scene frame
  RenderScene
  'continue rendering if needed
  If Not cmdStep.Enabled Then frmOutput.tmrRender.Enabled = True
End Sub

Private Sub scrCmp_Change()
  For lngNumber = 0 To UBound(cCore()) Step 1
    If scrCmp.Value > 0 Then
      cCore(lngNumber).sngCompression = (scrCmp.Value) / 1000 + 1
    Else
      If scrCmp.Value < 0 Then
        cCore(lngNumber).sngCompression = 1 + (scrCmp.Value) / 1000
      Else
        cCore(lngNumber).sngCompression = 1
      End If
    End If
    valCmp.Caption = cCore(lngNumber).sngCompression & vbNullString
  Next lngNumber
  cmdReset.Enabled = True
End Sub

Private Sub scrCmp_Scroll()
  scrCmp_Change
End Sub

Private Sub scrFade_Change()
  For lngNumber = 0 To UBound(cCore()) Step 1
    cCore(lngNumber).sngDecRadius = scrFade.Value / 200
    valFade.Caption = cCore(lngNumber).sngDecRadius & vbNullString
  Next lngNumber
  cmdReset.Enabled = True
End Sub

Private Sub scrFade_Scroll()
  scrFade_Change
End Sub

Private Sub scrHeight_Change()
  For lngNumber = 0 To UBound(cCore()) Step 1
    cCore(lngNumber).sngIncHeight = scrHeight.Value / 100
    valHeight.Caption = cCore(lngNumber).sngIncHeight & vbNullString
  Next lngNumber
  cmdReset.Enabled = True
End Sub

Private Sub scrHeight_Scroll()
  scrHeight_Change
End Sub

Private Sub scrWind_Change()
  For lngNumber = 0 To UBound(cCore()) Step 1
    cCore(lngNumber).sngWind = -scrWind.Value / 100
    valWind.Caption = cCore(lngNumber).sngWind & vbNullString
  Next lngNumber
  cmdReset.Enabled = True
End Sub

Private Sub scrWind_Scroll()
  scrWind_Change
End Sub
