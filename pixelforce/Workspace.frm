VERSION 5.00
Begin VB.Form Workspace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PixelForce - [RayTracer]"
   ClientHeight    =   8055
   ClientLeft      =   1725
   ClientTop       =   2130
   ClientWidth     =   11175
   Icon            =   "Workspace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   Begin PixelForce.ProgressIndicator ProgressIndicator 
      Height          =   270
      Left            =   5040
      TabIndex        =   36
      Top             =   7680
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   476
   End
   Begin VB.Frame frameInfo 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   33
      Top             =   6840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.Label infoHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   435
         Left            =   120
         TabIndex        =   35
         Top             =   300
         Width           =   6195
      End
      Begin VB.Label infoTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Syntax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   60
         Width           =   600
      End
   End
   Begin PixelForce.AlphaButton InsertTorus 
      Height          =   420
      Left            =   10560
      TabIndex        =   32
      Top             =   3600
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton InsertChamferBox 
      Height          =   420
      Left            =   10080
      TabIndex        =   31
      Top             =   3600
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton InsertPrism 
      Height          =   420
      Left            =   9600
      TabIndex        =   30
      Top             =   3600
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton InsertTube 
      Height          =   420
      Left            =   9120
      TabIndex        =   29
      Top             =   3600
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton InsertPyramid 
      Height          =   420
      Left            =   8640
      TabIndex        =   28
      Top             =   3600
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton InsertCone 
      Height          =   420
      Left            =   10560
      TabIndex        =   27
      Top             =   3120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton InsertCylinder 
      Height          =   420
      Left            =   10080
      TabIndex        =   26
      Top             =   3120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton InsertSphere 
      Height          =   420
      Left            =   9600
      TabIndex        =   25
      Top             =   3120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton InsertBox 
      Height          =   420
      Left            =   9120
      TabIndex        =   24
      Top             =   3120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton InsertPlane 
      Height          =   420
      Left            =   8640
      TabIndex        =   23
      Top             =   3120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton AddLight 
      Height          =   420
      Left            =   8640
      TabIndex        =   21
      Top             =   2040
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton AddTexture 
      Height          =   420
      Left            =   8640
      TabIndex        =   20
      Top             =   1560
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton AddMesh 
      Height          =   420
      Left            =   8640
      TabIndex        =   19
      Top             =   1080
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Timer Stopper 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10680
      Top             =   7560
   End
   Begin PixelForce.AlphaButton ButtonCancel 
      Height          =   420
      Left            =   8160
      TabIndex        =   16
      Top             =   120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton ButtonExit 
      Height          =   420
      Left            =   9600
      TabIndex        =   14
      Top             =   120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton ButtonResult 
      Height          =   420
      Left            =   1920
      TabIndex        =   13
      Top             =   120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.ComboBox Resolution 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Workspace.frx":014A
      Left            =   6240
      List            =   "Workspace.frx":0160
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   180
      Width           =   1815
   End
   Begin VB.TextBox SceneCode 
      DataSource      =   "465"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6960
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   600
      Width           =   8535
   End
   Begin PixelForce.AlphaButton ButtonRender 
      Height          =   420
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton ButtonSave 
      Height          =   420
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton ButtonOpen 
      Height          =   420
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton ButtonNew 
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton FeatureBlurring 
      Height          =   420
      Left            =   8640
      TabIndex        =   5
      Top             =   7560
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton FeatureFiltering 
      Height          =   420
      Left            =   8640
      TabIndex        =   6
      Top             =   7080
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton FeatureTexturing 
      Height          =   420
      Left            =   8640
      TabIndex        =   7
      Top             =   6600
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton FeatureLighting 
      Height          =   420
      Left            =   8640
      TabIndex        =   8
      Top             =   6120
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton FeatureAlphaBlending 
      Height          =   420
      Left            =   8640
      TabIndex        =   9
      Top             =   5640
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton FeatureZSort 
      Height          =   420
      Left            =   8640
      TabIndex        =   10
      Top             =   5160
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin PixelForce.AlphaButton FeatureDepthTest 
      Height          =   420
      Left            =   8640
      TabIndex        =   11
      Top             =   4680
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label infoMeshes 
      Alignment       =   2  'Center
      Caption         =   "Primitives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8640
      TabIndex        =   22
      Top             =   2760
      Width           =   2505
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   576
      X2              =   744
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   576
      X2              =   744
      Y1              =   201
      Y2              =   201
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   576
      X2              =   744
      Y1              =   305
      Y2              =   305
   End
   Begin VB.Line Separator3 
      BorderColor     =   &H80000014&
      X1              =   576
      X2              =   744
      Y1              =   65
      Y2              =   65
   End
   Begin VB.Line Separator2 
      BorderColor     =   &H80000010&
      X1              =   576
      X2              =   744
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Line Separator1 
      BorderColor     =   &H80000010&
      X1              =   576
      X2              =   744
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Label infoPrimitives 
      Alignment       =   2  'Center
      Caption         =   "System Objects"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8640
      TabIndex        =   18
      Top             =   720
      Width           =   2505
   End
   Begin VB.Label infoFeatures 
      Alignment       =   2  'Center
      Caption         =   "Features"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8640
      TabIndex        =   17
      Top             =   4320
      Width           =   2505
   End
   Begin VB.Label infoState 
      Caption         =   "PixelForce RayTracer 1.0 alpha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   15
      Top             =   7680
      Width           =   4905
   End
End
Attribute VB_Name = "Workspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'just to keep code clear
Option Explicit
Option Base 0

'path variables for load/save functions
Private LastPath As String
Private ThisPath As String

'tga pixelmap for saving tga image
Private TgaSaver As TgaFile

'configuration variables
Private ConfLine As String
Private ConfSeparate As Integer
Private ConfParameter As String
Private ConfValue As String

'syntax help variables
Private NextSep As Integer
Private PrevSep As Integer
Private CurLine As String
Private SearchBuffer As Integer

'add light code fragment
Private Sub AddLight_ABClick()
  With SceneCode
    'add code
    .Text = .Text & vbCrLf
    .Text = .Text & "<Light>" & vbCrLf
    .Text = .Text & "  Position 0.00 0.00 0.00" & vbCrLf
    .Text = .Text & "  Color 127 127 127" & vbCrLf
    .Text = .Text & "  Range 1.00" & vbCrLf
    .Text = .Text & "  Amplify 1.00" & vbCrLf
    .Text = .Text & "  Alpha 0" & vbCrLf
    .Text = .Text & "</Light>" & vbCrLf
    'move cursor to the end
    .SelStart = Len(.Text)
    .SelLength = 0
      .SetFocus
  End With
End Sub

Private Sub AddLight_ABRollOver()
  infoState.Caption = "Add Area-Light Declaration"
End Sub

'add mesh code fragment
Private Sub AddMesh_ABClick()
  With SceneCode
    'ask for file
    ThisPath = OpenFile(hwnd, "Add Mesh", fixPath & "Mesh\", "*.txt - Text Mesh Files|*.txt", "*.txt", &H1000 Or &H800)
    If Len(ThisPath) > 0 Then
      LastPath = ThisPath
      'cut 0x0 symbol
      If Right(ThisPath, 1) = Chr(0) Then ThisPath = Left(ThisPath, Len(ThisPath) - 1)
      'add code
      InsertMesh ThisPath
      .SetFocus
    End If
  End With
End Sub

Private Sub AddMesh_ABRollOver()
  infoState.Caption = "Add External Custom Mesh"
End Sub

'add texture code fragment
Private Sub AddTexture_ABClick()
  With SceneCode
    'ask for file
    ThisPath = OpenFile(hwnd, "Add Texture", fixPath & "Texture\", "*.tga - TrueVision X-File (Targa)|*.tga", "*.tga", &H1000 Or &H800)
    If Len(ThisPath) > 0 Then
      LastPath = ThisPath
      'cut 0x0 symbol
      If Right(ThisPath, 1) = Chr(0) Then ThisPath = Left(ThisPath, Len(ThisPath) - 1)
      'add code
      .Text = .Text & vbCrLf
      .Text = .Text & "<DiffuseMap>" & vbCrLf
      .Text = .Text & "  File " & ThisPath & vbCrLf
      .Text = .Text & "  Transparency on" & vbCrLf
      .Text = .Text & "  Generate32Bit off" & vbCrLf
      .Text = .Text & "  Alpha 0" & vbCrLf
      .Text = .Text & "</DiffuseMap>" & vbCrLf
      'move cursor to the end
      .SelStart = Len(.Text)
      .SelLength = 0
      .SetFocus
    End If
  End With
End Sub

Private Sub AddTexture_ABRollOver()
  infoState.Caption = "Add External Custom Texture"
End Sub

'cancel rendering
Private Sub ButtonCancel_ABClick()
  RenderStop = True
  Stopper.Enabled = True
  SceneCode.SetFocus
End Sub

Private Sub ButtonCancel_ABRollOver()
  infoState.Caption = "Stop Rendering Process"
End Sub

'exit program
Private Sub ButtonExit_ABClick()
  Unload Me
End Sub

Private Sub ButtonExit_ABRollOver()
  infoState.Caption = "Save Configuration And Close Application"
End Sub

'new scene
Private Sub ButtonNew_ABClick()
  'failed to load template? use default
  If Not LoadScene(fixPath & "Config\Scene.txt") Then
    With SceneCode
      .Text = vbNullString
      .Text = .Text & "" & vbCrLf
      .Text = .Text & "//PixelForce RayTracer Scene File" & vbCrLf
      .Text = .Text & "" & vbCrLf
      .Text = .Text & "" & vbCrLf
      .Text = .Text & "<BackBuffer>" & vbCrLf
      .Text = .Text & "  Clear on" & vbCrLf
      .Text = .Text & "  Color 0 0 0" & vbCrLf
      .Text = .Text & "  AntiAliasLevel 0" & vbCrLf
      .Text = .Text & "  AliasEdgeOnly on" & vbCrLf
      .Text = .Text & "</BackBuffer>" & vbCrLf
      .Text = .Text & "" & vbCrLf
      .Text = .Text & "<ClippingDistance>" & vbCrLf
      .Text = .Text & "  ZNear -100000" & vbCrLf
      .Text = .Text & "  ZFar 100000" & vbCrLf
      .Text = .Text & "</ClippingDistance>" & vbCrLf
      .Text = .Text & "" & vbCrLf
      .Text = .Text & "<Camera>" & vbCrLf
      .Text = .Text & "  Position 0.00 0.00 0.00" & vbCrLf
      .Text = .Text & "  Rotation 0.00 0.00" & vbCrLf
      .Text = .Text & "  Scale 1.00" & vbCrLf
      .Text = .Text & "</Camera>" & vbCrLf
      .Text = .Text & "" & vbCrLf
    End With
  End If
  SceneCode.SetFocus
End Sub

Private Sub ButtonNew_ABRollOver()
  infoState.Caption = "Create New Scene From Default Template"
End Sub

'open file
Private Sub ButtonOpen_ABClick()
  'ask for file
  ThisPath = OpenFile(hwnd, "Open Scene", fixPath & "Scene\", "*.txt - PixelForce RayTracer Scene Files|*.txt", "*.txt", &H1000 Or &H800)
  'load scene file
  If Len(ThisPath) > 0 Then
    LastPath = ThisPath
    LoadScene ThisPath
    SceneCode.SetFocus
  End If
End Sub

Private Sub ButtonOpen_ABRollOver()
  infoState.Caption = "Load Scene From File"
End Sub

Private Sub ButtonRender_ABClick()
  SceneCode.SetFocus
  'show output window
  Unload Output
  Load Output
  'disable controls
  RenderStop = False
  ButtonCancel.AB_Disabled = False
  ButtonCancel.AB_RenderIcon
  ButtonResult.AB_Disabled = True
  ButtonResult.AB_RenderIcon
  Resolution.Enabled = False
  ButtonRender.AB_Disabled = True
  ButtonRender.AB_RenderIcon
  'reset renderer
  ResetPrimitives
  ResetTextures
  ResetLights
  'launch parser
  ParseScene
  'start rendering
  RenderScene Output
End Sub

'show output window
Private Sub ButtonShowResult_ABClick()
  SceneCode.SetFocus
  Output.Show
End Sub

Private Sub ButtonRender_ABRollOver()
  infoState.Caption = "Start Rendering Process"
End Sub

'save output image file
Private Sub ButtonResult_ABClick()
  SceneCode.SetFocus
  'ask for file
  ThisPath = SaveFile(hwnd, "Save Output Image", fixPath & "Output\", "*.tga - TrueVision X-File (Targa)|*.tga", "*.tga", &H2)
  'save picture file
  If Len(ThisPath) > 0 Then
    LastPath = ThisPath
    ButtonResult.AB_Disabled = True
    ButtonResult.AB_RenderIcon
    'create tga file handler
    Set TgaSaver = New TgaFile
    'prepare image data
    With TgaSaver
      .Width = BufX
      .Height = BufY
      .CreateEmptyPixelMap
      .AlphaBits = 0
      .Comment = "Produced with PixelForce raytracer"
      .CommentLength = Len(.Comment)
      .Ending = .StandardEnding
      .FileName = ThisPath
      .Format = 24
      .LockAlpha = False
      .Rle = False
      .RleRatio = 0
      .RleBytes = 0
      .FileSize = 0
      'show info
      DoEvents
      infoState.Caption = "Writing TGA file... Please Wait!"
      'set image bits
      .SetBits Out()
      'try to save image
      If Not .SaveTga(.FileName) Then MsgBox "Failed To Save Output Image File.", vbCritical + vbOKOnly, "Error"
      .Destroy
    End With
    'destroy tga handler
    Set TgaSaver = Nothing
    infoState.Caption = "Ready."
    ButtonResult.AB_Disabled = False
    ButtonResult.AB_RenderIcon
  End If
End Sub

Private Sub ButtonResult_ABRollOver()
  infoState.Caption = "Save Current Rendering Results To File In TGA Format"
End Sub

'save file
Private Sub ButtonSave_ABClick()
  SceneCode.SetFocus
  'ask for file
  ThisPath = SaveFile(hwnd, "Save Scene", fixPath & "Scene\", "*.txt - PixelForce RayTracer Scene Files|*.txt", "*.txt", &H2)
  'save scene file
  If Len(ThisPath) > 0 Then
    LastPath = ThisPath
    SaveScene ThisPath
  End If
End Sub

Private Sub ButtonSave_ABRollOver()
  infoState.Caption = "Save Current Scene To File"
End Sub

'toggle alphablending
Private Sub FeatureAlphaBlending_ABClick()
  SceneCode.SetFocus
  FeatureAlphaBlending.AB_Pushed = Not FeatureAlphaBlending.AB_Pushed
  FeatureAlphaBlending.AB_RenderIcon
  AlphaBlending = FeatureAlphaBlending.AB_Pushed
  If Not AlphaBlending Then
    DepthSorting = False
    FeatureZSort.AB_Disabled = True
    FeatureZSort.AB_RenderIcon
  Else
    DepthSorting = FeatureZSort.AB_Pushed
    FeatureZSort.AB_Disabled = False
    FeatureZSort.AB_RenderIcon
  End If
End Sub

Private Sub FeatureAlphaBlending_ABRollOver()
  infoState.Caption = "Enable/Disable 128-Bit Alpha+Light Color Blending"
End Sub

'toggle anti-aliasing
Private Sub FeatureBlurring_ABClick()
  SceneCode.SetFocus
  FeatureBlurring.AB_Pushed = Not FeatureBlurring.AB_Pushed
  FeatureBlurring.AB_RenderIcon
  Antialiasing = FeatureBlurring.AB_Pushed
End Sub

Private Sub FeatureBlurring_ABRollOver()
  infoState.Caption = "Enable/Disable Triangle-Edge Depth-Stencil Antialiasing (Blurring)"
End Sub

'toggle depth test
Private Sub FeatureDepthTest_ABClick()
  SceneCode.SetFocus
  FeatureDepthTest.AB_Pushed = Not FeatureDepthTest.AB_Pushed
  FeatureDepthTest.AB_RenderIcon
  DepthTest = FeatureDepthTest.AB_Pushed
End Sub

Private Sub FeatureDepthTest_ABRollOver()
  infoState.Caption = "Enable/Disable 32-Bit Z-Buffer"
End Sub

'toggle bilinear texture filtering
Private Sub FeatureFiltering_ABClick()
  SceneCode.SetFocus
  FeatureFiltering.AB_Pushed = Not FeatureFiltering.AB_Pushed
  FeatureFiltering.AB_RenderIcon
  TextureFiltering = FeatureFiltering.AB_Pushed
End Sub

Private Sub FeatureFiltering_ABRollOver()
  infoState.Caption = "Enable/Disable Bilinear Filter For Textures"
End Sub

'toggle lighting
Private Sub FeatureLighting_ABClick()
  SceneCode.SetFocus
  FeatureLighting.AB_Pushed = Not FeatureLighting.AB_Pushed
  FeatureLighting.AB_RenderIcon
  Lighting = FeatureLighting.AB_Pushed
End Sub

Private Sub FeatureLighting_ABRollOver()
  infoState.Caption = "Enable/Disable 32-Bit Area Lights"
End Sub

'toggle texturing
Private Sub FeatureTexturing_ABClick()
  SceneCode.SetFocus
  FeatureTexturing.AB_Pushed = Not FeatureTexturing.AB_Pushed
  FeatureTexturing.AB_RenderIcon
  Texturing = FeatureTexturing.AB_Pushed
End Sub

Private Sub FeatureTexturing_ABRollOver()
  infoState.Caption = "Enable/Disable 32-Bit Diffuse Maps Rendering"
End Sub

'toggle z-sort
Private Sub FeatureZSort_ABClick()
  SceneCode.SetFocus
  FeatureZSort.AB_Pushed = Not FeatureZSort.AB_Pushed
  FeatureZSort.AB_RenderIcon
  DepthSorting = FeatureZSort.AB_Pushed
End Sub

Private Sub FeatureZSort_ABRollOver()
  infoState.Caption = "Enable/Disable Triangle Shell-Sorting Before AlphaBlending"
End Sub

'show the form
Private Sub Form_Load()
  'error handler
  On Error Resume Next
  Show
  DoEvents
  'load buttons
  FeatureDepthTest.AB_Pushed = True
  FeatureZSort.AB_Pushed = True
  FeatureAlphaBlending.AB_Pushed = True
  FeatureLighting.AB_Pushed = True
  FeatureTexturing.AB_Pushed = True
  FeatureFiltering.AB_Pushed = True
  FeatureBlurring.AB_Pushed = True
  ButtonNew.AB_Tooltip "New Scene"
  ButtonOpen.AB_Tooltip "Open Scene"
  ButtonSave.AB_Tooltip "Save Scene"
  ButtonResult.AB_Tooltip "Save Result"
  ButtonRender.AB_Caption "Render Scene at:"
  ButtonCancel.AB_Tooltip "Stop Rendering"
  ButtonExit.AB_Caption "Exit Program"
  AddMesh.AB_Caption "Add Mesh"
  AddTexture.AB_Caption "Add Texture"
  AddLight.AB_Caption "Add Light Source"
  FeatureDepthTest.AB_Caption "Z-Depth Test"
  FeatureZSort.AB_Caption "Z-Sort (For AlphaBlend)"
  FeatureAlphaBlending.AB_Caption "Alpha Blend"
  FeatureLighting.AB_Caption "Area Lighting"
  FeatureTexturing.AB_Caption "Texturing"
  FeatureFiltering.AB_Caption "Bilinear Texture Filter"
  FeatureBlurring.AB_Caption "Anti-Aliasing / Blurring"
  ButtonNew.AB_LoadIcon fixPath & "ui\Button_New.tga"
  ButtonOpen.AB_LoadIcon fixPath & "ui\Button_Open.tga"
  ButtonSave.AB_LoadIcon fixPath & "ui\Button_Save.tga"
  ButtonResult.AB_LoadIcon fixPath & "ui\Button_Result.tga"
  ButtonRender.AB_LoadIcon fixPath & "ui\Button_Render.tga"
  ButtonCancel.AB_LoadIcon fixPath & "ui\Button_Cancel.tga"
  ButtonExit.AB_LoadIcon fixPath & "ui\Button_Exit.tga"
  AddMesh.AB_LoadIcon fixPath & "ui\Setup_Mesh.tga"
  AddTexture.AB_LoadIcon fixPath & "ui\Setup_Texture.tga"
  AddLight.AB_LoadIcon fixPath & "ui\Setup_Light.tga"
  FeatureDepthTest.AB_LoadIcon fixPath & "ui\Feature_DepthTest.tga"
  FeatureZSort.AB_LoadIcon fixPath & "ui\Feature_ZSort.tga"
  FeatureAlphaBlending.AB_LoadIcon fixPath & "ui\Feature_AlphaBlending.tga"
  FeatureLighting.AB_LoadIcon fixPath & "ui\Feature_Lighting.tga"
  FeatureTexturing.AB_LoadIcon fixPath & "ui\Feature_Texturing.tga"
  FeatureFiltering.AB_LoadIcon fixPath & "ui\Feature_Filtering.tga"
  FeatureBlurring.AB_LoadIcon fixPath & "ui\Feature_Blurring.tga"
  'draw progressbar
  ProgressIndicator.PI_LoadIcons fixPath & "ui\Indicator_Back.tga", fixPath & "ui\Indicator_Fore.tga"
  'load primitive buttons
  InsertPlane.AB_Tooltip "Insert Plane"
  InsertPlane.AB_LoadIcon fixPath & "\ui\Add_Plane.tga"
  InsertBox.AB_Tooltip "Insert Box"
  InsertBox.AB_LoadIcon fixPath & "\ui\Add_Box.tga"
  InsertSphere.AB_Tooltip "Insert Sphere"
  InsertSphere.AB_LoadIcon fixPath & "\ui\Add_Sphere.tga"
  InsertCylinder.AB_Tooltip "Insert Cylinder"
  InsertCylinder.AB_LoadIcon fixPath & "\ui\Add_Cylinder.tga"
  InsertCone.AB_Tooltip "Insert Cone"
  InsertCone.AB_LoadIcon fixPath & "\ui\Add_Cone.tga"
  InsertPyramid.AB_Tooltip "Insert Pyramid"
  InsertPyramid.AB_LoadIcon fixPath & "\ui\Add_Pyramid.tga"
  InsertTube.AB_Tooltip "Insert Tube"
  InsertTube.AB_LoadIcon fixPath & "\ui\Add_Tube.tga"
  InsertPrism.AB_Tooltip "Insert Prism"
  InsertPrism.AB_LoadIcon fixPath & "\ui\Add_Prism.tga"
  InsertChamferBox.AB_Tooltip "Insert Chamfer Box"
  InsertChamferBox.AB_LoadIcon fixPath & "\ui\Add_ChamferBox.tga"
  InsertTorus.AB_Tooltip "Insert Torus"
  InsertTorus.AB_LoadIcon fixPath & "\ui\Add_Torus.tga"
  'set disabled buttons
  DoEvents
  ButtonResult.AB_Disabled = True
  ButtonCancel.AB_Disabled = True
  ButtonResult.AB_RenderIcon
  ButtonCancel.AB_RenderIcon
  'defaults
  DepthSorting = True
  DepthTest = True
  Texturing = True
  TextureFiltering = True
  Lighting = True
  AlphaBlending = True
  ClearBuffer = True
  Antialiasing = True
  EdgeOnly = False
  'new scene
  ButtonNew_ABClick
  'load confguration
  Open fixPath & "Config\Options.txt" For Input As #1
  'failed?
  If Not Err.Number = 0 Then
    'show message
    Err.Number = 0
    Close #1
    MsgBox "Failed To Load Configuration File. Default settings will be used.", vbExclamation + vbOKOnly, "Error"
    'set defaults
    Resolution.Text = "800x600"
  Else
    'parse config file
    Do While Not EOF(1)
      'get line
      Line Input #1, ConfLine
      'can be parsed?
      If Len(ConfLine) > 5 Then
        'find separator
        ConfSeparate = InStr(3, ConfLine, " ", vbTextCompare)
        If ConfSeparate > 0 Then
          'extract parameter name
          ConfParameter = UCase(Left(ConfLine, ConfSeparate - 1))
          'extract value
          ConfValue = Right(ConfLine, Len(ConfLine) - ConfSeparate)
          'target setting selection
          Select Case ConfParameter
            'set resolution
            Case "  RESOLUTION"
              Resolution.Text = ConfValue
            'disable misc features if needed
            Case "  DEPTHTEST"
              If UCase(ConfValue) = "OFF" Then FeatureDepthTest_ABClick
            Case "  ZSORT"
              If UCase(ConfValue) = "OFF" Then FeatureZSort_ABClick
            Case "  ALPHABLEND"
              If UCase(ConfValue) = "OFF" Then FeatureAlphaBlending_ABClick
            Case "  LIGHTING"
              If UCase(ConfValue) = "OFF" Then FeatureLighting_ABClick
            Case "  TEXTURING"
              If UCase(ConfValue) = "OFF" Then FeatureTexturing_ABClick
            Case "  FILTERING"
              If UCase(ConfValue) = "OFF" Then FeatureFiltering_ABClick
            Case "  MULTISAMPLING"
              If UCase(ConfValue) = "OFF" Then FeatureBlurring_ABClick
          End Select
        End If
      End If
    Loop
    'done, close it
    Close #1
  End If
  'load syntax help
  ReDim SyntaxArray(0)
  LoadCodeSyntax
End Sub

'terminate application
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'error handler
  On Error Resume Next
  'save configuration
  Open fixPath & "Config\Options.txt" For Output As #1
  'failed?
  If Not Err.Number = 0 Then
    'show message
    Err.Number = 0
    Close #1
    MsgBox "Failed To Save Configuration File.", vbExclamation + vbOKOnly, "Error"
  Else
    'write settings
    Print #1, vbNullString
    Print #1, "//PixelForce RayTracer Confguration File"
    Print #1, vbNullString
    Print #1, vbNullString
    Print #1, "<settings>"
    Print #1, vbNullString
    Print #1, "  Resolution " & Resolution.Text
    Print #1, vbNullString
    'write features status
    If FeatureDepthTest.AB_Pushed Then
      Print #1, "  DepthTest on"
    Else
      Print #1, "  DepthTest off"
    End If
    If FeatureZSort.AB_Pushed Then
      Print #1, "  ZSort on"
    Else
      Print #1, "  ZSort off"
    End If
    If FeatureAlphaBlending.AB_Pushed Then
      Print #1, "  AlphaBlend on"
    Else
      Print #1, "  AlphaBlend off"
    End If
    If FeatureLighting.AB_Pushed Then
      Print #1, "  Lighting on"
    Else
      Print #1, "  Lighting off"
    End If
    If FeatureTexturing.AB_Pushed Then
      Print #1, "  Texturing on"
    Else
      Print #1, "  Texturing off"
    End If
    If FeatureFiltering.AB_Pushed Then
      Print #1, "  Filtering on"
    Else
      Print #1, "  Filtering off"
    End If
    If FeatureBlurring.AB_Pushed Then
      Print #1, "  MultiSampling on"
    Else
      Print #1, "  MultiSampling off"
    End If
    Print #1, vbNullString
    Print #1, "</settings>"
    'done
    Close #1
  End If
  'kill renderer
  ReleaseRenderer
  End
End Sub

'add box declaration
Private Sub InsertBox_ABClick()
  InsertMesh "$LocalPath\Mesh\Box.txt"
End Sub

Private Sub InsertBox_ABRollOver()
  infoState.Caption = "Add Standard 'Box' Mesh Declaration"
End Sub

'add chamferbox declaration
Private Sub InsertChamferBox_ABClick()
  InsertMesh "$LocalPath\Mesh\ChamferBox.txt"
End Sub

Private Sub InsertChamferBox_ABRollOver()
  infoState.Caption = "Add Standard 'Chamfer Box' Mesh Declaration"
End Sub

'add cone declaration
Private Sub InsertCone_ABClick()
  InsertMesh "$LocalPath\Mesh\Cone.txt"
End Sub

Private Sub InsertCone_ABRollOver()
  infoState.Caption = "Add Standard 'Cone' Mesh Declaration"
End Sub

'add cylinder declaration
Private Sub InsertCylinder_ABClick()
  InsertMesh "$LocalPath\Mesh\Cylinder.txt"
End Sub

Private Sub InsertCylinder_ABRollOver()
  infoState.Caption = "Add Standard 'Cylinder' Mesh Declaration"
End Sub

'add plane declaration
Private Sub InsertPlane_ABClick()
  InsertMesh "$LocalPath\Mesh\Plane.txt"
End Sub

Private Sub InsertPlane_ABRollOver()
  infoState.Caption = "Add Standard 'Plane' Mesh Declaration"
End Sub

'add prism declaration
Private Sub InsertPrism_ABClick()
  InsertMesh "$LocalPath\Mesh\Prism.txt"
End Sub

Private Sub InsertPrism_ABRollOver()
  infoState.Caption = "Add Standard 'Prism' Mesh Declaration"
End Sub

'add pyramid declaration
Private Sub InsertPyramid_ABClick()
  InsertMesh "$LocalPath\Mesh\Pyramid.txt"
End Sub

Private Sub InsertPyramid_ABRollOver()
  infoState.Caption = "Add Standard 'Pyramid' Mesh Declaration"
End Sub

'add sphere declaration
Private Sub InsertSphere_ABClick()
  InsertMesh "$LocalPath\Mesh\Sphere.txt"
End Sub

Private Sub InsertSphere_ABRollOver()
  infoState.Caption = "Add Standard 'Sphere' Mesh Declaration"
End Sub

'add torus declaration
Private Sub InsertTorus_ABClick()
  InsertMesh "$LocalPath\Mesh\Torus.txt"
End Sub

Private Sub InsertTorus_ABRollOver()
  infoState.Caption = "Add Standard 'Torus' Mesh Declaration"
End Sub

'add tybe declaration
Private Sub InsertTube_ABClick()
  InsertMesh "$LocalPath\Mesh\Tube.txt"
End Sub

Private Sub InsertTube_ABRollOver()
  infoState.Caption = "Add Standard 'Tube' Mesh Declaration"
End Sub

'switch resolution
Private Sub Resolution_Click()
  SceneCode.SetFocus
  'hide output window
  Unload Output
  'cleanup, if initialized
  ReleaseRenderer
  'reinitialize renderer
  infoState.Caption = "Initializing Renderer At " & Resolution.Text
  Select Case Resolution.ListIndex
  Case 0
    InitializeRenderer 320, 200
  Case 1
    InitializeRenderer 512, 384
  Case 2
    InitializeRenderer 640, 480
  Case 3
    InitializeRenderer 800, 600
  Case 4
    InitializeRenderer 1024, 768
  Case 5
    InitializeRenderer 1280, 1024
  End Select
  'all done
  infoState.Caption = "Renderer Initialized For " & Resolution.Text & " Pixels Resolution"
End Sub

'show hints in status bar, while editing scene code
Private Sub SceneCode_Change()
  'find current line
  NextSep = InStr(SceneCode.SelStart + 1, SceneCode.Text, vbCrLf, vbBinaryCompare)
  PrevSep = FindStrLeft(SceneCode.SelStart, SceneCode.Text, vbCrLf)
  If NextSep = 0 And PrevSep <> 0 Then CurLine = Right(SceneCode.Text, Len(SceneCode.Text) - PrevSep + 1)
  If PrevSep = 0 And NextSep <> 0 Then CurLine = Left(SceneCode.Text, NextSep - 1)
  If PrevSep = 0 And NextSep = 0 Then CurLine = SceneCode.Text
  If PrevSep <> 0 And NextSep <> 0 Then CurLine = Mid(SceneCode.Text, PrevSep, NextSep - PrevSep)
  'to up case
  CurLine = UCase(CurLine)
  'no messages by default
  infoTip.Caption = vbNullString
  'syntax help
  If SyntaxOK Then
    'use external extended syntax
    For SearchBuffer = 1 To UBound(SyntaxArray()) Step 1
      If InStr(1, CurLine, SyntaxArray(SearchBuffer).KeyWord, vbTextCompare) > 0 And SyntaxArray(SearchBuffer).KeyWord <> "" Then
        infoTip.Caption = SyntaxArray(SearchBuffer).Syntax
        infoHelp.Caption = SyntaxArray(SearchBuffer).Info & vbCrLf & SyntaxArray(SearchBuffer).MoreInfo
      End If
    Next SearchBuffer
  Else
    'use built-in syntax
    If InStr(1, CurLine, "LIGHT", vbTextCompare) > 0 Then
      infoTip.Caption = "<Light> Position, Color, Range, Amplify, Alpha </Light>"
      infoHelp.Caption = "Adds an Area-Light to scene"
    End If
    If InStr(1, CurLine, "BACKBUFFER", vbTextCompare) > 0 Then
      infoTip.Caption = "<BackBuffer> Clear, Color, AntiAliasLevel, AliasEdgeOnly </BackBuffer>"
      infoHelp.Caption = "General scene configuration"
    End If
    If InStr(1, CurLine, "CLIPPINGDISTANCE", vbTextCompare) > 0 Then
      infoTip.Caption = "<ClippingDistance> ZNear, ZFar </ClippingDistance>"
      infoHelp.Caption = "Min and Max clipping range settings"
    End If
    If InStr(1, CurLine, "CAMERA", vbTextCompare) > 0 Then
      infoTip.Caption = "<Camera> Position, Rotation, Scale </Camera>"
      infoHelp.Caption = "Orthographic camera settings"
    End If
    If InStr(1, CurLine, "DIFFUSEMAP", vbTextCompare) > 0 Then
      infoTip.Caption = "<DiffuseMap> File, Transparency, Generate32Bit, Alpha </DiffuseMap>"
      infoHelp.Caption = "Adds diffuse texture map to scene"
    End If
    If InStr(1, CurLine, "MESH", vbTextCompare) > 0 Then
      infoTip.Caption = "<Mesh> File, Position, Rotation, Scale, Texture, Lighting, Alpha </Mesh>"
      infoHelp.Caption = "Adds external mesh file to scene"
    End If
    If InStr(1, CurLine, "ZFAR", vbTextCompare) > 0 Then
      infoTip.Caption = "ZFar (float MaxDistance)"
      infoHelp.Caption = "Max drawing depth"
    End If
    If InStr(1, CurLine, "ZNEAR", vbTextCompare) > 0 Then
      infoTip.Caption = "ZNear (float MinDistance)"
      infoHelp.Caption = "Min drawing depth"
    End If
    If InStr(1, CurLine, "ROTATION", vbTextCompare) > 0 Then
      infoTip.Caption = "Rotation (float AngleX) (float AngleY) (float AngleZ)"
      infoHelp.Caption = "Rotates object or camera by X,Y and Z" & vbCrLf & "There is no need to define Z coord, when using this in camera section"
    End If
    If InStr(1, CurLine, "POSITION", vbTextCompare) > 0 Then
      infoTip.Caption = "Position (float X) (float Y) (float Z)"
      infoHelp.Caption = "Sets position of object or camera"
    End If
    If InStr(1, CurLine, "SCALE", vbTextCompare) > 0 Then
      infoTip.Caption = "Scale (float X) (float Y) (float Z)"
      infoHelp.Caption = "Scales object or sets camera zoom factor" & vbCrLf & "Y and Z coords are optional when using this declaration in camera section"
    End If
    If InStr(1, CurLine, "CLEAR", vbTextCompare) > 0 Then
      infoTip.Caption = "Clear (boolean Enable)"
      infoHelp.Caption = "Allows to clear backbuffer before frame rendering" & vbCrLf & "hint: Boolean variable can be set to ON/OFF only!"
    End If
    If InStr(1, CurLine, "COLOR", vbTextCompare) > 0 Then
      infoTip.Caption = "Color (byte Red) (byte Green) (byte Blue)"
      infoHelp.Caption = "Sets backbuffer background or light color" & vbCrLf & "Use special 'Alpha (byte Transparency)' command to set alpha color"
    End If
    If InStr(1, CurLine, "ALPHA", vbTextCompare) > 0 Then
      infoTip.Caption = "Alpha (byte Transparency)"
      infoHelp.Caption = "Sets transparency level for objects"
    End If
    If InStr(1, CurLine, "FILE", vbTextCompare) > 0 Then
      infoTip.Caption = "File (string FullPathName)"
      infoHelp.Caption = "Loads texture/mesh file into memory." & vbCrLf & "You have ability to use '$LocalPath' for current rendering engine path."
    End If
    If InStr(1, CurLine, "LIGHTING", vbTextCompare) > 0 Then
      infoTip.Caption = "Lighting (boolean Enable)"
      infoHelp.Caption = "Object will be rendered with light pass if set to ON" & vbCrLf & "hint: Boolean variable can be set to ON/OFF only!"
    End If
    If InStr(1, CurLine, "TEXTURE", vbTextCompare) > 0 Then
      infoTip.Caption = "Texture (integer Number)"
      infoHelp.Caption = "Allows object to use defined texture map" & vbCrLf & "When map can not be found, object will not be textured."
    End If
    If InStr(1, CurLine, "TRANSPARENCY", vbTextCompare) > 0 Then
      infoTip.Caption = "Transparency (boolean Enable)"
      infoHelp.Caption = "Allows texture to be rendered with alpha-blending" & vbCrLf & "If you want to generate alpha channel for texture use 'Generate32Bit' command."
    End If
    If InStr(1, CurLine, "GENERATE32BIT", vbTextCompare) > 0 Then
      infoTip.Caption = "Generate32Bit (boolean Enable)"
      infoHelp.Caption = "Generates 8-Bit alpha channel from 24-bit color channel." & vbCrLf & "Will be executed only in case, when texture has no alpha channel."
    End If
    If InStr(1, CurLine, "RANGE", vbTextCompare) > 0 Then
      infoTip.Caption = "Range (float Distance)"
      infoHelp.Caption = "Sets light max distance"
    End If
    If InStr(1, CurLine, "AMPLIFY", vbTextCompare) > 0 Then
      infoTip.Caption = "Amplify (float Factor)"
      infoHelp.Caption = "Light core size factor"
    End If
    If InStr(1, CurLine, "ANTIALIASLEVEL", vbTextCompare) > 0 Then
      infoTip.Caption = "AntiAliasLevel (integer BlurRadius)"
      infoHelp.Caption = "Sets triangle-edge blurring radius" & vbCrLf & "Edge-AntiAliasing is experimental feature, use it at your own risk!"
    End If
    If InStr(1, CurLine, "ALIASEDGEONLY", vbTextCompare) > 0 Then
      infoTip.Caption = "AliasEdgeOnly (boolean Enable)"
      infoHelp.Caption = "Blurs ALL triangle edges when set to ON" & vbCrLf & "If this variable set to OFF, only scene outline will be multisampled."
    End If
  End If
  'resize textbox
  If infoTip.Caption <> vbNullString Then
    If Not frameInfo.Visible Then
      frameInfo.Visible = True
      SceneCode.Height = 416
      SceneCode.SelStart = SceneCode.SelStart
    End If
  Else
    If frameInfo.Visible Then
      frameInfo.Visible = False
      SceneCode.Height = 465
      SceneCode.SelStart = SceneCode.SelStart
    End If
  End If
End Sub

'show hints (on mouse click)
Private Sub SceneCode_Click()
  SceneCode_Change
End Sub

'show hints (when cursor moves)
Private Sub SceneCode_KeyDown(KeyCode As Integer, Shift As Integer)
  SceneCode_Change
End Sub

'show hints (when cursor moves)
Private Sub SceneCode_KeyUp(KeyCode As Integer, Shift As Integer)
  SceneCode_Change
End Sub

'enable controls
Private Sub Stopper_Timer()
  On Error Resume Next
  Show
  SceneCode.SetFocus
  Stopper.Enabled = False
  infoState.Caption = "Render Stopped"
  Resolution.Enabled = True
  ButtonCancel.AB_Disabled = True
  ButtonCancel.AB_RenderIcon
  ButtonRender.AB_Disabled = False
  ButtonRender.AB_RenderIcon
  ButtonResult.AB_Disabled = False
  ButtonResult.AB_RenderIcon
End Sub
