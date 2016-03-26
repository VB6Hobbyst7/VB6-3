VERSION 5.00
Begin VB.Form windowFrameBuffer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frame Buffer Options"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton buttonCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   2280
      TabIndex        =   27
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton buttonOk 
      Caption         =   "&Ok"
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
      Left            =   1200
      TabIndex        =   26
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox boxA 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      MaxLength       =   8
      TabIndex        =   21
      ToolTipText     =   "Border Color (Alpha Component / Transparency)"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox boxB 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   20
      ToolTipText     =   "Border Color (Blue Component)"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox boxG 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      MaxLength       =   8
      TabIndex        =   19
      ToolTipText     =   "Border Color (Green Component)"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox boxR 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   18
      ToolTipText     =   "Border Color (Red Component)"
      Top             =   4920
      Width           =   735
   End
   Begin VB.ComboBox listFilter 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   15
      ToolTipText     =   "Pixel color interpolation mode"
      Top             =   3960
      Width           =   2175
   End
   Begin VB.ComboBox listV 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   12
      ToolTipText     =   "Addressing mode along Y axis"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ComboBox listU 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Addressing mode along X axis"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox listPixel 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Pixel format"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox boxHeight 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   4
      ToolTipText     =   "Image height"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox boxWidth 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   2
      ToolTipText     =   "Image width"
      Top             =   600
      Width           =   615
   End
   Begin VB.Label infoA 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   25
      Top             =   5280
      Width           =   105
   End
   Begin VB.Label infoB 
      AutoSize        =   -1  'True
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   24
      Top             =   5280
      Width           =   90
   End
   Begin VB.Label infoG 
      AutoSize        =   -1  'True
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   23
      Top             =   4920
      Width           =   105
   End
   Begin VB.Label infoR 
      AutoSize        =   -1  'True
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   22
      Top             =   4920
      Width           =   105
   End
   Begin VB.Label infoBorder 
      AutoSize        =   -1  'True
      Caption         =   "&Border:"
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
      Left            =   480
      TabIndex        =   17
      Top             =   4920
      Width           =   540
   End
   Begin VB.Label infoMisc 
      AutoSize        =   -1  'True
      Caption         =   "Misc. Properties"
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
      TabIndex        =   16
      Top             =   4560
      Width           =   1350
   End
   Begin VB.Label infoFilter 
      AutoSize        =   -1  'True
      Caption         =   "&Filter:"
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
      Left            =   480
      TabIndex        =   14
      Top             =   3960
      Width           =   420
   End
   Begin VB.Label infoFetching 
      AutoSize        =   -1  'True
      Caption         =   "Fetching Samples"
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
      TabIndex        =   13
      Top             =   3600
      Width           =   1485
   End
   Begin VB.Label infoV 
      AutoSize        =   -1  'True
      Caption         =   "&V:"
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
      Left            =   480
      TabIndex        =   11
      Top             =   3000
      Width           =   150
   End
   Begin VB.Label infoU 
      AutoSize        =   -1  'True
      Caption         =   "&U:"
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
      Left            =   480
      TabIndex        =   9
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label infoAddressing 
      AutoSize        =   -1  'True
      Caption         =   "Addressing Mode"
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
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label infoPixel 
      AutoSize        =   -1  'True
      Caption         =   "&Pixel:"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   390
   End
   Begin VB.Label infoData 
      AutoSize        =   -1  'True
      Caption         =   "Data Storage"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label infoHeight 
      AutoSize        =   -1  'True
      Caption         =   "&Height:"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   525
   End
   Begin VB.Label infoWidth 
      AutoSize        =   -1  'True
      Caption         =   "&Width:"
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
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
   Begin VB.Label infoResolution 
      AutoSize        =   -1  'True
      Caption         =   "Resolution"
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
      Width           =   900
   End
End
Attribute VB_Name = "windowFrameBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' image dimensions
Public p_width As Long
Public p_height As Long

' pixel format
Public p_pixel As Long

' u, v addressing mode
Public p_u As Long
Public p_v As Long

' sampling filter
Public p_filter As Long

' border color
Public p_r As Single
Public p_g As Single
Public p_b As Single
Public p_a As Single


Public result As Long  ' accept or discard changes

Private freeze As Long ' freeze controls



'
' refresh all controls
'
Public Sub update()

  freeze = 1 ' lock them first

  ' image dimensions
  boxWidth.Text = p_width & vbNullString
  boxHeight.Text = p_height & vbNullString

  ' pixel format
  listPixel.ListIndex = clamp1i(p_pixel, 0, listPixel.ListCount - 1)

  ' u, v addressing mode
  listU.ListIndex = clamp1i(p_u, 0, listU.ListCount - 1)
  listV.ListIndex = clamp1i(p_v, 0, listV.ListCount - 1)

  ' sampling filter
  listFilter.ListIndex = clamp1i(p_filter, 0, listFilter.ListCount - 1)

  ' border color
  boxR.Text = format(p_r, "0.000") & vbNullString
  boxG.Text = format(p_g, "0.000") & vbNullString
  boxB.Text = format(p_b, "0.000") & vbNullString
  boxA.Text = format(p_a, "0.000") & vbNullString

  freeze = 0 ' unlock controls

End Sub



Private Sub boxA_Change()

  If (freeze <> 0) Then Exit Sub
  p_a = Val(boxA.Text)

End Sub



Private Sub boxB_Change()

  If (freeze <> 0) Then Exit Sub
  p_b = Val(boxB.Text)

End Sub



Private Sub boxG_Change()

  If (freeze <> 0) Then Exit Sub
  p_g = Val(boxG.Text)

End Sub



Private Sub boxHeight_Change()

  If (freeze <> 0) Then Exit Sub
  p_height = Val(boxHeight.Text)

End Sub



Private Sub boxR_Change()

  If (freeze <> 0) Then Exit Sub
  p_r = Val(boxR.Text)

End Sub



Private Sub boxWidth_Change()

  If (freeze <> 0) Then Exit Sub
  p_width = Val(boxWidth.Text)

End Sub



'
' discard changes
'
Private Sub buttonCancel_Click()

  result = 0
  Hide

End Sub



'
' accept changes
'
Private Sub buttonOk_Click()

  result = 1
  Hide

End Sub



'
' initialize controls
'
Private Sub Form_Load()

  freeze = 1 ' lock controls

  ' pixel format
  With listPixel
    .Clear
    .AddItem "RGBA: 32 Bit"
    .AddItem "RGBA: 128 Bit, FP"
  End With

  ' u addressing mode
  With listU
    .Clear
    .AddItem "Border"
    .AddItem "Clamp"
    .AddItem "Wrap"
    .AddItem "Mirror"
  End With

  ' v addressing mode
  With listV
    .Clear
    .AddItem "Border"
    .AddItem "Clamp"
    .AddItem "Wrap"
    .AddItem "Mirror"
  End With

  ' sampling filter
  With listFilter
    .Clear
    .AddItem "Nearest Neighbour"
    .AddItem "Bilinear"
  End With

  freeze = 0 ' unlock controls

End Sub



Private Sub listFilter_Click()

  If (freeze <> 0) Then Exit Sub
  p_filter = listFilter.ListIndex

End Sub



Private Sub listPixel_Click()

  If (freeze <> 0) Then Exit Sub
  p_pixel = listPixel.ListIndex

End Sub



Private Sub listU_Click()

  If (freeze <> 0) Then Exit Sub
  p_u = listU.ListIndex

End Sub



Private Sub listV_Click()

  If (freeze <> 0) Then Exit Sub
  p_v = listV.ListIndex

End Sub
