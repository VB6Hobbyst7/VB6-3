VERSION 5.00
Begin VB.Form windowNodeMixLayers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox boxMax2 
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
      Left            =   4200
      MaxLength       =   8
      TabIndex        =   15
      ToolTipText     =   "Remap To"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox boxMin2 
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
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   14
      ToolTipText     =   "Remap From"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox checkCos2 
      Caption         =   "Smooth"
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
      Left            =   5160
      TabIndex        =   12
      ToolTipText     =   "Smooth Cosine Interpolation"
      Top             =   2040
      Width           =   855
   End
   Begin VB.ComboBox listSource2 
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "Source Channel"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox boxOpacity 
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
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   7
      ToolTipText     =   "Source[1] Multiplier"
      Top             =   1080
      Width           =   735
   End
   Begin VB.ComboBox listBlendMode 
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Blending Formula"
      Top             =   600
      Width           =   4455
   End
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
      Left            =   4920
      TabIndex        =   2
      Top             =   2640
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
      Left            =   3840
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label infoOptional 
      AutoSize        =   -1  'True
      Caption         =   "Optional Inputs"
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
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label infoMax2 
      AutoSize        =   -1  'True
      Caption         =   "Max:"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label infoMin2 
      AutoSize        =   -1  'True
      Caption         =   "Min:"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   2040
      Width           =   300
   End
   Begin VB.Label infoSocket2 
      AutoSize        =   -1  'True
      Caption         =   "Opacity:"
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
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label infoOpacity 
      AutoSize        =   -1  'True
      Caption         =   "Opacity:"
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
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label infoBlendMode 
      AutoSize        =   -1  'True
      Caption         =   "Blend Mode:"
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
      TabIndex        =   4
      Top             =   600
      Width           =   885
   End
   Begin VB.Label infoGeneral 
      AutoSize        =   -1  'True
      Caption         =   "General"
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
      TabIndex        =   3
      Top             =   240
      Width           =   660
   End
   Begin VB.Label infoFrameBufferOptions 
      AutoSize        =   -1  'True
      Caption         =   "Frame Buffer Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Click here for frame buffer configuration dialog"
      Top             =   2760
      Width           =   1785
   End
End
Attribute VB_Name = "windowNodeMixLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' formula
Public p_mode As Long

' opacity
Public p_opacity As Single

' optional socket 2
Public p_src2 As Long
Public p_min2 As Single
Public p_max2 As Single
Public p_cos2 As Long


Public result As Long  ' accept or discard changes

Private freeze As Long ' freeze controls



'
' refresh all controls
'
Public Sub update()

  freeze = 1 ' lock them first

  ' formula
  listBlendMode.ListIndex = clamp1i(p_mode, 0, listBlendMode.ListCount - 1)
  
  ' opacity
  boxOpacity.Text = format(p_opacity, "0.000") & vbNullString
  
  ' optional socket 2
  listSource2.ListIndex = clamp1i(p_src2, 0, listSource2.ListCount - 1)
  boxMin2.Text = format(p_min2, "0.000") & vbNullString
  boxMax2.Text = format(p_max2, "0.000") & vbNullString
  checkCos2.Value = p_cos2

  freeze = 0 ' unlock controls

End Sub



Private Sub boxMax2_Change()

  If (freeze <> 0) Then Exit Sub
  p_max2 = Val(boxMax2.Text)

End Sub



Private Sub boxMin2_Change()

  If (freeze <> 0) Then Exit Sub
  p_min2 = Val(boxMin2.Text)

End Sub



Private Sub boxOpacity_Change()

  If (freeze <> 0) Then Exit Sub
  p_opacity = Val(boxOpacity.Text)

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



Private Sub checkCos2_Click()

  If (freeze <> 0) Then Exit Sub
  p_cos2 = checkCos2.Value

End Sub



'
' initialize controls
'
Private Sub Form_Load()

  freeze = 1 ' lock controls

  ' blending mode
  With listBlendMode
    .Clear
    .AddItem "Normal:   Lerp(A, B, Op)"
    .AddItem "Alpha:    Lerp(A, B, B[a] * Op)"
    .AddItem "Add:      A + (B * Op)"
    .AddItem "Subtract: A - (B * Op)"
    .AddItem "Multiply: A * (B * Op)"
  End With

  ' source 2 selector
  With listSource2
    .Clear
    .AddItem "Lum"
    .AddItem "Avg"
    .AddItem "R"
    .AddItem "G"
    .AddItem "B"
    .AddItem "A"
    .AddItem "Max"
    .AddItem "Min"
  End With

  freeze = 0 ' unlock controls

End Sub



'
' configure frame buffer
'
Private Sub infoFrameBufferOptions_Click()

  ' is assumed that 'selection' points to a vailid node element
  node(selection).kernel.buffer.show_options

End Sub



Private Sub listBlendMode_click()

  If (freeze <> 0) Then Exit Sub
  p_mode = listBlendMode.ListIndex

End Sub



Private Sub listSource2_Click()

  If (freeze <> 0) Then Exit Sub
  p_src2 = listSource2.ListIndex

End Sub
