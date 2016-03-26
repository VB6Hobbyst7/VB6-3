VERSION 5.00
Begin VB.Form windowNodeLuminosityContrast 
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
   Begin VB.CheckBox checkAlpha 
      Caption         =   "Alpha"
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
      TabIndex        =   23
      ToolTipText     =   "Preserve alpha channel"
      Top             =   600
      Width           =   735
   End
   Begin VB.ComboBox listSource1 
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
      TabIndex        =   19
      ToolTipText     =   "Source Channel"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CheckBox checkCos1 
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
      TabIndex        =   18
      ToolTipText     =   "Smooth Cosine Interpolation"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox boxMin1 
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
      TabIndex        =   17
      ToolTipText     =   "Remap From"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox boxMax1 
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
      TabIndex        =   16
      ToolTipText     =   "Remap To"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox boxLuminosity 
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
      TabIndex        =   15
      ToolTipText     =   "Luminosity Addition"
      Top             =   600
      Width           =   735
   End
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   11
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
      TabIndex        =   8
      ToolTipText     =   "Source Channel"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox boxContrast 
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
      Left            =   3480
      MaxLength       =   8
      TabIndex        =   6
      ToolTipText     =   "Contrast Multiplier"
      Top             =   600
      Width           =   735
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
   Begin VB.Label infoSocket1 
      AutoSize        =   -1  'True
      Caption         =   "Luminosity:"
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
      TabIndex        =   22
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label infoMin1 
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
      TabIndex        =   21
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label infoMax1 
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
      TabIndex        =   20
      Top             =   1560
      Width           =   360
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
      TabIndex        =   12
      Top             =   1200
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   2040
      Width           =   300
   End
   Begin VB.Label infoSocket2 
      AutoSize        =   -1  'True
      Caption         =   "Contrast:"
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
      TabIndex        =   7
      Top             =   2040
      Width           =   690
   End
   Begin VB.Label infoContrast 
      AutoSize        =   -1  'True
      Caption         =   "Contrast:"
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
      TabIndex        =   5
      Top             =   600
      Width           =   690
   End
   Begin VB.Label infoLuminosity 
      AutoSize        =   -1  'True
      Caption         =   "Luminosity:"
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
      Width           =   810
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
Attribute VB_Name = "windowNodeLuminosityContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' lum. / contrast
Public p_luminosity As Single
Public p_contrast As Single

' alpha channel
Public p_alpha As Long

' optional socket 1
Public p_src1 As Long
Public p_min1 As Single
Public p_max1 As Single
Public p_cos1 As Long

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

  ' lum. / contrast
  boxLuminosity.Text = format(p_luminosity, "0.000") & vbNullString
  boxContrast.Text = format(p_contrast, "0.000") & vbNullString

  ' alpha channel
  checkAlpha.Value = p_alpha
  
  ' optional socket 1
  listSource1.ListIndex = clamp1i(p_src1, 0, listSource1.ListCount - 1)
  boxMin1.Text = format(p_min1, "0.000") & vbNullString
  boxMax1.Text = format(p_max1, "0.000") & vbNullString
  checkCos1.Value = p_cos1

  ' optional socket 2
  listSource2.ListIndex = clamp1i(p_src2, 0, listSource2.ListCount - 1)
  boxMin2.Text = format(p_min2, "0.000") & vbNullString
  boxMax2.Text = format(p_max2, "0.000") & vbNullString
  checkCos2.Value = p_cos2

  freeze = 0 ' unlock controls

End Sub



Private Sub boxContrast_Change()

  If (freeze <> 0) Then Exit Sub
  p_contrast = Val(boxContrast.Text)

End Sub



Private Sub boxLuminosity_Change()

  If (freeze <> 0) Then Exit Sub
  p_luminosity = Val(boxLuminosity.Text)

End Sub



Private Sub boxMax1_Change()

  If (freeze <> 0) Then Exit Sub
  p_max1 = Val(boxMax1.Text)

End Sub



Private Sub boxMax2_Change()

  If (freeze <> 0) Then Exit Sub
  p_max2 = Val(boxMax2.Text)

End Sub



Private Sub boxMin1_Change()

  If (freeze <> 0) Then Exit Sub
  p_min1 = Val(boxMin1.Text)

End Sub



Private Sub boxMin2_Change()

  If (freeze <> 0) Then Exit Sub
  p_min2 = Val(boxMin2.Text)

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



Private Sub checkAlpha_Click()

  If (freeze <> 0) Then Exit Sub
  p_alpha = checkAlpha.Value

End Sub



Private Sub checkCos1_Click()

  If (freeze <> 0) Then Exit Sub
  p_cos1 = checkCos1.Value

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

  ' source 1 selector
  With listSource1
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



Private Sub listSource1_Click()

  If (freeze <> 0) Then Exit Sub
  p_src1 = listSource1.ListIndex

End Sub



Private Sub listSource2_Click()

  If (freeze <> 0) Then Exit Sub
  p_src2 = listSource2.ListIndex

End Sub
