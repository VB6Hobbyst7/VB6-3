VERSION 5.00
Begin VB.Form windowNodeRemapChannels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox boxMax0 
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
      TabIndex        =   38
      ToolTipText     =   "Remap To"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox boxMin0 
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
      TabIndex        =   37
      ToolTipText     =   "Remap From"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CheckBox checkCos0 
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
      TabIndex        =   36
      ToolTipText     =   "Smooth Cosine Interpolation"
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox listSource0 
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
      TabIndex        =   35
      ToolTipText     =   "Source Channel"
      Top             =   1560
      Width           =   855
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
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   29
      ToolTipText     =   "Fill Color (Red Component)"
      Top             =   600
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
      Left            =   2760
      MaxLength       =   8
      TabIndex        =   28
      ToolTipText     =   "Fill Color (Green Component)"
      Top             =   600
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
      Left            =   3960
      MaxLength       =   8
      TabIndex        =   27
      ToolTipText     =   "Fill Color (Blue Component)"
      Top             =   600
      Width           =   735
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
      Left            =   5160
      MaxLength       =   8
      TabIndex        =   26
      ToolTipText     =   "Fill Color (Alpha Component / Transparency)"
      Top             =   600
      Width           =   735
   End
   Begin VB.ComboBox listSource3 
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
      TabIndex        =   22
      ToolTipText     =   "Source Channel"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CheckBox checkCos3 
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
      TabIndex        =   21
      ToolTipText     =   "Smooth Cosine Interpolation"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox boxMin3 
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
      TabIndex        =   20
      ToolTipText     =   "Remap From"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox boxMax3 
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
      TabIndex        =   19
      ToolTipText     =   "Remap To"
      Top             =   3000
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
      TabIndex        =   15
      ToolTipText     =   "Source Channel"
      Top             =   2040
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
      TabIndex        =   14
      ToolTipText     =   "Smooth Cosine Interpolation"
      Top             =   2040
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
      TabIndex        =   13
      ToolTipText     =   "Remap From"
      Top             =   2040
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
      TabIndex        =   12
      ToolTipText     =   "Remap To"
      Top             =   2040
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
      TabIndex        =   11
      ToolTipText     =   "Remap To"
      Top             =   2520
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
      TabIndex        =   10
      ToolTipText     =   "Remap From"
      Top             =   2520
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
      TabIndex        =   8
      ToolTipText     =   "Smooth Cosine Interpolation"
      Top             =   2520
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
      TabIndex        =   5
      ToolTipText     =   "Source Channel"
      Top             =   2520
      Width           =   855
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
      Top             =   3600
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
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label infoMax0 
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
      TabIndex        =   41
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label infoMin0 
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
      TabIndex        =   40
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label infoSocket0 
      AutoSize        =   -1  'True
      Caption         =   "Red:"
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
      TabIndex        =   39
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label infoDefault 
      AutoSize        =   -1  'True
      Caption         =   "Default:"
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
      TabIndex        =   34
      Top             =   600
      Width           =   585
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
      Left            =   1320
      TabIndex        =   33
      Top             =   600
      Width           =   105
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
      Left            =   2520
      TabIndex        =   32
      Top             =   600
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
      Left            =   3720
      TabIndex        =   31
      Top             =   600
      Width           =   90
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
      Left            =   4920
      TabIndex        =   30
      Top             =   600
      Width           =   105
   End
   Begin VB.Label infoSocket3 
      AutoSize        =   -1  'True
      Caption         =   "Alpha:"
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
      TabIndex        =   25
      Top             =   3000
      Width           =   465
   End
   Begin VB.Label infoMin3 
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
      TabIndex        =   24
      Top             =   3000
      Width           =   300
   End
   Begin VB.Label infoMax3 
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
      TabIndex        =   23
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label infoSocket1 
      AutoSize        =   -1  'True
      Caption         =   "Green:"
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
      TabIndex        =   18
      Top             =   2040
      Width           =   495
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
      TabIndex        =   17
      Top             =   2040
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
      TabIndex        =   16
      Top             =   2040
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
      TabIndex        =   9
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
      TabIndex        =   7
      Top             =   2520
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
      TabIndex        =   6
      Top             =   2520
      Width           =   300
   End
   Begin VB.Label infoSocket2 
      AutoSize        =   -1  'True
      Caption         =   "Blue:"
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
      Top             =   2520
      Width           =   360
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
      Top             =   3720
      Width           =   1785
   End
End
Attribute VB_Name = "windowNodeRemapChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' default color
Public p_r As Single
Public p_g As Single
Public p_b As Single
Public p_a As Single

' optional socket 0
Public p_src0 As Long
Public p_min0 As Single
Public p_max0 As Single
Public p_cos0 As Long

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

' optional socket 3
Public p_src3 As Long
Public p_min3 As Single
Public p_max3 As Single
Public p_cos3 As Long


Public result As Long  ' accept or discard changes

Private freeze As Long ' freeze controls



'
' refresh all controls
'
Public Sub update()

  freeze = 1 ' lock them first


  ' default color
  boxR.Text = format(p_r, "0.000") & vbNullString
  boxG.Text = format(p_g, "0.000") & vbNullString
  boxB.Text = format(p_b, "0.000") & vbNullString
  boxA.Text = format(p_a, "0.000") & vbNullString
  
  ' optional socket 0
  listSource0.ListIndex = clamp1i(p_src0, 0, listSource1.ListCount - 1)
  boxMin0.Text = format(p_min0, "0.000") & vbNullString
  boxMax0.Text = format(p_max0, "0.000") & vbNullString
  checkCos0.Value = p_cos0

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

  ' optional socket 3
  listSource3.ListIndex = clamp1i(p_src3, 0, listSource3.ListCount - 1)
  boxMin3.Text = format(p_min3, "0.000") & vbNullString
  boxMax3.Text = format(p_max3, "0.000") & vbNullString
  checkCos3.Value = p_cos3

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



Private Sub boxR_Change()

  If (freeze <> 0) Then Exit Sub
  p_r = Val(boxR.Text)

End Sub



Private Sub boxMax0_Change()

  If (freeze <> 0) Then Exit Sub
  p_max0 = Val(boxMax0.Text)

End Sub



Private Sub boxMax1_Change()

  If (freeze <> 0) Then Exit Sub
  p_max1 = Val(boxMax1.Text)

End Sub



Private Sub boxMax2_Change()

  If (freeze <> 0) Then Exit Sub
  p_max2 = Val(boxMax2.Text)

End Sub



Private Sub boxMax3_Change()

  If (freeze <> 0) Then Exit Sub
  p_max3 = Val(boxMax3.Text)

End Sub



Private Sub boxMin0_Change()

  If (freeze <> 0) Then Exit Sub
  p_min0 = Val(boxMin0.Text)

End Sub



Private Sub boxMin1_Change()

  If (freeze <> 0) Then Exit Sub
  p_min1 = Val(boxMin1.Text)

End Sub



Private Sub boxMin2_Change()

  If (freeze <> 0) Then Exit Sub
  p_min2 = Val(boxMin2.Text)

End Sub



Private Sub boxMin3_Change()

  If (freeze <> 0) Then Exit Sub
  p_min3 = Val(boxMin3.Text)

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



Private Sub checkCos0_Click()

  If (freeze <> 0) Then Exit Sub
  p_cos0 = checkCos0.Value

End Sub



Private Sub checkCos1_Click()

  If (freeze <> 0) Then Exit Sub
  p_cos1 = checkCos1.Value

End Sub



Private Sub checkCos2_Click()

  If (freeze <> 0) Then Exit Sub
  p_cos2 = checkCos2.Value

End Sub



Private Sub checkCos3_Click()

  If (freeze <> 0) Then Exit Sub
  p_cos3 = checkCos3.Value

End Sub


'
' initialize controls
'
Private Sub Form_Load()

  freeze = 1 ' lock controls

  ' source 0 selector
  With listSource0
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

  ' source 3 selector
  With listSource3
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



Private Sub listSource0_Click()

  If (freeze <> 0) Then Exit Sub
  p_src0 = listSource0.ListIndex

End Sub



Private Sub listSource1_Click()

  If (freeze <> 0) Then Exit Sub
  p_src1 = listSource1.ListIndex

End Sub



Private Sub listSource2_Click()

  If (freeze <> 0) Then Exit Sub
  p_src2 = listSource2.ListIndex

End Sub



Private Sub listSource3_Click()

  If (freeze <> 0) Then Exit Sub
  p_src3 = listSource3.ListIndex

End Sub

