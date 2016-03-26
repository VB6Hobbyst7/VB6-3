VERSION 5.00
Begin VB.Form windowNodeGradient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox checkCos 
      Caption         =   "Cosine Interpolation"
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
      Left            =   1560
      TabIndex        =   40
      ToolTipText     =   "Enable smooth cosine color interpolation"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox boxA4 
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
      TabIndex        =   34
      ToolTipText     =   "Alpha Component / Transparency"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox boxB4 
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
      TabIndex        =   33
      ToolTipText     =   "Blue Component"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox boxG4 
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
      TabIndex        =   32
      ToolTipText     =   "Green Component"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox boxR4 
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
      TabIndex        =   31
      ToolTipText     =   "Red Component"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox boxA3 
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
      TabIndex        =   25
      ToolTipText     =   "Alpha Component / Transparency"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox boxB3 
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
      TabIndex        =   24
      ToolTipText     =   "Blue Component"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox boxG3 
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
      TabIndex        =   23
      ToolTipText     =   "Green Component"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox boxR3 
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
      TabIndex        =   22
      ToolTipText     =   "Red Component"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox boxR2 
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
      TabIndex        =   16
      ToolTipText     =   "Red Component"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox boxG2 
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
      TabIndex        =   15
      ToolTipText     =   "Green Component"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox boxB2 
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
      TabIndex        =   14
      ToolTipText     =   "Blue Component"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox boxA2 
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
      TabIndex        =   13
      ToolTipText     =   "Alpha Component / Transparency"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox boxA1 
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
      TabIndex        =   12
      ToolTipText     =   "Alpha Component / Transparency"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox boxB1 
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
      TabIndex        =   10
      ToolTipText     =   "Blue Component"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox boxG1 
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
      TabIndex        =   8
      ToolTipText     =   "Green Component"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox boxR1 
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
      TabIndex        =   6
      ToolTipText     =   "Red Component"
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
      Top             =   3000
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
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label infoA4 
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
      TabIndex        =   39
      Top             =   2040
      Width           =   105
   End
   Begin VB.Label infoB4 
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
      TabIndex        =   38
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label infoG4 
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
      TabIndex        =   37
      Top             =   2040
      Width           =   105
   End
   Begin VB.Label infoR4 
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
      TabIndex        =   36
      Top             =   2040
      Width           =   105
   End
   Begin VB.Label info4 
      AutoSize        =   -1  'True
      Caption         =   "Bottom, L:"
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
      TabIndex        =   35
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label infoA3 
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
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label infoB3 
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
      TabIndex        =   29
      Top             =   1560
      Width           =   90
   End
   Begin VB.Label infoG3 
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
      TabIndex        =   28
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label infoR3 
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
      TabIndex        =   27
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label info3 
      AutoSize        =   -1  'True
      Caption         =   "Bottom, R:"
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
      TabIndex        =   26
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label info2 
      AutoSize        =   -1  'True
      Caption         =   "Top, R:"
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
      TabIndex        =   21
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label infoR2 
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
      TabIndex        =   20
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label infoG2 
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
      TabIndex        =   19
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label infoB2 
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
      TabIndex        =   18
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label infoA2 
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
      TabIndex        =   17
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label infoA1 
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
      TabIndex        =   11
      Top             =   600
      Width           =   105
   End
   Begin VB.Label infoB1 
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
      TabIndex        =   9
      Top             =   600
      Width           =   90
   End
   Begin VB.Label infoG1 
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
      TabIndex        =   7
      Top             =   600
      Width           =   105
   End
   Begin VB.Label infoR1 
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
      TabIndex        =   5
      Top             =   600
      Width           =   105
   End
   Begin VB.Label info1 
      AutoSize        =   -1  'True
      Caption         =   "Top, L:"
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
      Width           =   510
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
      Top             =   3120
      Width           =   1785
   End
End
Attribute VB_Name = "windowNodeGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' color range
Public p_r1 As Single
Public p_g1 As Single
Public p_b1 As Single
Public p_a1 As Single
Public p_r2 As Single
Public p_g2 As Single
Public p_b2 As Single
Public p_a2 As Single
Public p_r3 As Single
Public p_g3 As Single
Public p_b3 As Single
Public p_a3 As Single
Public p_r4 As Single
Public p_g4 As Single
Public p_b4 As Single
Public p_a4 As Single

' smoothing
Public p_cos As Long


Public result As Long  ' accept or discard changes

Private freeze As Long ' freeze controls



'
' refresh all controls
'
Public Sub update()

  freeze = 1 ' lock them first

  ' color range
  boxR1.Text = format(p_r1, "0.000") & vbNullString
  boxG1.Text = format(p_g1, "0.000") & vbNullString
  boxB1.Text = format(p_b1, "0.000") & vbNullString
  boxA1.Text = format(p_a1, "0.000") & vbNullString
  boxR2.Text = format(p_r2, "0.000") & vbNullString
  boxG2.Text = format(p_g2, "0.000") & vbNullString
  boxB2.Text = format(p_b2, "0.000") & vbNullString
  boxA2.Text = format(p_a2, "0.000") & vbNullString
  boxR3.Text = format(p_r3, "0.000") & vbNullString
  boxG3.Text = format(p_g3, "0.000") & vbNullString
  boxB3.Text = format(p_b3, "0.000") & vbNullString
  boxA3.Text = format(p_a3, "0.000") & vbNullString
  boxR4.Text = format(p_r4, "0.000") & vbNullString
  boxG4.Text = format(p_g4, "0.000") & vbNullString
  boxB4.Text = format(p_b4, "0.000") & vbNullString
  boxA4.Text = format(p_a4, "0.000") & vbNullString

  ' smoothing
  checkCos.Value = p_cos

  freeze = 0 ' unlock controls

End Sub



Private Sub boxA1_Change()

  If (freeze <> 0) Then Exit Sub
  p_a1 = Val(boxA1.Text)

End Sub



Private Sub boxB1_Change()

  If (freeze <> 0) Then Exit Sub
  p_b1 = Val(boxB1.Text)

End Sub



Private Sub boxG1_Change()

  If (freeze <> 0) Then Exit Sub
  p_g1 = Val(boxG1.Text)

End Sub



Private Sub boxR1_Change()

  If (freeze <> 0) Then Exit Sub
  p_r1 = Val(boxR1.Text)

End Sub



Private Sub boxA2_Change()

  If (freeze <> 0) Then Exit Sub
  p_a2 = Val(boxA2.Text)

End Sub



Private Sub boxB2_Change()

  If (freeze <> 0) Then Exit Sub
  p_b2 = Val(boxB2.Text)

End Sub



Private Sub boxG2_Change()

  If (freeze <> 0) Then Exit Sub
  p_g2 = Val(boxG2.Text)

End Sub



Private Sub boxR2_Change()

  If (freeze <> 0) Then Exit Sub
  p_r2 = Val(boxR2.Text)

End Sub



Private Sub boxA3_Change()

  If (freeze <> 0) Then Exit Sub
  p_a3 = Val(boxA3.Text)

End Sub



Private Sub boxB3_Change()

  If (freeze <> 0) Then Exit Sub
  p_b3 = Val(boxB3.Text)

End Sub



Private Sub boxG3_Change()

  If (freeze <> 0) Then Exit Sub
  p_g3 = Val(boxG3.Text)

End Sub



Private Sub boxR3_Change()

  If (freeze <> 0) Then Exit Sub
  p_r3 = Val(boxR3.Text)

End Sub



Private Sub boxA4_Change()

  If (freeze <> 0) Then Exit Sub
  p_a4 = Val(boxA4.Text)

End Sub



Private Sub boxB4_Change()

  If (freeze <> 0) Then Exit Sub
  p_b4 = Val(boxB4.Text)

End Sub



Private Sub boxG4_Change()

  If (freeze <> 0) Then Exit Sub
  p_g4 = Val(boxG4.Text)

End Sub



Private Sub boxR4_Change()

  If (freeze <> 0) Then Exit Sub
  p_r4 = Val(boxR4.Text)

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



Private Sub checkCos_Click()

  If (freeze <> 0) Then Exit Sub
  p_cos = checkCos.Value

End Sub



'
' configure frame buffer
'
Private Sub infoFrameBufferOptions_Click()

  ' is assumed that 'selection' points to a vailid node element
  node(selection).kernel.buffer.show_options

End Sub
