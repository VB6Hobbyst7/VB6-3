VERSION 5.00
Begin VB.Form windowNodeImageTransform 
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
   Begin VB.TextBox boxPanX 
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
      MaxLength       =   7
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox boxPanY 
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
      Left            =   3000
      MaxLength       =   7
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox boxScaleX 
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
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox boxScaleY 
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
      Left            =   3000
      MaxLength       =   7
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox boxRotate 
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
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1560
      Width           =   855
   End
   Begin VB.CheckBox checkH 
      Caption         =   "Horizontal"
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
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox checkV 
      Caption         =   "Vertical"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   2040
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
   Begin VB.Label infoPan 
      AutoSize        =   -1  'True
      Caption         =   "Pan:"
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
      TabIndex        =   23
      Top             =   600
      Width           =   330
   End
   Begin VB.Label infoPanX 
      AutoSize        =   -1  'True
      Caption         =   "X:"
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
      Left            =   1080
      TabIndex        =   22
      Top             =   600
      Width           =   150
   End
   Begin VB.Label infoPanY 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
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
      Left            =   2760
      TabIndex        =   21
      Top             =   600
      Width           =   150
   End
   Begin VB.Label infoPx2 
      AutoSize        =   -1  'True
      Caption         =   "px."
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
      Left            =   2280
      TabIndex        =   20
      Top             =   600
      Width           =   240
   End
   Begin VB.Label infoPx3 
      AutoSize        =   -1  'True
      Caption         =   "px."
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
      Left            =   3960
      TabIndex        =   19
      Top             =   600
      Width           =   240
   End
   Begin VB.Label infoScale 
      AutoSize        =   -1  'True
      Caption         =   "Scale:"
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
      Top             =   1080
      Width           =   435
   End
   Begin VB.Label infoScaleX 
      AutoSize        =   -1  'True
      Caption         =   "X:"
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
      Left            =   1080
      TabIndex        =   17
      Top             =   1080
      Width           =   150
   End
   Begin VB.Label infoScaleY 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
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
      Left            =   2760
      TabIndex        =   16
      Top             =   1080
      Width           =   150
   End
   Begin VB.Label infoRotate 
      AutoSize        =   -1  'True
      Caption         =   "Rotate:"
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
      TabIndex        =   15
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label infoDeg 
      AutoSize        =   -1  'True
      Caption         =   "deg."
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
      Left            =   2280
      TabIndex        =   14
      Top             =   1560
      Width           =   330
   End
   Begin VB.Label infoP0 
      AutoSize        =   -1  'True
      Caption         =   "%"
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
      Left            =   2280
      TabIndex        =   13
      Top             =   1080
      Width           =   165
   End
   Begin VB.Label infoP1 
      AutoSize        =   -1  'True
      Caption         =   "%"
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
      Left            =   3960
      TabIndex        =   12
      Top             =   1080
      Width           =   165
   End
   Begin VB.Label infoFlip 
      AutoSize        =   -1  'True
      Caption         =   "Flip:"
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
      Top             =   2040
      Width           =   300
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
Attribute VB_Name = "windowNodeImageTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' rotation, pan, zoom and flip flags
Public p_rotate As Single
Public p_scale_x As Single
Public p_scale_y As Single
Public p_pan_x As Single
Public p_pan_y As Single
Public p_flip_h As Long
Public p_flip_v As Long


Public result As Long  ' accept or discard changes

Private freeze As Long ' freeze controls



'
' refresh all controls
'
Public Sub update()

  freeze = 1 ' lock them first

  ' rotation, pan, zoom and flip flags
  boxPanX.Text = p_pan_x & vbNullString
  boxPanY.Text = p_pan_y & vbNullString
  boxScaleX.Text = format(p_scale_x * 100, "0.00") & vbNullString
  boxScaleY.Text = format(p_scale_y * 100, "0.00") & vbNullString
  boxRotate.Text = format(p_rotate, "0.00") & vbNullString
  checkH.Value = p_flip_h
  checkV.Value = p_flip_v

  freeze = 0 ' unlock controls

End Sub



Private Sub boxPanX_Change()

  If freeze <> 0 Then Exit Sub
  p_pan_x = Val(boxPanX.Text)

End Sub



Private Sub boxPanY_Change()

  If freeze <> 0 Then Exit Sub
  p_pan_y = Val(boxPanY.Text)

End Sub



Private Sub boxRotate_Change()

  If freeze <> 0 Then Exit Sub
  p_rotate = Val(boxRotate.Text)

End Sub



Private Sub boxScaleX_Change()

  If freeze <> 0 Then Exit Sub
  p_scale_x = Val(boxScaleX.Text) / 100

End Sub



Private Sub boxScaleY_Change()

  If freeze <> 0 Then Exit Sub
  p_scale_y = Val(boxScaleY.Text) / 100

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



Private Sub checkH_Click()

  If freeze <> 0 Then Exit Sub
  p_flip_h = checkH.Value

End Sub



Private Sub checkV_Click()

  If freeze <> 0 Then Exit Sub
  p_flip_v = checkV.Value

End Sub



'
' configure frame buffer
'
Private Sub infoFrameBufferOptions_Click()

  ' is assumed that 'selection' points to a vailid node element
  node(selection).kernel.buffer.show_options

End Sub
