VERSION 5.00
Begin VB.UserControl ProgressIndicator 
   Alignable       =   -1  'True
   BackColor       =   &H8000000D&
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   ScaleHeight     =   18
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   232
   Begin VB.PictureBox Indicator 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1305
      TabIndex        =   1
      Top             =   15
      Width           =   1305
   End
   Begin VB.PictureBox Background 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   3000
      TabIndex        =   2
      Top             =   15
      Width           =   3000
   End
   Begin VB.Line FinalShade 
      BorderColor     =   &H80000014&
      X1              =   232
      X2              =   232
      Y1              =   0
      Y2              =   18
   End
   Begin VB.Label Percents 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   165
      Left            =   3120
      TabIndex        =   0
      Top             =   45
      Width           =   180
   End
   Begin VB.Line ShadeBottom 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   232
      Y1              =   17
      Y2              =   17
   End
   Begin VB.Line ShadeRight 
      BorderColor     =   &H80000014&
      X1              =   201
      X2              =   201
      Y1              =   0
      Y2              =   17.333
   End
   Begin VB.Line ShadeLeft 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   0
      Y1              =   17.333
      Y2              =   0
   End
   Begin VB.Line ShadeTop 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   232
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "ProgressIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'just to keep code clear
Option Explicit
Option Base 0

'fast "pset"
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'tga loader class
Private TgaLoader As TgaFile

'x,y coords
Private X As Integer
Private Y As Integer

'pixel
Private Color As ARGB32bit

'scale factor
Private Factor As Single

'startup
Private Sub UserControl_Initialize()
  PI_Percent 100
End Sub

'keep the same size
Private Sub UserControl_Resize()
  Height = 270
  Width = 3480
End Sub

'load images
Public Sub PI_LoadIcons(Back As String, Fore As String)
  'create tga class
  Set TgaLoader = New TgaFile
  'load background picture
  If TgaLoader.LoadTga(Back) Then
    'paint it
    For Y = 0 To TgaLoader.Height Step 1
      For X = 0 To TgaLoader.Width Step 1
        Color = TgaLoader.GetPixel(X, Y)
        SetPixelV Background.hdc, X - 1, Y - 1, RGB(Color.R, Color.G, Color.B)
      Next X
    Next Y
    Background.Refresh
  End If
  'load foreground picture
  If TgaLoader.LoadTga(Fore) Then
    'paint it
    For Y = 0 To TgaLoader.Height Step 1
      For X = 0 To TgaLoader.Width Step 1
        Color = TgaLoader.GetPixel(X, Y)
        SetPixelV Indicator.hdc, X - 1, Y - 1, RGB(Color.R, Color.G, Color.B)
      Next X
    Next Y
    Indicator.Refresh
  End If
  'delete tga class
  TgaLoader.Destroy
  Set TgaLoader = Nothing
End Sub

Public Sub PI_Percent(NewPercent As Single)
  'range check
  If NewPercent < 0 Then NewPercent = 0
  If NewPercent > 100 Then NewPercent = 100
  'set scale
  Factor = Background.Width / 100
  'update
  Indicator.Width = Factor * NewPercent
  Percents.Caption = Int(NewPercent) & "%"
End Sub
