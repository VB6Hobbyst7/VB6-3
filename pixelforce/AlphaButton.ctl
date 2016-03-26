VERSION 5.00
Begin VB.UserControl AlphaButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   ClipControls    =   0   'False
   ScaleHeight     =   93
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   303
   Begin VB.PictureBox ActiveArea 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   615
      Left            =   360
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Timer Tick 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3360
      Top             =   480
   End
   Begin VB.Label Message 
      AutoSize        =   -1  'True
      Caption         =   "Alpha Button"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   930
   End
End
Attribute VB_Name = "AlphaButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'just to keep code clear
Option Explicit
Option Base 0

'standard api rect type
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'standard api bitmapinfoheader type
Private Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

'standard api bitmapinfo type
Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
End Type

'standard api pointapi type
Private Type POINTAPI
  X As Long
  Y As Long
End Type

'draw bitmap
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
'fast "point"
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'current mouse position
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'get window rect (actually used to get position of usercontrol)
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'temp variables
Private A As Integer
Private R As Integer
Private G As Integer
Private B As Integer
Private X As Integer
Private Y As Integer

'pixelmap loader
Private IconLoader As TgaFile

'color map data
Private ColorArray() As ARGB32bit
Private ColorWidth As Integer
Private ColorHeight As Integer

'user control back color
Private BackColor As ARGB32bit

'result (visible) pixelmap
Private ImageArray() As ARGB32bit

'pixelmap descriptor
Private Bitmap As BITMAPINFO

'mouse currently in window rect (flag)
Private MouseIn As Boolean
'mouse button state flag
Private MousePush As Integer
'current mouse position
Private Mouse As POINTAPI
'current active-x control position
Private Button As RECT

'alpha amp
Private Fade As Integer

'temp variable (to store old alpha value)
Private Temp As Integer

'event similar to CommandButton_Click()
Public Event ABClick()
'event similat to PictureBox_MouseMove()
Public Event ABRollOver()

'default transparency level
Public AB_InactiveTransparency As Integer

'button pushed (for check-box buttons)
Public AB_Pushed As Boolean

'button disabled
Public AB_Disabled As Boolean

'icon load flag
Private iconOK As Boolean

'set tool tip for icon
Public Sub AB_Tooltip(Text As String)
  ActiveArea.ToolTipText = Text
End Sub

'set button caption
Public Sub AB_Caption(Text As String)
  Message.Caption = Text
  'update control size (autosize)
  UserControl_Resize
End Sub

'this will load tga file as icon
Public Function AB_LoadIcon(FileName As String) As Boolean
  'create tga loader class
  Set IconLoader = New TgaFile
  With IconLoader
    'try to load tga image
    iconOK = False
    AB_LoadIcon = .LoadTga(FileName)
    If AB_LoadIcon Then
      'allocate memory
      ReDim ColorArray(.Width, .Height)
      'set image dimensions
      ColorWidth = .Width
      ColorHeight = .Height
      'get pixelmap
      .GetBits ColorArray()
      'if tga file has no alpha channel, generate it
      If .AlphaBits = 0 Then
        'process every pixel
        For Y = 1 To .Height Step 1
          For X = 1 To .Width Step 1
            With ColorArray(X, Y)
              'alpha=(red+green+blue)/3
              .A = CByte(255 - (CLng(.R) + CLng(.G) + CLng(.B)) / 3)
            End With
          Next X
        Next Y
      End If
      'create output pixelmap
      ReDim ImageArray(ColorWidth - 1, ColorHeight - 1)
      'enable refresh timer
      Tick.Enabled = True
      'user control size depends on image dimensions
      iconOK = True
      UserControl_Resize
    End If
    'cleanup memory
    .Destroy
  End With
  'kill tga loader
  Set IconLoader = Nothing
End Function

'generate and paint pixelmap
Public Sub AB_RenderIcon()
  Temp = 0 'now it is used for grid, when button pushed
  'process every pixel
  For Y = 1 To ColorHeight Step 1
    For X = 1 To ColorWidth Step 1
      With BackColor
        'get alpha level
        A = CInt(CLng(ColorArray(X, Y).A) - CLng(Fade))
        'chech alpha range
        If A < 0 Then A = 0
        If A > 255 Then A = 255
        'toggle grid cell color flag
        If X < ColorWidth Then
          If Temp = -1 Then
            Temp = 1
          Else
            Temp = -1
          End If
        End If
        'button pushed
        If AB_Pushed Then
          'make per-pixel grid
          R = .R + 32 * Temp
          G = .G + 32 * Temp
          B = .B + 32 * Temp
        Else
          'set current bg color
          R = .R
          G = .G
          B = .B
        End If
        'check ranges
        CheckRGB
        'do alpha mix
        R = CByte(CLng(R) + (CLng(ColorArray(X, Y).R) - CLng(R)) / 255 * CLng(A))
        G = CByte(CLng(G) + (CLng(ColorArray(X, Y).G) - CLng(G)) / 255 * CLng(A))
        B = CByte(CLng(B) + (CLng(ColorArray(X, Y).B) - CLng(B)) / 255 * CLng(A))
        'check ranges
        CheckRGB
      End With
      'set result color
      With ImageArray(X - 1, ColorHeight - Y)
        'make gray color, if button disabled
        .A = CByte((CLng(R) + CLng(G) + CLng(B)) / 3)
        If Not AB_Disabled Then
          .R = R
          .G = G
          .B = B
        Else
          .R = .A
          .G = .A
          .B = .A
        End If
      End With
    Next X
  Next Y
  'set caption lebel settings, if we got caption
  If Len(Message.Caption) > 0 Then
    'bold fornt if pushed
    If AB_Pushed Then
      If Not Message.Font.Bold Then Message.Font.Bold = True
    Else
      If Message.Font.Bold Then Message.Font.Bold = False
    End If
    'underline, like hyper-link
    If MouseIn Then
      If Not Message.Font.Underline Then Message.Font.Underline = True
    Else
      If Message.Font.Underline Then Message.Font.Underline = False
    End If
    'resize control
    If Len(Message.Caption) > 0 Then
      If Width <> ColorWidth * Screen.TwipsPerPixelX + (Message.Width + 3) * Screen.TwipsPerPixelX Then Width = ColorWidth * Screen.TwipsPerPixelX + (Message.Width + 3) * Screen.TwipsPerPixelX
    Else
      If Width <> ColorWidth * Screen.TwipsPerPixelX Then Width = ColorWidth * Screen.TwipsPerPixelX
    End If
  End If
  'make button visible when icon not loaded
  If Not iconOK Then MouseIn = True
  'this will render 3d-box (outline)
  If (MouseIn Or AB_Pushed) And Not AB_Disabled Then
    Temp = MousePush
    If AB_Pushed Then MousePush = -1
    For X = 0 To ColorWidth - 1 Step 1
      With BackColor
        R = .R - 64 * MousePush
        G = .G - 64 * MousePush
        B = .B - 64 * MousePush
      End With
      With ImageArray(X, 0)
         CheckRGB
        .R = R
        .G = G
        .B = B
      End With
      With BackColor
        R = .R + 64 * MousePush
        G = .G + 64 * MousePush
        B = .B + 64 * MousePush
      End With
      With ImageArray(X, ColorHeight - 1)
         CheckRGB
        .R = R
        .G = G
        .B = B
      End With
    Next X
    For Y = 0 To ColorHeight - 1 Step 1
      With BackColor
        R = .R + 64 * MousePush
        G = .G + 64 * MousePush
        B = .B + 64 * MousePush
      End With
      With ImageArray(0, Y)
         CheckRGB
        .R = R
        .G = G
        .B = B
      End With
      With BackColor
        R = .R - 64 * MousePush
        G = .G - 64 * MousePush
        B = .B - 64 * MousePush
      End With
      With ImageArray(ColorWidth - 1, Y)
         CheckRGB
        .R = R
        .G = G
        .B = B
      End With
    Next Y
    MousePush = Temp
  End If
  'draw icon
  SetDIBits ActiveArea.hdc, ActiveArea.Image.Handle, 0, Bitmap.bmiHeader.biHeight, ImageArray(0, 0), Bitmap, 0
  'refresh image
  ActiveArea.Refresh
End Sub

'check r,g,b ranges
Private Function CheckRGB()
  If R < 0 Then R = 0
  If G < 0 Then G = 0
  If B < 0 Then B = 0
  If R > 255 Then R = 255
  If G > 255 Then G = 255
  If B > 255 Then B = 255
End Function

Private Sub Tick_Timer()
  'if button disabled - do not fade
  If AB_Disabled Then Tick.Enabled = False
  'get mouse
  GetCursorPos Mouse
  'get control
  GetWindowRect hwnd, Button
  With Button
    'mouse inside?
    If Mouse.X >= .Left And Mouse.X <= .Right And Mouse.Y >= .Top And Mouse.Y <= .Bottom Then
      MouseIn = True
      'fadein if not disabled
      If Fade > 0 And Not AB_Disabled Then
        'step 4
        Fade = Fade - 4
        If Fade <= 0 Then Fade = 0
      End If
    Else
      'mouse outside
      MouseIn = False
      'pushed?
      If AB_Pushed Then
        'remember initial alpha
        Temp = AB_InactiveTransparency
        'current_alpha=initial_alpha/2
        AB_InactiveTransparency = AB_InactiveTransparency / 2
        Fade = AB_InactiveTransparency
      End If
      'diabled? - current alpha level
      If AB_Disabled Then Fade = AB_InactiveTransparency
      'fade out is enabled
      If Fade <= AB_InactiveTransparency And Not AB_Disabled Then
        'step 4
        Fade = Fade + 4
        'fade complete?
        If Fade >= AB_InactiveTransparency Then
          'disable timer
          Fade = AB_InactiveTransparency
          Tick.Enabled = False
        End If
      End If
      'restore alpha
      If AB_Pushed Then AB_InactiveTransparency = Temp
    End If
  End With
  'render image
  AB_RenderIcon
End Sub

'startup
Private Sub UserControl_Initialize()
  'icon not loaded
  iconOK = False
  'no caption
  Message.Caption = vbNullString
  '28x28 button size by default
  ColorWidth = 28
  ColorHeight = 28
  '96 alpha level
  AB_InactiveTransparency = 96
  'current fade pass = current alpha
  Fade = AB_InactiveTransparency
  'mouse not inside control
  MouseIn = False
  'unpushed
  MousePush = 1
  'allocate memory for icon
  ReDim ColorArray(ColorWidth, ColorHeight)
  ReDim ImageArray(ColorWidth - 1, ColorHeight - 1)
  'get background color (for blending)
  With BackColor
    .A = 0
    .R = GetPixel(hdc, 0, 0) And 255
    .G = (GetPixel(hdc, 0, 0) And 65280) / 256
    .B = (GetPixel(hdc, 0, 0) And 16711680) / 65536
  End With
  'update everything
  UserControl_Resize
End Sub

'send all required events to usercontrol
Private Sub Message_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ActiveArea_MouseDown Button, Shift, X, Y
End Sub

Private Sub Message_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ActiveArea_MouseMove Button, Shift, X, Y
End Sub

Private Sub Message_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ActiveArea_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ActiveArea_MouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ActiveArea_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ActiveArea_MouseUp Button, Shift, X, Y
End Sub

Private Sub ActiveArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'if button enabled
  If Not AB_Disabled Then
    'we pushed it
    MousePush = -1
    'refresh
    Tick_Timer
  End If
End Sub

Private Sub ActiveArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'mouse moved inside control
  RaiseEvent ABRollOver
  'ok, we got it
  MouseIn = True
  'start refreshing timer (for fade in-out effects)
  Tick.Enabled = True
End Sub

Private Sub ActiveArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'mouse button released
  If Not AB_Disabled Then
    'UNpush the button
    MousePush = 1
    'refresh
    Tick_Timer
    'user clicked a button, raise event
    RaiseEvent ABClick
  End If
End Sub

'on resize control -> setup all elements
Private Sub UserControl_Resize()
  With Screen
    'control height is the same as pixelmap height
    Height = ColorHeight * .TwipsPerPixelY
    'icon render output picture resize
    ActiveArea.Move 0, 0, ColorWidth, ColorHeight
    'move caption lavel
    Message.Move ActiveArea.Width + 3, ScaleHeight / 2 - Message.Height / 2
    'caption set?
    If Len(Message.Caption) > 0 Then
      'resize control
      Width = ColorWidth * .TwipsPerPixelX + (Message.Width + 3) * .TwipsPerPixelX
    Else
      'resize control
      Width = ColorWidth * .TwipsPerPixelX
    End If
  End With
  'prepare bitmap header
  With Bitmap.bmiHeader
    .biSize = 40
    .biWidth = ColorWidth
    .biHeight = ColorHeight
    .biPlanes = 1
    .biBitCount = 32
    .biSizeImage = CLng(ColorWidth) * CLng(ColorHeight)
  End With
  'redraw button
  Tick_Timer
End Sub

'stop
Private Sub UserControl_Terminate()
  'memory cleanup
  Erase ColorArray()
  Erase ImageArray()
End Sub
