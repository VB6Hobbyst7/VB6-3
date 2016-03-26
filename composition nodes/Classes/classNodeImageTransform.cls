VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNodeImageTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' frame buffer object
Public buffer As classFrameBuffer


' properties
Private rotate As Single
Private scale_x As Single
Private scale_y As Single
Private pan_x As Single
Private pan_y As Single
Private flip_h As Long
Private flip_v As Long



'
' return existance of the output pin
'
Public Function get_output() As Long

  get_output = 1

End Function



'
' return input socket id
'
Public Function get_name(ByVal index As Long) As String

  ' choose socket
  Select Case (index)

    Case Is = 0: get_name = vbNullString

    Case Else: get_name = vbNullString ' unnamed socket
  End Select

End Function



'
' return input socket type
'
Public Function get_type(ByVal index As Long) As Long

  ' choose socket
  Select Case (index)

    Case Is = 0: get_type = 0

    Case Else: get_type = 1 ' optional socket
  End Select

End Function



'
' return number of input sockets
'
Public Function get_inputs() As Long

  get_inputs = 1 ' (5 sockets is max for 50% zoom)

End Function



'
' return node title
'
Public Function get_title() As String

  get_title = "Image Transform"

End Function



'
' open up settings window
'
Public Sub show_options()

  ' load properties dialog
  Dim dialog As New windowNodeImageTransform: Load dialog
  With dialog


  ' submit data into dialog

  ' rotation, scale and pan
  .p_pan_x = pan_x
  .p_pan_y = pan_y
  .p_scale_x = scale_x
  .p_scale_y = scale_y
  .p_rotate = rotate
  .p_flip_h = flip_h
  .p_flip_v = flip_v


  ' show up dialog and wait until it's hidden
  .update
  .Caption = get_title
  .Show vbModal


  ' read back modified data from dialog
  If .result <> 0 Then

    ' rotation, scale and pan
    pan_x = clamp1i(.p_pan_x, -1024, 1024)
    pan_y = clamp1i(.p_pan_y, -1024, 1024)
    scale_x = clamp1f(.p_scale_x, 0.01, 100)
    scale_y = clamp1f(.p_scale_y, 0.01, 100)
    rotate = clamp1f(.p_rotate, -360, 360)
    flip_h = clamp1i(.p_flip_h, 0, 1)
    flip_v = clamp1i(.p_flip_v, 0, 1)

  End If


  ' unload window
  End With
  Unload dialog: Set dialog = Nothing

End Sub



'
' render self
'
Public Function render(ByRef socket() As classFrameBuffer) As Long

  If (socket(0) Is Nothing) Then render = 0: Exit Function ' check socket 0

  ' input 0: image size
  Dim sx As Single: sx = CSng(socket(0).get_width)
  Dim sy As Single: sy = CSng(socket(0).get_height)
  
  ' input 0: center point
  Dim cx As Single: cx = sx * 0.5
  Dim cy As Single: cy = sy * 0.5
  
  ' rotation angle in radians
  Dim ra As Single: ra = rotate * pi1f / 180
  Dim rc As Single: rc = Cos(ra)
  Dim rs As Single: rs = Sin(ra)
  
  ' inverse scale
  Dim iu As Single: iu = 1 / scale_x
  Dim iv As Single: iv = 1 / scale_y

  ' current pixel coords
  Dim x As Long
  Dim y As Long

  ' process image
  y = 0
  Do While y < buffer.get_height

    ' process line
    x = 0
    Do While x < buffer.get_width


      ' transform pixel coordinates
      Dim u As Single: u = CSng(x)
      Dim v As Single: v = CSng(y)
      
      ' flip
      If (flip_h <> 0) Then u = sx - u
      If (flip_v <> 0) Then v = sy - v
      
      ' pan
      u = u - pan_x
      v = v - pan_y
      
      ' scale
      u = u * iu
      v = v * iv
      
      ' rotate around center
      u = u - cx: v = v - cy
      Dim u0 As Single: u0 = u * rc - v * rs
      Dim v0 As Single: v0 = u * rs + v * rc
      u = u0 + cx: v = v0 + cy


      ' sample pixel
      Dim color As float4: color = fetch_critical(socket(0), u, v)

      ' write pixel
      With color: buffer.store x, y, .x, .y, .z, .w: End With


      x = x + 1 ' next pixel in this line
    Loop

    If windowRender.cancel <> 0 Then Exit Do            ' leave rendering loop (if required)
    windowRender.progress = (1 / buffer.get_height) * y ' update progress
    DoEvents                                            ' process all events

    y = y + 1 ' next line
  Loop

  ' success
  render = 1

End Function



'
' store properties to file
'
Public Function file_write(ByVal file As Long) As Long

  ' error handler
  On Error Resume Next

  ' store frame buffer settings
  If (buffer.file_write(file) = 0) Then file_write = 0: Exit Function


  ' write properties
  Put #file, , pan_x
  Put #file, , pan_y
  Put #file, , scale_x
  Put #file, , scale_y
  Put #file, , rotate
  Put #file, , flip_h
  Put #file, , flip_v


  ' i/o error check
  If (Err.number <> 0) Then Err.Clear: file_write = 0: Exit Function ' fail
  file_write = 1 ' success

End Function



'
' restore properties from file (repair them if possible)
'
Public Function file_read(ByVal file As Long) As Long

  ' error handler
  On Error Resume Next

  ' restore frame buffer settings
  If (buffer.file_read(file) = 0) Then file_read = 0: Exit Function


  ' read properties
  Get #file, , pan_x: pan_x = clamp1i(pan_x, -1024, 1024)
  Get #file, , pan_y: pan_y = clamp1i(pan_y, -1024, 1024)
  Get #file, , scale_x: scale_x = clamp1f(scale_x, 0.01, 100)
  Get #file, , scale_y: scale_y = clamp1f(scale_y, 0.01, 100)
  Get #file, , rotate: rotate = clamp1f(rotate, -360, 360)
  Get #file, , flip_h: flip_h = clamp1i(flip_h, 0, 1)
  Get #file, , flip_v: flip_v = clamp1i(flip_v, 0, 1)


  ' i/o error check
  If (Err.number <> 0) Then Err.Clear: file_read = 0: Exit Function ' fail
  file_read = 1 ' success

End Function



'
' startup
'
Private Sub Class_Initialize()

  ' create frame buffer object
  Set buffer = New classFrameBuffer


  ' initialize default defaults
  pan_x = 0
  pan_y = 0
  scale_x = 1
  scale_y = 1
  rotate = 0
  flip_h = 0
  flip_v = 0


End Sub



'
' shutdown
'
Private Sub Class_Terminate()

  ' free frame buffer object
  Set buffer = Nothing

End Sub
