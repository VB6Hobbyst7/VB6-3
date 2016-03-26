VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNodeBoxBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' frame buffer object
Public buffer As classFrameBuffer


' kernel
Private kernel As Single

' direction
Private direction As Long

' optional socket 1
Private src1 As Long
Private min1 As Single
Private max1 As Single
Private cos1 As Long



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
    Case Is = 1: get_name = "kernel"

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
    Case Is = 1: get_type = 1

    Case Else: get_type = 1 ' optional socket
  End Select

End Function



'
' return number of input sockets
'
Public Function get_inputs() As Long

  get_inputs = 2 ' (5 sockets is max for 50% zoom)

End Function



'
' return node title
'
Public Function get_title() As String

  get_title = "Box Blur"

End Function



'
' open up settings window
'
Public Sub show_options()

  ' load properties dialog
  Dim dialog As New windowNodeBoxBlur: Load dialog
  With dialog


  ' submit data into dialog

  ' kernel
  .p_kernel = kernel

  ' direction
  .p_direction = direction
  
  ' optional socket 1
  .p_src1 = src1
  .p_min1 = min1
  .p_max1 = max1
  .p_cos1 = cos1


  ' show up dialog and wait until it's hidden
  .update
  .Caption = get_title
  .Show vbModal


  ' read back modified data from dialog
  If .result <> 0 Then

    ' kernel
    kernel = clamp1f(.p_kernel, 1, 32)

    ' direction
    direction = clamp1i(.p_direction, 0, 2)

    ' optional socket 1
    src1 = clamp1i(.p_src1, 0, 7)
    min1 = clamp1f(.p_min1, 1, 32)
    max1 = clamp1f(.p_max1, 1, 32)
    cos1 = clamp1i(.p_cos1, 0, 1)

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

  ' current pixel coords
  Dim x As Long
  Dim y As Long

  ' process image
  y = 0
  Do While y < buffer.get_height

    ' process line
    x = 0
    Do While x < buffer.get_width

      
      ' pixel coordinates
      Dim u As Single: u = CSng(x)
      Dim v As Single: v = CSng(y)


      ' kernel
      Dim k As Single: k = fetch_optional(socket(1), u, v, src1, min1, max1, cos1, kernel)


      ' compute normalizer and box coords

      Dim n As Single

      Dim x0 As Single
      Dim y0 As Single
      Dim x1 As Single
      Dim y1 As Single

      ' choose direction
      Select Case (direction)

        Case Is = 0:  ' u, v
          n = k * 2 + 1: n = 1 / (n * n)
          x0 = x - k
          y0 = y - k
          x1 = x + k
          y1 = y + k

        Case Is = 1:  ' u
          n = k * 2 + 1: n = 1 / n
          x0 = x - k
          y0 = y
          x1 = x + k
          y1 = y

        Case Is = 2:  ' v
          n = k * 2 + 1: n = 1 / n
          x0 = x
          y0 = y - k
          x1 = x
          y1 = y + k

      End Select


      ' reset accumulator
      Dim color As float4: color = vector4f(0, 0, 0, 0)

      ' v pass
      Dim dy As Single: dy = y0
      Do While (dy <= y1)

        ' h pass
        Dim dx As Single: dx = x0
        Do While (dx <= x1)

          ' accumulate sampled pixels from socket 0
          color = add4f(color, fetch_critical(socket(0), dx, dy))

          dx = dx + 1
        Loop

        dy = dy + 1
      Loop

      ' normalize accumulator
      color = clamp4f(scale4f(color, n), 0, 1)
      
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
  Put #file, , kernel
  Put #file, , direction
  Put #file, , src1
  Put #file, , min1
  Put #file, , max1
  Put #file, , cos1


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
  Get #file, , kernel: kernel = clamp1f(kernel, 1, 32)
  Get #file, , direction: direction = clamp1i(direction, 0, 2)
  Get #file, , src1: src1 = clamp1i(src1, 0, 7)
  Get #file, , min1: min1 = clamp1f(min1, 1, 32)
  Get #file, , max1: max1 = clamp1f(max1, 1, 32)
  Get #file, , cos1: cos1 = clamp1i(cos1, 0, 1)


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
  kernel = 2 ' 2 px
  direction = 0 ' u and v
  src1 = 1 ' average
  min1 = 0 ' from 1
  max1 = 32 ' to 32
  cos1 = 0 ' smooth off


End Sub



'
' shutdown
'
Private Sub Class_Terminate()

  ' free frame buffer object
  Set buffer = Nothing

End Sub
