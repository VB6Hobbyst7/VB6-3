VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNodeRemapChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' frame buffer object
Public buffer As classFrameBuffer


' default color
Private color As float4

' optional socket 0
Private src0 As Long
Private min0 As Single
Private max0 As Single
Private cos0 As Long

' optional socket 1
Private src1 As Long
Private min1 As Single
Private max1 As Single
Private cos1 As Long

' optional socket 2
Private src2 As Long
Private min2 As Single
Private max2 As Single
Private cos2 As Long

' optional socket 3
Private src3 As Long
Private min3 As Single
Private max3 As Single
Private cos3 As Long



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

    Case Is = 0: get_name = "red"
    Case Is = 1: get_name = "green"
    Case Is = 2: get_name = "blue"
    Case Is = 3: get_name = "alpha"

    Case Else: get_name = vbNullString ' unnamed socket
  End Select

End Function



'
' return input socket type
'
Public Function get_type(ByVal index As Long) As Long

  ' choose socket
  Select Case (index)

    Case Is = 0: get_type = 1
    Case Is = 1: get_type = 1
    Case Is = 2: get_type = 1
    Case Is = 3: get_type = 1

    Case Else: get_type = 1 ' optional socket
  End Select

End Function



'
' return number of input sockets
'
Public Function get_inputs() As Long

  get_inputs = 4 ' (5 sockets is max for 50% zoom)

End Function



'
' return node title
'
Public Function get_title() As String

  get_title = "Remap Channels"

End Function



'
' open up settings window
'
Public Sub show_options()

  ' load properties dialog
  Dim dialog As New windowNodeRemapChannels: Load dialog
  With dialog


  ' submit data into dialog


  ' default color
  .p_r = color.x
  .p_g = color.y
  .p_b = color.z
  .p_a = color.w

  ' optional socket 0
  .p_src0 = src0
  .p_min0 = min0
  .p_max0 = max0
  .p_cos0 = cos0

  ' optional socket 1
  .p_src1 = src1
  .p_min1 = min1
  .p_max1 = max1
  .p_cos1 = cos1
  
  ' optional socket 2
  .p_src2 = src2
  .p_min2 = min2
  .p_max2 = max2
  .p_cos2 = cos2
  
  ' optional socket 3
  .p_src3 = src3
  .p_min3 = min3
  .p_max3 = max3
  .p_cos3 = cos3


  ' show up dialog and wait until it's hidden
  .update
  .Caption = get_title
  .Show vbModal


  ' read back modified data from dialog
  If .result <> 0 Then

    ' default color
    color = clamp4f(vector4f(.p_r, .p_g, .p_b, .p_a), 0, 1)

    ' optional socket 0
    src0 = clamp1i(.p_src0, 0, 7)
    min0 = clamp1f(.p_min0, 0, 1)
    max0 = clamp1f(.p_max0, 0, 1)
    cos0 = clamp1i(.p_cos0, 0, 1)

    ' optional socket 1
    src1 = clamp1i(.p_src1, 0, 7)
    min1 = clamp1f(.p_min1, 0, 1)
    max1 = clamp1f(.p_max1, 0, 1)
    cos1 = clamp1i(.p_cos1, 0, 1)

    ' optional socket 2
    src2 = clamp1i(.p_src2, 0, 7)
    min2 = clamp1f(.p_min2, 0, 1)
    max2 = clamp1f(.p_max2, 0, 1)
    cos2 = clamp1i(.p_cos2, 0, 1)

    ' optional socket 3
    src3 = clamp1i(.p_src3, 0, 7)
    min3 = clamp1f(.p_min3, 0, 1)
    max3 = clamp1f(.p_max3, 0, 1)
    cos3 = clamp1i(.p_cos3, 0, 1)

  End If


  ' unload window
  End With
  Unload dialog: Set dialog = Nothing

End Sub



'
' render self
'
Public Function render(ByRef socket() As classFrameBuffer) As Long

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


      ' channels
      Dim r As Single: r = fetch_optional(socket(0), u, v, src0, min0, max0, cos0, color.x)
      Dim g As Single: g = fetch_optional(socket(1), u, v, src1, min1, max1, cos1, color.y)
      Dim b As Single: b = fetch_optional(socket(2), u, v, src2, min2, max2, cos2, color.z)
      Dim a As Single: a = fetch_optional(socket(3), u, v, src3, min3, max3, cos3, color.w)


      ' write pixel
      buffer.store x, y, r, g, b, a


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
  Put #file, , color
  Put #file, , src0
  Put #file, , min0
  Put #file, , max0
  Put #file, , cos0
  Put #file, , src1
  Put #file, , min1
  Put #file, , max1
  Put #file, , cos1
  Put #file, , src2
  Put #file, , min2
  Put #file, , max2
  Put #file, , cos3
  Put #file, , src3
  Put #file, , min3
  Put #file, , max3
  Put #file, , cos3


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
  Get #file, , color: color = clamp4f(color, 0, 1)
  Get #file, , src0: src1 = clamp1i(src1, 0, 7)
  Get #file, , min0: min1 = clamp1f(min1, 0, 1)
  Get #file, , max0: max1 = clamp1f(max1, 0, 1)
  Get #file, , cos0: cos1 = clamp1i(cos1, 0, 1)
  Get #file, , src1: src1 = clamp1i(src1, 0, 7)
  Get #file, , min1: min1 = clamp1f(min1, 0, 1)
  Get #file, , max1: max1 = clamp1f(max1, 0, 1)
  Get #file, , cos1: cos1 = clamp1i(cos1, 0, 1)
  Get #file, , src2: src2 = clamp1i(src2, 0, 7)
  Get #file, , min2: min2 = clamp1f(min2, 0, 1)
  Get #file, , max2: max2 = clamp1f(max2, 0, 1)
  Get #file, , cos2: cos2 = clamp1i(cos2, 0, 1)
  Get #file, , src3: src3 = clamp1i(src3, 0, 7)
  Get #file, , min3: min3 = clamp1f(min3, 0, 1)
  Get #file, , max3: max3 = clamp1f(max3, 0, 1)
  Get #file, , cos3: cos3 = clamp1i(cos3, 0, 1)


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
  color = vector4f(0, 0, 0, 0) ' black
  src0 = 2 ' red
  min0 = 0 ' from 0
  max0 = 1 ' to 1
  cos0 = 0 ' smooth off
  src1 = 3 ' green
  min1 = 0 ' from 0
  max1 = 1 ' to 1
  cos1 = 0 ' smooth off
  src2 = 4 ' blue
  min2 = 0 ' from 0
  max2 = 1 ' to 1
  cos2 = 0 ' smooth off
  src3 = 5 ' alpha
  min3 = 0 ' from 0
  max3 = 1 ' to 1
  cos3 = 0 ' smooth off


End Sub



'
' shutdown
'
Private Sub Class_Terminate()

  ' free frame buffer object
  Set buffer = Nothing

End Sub
