VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNodeMixLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' frame buffer object
Public buffer As classFrameBuffer


' formula
Private mode As Long

' opacity
Private opacity As Single

' optional socket 2
Private src2 As Long
Private min2 As Single
Private max2 As Single
Private cos2 As Long



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

    Case Is = 0: get_name = "a"
    Case Is = 1: get_name = "b"
    Case Is = 2: get_name = "opacity"

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
    Case Is = 1: get_type = 0
    Case Is = 2: get_type = 1

    Case Else: get_type = 1 ' optional socket
  End Select

End Function



'
' return number of input sockets
'
Public Function get_inputs() As Long

  get_inputs = 3 ' (5 sockets is max for 50% zoom)

End Function



'
' return node title
'
Public Function get_title() As String

  get_title = "Mix Layers"

End Function



'
' open up settings window
'
Public Sub show_options()

  ' load properties dialog
  Dim dialog As New windowNodeMixLayers: Load dialog
  With dialog


  ' submit data into dialog

  ' formula
  .p_mode = mode
  
  ' opacity
  .p_opacity = opacity
  
  ' optional socket 2
  .p_src2 = src2
  .p_min2 = min2
  .p_max2 = max2
  .p_cos2 = cos2


  ' show up dialog and wait until it's hidden
  .update
  .Caption = get_title
  .Show vbModal


  ' read back modified data from dialog
  If .result <> 0 Then

    ' formula
    mode = clamp1i(.p_mode, 0, 4)

    ' opacity
    opacity = clamp1f(.p_opacity, -10, 10)

    ' optional socket 2
    src2 = clamp1i(.p_src2, 0, 7)
    min2 = clamp1f(.p_min2, -10, 10)
    max2 = clamp1f(.p_max2, -10, 10)
    cos2 = clamp1i(.p_cos2, 0, 1)

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
  If (socket(1) Is Nothing) Then render = 0: Exit Function ' check socket 1

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

      ' sample pixel from socket #0
      Dim r As Single, g As Single, b As Single, a As Single
      socket(0).fetch_sample u, v, r, g, b, a
      
      ' sample pixel from socket #1
      Dim r1 As Single, g1 As Single, b1 As Single, a1 As Single
      socket(1).fetch_sample u, v, r1, g1, b1, a1


      ' opacity
      Dim o As Single: o = fetch_optional(socket(2), u, v, src2, min2, max2, cos2, opacity)


      ' choose blend formula
      Select Case (mode)
      
        Case Is = 0 ' normal
          r = lerp1f(r, r1, o)
          g = lerp1f(g, g1, o)
          b = lerp1f(b, b1, o)
          a = lerp1f(a, a1, o)

        Case Is = 1 ' alpha
          r = lerp1f(r, r1, o * a1)
          g = lerp1f(g, g1, o * a1)
          b = lerp1f(b, b1, o * a1)
          a = a + a1

        Case Is = 2 ' add
          r = r + (r1 * o)
          g = g + (g1 * o)
          b = b + (b1 * o)
          a = a + (a1 * o)

        Case Is = 3 ' subtract
          r = r - (r1 * o)
          g = g - (g1 * o)
          b = b - (b1 * o)
          a = a - (a1 * o)

        Case Is = 4 ' multiply
          r = r * (r1 * o)
          g = g * (g1 * o)
          b = b * (b1 * o)
          a = a * (a1 * o)
      
      End Select


      ' write pixel
      buffer.store x, y, clamp1f(r, 0, 1), clamp1f(g, 0, 1), clamp1f(b, 0, 1), clamp1f(a, 0, 1)


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
  Put #file, , mode
  Put #file, , opacity
  Put #file, , src2
  Put #file, , min2
  Put #file, , max2
  Put #file, , cos2


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
  Get #file, , mode: mode = clamp1i(mode, 0, 4)
  Get #file, , opacity: opacity = clamp1f(opacity, -100, 100)
  Get #file, , src2: src2 = clamp1i(src2, 0, 7)
  Get #file, , min2: min2 = clamp1f(min2, -100, 100)
  Get #file, , max2: max2 = clamp1f(max2, -100, 100)
  Get #file, , cos2: cos2 = clamp1i(cos2, 0, 1)


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
  mode = 0 ' normal
  opacity = 0.5 ' 50%
  src2 = 1 ' average
  min2 = 0 ' from 0
  max2 = 1 ' to 1
  cos2 = 0 ' smooth off


End Sub



'
' shutdown
'
Private Sub Class_Terminate()

  ' free frame buffer object
  Set buffer = Nothing

End Sub
