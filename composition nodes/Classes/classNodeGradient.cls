VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNodeGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' frame buffer object
Public buffer As classFrameBuffer


' color range
Private color1 As float4
Private color2 As float4
Private color3 As float4
Private color4 As float4

' smothing
Private smooth As Long



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



    Case Else: get_name = vbNullString ' unnamed socket
  End Select

End Function



'
' return input socket type
'
Public Function get_type(ByVal index As Long) As Long

  ' choose socket
  Select Case (index)



    Case Else: get_type = 1 ' optional socket
  End Select

End Function



'
' return number of input sockets
'
Public Function get_inputs() As Long

  get_inputs = 0 ' (5 sockets is max for 50% zoom)

End Function



'
' return node title
'
Public Function get_title() As String

  get_title = "Gradient"

End Function



'
' open up settings window
'
Public Sub show_options()

  ' load properties dialog
  Dim dialog As New windowNodeGradient: Load dialog
  With dialog


  ' submit data into dialog

  ' color range
  .p_r1 = color1.x
  .p_g1 = color1.y
  .p_b1 = color1.z
  .p_a1 = color1.w
  .p_r2 = color2.x
  .p_g2 = color2.y
  .p_b2 = color2.z
  .p_a2 = color2.w
  .p_r3 = color3.x
  .p_g3 = color3.y
  .p_b3 = color3.z
  .p_a3 = color3.w
  .p_r4 = color4.x
  .p_g4 = color4.y
  .p_b4 = color4.z
  .p_a4 = color4.w

  ' smoothing
  .p_cos = smooth


  ' show up dialog and wait until it's hidden
  .update
  .Caption = get_title
  .Show vbModal


  ' read back modified data from dialog
  If .result <> 0 Then

    ' color range
    color1 = clamp4f(vector4f(.p_r1, .p_g1, .p_b1, .p_a1), 0, 1)
    color2 = clamp4f(vector4f(.p_r2, .p_g2, .p_b2, .p_a2), 0, 1)
    color3 = clamp4f(vector4f(.p_r3, .p_g3, .p_b3, .p_a3), 0, 1)
    color4 = clamp4f(vector4f(.p_r4, .p_g4, .p_b4, .p_a4), 0, 1)

    ' smoothing
    smooth = clamp1i(.p_cos, 0, 1)

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

    ' interpolate line
    Dim c0 As float4
    Dim c1 As float4
    If (smooth <> 0) Then
      c0 = cosine4f(color1, color4, y / buffer.get_height)
      c1 = cosine4f(color2, color3, y / buffer.get_height)
    Else
      c0 = lerp4f(color1, color4, y / buffer.get_height)
      c1 = lerp4f(color2, color3, y / buffer.get_height)
    End If

    ' process line
    x = 0
    Do While x < buffer.get_width

      ' interpolate color in line
      Dim color As float4
      If (smooth <> 0) Then
        color = cosine4f(c0, c1, x / buffer.get_width)
      Else
        color = lerp4f(c0, c1, x / buffer.get_width)
      End If


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
  Put #file, , color1
  Put #file, , color2
  Put #file, , color3
  Put #file, , color4
  Put #file, , smooth


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
  Get #file, , color1: color1 = clamp4f(color1, 0, 1)
  Get #file, , color2: color2 = clamp4f(color2, 0, 1)
  Get #file, , color3: color3 = clamp4f(color3, 0, 1)
  Get #file, , color4: color4 = clamp4f(color4, 0, 1)
  Get #file, , smooth: smooth = clamp1i(smooth, 0, 1)


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
  color1 = vector4f(0, 0, 0, 0) ' black
  color2 = vector4f(1, 1, 1, 1) ' white
  color3 = vector4f(1, 1, 1, 1) ' white
  color4 = vector4f(0, 0, 0, 0) ' black
  smooth = 0 ' cosine off


End Sub



'
' shutdown
'
Private Sub Class_Terminate()

  ' free frame buffer object
  Set buffer = Nothing

End Sub
