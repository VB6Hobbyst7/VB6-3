VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNodeLuminosityContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' frame buffer object
Public buffer As classFrameBuffer


' lum / contrast
Private luminosity As Single
Private contrast As Single

' alpha channel
Private alpha As Long

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
    Case Is = 1: get_name = "luminosity"
    Case Is = 2: get_name = "contrast"

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

  get_title = "Lum. / Contrast"

End Function



'
' open up settings window
'
Public Sub show_options()

  ' load properties dialog
  Dim dialog As New windowNodeLuminosityContrast: Load dialog
  With dialog


  ' submit data into dialog

  ' lum / contrast
  .p_luminosity = luminosity
  .p_contrast = contrast

  ' alpha channel
  .p_alpha = alpha

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


  ' show up dialog and wait until it's hidden
  .update
  .Caption = get_title
  .Show vbModal


  ' read back modified data from dialog
  If .result <> 0 Then

    ' lum / contrast
    luminosity = clamp1f(.p_luminosity, -10, 10)
    contrast = clamp1f(.p_contrast, 0, 10)

    ' alpha channel
    alpha = clamp1i(.p_alpha, 0, 1)

    ' optional socket 1
    src1 = clamp1i(.p_src1, 0, 7)
    min1 = clamp1f(.p_min1, -10, 10)
    max1 = clamp1f(.p_max1, -10, 10)
    cos1 = clamp1i(.p_cos1, 0, 1)

    ' optional socket 2
    src2 = clamp1i(.p_src2, 0, 7)
    min2 = clamp1f(.p_min2, 0, 10)
    max2 = clamp1f(.p_max2, 0, 10)
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


      ' luminosity
      Dim l As Single: l = fetch_optional(socket(1), u, v, src1, min1, max1, cos1, luminosity)
      
      ' contrast
      Dim c As Single: c = fetch_optional(socket(2), u, v, src2, min2, max2, cos2, contrast)


      ' sample pixel
      Dim color As float4: color = fetch_critical(socket(0), u, v)

      
      ' shift luminosity and apply contrast
      Dim a As Single: a = color.w ' remember alpha
      color = add4fv(scale4f(sub4fv(color, 0.5), c), 0.5 + l)
      If (alpha <> 0) Then color.w = a ' restore alpha


      ' write pixel
      color = clamp4f(color, 0, 1)
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
  Put #file, , luminosity
  Put #file, , contrast
  Put #file, , alpha
  Put #file, , src1
  Put #file, , min1
  Put #file, , max1
  Put #file, , cos1
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
  Get #file, , luminosity: luminosity = clamp1f(luminosity, -10, 10)
  Get #file, , contrast: contrast = clamp1f(contrast, 0, 10)
  Get #file, , alpha: alpha = clamp1i(alpha, 0, 1)
  Get #file, , src1: src1 = clamp1i(src1, 0, 7)
  Get #file, , min1: min1 = clamp1f(min1, -10, 10)
  Get #file, , max1: max1 = clamp1f(max1, -10, 10)
  Get #file, , cos1: cos1 = clamp1i(cos1, 0, 1)
  Get #file, , src2: src2 = clamp1i(src2, 0, 7)
  Get #file, , min2: min2 = clamp1f(min2, 0, 10)
  Get #file, , max2: max2 = clamp1f(max2, 0, 10)
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
  luminosity = 0 ' +0%
  contrast = 1 ' 100%
  alpha = 0 ' preserve alpha ch.
  src1 = 1 ' average
  min1 = -1 ' from -1
  max1 = 1 ' to 1
  cos1 = 0 ' smooth off
  src2 = 1 ' average
  min2 = 0 ' from 0
  max2 = 2 ' to 2
  cos2 = 0 ' smooth off


End Sub



'
' shutdown
'
Private Sub Class_Terminate()

  ' free frame buffer object
  Set buffer = Nothing

End Sub
