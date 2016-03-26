VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNodeShiftHSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' require variable declaration
Option Explicit


' frame buffer object
Public buffer As classFrameBuffer


' hsl
Private hue As Single
Private saturation As Single
Private lightness As Single

' colorize flag
Private colorize As Long

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

    Case Is = 0: get_name = vbNullString
    Case Is = 1: get_name = "hue"
    Case Is = 2: get_name = "saturation"
    Case Is = 3: get_name = "lightness"

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

  get_title = "Shift H.S.L."

End Function



'
' open up settings window
'
Public Sub show_options()

  ' load properties dialog
  Dim dialog As New windowNodeShiftHSL: Load dialog
  With dialog


  ' submit data into dialog

  ' hsl
  .p_hue = hue
  .p_saturation = saturation
  .p_lightness = lightness

  ' colorize flag
  .p_colorize = colorize

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

    ' hsl
    hue = clamp1f(.p_hue, -360, 360)
    saturation = clamp1f(.p_saturation, -10, 10)
    lightness = clamp1f(.p_lightness, -10, 10)

    ' colorize flag
    colorize = clamp1i(.p_colorize, 0, 1)

    ' optional socket 1
    src1 = clamp1i(.p_src1, 0, 7)
    min1 = clamp1f(.p_min1, -360, 360)
    max1 = clamp1f(.p_max1, -360, 360)
    cos1 = clamp1i(.p_cos1, 0, 1)

    ' optional socket 2
    src2 = clamp1i(.p_src2, 0, 7)
    min2 = clamp1f(.p_min2, -10, 10)
    max2 = clamp1f(.p_max2, -10, 10)
    cos2 = clamp1i(.p_cos2, 0, 1)

    ' optional socket 3
    src3 = clamp1i(.p_src3, 0, 7)
    min3 = clamp1f(.p_min3, -10, 10)
    max3 = clamp1f(.p_max3, -10, 10)
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


      ' hue
      Dim h_ As Single: h_ = fetch_optional(socket(1), u, v, src1, min1, max1, cos1, hue)
      
      ' saturation
      Dim s_ As Single: s_ = fetch_optional(socket(2), u, v, src2, min2, max2, cos2, saturation)
      
      ' lightness
      Dim l_ As Single: l_ = fetch_optional(socket(3), u, v, src3, min3, max3, cos3, lightness)


      ' sample pixel from socket #0
      Dim samp As float4: samp = fetch_critical(socket(0), u, v)



      ' (Original Code - http://labs.adobe.com/wiki/index.php/HSL)
      ' ----------------------------------------------------------


      ' Convert sampled pixel from RGB space to HSL space.

      Dim sampMin As Single
      Dim sampMax As Single
      Dim sampDiff As Single
      Dim sampSum As Single
      Dim sampH As Single
      Dim sampS As Single
      Dim sampL As Single

      sampMin = min2f(samp.x, samp.y)
      sampMin = min2f(sampMin, samp.z)
      sampMax = max2f(samp.x, samp.y)
      sampMax = max2f(sampMax, samp.z)
      sampDiff = sampMax - sampMin
      sampSum = sampMax + sampMin
      sampL = sampSum * 0.5

      If (sampMin = sampMax) Then
        sampH = 0
      Else
        If (sampMax = samp.x) Then
          sampH = (60 * ((samp.y - samp.z) / sampDiff)) Mod 360
        Else
          If (sampMax = samp.y) Then
            sampH = 60 * ((samp.z - samp.x) / sampDiff) + 120
          Else
            If (sampMax = samp.z) Then
              sampH = 60 * ((samp.x - samp.y) / sampDiff) + 240
            Else
              sampH = 0
            End If
          End If
        End If
      End If
        
      If (sampMin = sampMax) Then
        sampS = 0
      Else
        If (sampL > 0.5) Then
          sampS = sampDiff / (2 - sampSum)
        Else
          sampS = sampDiff / sampSum
        End If
      End If



      ' Transform the sampled HSL values by the amounts specified
      ' by the hue, saturation, and lightness parameters.
        
      Dim outH As Single
      Dim outS As Single
      Dim outL As Single

      outH = wrap1f(sampH - h_, 360)
      If (colorize <> 0) Then
        outS = clamp1f(sampS + s_, 0, 1)
      Else
        outS = clamp1f(sampS * s_, 0, 1)
      End If
      outL = clamp1f(sampL + l_, 0, 1)



      ' Convert the transformed HSL values back to RGB space.

      Dim q As Single
      Dim p As Single
      Dim h As Single

      If (outL < 0.5) Then
        q = outL * (1 + outS)
      Else
        q = outL + outS - outL * outS
      End If

      p = 2 * outL - q
      h = outH / 360

      Const oneOverThree As Single = 1 / 3
      Const twoOverThree As Single = 2 / 3
      Const oneOverSix As Single = 1 / 6
      Dim t As float4: t = vector4f(h + oneOverThree, h, h - oneOverThree, 0)

      If (t.x < 0) Then
        t.x = t.x + 1
      Else
        If (t.x > 1) Then t.x = t.x - 1
      End If

      If (t.y < 0) Then
        t.y = t.y + 1
      Else
        If (t.y > 1) Then t.y = t.y - 1
      End If

      If (t.z < 0) Then
        t.z = t.z + 1
      Else
        If (t.z > 1) Then t.z = t.z - 1
      End If

      Dim c As float4: c = vector4f(0, 0, 0, samp.w)

      If (t.x < oneOverSix) Then
        c.x = p + (q - p) * 6 * t.x
      Else
        If (t.x >= oneOverSix And t.x < 0.5) Then
          c.x = q
        Else
          If (t.x >= 0.5 And t.x < twoOverThree) Then
              c.x = p + (q - p) * 6 * (twoOverThree - t.x)
          Else
              c.x = p
          End If
        End If
      End If

      If (t.y < oneOverSix) Then
        c.y = p + (q - p) * 6 * t.y
      Else
        If (t.y >= oneOverSix And t.y < 0.5) Then
          c.y = q
        Else
          If (t.y >= 0.5 And t.y < twoOverThree) Then
            c.y = p + (q - p) * 6 * (twoOverThree - t.y)
          Else
            c.y = p
          End If
        End If
      End If

      If (t.z < oneOverSix) Then
        c.z = p + (q - p) * 6 * t.z
      Else
        If (t.z >= oneOverSix And t.z < 0.5) Then
          c.z = q
        Else
          If (t.z >= 0.5 And t.z < twoOverThree) Then
            c.z = p + (q - p) * 6 * (twoOverThree - t.z)
          Else
            c.z = p
          End If
        End If
      End If



      ' write pixel
      c = clamp4f(c, 0, 1)
      With c: buffer.store x, y, .x, .y, .z, .w: End With


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
  Put #file, , hue
  Put #file, , saturation
  Put #file, , lightness
  Put #file, , colorize
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
  Get #file, , hue: hue = clamp1f(hue, -360, 360)
  Get #file, , saturation: saturation = clamp1f(saturation, -10, 10)
  Get #file, , lightness: lightness = clamp1f(lightness, -10, 10)
  Get #file, , colorize: colorize = clamp1i(colorize, 0, 1)
  Get #file, , src1: src1 = clamp1i(src1, 0, 7)
  Get #file, , min1: min1 = clamp1f(min1, -360, 360)
  Get #file, , max1: max1 = clamp1f(max1, -360, 360)
  Get #file, , cos1: cos1 = clamp1i(cos1, 0, 1)
  Get #file, , src2: src2 = clamp1i(src2, 0, 7)
  Get #file, , min2: min2 = clamp1f(min2, -10, 10)
  Get #file, , max2: max2 = clamp1f(max2, -10, 10)
  Get #file, , cos2: cos2 = clamp1i(cos2, 0, 1)
  Get #file, , src3: src3 = clamp1i(src3, 0, 7)
  Get #file, , min3: min3 = clamp1f(min3, -10, 10)
  Get #file, , max3: max3 = clamp1f(max3, -10, 10)
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
  hue = 0 ' +0%
  saturation = 0 ' +0%
  lightness = 0 ' +0%
  colorize = 1 ' on
  src1 = 1 ' average
  min1 = -180 ' from -180
  max1 = 180 ' to 180
  cos1 = 0 ' smooth off
  src2 = 1 ' average
  min2 = -1 ' from -1
  max2 = 1 ' to 1
  cos2 = 0 ' smooth off
  src3 = 1 ' average
  min3 = -1 ' from -1
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
