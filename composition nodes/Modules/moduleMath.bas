Attribute VB_Name = "moduleMath"

' require variable declaration
Option Explicit


' 4d vector (float32x4)
Public Type float4
  x As Single ' r
  y As Single ' g
  z As Single ' b
  w As Single ' a
End Type



' pi constant
Public Const pi1f As Single = 3.1415924



'
' optional input sampler
'
Public Function fetch_critical(ByRef frame As classFrameBuffer, ByVal u As Single, ByVal v As Single) As float4

  ' sample pixel from socket
  Dim r As Single, g As Single, b As Single, a As Single
  frame.fetch_sample u, v, r, g, b, a

  ' combine channels
  fetch_critical = vector4f(r, g, b, a)

End Function



'
' optional input sampler
'
Public Function fetch_optional(ByRef frame As classFrameBuffer, ByVal u As Single, ByVal v As Single, ByVal source As Long, ByVal min As Single, ByVal max As Single, ByVal smooth As Long, ByVal default As Single) As Single

  ' kernel
  If (Not (frame Is Nothing)) Then ' when socket is connected

    ' sample pixel from socket
    Dim r As Single, g As Single, b As Single, a As Single
    frame.fetch_sample u, v, r, g, b, a

    ' apply channel selector
    Dim f As Single: f = select_value(vector4f(r, g, b, a), source)

    ' remap value
    If (smooth <> 0) Then
      fetch_optional = lerp1f(min, max, f)
    Else
      fetch_optional = cosine1f(min, max, f)
    End If

  Else
    fetch_optional = default
  End If

End Function



'
' value selector
'
Public Function select_value(ByRef color As float4, ByVal how As Long)

  ' choose mode
  With color
    Select Case (how)
  
      ' luminance (r,g,b)
      Case Is = 0: select_value = .x * 0.299 + .y * 0.587 + .z * 0.114
  
      ' average (r,g,b)
      Case Is = 1: select_value = (.x + .y + .z) * 0.333333
  
      ' red
      Case Is = 2: select_value = .x
  
      ' green
      Case Is = 3: select_value = .y
  
      ' blue
      Case Is = 4: select_value = .z
  
      ' alpha
      Case Is = 5: select_value = .w
  
      ' max(r,g,b)
      Case Is = 6: select_value = max2f(max2f(.x, .y), color.z)
  
      ' min(r,g,b)
      Case Is = 7: select_value = min2f(min2f(.x, .y), color.z)

    End Select
  End With

  ' fit in 0...1
  select_value = clamp1f(select_value, 0, 1)

End Function



'
' convert long color into floating point color
'
Public Function rgba2fp(ByVal inp As Long) As float4

  ' extract channels
  With rgba2fp

    ' b,g and r
    .z = CSng(inp And &HFF&) / 255
    .y = CSng((inp And &HFF00&) \ &H100&) / 255
    .x = CSng((inp And &HFF0000) \ &H10000) / 255

    ' extracting alpha is a bit tricky - vb does not support Unsigned Longs :)
    Dim a As Long: a = inp And &HFF000000
    If (a < 0) Then
      .w = CSng((a Xor &H80000000) \ &H1000000 + 128) / 255
    Else
      .w = CSng(a \ &H1000000) / 255
    End If

  End With

End Function



'
' convert floating point color into long color
'
Public Function fp2rgba(ByRef inp As float4) As Long

  ' scale and clamp channels
  With inp
    Dim r As Byte: r = CByte(clamp1i(Int(.x * 255), 0, 255))
    Dim g As Byte: g = CByte(clamp1i(Int(.y * 255), 0, 255))
    Dim b As Byte: b = CByte(clamp1i(Int(.z * 255), 0, 255))
    Dim a As Byte: a = CByte(clamp1i(Int(.w * 255), 0, 255))
  End With

  ' pack them all into long color value
  If (a > 127) Then
    fp2rgba = ((a - 128) * &H1000000 Or &H80000000) Or (r * &H10000) Or (g * &H100&) Or b
  Else
    fp2rgba = (a * &H1000000) Or (r * &H10000) Or (g * &H100&) Or b
  End If

End Function



'
' wraps float value in range from 0 to 'range'
'
Public Function wrap1f(ByVal inp As Single, ByVal range As Single) As Single

  wrap1f = inp - CSng(Int(inp / range)) * range

End Function



'
' wraps integer value in range from 0 to 'range' (MOD operator does something different)
'
Public Function wrap1i(ByVal inp As Long, ByVal range As Long) As Long

  wrap1i = inp - Int(CSng(inp) / CSng(range)) * range

End Function



'
' mirrors an integer value in range from 0 to 'range'
'
Public Function mirror1i(ByVal inp As Long, ByVal range As Long) As Long

  Dim count As Long: count = Int(CSng(inp) / CSng(range))
  mirror1i = inp - count * range
  If (count Mod 2 <> 0) Then mirror1i = range - mirror1i - 1

End Function



'
' clamp float value in given range
'
Public Function clamp1f(ByVal inp As Single, ByVal min As Single, ByVal max As Single) As Single

  If (inp <= min) Then
    clamp1f = min   ' low bound
  Else

    If (inp >= max) Then
      clamp1f = max ' high bound
    Else

      clamp1f = inp ' no change

    End If

  End If

End Function



'
' clamp integer value in given range
'
Public Function clamp1i(ByVal inp As Long, ByVal min As Long, ByVal max As Long) As Long

  If (inp <= min) Then
    clamp1i = min   ' low bound
  Else

    If (inp >= max) Then
      clamp1i = max ' high bound
    Else

      clamp1i = inp ' no change

    End If

  End If

End Function



'
' linear interpolation
'
Public Function lerp1f(ByVal in1 As Single, ByVal in2 As Single, ByVal factor As Single) As Single

  lerp1f = in1 + (in2 - in1) * factor

End Function



'
' cosine interpolation
'
Public Function cosine1f(ByVal in1 As Single, ByVal in2 As Single, ByVal factor As Single) As Single

  cosine1f = in1 + (in2 - in1) * (1 - Cos(factor * pi1f)) * 0.5

End Function



'
' cubic interpolation
'
Public Function cubic1f(ByVal in1 As Single, ByVal in2 As Single, ByVal in3 As Single, ByVal in4 As Single, ByVal factor As Single) As Single

   Dim sq As Single: sq = factor * factor
   Dim a0 As Single: a0 = in4 - in3 - in1 + in2

   cubic1f = a0 * factor * sq + (in1 - in2 - a0) * sq + (in3 - in1) * factor + in2

End Function



'
' point inside circle test
'
Public Function point_vs_circle2d(ByVal u As Single, ByVal v As Single, ByVal x As Single, ByVal y As Single, ByVal r As Single) As Long

  ' perform test
  If (Sqr((x - u) ^ 2 + (y - v) ^ 2) <= r) Then
    point_vs_circle2d = 1 ' inside
  Else
    point_vs_circle2d = 0 ' outside
  End If

End Function



'
' point inside box test
'
Public Function point_vs_box2d(ByVal u As Single, ByVal v As Single, ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single) As Long

  ' perform test
  If (u >= x0 And v >= y0 And u <= x1 And v <= y1) Then
    point_vs_box2d = 1 ' inside
  Else
    point_vs_box2d = 0 ' outside
  End If

End Function



'
' box vs box intersection test
'
Public Function box_vs_box2d(ByVal u0 As Single, ByVal v0 As Single, ByVal u1 As Single, ByVal v1 As Single, ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single) As Long

  ' check if any of 4 points (box 1) inside box 2
  
  ' min, min
  box_vs_box2d = point_vs_box2d(u0, v0, x0, y0, x1, y1)
  If box_vs_box2d <> 0 Then Exit Function

  ' max, min
  box_vs_box2d = point_vs_box2d(u1, v0, x0, y0, x1, y1)
  If box_vs_box2d <> 0 Then Exit Function

  ' min, max
  box_vs_box2d = point_vs_box2d(u0, v1, x0, y0, x1, y1)
  If box_vs_box2d <> 0 Then Exit Function

  ' max, max
  box_vs_box2d = point_vs_box2d(u1, v1, x0, y0, x1, y1)
  If box_vs_box2d <> 0 Then Exit Function

  ' check x overlapping
  If u0 < x0 And u1 > x1 Then

    box_vs_box2d = 1
    If v0 >= y0 And v0 <= y1 Then Exit Function ' top line
    If v1 >= y0 And v1 <= y1 Then Exit Function ' bottom line

  End If

  ' check y overlapping
  If v0 < y0 And v1 > y1 Then

    If box_vs_box2d <> 0 Then Exit Function     ' dual (x,y) overlapping

    box_vs_box2d = 1
    If u0 >= x0 And u0 <= x1 Then Exit Function ' left line
    If u1 >= x0 And u1 <= x1 Then Exit Function ' right line

  End If

  box_vs_box2d = 0 ' no collision

End Function



'
' select min from 2 components
'
Public Function min2f(ByVal in1 As Single, ByVal in2 As Single) As Single

  If (in1 > in2) Then ' in2 less than in1 (in1 - discarded)
    min2f = in2
  Else ' in1 less or equal in2 (in2 - discarded)
    min2f = in1
  End If

End Function



'
' select max from 2 components
'
Public Function max2f(ByVal in1 As Single, ByVal in2 As Single) As Single

  If (in1 > in2) Then ' in2 less than in1 (in2 - discarded)
    max2f = in1
  Else ' in1 less or equal in2 (in1 - discarded)
    max2f = in2
  End If

End Function



'
' linear interpolation (4d)
'
Public Function lerp4f(ByRef in1 As float4, ByRef in2 As float4, ByVal factor As Single) As float4

  With lerp4f
    .x = in1.x + (in2.x - in1.x) * factor
    .y = in1.y + (in2.y - in1.y) * factor
    .z = in1.z + (in2.z - in1.z) * factor
    .w = in1.w + (in2.w - in1.w) * factor
  End With

End Function



'
' cosine interpolation (4d)
'
Public Function cosine4f(ByRef in1 As float4, ByRef in2 As float4, ByVal factor As Single) As float4

  Dim f As Single: f = (1 - Cos(factor * pi1f)) * 0.5

  With cosine4f
    .x = in1.x + (in2.x - in1.x) * f
    .y = in1.y + (in2.y - in1.y) * f
    .z = in1.z + (in2.z - in1.z) * f
    .w = in1.w + (in2.w - in1.w) * f
  End With

End Function



'
' create 4d-vector
'
Public Function vector4f(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal w As Single) As float4

  With vector4f
    .x = x
    .y = y
    .z = z
    .w = w
  End With

End Function



'
' clamp 4d-vector
'
Public Function clamp4f(ByRef v As float4, ByVal min As Single, ByVal max As Single) As float4

  With clamp4f

    ' x component
    If (v.x < min) Then
      .x = min
    Else
      If (v.x > max) Then
        .x = max
      Else
        .x = v.x
      End If
    End If

    ' y component
    If (v.y < min) Then
      .y = min
    Else
      If (v.y > max) Then
        .y = max
      Else
        .y = v.y
      End If
    End If

    ' z component
    If (v.z < min) Then
      .z = min
    Else
      If (v.z > max) Then
        .z = max
      Else
        .z = v.z
      End If
    End If

    ' w component
    If (v.w < min) Then
      .w = min
    Else
      If (v.w > max) Then
        .w = max
      Else
        .w = v.w
      End If
    End If

  End With

End Function



'
' add two 4d-vectors
'
Public Function add4f(ByRef in1 As float4, ByRef in2 As float4) As float4

  With add4f
    .x = in1.x + in2.x
    .y = in1.y + in2.y
    .z = in1.z + in2.z
    .w = in1.w + in2.w
  End With

End Function



'
' add a value to 4d-vector
'
Public Function add4fv(ByRef in1 As float4, ByVal in2 As Single) As float4

  With add4fv
    .x = in1.x + in2
    .y = in1.y + in2
    .z = in1.z + in2
    .w = in1.w + in2
  End With

End Function



'
' subtract a value from 4d-vector
'
Public Function sub4fv(ByRef in1 As float4, ByVal in2 As Single) As float4

  With sub4fv
    .x = in1.x - in2
    .y = in1.y - in2
    .z = in1.z - in2
    .w = in1.w - in2
  End With

End Function



'
' scale 4d-vector
'
Public Function scale4f(ByRef inp As float4, ByVal factor As Single) As float4

  With scale4f
    .x = inp.x * factor
    .y = inp.y * factor
    .z = inp.z * factor
    .w = inp.w * factor
  End With

End Function
