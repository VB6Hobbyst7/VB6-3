Attribute VB_Name = "moduleLayout"

' require variable declaration
Option Explicit
Option Base 0



Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

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

Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long



'
' render cubic spline (y-pass)
'
Private Sub draw_cubic_spline(ByRef target As PictureBox, ByVal x0 As Long, y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal color As Long, ByVal factor As Single, ByVal limit As Long)

  ' culling
  If box_vs_box2d(0, 0, view_width - 1, view_height - 1, x0, y0, x3, y3) = 0 Then Exit Sub

  ' compute delta
  Dim d As Single
  d = 1 / clamp1f(Sqr((x3 - x0) ^ 2 + (y3 - y0) ^ 2) * factor, 1, limit)

  ' first coords
  Dim u0 As Long
  Dim v0 As Long
  u0 = x0
  v0 = y0

  ' draw spline
  Dim f As Single
  f = d
  Do While f < 1

    ' compute next coords
    Dim u1 As Long
    Dim v1 As Long
    u1 = Int(cubic1f(x1, x0, x3, x2, f)) ' x is cubic
    v1 = Int(cubic1f(y1, y0, y3, y2, f)) ' y is cubic

    ' draw segment
    target.Line (u0, v0)-(u1, v1), color

    ' remember last coords
    u0 = u1
    v0 = v1

    ' next segment
    f = f + d

  Loop

  ' last segment
  target.Line (u0, v0)-(x3, y3), color

End Sub



'
' render cosine spline (x-pass)
'
Private Sub draw_spline(ByRef target As PictureBox, ByVal x0 As Long, y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal color As Long, ByVal factor As Single, ByVal limit As Long)

  ' culling
  If box_vs_box2d(0, 0, view_width - 1, view_height - 1, x0, y0, x1, y1) = 0 Then Exit Sub

  ' compute delta
  Dim d As Single
  d = 1 / clamp1f(Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2) * factor, 1, limit)

  ' first coords
  Dim u0 As Long
  Dim v0 As Long
  u0 = x0
  v0 = y0

  ' draw spline
  Dim f As Single
  f = d
  Do While f < 1

    ' compute next coords
    Dim u1 As Long
    Dim v1 As Long
    u1 = Int(lerp1f(x0, x1, f))   ' x is linear
    v1 = Int(cosine1f(y0, y1, f)) ' y is cosine

    ' draw segment
    target.Line (u0, v0)-(u1, v1), color

    ' remember last coords
    u0 = u1
    v0 = v1

    ' next segment
    f = f + d

  Loop

  ' last segment
  target.Line (u0, v0)-(x1, y1), color

End Sub



'
' render grid
'
Private Sub draw_grid(ByRef target As PictureBox, ByVal size As Long, ByVal color As Long)

  ' precompute interval
  Dim i As Single
  i = view_zoom * CSng(size)

  ' precompute delta values
  Dim su As Single
  Dim sv As Single
  su = view_pan_x + (view_zoom - 1) * view_pan_x
  sv = view_pan_y + (view_zoom - 1) * view_pan_y

  ' vertical lines
  Dim u As Single, x As Single
  u = i - wrap1f(su - CSng(view_width) * 0.5, i)
  Do While (u < view_width)

    x = Int(u)
    target.Line (x, 0)-(x, view_height - 1), color

    u = u + i
  Loop

  ' horizontal lines
  Dim v As Single, y As Single
  v = i - wrap1f(sv - CSng(view_height) * 0.5, i)
  Do While (v < view_height)

    y = Int(v)
    target.Line (0, y)-(view_width - 1, y), color

    v = v + i
  Loop

End Sub



'
' render complete layout
'
Public Sub layout_draw(ByRef target As PictureBox)

  ' fill background (.Cls is slow)
  target.Line (0, 0)-(view_width - 1, view_height - 1), RGB(47, 47, 47), BF

  ' draw grid lines
  draw_grid target, 50, RGB(55, 55, 55)

  ' axis
  Dim x As Long: x = map_u(0)
  Dim y As Long: y = map_v(0)
  If (x >= 0 And x < view_width) Then target.Line (x, 0)-(x, view_height - 1), RGB(79, 79, 79)
  If (y >= 0 And y < view_height) Then target.Line (0, y)-(view_width - 1, y), RGB(79, 79, 79)

  ' render connections
  Dim i As Long
  i = 0
  Do While i < nodes
    With node(i)

      ' check all input pins
      Dim j As Long, n As Long
      j = 0: n = .kernel.get_inputs
      Do While j < n

        ' get connection source
        Dim src As Long
        src = .socket(j)

        ' connected?
        If src <> -1 Then

          ' get spline coords (input pin)
          Dim u0 As Long
          Dim v0 As Long
          u0 = map_u(.x)
          v0 = map_v(.y) + 30 + 12 * j

          ' get spline coords (output pin)
          Dim u1 As Long
          Dim v1 As Long
          u1 = map_u(node(src).x + 128) + 5
          v1 = map_v(node(src).y) + 30
          
          ' cubic or cosine?
          If u0 - u1 < 0 Then
            Dim d As Long: d = (u0 - u1) * 5
            draw_cubic_spline target, u0, v0, u0 - d, v0, u1 + d, v1, u1, v1, RGB(255, 255, 255), 0.25, 50
          Else
            draw_spline target, u0, v0, u1, v1, RGB(255, 255, 255), 0.1, 50
          End If

        End If

        j = j + 1 ' next pin
      Loop

    End With
    i = i + 1
  Loop

  ' setup font
  target.FontName = "Small Fonts"
  target.ForeColor = RGB(255, 255, 255)

  ' render node frames
  i = 0
  Do While (i < nodes)
    With node(i)

      ' get transformed frame rectagle
      Dim x0 As Long: x0 = map_u(.x)
      Dim y0 As Long: y0 = map_v(.y)
      Dim x1 As Long: x1 = map_u(.x + 128)
      Dim y1 As Long: y1 = map_v(.y + 128)

      If (box_vs_box2d(0, 0, view_width - 1, view_height - 1, x0, y0, x1 + 5, y1 + 22) <> 0) Then


        ' node frame
        If (i = selection) Then
          target.Line (x0, y0)-(x1 + 5, y0 + 17), RGB(0, 95, 255), BF
        Else
          target.Line (x0, y0)-(x1 + 5, y0 + 17), RGB(95, 95, 95), BF
        End If
        target.Line (x0, y0 + 18)-(x1 + 5, y1 + 22), RGB(111, 111, 111), BF
        target.Line (x0, y0)-(x1 + 5, y1 + 22), RGB(191, 191, 191), B


        ' all thumbnails are valid?
        If (thumb_state = 0 And .kernel.get_output <> 0) Then

          ' thumbnail validation
          If (.thumb_valid = 0) Then
            .kernel.buffer.render_thumbnail .thumbnail(), thumb_width, thumb_height ' update it!
            .thumb_valid = 1 ' it is valid now
          End If

          ' setup thumbnail bitmap descriptor
          Dim b As BITMAPINFO
          With b.bmiHeader
            .biBitCount = 32
            .biPlanes = 1
            .biSize = Len(b)
            .biHeight = -thumb_height
            .biWidth = thumb_width
            .biSizeImage = thumb_width * thumb_height
          End With

          ' draw thumbnail
          SetDIBitsToDevice target.hdc, x0 + 3, y0 + 20, thumb_width, thumb_height, 0, 0, 0, thumb_height, .thumbnail(0, 0), b, 0

        End If


        ' caption text
        target.FontSize = 7
        target.CurrentX = x0 + 4
        target.CurrentY = y0 + 4
        target.Print .kernel.get_title


        ' render output pin
        If (.kernel.get_output <> 0) Then
          If (i = selection) Then
            target.DrawWidth = 4: target.Circle (x1 + 5, y0 + 30), 2, RGB(127, 0, 0)
            target.DrawWidth = 1: target.Circle (x1 + 5, y0 + 30), 4, RGB(255, 0, 0)
          Else
            target.DrawWidth = 4: target.Circle (x1 + 5, y0 + 30), 2, RGB(127, 127, 127)
            target.DrawWidth = 1: target.Circle (x1 + 5, y0 + 30), 4, RGB(255, 255, 255)
          End If
        End If


        ' render input pins
        n = .kernel.get_inputs
        If (n <> 0) Then
          j = 0
          Do While (j < n)

            If (i = selection) Then

              ' fill circle
              target.DrawWidth = 4
              If .kernel.get_type(j) <> 0 Then
                target.Circle (x0, y0 + 30 + 12 * j), 2, RGB(191, 191, 0)
              Else
                target.Circle (x0, y0 + 30 + 12 * j), 2, RGB(0, 127, 0)
              End If
              target.DrawWidth = 1

              ' draw label (if exist)
              Dim id As String: id = .kernel.get_name(j)
              If (Len(id) <> 0) Then
                target.FontSize = 6
                Dim l As Long: l = Int(target.TextWidth(id) + 10)

                ' background
                If (.kernel.get_type(j) <> 0) Then
                  target.Line (x0, y0 + 26 + 12 * j)-(x0 - l, y0 + 34 + 12 * j), RGB(255, 255, 0), B
                  target.Line (x0, y0 + 27 + 12 * j)-(x0 - l + 1, y0 + 33 + 12 * j), RGB(127, 127, 0), BF
                Else
                  target.Line (x0, y0 + 26 + 12 * j)-(x0 - l, y0 + 34 + 12 * j), RGB(0, 255, 0), B
                  target.Line (x0, y0 + 27 + 12 * j)-(x0 - l + 1, y0 + 33 + 12 * j), RGB(0, 127, 0), BF
                End If

                ' text
                target.CurrentX = x0 - l + 3
                target.CurrentY = y0 + 25 + 12 * j
                target.Print id

              End If

              ' circle outline
              If (.kernel.get_type(j) <> 0) Then
                target.Circle (x0, y0 + 30 + 12 * j), 4, RGB(255, 255, 0)
              Else
                target.Circle (x0, y0 + 30 + 12 * j), 4, RGB(0, 255, 0)
              End If

            Else
              target.DrawWidth = 4: target.Circle (x0, y0 + 30 + 12 * j), 2, RGB(127, 127, 127)
              target.DrawWidth = 1: target.Circle (x0, y0 + 30 + 12 * j), 4, RGB(255, 255, 255)
            End If

            j = j + 1 ' next pin
          Loop
        End If


      End If

    End With
    i = i + 1
  Loop

  ' floating wire
  If src_node <> -1 Then

    ' get spline points
    Dim su0 As Long: su0 = mouse_x
    Dim sv0 As Long: sv0 = mouse_y
    Dim su1 As Long: su1 = map_u(src_x)
    Dim sv1 As Long: sv1 = map_v(src_y)
    Dim du As Long: du = (su0 - su1) * 5 ' for cubic

    ' choose color and spline direction
    If src_pin = -1 Then

      ' from output
      If su0 < su1 Then
        draw_cubic_spline target, su0, sv0, su0 - du, sv0, su1 + du, sv1, su1, sv1, RGB(255, 0, 0), 0.25, 50
      Else
        draw_spline target, su0, sv0, su1, sv1, RGB(255, 0, 0), 0.1, 50
      End If

    Else

      ' from input
      If su0 > su1 Then
        draw_cubic_spline target, su0, sv0, su0 - du, sv0, su1 + du, sv1, su1, sv1, RGB(0, 255, 0), 0.25, 50
      Else
        draw_spline target, su0, sv0, su1, sv1, RGB(0, 255, 0), 0.1, 50
      End If

    End If

  End If

  ' update image
  target.Refresh

End Sub
