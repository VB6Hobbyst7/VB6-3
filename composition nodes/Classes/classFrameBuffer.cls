VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classFrameBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' require variable declaration
Option Explicit
Option Base 0


' hardcoded limitations

Private Const max_width As Long = 1024     ' max image width
Private Const max_height As Long = 1024    ' max image height



' image data
Private pixels_hdr() As float4 ' hdr
Private pixels_ldr() As Long   ' ldr


' image dimensions
Private width As Long
Private height As Long


' options

Private format As Long    ' pixel format (hdr/ldr)

Private filter As Long    ' interpolation mode

Private address_u As Long ' x addressing
Private address_v As Long ' y addressing

Private border As float4  ' border color



'
' generate thumbnail image
'
Public Sub render_thumbnail(ByRef p() As Long, ByVal w As Long, ByVal h As Long)

  ' check image size
  If (w < 0 Or w > max_width Or h < 0 Or h > max_height) Then Exit Sub ' wrong resolution

  ' re/allocate memory
  Erase p()
  ReDim p(0 To w - 1, 0 To h - 1) As Long

  ' compute aspect ratio
  Dim delta_u As Single: delta_u = CSng(width - 1) / CSng(w)
  Dim delta_v As Single: delta_v = CSng(height - 1) / CSng(h)

  Dim src_u As Single: Dim src_x As Long, dst_x As Long
  Dim src_v As Single: Dim src_y As Long, dst_y As Long

  dst_y = 0: src_y = 0: src_v = 0
  Do While (dst_y < h)

    dst_x = 0: src_x = 0: src_u = 0
    Do While (dst_x < w)

      If (format <> 0) Then
        p(dst_x, dst_y) = fp2rgba(pixels_hdr(src_x, src_y))
      Else
        p(dst_x, dst_y) = pixels_ldr(src_x, src_y)
      End If

      src_u = src_u + delta_u
      src_x = Int(src_u)
      dst_x = dst_x + 1
    Loop

    src_v = src_v + delta_v
    src_y = Int(src_v)
    dst_y = dst_y + 1
  Loop

End Sub



'
' return image width
'
Public Function get_width() As Long

  get_width = width

End Function



'
' return image height
'
Public Function get_height() As Long

  get_height = height

End Function



'
' write pixel
'
Public Sub store(ByVal x As Long, ByVal y As Long, ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)

  If (x < 0 Or y < 0 Or x >= width Or y >= height) Then Exit Sub ' wrong pixel coords

  ' store pixel data
  If (format <> 0) Then
    pixels_hdr(x, y) = vector4f(r, g, b, a)
  Else
    pixels_ldr(x, y) = fp2rgba(vector4f(r, g, b, a))
  End If

End Sub



'
' read pixel
'
Public Sub fetch(ByVal x As Long, ByVal y As Long, ByRef r As Single, ByRef g As Single, ByRef b As Single, ByRef a As Single)

  If (x < 0 Or y < 0 Or x >= width Or y >= height) Then r = 0: g = 0: b = 0: a = 0: Exit Sub ' wrong pixel coords

  ' aquire pixel data
  Dim color As float4
  If (format <> 0) Then
    color = pixels_hdr(x, y)
  Else
    color = rgba2fp(pixels_ldr(x, y))
  End If

  ' extract channels
  With color
    r = .x
    g = .y
    b = .z
    a = .w
  End With

End Sub



'
' sample pixel
'
Public Sub fetch_sample(ByVal u As Single, ByVal v As Single, ByRef r As Single, ByRef g As Single, ByRef b As Single, ByRef a As Single)

  ' output color
  Dim out As float4

  Select Case filter
    Case Is = 0: out = sample_nearest(u, v)  ' no filtering
    Case Is = 1: out = sample_bilinear(u, v) ' bilinear filter
  End Select

  ' extract channels
  With out
    r = .x
    g = .y
    b = .z
    a = .w
  End With

End Sub



'
' show configuration dialog
'
Public Function show_options() As Long

  ' load properties dialog
  Dim dialog As New windowFrameBuffer: Load dialog
  With dialog


  ' submit data into dialog

  ' image dimensions
  .p_width = width
  .p_height = height

  ' pixel format
  .p_pixel = format

  ' u, v addressing mode
  .p_u = address_u
  .p_v = address_v

  ' sampling filter
  .p_filter = filter

  ' border color
  .p_r = border.x
  .p_g = border.y
  .p_b = border.z
  .p_a = border.w


  ' show up dialog and wait until it's hidden
  .update
  .Show vbModal


  ' read back modified data from dialog
  If .result <> 0 Then

    ' image dimensions
    width = clamp1i(.p_width, 1, max_width)
    height = clamp1i(.p_height, 1, max_height)

    ' pixel format
    format = clamp1i(.p_pixel, 0, 1)

    ' u, v addressing mode
    address_u = clamp1i(.p_u, 0, 3)
    address_v = clamp1i(.p_v, 0, 3)

    ' sampling filter
    filter = clamp1i(.p_filter, 0, 1)

    ' border color
    border = clamp4f(vector4f(.p_r, .p_g, .p_b, .p_a), 0, 1)


    ' reallocate image
    create

  End If


  ' unload window
  End With
  Unload dialog: Set dialog = Nothing

End Function



'
' store properties to file
'
Public Function file_write(ByVal f As Long) As Long

  ' error handler
  On Error Resume Next


  ' write properties
  Put #f, , width
  Put #f, , height
  Put #f, , format
  Put #f, , address_u
  Put #f, , address_v
  Put #f, , filter
  Put #f, , border


  ' i/o error check
  If (Err.number <> 0) Then Err.Clear: file_write = 0: Exit Function ' fail
  file_write = 1 ' success

End Function



'
' restore properties from file (repair them if possible)
'
Public Function file_read(ByVal f As Long) As Long

  ' error handler
  On Error Resume Next


  ' read properties
  Get #f, , width: width = clamp1i(width, 1, max_width)
  Get #f, , height: height = clamp1i(height, 1, max_height)
  Get #f, , format: format = clamp1i(format, 0, 1)
  Get #f, , address_u: address_u = clamp1i(address_u, 0, 3)
  Get #f, , address_v: address_v = clamp1i(address_v, 0, 3)
  Get #f, , filter: filter = clamp1i(filter, 0, 1)
  Get #f, , border: border = clamp4f(border, 0, 1)


  ' allocate image
  create

  ' i/o error check
  If (Err.number <> 0) Then Err.Clear: file_read = 0: Exit Function ' fail
  file_read = 1 ' success

End Function



'
' allocate frame buffer
'
Private Sub create()

  ' cleanup
  Erase pixels_ldr()
  Erase pixels_hdr()

  ' allocate memory
  If (format <> 0) Then
    ReDim pixels_hdr(0 To width - 1, 0 To height - 1) As float4
  Else
    ReDim pixels_ldr(0 To width - 1, 0 To height - 1) As Long
  End If

End Sub



'
' sample pixel: nearest
'
Private Function sample_nearest(ByVal u As Single, ByVal v As Single) As float4

  ' recomputed pixel coordinates
  Dim x As Long: x = Int(u)
  Dim y As Long: y = Int(v)


  ' u
  Select Case (address_u)

    Case Is = 0 ' drop edges to border color
      If (x < 0 Or x >= width) Then sample_nearest = border: Exit Function

    Case Is = 1 ' clamp to edge
      x = clamp1i(x, 0, width - 1)

    Case Is = 2 ' wrap over the edge
      x = wrap1i(x, width)
      
    Case Is = 3 ' mirror from the edge
      x = mirror1i(x, width)

  End Select


  ' v
  Select Case (address_v)

    Case Is = 0 ' drop edges to border color
      If (y < 0 Or y >= height) Then sample_nearest = border: Exit Function

    Case Is = 1 ' clamp to edge
      y = clamp1i(y, 0, height - 1)

    Case Is = 2 ' wrap over the edge
      y = wrap1i(y, height)

    Case Is = 3 ' mirror from the edge
      y = mirror1i(y, height)

  End Select


  ' get pixel color
  If (format <> 0) Then
    sample_nearest = pixels_hdr(x, y)
  Else
    sample_nearest = rgba2fp(pixels_ldr(x, y))
  End If

End Function



'
' sample pixel: bilinear
'
Private Function sample_bilinear(ByVal u As Single, ByVal v As Single) As float4

  ' floor
  Dim x As Long: x = Int(u)
  Dim y As Long: y = Int(v)

  ' delta +
  Dim u_ratio As Single: u_ratio = u - CSng(x)
  Dim v_ratio As Single: v_ratio = v - CSng(y)

  ' delta -
  Dim u_opposite As Single: u_opposite = 1 - u_ratio
  Dim v_opposite As Single: v_opposite = 1 - v_ratio

  ' fetch four pixels
  Dim p00 As float4: p00 = sample_nearest(x, y)
  Dim p10 As float4: p10 = sample_nearest(x + 1, y)
  Dim p01 As float4: p01 = sample_nearest(x, y + 1)
  Dim p11 As float4: p11 = sample_nearest(x + 1, y + 1)

  ' blend them all according to their weights
  sample_bilinear = add4f(scale4f(add4f(scale4f(p00, u_opposite), scale4f(p10, u_ratio)), v_opposite), _
                          scale4f(add4f(scale4f(p01, u_opposite), scale4f(p11, u_ratio)), v_ratio))

End Function



'
' startup
'
Private Sub Class_Initialize()

  ' use default settings
  width = 128   ' 128px
  height = 128  ' by 128px
  format = 1    ' hdr
  filter = 0    ' nearest
  address_u = 2 ' wrap
  address_v = 2 ' wrap
  border = vector4f(0, 0, 0, 0) ' black & transparent

  ' allocate image
  create

End Sub



'
' shutdown
'
Private Sub Class_Terminate()

  ' memory cleanup
  Erase pixels_ldr()
  Erase pixels_hdr()

End Sub
