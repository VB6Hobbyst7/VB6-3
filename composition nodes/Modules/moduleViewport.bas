Attribute VB_Name = "moduleViewport"

' require variable declaration
Option Explicit
Option Base 0


' viewport dimensions
Public view_width As Long
Public view_height As Long

' scale factor
Public view_zoom As Single

' pan
Public view_pan_x As Single
Public view_pan_y As Single


' thumbnail resolution
Public thumb_width As Long
Public thumb_height As Long



'
' reset viewport
'
Public Sub view_identity()

  ' look at system origin
  view_pan_x = 0
  view_pan_y = 0

  ' 1:1 scale
  view_zoom = 1

  ' default thumbnail size
  thumb_width = 128
  thumb_height = 128

  ' invalidate all thumbnails (scale changed, we need them all to be resized too)
  Dim i As Long: i = 0: Do While (i < nodes): node(i).thumb_valid = 0: i = i + 1: Loop

End Sub



'
' scale viewport (zoom)
'
Public Sub view_scale(ByVal f As Single)

  ' update scale factor
  view_zoom = clamp1f(view_zoom + f, 0.5, 2) ' clamp in range from 50% (64x64) up to 200% (256x256)

  ' update thumbnail size
  thumb_width = view_zoom * 128
  thumb_height = view_zoom * 128

  ' invalidate all thumbnails (scale changed, we need them all to be resized too)
  Dim i As Long: i = 0: Do While (i < nodes): node(i).thumb_valid = 0: i = i + 1: Loop

End Sub



'
' scroll viewport (pan)
'
Public Sub view_scroll(ByVal x As Long, ByVal y As Long)

  ' update scrolling
  view_pan_x = view_pan_x + CSng(x) / view_zoom
  view_pan_y = view_pan_y + CSng(y) / view_zoom

End Sub



'
' project x coord from world into screen space
'
Public Function map_u(ByVal u As Single) As Long

  map_u = Int(CSng(view_width) * 0.5 + (u - view_pan_x) * view_zoom)

End Function



'
' project y coord from world into screen space
'
Public Function map_v(ByVal v As Single) As Long

  map_v = Int(CSng(view_height) * 0.5 + (v - view_pan_y) * view_zoom)

End Function



'
' project x coord from screen space into world
'
Public Function unmap_u(ByVal u As Long) As Single

  unmap_u = (CSng(u) - CSng(view_width) * 0.5) / view_zoom + view_pan_x

End Function



'
' project y coord from screen space into world
'
Public Function unmap_v(ByVal v As Long) As Single

  unmap_v = (CSng(v) - CSng(view_height) * 0.5) / view_zoom + view_pan_y

End Function
