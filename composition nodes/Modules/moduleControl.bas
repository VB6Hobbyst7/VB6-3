Attribute VB_Name = "moduleControl"

' require variable declaration
Option Explicit
Option Base 0


' last mouse coordinates
Public mouse_x As Long
Public mouse_y As Long

' selected and locked node indexes
Public selection As Long
Private locking As Long

' floating connection wire
Public src_node As Long ' source node
Public src_pin As Long  ' = -2: none, = -1: output, >= 0: input
Public src_x As Long    ' spline start point
Public src_y As Long

' thumbnail state
Public thumb_state As Long ' 0 - thumbs are valid, 1 - thumbs are not yet resized



'
' reset all mouse actions
'
Public Sub reset_actions()

  src_node = -1 ' drop floating connection
  locking = -1  ' release locking

  ' reset zoom state
  If (thumb_state = 1) Then thumb_state = 0

End Sub



Public Sub mouse_move(ByVal button As Long, ByVal x As Long, ByVal y As Long)

  Select Case (button)

    ' move node
    Case Is = vbLeftButton
      If (locking >= 0) Then
        With node(selection)
          .x = .x - (unmap_u(mouse_x) - unmap_u(x))
          .y = .y - (unmap_v(mouse_y) - unmap_v(y))
        End With
      End If

    ' pan
    Case Is = vbRightButton
      view_scroll mouse_x - x, mouse_y - y

    ' zoom
    Case Is = vbMiddleButton
      view_scale CSng(mouse_y - y) * 0.01

  End Select

  ' remeber mouse position
  mouse_x = x
  mouse_y = y

End Sub



Public Sub mouse_down(ByVal button As Long, ByVal x As Long, ByVal y As Long)

  ' zoom start (do not render thumbs now)
  If (button = vbMiddleButton) Then thumb_state = 1

  If (button = vbLeftButton) Then


    ' find node and pin indexes over mouse cursor
    node_indexes_at_coords x, y, src_node, src_pin

    If src_node <> -1 Then
      If src_pin = -2 Then

        selection = src_node ' activate node
        locking = src_node   ' and lock it (we are going to move it, right?)
        src_node = -1

      Else

        selection = src_node
        locking = -2       ' node is not locked, but cursor is fixed (we are connecting nodes)
        src_x = unmap_u(x) ' begin floating connection
        src_y = unmap_v(y)

      End If
    Else

      selection = -1 ' deactivate any selected (mouse is not over node or pin)
      locking = -1

    End If

  End If

  ' remeber mouse position
  mouse_x = x
  mouse_y = y

End Sub



Public Sub mouse_up(ByVal button As Long, ByVal x As Long, ByVal y As Long)

  ' zoom end (allow to render thumbnails)
  If (button = vbMiddleButton) Then thumb_state = 0

  If button = vbLeftButton Then

    ' release node lock
    locking = -1

    ' find out where wire dropped
    If (src_node <> -1) Then

      ' find node and pin indexes over mouse cursor
      Dim dst_node As Long ' target node
      Dim dst_pin As Long  ' = -2: none, = -1: output, >= 0: input
      node_indexes_at_coords x, y, dst_node, dst_pin

      ' connect or disconnect pins
      If (dst_node <> -1 And dst_pin <> -2) Then

        If (dst_node = src_node) Then
          disconnect_pin src_node, src_pin
        Else
          connect_pins src_node, src_pin, dst_node, dst_pin
        End If

      End If

      ' release source
      src_node = -1

    End If

  End If

  ' remeber mouse position
  mouse_x = x
  mouse_y = y

End Sub
