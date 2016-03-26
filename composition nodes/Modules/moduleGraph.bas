Attribute VB_Name = "moduleGraph"

' require variable declaration
Option Explicit
Option Base 0


' node descriptor structure
Public Type node_format

  ' position in the "world"
  x As Single
  y As Single

  ' node type and processor instance
  type As Long     ' what kernel to use
  kernel As Object ' can be classNodeUniformColor, classNodeImageTransform, e.t.c...

  ' input connections (node indexes, on which 'kernel' depends on)
  socket() As Long

  ' output validation flag (required to render entire graph at once, used in windowRender only)
  valid As Long

  ' frame buffer thumbnail (display only)
  thumbnail() As Long
  thumb_valid As Long ' is this thumbnail valid?

End Type


' composition graph storage
Public node() As node_format ' node descriptor heap
Public nodes As Long         ' total number of nodes in the composition



'
' initialize new layout (delete all nodes)
'
Public Sub graph_new()

  If (nodes <> 0) Then ' scene clean up required?

    ' sub element cleanup
    Dim i As Long
    i = 0
    Do While (i < nodes)
      With node(i)

        Set .kernel = Nothing ' destroy processor
        Erase .socket()       ' free dependency indexes
        Erase .thumbnail()    ' free thumbnail image

      End With
      i = i + 1
    Loop

    ' cleanup
    nodes = 0
    Erase node()

  End If

End Sub



'
' create new node at current cursor position
'
Public Function node_create(ByVal id As Long) As Long

  ' 1024 nodes - max limit, anybody needs more? fix it your self
  If (nodes = 1024) Then node_create = 0: Exit Function

  ' create processor core
  Dim core As Object: Set core = instance_create(id)
  If (core Is Nothing) Then node_create = 0: Exit Function ' wrong type

  ' allocate memory for new desciptor
  If (nodes = 0) Then
    ReDim node(0 To 0) As node_format              ' first allocation
  Else
    ReDim Preserve node(0 To nodes) As node_format ' all future allocations
  End If

  With node(nodes)

    ' initialize core
    Set .kernel = core
    .valid = 0

    ' center screen
    .x = unmap_u(view_width \ 2 - 64 * view_zoom)
    .y = unmap_v(view_height \ 2 - 64 * view_zoom)

    ' remember type
    .type = id

    ' initialize input sockets
    Dim n As Long: n = core.get_inputs
    If (n <> 0) Then

      ReDim .socket(0 To n - 1) As Long

      Dim i As Long: i = 0
      Do While (i < n)
        .socket(i) = -1 ' socket is not connected to anything
        i = i + 1
      Loop

    End If

    ' invalidate thumbnail
    .thumb_valid = 0

  End With

  selection = nodes ' select this node
  nodes = nodes + 1 ' one more node in the graph
  node_create = 1   ' success

End Function



'
' destroy node and all it's connections
'
Public Sub node_destroy(ByVal index As Long)

  ' update node connections
  Dim i As Long
  i = 0
  Do While (i < nodes)
    With node(i)

      ' process all connections
      Dim k As Long: k = .kernel.get_inputs
      Dim j As Long: j = 0
      Do While (j < k)

        If (.socket(j) = index) Then
          .socket(j) = -1 ' disconnect link
        Else
          If (.socket(j) > index) Then .socket(j) = .socket(j) - 1 ' relocate link
        End If

        j = j + 1 ' next connection
      Loop

    End With
    i = i + 1 ' next node
  Loop

  ' dispose node kernel and socket array
  With node(index)
    Set .kernel = Nothing
    Erase .socket()
    Erase .thumbnail()
  End With

  ' relocate nodes in the storage
  i = index
  Do While (i < nodes - 1)
    node(i) = node(i + 1)
    i = i + 1
  Loop

  ' resize node storage
  nodes = nodes - 1
  If (nodes <> 0) Then
    ReDim Preserve node(0 To nodes - 1) As node_format
  Else
    Erase node()
  End If

End Sub



'
' break all connections with given pin
'
Public Sub disconnect_pin(ByVal node0 As Long, ByVal pin0 As Long)

  If pin0 = -1 Then ' output pin

    ' check every node
    Dim i As Long
    i = 0
    Do While (i < nodes)
      With node(i)

        ' check every pin
        Dim k As Long: k = .kernel.get_inputs
        Dim j As Long: j = 0
        Do While j < k

          ' is this pin connected with given pin?
          If (.socket(j) = node0) Then .socket(j) = -1 ' break connection

          j = j + 1
        Loop

      End With
      i = i + 1
    Loop

  Else ' input pin
    node(node0).socket(pin0) = -1 ' break connection
  End If

End Sub



'
' establish connection between pins (output->input / input<-output)
'
Public Sub connect_pins(ByVal node0 As Long, ByVal pin0 As Long, ByVal node1 As Long, ByVal pin1 As Long)

  If (pin0 = -1) Then ' node0 is "output"
    If (pin1 <> -1) Then node(node1).socket(pin1) = node0 ' node1 is "input"
  Else                ' node0 is "input"
    If (pin1 = -1) Then node(node0).socket(pin0) = node1  ' node1 is "output"
  End If

End Sub



'
' aquire node and pin indexes at given point
'
Public Sub node_indexes_at_coords(ByVal x As Single, ByVal y As Single, ByRef n As Long, ByRef p As Long)

  ' reset
  n = -1
  p = -2

  ' check every node
  Dim i As Long
  i = nodes - 1
  Do While (i > -1)
    With node(i)

      ' get transformed frame rectagle
      Dim x0 As Long: x0 = map_u(.x)
      Dim y0 As Long: y0 = map_v(.y)
      Dim x1 As Long: x1 = map_u(.x + 128)
      Dim y1 As Long: y1 = map_v(.y + 128)

      ' check output pin
      If (point_vs_box2d(x, y, x1 + 5 - 4, y0 + 30 - 4, x1 + 5 + 4, y0 + 30 + 4) <> 0 And .kernel.get_output <> 0) Then
        n = i: p = -1
      Else

        ' check input pins
        Dim j As Long, k As Long
        j = 0: k = .kernel.get_inputs
        Do While (j < k)
          If (point_vs_box2d(x, y, x0 - 4, y0 + 30 + 12 * j - 4, x0 + 4, y0 + 30 + 12 * j + 4) <> 0) Then n = i: p = j: Exit Do
          j = j + 1 ' next input pin
        Loop

      End If

      ' mouse pointer over current node?
      If (n = -1) Then
        If (point_vs_box2d(x, y, x0, y0, x1 + 5, y1 + 22) <> 0) Then n = i: Exit Do
      End If

    End With
    If (n <> -1) Then Exit Do
    i = i - 1 ' next node
  Loop

End Sub
