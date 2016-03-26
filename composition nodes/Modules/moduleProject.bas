Attribute VB_Name = "moduleProject"

' require variable declaration
Option Explicit
Option Base 0


' project file signature and version
Private Const signature As Long = &H65646F6E ' "node"
Private Const version   As Long = &H5        ' 5



'
' save project to file
'
Public Function project_write(ByVal filename As String) As Long

  ' error handling
  On Error Resume Next: project_write = 0

  ' delete old file (if exist)
  If (Len(Dir(filename)) <> 0) Then Kill filename
  If (Err.number <> 0) Then Err.Clear: Exit Function

  ' get file handle
  Dim out As Long: out = FreeFile

  ' create and open output file
  Open filename For Binary Access Write As #out
  If (Err.number <> 0) Then Err.Clear: Exit Function

  ' store header (signature + version)
  Put #out, , signature
  Put #out, , version

  ' store node counter
  Put #out, , nodes

  ' write nodes one by one
  Dim i As Long: i = 0
  Do While (i < nodes)
    With node(i)

      ' store type
      Put #out, , .type

      ' store connections
      Dim j As Long: j = 0: Dim n As Long: n = .kernel.get_inputs
      Do While (j < n)
        Put #out, , .socket(j)
        j = j + 1 ' next pin
      Loop

      ' store node-specific data
      If (.kernel.file_write(out) = 0) Then Error 1 ' trigger an error on failure

      ' store position
      Put #out, , .x
      Put #out, , .y

    End With
    If (Err.number <> 0) Then Exit Do
    i = i + 1 ' next node
  Loop

  ' remember current viewport
  Put #out, , view_pan_x
  Put #out, , view_pan_y
  Put #out, , view_zoom

  ' flush data and close file
  Close #out

  ' check for errors
  If (Err.number <> 0) Then ' failed
    Err.Clear
  Else
    project_write = 1       ' successful
  End If

End Function



'
' load project from file
'
Public Function project_read(ByVal filename As String) As Long

  ' error handling
  On Error Resume Next: project_read = 0

  ' check file existance
  If (Len(Dir(filename)) <> 0) Then
    If (Err.number <> 0) Then Err.Clear: Exit Function

    ' get file handle
    Dim inp As Long: inp = FreeFile
  
    ' open source file
    Open filename For Binary Access Read As #inp
    If (Err.number <> 0) Then Err.Clear: Exit Function

    ' aquire header (signature + version)
    Dim signature_test As Long: Get #inp, , signature_test
    Dim version_test As Long: Get #inp, , version_test

    ' validate header
    If (signature = signature_test And version = version_test) Then

      ' read node counter and allocate composition graph
      Get #inp, , nodes: nodes = clamp1i(nodes, 0, 1024)
      If (nodes <> 0) Then
        ReDim node(0 To nodes - 1) As node_format

        ' read nodes one by one
        Dim i As Long: i = 0
        Do While (i < nodes)
          With node(i)

            ' get type
            Get #inp, , .type

            ' initialize kernel
            Set .kernel = instance_create(.type)
            If (.kernel Is Nothing) Then
              Error 1 ' unknown node type
            Else

              ' aquire connections
              Dim n As Long: n = .kernel.get_inputs
              If (n <> 0) Then
                ReDim .socket(0 To n - 1) As Long
                Dim j As Long: j = 0
                Do While (j < n)
                  Get #inp, , .socket(j)
                  j = j + 1 ' next pin
                Loop
              End If

              ' load node-specific data
              If (.kernel.file_read(inp) = 0) Then Error 1 ' trigger an error on failure

              ' get position
              Get #inp, , .x
              Get #inp, , .y

              ' invalidate thumbnail
              .thumb_valid = 0
            
            End If

          End With
          If (Err.number <> 0) Then Exit Do
          i = i + 1 ' next node
        Loop

      End If

      ' restore viewport
      Get #inp, , view_pan_x
      Get #inp, , view_pan_y
      Dim zoom As Single: Get #inp, , zoom: view_zoom = 0: view_scale zoom

    Else
      Error 1 ' wrong header or unsupported version
    End If

    ' close file
    Close #inp

  Else
    Error 1 ' file does not exist
  End If

  ' check for errors
  If (Err.number <> 0) Then ' failed
    Err.Clear
  Else


    ' check all data connections and break invalid
    i = 0
    Do While (i < nodes)
      With node(i)

        ' check all connections for current node
        j = 0: n = .kernel.get_inputs
        Do While (j < n)
  
          If (.socket(j) < -1) Then .socket(j) = -1     ' node index can not be less than -1
          If (.socket(j) >= nodes) Then .socket(j) = -1 ' node index is out of range
          If (.socket(j) <> -1) Then ' if socket is still connected, perform a final check
            If (node(.socket(j)).kernel.get_output = 0) Then .socket(j) = -1 ' this node has no output pin
          End If
  
          j = j + 1 ' next socket
        Loop

      End With
      i = i + 1 ' next node
    Loop


    project_read = 1       ' successful
  End If

End Function
