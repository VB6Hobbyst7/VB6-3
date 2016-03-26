VERSION 5.00
Begin VB.Form windowWorkspace 
   Caption         =   "Composition Nodes"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15360
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imageDisplay 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   11520
      Left            =   0
      ScaleHeight     =   768
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   0
      Top             =   0
      Width           =   15360
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu itemNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu submenuOpen 
         Caption         =   "&Open"
         Begin VB.Menu itemOpen 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu itemSaveAs 
         Caption         =   "&Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu itemSeparator0 
         Caption         =   "-"
      End
      Begin VB.Menu itemQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu submenuAdd 
         Caption         =   "&Add"
         Begin VB.Menu itemAdd 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu itemRemove 
         Caption         =   "&Remove"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu itemSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu itemProperties 
         Caption         =   "&Properties..."
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu menuView 
      Caption         =   "&View"
      Begin VB.Menu itemReset 
         Caption         =   "Rese&t"
      End
   End
   Begin VB.Menu menuRun 
      Caption         =   "R&un"
      Begin VB.Menu itemProcess 
         Caption         =   "Render &All..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu itemProcessOne 
         Caption         =   "Render S&elected..."
         Shortcut        =   ^{F5}
      End
   End
End
Attribute VB_Name = "windowWorkspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' require variable declaration
Option Explicit
Option Base 0


' current file name
Private file As String
Private counter As Long



'
' application entry point
'
Private Sub Form_Load()

  ' fill up "edit->add->..." menu
  Dim i As Long: i = 0
  Do While (i < instance_count)

    ' create menu item and aquire it's caption
    If (i <> 0) Then Load itemAdd(i)
    Dim title As String: title = instance_title(i)
    If (Len(title) <> 0) Then
      itemAdd(i).Caption = instance_title(i) ' node
    Else
      itemAdd(i).Caption = "-" ' separator
    End If

    i = i + 1
  Loop

  ' reset working file name
  counter = 1
  file = "untitled" & counter

  selection = -1

  reset_actions ' reset mouse actions
  view_identity ' reset viewport
  graph_new     ' initialize new composition graph

  Icon = Nothing

End Sub



'
' shutdown
'
Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)

  ' exit confirmation
  If (MsgBox("All unsaved work will be lost. Exit program?", vbQuestion Or vbYesNo, "Warning") = vbNo) Then
    cancel = 1
    Exit Sub
  End If

  ' cleanup and exit
  graph_new
  End

End Sub



'
' window resize
'
Private Sub Form_Resize()

  ' do nothing when window is minimized
  If (WindowState = vbMinimized) Then Exit Sub

  ' resize workspace image
  imageDisplay.Move 0, 0, ScaleWidth, ScaleHeight

  ' update viewport size
  view_width = ScaleWidth
  view_height = ScaleHeight

  ' and refresh image
  layout_draw imageDisplay

End Sub



'
' mouse button press
'
Private Sub imageDisplay_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)

  mouse_down button, Int(x), Int(y)

  ' refresh image
  layout_draw imageDisplay

End Sub



'
' mouse movement
'
Private Sub imageDisplay_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)

  mouse_move button, Int(x), Int(y)

  ' refresh image
  If (button <> 0) Then layout_draw imageDisplay

End Sub



'
' mouse button release
'
Private Sub imageDisplay_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)

  mouse_up button, Int(x), Int(y)

  ' refresh image
  layout_draw imageDisplay

End Sub



'
' create new node
'
Private Sub itemAdd_Click(index As Integer)

  ' add node
  If (node_create(index) <> 0) Then
    If (index > 10) Then itemProcessOne_Click
    layout_draw imageDisplay ' refresh image
  End If

End Sub



'
' reset current project
'
Private Sub itemNew_Click()

  ' confirmation
  If (MsgBox("All unsaved work will be lost. Initialize new project?", vbYesNo Or vbQuestion, "Warning") = vbYes) Then

    selection = -1

    reset_actions ' reset all mouse actions
    view_identity ' reset viewport
    graph_new     ' initialize new composition graph

    ' generate new file name
    counter = counter + 1
    file = "untitled" & counter

    ' refresh image
    layout_draw imageDisplay

  End If

End Sub



'
' load project from file
'
Private Sub itemOpen_Click(index As Integer)

  ' check file name
  Dim length As Long: length = Len(itemOpen(index).Caption)
  If (length <> 0) Then

    selection = -1

    reset_actions ' reset all mouse actions
    view_identity ' reset viewport
    graph_new     ' initialize new composition graph

    If (project_read(App.Path & "\" & itemOpen(index).Caption) <> 0) Then

      ' update file name
      file = Left(itemOpen(index).Caption, length - 5)

      ' generate thumbnails for all nodes
      Dim i As Long: i = 0
      Do While (i < nodes)
        node(i).kernel.buffer.render_thumbnail node(i).thumbnail(), thumb_width, thumb_height
        i = i + 1
      Loop

    Else ' load failed

      ' generate new file name
      counter = counter + 1
      file = "untitled" & counter

      view_identity ' reset viewport
      graph_new     ' initialize new composition graph

      MsgBox "Failed to load project.", vbCritical, "Error"

    End If

    layout_draw imageDisplay ' refresh image

  End If

End Sub



'
' render entrie composition graph
'
Private Sub itemProcess_Click()

  reset_actions                 ' reset all mouse actions

  Load windowRender
  windowRender.Show vbModal, Me ' begin rendering
  Unload windowRender

  layout_draw imageDisplay      ' refresh image

End Sub



'
' render selected node only
'
Private Sub itemProcessOne_Click()

  If (selection <> -1) Then
    reset_actions                   ' reset all mouse actions

    Load windowRender
    windowRender.number = selection ' process only selected node
    windowRender.Show vbModal, Me   ' begin rendering
    Unload windowRender

    layout_draw imageDisplay        ' refresh image
  Else
    MsgBox "Node is not selected.", vbExclamation, "Warning"
  End If

End Sub



'
' open up node configuration dialog
'
Private Sub itemProperties_Click()

  If (selection <> -1) Then ' check selection

    reset_actions ' reset all mouse actions

    ' open settings dialog
    node(selection).kernel.show_options
    node(selection).thumb_valid = 0 ' invalidate thumbnail

    layout_draw imageDisplay ' refresh image (this will also update all invalid thumbs)

  End If

End Sub



'
' exit program
'
Private Sub itemQuit_Click()

  Unload Me

End Sub



'
' delete selected node
'
Private Sub itemRemove_Click()

  If (selection <> -1) Then ' check selection

    ' confirmation
    If (MsgBox("Delete node #" & selection & " [" & node(selection).kernel.get_title & "]?", vbYesNo Or vbQuestion, "Warning") = vbYes) Then

      ' destroy node
      node_destroy selection

      selection = -1 ' deselect
      reset_actions  ' reset all mouse actions

      ' refresh image
      layout_draw imageDisplay

    End If

  End If

End Sub



'
' reset viewport
'
Private Sub itemReset_Click()

  reset_actions ' reset all mouse actions
  view_identity ' reset viewport

  ' refresh image
  layout_draw imageDisplay

End Sub



'
' save project to file
'
Private Sub itemSaveAs_Click()

  ' aquire file name
  Dim name As String: name = InputBox("File name:", "Save Project", file)
  Dim length As Long: length = Len(name)
  If (length <> 0) Then

    ' check file name
    Dim i As Long: i = 1
    Do While (i <= length)

      ' test every character (A-Z, a-z, 0-9)
      Dim char As Byte: char = Asc(Mid(name, i, 1))
      If ((char < 65 Or char > 90) And (char < 97 Or char > 122) And (char < 48 Or char > 57)) Then length = 0: Exit Do

      i = i + 1 ' next character
    Loop

    ' save binary
    If (length <> 0) Then

      ' get file name with full path
      Dim fullname As String: fullname = App.Path & "\" & name & ".node"

      ' confirm overwrite
      If (Len(Dir(fullname)) <> 0) Then
        If (Err.number <> 0) Then Err.Clear: MsgBox "Failed to save project.", vbCritical, "Error": Exit Sub
        If (MsgBox("File already eixsts, overwrite?", vbQuestion Or vbYesNo, "Warning") = vbNo) Then Exit Sub
      End If

      If (project_write(fullname) <> 0) Then
        file = name ' update file name
      Else
        MsgBox "Failed to save project.", vbCritical, "Error"
      End If

    Else
      MsgBox "Invailid filename (allowed characters are: A..Z, a..z, 0..9).", vbExclamation, "Warning"
    End If

  End If

End Sub



'
' update file list
'
Private Sub menuFile_Click()

  ' handle all errors
  On Error Resume Next

  ' unload all items
  Dim i As Long: i = itemOpen.ubound
  Do While (i >= 0)
    itemOpen(i).Caption = vbNullString
    If (i <> 0) Then Unload itemOpen(i)
    i = i - 1
  Loop

  ' enumerate all files in current folder one by one
  Dim index As Long: index = 0
  Dim name As String: name = Dir(App.Path & "/")
  Do While (Len(name) <> 0)

    ' error checking
    If (Err.number <> 0) Then Exit Do

    ' add all text files into menu
    If Right(LCase(name), 5) = ".node" Then

      If (index <> 0) Then Load itemOpen(index) ' create menu item (first already exists)
      itemOpen(index).Caption = name            ' set item title
      index = index + 1                         ' increment counter

    End If

    name = Dir ' fetch next file name
  Loop

  ' error cleanup
  If (Err.number <> 0) Then Err.Clear

End Sub
