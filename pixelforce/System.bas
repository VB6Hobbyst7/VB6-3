Attribute VB_Name = "System"

'just to keep code clear
Option Explicit
Option Base 0

'data structure (for file dialogs)
Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

'code line
Private SceneLine As String

'import functions from comdlg32.dll (i hate any OCX files connected)
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'file info
Private Dialog As OPENFILENAME

'scan position
Private Ind As Integer

'main procedure
Private Sub Main()
  If Len(Command) > 0 Then MsgBox Command, vbExclamation + vbOKOnly, "Note"
  Load Workspace
End Sub

'show open file dialog
Public Function OpenFile(hwnd As Long, Title As String, StartFolder As String, Filter As String, DefaultExtension As String, Flags As Long) As String
  'configure
  Filter = Replace(Filter, "|", Chr(0))
  If Not Right(Filter, 1) = Chr(0) Then Filter = Filter & Chr(0)
  With Dialog
    .hwndOwner = hwnd
    .hInstance = App.hInstance
    .lpstrFilter = Filter
    .lpstrFile = Space(254)
    .nMaxFile = 255
    .lpstrDefExt = DefaultExtension
    .lpstrFileTitle = Space(254)
    .nMaxFileTitle = 255
    .lpstrInitialDir = StartFolder
    .lpstrTitle = Title
    .Flags = Flags
  End With
  Dialog.lStructSize = Len(Dialog)
  'show
  OpenFile = IIf(GetOpenFileName(Dialog), Trim(Dialog.lpstrFile), vbNullString)
End Function

'show save file dialog
Public Function SaveFile(hwnd As Long, Title As String, StartFolder As String, Filter As String, DefaultExtension As String, Flags As Long) As String
  'configure
  Filter = Replace(Filter, "|", Chr(0))
  If Not Right(Filter, 1) = Chr(0) Then Filter = Filter & Chr(0)
  With Dialog
    .hwndOwner = hwnd
    .hInstance = App.hInstance
    .lpstrFilter = Filter
    .lpstrFile = Space(254)
    .nMaxFile = 255
    .lpstrDefExt = DefaultExtension
    .lpstrFileTitle = Space(254)
    .nMaxFileTitle = 255
    .lpstrInitialDir = StartFolder
    .lpstrTitle = Title
    .Flags = Flags
  End With
  Dialog.lStructSize = Len(Dialog)
  'show
  SaveFile = IIf(GetSaveFileName(Dialog), Trim(Dialog.lpstrFile), vbNullString)
End Function

'this function will return application path with "/" correction
Public Function fixPath(Optional datPath As String = vbNullString) As String
  If Len(datPath) = 0 Then datPath = App.Path
  If Right(datPath, 1) <> "\" Then datPath = datPath & "\"
  fixPath = datPath
End Function

'insert mesh code
Public Sub InsertMesh(file As String)
  With Workspace.SceneCode
    'add code
    .Text = .Text & vbCrLf
    .Text = .Text & "<Mesh>" & vbCrLf
    .Text = .Text & "  File " & file & vbCrLf
    .Text = .Text & "  Position 0.00 0.00 0.00" & vbCrLf
    .Text = .Text & "  Rotation 0.00 0.00 0.00" & vbCrLf
    .Text = .Text & "  Scale 1.00 1.00 1.00" & vbCrLf
    .Text = .Text & "  Texture 0" & vbCrLf
    .Text = .Text & "  Lighting on" & vbCrLf
    .Text = .Text & "  Alpha 255" & vbCrLf
    .Text = .Text & "</Mesh>" & vbCrLf
    'move cursor to the end
    .SelStart = Len(.Text)
    .SelLength = 0
    .SetFocus
  End With
End Sub

'load scene file
Public Function LoadScene(file As String)
  'error handler
  On Error Resume Next
  With Workspace.SceneCode
    .Text = vbNullString
    'open file
    Open file For Input As #1
    If Not Err.Number = 0 Then
      'show error, close file
      LoadScene = False
      Err.Clear
      Close #1
      MsgBox "Could Not Open Scene File.", vbCritical + vbOKOnly, "Error"
    Else
      'add lines
      Do While Not EOF(1)
        Line Input #1, SceneLine
        .Text = .Text & SceneLine & vbCrLf
      Loop
      'move cursor to the end
      .SelStart = Len(.Text)
      .SelLength = 0
      Close #1
      LoadScene = True
    End If
  End With
End Function

'save scene file
Public Function SaveScene(file As String)
  'error handler
  On Error Resume Next
  'create/open file
  Open file For Output As #1
  If Not Err.Number = 0 Then
    'show error, close file
    SaveScene = False
    Err.Clear
    Close #1
    MsgBox "Could Not Save Scene File.", vbCritical + vbOKOnly, "Error"
  Else
    'write code
    Print #1, Workspace.SceneCode.Text;
    'all done
    Close #1
    SaveScene = True
  End If
End Function

'this is something like instr function, but with another direction - scans string from ending to begining
Public Function FindStrLeft(Position As Integer, data As String, Find As String) As Integer
  For Ind = Position - Len(Find) + 1 To 1 Step -1
    If Mid(data, Ind, 2) = Find Then
      FindStrLeft = Ind + 2
      Exit Function
    End If
  Next Ind
  FindStrLeft = 0
End Function

