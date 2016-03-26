VERSION 5.00
Begin VB.Form windowRender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Render"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tickRefresh 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   1080
   End
   Begin VB.Timer tickLauncher 
      Interval        =   100
      Left            =   240
      Top             =   1080
   End
   Begin VB.CommandButton buttonStop 
      Cancel          =   -1  'True
      Caption         =   "&Stop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label infoPercent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5055
   End
   Begin VB.Shape shapeForeground 
      BorderColor     =   &H80000015&
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   240
      Top             =   600
      Width           =   2415
   End
   Begin VB.Shape shapeBackground 
      BorderColor     =   &H80000015&
      FillColor       =   &H80000010&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   240
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label infoRendering 
      Caption         =   "Starting up..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "windowRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' require variable declaration
Option Explicit
Option Base 0


' task information

Public number As Long     ' render selected node (0...+) or entrie graph (-1)?

Public cancel As Long     ' terminate rendering process (0 - no, 1 - yes)?
Public progress As Single ' current node progress (from 0 to 1 = 100%)

Private this As Long      ' current node number (overall progress)
Private total As Long     ' total number of nodes to render



'
' terminate rendering
'
Private Sub buttonStop_Click()

  buttonStop.Enabled = False
  cancel = 1

End Sub



'
' startup
'
Private Sub Form_Load()

  ' reset vars
  cancel = 0
  progress = 0

  number = -1 ' process all nodes by default
  total = nodes
  this = 1

  ' set progress bar width
  shapeForeground.width = 0

End Sub



'
' rendering process launcher
'
Private Sub tickLauncher_Timer()

  ' toggle timers
  tickLauncher.Enabled = False
  tickRefresh.Enabled = True

  ' referenced nodes
  Dim refs() As classFrameBuffer

  If number <> -1 Then

    total = 1 ' process single node only
    With node(number)

      ' show information
      infoRendering.Caption = "Processing node #" & number & " [" & .kernel.get_title & "]..."
      DoEvents

      ' build dependency list (no matter is invalid)
      build_dependencies number, refs(), 0

      ' invoke node rendering engine
      If .kernel.render(refs()) = 0 Then
        MsgBox "Failed to render node, verify all connections and settings.", vbExclamation, "Warning"
      Else
        .thumb_valid = 0 ' thumbnail is no longer valid
      End If

    End With

  Else ' process entrie composition graph (default)
    If total <> 0 Then

      ' invalidate all nodes
      Dim i As Long: i = 0
      Do While (i < nodes)
        node(i).valid = 0
        i = i + 1
      Loop

      ' render loop
      Do While (cancel = 0)

        Dim n As Long: n = 0 ' number of nodes processed in last iteration
        i = 0
        Do While (i < nodes)

          ' leave loop by user request
          If (cancel <> 0) Then Exit Do

          With node(i)
            If (.valid = 0) Then ' render only invalid nodes

              If (build_dependencies(i, refs(), 1) <> 0) Then ' query valid dependency list

                ' show information
                infoRendering.Caption = "Processing node #" & i & " [" & .kernel.get_title & "]..."
                DoEvents

                ' try to process node
                If .kernel.render(refs()) <> 0 Then
                  .thumb_valid = 0 ' thumbnail is no longer valid
                  .valid = 1       ' processed and valid now
                  n = n + 1        ' one more node in this iteration
                  this = this + 1  ' one more node in total
                  progress = 0     ' next possible node is not rendered at all
                End If
              
              End If

            End If
          End With

          DoEvents

          i = i + 1 'try next node
        Loop
        
        DoEvents
        If n = 0 Then Exit Do ' all possible nodes are processed

      Loop

    End If
  End If

  number = -2  ' render finished
  Erase refs() ' cleanup

  ' update ui
  progress = 1
  this = total
  shapeForeground.width = shapeBackground.width
  infoPercent.Caption = "100%"
  infoRendering.Caption = "Finished."
  buttonStop.Enabled = False

End Sub



'
' update progress bar
'
Private Sub tickRefresh_Timer()

  If total <> 0 Then

    ' compute progress bar length
    Dim f As Single
    f = (shapeBackground.width) / total
    f = f * (this - 1) + f * progress

    ' update progress bar
    shapeForeground.width = Int(f)
    infoPercent = Int(100 / shapeBackground.width * shapeForeground.width) & "%"

  End If

  layout_draw windowWorkspace.imageDisplay ' refresh image

  ' close window when complete
  If (number = -2) Then Unload Me

End Sub



'
' build dependeny list
'
Private Function build_dependencies(ByVal index As Long, ByRef list() As classFrameBuffer, ByVal valid As Long) As Long

  build_dependencies = 1 ' success by default
  With node(index)

    ' build dependecies
    Dim n As Long: n = .kernel.get_inputs
    If (n <> 0) Then

      ' allocate array
      Erase list(): ReDim list(0 To n - 1) As classFrameBuffer

      ' bind processors
      Dim i As Long: i = 0
      Do While (i < n)

        Dim src As Long: src = .socket(i)

        If (src <> -1) Then ' input connected?

          ' validate socket
          If (valid <> 0) Then

            If (node(src).valid <> 0) Then ' source node has vaild output?
              Set list(i) = node(src).kernel.buffer ' yep, bind validated reference
            Else
              Set list(i) = Nothing        ' no, unusable reference
              build_dependencies = 0       ' failure, a connected socket did not receive vailid data
            End If

          Else
            Set list(i) = node(src).kernel.buffer   ' bind reference (no validation required)
          End If

        Else
          Set list(i) = Nothing            ' unconnected socket (still can succeed if it is an optional input)
        End If

        i = i + 1 ' next input socket
      Loop

    Else ' no dependencies

      Erase list(): ReDim list(0 To 0) As classFrameBuffer
      Set list(0) = Nothing

    End If

  End With

End Function
