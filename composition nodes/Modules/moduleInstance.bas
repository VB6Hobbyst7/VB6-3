Attribute VB_Name = "moduleInstance"

' require variable declaration
Option Explicit


' number of node processing modules
Public Const instance_count As Long = 16



'
' create node processor instance by given type
'
Public Function instance_create(ByVal id As Long) As Object

  ' choose node processor by given type
  Select Case (id)

    ' type 0 - bitmap import
    'Case Is = 0: Set instance_create = New classNodeBitmapImport

    ' type 1 - bitmap export
    'Case Is = 1: Set instance_create = New classNodeBitmapExport

    ' type 2 - -----------------------------------------------------------

    ' type 3 - luminosity / contrast
    Case Is = 3: Set instance_create = New classNodeLuminosityContrast

    ' type 4 - shift h.s.l.
    Case Is = 4: Set instance_create = New classNodeShiftHSL

    ' type 5 - box blur
    Case Is = 5: Set instance_create = New classNodeBoxBlur

    ' type 6 - image transform
    Case Is = 6: Set instance_create = New classNodeImageTransform

    ' type 7 - remap channels
    Case Is = 7: Set instance_create = New classNodeRemapChannels

    ' type 8 - colorize
    'Case Is = 8: Set instance_create = New classNodeColorize

    ' type 9 - mix layers
    Case Is = 9: Set instance_create = New classNodeMixLayers

    ' type 10 - -----------------------------------------------------------

    ' type 11 - uniform color
    Case Is = 11: Set instance_create = New classNodeUniformColor

    ' type 12 - checkers
    Case Is = 12: Set instance_create = New classNodeCheckers

    ' type 13 - gradient
    Case Is = 13: Set instance_create = New classNodeGradient

    ' type 14 - env. map
    Case Is = 14: Set instance_create = New classNodeEnvMap

    ' type 15 - noise
    Case Is = 15: Set instance_create = New classNodeNoise

    ' unknown type
    Case Else: Set instance_create = Nothing
  End Select

End Function



'
' get node processor title
'
Public Function instance_title(ByVal id As Long) As String

  ' create instance
  Dim processor As Object: Set processor = instance_create(id)
  If (processor Is Nothing) Then
    instance_title = vbNullString        ' unknown type
  Else

    instance_title = processor.get_title ' get title
    Set processor = Nothing              ' destroy instance

  End If

End Function
