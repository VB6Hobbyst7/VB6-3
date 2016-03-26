Attribute VB_Name = "mEnvironment"
'================================================
' Module:        mEnvironment
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  clsFire.cls
' Last revision: 2006.06.26
'================================================

Option Explicit
Option Base 0

Public Type tColor
  R As Single
  G As Single
  B As Single
End Type

Public Type tTextureCoord
  U As Single
  V As Single
End Type

Public Type tFace
  A As Single
  B As Single
  C As Single
End Type

Public Type tVertex
  Position As D3DVECTOR
  normal As D3DVECTOR
  color As Long
  texture As tTextureCoord
End Type

Public Type tObject
  id As String
  Position As D3DVECTOR
  vertex() As tVertex
  face() As tFace
  normals_present As Boolean
  texture_present As Boolean
  map As String
  emissive As tColor
  vertex_stream() As tVertex
  vertex_buffer As Direct3DVertexBuffer8
  texture_buffer As Direct3DTexture8
  material_reference As Long
End Type

' Shader for static objects
Public Const lngStaticShader As Long = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
' Misc variables for staic objects
Public tStMat As D3DMATERIAL8
Public lngStVLen As Long
Private lngIndex As Long
Public lngOffset As Long

' All static meshes are stored here
Public tObjectList() As tObject

Public Sub BootEnv(SceneFile As String)
  frmSetup.Enabled = False
  DSX.smStop 1
  DSX.smStop 2
  DSX.smStop 3
  DSX.smStop 5
  DSX.smPlay 4
  ' Load static meshes
  ImportASE SceneFile
  ' Load all required textures and create data streams
  For lngIndex = 0 To UBound(tObjectList()) Step 1
    With tObjectList(lngIndex)
      ' Load texture if needed
      If Not .map = vbNullString Then
        ' load picture from package
        ReDim texRaw(0) As Byte
        If Not cPKG.pkExtract(.map, texRaw()) Then
          Set .texture_buffer = Nothing
          MsgBox "File can not be found in package. Unable to create binary stream." & vbCrLf & cPKG.pkNameHandle & " > " & .map, vbExclamation + vbOKOnly, "Warning"
        Else
          Set .texture_buffer = cD3DHLP.CreateTextureFromFileInMemoryEx(cD3DDev, texRaw(0), UBound(texRaw()) + 1, -1, -1, 0, 0, D3DFMT_X8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        End If
        Erase texRaw()
      Else
        'No Texture
        Set .texture_buffer = Nothing
      End If
      ' Create Buffer For Vertices
      ReDim .vertex_stream(UBound(.face()) * 3 + 2)
      ' enerate Vertex Stream
      For lngOffset = 0 To UBound(.face()) Step 1
        'also apply custom position & scale
        Select Case SceneFile
          Case "mesh\fxScene_01.ase"
            .vertex_stream(lngOffset * 3 + 0) = ProcessVertex(.vertex(.face(lngOffset).A), 0.3, 0.3, 0.3, 0, 24, 0)
            .vertex_stream(lngOffset * 3 + 1) = ProcessVertex(.vertex(.face(lngOffset).B), 0.3, 0.3, 0.3, 0, 24, 0)
            .vertex_stream(lngOffset * 3 + 2) = ProcessVertex(.vertex(.face(lngOffset).C), 0.3, 0.3, 0.3, 0, 24, 0)
          Case "mesh\fxScene_02.ase"
            .vertex_stream(lngOffset * 3 + 0) = ProcessVertex(.vertex(.face(lngOffset).A), 1.3, 1.3, 1.3, 0, 29, 0)
            .vertex_stream(lngOffset * 3 + 1) = ProcessVertex(.vertex(.face(lngOffset).B), 1.3, 1.3, 1.3, 0, 29, 0)
            .vertex_stream(lngOffset * 3 + 2) = ProcessVertex(.vertex(.face(lngOffset).C), 1.3, 1.3, 1.3, 0, 29, 0)
          Case "mesh\fxScene_03.ase"
            .vertex_stream(lngOffset * 3 + 0) = ProcessVertex(.vertex(.face(lngOffset).A), 1, 1, 1, 0, 47, 0)
            .vertex_stream(lngOffset * 3 + 1) = ProcessVertex(.vertex(.face(lngOffset).B), 1, 1, 1, 0, 47, 0)
            .vertex_stream(lngOffset * 3 + 2) = ProcessVertex(.vertex(.face(lngOffset).C), 1, 1, 1, 0, 47, 0)
          Case "mesh\fxScene_04.ase"
            .vertex_stream(lngOffset * 3 + 0) = ProcessVertex(.vertex(.face(lngOffset).A), 1, 1, 1, 0, 15, 0)
            .vertex_stream(lngOffset * 3 + 1) = ProcessVertex(.vertex(.face(lngOffset).B), 1, 1, 1, 0, 15, 0)
            .vertex_stream(lngOffset * 3 + 2) = ProcessVertex(.vertex(.face(lngOffset).C), 1, 1, 1, 0, 15, 0)
          Case "mesh\fxScene_05.ase"
            .vertex_stream(lngOffset * 3 + 0) = ProcessVertex(.vertex(.face(lngOffset).A), 0.3, 0.3, 0.3, 0, 0, 0)
            .vertex_stream(lngOffset * 3 + 1) = ProcessVertex(.vertex(.face(lngOffset).B), 0.3, 0.3, 0.3, 0, 0, 0)
            .vertex_stream(lngOffset * 3 + 2) = ProcessVertex(.vertex(.face(lngOffset).C), 0.3, 0.3, 0.3, 0, 0, 0)
          Case "mesh\fxScene_06.ase"
            .vertex_stream(lngOffset * 3 + 0) = ProcessVertex(.vertex(.face(lngOffset).A), 0.4, 0.3, 0.4, 0, 0, 0)
            .vertex_stream(lngOffset * 3 + 1) = ProcessVertex(.vertex(.face(lngOffset).B), 0.4, 0.3, 0.4, 0, 0, 0)
            .vertex_stream(lngOffset * 3 + 2) = ProcessVertex(.vertex(.face(lngOffset).C), 0.4, 0.3, 0.4, 0, 0, 0)
          Case "mesh\fxScene_07.ase"
            .vertex_stream(lngOffset * 3 + 0) = ProcessVertex(.vertex(.face(lngOffset).A), 0.3, 0.3, 0.3, 0, 1, 0)
            .vertex_stream(lngOffset * 3 + 1) = ProcessVertex(.vertex(.face(lngOffset).B), 0.3, 0.3, 0.3, 0, 1, 0)
            .vertex_stream(lngOffset * 3 + 2) = ProcessVertex(.vertex(.face(lngOffset).C), 0.3, 0.3, 0.3, 0, 1, 0)
          Case "mesh\fxScene_08.ase"
            .vertex_stream(lngOffset * 3 + 0) = ProcessVertex(.vertex(.face(lngOffset).A), 1, 1, 1, 0, 30, 0)
            .vertex_stream(lngOffset * 3 + 1) = ProcessVertex(.vertex(.face(lngOffset).B), 1, 1, 1, 0, 30, 0)
            .vertex_stream(lngOffset * 3 + 2) = ProcessVertex(.vertex(.face(lngOffset).C), 1, 1, 1, 0, 30, 0)
        End Select
      Next lngOffset
      ' Determine Vertex Data Length
      lngStVLen = Len(tObjectList(0).vertex_stream(0))
      ' Update VertexBuffer
      Set .vertex_buffer = cD3DDev.CreateVertexBuffer(lngStVLen * (UBound(.vertex_stream()) + 1), 0, lngStaticShader, D3DPOOL_DEFAULT)
      D3DVertexBuffer8SetData .vertex_buffer, 0, lngStVLen * (UBound(.vertex_stream()) + 1), 0, .vertex_stream(0)
    End With
  Next lngIndex
  ' Release effects
  For lngNumber = 0 To UBound(cCore()) Step 1
    If ObjPtr(cCore(lngNumber)) Then cCore(lngNumber).Release
  Next lngNumber
  ' Init effects and setup camera
  Select Case SceneFile
    Case "mesh\fxScene_01.ase"
      'preset - FirePit
      sngCamFinalDistance = 75
      sngCamFinalHeight = 20
      ReDim cCore(0)
      Set cCore(0) = New clsFire
      With cCore(0)
        .lngLayersCount = 128
        .sFlameTexture = "texture\flame\fxFlame_01.dds"
        .sngRadius = 10
        .sngPositionX = 0
        .sngPositionY = 8
        .sngPositionZ = 0
        .sngLightIndex = 0
        .sngDecRadius = 0.15
        .sngIncHeight = 0.3
        .sngWind = 0
        .sngCompression = 1
        .sngRed = 1
        .sngGreen = 1
        .sngBlue = 1
        .sngLightRed = 1
        .sngLightGreen = 0.8
        .sngLightBlue = 0.5
        .sngLightRange = 1000
        .Handle cD3DDev, cD3DHLP, cPKG
        .Initialize
      End With
    Case "mesh\fxScene_02.ase"
      'preset - Candle (3 Emmiters)
      sngCamFinalDistance = 60
      sngCamFinalDistance = 5
      ReDim cCore(2)
      For lngNumber = 0 To UBound(cCore()) Step 1
        Set cCore(lngNumber) = New clsFire
        With cCore(lngNumber)
          .lngLayersCount = 128
          .sFlameTexture = "texture\flame\fxFlame_02.dds"
          .sngRadius = 4
          .sngLightIndex = lngNumber
          .sngDecRadius = 0.05
          .sngIncHeight = 0.1
          .sngWind = 0
          .sngCompression = 0.945
          .sngRed = 1
          .sngGreen = 0.9
          .sngBlue = 0.9
          .sngLightRed = 0.5
          .sngLightGreen = 0.3
          .sngLightBlue = 0.2
          .sngLightRange = 1000
          .Handle cD3DDev, cD3DHLP, cPKG
          .Initialize
        End With
      Next lngNumber
      'set effect positions
      With cCore(0)
        .sngPositionX = 4.5 * 1.3
        .sngPositionY = 23
        .sngPositionZ = 7.2 * 1.3
      End With
      With cCore(1)
        .sngPositionX = 1.3 * 1.3
        .sngPositionY = 16
        .sngPositionZ = -5.9 * 1.3
      End With
      With cCore(2)
        .sngPositionX = -6.5 * 1.3
        .sngPositionY = 14
        .sngPositionZ = 1.1 * 1.3
      End With
    Case "mesh\fxScene_03.ase"
      'preset - StreetLamp
      sngCamFinalDistance = 90
      sngCamFinalHeight = 0
      ReDim cCore(0)
      Set cCore(0) = New clsFire
      With cCore(0)
        .lngLayersCount = 128
        .sFlameTexture = "texture\flame\fxFlame_02.dds"
        .sngRadius = 5
        .sngPositionX = 0
        .sngPositionY = 8
        .sngPositionZ = 0
        .sngLightIndex = 0
        .sngDecRadius = 0.05
        .sngIncHeight = 0.1
        .sngWind = 0
        .sngCompression = 1.01
        .sngRed = 0.5
        .sngGreen = 0.5
        .sngBlue = 1
        .sngLightRed = 0.6
        .sngLightGreen = 0.6
        .sngLightBlue = 1
        .sngLightRange = 1000
        .Handle cD3DDev, cD3DHLP, cPKG
        .Initialize
      End With
    Case "mesh\fxScene_04.ase"
      'preset - Troch
      sngCamFinalDistance = 40
      sngCamFinalHeight = 30
      ReDim cCore(0)
      Set cCore(0) = New clsFire
      With cCore(0)
        .lngLayersCount = 128
        .sFlameTexture = "texture\flame\fxFlame_01.dds"
        .sngRadius = 5
        .sngPositionX = 0
        .sngPositionY = 33
        .sngPositionZ = 0
        .sngLightIndex = 0
        .sngDecRadius = 0.15
        .sngIncHeight = 0.3
        .sngWind = 0
        .sngCompression = 1
        .sngRed = 0.9
        .sngGreen = 0.7
        .sngBlue = 0.7
        .sngLightRed = 0.9
        .sngLightGreen = 0.8
        .sngLightBlue = 0.7
        .sngLightRange = 1000
        .Handle cD3DDev, cD3DHLP, cPKG
        .Initialize
      End With
    Case "mesh\fxScene_05.ase"
      'preset - Furn
      sngCamFinalDistance = 35
      sngCamFinalHeight = 7
      ReDim cCore(0)
      Set cCore(0) = New clsFire
      With cCore(0)
        .lngLayersCount = 128
        .sFlameTexture = "texture\flame\fxFlame_01.dds"
        .sngRadius = 10
        .sngPositionX = 0
        .sngPositionY = 23
        .sngPositionZ = 0
        .sngLightIndex = 0
        .sngDecRadius = 0.15
        .sngIncHeight = 0.3
        .sngWind = 0
        .sngCompression = 1
        .sngRed = 0.9
        .sngGreen = 0.7
        .sngBlue = 0.7
        .sngLightRed = 0.9
        .sngLightGreen = 0.8
        .sngLightBlue = 0.7
        .sngLightRange = 1000
        .Handle cD3DDev, cD3DHLP, cPKG
        .Initialize
      End With
    Case "mesh\fxScene_06.ase"
      'preset - Lamp
      sngCamFinalDistance = 30
      sngCamFinalHeight = 15
      ReDim cCore(0)
      Set cCore(0) = New clsFire
      With cCore(0)
        .lngLayersCount = 128
        .sFlameTexture = "texture\flame\fxFlame_02.dds"
        .sngRadius = 0.7
        .sngPositionX = 0
        .sngPositionY = 35.5
        .sngPositionZ = 0
        .sngLightIndex = 0
        .sngDecRadius = 0.05
        .sngIncHeight = 0.1
        .sngWind = 0
        .sngCompression = 0.945
        .sngRed = 1
        .sngGreen = 1
        .sngBlue = 0.9
        .sngLightRed = 0.5
        .sngLightGreen = 0.5
        .sngLightBlue = 0.3
        .sngLightRange = 1000
        .Handle cD3DDev, cD3DHLP, cPKG
        .Initialize
      End With
    Case "mesh\fxScene_07.ase"
      'preset - GroundFire
      sngCamFinalDistance = 25
      sngCamFinalHeight = 1
      ReDim cCore(0)
      Set cCore(0) = New clsFire
      With cCore(0)
        .lngLayersCount = 128
        .sFlameTexture = "texture\flame\fxFlame_01.dds"
        .sngRadius = 10
        .sngPositionX = 0
        .sngPositionY = -5
        .sngPositionZ = 0
        .sngLightIndex = 0
        .sngDecRadius = 0.15
        .sngIncHeight = 0.3
        .sngWind = 0
        .sngCompression = 1
        .sngRed = 0.9
        .sngGreen = 0.7
        .sngBlue = 0.7
        .sngLightRed = 0.9
        .sngLightGreen = 0.8
        .sngLightBlue = 0.7
        .sngLightRange = 1000
        .Handle cD3DDev, cD3DHLP, cPKG
        .Initialize
      End With
    Case "mesh\fxScene_08.ase"
      'preset - Chandelier (8 Emmiters)
      sngCamFinalDistance = 50
      sngCamFinalHeight = 30
      ReDim cCore(7)
      For lngNumber = 0 To UBound(cCore()) Step 1
        Set cCore(lngNumber) = New clsFire
        With cCore(lngNumber)
          .lngLayersCount = 64
          .sFlameTexture = "texture\flame\fxFlame_02.dds"
          .sngRadius = 2
          .sngLightIndex = lngNumber
          .sngDecRadius = 0.05
          .sngIncHeight = 0.1
          .sngWind = 0
          .sngCompression = 0.945
          .sngRed = 1
          .sngGreen = 0.9
          .sngBlue = 0.9
          .sngLightRed = 0.2
          .sngLightGreen = 0.15
          .sngLightBlue = 0.1
          .sngLightRange = 1000
          .Handle cD3DDev, cD3DHLP, cPKG
          .Initialize
        End With
      Next lngNumber
      'set effect positions
      With cCore(0)
        .sngPositionX = 0
        .sngPositionY = 20
        .sngPositionZ = 25
      End With
      With cCore(1)
        .sngPositionX = -17.5
        .sngPositionY = 20
        .sngPositionZ = 17.5
      End With
      With cCore(2)
        .sngPositionX = -25
        .sngPositionY = 20
        .sngPositionZ = 0
      End With
      With cCore(3)
        .sngPositionX = -17.5
        .sngPositionY = 20
        .sngPositionZ = -17.5
      End With
      With cCore(4)
        .sngPositionX = 0
        .sngPositionY = 20
        .sngPositionZ = -25
      End With
      With cCore(5)
        .sngPositionX = 17.5
        .sngPositionY = 20
        .sngPositionZ = -17.5
      End With
      With cCore(6)
        .sngPositionX = 25
        .sngPositionY = 20
        .sngPositionZ = 0
      End With
      With cCore(7)
        .sngPositionX = 17.5
        .sngPositionY = 20
        .sngPositionZ = 17.5
      End With
  End Select
  
  ' Refresh gui
  frmSetup.cmdTopView.Caption = "Top"
  GUI_SetScrolls
  frmSetup.Enabled = True
End Sub

Public Sub GUI_SetScrolls()
  ' Reset gui controls for current preset
  With frmSetup
    If .cmbEnv.ListIndex = 0 Or .cmbEnv.ListIndex = 3 Or .cmbEnv.ListIndex = 4 Or .cmbEnv.ListIndex = 6 Then
      .scrCmp.Value = 0
      .scrFade.Value = 30
      .scrHeight.Value = 30
      .scrWind.Value = 0
    End If
    If .cmbEnv.ListIndex = 1 Or .cmbEnv.ListIndex = 5 Or .cmbEnv.ListIndex = 7 Then
      .scrCmp.Value = -35
      .scrFade.Value = 10
      .scrHeight.Value = 10
      .scrWind.Value = 0
    End If
    If .cmbEnv.ListIndex = 2 Then
      .scrCmp.Value = 13
      .scrFade.Value = 10
      .scrHeight.Value = 20
      .scrWind.Value = 0
    End If
    Select Case .cmbEnv.ListIndex
      Case 0: DSX.smPlay 1
      Case 1: DSX.smPlay 3
      Case 2: DSX.smPlay 3
      Case 3: DSX.smPlay 2
      Case 4: DSX.smPlay 1
      Case 5: DSX.smPlay 3
      Case 6: DSX.smPlay 1
      Case 7: DSX.smPlay 5
    End Select
    .cmdReset.Enabled = False
  End With
End Sub

