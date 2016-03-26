VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Module:        clsFire
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  clsPackage.cls
' Last revision: 2006.06.23
'================================================

'------- FX FIRE CLASS -------

'TO DO: - Light Attenuations (Depend On Light Range)
'       - Light Pulsing Modes


Option Base 0
Option Explicit

Private q As Single

' Effect configuration
Public sngPitch As Single
Public sngAngle As Single
Public lngLayersCount As Long
Public sFlameTexture As String
Public sngRadius As Single
Public sngPositionX As Single
Public sngPositionY As Single
Public sngPositionZ As Single
Public sngLightIndex As Long
Public sngIncHeight As Single
Public sngDecRadius As Single
Public sngCompression As Single
Public sngWind As Single
Public sngRed As Single
Public sngGreen As Single
Public sngBlue As Single
Public sngLightRange As Single
Public sngLightRed As Single
Public sngLightGreen As Single
Public sngLightBlue As Single

Private cPK As clsPackage

Private cHLD3DDev As Direct3DDevice8
Private cHLD3DHLP As D3DX8

' Effect variables and objects
Private Const lngShader As Long = D3DFVF_XYZ Or D3DFVF_TEX1
Private Type tVertex
  posX As Single
  posY As Single
  posZ As Single
  txU As Single
  txV As Single
End Type
Private tData() As D3DVECTOR4
Private tStream() As tVertex
Private cBuffer As Direct3DVertexBuffer8
Private cTexture As Direct3DTexture8
Private lngVertexLen As Long
Private tMaterial As D3DMATERIAL8
Private tLight As D3DLIGHT8
Private tWorldMat As D3DMATRIX
Private tWorldMatBackup As D3DMATRIX
Private sngCompress As Single

' Counter
Private lngI As Long
Private lngJ As Single

' Capture DirectX objects & packager
Public Sub Handle(cD3DDev As Direct3DDevice8, cD3DHLP As D3DX8, Packager As clsPackage)
  Set cHLD3DDev = cD3DDev
  Set cHLD3DHLP = cD3DHLP
  Set cPK = Packager
End Sub

Public Sub Initialize()

On Error Resume Next

  ' Load flame texture
  ReDim raw(0) As Byte
  If Not cPK.pkExtract(sFlameTexture, raw()) Then
    Set cTexture = Nothing
    MsgBox "File can not be found in package. Unable to create binary stream." & vbCrLf & cPKG.pkNameHandle & " > " & sFlameTexture, vbExclamation + vbOKOnly, "Warning"
  Else
    Set cTexture = cHLD3DHLP.CreateTextureFromFileInMemoryEx(cHLD3DDev, raw(0), UBound(raw()) + 1, -1, -1, 0, 0, D3DFMT_X8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
   
    If Not Err.Number = 0 Then
      Err.Clear
      Set cTexture = Nothing
      MsgBox "Unable to load texture: " & sFlameTexture, vbExclamation + vbOKOnly, "Warning"
    End If
  End If
  ' Create initial data
  ReDim tData(lngLayersCount - 1)
  ReDim tStream((lngLayersCount - 1) * 6 * 2)
  For lngI = 0 To lngLayersCount - 1 Step 1
    With tData(lngI)
      .X = (sngRadius / 2) - Rnd * sngRadius
      .Y = 0
      .Z = Rnd * sngRadius
      .w = 1 - Rnd * 2
    End With
  Next lngI
  ' Calculate vertex length
  lngVertexLen = Len(tStream(0))
  sngCompress = sngRadius / (lngLayersCount - 1)
  ' Process first frame
  Process
End Sub

Public Sub Render()
  'default emissive material
  With tMaterial
    .emissive.A = 1
    .emissive.R = sngRed
    .emissive.G = sngGreen
    .emissive.B = sngBlue
  End With
  'light setup
  With tLight
    .Type = D3DLIGHT_POINT
    'try to experiment with attenuations to change light radius...
    .Attenuation0 = 0.07 + Sin(Timer * 7) * 0.07
    .Attenuation1 = 0.01 + Cos(Timer) * 0.005
    .Attenuation2 = Abs(Sin(Timer) * 0.0001)
    .Position.X = sngPositionX
    .Position.Y = sngPositionY + 10
    .Position.Z = sngPositionZ
    .Range = sngLightRange
    .diffuse.A = 1
    .diffuse.R = sngLightRed
    .diffuse.G = sngLightGreen
    .diffuse.B = sngLightBlue
    .specular = .diffuse
  End With
  With cHLD3DDev
    'enable light
    .SetLight sngLightIndex, tLight
    .LightEnable sngLightIndex, 1
    'set fire material
    .SetMaterial tMaterial
    'flame texture
    .SetTexture 0, cTexture
    'filtering (fast)
    .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
    .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_ANISOTROPIC
    'alpha
    .SetRenderState D3DRS_ALPHABLENDENABLE, 1
    .SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTALPHA
    .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    'triangle culling
    .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    'get world matrix
    .GetTransform D3DTS_WORLD, tWorldMatBackup
    tWorldMat = tWorldMatBackup
    'fire always faces the camera
    D3DXMatrixRotationY tWorldMat, sngAngle
    'set fire position
    With tWorldMat
      .m41 = sngPositionX
      .m42 = sngPositionY
      .m43 = sngPositionZ
    End With
    .SetTransform D3DTS_WORLD, tWorldMat
    'render fire
    .SetVertexShader lngShader
    .SetStreamSource 0, cBuffer, lngVertexLen
    .SetRenderState D3DRS_ZWRITEENABLE, 0
    .DrawPrimitive D3DPT_TRIANGLELIST, 0, (UBound(tStream()) + 1) / 3
    .SetRenderState D3DRS_ZWRITEENABLE, 1
    'turn back settings
    .SetTransform D3DTS_WORLD, tWorldMatBackup
    .SetTexture 0, Nothing
    .SetRenderState D3DRS_ALPHABLENDENABLE, 0
  End With
End Sub

Public Sub Process()
  Dim Ca As Single
  Dim Sa As Single
  'Prepare Angles
  Ca = Cos(sngPitch)
  Sa = Sin(sngPitch)
  'process fire particles
  For lngI = 0 To lngLayersCount - 1 Step 1
    With tData(lngI)
      'radius check
      If .Z < 0.1 Then
        'too small? reset this particle
        .Y = 0
        .X = (sngRadius / 2) - Rnd * sngRadius
        .Z = sngRadius
      End If
      'position
      .X = .X * sngCompression + sngWind
      .Y = .Y + sngIncHeight
      .Z = .Z - sngDecRadius
      If .w < 0 Then
        .w = .w - Rnd * 0.05 - 0.05
      Else
        .w = .w + Rnd * 0.05 + 0.05
      End If
      q = .w
      'layers z-order
      lngJ = lngI * sngCompress - (sngRadius / 2)
      'create 2 triangles
      Transform Ca, Sa, MakeVertex(.X - .Z / 2, .Y - .Z / 2, lngJ, 0, 0), tStream(lngI * 6 + 0), .X, .Y, lngJ
      Transform Ca, Sa, MakeVertex(.X + .Z / 2, .Y - .Z / 2, lngJ, 1, 0), tStream(lngI * 6 + 1), .X, .Y, lngJ
      Transform Ca, Sa, MakeVertex(.X - .Z / 2, .Y + .Z / 2, lngJ, 0, -1), tStream(lngI * 6 + 2), .X, .Y, lngJ
      Transform Ca, Sa, MakeVertex(.X + .Z / 2, .Y - .Z / 2, lngJ, 1, 0), tStream(lngI * 6 + 3), .X, .Y, lngJ
      Transform Ca, Sa, MakeVertex(.X + .Z / 2, .Y + .Z / 2, lngJ, 1, -1), tStream(lngI * 6 + 4), .X, .Y, lngJ
      Transform Ca, Sa, MakeVertex(.X - .Z / 2, .Y + .Z / 2, lngJ, 0, -1), tStream(lngI * 6 + 5), .X, .Y, lngJ
    End With
  Next lngI
  'reload vertex buffer
  Set cBuffer = cHLD3DDev.CreateVertexBuffer(lngVertexLen * (UBound(tStream()) + 1), 0, lngShader, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData cBuffer, 0, lngVertexLen * (UBound(tStream()) + 1), 0, tStream(0)
End Sub

Public Sub Release()
  On Error Resume Next
  ' Disable light
  cHLD3DDev.LightEnable sngLightIndex, 0
  ' Cleanup
  Erase tData()
  Erase tStream()
  ' Release effect objects
  Set cTexture = Nothing
  Set cBuffer = Nothing
  ' Release DirectX objects
  Set cHLD3DDev = Nothing
  Set cHLD3DHLP = Nothing
  ' Release packager
  Set cPK = Nothing
End Sub

Private Function MakeVertex(X As Single, Y As Single, Z As Single, U As Single, V As Single) As tVertex
  With MakeVertex
    .posX = X
    .posY = Y
    .posZ = Z
    .txU = U
    .txV = V
  End With
End Function

' Rotate vertex by Z axis
Private Function Transform(Ca As Single, Sa As Single, src As tVertex, dst As tVertex, cx As Single, cy As Single, cz As Single)
  Dim Yy As Single
  'transformation
  dst = src
  With dst
    'subtract center
    .posX = .posX - cx
    .posY = .posY - cy
    .posZ = .posZ - cz
    'rotate by Z axis
    Yy = .posY * Ca - .posZ * Sa
    .posZ = .posY * Sa + .posZ * Ca
    .posY = Yy
    'add center
    .posX = .posX + cx
    .posY = .posY + cy
    .posZ = .posZ + cz
    If Not bTexRotate Then Exit Function
    'texture rotation (more realistic sometimes)
    .txU = (.txU - 0.5) * 2
    .txV = (.txV - 0.5) * 2
    If .txV = 1 And .txU = 1 Then
      .txU = Sin(q + 45 * 3.14 / 180)
      .txV = Cos(q + 45 * 3.14 / 180)
    Else
      If .txV = -1 And .txU = -1 Then
        .txU = Sin(q + 225 * 3.14 / 180)
        .txV = Cos(q + 225 * 3.14 / 180)
      Else
        If .txV = 1 And .txU = -1 Then
          .txU = Sin(q + 135 * 3.14 / 180)
          .txV = Cos(q + 135 * 3.14 / 180)
        Else
          If .txV = -1 And .txU = 1 Then
            .txU = Sin(q + 315 * 3.14 / 180)
            .txV = Cos(q + 315 * 3.14 / 180)
          End If
        End If
      End If
    End If
    .txU = .txU / 2 + 0.5
    .txV = .txV / 2 + 0.5
    .txU = .txU * 1.1
    .txV = .txV * 1.1
  End With
End Function
