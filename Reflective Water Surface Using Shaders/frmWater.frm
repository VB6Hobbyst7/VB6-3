VERSION 5.00
Begin VB.Form frmWater 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refective Water Surface Using Shaders"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   ForeColor       =   &H80000008&
   LinkTopic       =   "Handle"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Frame 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmWater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Module:        frmWater
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  None
' Last revision: 2007.03.19
'================================================

Option Explicit
Option Base 0


'general objects
Private objDX As DirectX8
Private objD3D As Direct3D8
Private objD3DDev As Direct3DDevice8
Private objD3DHlp As D3DX8

'configuration structures
Private devDisplay As D3DDISPLAYMODE
Private devOptions As D3DPRESENT_PARAMETERS

'textures
Private texPalette As Direct3DTexture8
Private texHeight As Direct3DTexture8
Private texEnvironment As Direct3DCubeTexture8
Private texNormal As Direct3DTexture8

'normal map generation
Private texDesc As D3DSURFACE_DESC
Private texRectSrc As D3DLOCKED_RECT
Private texRectDst As D3DLOCKED_RECT

'pixel shader
Private shpProgram As String
Private shpCode As D3DXBuffer
Private shpLength As Long
Private shpArray() As Long
Private shpHandle As Long

'vertex shader
Private shvProgram As String
Private shvCode As D3DXBuffer
Private shvLength As Long
Private shvArray() As Long
Private shvHandle As Long
Private shvDeclare() As Long

'water surface settings
Private fxResH As Single
Private fxResV As Single
Private fxMinX As Single
Private fxMaxX As Single
Private fxMinY As Single
Private fxMaxY As Single

'vertex format
Private Type fmtVertex
  vecPos As D3DVECTOR  'position
  vecNorm As D3DVECTOR 'normal
  vecScl As D3DVECTOR  'wave height scale vector
  texU1 As Single      'texture coords
  texV1 As Single
  texW1 As Single
  T As D3DVECTOR
End Type

'vertex declaration
Private Const declVertex As Long = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1 Or D3DFVF_TEXCOORDSIZE3_0

'temp counters for vertex & index buffers generation
Private I As Long
Private Ii As Single
Private J As Long
Private Jj As Single
Private K As Long

'water surface objects & variables
Private numFaces As Long
Private numVertices As Long
Private arrIndex() As Long
Private arrVertex() As fmtVertex
Private mhIndex As Direct3DIndexBuffer8
Private mhVertex As Direct3DVertexBuffer8

'camera matrices
Private matTemp As D3DMATRIX
Private matWorld As D3DMATRIX
Private matView As D3DMATRIX
Private matProj As D3DMATRIX

'aspect ratio
Private aspectRatio As Single

'camera position & lookat
Private CamXPos As Single
Private CamYPos As Single
Private CamZPos As Single
Private CamXAt As Single
Private CamYAt As Single
Private CamZAt As Single
Private vecEye As D3DVECTOR
Private vecLookAt As D3DVECTOR
Private vecUp As D3DVECTOR

'pi constant
Private Const Pi As Single = 3.14159265358979

'global time (for animation)
Private glTime As Single

'vertex shader registers
Private c0 As D3DVECTOR4
Private c1 As D3DVECTOR4
Private c2 As D3DVECTOR4
Private c3 As D3DVECTOR4
Private c8 As D3DVECTOR4
Private c10 As D3DVECTOR4
Private c11 As D3DVECTOR4
Private c12 As D3DVECTOR4
Private c13 As D3DVECTOR4
Private c14 As D3DVECTOR4
Private c15 As D3DVECTOR4
Private c16 As D3DVECTOR4
Private c17 As D3DVECTOR4

'pixel shader registers
Private c0p As D3DVECTOR4
Private c1p As D3DVECTOR4

Private Sub Form_Load()
  
  'show the form
  Show
  DoEvents

  'create directx8 core objects
  Set objDX = New DirectX8
  Set objD3D = objDX.Direct3DCreate
  Set objD3DHlp = New D3DX8
  
  'aquire current display mode
  objD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, devDisplay
  
  'configure rendering device
  With devOptions
    .AutoDepthStencilFormat = D3DFMT_D16   '16-bit z-buffer
    .BackBufferCount = 1                   'only one backbuffer
    .BackBufferFormat = devDisplay.Format  'set current color depth
    .BackBufferHeight = ScaleHeight        'set current resolution
    .BackBufferWidth = ScaleWidth
    .EnableAutoDepthStencil = 1            'enable depth stencil
    .flags = 0                             'nothing special we could use here
    .FullScreen_PresentationInterval = 0   'we are not going to use fullscreen
    .FullScreen_RefreshRateInHz = 0
    .hDeviceWindow = hWnd                  'attach device to current window
    .MultiSampleType = D3DMULTISAMPLE_NONE 'no antialiasing
    .SwapEffect = D3DSWAPEFFECT_DISCARD    'discard
    .Windowed = 1                          'run in window
  End With

  'create rendering device
  Set objD3DDev = objD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_MIXED_VERTEXPROCESSING, devOptions)
  
  'load textures
  Set texPalette = objD3DHlp.CreateTextureFromFile(objD3DDev, App.Path & "\Texture Water Gradient.tga")
  Set texHeight = objD3DHlp.CreateTextureFromFile(objD3DDev, App.Path & "\Texture Height Map.dds")
  Set texEnvironment = objD3DHlp.CreateCubeTextureFromFile(objD3DDev, App.Path & "\Texture Environment Cube Map.dds")
  
  'version 1.4 pixel shader program code
  'input registers
  'c0: commonConst (0, 0.5, 1, 0.25)
  'c1: highlightColor (0.8, 0.76, 0.62, 1)
  shpProgram = "ps.1.4                       //need version 1.4 pixel shader           " & vbCrLf & _
               "texld r0, t0                 //bump map 0                              " & vbCrLf & _
               "texld r1, t1                 //sample bump map 1                       " & vbCrLf & _
               "texcrd r2.rgb, t2            //view vector                             " & vbCrLf & _
               "texcrd r3.rgb, t3            //binormal                                " & vbCrLf & _
               "texcrd r4.rgb, t4            //tangent                                 " & vbCrLf & _
               "texcrd r5.rgb, t5            //normal                                  " & vbCrLf & _
               "add_d2 r0.xy, r0, r1         //scaled average of 2 bumpmaps xy offsets " & vbCrLf & _
               "mul r1.rgb, r0.x, r3         //x offset                                " & vbCrLf & _
               "mad r1.rgb, r0.y, r4, r1     //y offset                                " & vbCrLf & _
               "mad r1.rgb, r0.z, r5, r1     //put bumpmap normal into world space     " & vbCrLf & _
               "dp3 r0.rgb, r1, r2           //v.n                                     " & vbCrLf & _
               "mad r2.rgb, r1, r0_x2, r2    //r=2n(v.n)-v                             " & vbCrLf & _
               "mov_sat r1, r0_x2            //2*n.n (sample over range of 1d map)     " & vbCrLf & _
               "phase                        //start rendering phase                   " & vbCrLf & _
               "texld r2, r2                 //cubic env map                           " & vbCrLf & _
               "texld r3, r1                 //index fresnel map using 2*v.n           " & vbCrLf & _
               "mul r2.rgb, r2, r2           //square the environment map              " & vbCrLf & _
               "+mul r2.a, r2.g, r2.g        //use green channel of env map as specular" & vbCrLf & _
               "mul r2.rgb, r2, 1-r0.r       //fresnel term                            " & vbCrLf & _
               "+mul r2.a, r2.a, r2.a        //specular highlight ^4                   " & vbCrLf & _
               "add_d4_sat r2.rgb, r2, r3_x2 //+=waterColor                            " & vbCrLf & _
               "+mul r2.a, r2.a, r2.a        //specular highlight ^8                   " & vbCrLf & _
               "mad_sat r0, r2.a, c1, r2     //+=Specular highlight*highlightColor     "

  'version 1.1 vertex shader program code
  'input registers
  'v0  : vertex position
  'v3  : vertex normal
  'v7  : vertex texture coords u,v
  'v8  : vertex tangent (v direction)
  'v5  : wave height Scale
  'c0  : { 0.0, 0.5, 1.0, 2.0}
  'c1  : { 4.0, 0.5*pi, pi, 2*pi}
  'c2  : {1, -1/3!, 1/5!, -1/7!}    (for sin)
  'c3  : {1/2!, -1/4!, 1/6!, -1/8!} (for cos)
  'c4-7: composite world-view-projection matrix
  'c8  : modelSpace camera position
  'c9  : modelSpace light position
  'c10 : {fixup factor for taylor series imprecision}         (1.02, 0.1, 0, 0)
  'c11 : {waveHeight0, waveHeight1, waveHeight2, waveHeight3} (80.0, 100.0, 5.0, 5.0)
  'c12 : {waveOffset0, waveOffset1, waveOffset2, waveOffset3} (0.0, 0.2, 0.0, 0.0)
  'c13 : {waveSpeed0, waveSpeed1, waveSpeed2, waveSpeed3}     (0.2, 0.15, 0.4, 0.4)
  'c14 : {waveDirX0, waveDirX1, waveDirX2, waveDirX3}         (0.25, 0.0, -0.7, -0.8)
  'c15 : {waveDirY0, waveDirY1, waveDirY2, waveDirY3}         (0.0, 0.15, -0.7, 0.1)
  'c16 : {time, sin(time)}
  'c17 : {basetexcoord distortion x0, y0, x1, y1}             (0.031, 0.04, -0.03, 0.02)
  'c18 : world martix
  shvProgram = "vs.1.1                        //need version 1.1 vertex shader                      " & vbCrLf & _
               "mul r0, c14, v7.x             //use tex coords as inputs to sinusoidal warp         " & vbCrLf & _
               "mad r0, c15, v7.y, r0         //use tex coords as inputs to sinusoidal warp         " & vbCrLf & _
               "mov r1, c16.x                 //time...                                             " & vbCrLf & _
               "mad r0, r1, c13, r0           //add scaled time to move bumps according to frequency" & vbCrLf & _
               "add r0, r0, c12               //starting time offset                                " & vbCrLf & _
               "frc r0.xy, r0                 //take frac of all 4 components                       " & vbCrLf & _
               "frc r1.xy, r0.zwzw            //                                                    " & vbCrLf & _
               "mov r0.zw, r1.xyxy            //                                                    " & vbCrLf & _
               "mul r0, r0, c10.x             //multiply by fixup factor (due to inaccuracy)        " & vbCrLf & _
               "sub r0, r0, c0.y              //subtract .5                                         " & vbCrLf & _
               "mul r0, r0, c1.w              //mult tex coords by 2pi  coords range from(-pi to pi)" & vbCrLf & _
               "mul r5, r0, r0                //(wave vec)^2                                        " & vbCrLf & _
               "mul r1, r5, r0                //(wave vec)^3                                        " & vbCrLf & _
               "mul r6, r1, r0                //(wave vec)^4                                        " & vbCrLf & _
               "mul r2, r6, r0                //(wave vec)^5                                        " & vbCrLf & _
               "mul r7, r2, r0                //(wave vec)^6                                        " & vbCrLf & _
               "mul r3, r7, r0                //(wave vec)^7                                        " & vbCrLf & _
               "mul r8, r3, r0                //(wave vec)^8                                        " & vbCrLf & _
               "mad r4, r1, c2.y, r0          //(wave vec)-((wave vec)^3)/3!                        " & vbCrLf & _
               "mad r4, r2, c2.z, r4          //+((wave vec)^5)/5!                                  " & vbCrLf & _
               "mad r4, r3, c2.w, r4          //-((wave vec)^7)/7!                                  " & vbCrLf & _
               "mov r0, c0.z                  //1                                                   " & vbCrLf & _
               "mad r5, r5, c3.x ,r0          //-(wave vec)^2/2!                                    " & vbCrLf
  shvProgram = shvProgram & _
               "mad r5, r6, c3.y, r5          //+(wave vec)^4/4!                                    " & vbCrLf & _
               "mad r5, r7, c3.z, r5          //-(wave vec)^6/6!                                    " & vbCrLf & _
               "mad r5, r8, c3.w, r5          //+(wave vec)^8/8!                                    " & vbCrLf & _
               "sub r0, c0.z, v5.x            //...1-wave scale                                     " & vbCrLf & _
               "mul r4, r4, r0                //scale sin                                           " & vbCrLf & _
               "mul r5, r5, r0                //scale cos                                           " & vbCrLf & _
               "dp4 r0, r4, c11               //multiply wave heights by waves                      " & vbCrLf & _
               "mul r0.xyz, v3, r0            //multiply wave magnitude at this vertex by normal    " & vbCrLf & _
               "add r0.xyz, r0, v0            //add to position                                     " & vbCrLf & _
               "mov r0.w, c0.z                //homogenous component                                " & vbCrLf & _
               "m4x4 oPos, r0, c4             //OutPos=ObjSpacePos*World-View-Projection Matrix     " & vbCrLf & _
               "mul r1, r5, c11               //cos*waveheight                                      " & vbCrLf & _
               "dp4 r9.x, -r1, c14            //normal x offset                                     " & vbCrLf & _
               "dp4 r9.yzw, -r1, c15          //normal y offset and tangent offset                  " & vbCrLf & _
               "mov r5, v3                    //starting normal                                     " & vbCrLf & _
               "mad r5.xy, r9, c10.y, r5      //warped normal move nx, ny according to              " & vbCrLf & _
               "mov r4, v8                    //tangent                                             " & vbCrLf & _
               "mad r4.z, -r9.x, c10.y, r4.z  //warped tangent vector                               " & vbCrLf & _
               "mov r10, r5                   //                                                    " & vbCrLf & _
               "m3x3 r5, r10, c18             //transform normal                                    " & vbCrLf & _
               "dp3 r10.x, r5, r5             //                                                    " & vbCrLf & _
               "rsq r10.y, r10.x              //                                                    " & vbCrLf & _
               "mul r5, r5, r10.y             //normalize normal                                    " & vbCrLf
  shvProgram = shvProgram & _
               "mov r10, r4                   //                                                    " & vbCrLf & _
               "m3x3 r4, r10, c18             //transform tangent                                   " & vbCrLf & _
               "dp3 r10.x, r4, r4             //                                                    " & vbCrLf & _
               "rsq r10.y, r10.x              //                                                    " & vbCrLf & _
               "mul r4, r4, r10.y             //normalize tangent                                   " & vbCrLf & _
               "mul r3, r4.yzxw, r5.zxyw      //                                                    " & vbCrLf & _
               "mad r3, r4.zxyw, -r5.yzxw, r3 //xprod to find binormal                              " & vbCrLf & _
               "mov r10, r0                   //                                                    " & vbCrLf & _
               "m4x4 r0, r10, c18             //transform vertex position                           " & vbCrLf & _
               "sub r2, c8,  r0               //view vector                                         " & vbCrLf & _
               "dp3 r10.x, r2, r2             //                                                    " & vbCrLf & _
               "rsq r10.y, r10.x              //                                                    " & vbCrLf & _
               "mul r2, r2, r10.y             //normalized view vector                              " & vbCrLf & _
               "mov r0, c16.x                 //                                                    " & vbCrLf & _
               "mul r0, r0, c17.xyxy          //                                                    " & vbCrLf & _
               "frc r0.xy, r0                 //frc of incoming time                                " & vbCrLf & _
               "add r0, v7, r0                //add time to tex coords                              " & vbCrLf & _
               "mov oT0, r0                   //distorted tex coord 0                               " & vbCrLf & _
               "mov r0, c16.x                 //                                                    " & vbCrLf & _
               "mul r0, r0, c17.zwzw          //                                                    " & vbCrLf & _
               "frc r0.xy, r0                 //frc of incoming time                                " & vbCrLf & _
               "add r0, v7, r0                //add time to tex coords                              " & vbCrLf & _
               "mov oT1, r0.yxzw              //distorted tex coord 1                               " & vbCrLf
  shvProgram = shvProgram & _
               "mov oT2, r2                   //pass in view vector (worldspace)                    " & vbCrLf & _
               "mov oT3, r3                   //binormal                                            " & vbCrLf & _
               "mov oT4, r4                   //tangent                                             " & vbCrLf & _
               "mov oT5, r5                   //normal                                              "

  'compile pixel shader
  Set shpCode = objD3DHlp.AssembleShader(shpProgram, 0, Nothing, vbNullString)
  shpLength = shpCode.GetBufferSize / 4
  ReDim shpArray(shpLength - 1) As Long
  objD3DHlp.BufferGetData shpCode, 0, 4, shpLength, shpArray(0)
  shpHandle = objD3DDev.CreatePixelShader(shpArray(0))
  
  'declare vertex shader
  ReDim shvDeclare(0 To 6) As Long
  shvDeclare(0) = (&H20000000 And &HE0000000) Or 0                   'begin token       -> d3dvsd_stream(0)
  shvDeclare(1) = (&H40000000 And &HE0000000) Or (2 * 2 ^ 16) Or (0) 'position          -> d3dvsd_reg(0,d3dvsdt_float3)
  shvDeclare(2) = (&H40000000 And &HE0000000) Or (2 * 2 ^ 16) Or (3) 'normal            -> d3dvsd_reg(3,d3dvsdt_float3)
  shvDeclare(3) = (&H40000000 And &HE0000000) Or (2 * 2 ^ 16) Or (5) 'wave height scale -> d3dvsd_reg(5,d3dvsdt_float3)
  shvDeclare(4) = (&H40000000 And &HE0000000) Or (2 * 2 ^ 16) Or (7) 'texture coords    -> d3dvsd_reg(7,d3dvsdt_float3)
  shvDeclare(5) = (&H40000000 And &HE0000000) Or (2 * 2 ^ 16) Or (8) 'tangent           -> d3dvsd_reg(8,d3dvsdt_float3)
  shvDeclare(6) = &HFFFFFFFF                                         'end token         -> d3dvsd_end
  
  'compile vertex shader
  Set shvCode = objD3DHlp.AssembleShader(shvProgram, 0, Nothing, vbNullString)
  shvLength = shvCode.GetBufferSize / 4
  ReDim shvArray(shvLength - 1) As Long
  objD3DHlp.BufferGetData shvCode, 0, 4, shvLength, shvArray(0)
  objD3DDev.CreateVertexShader shvDeclare(0), shvArray(0), shvHandle, 0
  
  'aquire source texture parameters
  texHeight.GetLevelDesc 0, texDesc
  'generate normal map
  Set texNormal = objD3DDev.CreateTexture(texDesc.Width, texDesc.Height, 1, 0, D3DFMT_V8U8, D3DPOOL_MANAGED)
  'unlock texture rects
  texNormal.LockRect 0, texRectDst, ByVal 0, 0
  texHeight.LockRect 0, texRectDst, ByVal 0, 0
  
  
  
  Set texNormal = texHeight
  
  
  
  'unlock texture rects
  texNormal.UnlockRect 0
  texHeight.UnlockRect 0
  
  'setup water surface
  fxResH = 120
  fxResV = 40
  fxMinX = -20
  fxMaxX = 20
  fxMinY = -10
  fxMaxY = 10
  
  'create water mesh index buffer
  numFaces = fxResH * fxResV * 2
  ReDim arrIndex(0 To numFaces * 3 - 1) As Long
  K = 0
  For I = 0 To fxResH - 1 Step 1
    For J = 0 To fxResV - 1 Step 1
      arrIndex(K + 0) = (I * (fxResV + 1) + J)
      arrIndex(K + 1) = ((I + 1) * (fxResV + 1) + J + 1)
      arrIndex(K + 2) = (I * (fxResV + 1) + J + 1)
      arrIndex(K + 3) = (I * (fxResV + 1) + J)
      arrIndex(K + 4) = ((I + 1) * (fxResV + 1) + J)
      arrIndex(K + 5) = ((I + 1) * (fxResV + 1) + J + 1)
      K = K + 6
    Next J
  Next I
  Set mhIndex = objD3DDev.CreateIndexBuffer(numFaces * 3 * Len(arrIndex(0)), D3DUSAGE_WRITEONLY, D3DFMT_INDEX32, D3DPOOL_DEFAULT)
  D3DIndexBuffer8SetData mhIndex, 0, numFaces * 3 * Len(arrIndex(0)), 0, arrIndex(0)
  
  'create water mesh vertex buffer
  numVertices = (fxResH + 1) * (fxResV + 1)
  ReDim arrVertex(0 To numVertices - 1) As fmtVertex
  K = 0
  For I = 0 To fxResH Step 1
    For J = 0 To fxResV Step 1
      Ii = CSng(I) / CSng(fxResH)
      Jj = CSng(J) / CSng(fxResV)
      With arrVertex(K)
        'position
        .vecPos.X = fxMinX + Ii * (fxMaxX - fxMinX)
        .vecPos.Y = fxMinY + Jj * (fxMaxY - fxMinY)
        .vecPos.z = 0
        'normal
        .vecNorm.X = 0
        .vecNorm.Y = 0
        .vecNorm.z = -1
        'texture
        .texU1 = Ii
        .texV1 = Jj
        .texW1 = 0
        .T.X = 0
        .T.Y = 1
        .T.z = 0
        'scale
        Ii = Abs(Ii - 0.5)
        Jj = Abs(Jj - 0.5)
        If Ii > Jj Then
          .vecScl.X = 2 * Ii
        Else
          .vecScl.X = 2 * Jj
        End If
        .vecScl.Y = 0
        .vecScl.z = 0
      End With
      K = K + 1
    Next J
  Next I
  Set mhVertex = objD3DDev.CreateVertexBuffer(numVertices * Len(arrVertex(0)), D3DUSAGE_WRITEONLY, 0, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData mhVertex, 0, numVertices * Len(arrVertex(0)), 0, arrVertex(0)
  
  'configure scene
  With objD3DDev
    .SetRenderState D3DRS_ZENABLE, 0             'no depth testing
    .SetRenderState D3DRS_LIGHTING, 0            'no lighting
    .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE 'no culling
    'setup texture filtering
    For K = 0 To 3 Step 1
      .SetTextureStageState K, D3DTSS_MINFILTER, D3DTEXF_LINEAR
      .SetTextureStageState K, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
      .SetTextureStageState K, D3DTSS_MIPFILTER, D3DTEXF_LINEAR
    Next K
  End With
  
  'setup camera
  CamXPos = 0
  CamYPos = 0
  CamZPos = -20
  CamXAt = 0
  CamYAt = 0
  CamZAt = 0
  aspectRatio = ScaleWidth / ScaleHeight
  vecEye.X = CamXPos
  vecEye.Y = CamYPos
  vecEye.z = CamZPos
  vecLookAt.X = CamXAt
  vecLookAt.Y = CamYAt
  vecLookAt.z = CamZAt
  vecUp.X = 0
  vecUp.Y = 1
  vecUp.z = 0
   
  
  'setup transform matrices
  D3DXMatrixIdentity matWorld
  D3DXMatrixLookAtLH matView, vecEye, vecLookAt, vecUp
  D3DXMatrixPerspectiveFovLH matProj, Pi / 2, aspectRatio, 1, 1000
  
  'reset time
  glTime = 0
  
  'start rendering
  With Frame
    .Interval = 10
    .Enabled = True
  End With
  
End Sub

'generate render frame
Private Sub Frame_Timer()
  
  With objD3DDev
    'clear z-buffer
    .Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H0, 1, 0
    'open renderer
    .BeginScene
    
    Static mat As D3DMATRIX
    Static matInitial As D3DMATRIX
    D3DXMatrixIdentity matInitial
    D3DXMatrixRotationX mat, Pi * 1.5
    D3DXMatrixMultiply matInitial, matInitial, mat
    D3DXMatrixTranslation mat, 0, -1.5, -20
    D3DXMatrixMultiply matInitial, matInitial, mat
    matWorld = matInitial
    
    'apply transformation
    .SetTransform D3DTS_WORLD, matWorld
    .SetTransform D3DTS_VIEW, matView
    .SetTransform D3DTS_PROJECTION, matProj
    
    'inc time
    glTime = glTime + 0.02
    
    'configure vertex shader
    c0 = vec4f(0, 0.5, 1, 2)
    c1 = vec4f(4, 0.5 * Pi, Pi, 2 * Pi)
    c2 = vec4f(1, -1 / 6, 1 / 120, -1 / 5040)
    c3 = vec4f(1 / 2, -1 / 24, 1 / 720, -1 / 40320)
    c8 = vec4f(CamXPos, CamYPos, CamZPos, 1)
    c10 = vec4f(1.02, 0.1, 0, 0)
    c11 = vec4f(0.4, 0.5, 0.025, 0.025)
    c12 = vec4f(0, 0.2, 0, 0)
    c13 = vec4f(0.2, 0.15, 0.4, 0.4)
    c14 = vec4f(2.5, 0, -7, -8)
    c15 = vec4f(0, 1.5, -7, 1)
    c16 = vec4f(glTime * 0.75, Sin(glTime), 0, 0)
    c17 = vec4f(0.031, 0.04, -0.03, 0.02)
    .SetVertexShaderConstant 0, c0, 1
    .SetVertexShaderConstant 1, c1, 1
    .SetVertexShaderConstant 2, c2, 1
    .SetVertexShaderConstant 3, c3, 1
    .SetVertexShaderConstant 8, c8, 1
    .SetVertexShaderConstant 10, c10, 1
    .SetVertexShaderConstant 11, c11, 1
    .SetVertexShaderConstant 12, c12, 1
    .SetVertexShaderConstant 13, c13, 1
    .SetVertexShaderConstant 14, c14, 1
    .SetVertexShaderConstant 15, c15, 1
    .SetVertexShaderConstant 16, c16, 1
    .SetVertexShaderConstant 17, c17, 1
    D3DXMatrixMultiply matTemp, matView, matProj
    D3DXMatrixMultiply matTemp, matWorld, matTemp
    D3DXMatrixTranspose matTemp, matTemp
    .SetVertexShaderConstant 4, matTemp, 4
    D3DXMatrixTranspose matTemp, matWorld
    .SetVertexShaderConstant 18, matTemp, 4
    
    'configure pixel shader
    c0p = vec4f(0, 0.5, 1, 0.25)
    c1p = vec4f(0.8, 0.76, 0.62, 1)
    .SetPixelShaderConstant 0, c0p, 1
    .SetPixelShaderConstant 1, c1p, 1
    
    'setup textures
    .SetTexture 0, texNormal
    .SetTexture 1, texNormal
    .SetTexture 2, texEnvironment
    .SetTexture 3, texPalette
    
    'use proper texture coords
    .SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
    
    'apply vertex shader & pixel shader
    .SetVertexShader shvHandle
    .SetPixelShader shpHandle
    
    'finally render water surface
    .SetStreamSource 0, mhVertex, Len(arrVertex(0))
    .SetIndices mhIndex, 0
    .DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, numVertices, 0, numFaces
    
    'close renderer
    .EndScene
    'swap buffers
    .Present ByVal 0, ByVal 0, 0, ByVal 0
  End With

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  'stop timer
  Frame.Enabled = False
  
  'dispose index & vertex buffers
  Erase arrVertex()
  Erase arrIndex()
  Set mhIndex = Nothing
  Set mhVertex = Nothing
  
  'kill vertex shader
  objD3DDev.DeleteVertexShader shvHandle
  Set shvCode = Nothing
  Erase shvArray()
  Erase shvDeclare()
  
  'kill pixel shader
  objD3DDev.DeletePixelShader shpHandle
  Set shpCode = Nothing
  Erase shpArray()
  
  'release textures
  objD3DDev.SetTexture 0, Nothing
  objD3DDev.SetTexture 1, Nothing
  objD3DDev.SetTexture 2, Nothing
  objD3DDev.SetTexture 3, Nothing
  Set texPalette = Nothing
  Set texHeight = Nothing
  Set texEnvironment = Nothing
  Set texNormal = Nothing
  
  'destroy all objects in a proper oreder
  Set objD3DHlp = Nothing
  Set objD3DDev = Nothing
  Set objD3D = Nothing
  Set objDX = Nothing
  
  'full stop
  End
  
End Sub

'helper vector creation function
Private Function vec4f(X As Single, Y As Single, z As Single, w As Single) As D3DVECTOR4
  With vec4f
    .X = X
    .Y = Y
    .z = z
    .w = w
  End With
End Function
