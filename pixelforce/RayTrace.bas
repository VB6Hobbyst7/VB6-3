Attribute VB_Name = "RayTrace"

'just to keep code clear
Option Base 0
Option Explicit

'fast "pset"
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'system timer
Private Declare Function GetTickCount Lib "kernel32" () As Long

'vertex point (float) format
Public Type Vertex
  x As Single 'x coord
  y As Single 'y coord
  Z As Single 'z coord
  U As Single 'texture x coord (from 0 to 1)
  V As Single 'texture y coord (from 0 to 1)
End Type

'128-bit color (for blending)
Private Type ARGB128Bit
  A As Long 'alpha 32-bit
  R As Long 'red 32-bit
  G As Long 'green 32-bit
  B As Long 'blue 32-bit
End Type

'triangle structure
Private Type Primitive
  Src1 As Vertex         'original vertex a
  Src2 As Vertex         'original vertex b
  Src3 As Vertex         'original vertex c
  Dst1 As Vertex         'transformed vertex a
  Dst2 As Vertex         'transformed vertex b
  Dst3 As Vertex         'transformed vertex c
  Id As Integer          'texture index
  MinX As Single         'min x bound rect
  MinY As Single         'min y bound rect
  MaxX As Single         'max x bound rect
  MaxY As Single         'min y bound rect
  Delta As Single        'temp variable for interpolation
  Alpha As Integer       'transparency value
  LightAffect As Boolean 'light affected
  CommonDepth As Single  'base triangle depth (for sorting)
End Type

'texture diffuse-map structure
Private Type Texture
  Width As Integer    'x size
  Height As Integer   'y size
  Bits() As ARGB32bit 'bitmap data in argb 32-bit format
End Type

'light source structure
Private Type Light
  Src As Vertex      'original position
  Dst As Vertex      'transformed position
  Color As ARGB32bit 'main color
  Range As Single    'max light range
  Amplify As Single  'core amplify
End Type

'z-buffer structure
Private Type DepthMap
  Z As Single  'pixel depth
  i As Integer 'primitive id
End Type

'stop flag
Public RenderStop As Boolean

'data buffers
Private TriBuff() As Primitive
Private TexBuff() As Texture
Private LitBuff() As Light

'data counters
Private TriCount As Integer
Private TexCount As Integer
Private LitCount As Integer

'output image
Public Out() As ARGB32bit
'depth (z) buffer
Public Depth() As DepthMap

'processed point
Private p As Vertex

'camera rotation variables
Private Alpha As Single
Private SinAlpha As Single
Private CosAlpha As Single
Private Beta As Single
Private SinBeta As Single
Private CosBeta As Single

'temp variables for interpolation
Private U As Single
Private V As Single

'temp variables for buffer rendering
Private x As Integer
Private y As Integer

'start time
Private RenderStartTime As Single

'temp variables for different helper functions
Private FX As Single
Private FY As Single
Private pX As Integer
Private pY As Integer
Private pZ As Single

'primitive counter
Private i As Integer
'light counter
Private J As Integer

'backbuffer dimension
Public BufX As Integer
Public BufY As Integer

'blur color
Private L As ARGB128Bit

'pixel counters
Private LastTimer As Long
Private ThisPixel As Long
Private PPS As Long

'triangle coefficents
Private Q1 As Single
Private Q2 As Single
Private Q3 As Single

'temp a,r,g,b components for color blending functions
Private CA As Single
Private CR As Single
Private CG As Single
Private CB As Single

'depth test result
Private DepthOK As Boolean

'aa vars
Private st1 As Single
Private st2 As Single
Private st3 As Single
Private st As Single
Private ps As Single

'pi constant
Public Const Pi As Single = 3.14159265358979

'variables for shell sort
Private ShellTable() As Integer
Private TempElement As Primitive
Private ShellFragment As Integer
Private ShellPass As Integer
Private FragmentPass As Integer
Private ElementPass As Integer

'photon mapping
Private Const nrPhotons As Integer = 1000     ' number of photons emitted
Private Const nrBounces As Integer = 3        ' number of times each photon bounces
Private Const LightPhotons As Boolean = True  ' enable photon lighting?
Private Const sqradius As Single = 0.7        ' photon integration area (squared for efficiency)
Private Const Exposure As Single = 50#        ' number of photons integrated at brightest pixel
Private numPhotons(2, 4) As Integer           ' photon count for each scene object
Private Photons(2, 5, 5000, 3, 3) As Single   ' allocated memory for per-object photon info

'antialiasing settings
Public EdgeAliasLevel As Integer
Public EdgeOnly As Boolean

'clipping range
Public ZNear As Single
Public ZFar As Single

'scene shift
Public ShiftX As Single
Public ShiftY As Single
Public ShiftZ As Single

'scale factor
Public SceneScale As Single

'back color
Public BackBufferColor As ARGB32bit

'features configuration flags
Public Antialiasing As Boolean
Public DepthSorting As Boolean
Public DepthTest As Boolean
Public Lighting As Boolean
Public TextureFiltering As Boolean
Public Texturing As Boolean
Public AlphaBlending As Boolean
Public ClearBuffer As Boolean

'remove all triangles
Public Sub ResetPrimitives()
  'memoru cleanup
  ReDim TriBuff(0) As Primitive
  TriCount = 0
End Sub

'add triangle to buffer
Public Sub AddPrimitive(Vertex1 As Vertex, Vertex2 As Vertex, Vertex3 As Vertex, TextureID As Integer, AlphaBlend As Integer, LightAffect As Boolean)
  If TextureID > TexCount Then TextureID = 0
  'one more element
  TriCount = TriCount + 1
  ReDim Preserve TriBuff(TriCount) As Primitive
  'set triangle data
  With TriBuff(TriCount)
    .Src1 = Vertex1
    .Src2 = Vertex2
    .Src3 = Vertex3
    .Id = TextureID
    .Alpha = AlphaBlend
    .LightAffect = LightAffect
  End With
End Sub

Public Sub ResetTextures()
  'memory cleanup
  ReDim TexBuff(0) As Texture
  TexCount = 0
  'create null texture
  With TexBuff(0)
    '2x2 pixels texture
    .Width = 1
    .Height = 1
    'allocate 4 pixels, 32bit for each for null texture
    ReDim .Bits(.Width, .Height) As ARGB32bit
  End With
End Sub

'add texture to buffer
Public Sub AddTexture(Width As Integer, Height As Integer, Bits() As ARGB32bit)
  Workspace.infoState.Caption = "Adding Texture: " & Width & "x" & Height & "..."
  'one more element
  TexCount = TexCount + 1
  ReDim Preserve TexBuff(TexCount) As Texture
  'set texture data
  With TexBuff(TexCount)
    'set dimension
    .Width = Width
    .Height = Height
    'allocate memory
    ReDim .Bits(Width, Height)
    'copy pixels
    .Bits() = Bits()
  End With
End Sub

'delete all lights
Public Sub ResetLights()
  'memory cleanup
  ReDim LitBuf(0) As Light
  LitCount = 0
End Sub

'create new light source
Public Sub AddLight(x As Single, y As Single, Z As Single, Range As Single, Amplify As Single, Color As ARGB32bit)
  'one more element
  LitCount = LitCount + 1
  ReDim Preserve LitBuff(LitCount) As Light
  'set light
  With LitBuff(LitCount)
    .Src.x = x
    .Src.y = y
    .Src.Z = Z
    .Amplify = Amplify
    .Range = Range
    .Color = Color
  End With
End Sub

'destroy renderer
Public Sub ReleaseRenderer()
  'cleanup memory
  Erase LitBuff()
  Erase TexBuff()
  Erase TriBuff()
  Erase Depth()
  Erase Out()
End Sub

'start renderer
Public Sub InitializeRenderer(Width As Integer, Height As Integer)
  'allocate memory for buffers
  ReDim Depth(Width, Height) As DepthMap
  ReDim Out(Width, Height) As ARGB32bit
  'remember buffer size
  BufX = Width
  BufY = Height
End Sub

'set scene rotation and prepare angles
Public Sub CameraRotation(x As Single, y As Single)
  'convert to radians
  Alpha = x * Pi / 180
  Beta = y * Pi / 180
  'prepare angles
  SinAlpha = Sin(Alpha)
  CosAlpha = Cos(Alpha)
  SinBeta = Sin(Beta)
  CosBeta = Cos(Beta)
End Sub

'create vertex type helper function
Public Function CreatePoint(x As Single, y As Single, Z As Single, U As Single, V As Single) As Vertex
  With CreatePoint
    .x = x
    .y = y
    .Z = Z
    .U = U
    .V = V
  End With
End Function

'createcolor type helper function
Public Function CreatePixel32Bit(A As Single, R As Single, G As Single, B As Single) As ARGB32bit
  With CreatePixel32Bit
    .A = A
    .R = R
    .G = G
    .B = B
  End With
End Function

'process scene
Public Sub RenderScene(Target As Object)
  DoEvents
  Output.infoTriangles.Caption = "Primitives: " & TriCount
  Output.infoLights.Caption = "Area Lights: " & LitCount
  Output.infoMaps.Caption = "Texture Maps: " & TexCount
  RenderStartTime = Timer
  Workspace.infoState.Caption = "Transforming lights..."
  DoEvents
  'process lights
  For i = 1 To LitCount Step 1
    With LitBuff(i)
      'transform position
      .Dst = Transform(.Src, 0, BufX / 2, BufY / 2)
    End With
  Next i
  Workspace.infoState.Caption = "Transforming triangles..."
  DoEvents
  'process triangles
  For i = 1 To TriCount Step 1
    With TriBuff(i)
      'transform vertices
      .Dst1 = Transform(.Src1, .Id, BufX / 2, BufY / 2)
      .Dst2 = Transform(.Src2, .Id, BufX / 2, BufY / 2)
      .Dst3 = Transform(.Src3, .Id, BufX / 2, BufY / 2)
      'calclate common depth value
      'If DepthSorting Then .CommonDepth = (.Dst1.Z + .Dst2.Z + .Dst3.Z) / 3 '(alternative)
      If DepthSorting Then .CommonDepth = Min3(.Dst1.Z, .Dst2.Z, .Dst3.Z) 'far point
      'find bounding rectcangle
      .MinX = Min3(.Dst1.x, .Dst2.x, .Dst3.x)
      .MaxX = Max3(.Dst1.x, .Dst2.x, .Dst3.x)
      .MinY = Min3(.Dst1.y, .Dst2.y, .Dst3.y)
      .MaxY = Max3(.Dst1.y, .Dst2.y, .Dst3.y)
      'calculate delta value for interpolation
      .Delta = ((.Dst2.x - .Dst1.x) * (.Dst3.y - .Dst1.y) - (.Dst2.y - .Dst1.y) * (.Dst3.x - .Dst1.x))
      'handle division by zero error
      If Not .Delta = 0 Then .Delta = 1 / .Delta
    End With
    DoEvents
    Output.infoTime.Caption = "Rendering Time: " & Int((Timer - RenderStartTime) * 100) / 100 & " s."
    If RenderStop Then Exit Sub
  Next i
  'sort triangles for rendering process speed up and alphablending
  If DepthSorting Then
    Workspace.infoState.Caption = "Proceeding with Z-Sort..."
    DoEvents
    'create shell table
    ReDim ShellTable(0 To 4) As Integer
    ShellTable(0) = 9
    ShellTable(1) = 5
    ShellTable(2) = 3
    ShellTable(3) = 2
    ShellTable(4) = 1
    'scan fragments
    For ShellPass = 0 To UBound(ShellTable()) Step 1
      ShellFragment = ShellTable(ShellPass)
      'do sorting for current fragment
      For FragmentPass = ShellFragment To UBound(TriBuff())
        DoEvents
        Output.infoTime.Caption = "Rendering Time: " & Int((Timer - RenderStartTime) * 100) / 100 & " s."
        If RenderStop Then Exit Sub
        TempElement = TriBuff(FragmentPass)
        For ElementPass = FragmentPass - ShellFragment To 1 Step -ShellFragment
          'check depth
          If TempElement.CommonDepth <= TriBuff(ElementPass).CommonDepth Then Exit For
          TriBuff(ElementPass + ShellFragment) = TriBuff(ElementPass)
        Next ElementPass
        TriBuff(ElementPass + ShellFragment) = TempElement
      Next FragmentPass
    Next ShellPass
    'memory clanup
    Erase ShellTable()
  End If
  'render depthbuffer & backbuffer
  Workspace.infoState.Caption = "Rendering..."
  DoEvents
  LastTimer = GetTickCount
  ThisPixel = 0
  For y = 1 To BufY Step 1
    For x = 1 To BufX Step 1
      'set initial depth
      If DepthTest Then
        Depth(x, y).Z = ZFar + 1
        Depth(x, y).i = 0
      End If
      If ClearBuffer Then Out(x, y) = BackBufferColor
      'process all triangles
      For i = 1 To TriCount Step 1
        With TriBuff(i)
          'current point in triangle bounding rectangle?
          If x >= .MinX And x <= .MaxX And y >= .MinY And y <= .MaxY Then
            'check if this point belongs current triangle
            Q1 = (.Dst2.x - x) * (.Dst3.y - y) - (.Dst2.y - y) * (.Dst3.x - x)
            Q2 = (.Dst1.x - x) * (.Dst2.y - y) - (.Dst1.y - y) * (.Dst2.x - x)
            Q3 = (.Dst3.x - x) * (.Dst1.y - y) - (.Dst3.y - y) * (.Dst1.x - x)
            'point in triangle?
            If (Q1 >= 0 And Q2 >= 0 And Q3 >= 0) Or (Q1 <= 0 And Q2 <= 0 And Q3 <= 0) Then
              'do depth test
              If DepthTest Then
                'interpolate depth value for current point
                p.Z = Interpolate(.Dst1.x, .Dst1.y, .Dst2.x, .Dst2.y, .Dst3.x, .Dst3.y, .Dst1.Z, .Dst2.Z, .Dst3.Z, .Delta, x, y)
                'apply clipping range
                If p.Z >= ZNear And p.Z <= ZFar Then
                  'compare
                  If p.Z < Depth(x, y).Z Then
                    'test passed, update zbuffer
                    Depth(x, y).Z = p.Z
                    Depth(x, y).i = i
                    DepthOK = True
                  Else
                    DepthOK = False
                  End If
                End If
              Else
                DepthOK = True
              End If
              'visible pixel?
              If DepthOK Then
                If Texturing Then
                  'interpolate texture coordinates
                  If .Id > 0 Then
                    p.U = Interpolate(.Dst1.x, .Dst1.y, .Dst2.x, .Dst2.y, .Dst3.x, .Dst3.y, .Dst1.U, .Dst2.U, .Dst3.U, .Delta, x, y)
                    p.V = Interpolate(.Dst1.x, .Dst1.y, .Dst2.x, .Dst2.y, .Dst3.x, .Dst3.y, .Dst1.V, .Dst2.V, .Dst3.V, .Delta, x, y)
                    'set textured pixel
                    If TextureFiltering Then
                      If AlphaBlending Then
                        'do alphablending
                        Out(x, y) = Pixel128To32Bit(BlendColor(Pixel32To128Bit(Out(x, y)), Filter(p.U, p.V, .Id), .Alpha))
                      Else
                        Out(x, y) = Pixel128To32Bit(Filter(p.U, p.V, .Id))
                      End If
                    Else
                      If AlphaBlending Then
                        'do alphablending (with no filtering)
                        Out(x, y) = Pixel128To32Bit(BlendColor(Pixel32To128Bit(Out(x, y)), Pixel32To128Bit(TexBuff(.Id).Bits(p.U, p.V)), .Alpha))
                      Else
                        Out(x, y) = TexBuff(.Id).Bits(p.U, p.V)
                      End If
                    End If
                  Else
                    If AlphaBlending Then
                      'do alphablending (with no texture)
                      Out(x, y) = Pixel128To32Bit(BlendColor(Pixel32To128Bit(Out(x, y)), Pixel32To128Bit(CreatePixel32Bit(0, 127, 127, 127)), .Alpha))
                    Else
                      'no texture
                      Out(x, y) = CreatePixel32Bit(127, 63, 63, 63)
                    End If
                  End If
                End If
                'apply lighting
                If Lighting And .LightAffect Then
                  p.x = x
                  p.y = y
                  For J = 1 To LitCount Step 1
                    'special area lights with blending
                    If AlphaBlending Then
                      Out(x, y) = Pixel128To32Bit(BlendColor(Pixel32To128Bit(Out(x, y)), Pixel32To128Bit(LightColor(p, J)), 0, True))
                    Else
                      Out(x, y) = Pixel128To32Bit(AddColor(Pixel32To128Bit(Out(x, y)), LightColor(p, J)))
                    End If
                  Next J
                End If
              End If
            End If
          End If
        End With
      Next i
      'draw fully-processed pixel on the screen space
      SetPixelV Target.hdc, x - 1, y - 1, rgb(Out(x, y).R, Out(x, y).G, Out(x, y).B)
      ThisPixel = ThisPixel + 1
      If LastTimer <= GetTickCount - 1000 Then
        PPS = ThisPixel
        ThisPixel = 0
        LastTimer = GetTickCount
      End If
    Next x
    'draw scanline & refresh image
    Output.infoTime.Caption = "Rendering Time: " & Int((Timer - RenderStartTime) * 100) / 100 & " s."
    Workspace.ProgressIndicator.PI_Percent 100 / BufY * y
    Target.Line (0, y + 1)-(Target.ScaleWidth, y + 1), vbRed
    If PPS > 1000 Then
      Workspace.infoState.Caption = "Rendering Line " & y & " Of " & BufY & " (Speed: " & Int(PPS / 1000) & "K PPS)"
    Else
      Workspace.infoState.Caption = "Rendering Line " & y & " Of " & BufY & " (Speed: " & PPS & " PPS)"
    End If
    DoEvents
    If RenderStop Then Exit Sub
  Next y
  'do edge-antialiasing
  If Antialiasing Then
    For y = 1 To BufY - 1 Step 1
      For x = 1 To BufX - 1 Step 1
        If Depth(x + 1, y).i <> Depth(x, y).i Or Depth(x, y + 1).i <> Depth(x, y).i Then
          If EdgeOnly Then
            If (Depth(x + 1, y).i = 0 Or Depth(x, y).i = 0) Or (Depth(x, y + 1).i = 0 Or Depth(x, y).i = 0) Then Blur x, y, EdgeAliasLevel
          Else
            Blur x, y, EdgeAliasLevel
          End If
        End If
      Next x
      'draw scanline & refresh image
      Workspace.infoState.Caption = "Performing Edge Anti-Alias Pass: " & y & " of " & BufY - 1 & "..."
      Output.infoTime.Caption = "Rendering Time: " & Int((Timer - RenderStartTime) * 100) / 100 & " s."
      Workspace.ProgressIndicator.PI_Percent 100 / BufY * y
      Output.Refresh
      DoEvents
      If RenderStop Then Exit Sub
    Next y
  End If
  'refresh ui
  With Workspace
    .ProgressIndicator.PI_Percent 100
    .infoState.Caption = "Render Completed In " & Int((Timer - RenderStartTime) * 100) / 100 & " s."
    Output.infoTime.Caption = "Rendering Time: " & Int((Timer - RenderStartTime) * 100) / 100 & " s."
    .Resolution.Enabled = True
    .ButtonRender.AB_Disabled = False
    .ButtonRender.AB_RenderIcon
    .ButtonResult.AB_Disabled = False
    .ButtonResult.AB_RenderIcon
    .ButtonCancel.AB_Disabled = True
    .ButtonCancel.AB_RenderIcon
  End With
End Sub

'blurring function
Private Sub Blur(DX As Integer, DY As Integer, Range As Integer)
  If Not (DX > Range And DY > Range And DX < BufX - Range And DY < BufY - Range) Then Exit Sub
  'pass all pixels in rectangle
  For FY = DY - Range To DY + Range Step 1
    For FX = DX - Range To DX + Range Step 1
      'get source pixel
      L = Pixel32To128Bit(Out(FX, FY))
      'add nearby pixles
      L = AddColor(L, Out(FX + 1, FY))
      L = AddColor(L, Out(FX, FY + 1))
      L = AddColor(L, Out(FX - 1, FY))
      L = AddColor(L, Out(FX, FY - 1))
      'div 5
      With L
        .R = .R / 5
        .G = .G / 5
        .B = .B / 5
      End With
      'set new pixel
      Out(FX, FY) = Pixel128To32Bit(L)
      SetPixelV Output.hdc, FX - 1, FY - 1, rgb(Out(FX, FY).R, Out(FX, FY).G, Out(FX, FY).B)
    Next FX
  Next FY
End Sub

'interpolation function
Private Function Interpolate(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single, V1 As Single, V2 As Single, V3 As Single, Delta As Single, DX As Integer, DY As Integer) As Single
  U = (((X2 - DX) * (Y3 - DY)) - ((Y2 - DY) * (X3 - DX))) * Delta
  V = (((X3 - DX) * (Y1 - DY)) - ((Y3 - DY) * (X1 - DX))) * Delta
  'retun interpolated value
  Interpolate = U * V1 + V * V2 + (1 - U - V) * V3
End Function

'point transformation
Private Function Transform(Point As Vertex, Id As Integer, MovX As Single, MovY As Single) As Vertex
  Transform = Point
  With Transform
    'scale & repose source coordinates
    .x = .x * SceneScale + ShiftX
    .y = .y * SceneScale + ShiftY
    .Z = .Z * SceneScale + ShiftZ
    'rotate point by x & y axis
    FX = .Z * CosAlpha - .x * SinAlpha
    .x = .Z * SinAlpha + .x * CosAlpha
    .Z = FX
    FY = .y * CosBeta - .Z * SinBeta
    .Z = .y * SinBeta + .Z * CosBeta
    'vertical flip
    .y = -FY
    'scale texture coordinates
    If Id > 0 Then
      .U = .U * (TexBuff(Id).Width - 1) + 1
      .V = .V * (TexBuff(Id).Height - 1) + 1
    End If
    'scroll coordinates
    .x = .x + MovX
    .y = .y + MovY
  End With
End Function

'do bilinear texture filtering
Private Function Filter(DU As Single, DV As Single, Id As Integer) As ARGB128Bit
  On Error Resume Next
  'prepare coords
  pX = Fix(DU)
  pY = Fix(DV)
  FX = (DU - pX)
  FY = (DV - pY)
  With TexBuff(Id)
    'filter r,g,b
    If AlphaBlending Then Filter.A = (FX * ((FY * CInt(.Bits(pX + 1, pY + 1).A)) + ((1 - FY) * CInt(.Bits(pX + 1, pY).A)))) + ((1 - FX) * ((FY * CInt(.Bits(pX, pY + 1).A)) + ((1 - FY) * CInt(.Bits(pX, pY).A))))
    Filter.R = (FX * ((FY * CInt(.Bits(pX + 1, pY + 1).R)) + ((1 - FY) * CInt(.Bits(pX + 1, pY).R)))) + ((1 - FX) * ((FY * CInt(.Bits(pX, pY + 1).R)) + ((1 - FY) * CInt(.Bits(pX, pY).R))))
    Filter.G = (FX * ((FY * CInt(.Bits(pX + 1, pY + 1).G)) + ((1 - FY) * CInt(.Bits(pX + 1, pY).G)))) + ((1 - FX) * ((FY * CInt(.Bits(pX, pY + 1).G)) + ((1 - FY) * CInt(.Bits(pX, pY).G))))
    Filter.B = (FX * ((FY * CInt(.Bits(pX + 1, pY + 1).B)) + ((1 - FY) * CInt(.Bits(pX + 1, pY).B)))) + ((1 - FX) * ((FY * CInt(.Bits(pX, pY + 1).B)) + ((1 - FY) * (.Bits(pX, pY).B))))
  End With
End Function

'calculate light color
Private Function LightColor(Point As Vertex, Id As Integer) As ARGB32bit
  With LitBuff(Id)
    'get distance between light source and current point
    FX = Sqr((.Dst.x - Point.x) ^ 2 + (.Dst.y - Point.y) ^ 2 + (.Dst.Z - Point.Z) ^ 2)
    'too far
    If FX >= .Range * Abs(SceneScale) Then
      'no lighting
      With LightColor
        .A = 0
        .R = 0
        .G = 0
        .B = 0
      End With
    Else
      'calculate mul factor
      FX = (Pi / 2) - ((Pi / 2) / (.Range * Abs(SceneScale)) * FX)
      'range check
      FX = Sin(FX) * .Amplify
      If FX < 0 Then FX = 0
      If FX > 1 Then FX = 1
      'mul color
      With .Color
        LightColor.A = .A * FX
        LightColor.R = .R * FX
        LightColor.G = .G * FX
        LightColor.B = .B * FX
      End With
    End If
  End With
End Function

'add color
Private Function AddColor(Pixel1 As ARGB128Bit, Pixel2 As ARGB32bit) As ARGB128Bit
  With AddColor
    'add a,r,g,b
    CA = CInt(Pixel1.A) + CInt(Pixel2.A)
    CR = CInt(Pixel1.R) + CInt(Pixel2.R)
    CG = CInt(Pixel1.G) + CInt(Pixel2.G)
    CB = CInt(Pixel1.B) + CInt(Pixel2.B)
    'return color
    .A = CA
    .R = CR
    .G = CG
    .B = CB
  End With
End Function

'add color
Private Function BlendColor(Pixel1 As ARGB128Bit, Pixel2 As ARGB128Bit, AddAlpha As Integer, Optional LightBlend As Boolean = False) As ARGB128Bit
  With BlendColor
    'add a,r,g,b
    If LightBlend Then
      CA = CInt(Pixel1.A) + CInt(Pixel2.A)
      If CA > 255 Then CA = 255
      If CA < 0 Then CA = 0
      CR = CInt(Pixel1.R) + (CInt(Pixel2.R)) / 255 * CA
      CG = CInt(Pixel1.G) + (CInt(Pixel2.G)) / 255 * CA
      CB = CInt(Pixel1.B) + (CInt(Pixel2.B)) / 255 * CA
    Else
      CA = CInt(Pixel2.A) + CInt(AddAlpha)
      If CA > 255 Then CA = 255
      If CA < 0 Then CA = 0
      CR = CInt(Pixel1.R) + (CInt(Pixel2.R) - CInt(Pixel1.R)) / 255 * CA
      CG = CInt(Pixel1.G) + (CInt(Pixel2.G) - CInt(Pixel1.G)) / 255 * CA
      CB = CInt(Pixel1.B) + (CInt(Pixel2.B) - CInt(Pixel1.B)) / 255 * CA
    End If
    'return color
    .A = CA
    .R = CR
    .G = CG
    .B = CB
  End With
End Function

'convert 32bit color to 128bit color
Private Function Pixel32To128Bit(Pixel As ARGB32bit) As ARGB128Bit
  'set a,r,g,b
  With Pixel32To128Bit
    .A = Pixel.A
    .R = Pixel.R
    .G = Pixel.G
    .B = Pixel.B
  End With
End Function

'convert 128bit color to 32bit color
Private Function Pixel128To32Bit(Pixel As ARGB128Bit) As ARGB32bit
  'get a,r,g,b
  With Pixel
    CA = .A
    CR = .R
    CG = .G
    CB = .B
  End With
  'range check
  If CA > 255 Then CA = 255
  If CR > 255 Then CR = 255
  If CG > 255 Then CG = 255
  If CB > 255 Then CB = 255
  If CA < 0 Then CA = 255
  If CR < 0 Then CR = 255
  If CG < 0 Then CG = 255
  If CB < 0 Then CB = 255
  'set a,r,g,b
  With Pixel128To32Bit
    .A = CA
    .R = CR
    .G = CG
    .B = CB
  End With
End Function

'return min value
Private Function Min3(A As Single, B As Single, c As Single) As Single
  Min3 = A
  If B < Min3 Then Min3 = B
  If c < Min3 Then Min3 = c
End Function

'return max value
Private Function Max3(A As Single, B As Single, c As Single) As Single
  Max3 = A
  If B > Max3 Then Max3 = B
  If c > Max3 Then Max3 = c
End Function

'---------------------------------------
' Photon Mapping
'---------------------------------------

Private Sub storePhoton(intType As Integer, intID As Integer, sngLocation() As Single, _
                        sngDirection() As Single, sngEnergy() As Single)
    
    Photons(intType, intID, numPhotons(intType, intID), 0) = sngLocation   ' location
    Photons(intType, intID, numPhotons(intType, intID), 1) = sngDirection  ' direction
    Photons(intType, intID, numPhotons(intType, intID), 2) = sngEnergy     ' attenuated energy (color)
    numPhotons(intType, intID) = numPhotons(intType, intID) + 1
End Sub

' shadow photons
Private Sub shadowPhoton(sngRay() As Single)
    Dim shadow(4) As Single
    Dim tPoint() As Single
    Dim tType, tIndex As Integer
    Dim bumpedPoint() As Single
    
End Sub

' photon visualization
Private Sub drawPhoton(rgb() As Single, p() As Single)
    
    Dim x, y As Integer
    
    If (view3D And p(2) > 0#) Then
        x = (szimg / 2) + CInt(szimg * p(0) / p(2))
        y = (szimg / 2) + CInt(szimg * -p(1) / p(2))
        If y <= szimg Then
            point(x,y)
            
End Sub
                        

