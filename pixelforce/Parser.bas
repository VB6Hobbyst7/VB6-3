Attribute VB_Name = "Parser"

'just to keep code clear
Option Explicit
Option Base 0

'syntax help
Private Type SyntaxEntry
  KeyWord As String
  Syntax As String
  Info As String
  MoreInfo As String
End Type
Public SyntaxArray() As SyntaxEntry
Public SyntaxOK As Boolean

'mesh rotation variables
Private Alpha As Single
Private SinAlpha As Single
Private CosAlpha As Single
Private Beta As Single
Private SinBeta As Single
Private CosBeta As Single
Private Gamma As Single
Private SinGamma As Single
Private CosGamma As Single
Private FX As Single
Private FY As Single
Private FZ As Single
Private FT1 As Integer
Private FT2 As Integer
Private FT3 As Integer

'mesh loader function
Private GeoFrag As String
Private GeoData As Boolean
Private Vertices As Integer
Private TVerts As Integer
Private Faces As Integer
Private Buffer() As Vertex
Private TBuffer() As Vertex
Private VX1 As Vertex
Private VX2 As Vertex
Private VX3 As Vertex
Private FA As Integer
Private FB As Integer
Private FC As Integer

'for cutline function
Private Separator As Integer

'parsing temp variables
Private Code As String
Private Fragment As String
Private Section As String
Private CurrentLine As Integer

'for value cutting functions
Private val_space As Boolean
Private val_index As Integer
Private val_buffer As String
Private val_code As Byte

'texture loader
Private TL As TgaFile
Private Bits() As ARGB32bit

'locked scale
Private LS As Single

'texture processing vars
Private X As Integer
Private Y As Integer
Private A As Integer

'buffer variables from code
Private pR As Single
Private pG As Single
Private pB As Single
Private pCl As String
Private pAA As Single
Private pZn As Single
Private pZf As Single
Private pSx As Single
Private pSy As Single
Private pSz As Single
Private pPx As Single
Private pPy As Single
Private pPz As Single
Private pRx As Single
Private pRy As Single
Private pRz As Single
Private pE As Single
Private pA As Integer
Private pM As Single
Private pT As Integer
Private pL As String
Private pF As String
Private pC As String
Private pN As String
Private pO As String

'syntax helper vars
Private sKW As String
Private sSX As String
Private sIO As String
Private sMI As String

'load external syntax help
Public Sub LoadCodeSyntax()
  'error handler
  On Error Resume Next
  Open fixPath & "Config\Code.txt" For Input As #1
  'failed?
  If Not Err.Number = 0 Then
    Err.Clear
    Close #1
    SyntaxOK = False
    MsgBox "Could not load external code syntax file. Internal help will be used.", vbCritical + vbOKOnly, "Error"
  Else
    'load xml-style data
    Section = vbNullString
    CurrentLine = 0
    Do While Not EOF(1)
      'get line
      Line Input #1, Fragment
      If Section = vbNullString Then
        If InStr(1, UCase(Fragment), "<ENTRY>", vbTextCompare) > 0 Then Section = "ENTRY"
        'reset vars
        sKW = vbNullString
        sSX = vbNullString
        sIO = vbNullString
        sMI = vbNullString
      Else
        'keyword
        If InStr(1, UCase(Fragment), "KEYWORD", vbTextCompare) > 0 Then sKW = Right(Fragment, Len(Fragment) - InStr(1, UCase(Fragment), "KEYWORD", vbTextCompare) - 7)
        'syntax
        If InStr(1, UCase(Fragment), "SYNTAX", vbTextCompare) > 0 Then sSX = Right(Fragment, Len(Fragment) - InStr(1, UCase(Fragment), "SYNTAX", vbTextCompare) - 6)
        'info
        If InStr(1, UCase(Fragment), "INFO", vbTextCompare) > 0 Then sIO = Right(Fragment, Len(Fragment) - InStr(1, UCase(Fragment), "INFO", vbTextCompare) - 4)
        'moreinfo
        If InStr(1, UCase(Fragment), "MORE", vbTextCompare) > 0 Then sMI = Right(Fragment, Len(Fragment) - InStr(1, UCase(Fragment), "MORE", vbTextCompare) - 4)
      End If
      'close section?
      If InStr(1, Fragment, "</", vbTextCompare) > 0 And InStr(1, Fragment, ">", vbTextCompare) > 0 And InStr(1, UCase(Fragment), Section, vbTextCompare) > 0 Then
        'add entry into array
        CurrentLine = CurrentLine + 1
        ReDim Preserve SyntaxArray(CurrentLine) As SyntaxEntry
        With SyntaxArray(CurrentLine)
          .Info = sIO
          .KeyWord = sKW
          .MoreInfo = sMI
          .Syntax = sSX
        End With
      End If
    Loop
    'all done
    Close #1
    SyntaxOK = True
  End If
End Sub

'put data in buffer, and clean it up (remove garbage)
Public Sub val_init(data As String)
  'reset
  val_buffer = vbNullString
  val_space = False
  'scan line
  For val_index = 1 To Len(data) Step 1
    'get symbol
    val_code = Asc(Mid(data, val_index, 1))
    'check for tab or space
    If val_code = 32 Or val_code = 9 Then
      'prevent multi-spaces or multi-tabs
      If Not val_space Then
        val_buffer = val_buffer & Chr(32)
      End If
      'space found
      val_space = True
    Else
      'it is not space
      val_space = False
      'copy symbol
      val_buffer = val_buffer & Chr(val_code)
    End If
  Next val_index
  'remove double spaces in the begining and in the end of line
  If Left(val_buffer, 1) = Chr(32) Then val_buffer = Right(val_buffer, Len(val_buffer) - 1)
  If Not Right(val_buffer, 1) = Chr(32) Then val_buffer = val_buffer & Chr(32)
End Sub

'return string from buffer & update it
Public Function str_cutout() As String
  'fing space
  val_index = InStr(1, val_buffer, Chr(32), vbTextCompare)
  If val_index > 0 Then
    'cut out string
    str_cutout = Left(val_buffer, val_index - 1)
    val_buffer = Right(val_buffer, Len(val_buffer) - val_index)
  Else
    'nothing left
    val_buffer = vbNullString
    str_cutout = vbNullString
  End If
End Function

'return single value from buffer & update it
Public Function val_cutout() As Single
  'fing space
  val_index = InStr(1, val_buffer, Chr(32), vbTextCompare)
  If val_index > 0 Then
    'cut out value
    val_cutout = Val(Left(val_buffer, val_index - 1))
    val_buffer = Right(val_buffer, Len(val_buffer) - val_index)
  Else
    'nothing left
    val_buffer = vbNullString
    val_cutout = 0
  End If
End Function

'cuts a line from text
Private Function CutLine(CutData As String, Optional CutSymbol As String = vbCrLf) As String
  'find symbol
  Separator = InStr(1, CutData, CutSymbol, vbBinaryCompare)
  'got symbol?
  If Separator > 0 Then
    'return line
    CutLine = Left(CutData, Separator - 1)
    'remove line from text
    CutData = Right(CutData, Len(CutData) - (Separator + (Len(CutSymbol) - 1)))
  Else
    'return entrie text
    CutLine = CutData
    CutData = vbNullString
  End If
End Function

'parser procedure
Public Sub ParseScene()
  'error handler
  On Error GoTo Failed
  'lock image scale with resolution
  LS = BufX
  If BufY < LS Then LS = BufY
  LS = LS * 0.5
  'get code
  Code = Workspace.SceneCode.Text
  'null section
  Section = vbNullString
  CurrentLine = 0
  'start parsing it
  Do While Len(Code) > 0
    Workspace.infoState.Caption = "Parsing..."
    DoEvents
    'get line
    Fragment = UCase(CutLine(Code, vbCrLf))
    CurrentLine = CurrentLine + 1
    'remove any comments from line :)
    If InStr(1, Fragment, "//", vbTextCompare) > 0 Then Fragment = Left(Fragment, InStr(1, Fragment, "//", vbTextCompare) - 1)
    'open section?
    If Section = vbNullString Then
      If InStr(1, Fragment, "<BACKBUFFER>", vbTextCompare) > 0 Then Section = "BACKBUFFER"
      If InStr(1, Fragment, "<CAMERA>", vbTextCompare) > 0 Then Section = "CAMERA"
      If InStr(1, Fragment, "<LIGHT>", vbTextCompare) > 0 Then Section = "LIGHT"
      If InStr(1, Fragment, "<DIFFUSEMAP>", vbTextCompare) > 0 Then Section = "DIFFUSEMAP"
      If InStr(1, Fragment, "<MESH>", vbTextCompare) > 0 Then Section = "MESH"
      If InStr(1, Fragment, "<CLIPPINGDISTANCE>", vbTextCompare) > 0 Then Section = "CLIPPINGDISTANCE"
      'reset all vars
      pR = 0
      pG = 0
      pB = 0
      pCl = "ON"
      pAA = 0
      pZn = -1
      pZf = 1
      pSx = 1
      pSy = 1
      pSz = 1
      pPx = 0
      pPy = 0
      pPz = 0
      pRx = 0
      pRy = 0
      pRz = 0
      pE = 1
      pA = 0
      pM = 1
      pT = 0
      pL = "ON"
      pF = vbNullString
      pC = "ON"
      pN = "ON"
      pO = "ON"
    Else 'process variables for sections
      'color
      If InStr(1, Fragment, "COLOR", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "COLOR", vbTextCompare) - 5)
        pR = val_cutout
        pG = val_cutout
        pB = val_cutout
      End If
      'clear
      If InStr(1, Fragment, "CLEAR", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "CLEAR", vbTextCompare) - 5)
        pCl = str_cutout
      End If
      'aliasedgeonly
      If InStr(1, Fragment, "ALIASEDGEONLY", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "ALIASEDGEONLY", vbTextCompare) - 13)
        pO = str_cutout
      End If
      'antialiaslevel
      If InStr(1, Fragment, "ANTIALIASLEVEL", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "ANTIALIASLEVEL", vbTextCompare) - 14)
        pAA = val_cutout
      End If
      'znear
      If InStr(1, Fragment, "ZNEAR", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "ZNEAR", vbTextCompare) - 5)
        pZn = val_cutout
      End If
      'zfar
      If InStr(1, Fragment, "ZFAR", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "ZFAR", vbTextCompare) - 4)
        pZf = val_cutout
      End If
      'scale
      If InStr(1, Fragment, "SCALE", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "SCALE", vbTextCompare) - 5)
        pSx = val_cutout
        pSy = val_cutout
        pSz = val_cutout
      End If
      'position
      If InStr(1, Fragment, "POSITION", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "POSITION", vbTextCompare) - 8)
        pPx = val_cutout
        pPy = val_cutout
        pPz = val_cutout
      End If
      'rotation
      If InStr(1, Fragment, "ROTATION", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "ROTATION", vbTextCompare) - 8)
        pRx = val_cutout
        pRy = val_cutout
        pRz = val_cutout
      End If
      'range
      If InStr(1, Fragment, "RANGE", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "RANGE", vbTextCompare) - 5)
        pE = val_cutout
      End If
      'alpha
      If InStr(1, Fragment, "ALPHA", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "ALPHA", vbTextCompare) - 5)
        pA = val_cutout
      End If
      'amplify
      If InStr(1, Fragment, "AMPLIFY", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "AMPLIFY", vbTextCompare) - 7)
        pM = val_cutout
      End If
      'texture
      If InStr(1, Fragment, "TEXTURE", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "TEXTURE", vbTextCompare) - 7)
        pT = val_cutout
      End If
      'lighting
      If InStr(1, Fragment, "LIGHTING", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "LIGHTING", vbTextCompare) - 8)
        pL = str_cutout
      End If
      'file
      If InStr(1, Fragment, "FILE", vbTextCompare) > 0 Then
        pF = Right(Fragment, Len(Fragment) - InStr(1, Fragment, "FILE", vbTextCompare) - 4)
        'replace "$LOCALPATH\" to current path
        pF = Replace(pF, "$LOCALPATH\", UCase(fixPath), 1, -1, vbTextCompare)
      End If
      'transparency
      If InStr(1, Fragment, "TRANSPARENCY", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "TRANSPARENCY", vbTextCompare) - 12)
        pC = str_cutout
      End If
      'generate32bit
      If InStr(1, Fragment, "GENERATE32BIT", vbTextCompare) > 0 Then
        val_init Right(Fragment, Len(Fragment) - InStr(1, Fragment, "GENERATE32BIT", vbTextCompare) - 13)
        pN = str_cutout
      End If
    End If
    'close section?
    If InStr(1, Fragment, "</", vbTextCompare) > 0 And InStr(1, Fragment, ">", vbTextCompare) > 0 And InStr(1, UCase(Fragment), Section, vbTextCompare) > 0 Then
      'choose section & apply parameters
      Select Case Section
        Case "BACKBUFFER"
          BackBufferColor.R = pR
          BackBufferColor.G = pG
          BackBufferColor.B = pB
          If pCl = "ON" Then
            ClearBuffer = True
          Else
            ClearBuffer = False
          End If
          EdgeAliasLevel = pAA
          If pO = "ON" Then
            EdgeOnly = True
          Else
            EdgeOnly = False
          End If
        Case "CAMERA"
          ShiftX = pPx
          ShiftY = pPy
          ShiftZ = pPz
          SceneScale = pSx * LS
          CameraRotation pRx, pRy
        Case "DIFFUSEMAP"
          Set TL = New TgaFile
          Workspace.infoState.Caption = "Loading Texture: " & pF
          DoEvents
          If Not TL.LoadTga(pF) Then
            MsgBox "Could not load texture: " & pF & vbCrLf & "Line: " & CurrentLine, vbExclamation + vbOKOnly, "Warning"
          Else
            Workspace.infoState.Caption = "Adding Texture: " & pF & vbCrLf & TL.Width & "x" & TL.Height & " pixels"
            DoEvents
            ReDim Bits(0, 0) As ARGB32bit
            TL.GetBits Bits()
            If TL.AlphaBits = 8 Then
              If pC = "OFF" Then
                Workspace.infoState.Caption = "Removing Alpha Channel: " & pF & vbCrLf & TL.Width & "x" & TL.Height & " pixels"
                DoEvents
                For Y = 0 To TL.Height Step 1
                  For X = 0 To TL.Width Step 1
                    Bits(X, Y).A = 0
                  Next X
                Next Y
              End If
            Else
              If pN = "ON" Then
                Workspace.infoState.Caption = "Generating Alpha Channel: " & pF & vbCrLf & TL.Width & "x" & TL.Height & " pixels"
                DoEvents
                For Y = 0 To TL.Height Step 1
                  For X = 0 To TL.Width Step 1
                    A = (CInt(Bits(X, Y).R) + CInt(Bits(X, Y).G) + CInt(Bits(X, Y).B)) / 3 + pA
                    If A < 0 Then A = 0
                    If A > 255 Then A = 255
                    Bits(X, Y).A = CByte(A)
                  Next X
                Next Y
              End If
            End If
            Workspace.infoState.Caption = "Adding Texture: " & pF & vbCrLf & TL.Width & "x" & TL.Height & " pixels"
            DoEvents
            AddTexture TL.Width, TL.Height, Bits()
            Erase Bits()
          End If
          TL.Destroy
          Set TL = Nothing
        Case "LIGHT"
          AddLight pPx, pPy, pPz, pE, pM, CreatePixel32Bit(pA + 0, pR, pG, pB)
        Case "MESH"
          Workspace.infoState.Caption = "Loading Mesh: " & pF
          DoEvents
          LoadMeshFile pF
        Case "CLIPPINGDISTANCE"
          ZNear = pZn
          ZFar = pZf
      End Select
      'no section
      Section = vbNullString
    End If
  Loop
  'error handler
  Exit Sub
Failed:
  Err.Clear
  MsgBox "Error at line: " & CurrentLine, vbExclamation + vbOKOnly, "Warning"
End Sub

'loads primitives into raytracer's buffer
Private Sub LoadMeshFile(file As String)
  'error handler
  On Error Resume Next
  Open file For Input As #1
  'failed?
  If Not Err.Number = 0 Then
    MsgBox "Unable to load mesh file: " & file, vbCritical + vbOKOnly, "Error"
    Err.Clear
    Close #1
  Else
    'set mesh rotation and prepare angles
    'convert to radians
    Alpha = pRx * Pi / 180
    Beta = pRy * Pi / 180
    Gamma = pRz * Pi / 180
    'prepare angles
    SinAlpha = Sin(Alpha)
    CosAlpha = Cos(Alpha)
    SinBeta = Sin(Beta)
    CosBeta = Cos(Beta)
    SinGamma = Sin(Gamma)
    CosGamma = Cos(Gamma)
    'search for geometry declaration
    GeoData = False
    Do While Not EOF(1)
      Line Input #1, GeoFrag
      If GeoFrag = "<geometry>" Then GeoData = True
      If GeoFrag = "</geometry>" Then GeoData = False
      If GeoData Then
        'get number of vertices and faces
        Input #1, Vertices, TVerts, Faces
        'allocate buffer for vertices
        ReDim Buffer(Vertices) As Vertex
        'load data from file into buffer
        For X = 1 To Vertices Step 1
          With Buffer(X)
            'load point position in 3d-space
            Input #1, .X
            Input #1, .Y
            Input #1, .Z
            'transform coords
            Buffer(X) = Transform(Buffer(X))
            'move & scale
            .X = .X * pSx + pPx
            .Y = .Y * pSy + pPy
            .Z = .Z * pSz + pPz
          End With
        Next X
        'allocate buffer for tvertices
        ReDim TBuffer(TVerts) As Vertex
        'load data from file into buffer
        For X = 1 To TVerts Step 1
          With TBuffer(X)
            'load texture coords
            Input #1, .U
            Input #1, .V
          End With
        Next X
        'now create primitives
        For X = 1 To Faces Step 1
          'get triangle points
          Input #1, FA
          Input #1, FB
          Input #1, FC
          'get triangle textured vertices
          Input #1, FT1
          Input #1, FT2
          Input #1, FT3
          'add triangle to scene
          VX1 = Buffer(FA)
          VX1.U = TBuffer(FT1).U
          VX1.V = TBuffer(FT1).V
          VX2 = Buffer(FB)
          VX2.U = TBuffer(FT2).U
          VX2.V = TBuffer(FT2).V
          VX3 = Buffer(FC)
          VX3.U = TBuffer(FT3).U
          VX3.V = TBuffer(FT3).V
          If pL = "ON" Then
            AddPrimitive VX1, VX2, VX3, pT, pA, True
          Else
            AddPrimitive VX1, VX2, VX3, pT, pA, False
          End If
        Next X
      End If
    Loop
    Close #1
  End If
End Sub

'point transformation
Private Function Transform(Point As Vertex) As Vertex
  Transform = Point
  With Transform
    'rotate point by x, y & z axis
    FZ = .Z * CosAlpha - .X * SinAlpha
    .X = .Z * SinAlpha + .X * CosAlpha
    .Z = FZ
    FY = .Y * CosBeta - .Z * SinBeta
    .Z = .Y * SinBeta + .Z * CosBeta
    .Y = FY
    FX = .X * CosGamma - .Y * SinGamma
    .Y = .X * SinGamma + .Y * CosGamma
    .X = FX
  End With
End Function

