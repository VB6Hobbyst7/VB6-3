Attribute VB_Name = "mImport"
'================================================
' Module:        mImport
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  clsBinaryFile.cls
' Last revision: 2006.06.23
'================================================

'------ 3D STUDIO MAX 5.0 ASCII SCENE EXPORT 2.0 IMPORTER ------

Option Explicit
Option Base 0

' Stores tree data and current level
Private sRootInfo() As String
Private iRootLevel As Integer

Private Type tMaterial
  map_file As String
  ambient_r As Single
  ambient_g As Single
  ambient_b As Single
End Type

Private sRootFrame As String
Private tMaterialList() As tMaterial
Private iCurrentMaterial As Integer
Private iCurrentObject As Integer
Private lngCurrentFace As Long
Private lngCurrentVertex As Long

' Used for status window
Private lngLastTimer As Long
Private lngFileTotal As Long
Private lngFileDone As Long

' Buffer and temporary variables
Private bValSpace As Boolean
Private iValIndex As Integer
Private sValBuffer As String
Private btValCode As Byte

' CleanUp data string
Private Sub ValInit(data As String)
  'Reset Variables
  sValBuffer = vbNullString
  bValSpace = False
  'Scan The String
  For iValIndex = 1 To Len(data) Step 1
    'Get Symbol
    btValCode = Asc(Mid(data, iValIndex, 1))
    'If Symbol Is TAB Or SPACE, Replace Tt To Space And Prevent Appearing 2 Or More Spaces At Once
    If btValCode = 32 Or btValCode = 9 Then
      If Not bValSpace Then
        sValBuffer = sValBuffer & Chr(32)
      End If
      bValSpace = True
    Else
      'Not A SPACE Or TAB, Add A Symbol To Buffer
      bValSpace = False
      sValBuffer = sValBuffer & Chr(btValCode)
    End If
  Next iValIndex
  'CheckUp For Cpaces In The Begining And In The End Of String
  If Left(sValBuffer, 1) = Chr(32) Then sValBuffer = Right(sValBuffer, Len(sValBuffer) - 1)
  If Not Right(sValBuffer, 1) = Chr(32) Then sValBuffer = sValBuffer & Chr(32)
End Sub

' Cut Buffer
Private Function ValCutout() As Single
  'Find Space
  iValIndex = InStr(1, sValBuffer, Chr(32), vbTextCompare)
  If iValIndex > 0 Then
    'Return The Value
    ValCutout = Val(Left(sValBuffer, iValIndex - 1))
    'Cut String
    sValBuffer = Right(sValBuffer, Len(sValBuffer) - iValIndex)
  Else
    'No Space? End Of String
    sValBuffer = vbNullString
    ValCutout = 0
  End If
End Function

Private Function GetVal(search As String) As Long
  GetVal = Val(Right(sValBuffer, Len(sValBuffer) - InStr(1, sValBuffer, search, vbTextCompare) - Len(search)))
End Function

Private Sub RootReset()
  ' Initialize tree buffer
  iRootLevel = 0
  ReDim sRootInfo(iRootLevel)
  sRootInfo(iRootLevel) = "root"
End Sub

Private Sub RootIncrease(description As String)
  ' Increase buffer size
  iRootLevel = iRootLevel + 1
  ReDim Preserve sRootInfo(iRootLevel)
  sRootInfo(iRootLevel) = description
End Sub

Private Sub RootDecrease()
  ' Decrease buffer size
  iRootLevel = iRootLevel - 1
  ReDim Preserve sRootInfo(iRootLevel)
End Sub

Private Function RootStatus() As String
  ' Return current tree level info
  RootStatus = sRootInfo(iRootLevel)
End Function

Private Sub RootShutdown()
  ' Cleanup
  Erase sRootInfo()
End Sub

Public Sub ImportASE(sFile As String)
  'Show Status Window
  Load frmStatus
  frmStatus.FileName.Caption = sFile
  'Reset All Variables And Get Ready To Import Data
  lngLastTimer = Int(Timer * 5)
  lngFileDone = 0
  'CleanUp Buffer
  ReDim tObjectList(0)
  'Reset Counters
  iCurrentObject = -1
  iCurrentMaterial = -1
  'Create Tree System
  RootReset
  'open virtual sFile
  If Not cVirtualFile.vfOpen(sFile, cPKG) Then
    cVirtualFile.vfClose
    MsgBox "sFile can not be found in package. Unable to create virtual sFile." & vbCrLf & cPKG.pkNameHandle & " > " & sFile, vbCritical + vbOKOnly, "Error"
    ReDim tObjectList(0)
    ReDim tObjectList(0).vertex(0)
    ReDim tObjectList(0).face(0)
    Unload frmStatus
    Exit Sub
    'Shutdown
  End If
  lngFileTotal = cVirtualFile.lngLength + 1
  Do While Not cVirtualFile.vfEof
    'read text line form virtual sFile
    sRootFrame = cVirtualFile.vfLine
    lngFileDone = lngFileDone + Len(sRootFrame) + 2
    'Update By Timer
    If Not lngLastTimer = Int(Timer * 15) Then
      'Update Status Information And Progress Bar
      With frmStatus
        .Caption = "importing..."
        .percent.Caption = Int(100 / lngFileTotal * lngFileDone) & "%"
        .foreground.Width = Int((.background.Width - 2) / lngFileTotal * lngFileDone)
      End With
      DoEvents
      lngLastTimer = Int(Timer * 15)
    End If
    
    Select Case RootStatus
      
      Case "root"
        'Found A Material List Declaration? Increase Tree Level From "ROOT" To "MATERIAL_LIST"
        If InStr(1, sRootFrame, "*MATERIAL_LIST") Then RootIncrease "material_list"
        'If Mesh Declaration Found
        If InStr(1, sRootFrame, "*GEOMOBJECT") Then
          'Increase Tree Level
          RootIncrease "geometry_object"
          'Increase Object Counter
          iCurrentObject = iCurrentObject + 1
          'Create A Buffer
          ReDim Preserve tObjectList(iCurrentObject)
          'Apply Texture And Other Parameters To The Object, If Exists
          If UBound(tMaterialList()) >= iCurrentObject Then
            Dim mat_id As Integer
            mat_id = iCurrentObject
            'Texture Map
            tObjectList(iCurrentObject).map = tMaterialList(mat_id).map_file
            'Emissive Material Info
            tObjectList(iCurrentObject).emissive.R = tMaterialList(mat_id).ambient_r
            tObjectList(iCurrentObject).emissive.G = tMaterialList(mat_id).ambient_g
            tObjectList(iCurrentObject).emissive.B = tMaterialList(mat_id).ambient_b
          End If
        End If
      
      Case "material_list"
        If InStr(1, sRootFrame, "*MATERIAL") And Right(sRootFrame, 1) = "{" Then RootIncrease "material_data"
        'Find Out How Many Materials Are In The sFile
        If InStr(1, sRootFrame, "*MATERIAL_COUNT") Then
          ReDim tMaterialList(Val(Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, Chr(32)))))
        End If
      
      'Parsing Material Data
      Case "material_data"
        'Declaration Of Ambient Material Info
        If InStr(1, sRootFrame, "*MATERIAL_AMBIENT") Then
          iCurrentMaterial = iCurrentMaterial + 1
          'Put Data String Into Buffer
          ValInit Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, "*MATERIAL_AMBIENT") - 17)
          'Use Current Material
          With tMaterialList(iCurrentMaterial)
            'Parse Material RGB Info
            .ambient_r = ValCutout
            .ambient_g = ValCutout
            .ambient_b = ValCutout
          End With
        End If
        'Find Texture Diffuse Map And Ambient Material Info
        If InStr(1, sRootFrame, "*MAP_DIFFUSE") Then
          RootIncrease "map_diffuse"
        End If
      
      'Get Diffuse Texture FileName, If Declaration Found
      Case "map_diffuse"
        If InStr(1, sRootFrame, "*BITMAP") And Right(sRootFrame, 1) = Chr(34) Then
          tMaterialList(iCurrentMaterial).map_file = Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, Chr(34)))
          tMaterialList(iCurrentMaterial).map_file = Left(tMaterialList(iCurrentMaterial).map_file, Len(tMaterialList(iCurrentMaterial).map_file) - 1)
        End If
    
      'Get Info About Geometry Object
      Case "geometry_object"
        'MATERIAL INDEX!!!
        If InStr(1, sRootFrame, "*MATERIAL_REF") Then
          tObjectList(iCurrentObject).material_reference = Val(Right(sRootFrame, Len(sRootFrame) - (InStr(1, sRootFrame, "*MATERIAL_REF") + 13)))
        End If
        'Get Mesh Name
        If InStr(1, sRootFrame, "*NODE_NAME") Then
          tObjectList(iCurrentObject).id = Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, Chr(34)))
          tObjectList(iCurrentObject).id = Left(tObjectList(iCurrentObject).id, Len(tObjectList(iCurrentObject).id) - 1)
        End If
        'Mesh Info Declaration, Increase Tree Level
        If InStr(1, sRootFrame, "*NODE_TM") Then RootIncrease "geometry_info"
        'Mesh Data Declaration, Increase Tree Level
        If InStr(1, sRootFrame, "*MESH") Then RootIncrease "mesh_data"
        
        
      'Now We Are Going To Read Information About Our Mesh
      Case "geometry_info"
        'Found Declaration, About Where Our Object Locates
        If InStr(1, sRootFrame, "*TM_POS") Then
          'Put Data String Into Buffer
          ValInit Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, "*TM_POS") - 7)
          'Parse Buffer And Get Position
          With tObjectList(iCurrentObject).Position
            .X = ValCutout
            .Y = ValCutout
            .Z = ValCutout
          End With
        End If
      
      'We Are In Mesh Section
      Case "mesh_data"
        'If Found A Number Of Faces Or Vertices, Create A Buffer For Them
        If InStr(1, sRootFrame, "*MESH_NUMFACES") Then ReDim tObjectList(iCurrentObject).face(Val(Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, Chr(32)))))
        If InStr(1, sRootFrame, "*MESH_NUMVERTEX") Then ReDim tObjectList(iCurrentObject).vertex(Val(Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, Chr(32)))))
        'Vertex List Declaration
        If InStr(1, sRootFrame, "*MESH_VERTEX_LIST") Then
          'Increase Tree Level
          RootIncrease "vertex_list"
          'Reset Counter
          lngCurrentVertex = 0
        End If
        'skip tfaces
        If InStr(1, sRootFrame, "*MESH_TFACELIST") Then RootIncrease "tface_list"
        'Face List Declaration
        If InStr(1, sRootFrame, "*MESH_FACE_LIST") Then
          'Increase Tree Level
          RootIncrease "face_list"
          'Reset Counter
          lngCurrentFace = 0
        End If
        'TVertex (Textured Vertex) List Declaration
        If InStr(1, sRootFrame, "*MESH_TVERTLIST") Then
          'Increase Tree Level
          RootIncrease "tvert_list"
          'Reset Counter
          lngCurrentVertex = 0
        End If
        'If Found Normals Declaration, Increase Tree Level
        If InStr(1, sRootFrame, "*MESH_NORMALS") Then RootIncrease "normals"
    
      'When Listing Vertices
      Case "vertex_list"
        'Find Vertex Declaration
        If InStr(1, sRootFrame, "*MESH_VERTEX") Then
          'Increase Current Vertex Number
          lngCurrentVertex = lngCurrentVertex + 1
          'Use Current Vertex
          With tObjectList(iCurrentObject).vertex(lngCurrentVertex).Position
            'Put Data String Into Buffer
            ValInit Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, "*MESH_VERTEX") - 12)
            'Remove Number (We Don't Need It)
            ValCutout
            'Parse Vertex Data
            .X = ValCutout
            .Z = ValCutout
            .Y = ValCutout
          End With
        End If
      
      'When Listing Faces
      Case "face_list"
        'Find Face Declaration
        If InStr(1, sRootFrame, "*MESH_FACE") Then
          'Increase Current Face Number
          lngCurrentFace = lngCurrentFace + 1
          'Use Current Face
          With tObjectList(iCurrentObject).face(lngCurrentFace)
            'Put Data String Into Buffer
            ValInit Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, "*MESH_FACE") - 10)
            'Parse Face Data And Add Some Corrections
            .A = GetVal("A:") + 1
            .B = GetVal("B:") + 1
            .C = GetVal("C:") + 1
          End With
        End If
      
      'If We Are Currently At TVertex Section
      Case "tvert_list"
        'Find TVertex Declaration
        If InStr(1, sRootFrame, "*MESH_TVERT") Then
          'Increase Current Vertex Number
          lngCurrentVertex = lngCurrentVertex + 1
          'Our Object Has Texture Coordinates
          tObjectList(iCurrentObject).texture_present = True
          'Put Data String Into Buffer
          ValInit Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, "*MESH_TVERT") - 11)
          'Remove Number (We Don't Need It, But Sometimes It Is Very Important)
          ValCutout
          'Sometimes There Are Much More TVerts Than Verts. So, We Can Be Out Of Range, Fix It!
          If lngCurrentVertex <= UBound(tObjectList(iCurrentObject).vertex()) Then
            'Use Current TVertex
            With tObjectList(iCurrentObject).vertex(lngCurrentVertex).texture
              'Parse Vertex Data
              .U = ValCutout
              .V = ValCutout
            End With
          End If
        End If
      
      'If We Are Currently Listing Normals
      Case "normals"
        'Find Normal Declaration
        If InStr(1, sRootFrame, "*MESH_VERTEXNORMAL") Then
          'Our Object Has Normals
          tObjectList(iCurrentObject).normals_present = True
          'Put Data String Into Buffer
          ValInit Right(sRootFrame, Len(sRootFrame) - InStr(1, sRootFrame, "*MESH_VERTEXNORMAL") - 18)
          'Apply Normals To A Vertex
          With tObjectList(iCurrentObject).vertex(ValCutout + 1).normal
            'Parse String With Normals
            .X = ValCutout
            .Z = ValCutout
            .Y = ValCutout
          End With
        End If
    
    End Select
    'If We Got } Bracket, We Should Decrease Tree Level
    If InStr(1, sRootFrame, "}") And Not RootStatus = "root" Then RootDecrease
  Loop
  'Close virtual sFile
  cVirtualFile.vfClose
  'We Don't Need Material List Anymore
  Erase tMaterialList()
  'CleanUp Tree System
  RootShutdown
  'Hide Status Window
  Unload frmStatus
End Sub

