VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TgaFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'just to keep code clear
Option Explicit
Option Base 0

'pixel format
Public Type ARGB32bit
  B As Byte
  G As Byte
  R As Byte
  A As Byte
End Type

'image buffer
Private PixelMap() As ARGB32bit

'image dimensions
Public Width As Integer
Public Height As Integer

'pixel format
Public Format As Byte
Public AlphaBits As Byte
Public LockAlpha As Boolean

'compression info
Public Rle As Boolean
Public RleRatio As Single
Public RleBytes As Long

'comment
Public CommentLength As Byte
Public Comment As String
Public Ending As String

'tga file
Public FileName As String
Public FileSize As Long

'temp variables for load/save functions
Private Temp As Integer 'alpha accumulator
Private ColorMapType As Byte 'palette type
Private ImageType As Byte 'file format (rle compressed/not compressed)
Private ColorMapOffset As Integer 'palette offset
Private ColorMapLength As Integer 'palette size
Private ColorMapDepth As Byte 'palette pixel depth
Private XOffset As Integer 'image x position
Private YOffset As Integer 'image y position
Private pX As Integer 'pixel x coord
Private pY As Integer 'pixel y coord
Private CommentOffset As Long 'comment byte index
Private CommentByte As Byte 'comment's byte
Private RleBlockIndex As Long 'rle block index
Private RleBlockLength As Byte 'rle block size
Private RleBlockType As Byte 'rle block type (single/multiple)
Private RleBlockPixel As ARGB32bit 'rle single pixel declaration
Private RlePixelStream() As ARGB32bit 'unique pixel stream
Private RleByte As Byte 'temp byte

'get image data
Public Sub GetBits(Bits() As ARGB32bit)
  ReDim Bits(Width, Height)
  Bits() = PixelMap()
End Sub

'set image data
Public Sub SetBits(Bits() As ARGB32bit)
  ReDim PixelMap(Width, Height)
  PixelMap() = Bits()
End Sub

'return standard file ending
Public Function StandardEnding() As String
  StandardEnding = String(8, Chr(0)) & "TRUEVISION-XFILE." & Chr(0)
End Function

'pixel map memory usage
Public Function Memory(NullFix As Boolean) As Long
  If NullFix Then
    'used by image buffer
    Memory = Width + 1
    Memory = Memory * (Height + 1)
  Else
    'used by image pixels
    Memory = Width
    Memory = Memory * Height
  End If
  If AlphaBits = 8 Then
    'mul argb
    Memory = Memory * 4
  Else
    'mul rgb
    Memory = Memory * 3
  End If
End Function

'erase image buffer & reset variables
Public Sub Destroy()
  'memory cleanup
  Erase PixelMap()
  ReDim PixelMap(0, 0)
  'reset variables
  LockAlpha = False
  Width = 0
  Height = 0
  Format = 0
  Rle = False
  RleRatio = 1
  RleBytes = 0
  AlphaBits = 0
  FileName = vbNullString
  FileSize = 0
  CommentLength = 0
  Comment = vbNullString
  'standard tga file ending (not required, just to keep format)
  Ending = StandardEnding
End Sub

'create empty image buffer
Public Sub CreateEmptyPixelMap()
  Erase PixelMap()
  ReDim PixelMap(Width, Height)
End Sub

'put a,r,g,b pixel components
Public Sub SetPixel(X As Integer, Y As Integer, Pixel As ARGB32bit)
  PixelMap(X, Y) = Pixel
End Sub

'get a,r,g,b pixel components
Public Function GetPixel(X As Integer, Y As Integer) As ARGB32bit
  GetPixel = PixelMap(X, Y)
End Function

'save pixel map into tga file
Public Function SaveTga(Name As String) As Boolean
  'error handling
  On Error GoTo Failed
  SaveTga = False
  'file exists? kill it
  If Len(Dir(Name)) > 0 Then Kill Name
  'open it
  Open Name For Binary Access Write As #1
  'write tga header
  'comment length
  CommentLength = Len(Comment)
  Put #1, , CommentLength
  'color map (palette) type - not required
  ColorMapType = 0
  Put #1, , ColorMapType
  'image type - set uncompressed type (2) - compression is not finished yet :(
  ImageType = 2
  Put #1, , ImageType
  'palette offset - not required
  ColorMapOffset = 0
  Put #1, , ColorMapOffset
  'palette length - not required
  ColorMapLength = 0
  Put #1, , ColorMapLength
  'palette pixel depth - not required
  ColorMapDepth = 0
  Put #1, , ColorMapDepth
  'image position - not required
  XOffset = 0
  YOffset = 0
  Put #1, , XOffset
  Put #1, , YOffset
  'image dimensions
  Put #1, , Width
  Put #1, , Height
  'pixel format - supported only 24 and 32 bit formats
  Put #1, , Format
  'alpha channel bits - supported only 8 bit format
  Put #1, , AlphaBits
  'write comment, if exists
  For CommentOffset = 1 To CommentLength Step 1
    CommentByte = Asc(Mid(Comment, CommentOffset, 1))
    Put #1, , CommentByte
  Next CommentOffset
  'write image data
  If ImageType = 2 Then '(uncompressed data)
    'write pixels
    For pY = 1 To Height Step 1
      For pX = 1 To Width Step 1
        'vertical flip
        With PixelMap(pX, Height - (pY - 1))
          'set r,g,b
          Put #1, , .B
          Put #1, , .G
          Put #1, , .R
          'need alpha?
          If AlphaBits = 8 Then Put #1, , .A
        End With
      Next pX
      'refresh progress bar
      DoEvents
      Workspace.ProgressIndicator.PI_Percent 100 / Height * pY
    Next pY
  End If
  'write ending
  For CommentOffset = 1 To Len(Ending) Step 1
    CommentByte = Asc(Mid(Ending, CommentOffset, 1))
    Put #1, , CommentByte
  Next CommentOffset
  'close file
  Close #1
  'set new file
  FileSize = FileLen(Name)
  FileName = Name
  SaveTga = True
  'error handling
  Exit Function
Failed:
  'memory clenup
  Erase RlePixelStream()
  Close #1
  SaveTga = False
End Function

'load tga file into memory
Public Function LoadTga(Name As String) As Boolean
  'error handling
  On Error GoTo Failed
  'reset
  LoadTga = False
  Destroy
  'file exists?
  If Len(Dir(Name)) > 0 Then
    'get file size
    FileSize = FileLen(Name)
    'open it
    Open Name For Binary Access Read As #1
      'comment length
      Get #1, , CommentLength
      'color map (palette) type - not required
      Get #1, , ColorMapType
      'image type - supported only 2 and 10(rle compressed) true color (24/32 bit)
      Get #1, , ImageType
      If ImageType = 10 Or ImageType = 2 Then
        'set compression flag
        If ImageType = 10 Then
          Rle = True
        Else
          Rle = False
        End If
        'palette offset - not required
        Get #1, , ColorMapOffset
        'palette length - not required
        Get #1, , ColorMapLength
        'palette pixel depth - not required
        Get #1, , ColorMapDepth
        'image position - not required
        Get #1, , XOffset
        Get #1, , YOffset
        'image dimensions
        Get #1, , Width
        Get #1, , Height
        'allocate memory for image
        CreateEmptyPixelMap
        'get pixel format - supported only 24 and 32 bit formats
        Get #1, , Format
        If Format = 24 Or Format = 32 Then
          'get alpha channel bits - supported only 8 bit format
          Get #1, , AlphaBits
          If Not (Format = 32 And AlphaBits <> 8) Then
            'read comment, if exists
            For CommentOffset = 1 To CommentLength Step 1
              Get #1, , CommentByte
              Comment = Comment & Chr(CommentByte)
            Next CommentOffset
            'load image data
            If ImageType = 2 Then '(uncompressed)
              'read pixels
              For pY = 1 To Height Step 1
                For pX = 1 To Width Step 1
                  'vertical flip
                  With PixelMap(pX, Height - (pY - 1))
                    'get r,g,b
                    Get #1, , .B
                    Get #1, , .G
                    Get #1, , .R
                    'detect alpha
                    If AlphaBits = 8 Then
                      Get #1, , .A
                    Else
                      .A = 255
                    End If
                  End With
                Next pX
              Next pY
            Else '(rle compressed)
              'reset coords
              pX = 0
              pY = 1
              'decompression routine
              Do While pY <= Height
                'get block header
                Get #1, , RleBlockType
                RleBytes = RleBytes + 1
                'pixel stream?
                If RleBlockType > 128 Then
                  RleBlockLength = RleBlockType - 128
                  'get pixel declaration
                  With RleBlockPixel
                    Get #1, , .B
                    Get #1, , .G
                    Get #1, , .R
                    RleBytes = RleBytes + 3
                    'detect alpha
                    If AlphaBits = 8 Then
                      Get #1, , .A
                      RleBytes = RleBytes + 1
                    Else
                      .A = 255
                    End If
                  End With
                Else
                  'unique pixel stream?
                  RleBlockLength = RleBlockType
                End If
                'read pixels
                For RleBlockIndex = 0 To RleBlockLength Step 1
                  'get pixel declaration, if unique pixel stream found
                  If RleBlockType < 128 Then
                    With RleBlockPixel
                      Get #1, , .B
                      Get #1, , .G
                      Get #1, , .R
                      RleBytes = RleBytes + 3
                      'detect alpha
                      If AlphaBits = 8 Then
                        Get #1, , .A
                        RleBytes = RleBytes + 1
                      Else
                        .A = 255
                      End If
                    End With
                  End If
                  'next column
                  pX = pX + 1
                  'enf of line?
                  If pX > Width Then
                    'cr & lf :)
                    pX = 1 'reset column
                    pY = pY + 1 'next row
                    'y bound reached?
                    If pY > Height Then Exit Do
                  End If
                  'vertical flip
                  PixelMap(pX, Height - (pY - 1)) = RleBlockPixel
                Next RleBlockIndex
              Loop
              Seek #1, Seek(1) - 4
              'calculate compression ratio
              If Memory(False) > 0 Then RleRatio = 1 - RleBytes / Memory(False)
            End If
            'load ending (if present)
            Ending = vbNullString
            Do While Seek(1) <= FileSize
              Get #1, , CommentByte
              Ending = Ending & Chr(CommentByte)
            Loop
            FileName = Name
            'image loaded
            LoadTga = True
          End If
        End If
      End If
    Close #1
  End If
  'load failed? reset
  If Not LoadTga Then Destroy
  'error handling
  Exit Function
Failed:
  'memory clenup
  Close #1
  LoadTga = False
  Destroy
End Function

'reset
Private Sub Class_Initialize()
  Destroy
End Sub

'memory cleanup
Private Sub Class_Terminate()
  Destroy
  Erase PixelMap()
End Sub
