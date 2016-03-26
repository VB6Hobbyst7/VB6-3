VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Module:        clsPackage
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  None
' Last revision: 2006.06.23
'================================================

'------ PACKAGE EXTRACTOR CLASS -----


Option Explicit
Option Base 0

' File record
Private Type pkEntry
  stOffset As Long   ' data stream offset
  stLength As Long   ' data stream length
  stName As String   ' stream id
End Type

Private pkIndex As Long         ' counter
Private pkOffset As Long        ' another counter
Private pkIdLength As Long      ' file name length
Private pkByte As Byte          ' byte from id
Private pkFat() As pkEntry      ' file allocation table
Private pkHeader As String * 3  ' file header (PKG)

Public pkNameHandle As String   ' package filename

Public pkEntries As Long        ' package entries
Public pkVersion As Long        ' package version

'load package f.a.t. into memory
Public Function pkOpen(pkName As String) As Boolean
  'check file existance
  If Len(Dir(pkName)) > 0 Then
    'reset
    pkClose
    'open package
    Open pkName For Binary Access Read As #1
    'get file header
    Get #1, , pkHeader
    If Not pkHeader = "PKG" Then
      'close file
      Close #1
      pkOpen = False
      Exit Function
    End If
    'get package version
    Get #1, , pkVersion
    If pkVersion > 1 Then
      'unsupported version
      Close #1
      pkOpen = False
      Exit Function
    End If
    'get entries count
    Get #1, , pkEntries
    If pkEntries < 1 Then
      'no files in archive
      Close #1
      pkOpen = False
      Exit Function
    End If
    'create file table
    ReDim pkFat(pkEntries - 1)
    'read files
    For pkIndex = 0 To pkEntries - 1 Step 1
      With pkFat(pkIndex)
        'read data stram position
        Get #1, , .stOffset
        'get stream length
        Get #1, , .stLength
        'get id length
        Get #1, , pkIdLength
        'check name length
        If pkIdLength < 1 Then
          'close file
          pkClose
          pkOpen = False
          Exit Function
        End If
        'read id
        .stName = vbNullString
        For pkOffset = 1 To pkIdLength Step 1
          'read byte
          Get #1, , pkByte
          'create id
          .stName = .stName & Chr(pkByte)
        Next pkOffset
      End With
    Next pkIndex
    'success
    pkOpen = True
    pkNameHandle = pkName
    Close #1
  Else
    'file not found
    pkOpen = False
  End If
End Function

'extract data stream into array
Public Function pkExtract(stName As String, stRaw() As Byte) As Boolean
  'search stream
  For pkIndex = 0 To pkEntries - 1 Step 1
    With pkFat(pkIndex)
      'when stream found
      If .stName = stName Then
        'allocte memory
        ReDim stRaw(.stLength - 1)
        Open pkNameHandle For Binary Access Read As #1
        'move to stream
        Seek #1, .stOffset
        'read bytes
        Get #1, , stRaw()
        'stream extracted, finish
        pkExtract = True
        Close #1
        Exit Function
      End If
    End With
  Next pkIndex
  'stream not found
  pkExtract = False
End Function

'close package
Public Sub pkClose()
  'cleanup memory
  Erase pkFat()
  'reset length
  pkEntries = -1
  'close file
  Close #1
End Sub

'startup
Private Sub Class_Initialize()
  'reset variables
  pkEntries = -1
  ReDim pkFat(0)
End Sub

'shutdown
Private Sub Class_Terminate()
  'memory cleanup
  pkClose
End Sub

