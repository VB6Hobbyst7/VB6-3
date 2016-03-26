VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBinaryFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Module:        clsBinaryFile
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  clsPackage.cls
' Last revision: 2006.06.26
'================================================

'------ VIRTUAL FILE CLASS -----

' History: 26 Jun 06 - Added functions for reading long, integer, and arrays values.

Option Explicit
Option Base 0

Private Declare Sub vfCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private btRaw() As Byte
Private lngSymbol As Long

Public lngOffset As Long
Public lngLength As Long

'read long value
Public Function vfLong() As Long
  'get bytes
  vfCopyMemory vfLong, btRaw(lngOffset), 4
  'move cursor
  lngOffset = lngOffset + 4
End Function

'read integer value
Public Function vfInteger() As Integer
  'get bytes
  vfCopyMemory vfInteger, btRaw(lngOffset), 2
  'move cursor
  lngOffset = lngOffset + 2
End Function

'get byte array
Public Sub vfArray(lngLength As Long, vfResult() As Byte)
  'allocate memory
  ReDim vfResult(lngLength - 1)
  'copy data
  vfCopyMemory vfResult(0), btRaw(lngOffset), lngLength
  'move cursor
  lngOffset = lngOffset + lngLength
End Sub

'load file into memory
Public Function vfOpen(vfName As String, pkObject As clsPackage) As Boolean
  'reset
  vfClose
  'create array
  ReDim btRaw(0)
  'load file
  vfOpen = pkObject.pkExtract(vfName, btRaw())
  'when stream found
  If vfOpen Then
    'reset cursor
    lngOffset = 0
    'ger stream length
    lngLength = UBound(btRaw())
  Else
    'close file
    vfClose
  End If
End Function

'end of file
Public Function vfEof() As Boolean
  'check
  If lngLength < lngOffset Then
    'yes
    vfEof = True
  Else
    'no
    vfEof = False
  End If
End Function

'close file
Public Sub vfClose()
  'cleanup memory
  Erase btRaw()
  'reset length
  lngLength = -1
  'reset cursor
  lngOffset = -1
End Sub

'input line
Public Function vfLine() As String
  'read bytes while not {cr} & {lf}
  vfLine = vbNullString
  For lngSymbol = lngOffset To lngLength - 1 Step 1
    'when found line feed
    If btRaw(lngSymbol) = Asc(vbCr) And btRaw(lngSymbol + 1) = Asc(vbLf) Then
      'move cursor & finish
      lngOffset = lngSymbol + 2
      Exit Function
    End If
    'add symbol to string
    vfLine = vfLine & Chr(btRaw(lngSymbol))
  Next lngSymbol
  'add last symbol
  vfLine = vfLine & Chr(btRaw(lngLength))
  'eof
  lngOffset = lngLength + 1
End Function

'startup
Private Sub Class_Initialize()
  'reset variables
  lngLength = -1
  lngOffset = -1
End Sub

'shutdown
Private Sub Class_Terminate()
  'memory cleanup
  vfClose
End Sub


