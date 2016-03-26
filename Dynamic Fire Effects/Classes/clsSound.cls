VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Module:        clsPackage
' Author:        Warren Galyen
' Website:       http://www.mechanikadesign.com
' Dependencies:  None
' Last revision: 2006.06.26
'================================================

'------ FIRE SOUND CLASS -----

Option Base 0
Option Explicit

Private Type tDSStream
  Looping As Boolean
  Descriptor As DSBUFFERDESC
  SecondaryBuffer As DirectSoundSecondaryBuffer8
End Type

' Core objects, stream buffer
Private cDirectSound As DirectSound8
Private tDSBuffer() As tDSStream

' Additional objects to load sound form package
Private btRaw() As Byte
Private cFile As clsBinaryFile

Private lngIndex As Long

Public Sub devRelease()
  'destroy streams
  For lngIndex = UBound(tDSBuffer()) To 0 Step -1
    smDestroy lngIndex
  Next lngIndex
  'erase buffer
  Erase tDSBuffer()
  'release directsound
  Set cDirectSound = Nothing
End Sub

'startup
Public Sub devInitialize(dxObject As DirectX8, lnkHandle As Long)
  'create direct sound
  Set cDirectSound = dxObject.DirectSoundCreate(vbNullString)
  cDirectSound.SetCooperativeLevel lnkHandle, DSSCL_PRIORITY
  '1 sound stream by default
  ReDim tDSBuffer(0)
End Sub

'stop playing
Public Sub smStop(iStream As Long)
  tDSBuffer(iStream).SecondaryBuffer.Stop
  'set cursor to 0
  tDSBuffer(iStream).SecondaryBuffer.SetCurrentPosition 0
End Sub

'toggle sound looping
Public Sub smRepeatEnable(iStream As Long, Optional swRepeat As Boolean = True)
  tDSBuffer(iStream).Looping = swRepeat
End Sub

'play sound
Public Sub smPlay(iStream As Long)
  If tDSBuffer(iStream).Looping Then
    tDSBuffer(iStream).SecondaryBuffer.Play DSBPLAY_LOOPING
  Else
    tDSBuffer(iStream).SecondaryBuffer.Play DSBPLAY_DEFAULT
  End If
End Sub

'create stream
Public Function smCreate(iStream As Long, vfName As String, pkObject As clsPackage) As Boolean
  With tDSBuffer(iStream)
    'create new virtual file
    Set cFile = New clsBinaryFile
    ReDim btRaw(0)
    'load file into memory
    If cFile.vfOpen(vfName, pkObject) Then
      .Looping = False
      'read file header & setup stream format
      With .Descriptor.fxFormat
        .lAvgBytesPerSec = cFile.vfLong
        .lSamplesPerSec = cFile.vfLong
        .nBitsPerSample = cFile.vfInteger
        .nBlockAlign = cFile.vfInteger
        .nChannels = cFile.vfInteger
        .nFormatTag = cFile.vfInteger
      End With
      'get sound data
      cFile.vfArray cFile.lngLength + 1 - 16, btRaw()
      'configure stream
      .Descriptor.lBufferBytes = UBound(btRaw()) + 1
      .Descriptor.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN
      'create sound buffer
      Set .SecondaryBuffer = cDirectSound.CreateSoundBuffer(.Descriptor)
      'write sound data into buffer
      .SecondaryBuffer.SetCurrentPosition 0
      .SecondaryBuffer.WriteBuffer 0, .Descriptor.lBufferBytes, btRaw(0), DSBLOCK_ENTIREBUFFER
      'ok
      smCreate = True
    Else
      'no file or package
      smCreate = False
    End If
    'memory cleanup
    Erase btRaw()
    cFile.vfClose
    Set cFile = Nothing
  End With
End Function

'set volume
Public Sub smVolume(iStream As Long, smMaxVolume As Long)
  tDSBuffer(iStream).SecondaryBuffer.SetVolume smMaxVolume
End Sub

'resize stream buffer
Public Sub smBuffer(iStreams As Long)
  ReDim Preserve tDSBuffer(iStreams)
End Sub

'destroy stream
Public Sub smDestroy(iStream As Long)
  With tDSBuffer(iStream)
    .Descriptor.lBufferBytes = 0
    'destroy buffer
    Set .SecondaryBuffer = Nothing
  End With
End Sub

