Attribute VB_Name = "basSoundRec_1"
Option Explicit

      Public Const MMSYSERR_NOERROR = 0
      Public Const MAXPNAMELEN = 32
      Public Const MIXER_LONG_NAME_CHARS = 64
      Public Const MIXER_SHORT_NAME_CHARS = 16
      Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
      Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
      Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
      Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
      Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
      
      Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
                     (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
                     
      Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = _
                     (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
      
      Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = _
                     (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
      
      Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
      Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
      
      Public Const MIXERCONTROL_CONTROLTYPE_FADER = _
                     (MIXERCONTROL_CT_CLASS_FADER Or _
                     MIXERCONTROL_CT_UNITS_UNSIGNED)
      
      Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = _
                     (MIXERCONTROL_CONTROLTYPE_FADER + 1)




Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&


Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Public Const CALLBACK_FUNCTION = &H30000
Public Const CALLBACK_WINDOW = &H10000      '  dwCallback is a HWND
Public Const MM_WIM_DATA = &H3C0
Public Const WHDR_DONE = &H1         '  done bit
Public Const WIM_DATA = MM_WIM_DATA
Public Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin
Public Const NUM_BUFFERS = 10
Public Const BUFFER_SIZE = 8192
Public Const DEVICEID = 0
Public Const GWL_WNDPROC = -4

Public Const WINSOCK_SERVER = 1
Public hWaveIn As Long
Public hmixer As Long                  ' mixer handle
Public rc As Long                       ' return code
Public ok As Boolean                    ' boolean return code
Public mxcd As MIXERCONTROLDETAILS          ' control info
Public volCtrl As MIXERCONTROL  ' waveout volume control
Public outputVolCtrl As MIXERCONTROL
Public micCtrl As MIXERCONTROL  ' microphone volume control
Public vol As MIXERCONTROLDETAILS_SIGNED    ' control's signed value
Public volume As Long                       ' volume value
Public volHmem As Long                      ' handle to volume memory
Public pPos As Long
Public X, y As Long

Type WAVEHDR
   lpData As Long          ' Address of the waveform buffer.
   dwBufferLength As Long  ' Length, in bytes, of the buffer.
   dwBytesRecorded As Long ' When the header is used in input, this member specifies how much
                           ' data is in the buffer.

   dwUser As Long          ' User data.
   dwFlags As Long         ' Flags supplying information about the buffer. Set equal to zero.
   dwLoops As Long         ' Number of times to play the loop. Set equal to zero.
   lpNext As Long          ' Not used
   reserved As Long        ' Not used
End Type

Type MIXERCONTROLDETAILS_SIGNED
   lValue As Long
End Type

Type WAVEFORMAT
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long


Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Sub CopyStringFromStruct Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal a As String, p As Any, ByVal cb As Long)
Public Declare Sub CopyStructFromPtr Lib "KERNEL32" _
                     Alias "RtlMoveMemory" _
                     (struct As Any, _
                     ByVal ptr As Long, _
                     ByVal cb As Long)

Public Declare Sub CopyPtrFromStruct Lib "KERNEL32" _
                     Alias "RtlMoveMemory" _
                     (ByVal ptr As Long, _
                     struct As Any, _
                     ByVal cb As Long)
                     
Public Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal msg As Long, _
    ByVal wParam As Long, ByRef lParam As WAVEHDR) As Long

Public Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hWnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public i As Integer
Public j As Integer

Public msg As String * 200

'Public hWaveIn As Long
Public wformat As WAVEFORMAT
Public hMem(NUM_BUFFERS) As Long
Public inHdr(NUM_BUFFERS) As WAVEHDR
Public fRecording As Boolean
Dim lpPrevWndProc As Long

Dim hWnd As Long            ' window handle
'---End the standard wave krap of old Enter Direct X
Public DX7 As New DirectX7, DS As DirectSound
Public BufDesc As DSBUFFERDESC, PCM As WAVEFORMATEX, pcm2 As WAVEFORMATEX

'Primary Buffer Object
Public PBuff As DirectSoundBuffer, PDesc As DSBUFFERDESC
Public DSBuffer As DirectSoundBuffer
Public DSBuffer2 As DirectSoundBuffer

Public curs As DSCURSORS, ByteArray() As Byte, FL As Long, ST As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long




Function StartInput() As Boolean
If fRecording Then
        StartInput = True
        Exit Function
End If
    
    'wformat.wBitsPerSample = 16
    'wformat.nSamplesPerSec = 44100
    'wformat.nBlockAlign = wformat.nChannels * wformat.wBitsPerSample / 8
    'wformat.nAvgBytesPerSec = wformat.nSamplesPerSec * wformat.nBlockAlign
    'wformat.cbSize = Len(wformat)
    
    For i = 0 To NUM_BUFFERS - 1
        hMem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(i).lpData = GlobalLock(hMem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next

    rc = waveInOpen(hWaveIn, -1, wformat, frmChat.hWnd, 0, CALLBACK_WINDOW)
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
End Function

Sub addData(iHdr As WAVEHDR)
Dim iRet As Long
Dim sBuff  As String
    
    rc = waveInAddBuffer(hWaveIn, iHdr, Len(iHdr))
    
    sBuff = Space(BUFFER_SIZE)
    CopyMemory ByVal sBuff, ByVal iHdr.lpData, BUFFER_SIZE
    'mmioWrite hmmioIn, sBuff, BUFFER_SIZE

End Sub
' Stop receiving audio input on the soundcard
Sub StopInput()
Dim iRet As Long
Dim icount As Long
    
    fRecording = False
    iRet = waveInReset(hWaveIn)
    iRet = waveInStop(hWaveIn)
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hMem(i)
    Next
    
    iRet = waveInClose(hWaveIn)
    
    'Ascent out of Data Chunk
    If (mmioAscend(hmmioIn, mmckinfoSubchunkIn, 0) <> 0) Then
        MsgBox "Cannot ascend out of DATA CHUNK"
        mmioClose hmmioIn, 0
    End If
    
    If (mmioAscend(hmmioIn, mmckinfoSubchunkIn, 0) <> 0) Then
        MsgBox "Could Ascend Out Of wformat Chunk"
        mmioClose hmmioIn, 0
        Exit Sub
    End If

    'Asecnd out of the RIFF chunk
    If (mmioAscend(hmmioIn, mmckinfoParentIn, 0) <> 0) Then
        MsgBox "Cannot ascend out of RIFF CHUNK"
        mmioClose hmmioIn, 0
    End If
    
    mmioClose hmmioIn, 0
    
End Sub

Public Sub Initialize(hwndIn As Long)
    hWnd = hwndIn
    lpPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByRef wavhdr As WAVEHDR) As Long

 If uMsg = WIM_DATA Then
    frmChat.ReceiveSoundBytes
    
 End If
 WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, wavhdr)
End Function

Public Sub Initialize_Engine()
Set DS = DX7.DirectSoundCreate(vbNullString) 'Create the DirectSound Object
DS.SetCooperativeLevel frmChat.hWnd, DSSCL_EXCLUSIVE 'Set the Cooperative Level

'Fill the buffer info structures. (format & description)
PCM.nSize = LenB(PCM)
PCM.nFormatTag = WAVE_FORMAT_PCM
PCM.nChannels = 1
PCM.lSamplesPerSec = 44100
PCM.nBitsPerSample = 16
PCM.nBlockAlign = PCM.nBitsPerSample / 8 * PCM.nChannels
PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign
BufDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_STATIC 'No need to set the buffer size.

'Create Primary Buffer
PDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_PRIMARYBUFFER
Set PBuff = DS.CreateSoundBuffer(PDesc, pcm2)
PBuff.SetFormat PCM
End Sub

Public Sub Terminate_Engine() 'Clear buffers from memory &
                              'Kill DirectX Objects

Set DSBuffer = Nothing
Set PBuff = Nothing
Set DS = Nothing
Set DX7 = Nothing
End Sub

Public Sub OpenFile(FileName As String)
FL = FileLen(FileName)
ReDim ByteArray(FL - 44)
Set DSBuffer = DS.CreateSoundBufferFromFile(FileName, BufDesc, PCM)
DSBuffer.ReadBuffer curs.lPlay, 0, ByteArray(0), DSBLOCK_ENTIREBUFFER

'Open FileName For Binary As #1
'Get #1, 44, ByteArray()
'Close #1
End Sub

Sub Play()

DSBuffer.Play DSBPLAY_LOOPING

End Sub
