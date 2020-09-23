Attribute VB_Name = "ACM_Defs"
Option Explicit

Public Recording As Boolean             ' Public Recording Status Indicator...
Public RecDeviceFree As Boolean         ' Public Recording Device Status Indicator...
Public Playing As Boolean               ' Public Recording Status Indicator...
Public PlayDeviceFree As Boolean        ' Public Recording Device Status Indicator...

Public waveChunkSize As Long            ' size of wave data buffer
Public waveCodec As Long                ' acm codec compression format
Public TIMESLICE As Single              ' recording interval...


'== ACM API Constants ================================================
Public Const ACMERR_BASE = 512
Public Const ACMERR_NOTPOSSIBLE = (ACMERR_BASE + 0)
Public Const ACMERR_BUSY = (ACMERR_BASE + 1)
Public Const ACMERR_UNPREPARED = (ACMERR_BASE + 2)
Public Const ACMERR_CANCELED = (ACMERR_BASE + 3)

' AcmStreamSize Flags...
Public Const ACM_STREAMSIZEF_SOURCE = &H0&
Public Const ACM_STREAMSIZEF_DESTINATION = &H1&
Public Const ACM_STREAMSIZEF_QUERYMASK = &HF&

' acmStreamConvert Flags...
Public Const ACM_STREAMCONVERTF_BLOCKALIGN = &H4&
Public Const ACM_STREAMCONVERTF_START = &H10&
Public Const ACM_STREAMCONVERTF_END = &H20&

' Done Bits For ACMSTREAMHEADER.fdwStatus
Public Const ACMSTREAMHEADER_STATUSF_DONE = &H10000
Public Const ACMSTREAMHEADER_STATUSF_PREPARED = &H20000
Public Const ACMSTREAMHEADER_STATUSF_INQUEUE = &H100000

' Done Bits For acmStreamOpen Formats
Public Const ACM_STREAMOPENF_QUERY = &H1&
Public Const ACM_STREAMOPENF_ASYNC = &H2&
Public Const ACM_STREAMOPENF_NONREALTIME = &H4&

'== ACM API Declarations ================================================
Public Declare Function acmStreamOpen Lib "MSACM32" (hAS As Long, ByVal hADrv As Long, wfxSrc As WAVEFORMATEX, wfxDst As WAVEFORMATEX, ByVal wFltr As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Public Declare Function acmStreamClose Lib "MSACM32" (ByVal hAS As Long, ByVal dwClose As Long) As Long
Public Declare Function acmStreamPrepareHeader Lib "MSACM32" (ByVal hAS As Long, hASHdr As ACMSTREAMHEADER, ByVal dwPrepare As Long) As Long
Public Declare Function acmStreamUnprepareHeader Lib "MSACM32" (ByVal hAS As Long, hASHdr As ACMSTREAMHEADER, ByVal dwUnPrepare As Long) As Long
Public Declare Function acmStreamConvert Lib "MSACM32" (ByVal hAS As Long, hASHdr As ACMSTREAMHEADER, ByVal dwConvert As Long) As Long
Public Declare Function acmStreamReset Lib "MSACM32" (ByVal hAS As Long, ByVal dwReset As Long) As Long
Public Declare Function acmStreamSize Lib "MSACM32" (ByVal hAS As Long, ByVal cbInput As Long, dwOutBytes As Long, ByVal dwSize As Long) As Long

'== ACM User Defined Datatypes ================================================
Type WAVEFILTER
    cbStruct      As Long
    dwFilterTag   As Long
    fdwFilter     As Long
    dwReserved(5) As Long
End Type

Type ACMSTREAMHEADER            ' [ACM STREAM HEADER TYPE]
    cbStruct As Long            ' Size of header in bytes
    dwStatus As Long            ' Conversion status buffer
    dwUser As Long              ' 32 bits of user data specified by application
    pbSrc As Long               ' Source data buffer pointer
    cbSrcLength As Long         ' Source data buffer size in bytes
    cbSrcLengthUsed As Long     ' Source data buffer size used in bytes
    dwSrcUser As Long           ' 32 bits of user data specified by application
    cbDst As Long               ' Dest data buffer pointer
    cbDstLength As Long         ' Dest data buffer size in bytes
    cbDstLengthUsed As Long     ' Dest data buffer size used in bytes
    dwDstUser As Long           ' 32 bits of user data specified by application
    dwReservedDriver(9) As Long ' Reserved and should not be used
End Type

Public Function RecordWave(hWND As Long, ByVal TCPSocket As Variant) As Boolean
' Records Audio Sounds To A String Buffer And Sends Buffer To TCP/IP Socket...
'------------------------------------------------------------------
    Dim rc As Long                                      ' Function Return Code
    Dim hAS As Long                                     ' ACM stream device
    Dim cWavefmt As WAVEFORMATEX                        ' Wave compression format
    Dim acmHdr As ACMSTREAMHEADER                       ' ACM stream header
    Dim acmHdr_x As ACMSTREAMHEADER                     ' <<Double Buffering>> ACM stream header
    Dim hWaveIn As Long                                 ' Handle To An Input Wave Device
    Dim waveFmt As WAVEFORMATEX                         ' Wave compression format
    Dim WaveInHDR As WAVEHDR                            ' Handle To An Input Wave Device Header
    Dim WaveInHDR_x As WAVEHDR                          ' <<Double Buffering>> Handle To An xtra Input Wave Device Header
'------------------------------------------------------------------
    RecDeviceFree = False                               ' Allocate Recording Device
    
    Do While Not PlayDeviceFree                         ' Wait For Play Device To Free
        DoEvents                                        ' Yield Events...
    Loop                                                ' Check Play Device Status
    
    Call InitWaveFormat(waveFmt, WAVE_FORMAT_PCM, TIMESLICE)   ' Set current wave format
    
    ' Open Input Wave Device, Let WAVE_MAPPER Pick The Best Device...
    rc = waveInOpen(hWaveIn, WAVE_MAPPER, waveFmt, 0&, 0&, CALLBACK_NULL)
    If Not AudioErrorHandler(rc, "WaveInOpen") Then Exit Function ' Validate Function Return Code
    
    '<<Double Buffering>> Initialize Wave Header Format Information
    Call InitWaveHDR(WaveInHDR_x, waveFmt, (waveFmt.nAvgBytesPerSec * TIMESLICE))
        
    ' Initialize Wave Header Format Information
    Call InitWaveHDR(WaveInHDR, waveFmt, (waveFmt.nAvgBytesPerSec * TIMESLICE))
    
    ' <<Double Buffering>> Prepare Input Wave Device Header
    rc = waveInPrepareHeader(hWaveIn, WaveInHDR_x, Len(WaveInHDR_x)) ' Validate Return Code
    If Not AudioErrorHandler(rc, "waveInPrepareHeader_x") Then GoTo ErrorRecordWave

    ' Prepare Input Wave Device Header
    rc = waveInPrepareHeader(hWaveIn, WaveInHDR, Len(WaveInHDR)) ' Validate Return Code
    If Not AudioErrorHandler(rc, "waveInPrepareHeader") Then GoTo ErrorRecordWave
    
    ' <<Double Buffering>> Wait For Wave (xtra)Header CallBack
    Call WaitForCallBack(WaveInHDR_x.dwFlags, WHDR_PREPARED)
    
    ' Wait For Wave Header CallBack
    Call WaitForCallBack(WaveInHDR.dwFlags, WHDR_PREPARED)
    
    ' <<Double Buffering>> Add Input Wave (xtra)Buffer To Wave Input Device
    rc = waveInAddBuffer(hWaveIn, WaveInHDR_x, Len(WaveInHDR_x)) ' Validate Return Code
    If Not AudioErrorHandler(rc, "waveInAddBuffer_x") Then GoTo ErrorRecordWave

    ' Add Input Wave Buffer To Wave Input Device
    rc = waveInAddBuffer(hWaveIn, WaveInHDR, Len(WaveInHDR)) ' Validate Return Code
    If Not AudioErrorHandler(rc, "waveInAddBuffer") Then GoTo ErrorRecordWave
        
    ' <<Double Buffering>> Wait For Wave (xtra)Header CallBack
    Call WaitForCallBack(WaveInHDR_x.dwFlags, WHDR_PREPARED)
    
    ' Wait For Wave Header CallBack
    Call WaitForCallBack(WaveInHDR.dwFlags, WHDR_PREPARED)
    
    Call InitWaveFormat(cWavefmt, waveCodec, TIMESLICE)   ' Set current wave format
    
    ' Open/Configure an acm Stream Handle For Compression
    rc = acmStreamOpen(hAS, 0&, waveFmt, cWavefmt, 0&, 0&, 0&, ACM_STREAMOPENF_NONREALTIME)
    Call AudioErrorHandler(rc, "acmStreamOpen")
    
    ' Initialize Audio Compression Manager Streaming Headers
    Call InitAcmHDR(hAS, acmHdr, WaveInHDR)
    Call InitAcmHDR(hAS, acmHdr_x, WaveInHDR_x)
    
    ' Prepare acm Stream Header
    rc = acmStreamPrepareHeader(hAS, acmHdr, 0&)
    Call AudioErrorHandler(rc, "acmStreamPrepareHeader")
    
    ' Prepare acm Stream Header
    rc = acmStreamPrepareHeader(hAS, acmHdr_x, 0&)
    Call AudioErrorHandler(rc, "acmStreamPrepareHeader_x")
        
    ' <<Double Buffering>> Wait For Wave (xtra)Header CallBack
    Call WaitForACMCallBack(acmHdr_x.dwStatus, ACMSTREAMHEADER_STATUSF_PREPARED)
    
    ' Wait For Wave Header CallBack
    Call WaitForACMCallBack(acmHdr.dwStatus, ACMSTREAMHEADER_STATUSF_PREPARED)
    
    ' Start Input Wave Device Recording...
    rc = waveInStart(hWaveIn)                           ' Validate Return Code
    If Not AudioErrorHandler(rc, "waveInStart") Then GoTo ErrorRecordWave
    
    Do
        ' <<Double Buffering>> Wait For Wave (xtra)Header CallBack
        Call WaitForCallBack(WaveInHDR_x.dwFlags, WHDR_DONE)
    
        ' <<Double Buffering>> Compress acm Stream Wave Buffer
        rc = acmStreamConvert(hAS, acmHdr_x, ACM_STREAMCONVERTF_BLOCKALIGN)
        If Not AudioErrorHandler(rc, "acmStreamConvert_x") Then GoTo ErrorRecordWave
        'Alter the code to return a Binary array
        'rc = SendSoundAll(TCPSocket, acmHdr_x) ' <<Double Buffering>> Send Sound Buffer To TCPSocket
        '****************************************
        If Not Recording Then Exit Do                       ' Evaluate Recording Stop Flag
        
        ' <<Double Buffering>> Add Input Wave (xtra)Buffer To Wave Input Device
        rc = waveInAddBuffer(hWaveIn, WaveInHDR_x, Len(WaveInHDR_x)) ' Validate Return Code
        If Not AudioErrorHandler(rc, "waveInAddBuffer_x") Then GoTo ErrorRecordWave
        
        Call WaitForCallBack(WaveInHDR.dwFlags, WHDR_DONE)  ' Wait For Wave Header CallBack
        
        ' Convert/Compress acm Stream Wave Buffer
        rc = acmStreamConvert(hAS, acmHdr, ACM_STREAMCONVERTF_BLOCKALIGN)
        If Not AudioErrorHandler(rc, "acmStreamConvert") Then GoTo ErrorRecordWave
        'Alter for Binary Array
        'rc = SendSoundAll(TCPSocket, acmHdr)           ' Send Sound Buffer To TCPSocket
        '**************************************8
        If Not Recording Then Exit Do                       ' Evaluate Recording Stop Flag
    
        ' Add Input Wave Buffer To Wave Input Device
        rc = waveInAddBuffer(hWaveIn, WaveInHDR, Len(WaveInHDR)) ' Validate Return Code
        If Not AudioErrorHandler(rc, "waveInAddBuffer") Then GoTo ErrorRecordWave
    Loop While Recording                                   ' Continue Recording...
    
    ' <<Double Buffering>> UnPrepare acm Stream Header
    rc = acmStreamUnprepareHeader(hAS, acmHdr_x, 0&)
    Call AudioErrorHandler(rc, "acmStreamUnprepareHeader_x")
    
    ' UnPrepare acm Stream Header
    rc = acmStreamUnprepareHeader(hAS, acmHdr, 0&)
    Call AudioErrorHandler(rc, "acmStreamUnprepareHeader")
    
    ' Free globally allocated and locked memory variables...
    Call FreeAcmHdr(acmHdr_x)                           ' Free extra wave header memory
    Call FreeAcmHdr(acmHdr)                             ' Free wave header memory
    
    ' Close acm Stream Handle
    rc = acmStreamClose(hAS, 0&)
    Call AudioErrorHandler(rc, "acmStreamClose")
    
    ' <<Double Buffering>> Wait For Wave (xtra)Header CallBack
    Call WaitForCallBack(WaveInHDR_x.dwFlags, WHDR_DONE)
    
    ' Wait For Wave Header CallBack
    Call WaitForCallBack(WaveInHDR.dwFlags, WHDR_DONE)
    
    ' Stop Input Wave Device
    rc = waveInStop(hWaveIn)                            ' Validate Return Code
    If Not AudioErrorHandler(rc, "waveInStop") Then GoTo ErrorRecordWave
   
   ' UnPrepare Input Wave Device Header
    rc = waveInUnprepareHeader(hWaveIn, WaveInHDR, Len(WaveInHDR)) ' Validate Return Code
    If Not AudioErrorHandler(rc, "waveInUnPrepareHeader") Then GoTo ErrorRecordWave
   
   ' <<Double Buffering>> UnPrepare Input Wave Device (xtra)Header
    rc = waveInUnprepareHeader(hWaveIn, WaveInHDR_x, Len(WaveInHDR_x)) ' Validate Return Code
    If Not AudioErrorHandler(rc, "waveInUnPrepareHeader_x") Then GoTo ErrorRecordWave
    
    ' Close Input Wave Device
    rc = waveInClose(hWaveIn)                           ' Validate Return Code
    If Not AudioErrorHandler(rc, "waveInClose") Then Exit Function
    
    ' Clean Up Memory Data...
    rc = FreeWaveHDR(WaveInHDR)                         ' Free Wave Header Data
    rc = FreeWaveHDR(WaveInHDR_x)                       ' Free Extra Wave Header Data
    
    RecordWave = True                                   ' Return Success
    RecDeviceFree = True                                ' Free Recording Device
    Exit Function                                       ' Exit
'------------------------------------------------------------------
ErrorRecordWave:                                        ' Clean Up Environment(Brute force no error handling)...
'------------------------------------------------------------------
    rc = acmStreamUnprepareHeader(hAS, acmHdr, 0&)      ' Attempt To UnPrepare acm Stream Header
    rc = acmStreamUnprepareHeader(hAS, acmHdr_x, 0&)    ' Attempt To UnPrepare acm Stream (xtra)Header
    Call FreeAcmHdr(acmHdr)                             ' Free wave header memory
    Call FreeAcmHdr(acmHdr_x)                           ' Free extra wave header memory
    rc = acmStreamClose(hAS, 0&)                        ' Attempt To Close acm Stream Handle
    
    rc = waveInStop(hWaveIn)                            ' Attempt To Stop WaveInput Device
    rc = waveInReset(hWaveIn)                           ' Attempt To Reset WaveInput Device
    rc = waveInUnprepareHeader(hWaveIn, WaveInHDR, Len(WaveInHDR)) ' Attempt To Unprepare WaveInput Header
    rc = waveInUnprepareHeader(hWaveIn, WaveInHDR_x, Len(WaveInHDR_x)) ' Attempt To Unprepare WaveInput (xtra)Header
    rc = waveInClose(hWaveIn)                           ' Attempt To Close Wave Input Device
    rc = FreeWaveHDR(WaveInHDR)                         ' Free Wave Header Data
    rc = FreeWaveHDR(WaveInHDR_x)                       ' Free Extra Wave Header Data
    
    RecDeviceFree = True                                ' Free Recording Device
    Exit Function                                       ' Exit
'------------------------------------------------------------------
End Function
'------------------------------------------------------------------

Private Sub InitWaveFormat(waveFmt As WAVEFORMATEX, fmtType As Long, Time_Slice As Single)
' Initializes Wave Format Data Type
'------------------------------------------------------------------
    Dim i As Long
'------------------------------------------------------------------
    Select Case fmtType
    Case WAVE_FORMAT_ADPCM
        waveFmt.wFormatTag = WAVE_FORMAT_ADPCM          ' wave format type
        waveFmt.nChannels = 1                           ' number of channels - mono
        waveFmt.wBitsPerSample = 4                      ' bits/sample of TRUESPEECH - not used.
        waveFmt.nSamplesPerSec = c8_0kHz                ' sample rate kHz
        waveFmt.nAvgBytesPerSec = 4055                  ' Bytes/Sec
        waveFmt.nBlockAlign = 256                       ' block size of data
        waveFmt.cbSize = 2                              ' extra bytes used for WaveFormatEx
        waveFmt.xBytes(0) = &HF9                        ' Fact Chunk - Byte 0
        waveFmt.xBytes(1) = &H1                         ' Fact Chunk - Byte 1
    Case WAVE_FORMAT_MSN_AUDIO          ' Initialize Wave Format - WAVE_FORMAT_MSN_AUDIO
        waveFmt.wFormatTag = WAVE_FORMAT_MSN_AUDIO      ' wave format type
        waveFmt.nChannels = 1                           ' number of channels - mono
        waveFmt.wBitsPerSample = 0                      ' bits/sample of TRUESPEECH - not used.
        waveFmt.cbSize = 4                              ' extra bytes used for WaveFormatEx
        waveFmt.xBytes(0) = &H40                        ' Fact Chunk - Byte 0
        waveFmt.xBytes(1) = &H1                         ' Fact Chunk - Byte 1
'<<< 8.0 kHz - 8200 Bauds >>>  (Fair, No FeedBack)
        waveFmt.nSamplesPerSec = c8_0kHz                ' sample rate kHz
        waveFmt.nAvgBytesPerSec = 1025                  ' Bytes/Sec
        waveFmt.nBlockAlign = 41                        ' block size of data
        waveFmt.xBytes(2) = &H8                         ' Fact Chunk - Byte 2
        waveFmt.xBytes(3) = &H20                        ' Fact Chunk - Byte 3
'<<< 8.0 kHz - 10000 Bauds >>> (Excellent, No FeedBack)
'        WaveFmt.nSamplesPerSec = c8_0kHz                ' sample rate kHz
'        WaveFmt.nAvgBytesPerSec = 1250                  ' Bytes/Sec
'        WaveFmt.nBlockAlign = 50                        ' block size of data
'        WaveFmt.xBytes(2) = &H10                        ' Fact Chunk - Byte 2
'        WaveFmt.xBytes(3) = &H27                        ' Fact Chunk - Byte 3
'<<< 11.025 kHz - 11301 Bauds >>> (Bad, FeedBack)
'<<< 11.025 kHz - 12128 Bauds >>> (Bad, FeedBack)
'<<< 11.025 kHz - 13782 Bauds >>> (Bad, FeedBack)
    Case WAVE_FORMAT_GSM610             ' Initialize Wave Format - WAVE_FORMAT_GSM610
        waveFmt.wFormatTag = WAVE_FORMAT_GSM610         ' wave format type
        waveFmt.nChannels = 1                           ' number of channels - mono
        waveFmt.nSamplesPerSec = c8_0kHz                ' sample rate kHz
        waveFmt.nAvgBytesPerSec = 1625                  ' Bytes/Sec
        waveFmt.nBlockAlign = 65                        ' block size of data
        waveFmt.wBitsPerSample = 0                      ' bits/sample of TRUESPEECH - not used.
        waveFmt.cbSize = 2                              ' extra bytes used for WaveFormatEx
        waveFmt.xBytes(0) = &H40                        ' Fact Chunk - Byte 0
        waveFmt.xBytes(1) = &H1                         ' Fact Chunk - Byte 1
    Case WAVE_FORMAT_PCM                ' Initialize Wave Format - WAVE_FORMAT_PCM
        waveFmt.wFormatTag = WAVE_FORMAT_PCM                ' format type
        waveFmt.nChannels = WAVE_FORMAT_1M08                ' number of channels (i.e. mono, stereo, etc.)
        waveFmt.nSamplesPerSec = c8_0kHz                    ' sample rate 8.0 kHz
        waveFmt.nAvgBytesPerSec = waveFmt.nSamplesPerSec    ' for buffer estimation
        waveFmt.wBitsPerSample = 8                          ' [8, 16, or 0]
        waveFmt.nBlockAlign = waveFmt.nChannels * waveFmt.wBitsPerSample / 8 '  block size of data
        waveFmt.cbSize = 0                                  ' Not Used If [wFormatTag= WAVE_FORMAT_PCM]
    End Select
'------------------------------------------------------------------
End Sub

Public Function AudioErrorHandler(rc As Long, fcnName As String) As Boolean
'------------------------------------------------------------------
    Dim msg As String               ' Error Message Body
'------------------------------------------------------------------
    AudioErrorHandler = False       ' Return Failure
    
'   Select Case rc Or Err.LastDllError
    Select Case rc
    Case MMSYSERR_NOERROR           ' no error
        AudioErrorHandler = True    ' Return Success
        Exit Function               ' Exit Function
    Case MMSYSERR_ERROR             ' unspecified error
        msg = "Unspecified Error."
    Case MMSYSERR_BADDEVICEID       ' device ID out of range
        msg = "device ID out of range"
    Case MMSYSERR_NOTENABLED        ' driver failed enable
        msg = "driver failed enable"
    Case MMSYSERR_ALLOCATED         ' device already allocated
        msg = "device already allocated"
    Case MMSYSERR_INVALHANDLE       ' device handle is invalid
        msg = "device handle is invalid"
    Case MMSYSERR_NODRIVER          ' no device driver present
        msg = "no device driver present"
    Case MMSYSERR_NOMEM             ' memory allocation error
        msg = "memory allocation error"
    Case MMSYSERR_NOTSUPPORTED      ' function isn't supported
        msg = "function isn't supported"
    Case MMSYSERR_BADERRNUM         ' error value out of range
        msg = "error value out of range"
    Case MMSYSERR_INVALFLAG         ' invalid flag passed
        msg = "invalid flag passed"
    Case MMSYSERR_INVALPARAM        ' invalid parameter passed
        msg = "invalid parameter passed"
    Case MMSYSERR_LASTERROR         ' last error in range
        msg = "last error in range"
    Case WAVERR_BADFORMAT           ' unsupported wave format
        msg = "unsupported wave format"
    Case WAVERR_STILLPLAYING        ' still something playing
        msg = "still something playing"
    Case WAVERR_UNPREPARED          ' header not prepared
        msg = "header not prepared"
    Case WAVERR_LASTERROR           ' last error in range
        msg = "last error in range"
    Case WAVERR_SYNC                ' device is synchronous
        msg = "device is synchronous"
    Case ACMERR_NOTPOSSIBLE         ' The requested operation cannot be performed
        msg = "The requested operation cannot be performed"
    Case ACMERR_BUSY                ' The stream header specified is currently in use and cannot be unprepared
        msg = "The acm stream header busy"
    Case ACMERR_UNPREPARED
        msg = "The acm stream header is not prepared"
    Case ACMERR_CANCELED
        msg = "The acm operation has been canceled"
    Case ERROR_SHARING_VIOLATION    ' The process cannot access the file because it is being used by another process.
        msg = "The process cannot access the file because it is being used by another process."
    Case Else                       ' Unknown MM Error!
        msg = "Unknown MM Error!"
    End Select
    
    ' Format Text Body Of Message
    msg = "Error In " & fcnName & _
          " rc= " & Str$(rc) & _
          " MSG= " & msg & _
          " LastDllError= " & Hex(err.LastDllError) & _
          " Source= " & err.Source & vbCrLf
    
    Debug.Print msg                 ' Print Error Message
    MsgBox msg
    Exit Function                   ' Exit
'------------------------------------------------------------------
End Function
'------------------------------------------------------------------



Private Sub InitWaveHDR(WaveHeader As WAVEHDR, waveFmt As WAVEFORMATEX, BuffSize As Long)
' Initialize's An Input Wave Header's DataBuffer And Size Members...
'--------------------------------------------------------------
    Dim rc As Long                                      ' Function Return Code...
'--------------------------------------------------------------
    'WaveHeader.hData = GlobalAlloc(GMEM_MOVEABLE Or GMEM_SHARE Or GMEM_ZEROINIT, BuffSize) ' Allocate Global Memory
    'WaveHeader.lpData = GlobalLock(WaveHeader.hData)    ' Lock Memory handle

    'WaveHeader.dwBufferLength = BuffSize                ' Get Wave Buffer Size
    'WaveHeader.dwFlags = 0                              ' Must Be Set To 0 For (waveOutPrepareHeader & waveInPrepareHeader)
'--------------------------------------------------------------
End Sub

Public Sub WaitForCallBack(CallBackBit As Long, cbFlag As Long)
' Waits For Asynchronous Function Callback Bit To Be Set.
'--------------------------------------------------------------
    Do Until (((CallBackBit And cbFlag) = cbFlag) Or _
               (CallBackBit = WHDR_PREPARED) Or _
               (CallBackBit = 0))       ' Check For (CallBack Bit Or Null)...
        DoEvents                        ' Post Events...
    Loop
'--------------------------------------------------------------
End Sub

Private Sub InitAcmHDR(hAS As Long, acmHdr As ACMSTREAMHEADER, wavHdr As WAVEHDR)
' Initialize's An Input Wave Header's DataBuffer And Size Members...
'--------------------------------------------------------------
    Dim rc As Long                                      ' Function Return Code...
    Dim OutBytes As Long
'--------------------------------------------------------------
    acmHdr.cbStruct = Len(acmHdr)                       ' Size of header in bytes
    acmHdr.dwStatus = 0                                 ' Must be initialized to 0
    acmHdr.dwUser = 0                                   ' clear user def info
    acmHdr.cbSrcLengthUsed = 0                          ' Must be initialized to 0
    acmHdr.cbDstLengthUsed = 0                          ' Must be initialized to 0
    
    acmHdr.pbSrc = wavHdr.lpData                        ' Copy address of unprocessed data
    acmHdr.cbSrcLength = wavHdr.dwBufferLength          ' Copy size of unprocessed data
    
    rc = acmStreamSize(hAS, acmHdr.cbSrcLength, acmHdr.cbDstLength, ACM_STREAMSIZEF_SOURCE)
    Call AudioErrorHandler(rc, "acmStreamSize")
    
    ' Allocate memory for de/compression
    acmHdr.dwDstUser = GlobalAlloc(GMEM_MOVEABLE Or GMEM_SHARE Or GMEM_ZEROINIT, acmHdr.cbDstLength)                                     ' Allocate Global Memory
    acmHdr.cbDst = GlobalLock(acmHdr.dwDstUser)         ' Lock Memory handle
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------

Public Sub WaitForACMCallBack(CallBackBit As Long, cbFlag As Long)
' Waits For Asynchronous Function Callback Bit To Be Set.
'--------------------------------------------------------------
    Do Until (((CallBackBit And cbFlag) = cbFlag) Or _
               (CallBackBit = 0))       ' Check For (CallBack Bit Or Null)...
        DoEvents                        ' Post Events...
    Loop
'--------------------------------------------------------------
End Sub
Private Sub FreeAcmHdr(acmHdr As ACMSTREAMHEADER)
' Initialize's An Input Wave Header's DataBuffer And Size Members...
'--------------------------------------------------------------
    Dim rc As Long                                      ' Function Return Code...
'--------------------------------------------------------------
    rc = GlobalUnlock(acmHdr.cbDst)                     ' Unlock Global Memory
    rc = GlobalFree(acmHdr.dwDstUser)                   ' Free Global Memory
'--------------------------------------------------------------
End Sub
Private Function FreeWaveHDR(WaveHeader As WAVEHDR) As Boolean
'--------------------------------------------------------------
    Dim rc As Long                                      ' Function return code
'--------------------------------------------------------------
    rc = GlobalUnlock(WaveHeader.lpData)                ' Unlock Global Memory
    rc = GlobalFree(WaveHeader.hData)                   ' Free Global Memory
    
    FreeWaveHDR = True                                  ' Set Default Return Code
'--------------------------------------------------------------
End Function

