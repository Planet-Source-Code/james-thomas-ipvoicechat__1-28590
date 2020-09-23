VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   3840
      TabIndex        =   7
      Top             =   960
      Width           =   1335
      Begin VB.OptionButton opt44 
         Caption         =   "44100 Hz"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opt225 
         Caption         =   "22050 Hz"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton opt11 
         Caption         =   "11025 Hz"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkTalk 
      Caption         =   "Talk"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   3405
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "Network"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3466
            Key             =   "Connection"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3466
            Key             =   "Recording"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Users"
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdConnect 
         Default         =   -1  'True
         Height          =   300
         Left            =   3240
         Picture         =   "frmChat.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Connect"
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox txtConnection 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Show Users"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   3000
         Width           =   1935
      End
      Begin MSComctlLib.TreeView tvwOnlineUsers 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4260
         _Version        =   393217
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   6
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Connect To:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox picPeakVol 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4440
      Top             =   2520
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   4080
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1001
   End
   Begin VB.Menu mnuConnection 
      Caption         =   "Connection"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSounds 
      Caption         =   "Sounds"
      Begin VB.Menu mnuSend 
         Caption         =   "Send"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm" Alias _
           "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) _
           As Long
           
'Dim psa(1) As DSBPOSITIONNOTIFY
Dim dx As New DirectX7
Implements DirectXEvent
Dim b_play As Boolean
Dim bWrite As Boolean
Sub SetEvents()
EventID(0) = dx.CreateEvent(Me)
EventID(1) = dx.CreateEvent(Me)
If gDSCB Is Nothing Then Call Init(hWnd)
EVNT(0).hEventNotify = EventID(0)
EVNT(0).lOffset = 0
EVNT(1).hEventNotify = EventID(1)
EVNT(1).lOffset = (gDSCBD.lBufferBytes \ 2)
gDSCB.SetNotificationPositions 2, EVNT()
    
End Sub
Private Sub chkTalk_Click()
If chkTalk.Value = 1 Then
    Init hWnd
    SetEvents
    StartCapture
Else
    'StopInput   ' Stop receiving audio input
    StopCapture
End If
End Sub

Private Sub cmdConnect_Click()
Select Case wskClient.State
    Case sckListening, sckError
        wskClient.Close
        wskClient.Connect txtConnection, 1001
    Case sckConnected
        If MsgBox("You are currently Connected to " & wskClient.RemoteHostIP & vbCrLf & "Are you sure you want to disconnect?", vbYesNo) = vbYes Then
            wskClient.Close
            wskClient.Connect txtConnection, 1001
        End If
    Case sckClosed
        wskClient.Connect txtConnection, 1001
End Select

Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
hThread = CreateThread(ByVal 0&, ByVal 0&, AddressOf AsyncThread, ByVal 0&, ByVal 0&, hThreadID)
SetThreadPriority hThread, THREAD_PRIORITY_HIGHEST
SetPriorityClass hThread, HIGH_PRIORITY_CLASS
CloseHandle hThread
End Sub

Private Sub DirectXEvent_DXCallback(ByVal EventID As Long)

Select Case EventID
    Case EVNT(1).hEventNotify
        CopyBuffer 2
End Select
End Sub

Private Sub Form_Load()
wskClient.LocalPort = 1001
wskClient.Listen
'lFile = FreeFile
'opt44.Value = True
'' Open the mixer with deviceID.
'rc = mixerOpen(hmixer, 0, 0, 0, 0)
'If ((MMSYSERR_NOERROR <> rc)) Then
'    MsgBox "Couldn't open the mixer."
'    Exit Sub
'End If''

'Get the wavein(Microphone) volume control
'ok = GetVolumeControl(hmixer, _
'                              MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, _
'                              MIXERCONTROL_CONTROLTYPE_VOLUME, _
'                              micCtrl)
'If (ok = True) Then
'    'Initialize Me.hWnd
'Else
'     MsgBox "Couldn't get wavein meter"
'End If

' Get the output volume meter
'ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
'If (ok = True) Then
'    ProgressBar1.Min = 0
'    ProgressBar1.Max = outputVolCtrl.lMaximum
'    'hDisplayThread = CreateThread(ByVal 0&, ByVal 0&, AddressOf DisplayWave, ByVal 0&, ByVal 0&, hDisplayID)
'    'SetThreadPriority hDisplayThread, THREAD_PRIORITY_HIGHEST
'    'SetPriorityClass hDisplayThread, HIGH_PRIORITY_CLASS
'    'CloseHandle hDisplayThread
'    Timer1.Interval = 50
'    Timer1.Enabled = True
'Else
'    MsgBox "Couldn't get waveout meter"
'End If'

   
' Initialize mixercontrol structure
'mxcd.cbStruct = Len(mxcd)
'volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
'mxcd.paDetails = GlobalLock(volHmem)
'mxcd.cbDetails = Len(volume)
'mxcd.cChannels = 1
'
'DX
Initialize_Engine

End Sub



Private Sub Form_Unload(Cancel As Integer)
If hThread <> 0 Then
    ExitThread GetExitCodeThread(hThread, 0)
    TerminateThread hThread, 0
End If
If hDisplayThread <> 0 Then
    ExitThread GetExitCodeThread(hDisplayThread, 0)
    TerminateThread hDisplayThread, 0
End If
wskClient.Close
Terminate_Engine
End
End Sub

Private Sub mnuClose_Click()
wskClient.Close
End Sub

Private Sub mnuExit_Click()
wskClient.Close
Unload Me
End Sub

Private Sub mnuNew_Click()
wskClient.Close
cmdConnect_Click
End Sub

Private Sub mnuSend_Click()
'Me.wskClient.SendData PrepareAVoice(GetWaveRes(102))
End Sub

Private Sub Timer1_Timer()
' Get the current output level
'On Error Resume Next
If (ProgressBar1.Enabled = True) Then
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then volume = -volume
    If volume > ProgressBar1.Max Then
        ProgressBar1.Value = ProgressBar1.Max
    Else
        ProgressBar1.Value = volume
    End If
    If (volume > picPeakVol.Height) Then volume = volume / (picPeakVol.Height / 4)
    y = (picPeakVol.Height / 2)
    picPeakVol.Line (pPos, y - volume)-(pPos, y + volume), vbRed
    lasty = y
    pPos = pPos + 10
    If pPos >= picPeakVol.Width Then               '5000 = picpeakvol.width
        picPeakVol.Cls
        pPos = 0
    End If
    
End If
End Sub

Private Sub tvwOnlineUsers_Expand(ByVal Node As MSComctlLib.Node)
If Node.Tag = "N" Then
   ' NodeExpand Node
End If
End Sub

Private Sub tvwOnlineUsers_NodeClick(ByVal Node As MSComctlLib.Node)
Me.txtConnection = Node.Text
End Sub

Private Sub wskClient_Close()
wskClient.Close
stbStatus.Panels("Connection").Text = "Connection Closed"
End Sub

Private Sub wskClient_Connect()
stbStatus.Panels("Connection").Text = "Connected To: " & wskClient.RemoteHostIP
End Sub

Private Sub wskClient_ConnectionRequest(ByVal requestID As Long)

If MsgBox("You Have An Incoming Connection.", vbYesNo) = vbYes Then
    If wskClient.State <> sckClosed Then
        wskClient.Close
    End If
    wskClient.Accept requestID
    stbStatus.Panels("Connection").Text = "Connected To: " & wskClient.RemoteHostIP
End If

End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
'On Error GoTo fin
Dim ret As Variant
Dim bt() As Byte
Dim Blankbuff() As Byte
Dim RetVal As Variant
Dim RemainSoundBuffer As Long
ReDim bt(bytesTotal) As Byte
wskClient.GetData bt, vbArray Or vbByte, bytesTotal

'bt = ret
If TotalBytesRecieved = 0 Then
    CopyMemory ByVal VarPtr(taciGram.DataType), ByVal VarPtr(bt(0)), LenB(taciGram.DataType)
End If
'After recieving the data gram check the type to see what needs to be processed
Select Case taciGram.DataType
    Case msg_Recordset
        CopyMemory ByVal VarPtr(taciGram.BlobSize), ByVal VarPtr(bt(0)) + LenB(taciGram.DataType), LenB(taciGram.BlobSize)
        ReDim Destbt(taciGram.BlobSize) As Byte
        RetVal = LenB(taciGram.BlobSize) + LenB(taciGram.DataType)
        For ret = RetVal To UBound(bt)
            Destbt(ret - RetVal) = bt(ret)
        Next ret
        'Set Bag = New DataBag
        Bag.Attach Destbt
    Case msg_Message
        Dim lpsize As Long
        Dim lpData As Long
        Dim taciMsg As taci_Message
        CopyMemory ByVal VarPtr(taciGram.BlobSize), ByVal VarPtr(bt(0)) + LenB(taciGram.DataType), LenB(taciGram.BlobSize)
        ReDim Destbt(taciGram.BlobSize) As Byte
        RetVal = LenB(taciGram.BlobSize) + LenB(taciGram.DataType)
        For ret = RetVal To UBound(bt)
            Destbt(ret - RetVal) = bt(ret)
        Next ret
        lpData = AttachFromSend(Destbt, lpsize, msg_Message)
        CopyMemory ByVal VarPtr(taciMsg), ByVal lpData, lpsize
    Case msg_WorkQueue
    
    Case msg_Reminder
    
    Case msg_StationID
        'sckClient(1).SendData PrepareAStationID
        Stop
    Case msg_Voice
        Dim vntFlags As Long
        If TotalBytesRecieved = 0 Then
            CopyMemory ByVal VarPtr(taciGram.BlobSize), ByVal VarPtr(bt(0)) + LenB(taciGram.DataType), LenB(taciGram.BlobSize)
            ReDim Destbt(taciGram.BlobSize) As Byte
            RetVal = LenB(taciGram.BlobSize) + LenB(taciGram.DataType)
            'right here the byte array may be excl uding some bytes
            For ret = RetVal To UBound(bt)
                Destbt(ret - RetVal) = bt(ret)
            Next ret
            TotalBytesRecieved = (bytesTotal - RetVal) - 1
            If taciGram.BlobSize <= TotalBytesRecieved Then bTotalBytesRecieved = True
        Else
            For ret = 0 To UBound(bt)
                Destbt(TotalBytesRecieved + ret) = bt(ret)
            Next ret
            TotalBytesRecieved = TotalBytesRecieved + bytesTotal
            If taciGram.BlobSize <= TotalBytesRecieved Then bTotalBytesRecieved = True
        End If
        If bTotalBytesRecieved = True Then
            'Play the Wave
            stbStatus.Panels("Recording").Text = UBound(Destbt) & "@" & Timer
            'If b_play = False Then
                With WaveF1
                    .nChannels = 1
                    .lExtra = 0
                    .nFormatTag = WAVE_FORMAT_PCM
                    .lSamplesPerSec = SAMPLE_RATE
                    .lAvgBytesPerSec = SAMPLE_RATE
                    .nBlockAlign = 1
                    .nBitsPerSample = 8
                End With
            
                lngNotificationSize = (WaveF1.lSamplesPerSec * 2) \ 2
                BufDesc.lBufferBytes = UBound(Destbt) 'lngNotificationSize * 2
                lngLastBit = (UBound(Destbt) \ BufDesc.lBufferBytes) * BufDesc.lBufferBytes
                                                'Create a half second buffer.
                BufDesc.lFlags = DSBCAPS_GETCURRENTPOSITION2 Or DSBCAPS_CTRLPOSITIONNOTIFY
                Set DSBuffer = DS.CreateSoundBuffer(BufDesc, WaveF1)
                'DSBuffer.Stop
                DSBuffer.WriteBuffer 0, UBound(Destbt), Destbt(0), DSBLOCK_FROMWRITECURSOR
                DSBuffer.Play DSBPLAY_DEFAULT

            'Else
                'don't let the buffer be written over until
                'it is ready to end
                'DSBuffer.GetCurrentPosition cur
                'I need to know if it is ok to write to the buffer
                '
                bWrite = True
                'Do Until bWrite = False
                '    DSBuffer.GetCurrentPosition cur
                '    DoEvents
                'Loop
                'DSBuffer.p
                'DSBuffer.SetCurrentPosition 0
                'DSBuffer.GetCurrentPosition cur
                'DSBuffer.WriteBuffer 0, UBound(Destbt), Destbt(0), DSBLOCK_DEFAULT
                'DSBuffer.SetCurrentPosition 0
                'DSBuffer.GetCurrentPosition cur
                'WritePosition = Abs((cur.lWrite + UBound(Destbt)) - BufDesc.lBufferBytes)
                'psa(0).lOffset = WritePosition
                'psa(0).hEventNotify = EventID(0)
                'DSBuffer.SetNotificationPositions 2, psa()
               'DSBuffer.Play DSBPLAY_DEFAULT
            'End If
            LastByteSize = UBound(Destbt)
            bTotalBytesRecieved = False
            TotalBytesRecieved = 0
        End If
End Select
Exit Sub
fin:
Select Case err.Number
    Case 9 'subscript out of range this is the first time into the sub
        ReDim Destbt(taciGram.BlobSize) As Byte
        Resume Next
    Case Else
        Stop
End Select
End Sub
Public Sub PlayWaveRes(vntResourceID As Variant, Optional vntFlags)
      '-----------------------------------------------------------------
      ' WARNING:  If you want to play sound files asynchronously in
      '           Win32, then you MUST change bytSound() from a local
      '           variable to a module-level or static variable. Doing
      '           this prevents your array from being destroyed before
      '           sndPlaySound is complete. If you fail to do this, you
      '           will pass an invalid memory pointer, which will cause
      '           a GPF in the Multimedia Control Interface (MCI).
      '-----------------------------------------------------------------
      Dim bytSound() As Byte ' Always store binary data in byte arrays!

      bytSound = LoadResData(vntResourceID, "WAVE")

      If IsMissing(vntFlags) Then
         vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
      End If

      If (vntFlags And SND_MEMORY) = 0 Then
         vntFlags = vntFlags Or SND_MEMORY
      End If

      sndPlaySound bytSound(0), vntFlags
      End Sub

Public Function GetWaveRes(resourceID As Long) As Byte()
    GetWaveRes = LoadResData(resourceID, "WAVE")
End Function

Public Function LoadUsers()
fnGetDomainName strDomain '= GetLoginDomain
fnGetPDCName "", strDomain, strPDC
stbStatus.Panels(1).Text = "Scaning Network......"
If optServer.Value = False Then
    Dim nodeCount As Long
    'Get domain and PDC
    fnGetDomainName strDomain '= GetLoginDomain
    fnGetPDCName "", strDomain, strPDC
    On Error Resume Next
    'Get the users home path
    strHome = getUserHome
    SessionEnum strPDC, "", ""
    tvwOnlineUsers.Nodes.Add , tvwFirst, strPDC, strPDC
    tvwOnlineUsers.Sorted = True
    For a = 0 To UBound(msSessionInfo)
        If Trim(msSessionInfo(a).CompName) <> "" Then
            If UCase(msSessionInfo(a).UserName) = "CD9259" Then Stop
            If msSessionInfo(a).UserName = "" Then
                msSessionInfo(a).UserName = GetLoginUser(msSessionInfo(a).CompName)
            End If
            If msSessionInfo(a).UserName <> "" Then
                msSessionInfo(a).UserName = UCase(msSessionInfo(a).UserName)
                tvwOnlineUsers.Nodes.Add strPDC, tvwChild, msSessionInfo(a).UserName, getRealName(strPDC, msSessionInfo(a).UserName) & " (" & msSessionInfo(a).UserName & ")"
                tvwOnlineUsers.Nodes(msSessionInfo(a).UserName).Tag = msSessionInfo(a).CompName
                nodeCount = tvwOnlineUsers.Nodes.Count
                tvwOnlineUsers.Nodes.Add msSessionInfo(a).UserName, tvwChild, msSessionInfo(a).CompName, msSessionInfo(a).CompName
                If tvwOnlineUsers.Nodes.Count > nodeCount Then
                    tvwOnlineUsers.Nodes(tvwOnlineUsers.Nodes.Count).Tag = msSessionInfo(a).CompName
                End If
                DoEvents
            End If
        End If
    Next a
    
Else
    Timer1.Enabled = False
    SERVERTYPE = SV_TYPE_ALL
    Show
    Call FillDomainTree(SV_TYPE_DOMAIN_ENUM, Me.tvwOnlineUsers, strPDC)
    Timer1.Enabled = True
End If
Stop
stbStatus.Panels(1).Text = "Network Connections: " & tvwOnlineUsers.Nodes.Count
End Function

Public Sub ReceiveSoundBytes()
Dim iRet As Long
Dim sBuff() As Byte

   ' Process sound buffer if recording
   If (fRecording) Then
      For i = 0 To (NUM_BUFFERS - 1)
         If inHdr(i).dwFlags And WHDR_DONE Then
            rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
            If rc <> 0 Then
                MsgBox "Failed"
                Exit Sub
            End If
            ReDim sBuff(BUFFER_SIZE) As Byte
            'right hear is where the transfer should occur for sending over the
            'winsock conection
            CopyMemory ByVal sBuff(0), ByVal inHdr(i).lpData, BUFFER_SIZE
            vntFlags = SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY
            If (vntFlags And SND_MEMORY) = 0 Then
                vntFlags = vntFlags Or SND_MEMORY
            End If
            'Play the Wave
            'sndPlaySound sBuff(0), vntFlags
            'now you can send the data
            stbStatus.Panels("Recording").Text = UBound(sBuff) & "@" & Timer
            wskClient.SendData PrepareAVoice(sBuff)
            'mmioWrite hmmioIn, sBuff, BUFFER_SIZE

         End If
      Next

   End If
   
End Sub


''-----------------------------------------
''Start the capture buffer rolling
''-----------------------------------------
Sub StartCapture()
    gDSCB.Start DSCBSTART_LOOPING
End Sub
Sub StopCapture()
gDSCB.Stop
End Sub

