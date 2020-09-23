VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSoundRec 
   Caption         =   "WAV Recorder / caesardutta@xoommail.com V2.01"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   2400
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save Wave File"
      Filter          =   "*.WAV"
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save As..."
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   3615
      Begin VB.CheckBox chkStereo 
         Caption         =   "Stereo / Mono"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.HScrollBar scrLevel 
      Height          =   255
      LargeChange     =   300
      Left            =   3720
      SmallChange     =   300
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton opt11 
         Caption         =   "11025 Hz"
         Height          =   195
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opt225 
         Caption         =   "22050 Hz"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opt44 
         Caption         =   "44100 Hz"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Start Recording"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Min            Level            Max"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmSoundRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TS
    sfld As String * 255
End Type


Dim volCtrl As MIXERCONTROL ' waveout volume control
Dim micCtrl As MIXERCONTROL ' microphone volume control
Dim rc As Long              ' return code
Dim ok As Boolean           ' boolean return code
Dim sFile As String
Dim blFileSaved As Boolean

Private Sub Check1_Click()
Dim sFmt As String * 255
Dim iRet As Integer

   If (Check1.Value = 1) Then
   
   If (Not blFileSaved) And (Trim(sFile) <> "") Then
        
        iRet = MsgBox("File Not Saved. Proceed?", vbYesNoCancel, "WAV Recorder")
        If iRet = vbCancel Then
            'Do nothing
            Exit Sub
        ElseIf iRet = vbNo Then
            cmdSave.Value = True
        Else
            Kill sFile
        End If
        
   End If
   
    blFileSaved = False
    cmdSave.Enabled = False
    Check1.Caption = "Stop Recording"
    'Create the Wave File
    sFile = App.Path & "\" & Format(Now, "MMDDYYYYHH24MISSAMPM") & ".WAV"
    hmmioIn = mmioOpen(sFile, mmioinf, (MMIO_CREATE Or MMIO_WRITE))  'Or MMIO_ALLOCBUF
    If hmmioIn = 0 Then
      MsgBox "Failed to create WAV file"
      Exit Sub
    End If
    
    'Set The WAV wformat
    wformat.wFormatTag = 1
    If chkStereo.Value = 1 Then
        wformat.nChannels = 2
    Else
        wformat.nChannels = 1
    End If
    
    wformat.wBitsPerSample = 16
    If opt44.Value = True Then
        wformat.nSamplesPerSec = 44100
    ElseIf opt11.Value = True Then
        wformat.nSamplesPerSec = 11025
    ElseIf opt225.Value = True Then
        wformat.nSamplesPerSec = 22050
    End If
    wformat.nBlockAlign = wformat.nChannels * wformat.wBitsPerSample / 8
    wformat.nAvgBytesPerSec = wformat.nSamplesPerSec * wformat.nBlockAlign
    wformat.cbSize = Len(wformat)
    
    'Create the RIFF Chunk
    mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
    If (mmioCreateChunk(hmmioIn, mmckinfoParentIn, MMIO_CREATERIFF) <> 0) Then
        MsgBox "Failed To Create RIFF CHUNK"
        mmioClose hmmioIn, 0
        Exit Sub
    End If
    
    'Create the Fmt Chunk
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("fmt", 0)
    mmckinfoSubchunkIn.ckSize = Len(wformat)
    If (mmioCreateChunk(hmmioIn, mmckinfoSubchunkIn, 0) <> 0) Then
        MsgBox "Failed To Create FMT CHUNK"
        mmioClose hmmioIn, 0
        Exit Sub
    End If
    
    CopyStringFromStruct sFmt, wformat, Len(wformat)
    
    If (mmioWrite(hmmioIn, sFmt, Len(wformat)) <> Len(wformat)) Then
        MsgBox "Could Not Write wformat"
        mmioClose hmmioIn, 0
        Exit Sub
    End If
    
    'Create Data Chunk
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
    If (mmioCreateChunk(hmmioIn, mmckinfoSubchunkIn, 0) <> 0) Then
        MsgBox "Failed To Create DATA CHUNK"
        mmioClose hmmioIn, 0
        Exit Sub
    End If
    Frame1.Enabled = False
    StartInput  ' Start receiving audio input

   Else
      Frame1.Enabled = True
      StopInput   ' Stop receiving audio input
      Check1.Caption = "Start Recording"
      cmdSave.Enabled = True
   End If
   
End Sub

Private Sub cmdSave_Click()
Dim o_file As New Scripting.FileSystemObject
    On Error Resume Next
    cmdlg.ShowSave
    If cmdlg.FileName <> "" Then
        o_file.CopyFile sFile, cmdlg.FileName & ".Wav", False
        Kill sFile
        blFileSaved = True
    End If
Set o_file = Nothing
End Sub

Private Sub Form_Load()
    Initialize Me.hwnd
    chkStereo.Value = 1
    opt44.Value = True
    cmdSave.Enabled = False
    scrLevel.Value = 32767
    
    ' Open the mixer with deviceID 0.
    rc = mixerOpen(hMixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Couldn't open the mixer."
        Exit Sub
    End If
    
     'Get the wavein volume control
    ok = GetVolumeControl(hMixer, _
                     MIXERLINE_COMPONENTTYPE_DST_WAVEIN, _
                     MIXERCONTROL_CONTROLTYPE_VOLUME, _
                     volCtrl)
    If (ok = True) Then
        'Adjust the scroll bar
    End If

End Sub

Private Sub scrLevel_Change()
    SetVolumeControl hMixer, volCtrl, (scrLevel.Value * 1.8)
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim iRet As Integer

   If (fRecording = True) Then
       StopInput
   End If
   
   If (Not blFileSaved) And (Trim(sFile) <> "") Then
        
        iRet = MsgBox("Quit Without Saving ?", vbYesNoCancel, "WAV Recorder")
        If iRet = vbCancel Then
            Cancel = 1
        ElseIf iRet = vbNo Then
            cmdSave.Value = True
        Else
            Kill sFile
        End If
        
   End If
   
End Sub

Public Sub Gandu()
Dim iRet As Long
Dim sBuff  As String

   ' Process sound buffer if recording
   If (fRecording) And Check1.Value = 1 Then
      For i = 0 To (NUM_BUFFERS - 1)
         If inHdr(i).dwFlags And WHDR_DONE Then
            rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
            If rc <> 0 Then
                MsgBox "Failed"
                Exit Sub
            End If

            sBuff = Space(BUFFER_SIZE)
            'right hear is where the transfer should occur for sending over the
            'winsock conection
            CopyMemory ByVal sBuff, ByVal inHdr(i).lpData, BUFFER_SIZE

            mmioWrite hmmioIn, sBuff, BUFFER_SIZE

         End If
      Next
   End If
   
End Sub
