Attribute VB_Name = "basThread"
Public Const THREAD_BASE_PRIORITY_IDLE = -15
Public Const THREAD_BASE_PRIORITY_LOWRT = 15
Public Const THREAD_BASE_PRIORITY_MIN = -2
Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Public Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Public Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Public Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
Public Const THREAD_PRIORITY_NORMAL = 0
Public Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100
Public Declare Function SetThreadPriority Lib "KERNEL32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function SetPriorityClass Lib "KERNEL32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetThreadPriority Lib "KERNEL32" (ByVal hThread As Long) As Long
Public Declare Function GetPriorityClass Lib "KERNEL32" (ByVal hProcess As Long) As Long

Public Declare Function CreateThread Lib "KERNEL32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function TerminateThread Lib "KERNEL32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Public Declare Sub ExitThread Lib "KERNEL32" (ByVal dwExitCode As Long)
Public Declare Function GetExitCodeThread Lib "KERNEL32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public hThread As Long, hThreadID As Long, hDisplayThread As Long, hDisplayID As Long
Public lFile As Long

Public Sub AsyncThread()
frmChat.Timer1.Enabled = False
strDomain = ""
strPDC = ""
fnGetDomainName strDomain '= GetLoginDomain
fnGetPDCName "", strDomain, strPDC
frmChat.stbStatus.Panels(1).Text = "Scaning Network......"
SERVERTYPE = SV_TYPE_ALL
Call FillDomainTree(SV_TYPE_DOMAIN_ENUM, frmChat.tvwOnlineUsers, strPDC)
frmChat.stbStatus.Panels(1).Text = "Network Connections: " & frmChat.tvwOnlineUsers.Nodes.Count
hThread = 0
frmChat.Timer1.Enabled = True
End Sub
Public Sub DisplayWave()
'On Error Resume Next
Do Until hDisplayID = 0
    If (frmChat.ProgressBar1.Enabled = True) Then
        mxcd.dwControlID = outputVolCtrl.dwControlID
        mxcd.item = outputVolCtrl.cMultipleItems
        rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
        CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
        If (volume < 0) Then volume = -volume
        If volumne > frmChat.ProgressBar1.Max Then
            frmChat.ProgressBar1.Value = frmChat.ProgressBar1.Max
        Else
            frmChat.ProgressBar1.Value = volume
        End If
        'If (volume > frmChat.picPeakVol.Height) Then volume = volume / (frmChat.picPeakVol.Height / 4)
        'y = (frmChat.picPeakVol.Height / 2)
        'frmChat.picPeakVol.Line (pPos, y - volume)-(pPos, y + volume), vbRed
        'lasty = y
        'pPos = pPos + 10
        'If pPos >= frmChat.picPeakVol.Width Then
        '    frmChat.picPeakVol.Cls
        '    pPos = 0
        'End If
    End If
    
Loop
hDisplayThread = 0
End Sub
