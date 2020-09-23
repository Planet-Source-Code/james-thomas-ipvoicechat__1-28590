Attribute VB_Name = "netUtils"

Public Const SV_TYPE_WORKSTATION = &H1
Public Const SV_TYPE_SERVER = &H2
Public Const SV_TYPE_SQLSERVER = &H4
Public Const SV_TYPE_DOMAIN_CTRL = &H8
Public Const SV_TYPE_DOMAIN_BAKCTRL = &H10
Public Const SV_TYPE_TIMESOURCE = &H20
Public Const SV_TYPE_AFP = &H40
Public Const SV_TYPE_NOVELL = &H80
Public Const SV_TYPE_DOMAIN_MEMBER = &H100
Public Const SV_TYPE_LOCAL_LIST_ONLY = &H40000000
Public Const SV_TYPE_PRINT = &H200
Public Const SV_TYPE_DIALIN = &H400
Public Const SV_TYPE_XENIX_SERVER = &H800
Public Const SV_TYPE_MFPN = &H4000
Public Const SV_TYPE_NT = &H1000
Public Const SV_TYPE_WFW = &H2000
Public Const SV_TYPE_SERVER_NT = &H8000
Public Const SV_TYPE_POTENTIAL_BROWSER = &H10000
Public Const SV_TYPE_BACKUP_BROWSER = &H20000
Public Const SV_TYPE_MASTER_BROWSER = &H40000
Public Const SV_TYPE_DOMAIN_MASTER = &H80000
Public Const SV_TYPE_DOMAIN_ENUM = &H80000000
Public Const SV_TYPE_WINDOWS = &H400000
Public Const SV_TYPE_ALL = &HFFFFFFFF

Public SERVERTYPE  As Long

Public Const RESOURCE_CONNECTED As Long = &H1&
Public Const RESOURCE_GLOBALNET As Long = &H2&
Public Const RESOURCE_REMEMBERED As Long = &H3&
Public Const RESOURCEDISPLAYTYPE_DIRECTORY& = &H9
Public Const RESOURCEDISPLAYTYPE_DOMAIN& = &H1
Public Const RESOURCEDISPLAYTYPE_FILE& = &H4
Public Const RESOURCEDISPLAYTYPE_GENERIC& = &H0
Public Const RESOURCEDISPLAYTYPE_GROUP& = &H5
Public Const RESOURCEDISPLAYTYPE_NETWORK& = &H6
Public Const RESOURCEDISPLAYTYPE_ROOT& = &H7
Public Const RESOURCEDISPLAYTYPE_SERVER& = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE& = &H3
Public Const RESOURCEDISPLAYTYPE_SHAREADMIN& = &H8
Public Const RESOURCETYPE_ANY As Long = &H0&
Public Const RESOURCETYPE_DISK As Long = &H1&
Public Const RESOURCETYPE_PRINT As Long = &H2&
Public Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&
Public Const RESOURCEUSAGE_ALL As Long = &H0&
Public Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
Public Const RESOURCEUSAGE_CONTAINER As Long = &H2&
Public Const RESOURCEUSAGE_RESERVED As Long = &H80000000
Public Const NO_ERROR = 0
Public Const ERROR_MORE_DATA = 234                        'L    // dderror
Public Const RESOURCE_ENUM_ALL As Long = &HFFFF

Public Const FILTER_TEMP_DUPLICATE_ACCOUNT As Long = &H1&
Public Const FILTER_NORMAL_ACCOUNT As Long = &H2&
Public Const FILTER_PROXY_ACCOUNT As Long = &H4&
Public Const FILTER_INTERDOMAIN_TRUST_ACCOUNT As Long = &H8&
Public Const FILTER_WORKSTATION_TRUST_ACCOUNT As Long = &H10&
Public Const FILTER_SERVER_TRUST_ACCOUNT As Long = &H20&
Public Const NERR_Success As Long = 0&

Public Type USER_INFO
    Name As String
    Comment As String
    UserComment As String
    FullName As String
End Type

Public Type USER_INFO_API
    Name As Long
    Comment As Long
    UserComment As Long
    FullName As Long
End Type

Public Type USER_INFO_2
    UserName As String
    logondomain As String
    LogonServer As String
    OtherDomains As String
End Type

Public UserInfo(0 To 1000) As USER_INFO



Type SERVER_INFO_100
    sv100_platform_id As Long
    sv100_servername As Long
End Type

Public Type SERVER_INFO_101
    dw_platform_id As Long
    ptr_name As Long
    dw_ver_major As Long
    dw_ver_minor As Long
    dw_type As Long
    ptr_comment As Long
End Type


Public Type NetResource
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    pLocalName As Long
    pRemoteName As Long
    pComment As Long
    pProvider As Long
End Type
Public Type NETRESOURCE_REAL
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    sLocalName As String
    sRemoteName As String
    sComment As String
    sProvider As String
End Type
Public Declare Function NetApiBufferSize Lib "netapi32.dll" _
(lpBuffer As Any, lpLength As Long) As Long
Public Declare Function NetServerEnum Lib "netapi32.dll" (vServername As Any, _
    ByVal lLevel As Long, vBufptr As Any, lPrefmaxlen As Long, _
    lEntriesRead As Long, lTotalEntries As Long, vServerType As Any, _
    ByVal sDomain As String, vResumeHandle As Any) As Long
Public Declare Function NetWkstaUserEnum Lib "netapi32.dll" _
(ByVal strServerName As String, ByVal dwLevel As Long, _
lpBuffer As Long, ByVal dwPrefMaxLen As Long, _
lpdEntriesRead As Long, lpdTotalEntries As Long, _
lpdResumehandle As Long) As Long

Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NetResource, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Public Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As NetResource, lpBufferSize As Long) As Long
Public Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Public Declare Function VarPtrAny Lib "vb40032.dll" Alias "VarPtr" (lpObject As Any) As Long
Public Declare Sub CopyMem Lib "KERNEL32" Alias "RtlMoveMemory" (lpTo As Any, lpFrom As Any, ByVal lLen As Long)
Public Declare Sub CopyMemByPtr Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal lpTo As Long, ByVal lpFrom As Long, ByVal lLen As Long)
Public Declare Function lStrCpy Lib "KERNEL32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Public Declare Function lstrlen Lib "KERNEL32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Function getusername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Declares

Public Declare Function WNetGetUser Lib "Mpr" Alias "WNetGetUserA" (lpName As Any, ByVal lpUserName$, lpnLength&) As Long
Public Declare Function NetSessionEnum Lib "netapi32.dll" (ServerName As Byte, UncClientName As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long, ByVal PreMaxLen As Long, EntriesRead As Long, TotalEntries As Long, Resume_Handle As Long) As Long
Public Declare Function lstrlenW Lib "KERNEL32" (ByVal lpString As Long) As Long
Public Declare Function StrLen Lib "KERNEL32" Alias "lstrlenW" (ByVal ptr As Long) As Long
Public Declare Function NetWkstaGetInfo Lib "Netapi32" (strServer As Any, ByVal lLevel&, pbBuffer As Any) As Long
Public Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

Public Declare Function PtrToStr Lib "KERNEL32" Alias "lstrcpyW" (RetVal As Byte, ByVal ptr As Long) As Long
Public Declare Function NetGetDCName Lib "netapi32.dll" (ServerName As Byte, DomainName As Byte, DCNPtr As Long) As Long
Public Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (ByVal ptr As Long) As Long
Public Declare Function lstrcpyW Lib "kernel32.dll" (bRet As Byte, ByVal lPtr As Long) As Long
Public Declare Function NetUserEnum Lib "netapi32.dll" (ServerName As Byte, ByVal Level As Long, ByVal Filter As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHwnd As Long) As Long
Public Declare Function NetUserGetInfo Lib "netapi32.dll" (ServerName As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'User Types
Public Type WKSTA_INFO_100
    dw_platform_id As Long
    ptr_computername As Long
    ptr_langroup As Long
    dw_ver_major As Long
    dw_ver_minor As Long
End Type

Public Type Session_Info_10
   sesi10_cname                       As Long
   sesi10_username                    As Long
   sesi10_time                        As Long
   sesi10_idle_time                   As Long
End Type

Public Type USER_INFO_10_API
  Name As Long
  Comment As Long
  UsrComment As Long
  FullName As Long
End Type

Public Type USERINFO_2_API
  usri2_name As Long
  usri2_password As Long
  usri2_password_age As Long
  usri2_priv As Long
  usri2_home_dir As Long
  usri2_comment As Long
  usri2_flags As Long
  usri2_script_path As Long
  usri2_auth_flags As Long
  usri2_full_name As Long
  usri2_usr_comment As Long
  usri2_parms As Long
  usri2_workstations As Long
  usri2_last_logon As Long
  usri2_last_logoff As Long
  usri2_acct_expires As Long
  usri2_max_storage As Long
  usri2_units_per_week As Long
  usri2_logon_hours As Long
  usri2_bad_pw_count As Long
  usri2_num_logons As Long
  usri2_logon_server As Long
  usri2_country_code As Long
  usri2_code_page As Long
End Type

Public Type UDT_Session_Info
    CompName                   As String
    UserName                   As String
    Time                       As Long
    IdleTime                   As Long
End Type

Public Type UDT_User_Info
    Name As String
    Comment As String
    UsrComment As String
    FullName As String
End Type

Public Type WKSTA_INFO_101
    wki101_platform_id As Long
    wki101_computername As Long
    wki101_langroup As Long
    wki101_ver_major As Long
    wki101_ver_minor As Long
    wki101_lanroot As Long
End Type
 
Public Type WKSTA_USER_INFO_1
    wkui1_username As Long
    wkui1_logon_domain As Long
    wkui1_logon_server As Long
    wkui1_oth_domains As Long
End Type

Public iState As OnlineState 'local copy
Public Enum OnlineState
    Scanning = 1
    Complete = 2
End Enum


'Constants
Public Const WKSTA_LEVEL_100 = 100
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const NV_MYEVENT As Long = &H5000&
Private Const BM_CLICK = &HF5
Const WM_CLOSE = &H10

'Public variables
Public msSessionInfo() As UDT_Session_Info
Public msUserInfo As UDT_User_Info
Public strPDC As String
Public strAddUser As String
Public strDomain As String
Public strHome As String

Public lInstance As Long
Public Function fnGetDomainName(strDomain) As String
On Error Resume Next
Dim lngReturn As Long
Dim lngTemp As Long
Dim strTemp As String
Dim bDomain(99) As Byte
Dim bServer() As Byte
Dim lngBuffPtr As Long
Dim typeWorkstation As WKSTA_INFO_100

    fnGetDomainName = 0
    
    bServer = "" + vbNullChar
    
    lngReturn = NetWkstaGetInfo( _
        bServer(0), _
        WKSTA_LEVEL_100, _
        lngBuffPtr)
        
    If lngReturn <> 0 Then
        fnGetDomainName = lngReturn
        Exit Function
    End If
        
    CopyMem typeWorkstation, _
        ByVal lngBuffPtr, _
        Len(typeWorkstation)
        
    lngTemp = typeWorkstation.ptr_langroup
    
    lngReturn = PtrToStr( _
        bDomain(0), _
        lngTemp)
        
    strTemp = Left( _
        bDomain, _
        StrLen(lngTemp))

    strDomain = strTemp
    
End Function

Public Function fnGetPDCName(strServer As String, strDomain As String, strPDCName As String) As Long
On Error Resume Next
Dim lngReturn As Long
Dim lngDCNPtr As Long
Dim bDomain() As Byte
Dim bServer() As Byte
Dim bPDCName(100) As Byte

    fnGetPDCName = 0
    
    bServer = strServer & vbNullChar
    bDomain = strDomain & vbNullChar
    lngReturn = NetGetDCName( _
        bServer(0), _
        bDomain(0), _
        lngDCNPtr)
    
    If lngReturn <> 0 Then
        fnGetPDCName = lngReturn
        Exit Function
    End If
    
    lngReturn = PtrToStr(bPDCName(0), lngDCNPtr)
    lngReturn = NetAPIBufferFree(lngDCNPtr)
    strPDCName = bPDCName()
    strPDCName = Mid$(strPDCName, 1, InStr(strPDCName, Chr$(0)) - 1)
End Function

Public Function GetPrimaryDCName(ByVal DName As String) As String
On Error Resume Next
    Dim DCName As String, DCNPtr As Long
    Dim DNArray() As Byte, DCNArray(100) As Byte
    Dim Result As Long
    DNArray = DName & vbNullChar
    ' Lookup the Primary Domain Controller
    Result = NetGetDCName(0&, DNArray(0), DCNPtr)
    
    If Result <> 0 Then
      err.Raise vbObjectError + 4000, "CNetworkInfo", Result
      Exit Function
    End If
     
    lstrcpyW DCNArray(0), DCNPtr
    Result = NetAPIBufferFree(DCNPtr)
    DCName = DCNArray()
     
    GetPrimaryDCName = Left(DCName, InStr(DCName, Chr(0)) - 1)
End Function

Public Function userExists(strServer As String, strUsername As String) As Boolean
On Error Resume Next
    Dim UserInfo As USER_INFO_10_API
    Dim lngReturn As Long
    Dim baServerName() As Byte
    Dim baUserName() As Byte
    Dim lngptrUserInfo As Long
    
    'set variables
    baServerName = strServer & Chr$(0)
    baUserName = strUsername & Chr$(0)
    
    'get user info
    lngReturn = NetUserGetInfo(baServerName(0), baUserName(0), 10, lngptrUserInfo)

    'any errors?
    If lngReturn <> 0 Then
        userExists = False
    Else
        userExists = True
    End If
    
    'Free the mem
    NetAPIBufferFree lngptrUserInfo
End Function

Public Function localUserName() As String
    Dim strUsername As String * 255
    Dim lngLength As Long
    Dim lngResult As Long
    
    lngLength = 255
    lngResult = getusername(strUsername, lngLength)
    If lngResult <> 1 Then
        MsgBox "An error occurred with localUserName() - No " & Str(lngResult), vbCritical, "Error in getUserName"
        Exit Function
    End If
    localUserName = Left(strUsername, lngLength)
End Function

Function SessionEnum(sServerName As String, sClientName As String, sUserName As String)
On Error Resume Next
   Dim bFirstTime           As Boolean
   Dim lRtn                 As Long
   Dim ServerName()         As Byte
   Dim UncClientName()      As Byte
   Dim UserName()           As Byte
   Dim lptrBuffer           As Long
   Dim lEntriesRead         As Long
   Dim lTotalEntries        As Long
   Dim lResume              As Long
   Dim i                    As Integer
   Dim psComputerName               As String
   Dim psUserName                   As String
   Dim plActiveTime                 As Long
   Dim plIdleTime                   As Long
   Dim typSessionInfo()             As Session_Info_10
    
    lPrefmaxlen = 65535
     
    ServerName = sServerName & vbNullChar
    UncClientName = sClientName & vbNullChar
    UserName = sUserName & vbNullChar
    
Do
   lRtn = NetSessionEnum(ServerName(0), UncClientName(0), UserName(0), 10, lptrBuffer, lPrefmaxlen, lEntriesRead, lTotalEntries, lResume)
     
    If lRtn <> 0 Then
        SessionEnum = lRtn
        Exit Function
    End If

If lTotalEntries <> 0 Then


    ReDim typSessionInfo(0 To lEntriesRead - 1)
    ReDim msSessionInfo(0 To lEntriesRead - 1)
     
    CopyMem typSessionInfo(0), ByVal lptrBuffer, Len(typSessionInfo(0)) * lEntriesRead
     
    For i = 0 To lEntriesRead - 1
     
        msSessionInfo(i).CompName = PointerToStringW(typSessionInfo(i).sesi10_cname)
        msSessionInfo(i).UserName = PointerToStringW(typSessionInfo(i).sesi10_username)
        msSessionInfo(i).Time = typSessionInfo(i).sesi10_time
        msSessionInfo(i).IdleTime = typSessionInfo(i).sesi10_idle_time
    Next i
    End If
Loop Until lEntriesRead = lTotalEntries
   
    If lptrBuffer <> 0 Then
        NetAPIBufferFree lptrBuffer
    End If
     
End Function

Public Function PointerToStringW(lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
    
   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMem Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function

Public Function getRealName(strServer As String, strUsername As String) As String
On Error Resume Next
    Dim lngReturn As Long
    Dim baServerName() As Byte
    Dim baUserName() As Byte
    Dim lngptrUserInfo As Long
    Dim UserInfo As USER_INFO_10_API
    Dim strName As String
    Dim a As Integer
    
    'set variables
    baServerName = strServer & Chr$(0)
    baUserName = strUsername & Chr$(0)
    
    'get user info
    DoEvents
    lngReturn = NetUserGetInfo(baServerName(0), baUserName(0), 10, lngptrUserInfo)
    'any errors?
    If lngReturn <> 0 Then
      getRealName = ""
      Exit Function
    End If
    'Turn the pointer into a variable
    CopyMem UserInfo, ByVal lngptrUserInfo, Len(UserInfo)
    strName = PointerToStringW(UserInfo.FullName)
    NetAPIBufferFree lngptrUserInfo
    getRealName = strName
End Function

Public Function computername() As String
Dim strString As String
    'Create a buffer
    strString = String(255, Chr$(0))
    'Get the computer name
    GetComputerName strString, 255
    'remove the unnecessary chr$(0)'s
    strString = Left$(strString, InStr(1, strString, Chr$(0)))
    'Show the computer name
    computername = strString
End Function
Public Function GetLoginDomain() As String
'---------------------------------------------------------------------------
' Found Domain/Workgroup and Logon Domain
'---------------------------------------------------------------------------
Dim ret As Long, Buffer(512) As Byte, i As Integer
Dim wk101 As WKSTA_INFO_101, pwk101 As Long
Dim wk1 As WKSTA_USER_INFO_1, pwk1 As Long
Dim cbusername As Long, UserName As String
Dim computername As String, langroup As String, logondomain As String
Dim X As String
            

' Clear all of the display values.
computername = "": langroup = "": UserName = "": logondomain = ""
' Windows 95 or NT - call WNetGetUser to get the name of the user.
UserName = Space(256)
cbusername = Len(UserName)
ret = WNetGetUser(ByVal 0&, UserName, cbusername)
If ret = 0 Then
    ' Success - strip off the null.
    UserName = Left(UserName, InStr(UserName, Chr(0)) - 1)
Else
    UserName = ""
End If

'================================================================
' The following section works only under Windows NT
'================================================================
'NT only - call NetWkstaGetInfo to get computer name and lan group
ret = NetWkstaGetInfo(ByVal 0&, 101, pwk101)
RtlMoveMemory wk101, ByVal pwk101, Len(wk101)
lstrcpyW Buffer(0), wk101.wki101_computername
' Get every other byte from Unicode string.
i = 0
Do While Buffer(i) <> 0
    computername = computername & Chr(Buffer(i))
    i = i + 2
Loop
lstrcpyW Buffer(0), wk101.wki101_langroup
i = 0
Do While Buffer(i) <> 0
langroup = langroup & Chr(Buffer(i))
    i = i + 2
Loop
ret = NetAPIBufferFree(pwk101)

GetLoginDomain = langroup
End Function



Public Sub TypeOfServer(Combo1 As ComboBox)

  If (Combo1 = "LAN Manager Workstations") Or (Combo1 = "(Nothing Specific)") Then
    SERVERTYPE = SV_TYPE_WORKSTATION
End If

If Combo1 = "LAN Manager Servers" Then
   SERVERTYPE = SV_TYPE_SERVER
End If

If Combo1 = "SQL Servers" Then
    SERVERTYPE = SV_TYPE_SQLSERVER
End If

If Combo1 = "Primary Domain Controllers" Then
    SERVERTYPE = SV_TYPE_DOMAIN_CTRL
End If

If Combo1 = "Backup Domain Controllers" Then
    SERVERTYPE = SV_TYPE_DOMAIN_BAKCTRL
End If

If Combo1 = "Timesource Servers" Then
    SERVERTYPE = SV_TYPE_TIMESOURCE
End If

If Combo1 = "Apple File Protocol Servers" Then
    SERVERTYPE = SV_TYPE_AFP
End If

If Combo1 = "Novell Servers" Then
    SERVERTYPE = SV_TYPE_NOVELL
End If

If Combo1 = "LM 2.x Domain Members" Then
    SERVERTYPE = SV_TYPE_DOMAIN_MEMBER
End If

If Combo1 = "Local Browse List (MB Only)" Then
    SERVERTYPE = SV_TYPE_LOCAL_LIST_ONLY
End If

If Combo1 = "Print Servers" Then
    SERVERTYPE = SV_TYPE_PRINT
End If

If Combo1 = "Dial-in Servers" Then
    SERVERTYPE = SV_TYPE_DIALIN
End If

If Combo1 = "Xenix Servers" Then
   SERVERTYPE = SV_TYPE_XENIX_SERVER
End If

If Combo1 = "MS Novell File & Print Servers" Then
    SERVERTYPE = SV_TYPE_MFPN
End If

If Combo1 = "Windows NT (S&W)" Then
    SERVERTYPE = SV_TYPE_NT
End If

If Combo1 = "WfW Servers" Then
    SERVERTYPE = SV_TYPE_WFW
End If

If Combo1 = "Non-DC NT Servers" Then
    SERVERTYPE = SV_TYPE_SERVER_NT
End If

If Combo1 = "Potential Master Browsers" Then
    SERVERTYPE = SV_TYPE_POTENTIAL_BROWSER
End If

If Combo1 = "Backup Master Browsers" Then
    SERVERTYPE = SV_TYPE_BACKUP_BROWSER
End If

If Combo1 = "Master Browser Servers" Then
    SERVERTYPE = SV_TYPE_MASTER_BROWSER
End If

If Combo1 = "Domain Master Browsers" Then
    SERVERTYPE = SV_TYPE_DOMAIN_MASTER
End If

If Combo1 = "Windows 95 and Later" Then
    SERVERTYPE = SV_TYPE_WINDOWS
End If

If Combo1 = "All Server Types" Then
    SERVERTYPE = SV_TYPE_ALL
End If

End Sub

Public Function FillDomainTree(lType As Long, tvw As TreeView, strthepdc As String) As Boolean
Dim lReturn As Long
Dim Server_Info As Long
Dim lEntries As Long
Dim lTotal As Long
Dim lMax As Long
Dim vResume As Variant
Dim tServer_info_101 As SERVER_INFO_101
Dim sServer As String
Dim sDomain As String
Dim lServerInfo101StructPtr As Long
Dim X As Long, i As Long
Dim bBuffer(512) As Byte
Dim nodex As Node
tvw.Nodes.Clear
Set nodex = tvw.Nodes.Add(, , "R", "Network Domains")
nodex.Expanded = True
lReturn = NetServerEnum(ByVal 0&, 101, Server_Info, lMax, lEntries, lTotal, ByVal lType, sDomain, vResume)
If lReturn <> 0 Then
    Exit Function
End If
X = 1
lServerInfo101StructPtr = Server_Info
Do While X <= lTotal
    RtlMoveMemory tServer_info_101, ByVal lServerInfo101StructPtr, Len(tServer_info_101)
    lstrcpyW bBuffer(0), tServer_info_101.ptr_name
    i = 0
    On Error Resume Next
    Do While bBuffer(i) <> 0
        sServer = sServer & Chr$(bBuffer(i))
        i = i + 2
    Loop
    Set nodex = tvw.Nodes.Add("R", tvwChild, sServer, sServer)
    nodex.Expanded = True
    Call AddDomainServers(SERVERTYPE, tvw, sServer, strthepdc)
    DoEvents
    X = X + 1
    sServer = ""
    lServerInfo101StructPtr = lServerInfo101StructPtr + Len(tServer_info_101)
Loop
lReturn = NetAPIBufferFree(Server_Info)
End Function
Public Sub AddDomainServers(lType As Long, tvw As TreeView, Parentkey As String, strthepdc As String)
Dim lReturn As Long
Dim Server_Info As Long
Dim lEntries As Long
Dim lTotal As Long
Dim lMax As Long
Dim vResume As Variant
Dim tServer_info_101 As SERVER_INFO_101
Dim sServer As String
Dim str_User As String
Dim str_Realname As String
Dim sDomain As String
Dim lServerInfo101StructPtr As Long
Dim X As Long, i As Long
Dim bBuffer(512) As Byte
Dim strBuffer As String
Dim nodex As Node
sDomain = StrConv(Parentkey, vbUnicode)
lReturn = NetServerEnum(ByVal 0&, 101, Server_Info, lMax, lEntries, lTotal, ByVal lType, sDomain, vResume)
If lReturn <> 0 Then
    Exit Sub
End If
X = 1
lServerInfo101StructPtr = Server_Info
Do While X <= lTotal
    RtlMoveMemory tServer_info_101, ByVal lServerInfo101StructPtr, Len(tServer_info_101)
    lstrcpyW bBuffer(0), tServer_info_101.ptr_name
    'call CopyMemoryAny( VarPtr(strBuffer), tServer_info_101.ptr_name,
    i = 0
    'On Error GoTo fin
    Do While bBuffer(i) <> 0
        sServer = sServer & Chr$(bBuffer(i))
        i = i + 2
    Loop
    'Here you can make this dependent of the OS
    str_User = UCase(GetLoginUser(sServer))
    str_Realname = getRealName(strPDC, str_User)

    If str_User <> "" Then
        Set nodex = tvw.Nodes.Add(Parentkey, tvwChild, str_User, str_Realname & " (" & str_User & ")")
        Set nodex = tvw.Nodes.Add(str_User, tvwChild, sServer, sServer)
        nodex.Expanded = True
    End If
    X = X + 1
    sServer = ""
    lServerInfo101StructPtr = lServerInfo101StructPtr + Len(tServer_info_101)
Loop
lReturn = NetAPIBufferFree(Server_Info)
Exit Sub
fin:
Select Case err.Number
    Case 35601
        Resume Next
    Case 35602
        Resume Next
    Case Else
        MsgBox err.Number & err.Description & err.Source
        Resume Next
End Select
End Sub
Public Function GetLocalSystemName()
    Dim lReturnCode As Long
    Dim bBuffer(512) As Byte
    Dim i As Integer
    Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
    Dim lwkstaInfo100StructPtr As Long
    Dim sLocalName As String
    
    lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
 
    lwkstaInfo100StructPtr = lwkstaInfo100
                 
    If lReturnCode = 0 Then
                 
        RtlMoveMemory twkstaInfo100, ByVal _
        lwkstaInfo100StructPtr, Len(twkstaInfo100)
         
        lstrcpyW bBuffer(0), twkstaInfo100.ptr_computername

        i = 0
        Do While bBuffer(i) <> 0
            sLocalName = sLocalName & Chr(bBuffer(i))
            i = i + 2
        Loop
            
        GetLocalSystemName = sLocalName
         
    End If

End Function

Public Function GetDomainName() As String
    
    Dim lReturnCode As Long
    Dim bBuffer(512) As Byte
    Dim i As Integer
    Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
    Dim lwkstaInfo100StructPtr As Long
    Dim sDomainName As String
    
    lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
 
    lwkstaInfo100StructPtr = lwkstaInfo100
                 
    If lReturnCode = 0 Then
                 
        RtlMoveMemory twkstaInfo100, ByVal lwkstaInfo100StructPtr, Len(twkstaInfo100)
         
        lstrcpyW bBuffer(0), twkstaInfo100.ptr_langroup
        
        
        i = 0
        Do While bBuffer(i) <> 0
            sDomainName = sDomainName & Chr(bBuffer(i))
            i = i + 2
        Loop
            
        GetDomainName = sDomainName
         
    End If
        
End Function


  Public Property Get UserName() As Variant
          Dim sBuffer As String
          Dim lSize As Long
          sBuffer = Space$(255)
          lSize = Len(sBuffer)
          Call getusername(sBuffer, lSize)
          UserName = Left$(sBuffer, lSize)
          
     End Property

Public Function GetUsers(ServerName As String) As Long
    Dim lpBuffer As Long
    Dim nRet As Long
    Dim EntriesRead As Long
    Dim TotalEntries As Long
    Dim ResumeHandle As Long
    Dim uUser As USER_INFO_API
    Dim bServer() As Byte
    Dim PDC() As Byte
    Dim i As Integer

    If Trim(ServerName) = "" Then
        'Local users
        bServer = vbNullString
    Else
        'Check the syntax of the ServerName string
        If InStr(ServerName, "\\") = 1 Then
            bServer = ServerName & vbNullChar
        Else
            bServer = "\\" & ServerName & vbNullChar
        End If
    End If
    i = 0
    ResumeHandle = 0
    Do
        'Start to enumerate the Users
        If Trim(ServerName) = "" Then
            nRet = NetUserEnum(vbNullString, 10, FILTER_NORMAL_ACCOUNT, lpBuffer, 1, EntriesRead, TotalEntries, ResumeHandle)
        Else
            nRet = NetUserEnum(bServer(0), 10, FILTER_NORMAL_ACCOUNT, lpBuffer, 1, EntriesRead, TotalEntries, ResumeHandle)
        End If
        'Fill the data structure for the User
        If nRet = ERROR_MORE_DATA Then
            CopyMem uUser, ByVal lpBuffer, Len(uUser)
            UserInfo(i).Name = PointerToStringW(uUser.Name)
            UserInfo(i).Comment = PointerToStringW(uUser.Comment)
            UserInfo(i).UserComment = PointerToStringW(uUser.UserComment)
            UserInfo(i).FullName = PointerToStringW(uUser.FullName)
            i = i + 1
        End If
        If lpBuffer Then
            Call NetAPIBufferFree(lpBuffer)
        End If
    Loop While nRet = ERROR_MORE_DATA
    'Return the number of Users
    GetUsers = i
End Function


Public Function GetLoginUser(ByVal strHostName As String) As String
    Dim lngLevel As Long
    Dim lngPrefmaxlen As Long
    Dim lngEntriesRead As Long
    Dim lngTotalEntries As Long
    Dim lngResumeHandle As Long
    Dim lngReturn As Long
    Dim lngLength As Long
    Dim lngBuffer As Long
    Dim typWkStaInfo(0 To 1000) As WKSTA_USER_INFO_1
    Dim intCount As Integer
    Dim CurrentInfo As USER_INFO_2

    'Check for the right syntax for the servername
    'Convert it to unicode because the C function wants a LPCWSTR
    'ie LongPointer to a unicode string, the C stands for a constants
    'vbNullString for the local Machine
    If strHostName = "" Then
        strHostName = vbNullString
    Else
        If InStr(strHostName, "\\") <> 0 Then
            strHostName = StrConv(strHostName & vbNullChar, vbUnicode)
        Else
            strHostName = StrConv("\\" & strHostName & vbNullChar, vbUnicode)
        End If
    End If
    'set the resumehandle to the first entry
    lngResumeHandle = 0
    'Call the function, the -1 passed to dwPrefMaxLen lets the function create its
    'own buffer that will hold all the data returned, I choose to enumerate at level 1
    'you can pass a level 0, feel free to modify
    DoEvents
    lngReturn = NetWkstaUserEnum(strHostName, &H1, lngBuffer, -1, lngEntriesRead, lngTotalEntries, lngResumeHandle)
    'if successful ie NERR_Success get the info
    DoEvents
    If lngReturn = NERR_Success Then
        'initialize the count variable
        intCount = 0
        'Get the size of the memory allocated
        lngReturn = NetApiBufferSize(ByVal lngBuffer, lngLength)
        'Copy the memory into the array so we can get the information out
        'I imagine this could cause really strange things to happen if you happen to
        'have more then 1000 users logged into this workstation. I tried to dump the info into
        'a dynamic array and VB keep generating a Doctor Watson error everytime the
        'sub exited beats me why, If anybody know email me
        CopyMem typWkStaInfo(0), ByVal lngBuffer, lngLength
        'Get the info out and add it too are collection
        For intCount = 0 To lngTotalEntries - 1
            'temporay object to hold the info
            
            'The info returned is actually a LP, which we have to convert
            'I used Andrea Tincani's function which transforms the returned LPWSTR to a string
            GetLoginUser = PointerToStringW(typWkStaInfo(intCount).wkui1_username)
            'CurrentInfo.UserName = PointerToStringW(typWkStaInfo(intCount).wkui1_username)
            'CurrentInfo.logondomain = PointerToStringW(typWkStaInfo(intCount).wkui1_logon_domain)
            'CurrentInfo.LogonServer = PointerToStringW(typWkStaInfo(intCount).wkui1_logon_server)
            'CurrentInfo.OtherDomains = PointerToStringW(typWkStaInfo(intCount).wkui1_oth_domains)
            'One more done
            intCount = intCount + 1
        Next
    Else
        'our function failed lets find out why
       Exit Function
    End If
    'We have to free up the Memory the funtion allocated for our data
    If lngBuffer Then
        Call NetAPIBufferFree(ByVal lngBuffer)
    End If

End Function
