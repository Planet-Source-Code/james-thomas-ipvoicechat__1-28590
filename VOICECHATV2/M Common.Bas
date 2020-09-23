Attribute VB_Name = "MCommon"

Option Explicit

'===============================================================================
'   Constants
'===============================================================================
Private Const Module                As String = "MCommon"

Private Const GMEM_ZEROINIT         As Long = &H40&
Private Const GMEM_FIXED            As Long = &H0&
Private Const GPTR                  As Long = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const LF_FACESIZE           As Long = 32

Private Const WM_USER               As Long = &H400&
Private Const ERROR_SUCCESS         As Long = 0
Private Const INVALID_HANDLE_VALUE  As Long = -1
Private Const LOGPIXELSX            As Long = 88
Private Const LOGPIXELSY            As Long = 90

' ListBox messages
Private Const LB_ERR                As Long = (-1)
Private Const LB_FINDSTRING         As Long = &H18F&
Private Const LB_FINDSTRINGEXACT    As Long = &H1A2&
Private Const LB_SETCURSEL          As Long = &H186&
Private Const LB_GETTEXT            As Long = &H189&
Private Const LB_GETTEXTLEN         As Long = &H18A&

' ComboBox messages
Private Const CB_ERR                As Long = (-1)
Private Const CB_FINDSTRING         As Long = &H14C&
Private Const CB_FINDSTRINGEXACT    As Long = &H158&
Private Const CB_SETCURSEL          As Long = &H14E&
Private Const CB_GETLBTEXT          As Long = &H148&
Private Const CB_GETLBTEXTLEN       As Long = &H149&
Private Const CB_SHOWDROPDOWN       As Long = &H14F&
Private Const CB_GETDROPPEDSTATE    As Long = &H157&
Private Const CB_SETEXTENDEDUI      As Long = &H155&
Private Const CB_GETDROPPEDWIDTH    As Long = &H15F&
Private Const CB_SETDROPPEDWIDTH    As Long = &H160&
Private Const MAX_COMPUTERNAME_LENGTH As Long = 15&

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000&
Private Const CLR_INVALID           As Long = &HFFFFFFFF
' Used by ShellWait
Private Const SYNCHRONIZE                 As Long = &H100000
Private Const PROCESS_QUERY_INFORMATION   As Long = &H400&
Private Const STILL_ACTIVE                As Long = &H103&
Private Const ERROR_ACCESS_DENIED         As Long = 5
'Private Const INVALID_HANDLE_VALUE        As Long = -1

'===============================================================================
'   Enums
'===============================================================================

Public Enum ProperCaseRules         ' Used by Proper case
    pcDelimitSpaces = 0
    pcDelimitNonAlpha = 1
End Enum

Public Enum EncodeFormats           ' Used by EncodeData
    efHexBytes = 0
    efHexWORDS = 1
    efHexDWORDs = 2
    efHexQWORDs = 3
End Enum

Public Enum ShellVirtualFolders     ' ShellFolder Browsing
    CSIDL__NA = -1                  ' SPI Custom value indicating that we are NOT browsing a virtual folder
    CSIDL_DESKTOP = &H0             ' Windows desktop - virtual folder at the root of the name space.
    CSIDL_INTERNET = &H1            ' Virtual folder representing the Internet
    CSIDL_PROGRAMS = &H2            ' File system directory that contains the user's program groups (which are also file system directories).
    CSIDL_CONTROLS = &H3            ' Control Panel - virtual folder containing icons for the control panel applications.
    CSIDL_PRINTERS = &H4            ' Printers folder - virtual folder containing installed printers.
    CSIDL_PERSONAL = &H5            ' File system directory that serves as a common respository for documents.
    CSIDL_FAVORITES = &H6           ' Favorite Places - file system folder that stores web (or other) links
    CSIDL_STARTUP = &H7             ' File system directory that corresponds to the user's Startup program group.
    CSIDL_RECENT = &H8              ' File system directory that contains the user's most recently used documents.
    CSIDL_SENDTO = &H9              ' File system directory that contains Send To menu items.
    CSIDL_BITBUCKET = &HA           ' Recycle bin - file system directory containing file objects in the user's recycle bin.
                                    ' The location of this directory is not in the registry; it is marked with the hidden
                                    ' and system attributes to prevent the user from moving or deleting it.
    CSIDL_STARTMENU = &HB           ' File system directory containing Start menu items.
    CSIDL_DESKTOPDIRECTORY = &H10   ' File system directory used to physically store file objects on the desktop (not to be confused
                                    ' with the desktop folder itself).
    CSIDL_DRIVES = &H11             ' My Computer - virtual folder containing everything on the local computer:
                                    ' storage devices, printers, and Control Panel. The folder may also contain mapped network drives.
    CSIDL_NETWORK = &H12            ' Network Neighborhood - virtual folder representing the top level of the network hierarchy.
    CSIDL_NETHOOD = &H13            ' File system directory containing objects that appear in the network neighborhood.
    CSIDL_FONTS = &H14              ' Virtual folder containing fonts.
    CSIDL_TEMPLATES = &H15          ' File system directory that serves as a common repository for document templates.
    CSIDL_COMMON_STARTMENU = &H16   '
    CSIDL_COMMON_PROGRAMS = &H17    '
    CSIDL_COMMON_STARTUP = &H18     '
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19  '
    CSIDL_APPDATA = &H1A            '
    CSIDL_PRINTHOOD = &H1B          '
    CSIDL_ALTSTARTUP = &H1D         ' // DBCS
    CSIDL_COMMON_ALTSTARTUP = &H1E  ' // DBCS
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
End Enum

Private Enum OSVersionEnum
    VER_PLATFORM_WIN32s = 0
    VER_PLATFORM_WIN32_WINDOWS = 1
    VER_PLATFORM_WIN32_NT = 2
End Enum

Public Enum VARENUM
    VT_EMPTY = 0
    VT_NULL = 1
    VT_I2 = 2
    VT_I4 = 3
    VT_R4 = 4
    VT_R8 = 5
    VT_CY = 6
    VT_DATE = 7
    VT_BSTR = 8
    VT_DISPATCH = 9
    VT_ERROR = 10
    VT_BOOL = 11
    VT_VARIANT = 12
    VT_UNKNOWN = 13
    VT_DECIMAL = 14
    VT_I1 = 16
    VT_UI1 = 17
    VT_UI2 = 18
    VT_UI4 = 19
    VT_I8 = 20
    VT_UI8 = 21
    VT_INT = 22
    VT_UINT = 23
    VT_VOID = 24
    VT_HRESULT = 25
    VT_PTR = 26
    VT_SAFEARRAY = 27
    VT_CARRAY = 28
    VT_USERDEFINED = 29
    VT_LPSTR = 30
    VT_LPWSTR = 31
    VT_FILETIME = 64
    VT_BLOB = 65
    VT_STREAM = 66
    VT_STORAGE = 67
    VT_STREAMED_OBJECT = 68
    VT_STORED_OBJECT = 69
    VT_BLOB_OBJECT = 70
    VT_CF = 71
    VT_CLSID = 72
    VT_VECTOR = &H1000&
    VT_ARRAY = &H2000&
    VT_BYREF = &H4000&
    VT_RESERVED = &H8000&
    VT_ILLEGAL = &HFFFF&
    VT_ILLEGALMASKED = &HFFF&
    VT_TYPEMASK = &HFFF&
End Enum

'/*
' * Queue status flags for GetQueueStatus() and MsgWaitForMultipleObjects()
' */
Public Enum QueueStatusFlags
    QS_KEY = &H1
    QS_MOUSEMOVE = &H2
    QS_MOUSEBUTTON = &H4
    QS_POSTMESSAGE = &H8
    QS_TIMER = &H10
    QS_PAINT = &H20
    QS_SENDMESSAGE = &H40
    QS_HOTKEY = &H80
    QS_ALLPOSTMESSAGE = &H100

    QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)

    QS_INPUT = (QS_MOUSE Or QS_KEY)

    QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)

    QS_ALLINPUT = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY Or QS_SENDMESSAGE)
End Enum

Public Enum RandomStringOptions
    rsoAllChars = 0
    rsoAllCharsExtended = 1
    rsoKeyboardChars = 2
    rsoAlphaNumericChars = 3
End Enum


'===============================================================================
'   Types
'===============================================================================

Private Type TMemRecord
    hMemObject      As Long                 ' Pointer to a memory object
    cbMemObject     As Long                 ' Size of memory object in bytes
End Type

Private Type POINTAPI
    X               As Long
    y               As Long
End Type

Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
' Optional filler:
'   Width           As Long
'   Height          As Long
End Type

Private Type GUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

Private Type PICTDESC                       ' This version is generic & supports all 4 types
    cbSizeofstruct  As Long                 ' defines the version of structure
    picType         As Long
    hImage          As Long                 ' hBitmap, hIcon, hMeta or hEnahanceMeta
    xExt_Or_hPal    As Long                 ' handle to a palette [bitmap] or width [metafile]
    yExt            As Long                 ' height [metafile]
End Type

Private Type LOGFONT
    lfHeight        As Long
    lfWidth         As Long
    lfEscapement    As Long
    lfOrientation   As Long
    lfWeight        As Long
    lfItalic        As Byte
    lfUnderline     As Byte
    lfStrikeOut     As Byte
    lfCharSet       As Byte
    lfOutPrecision  As Byte
    lfClipPrecision As Byte
    lfQuality       As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type FONTDESC
    cbSizeofstruct  As Long                 ' Size of the FONTDESC structure.
    lpstrName       As String               ' Pointer to the caller-owned string specifying the font name.
    cySize          As Long                 ' Initial point size of the font in CY units.
    sWeight         As Integer              ' Initial weight of the font. If the weight is below 550
                                            ' (the average of FW_NORMAL, 400, and FW_BOLD, 700), then
                                            ' the Bold property is also initialized to FALSE.
                                            ' If the weight is above 550, the Bold property is set to TRUE.
    sCharset        As Integer              ' Initial character set of the font.
    fItalic         As Long                 ' Initial italic state of the font.
    fUnderline      As Long                 ' Initial underline state of the font.
    fStrikethrough  As Long                 ' Initial strikethrough state of the font.
End Type

Public Type SHITEMID                        ' Shell item identifier
     cb As Long
     abID As Byte
End Type

Public Type ITEMIDLIST                      ' Shell item identifier list
     mkid As SHITEMID
End Type

Public Type BROWSEINFO                      ' shell structure used to browse for folders
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type

Public Enum BrowseFolderFlags               ' flags used to browse for folders
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4
    BIF_RETURNFSANCESTORS = &H8
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
End Enum

Private Type OSVERSIONINFO                      ' operating system version information [used by GetVersion]
    dwOSVersionInfoSize     As Long
    dwMajorVersion          As Long
    dwMinorVersion          As Long
    dwBuildNumber           As Long
    dwPlatformId            As Long
    szCSDVersion            As String * 128     ' Maintenance string for PSS usage
End Type

' The QUAD structures are used to convert between
' one single 64 bit value and two 32 bits long values.
Private Type TQuad
    Value                   As Currency
End Type

Private Type TQuadParts
    LoWords                 As Long
    HiWords                 As Long
End Type

Public Type SAFEARRAYBOUND
    cElements               As Long             ' # of elements in the array dimension
    lLbound                 As Long             ' lower bounds of the array dimension
End Type

Public Type SAFEARRAY
    cDims                   As Integer          ' // Count of dimensions in this array.
    fFeatures               As Integer          ' // Flags used by the SafeArray
                                                ' // routines documented below.
    cbElements              As Long             ' // Size of an element of the array.
                                                ' // Does not include size of
                                                ' // pointed-to data.
    cLocks                  As Long             ' // Number of times the array has been
                                                ' // locked without corresponding unlock.
    pvData                  As Long             ' // Pointer to the data.
    ' Should be sized to cDims:
    rgsabound()             As SAFEARRAYBOUND   ' // One bound for each dimension.
End Type

'===============================================================================
'   Private members
'===============================================================================
Private OSVersion           As OSVERSIONINFO    ' private structure for various Public functions
Private Success             As Boolean          ' general return code [boolean]
Private ReturnVal           As Long             ' general return code [various]
Private ErrCode             As Long             ' general return code [error]

'===============================================================================
'   Declares
'===============================================================================
Public Declare Sub SetLastError Lib "KERNEL32" (ByVal dwErrCode As Long)

Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32.dll" (ByVal Pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetPathFromIDListW Lib "shell32.dll" (ByVal Pidl As Long, ByVal pszPath As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, Pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetFileAttributesA Lib "KERNEL32" (ByVal lpFilename As String) As Long
Private Declare Function GetFileAttributesW Lib "KERNEL32" (ByVal lpFilename As Long) As Long

Private Declare Function GetTempFileNameW Lib "KERNEL32" (ByVal lpszPath As Long, ByVal lpPrefixString As Long, ByVal wUnique As Long, ByVal lpTempFileName As Long) As Long
Private Declare Function GetTempPathW Lib "KERNEL32" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function GetTempFileNameA Lib "KERNEL32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPathA Lib "KERNEL32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function WNetGetUserA Lib "mpr.dll" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Private Declare Function WNetGetUserW Lib "mpr.dll" (ByVal lpName As Long, ByVal lpUserName As Long, lpnLength As Long) As Long
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameW Lib "advapi32.dll" (ByVal lpBuffer As Long, nSize As Long) As Long

Private Declare Function GetComputerNameA Lib "KERNEL32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameW Lib "KERNEL32" (ByVal lpBuffer As Long, nSize As Long) As Long

Private Declare Function GetModuleFileNameA Lib "KERNEL32" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleFileNameW Lib "KERNEL32" (ByVal hModule As Long, ByVal lpFilename As Long, ByVal nSize As Long) As Long

Private Declare Function RasGetErrorStringA Lib "RasApi32.DLL" (ByVal uErrorValue As Long, ByVal lpszErrorString As String, ByVal cBufSize As Long) As Long
Private Declare Function RasGetErrorStringW Lib "RasApi32.DLL" (ByVal uErrorValue As Long, ByVal lpszErrorString As Long, ByVal cBufSize As Long) As Long

Public Declare Sub CopyMemoryAny Lib "KERNEL32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nBytes As Long)
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSource As Long, ByVal nBytes As Long)

Public Declare Sub ZeroMemory Lib "KERNEL32" Alias "RtlZeroMemory" (ByVal lpSource As Long, ByVal nBytes As Long)

Private Declare Function FormatMessageA Lib "KERNEL32" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function FormatMessageW Lib "KERNEL32" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Function lstrlenW Lib "KERNEL32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenA Lib "KERNEL32" (lpString As Any) As Long

Private Declare Function GetObjectApiA Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectApiW Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function OpenProcess Lib "KERNEL32" _
  (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "KERNEL32" _
  (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal pclsid As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function OleCreateFontIndirect Lib "OLEPRO32.DLL" (pFontDesc As FONTDESC, riid As GUID, ppvObj As IFont) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function LPtoDP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function DPtoLP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

' LastDllErrorMsg related functions
Private Declare Function GetModuleHandleA Lib "KERNEL32" (ByVal lpModuleName As String) As Long
Private Declare Function InternetGetLastResponseInfoA Lib "WinInet.dll" _
    (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Long
Private Declare Function WNetGetLastErrorA Lib "mpr.dll" _
    (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, _
     ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Private Declare Function WNetGetLastErrorW Lib "mpr.dll" _
    (lpError As Long, ByVal lpErrorBuf As Long, ByVal nErrorBufSize As Long, _
     ByVal lpNameBuf As Long, ByVal nNameBufSize As Long) As Long
     
Private Declare Function SafeArrayGetDim Lib "oleaut32" (psa As Long) As Long

Private Declare Function InternetOpen Lib "WinInet.dll" Alias "InternetOpenA" _
           (ByVal lpszCallerName As String, _
            ByVal dwAccessType As Long, _
            ByVal lpszProxyName As String, _
            ByVal lpszProxyBypass As String, _
            ByVal dwFlags As Long) As Long

Private Declare Function InternetCloseHandle Lib "WinInet.dll" _
           (ByVal hInternetHandle As Long) As Boolean


Private Declare Function InternetOpenUrl Lib "WinInet.dll" Alias "InternetOpenUrlA" _
           (ByVal hInternet As Long, ByVal lpszUrl As String, lpszHeaders As Any, _
            ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Private Declare Function InternetReadFile Lib "WinInet.dll" _
           (ByVal hFile As Long, lpBuffer As Any, ByVal dwNumberOfBytesToRead As Long, _
            lpdwNumberOfBytesRead As Long) As Long

Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long

Private Declare Function WaitForMultipleObjects Lib "KERNEL32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function WaitForMultipleObjectsEx Lib "KERNEL32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private Declare Function CreateEvent Lib "KERNEL32" Alias "CreateEventA" (lpEventAttributes As Any, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public MyThread As Long
Public iState As Integer
Public Function T2PX(ByVal Twips As Long) As Long
'===============================================================================
'   T2PX - Converts the supplied horizontal twips to pixels.
'   Use with nested calculations so it's easier to read.
'===============================================================================
    T2PX = Twips / TPPX
End Function
Public Function T2PY(ByVal Twips As Long) As Long
'===============================================================================
'   T2PY - Converts the supplied vertical twips to pixels.
'   Use with nested calculations so it's easier to read.
'===============================================================================
    T2PY = Twips / TPPY
End Function

Public Function P2TX(ByVal Pixels As Long) As Long
'===============================================================================
'   P2TX - Converts the supplied horizontal pixels to twips.
'   Use with nested calculations so it's easier to read.
'===============================================================================
    P2TX = Pixels * TPPX
End Function
Public Function P2TY(ByVal Pixels As Long) As Long
'===============================================================================
'   P2TY - Converts the supplied vertical pixels to twips.
'   Use with nested calculations so it's easier to read.
'===============================================================================
    P2TY = Pixels * TPPY
End Function

Public Function TPPX(Optional Multiplier As Integer = 1) As Long
'===============================================================================
'   TPPX - Returns the number of horizontal twips per pixel.
'   Use with nested calculations so it's easier to read.
'===============================================================================
    TPPX = Screen.TwipsPerPixelX * Multiplier
End Function
Public Function TPPY(Optional Multiplier As Integer = 1) As Long
'===============================================================================
'   TPPY - Returns the number of vertical twips per pixel.
'   Use with nested calculations so it's easier to read.
'===============================================================================
    TPPY = Screen.TwipsPerPixelY * Multiplier
End Function

Public Function CharCount(ByVal Expression As String, ByVal FindString As String, _
                          Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
'===============================================================================
'   CharCount - Counts the number of substring occurrances within a string.
'
'   Expression      String to evaluate.
'   FindString      String sought.
'   Compare         Binary or Textual compare.
'
'   RETURNS         The number of substring occurrances.
'===============================================================================

    Dim iPos            As Long
    
    If Compare = vbBinaryCompare Then
        Expression = UCase$(Expression)
        FindString = UCase$(FindString)
    End If
    
    iPos = InStr(Expression, FindString)
    Do While iPos > 0
        If iPos > 0 Then
            CharCount = CharCount + 1
            iPos = iPos + Len(FindString)
            iPos = InStr(iPos, Expression, FindString)
        End If
    Loop
    
End Function

Public Function Split(ByVal Expression As String, Optional Delimiter As String = " ", _
                      Optional Limit As Long = -1, _
                      Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant
'===============================================================================
'   Split - This function mimics the VBA6.Split function which is not
'   available in VB5. This function helps VB5 & VB6 projects co-exist.
'
'   NOTE: Note, anyone making adjustments to this function should keep in
'   mind that it is intended to behave EXACTLY like the VBA6.Split
'   function (flaws and all, including operands).
'
'   WARNING: Like the VB6.Split function, do not use Split for
'   parsing a comma delimited string, as it does not conform to
'   the rules that govern the parsing of the comma delimited
'   file format.
'
'   Expression      String to parse.
'   Delimiter       Character to delimit fields.
'   Limit
'   Expression      Required. String expression containing substrings and
'                   delimiters. If expression is a zero-length string(""),
'                   Split returns an empty array, that is, an array with no
'                   elements and no data.
'   Delimiter       Optional. String character used to identify substring
'                   limits. If omitted, the space character (" ") is
'                   assumed to be the delimiter. If delimiter is a zero-length
'                   string, a single-element array containing the entire
'                   expression string is returned.
'   Limit           Optional. Number of substrings to be returned; A value
'                   of –1 (default) indicates that all substrings are returned.
'   Compare         Optional. Numeric value indicating the kind of comparison
'                   to use when evaluating substrings. See Settings section
'                   for values.
'
'   RETURNS         A Variant() array of substrings.
'===============================================================================
Const AvgSize   As Integer = 10 ' fudge factor: average field size
Dim ArrayIncr   As Long         ' array resizing increment
Dim j           As Long         ' work subscript
Dim k           As Long         ' work subscript
Dim StrPos()    As Long         ' string map: String field position
Dim StrLen()    As Long         ' string map: String field length
Dim ExprLen     As Long         ' length of the supplied string
Dim nElements   As Long         ' # of elements parsed
Dim RetArray    As Variant      ' not very efficient but the only way to RETURN a dynamic array in VB5

    Split = Empty                                   ' by default return a value of type Empty
    
    ExprLen = Len(Expression)                       ' save this for performance
    If ExprLen = 0 Then Exit Function               ' exit if nothing to parse
    
    If StrComp(Right$(Expression, Len(Delimiter)), _
       Delimiter, Compare) = 0 Then                 ' does the string have a trailing delimiter?
        Expression = Left$ _
            (Expression, ExprLen - Len(Delimiter))  ' strip trailing delimiter
        ExprLen = ExprLen - Len(Delimiter)          ' track new string length
    End If
    
    
    ArrayIncr = ExprLen / AvgSize                   ' make a guess of how many elements needed
    If ArrayIncr < 10 Then ArrayIncr = 20           ' size array at least 20 elements at a time
    
    Select Case Limit                               ' a limit has been placed on the # of elements
    Case Is > 0                                     ' a limit supplied
        ' all is cool
    Case 0, Is < -1                                 ' no elements allowed
        Exit Function                               ' we're done
    End Select
    
    If Len(Delimiter) = 0 Then                      ' delimiter cleared
        Split = Array(Expression)                   ' return a single element array
    End If
    
    ReDim StrPos(1 To ArrayIncr)                    ' initialize the array for looping
    ReDim StrLen(1 To ArrayIncr)                    ' initialize the array for looping
    
    k = 1                                           ' start parsing from position 1
    j = InStr(k, Expression, Delimiter)             ' locate the first delimiter
    If j = 0 Then                                   ' nothing found
        Split = Array(Expression)                   ' return a single element array
        Exit Function                               ' we're done
    Else                                            ' first element found
        While j > 0                                             ' loop until nothing left to parse
            nElements = nElements + 1                           ' track # of elements thus far
            StrPos(nElements) = k                               ' track the start of the read
            StrLen(nElements) = j - k                           ' track the # of characters to read
            k = j + Len(Delimiter)                              ' track next position
            j = InStr(k, Expression, Delimiter)                 ' get next position
            If nElements = UBound(StrPos) Then                  ' time to resize the array
                ReDim Preserve _
                    StrPos(1 To UBound(StrPos) + ArrayIncr)     ' resize the array & make room
                ReDim Preserve _
                    StrLen(1 To UBound(StrLen) + ArrayIncr)     ' resize the array & make room
            End If
        Wend
        
        ReDim RetArray(nElements)                               ' size the array to fit
        For j = 1 To nElements                                  ' loop thru each element
            RetArray(j) = _
                Mid$(Expression, StrPos(j), StrLen(j))          ' copy next substring into the array
        Next j
      
        Split = RetArray                                        ' return the array
    End If
  
End Function

Public Function Replace(ByVal Expression As String, ByVal FindString As String, ByVal ReplaceString As String, _
                        Optional ByVal Start As Long = 1, Optional ByVal ReplaceCount As Long = -1, _
                        Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As String
'===============================================================================
'   Replace - This function mimics the VBA6.Replace function which is not
'   available in VB5. This function helps VB5 & VB6 projects co-exist.
'
'   NOTE: Note, anyone making adjustments to this function should keep in
'   mind that it is intended to behave EXACTLY like the VBA6.Replace
'   function (flaws and all, including operands).
'
'   Expression      String to be searched.
'   FindString      String being sought.
'   ReplaceString   Replacement substring.
'   Start           Optional. Position within expression where substring
'                   search is to begin. If omitted, 1 is assumed.
'   ReplaceCount    Optional. Number of substring substitutions to perform.
'                   If omitted, the default value is –1, which means make
'                   all possible substitutions.
'   Compare         Optional. Numeric value indicating the kind of comparison
'                   to use when evaluating substrings. See Settings section
'                   for values.
'===============================================================================
Dim FindLen     As Long     ' length of search string
Dim Pos         As Long     ' position where search string located
Dim nCount      As Long     ' # of replacements

' TODO: This code bombs when there are embedded NULL characters
' in the string. The Len() function misreports; use lstrlenW

    On Error GoTo ErrHandler                            ' trap all errors
    
    FindLen = Len(FindString)                           ' ID length of search string
  
    Pos = 1                                             ' initialize the loop
    Do While Pos > 0
        Pos = InStr(Pos, Expression, FindString)        ' locate search string
        'Found, Is the string being replaced the same length as the search string
        'If not I must divide the string at the search string position
        If Pos > 0 Then                                 ' substring found
            nCount = nCount + 1
            Expression = Left$(Expression, Pos - 1) & _
                ReplaceString & _
                Mid$(Expression, Pos + FindLen)         ' make the replacement
            Pos = Pos + (Len(ReplaceString) - 1)        ' Advance our position to next find segment
        End If
        If nCount = ReplaceCount Then Exit Do           ' exit loop if max replacements reached
    Loop
    

ExitLabel:
    'After replacing the String, assign to ReturnString...
    Replace = Expression

Exit Function
Resume
ErrHandler:
    Debug.Print err.Description: Debug.Assert 0
    Resume ExitLabel
  
End Function

Public Function BrowseForFolder(hWnd As Long, _
                                Optional TopFolder As ShellVirtualFolders = CSIDL_DESKTOP, _
                                Optional Title As String = "Select folder", _
                                Optional Flags As BrowseFolderFlags = BIF_RETURNONLYFSDIRS, _
                                Optional DefaultFolder As String) As String
'===============================================================================
'   BrowseForFolder - Allows the user to browse for a folder using
'   the shell namespace.
'
'   TopFolder       Identifies the root folder.
'   Title           Attaches a description above the TreeView.
'   Flags           Misc flags defining behavior.
'   DefaultFolder   Reserved.
'===============================================================================
Dim bi      As BROWSEINFO   ' input structure used by Shell API
Dim idl     As ITEMIDLIST   ' input structure used by Shell API
Dim Result  As Long         ' return code
Dim Pidl    As Long         ' pointer to an Item ID List
Dim Path    As String       ' resulting path returned
Dim Pos     As Integer      ' parsing index

    BrowseForFolder = DefaultFolder                     ' Return default value if canceled

    bi.hOwner = hWnd
    Result = SHGetSpecialFolderLocation(hWnd, TopFolder, idl)
    
    bi.pidlRoot = idl.mkid.cb
    bi.lpszTitle = Title                                ' set the dialog title
    bi.ulFlags = Flags                                  ' input options
    Pidl = SHBrowseForFolder(bi)                        ' display browse dialog
      
' DON'T do this. It bombs on the WIDE version.
'   Path = Space$(520)                                  ' size of the buffer to fit
    Path = String(520, 0)                               ' size of the buffer to fit & pad with nulls
    
    If CBool(SHGetPathFromIDList(Pidl, Path)) Then      ' copy the path to local buffer
        Pos = InStr(Path, Chr$(0))                      ' find null terminator
        BrowseForFolder = Left$(Path, Pos - 1)          ' strip off null terminators
    End If

    CoTaskMemFree Pidl

End Function

Public Function HiWord(ByVal lParam As Long) As Long
'===============================================================================
'   HiWord - Returns Hi Word part of the supplied dual long value which is often needed
'   when working with long values returned by API calls.
'
'   lParam      Input. Dual value returned by API call.
'
'   RETURNS     the high order word of the supplied long value
'===============================================================================
   HiWord = (lParam And &HFFFF0000) \ &H10000
End Function

Public Function LoWord(ByVal lParam As Long) As Long
'===============================================================================
'   LoWord - Returns Lo Word part of the supplied dual long value which is often needed
'   when working with long values returned by API calls.
'
'   lParam      Input. Dual value returned by API call.
'
'   RETURNS     the low order word of the supplied long value
'===============================================================================
    If (lParam And &H8000&) <> 0 Then
       LoWord = lParam Or &HFFFF0000
    Else
       LoWord = lParam And &HFFFF&
    End If
      
End Function

Public Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
'===============================================================================
'   MakeDWord - Packs 2 16-bit integers into a 32-bit Long.
'
'   HiWord          The most significant 16 bits.
'
'   LoWord          The least significant 16 bits.
'
'   RETURNS:        A Long holding the full 32-bit value.
'===============================================================================
    
    MakeDWord = (HiWord * 65536) + (LoWord And &HFFFF&)

End Function

Public Function QuadFromLongs(ByVal LoPart As Long, ByVal HiPart As Long) As Currency
'===============================================================================
'   QuadFromLongs - Convert two 32 bit long values into a single 64 bit integer.
'
'   This function is typical for use with I/O functions that store values for
'   files or drives larger than 2147483647 bytes (~ 2GB).
'
'   IMPORTANT NOTE: The value returned by this function should never be passed
'   to an API function as "ByRef."  This is because the underlying value (and
'   the Currency data type) is scaled by a factor of ten.
'
'   NOTE: The calling convention differs from that of the mate subroutine
'   "QuadToLongs" so as not to get the two confused.
'
'   LoPart          Low order 32 bits.
'   HiPart          High order 32 bits.
'
'   RETURNS         A single 64 bit (unscaled) integer.
'===============================================================================
    
    Dim Quad        As TQuad
    Dim QuadParts   As TQuadParts
    
    ' Save Hi & Lo DWORDs
    QuadParts.LoWords = LoPart
    QuadParts.HiWords = HiPart
    
    ' Copy both parts into a single Currency value
    LSet Quad = QuadParts
    
    ' Return the result
    QuadFromLongs = Quad.Value * 10000

End Function

Public Sub QuadToLongs(ByVal Value As Currency, LoPart As Long, HiPart As Long)
'===============================================================================
'   QuadToLongs - Convert two 32 bit long values into a single 64 bit integer.
'
'   This function is typical for use with I/O functions that store values for
'   files or drives larger than 2147483647 bytes (~ 2GB).
'
'   NOTE: The calling convention differs from that of the mate subroutine
'   "QuadFromLongs" so as not to get the two confused.
'
'   Value           The 64 bit (unscaled) integer to convert.
'   LoPart          Output. Low order 32 bits of Value.
'   HiPart          Output. High order 32 bits of Value.
'===============================================================================
    
    Dim Quad        As TQuad
    Dim QuadParts   As TQuadParts
    
    ' Save 64 bit value & scale
    Quad.Value = Value / 10000
    
    ' Copy single Currency value into twop Long values
    LSet QuadParts = Quad
    
    ' Return the results in output buffers
    LoPart = QuadParts.LoWords
    HiPart = QuadParts.HiWords
    
End Sub
Public Function ProcAddress(ByVal lpFunction As Long) As Long
'===============================================================================
'   ProcAddress - Returns the address of the supplied function.
'
'   lpFunction      Function name. Passed "AddressOf <name>"
'
'   RETURNS         The 32 bit address of the function.
'===============================================================================
    ProcAddress = lpFunction
End Function

Public Function IsWinNt() As Boolean
'===============================================================================
'   IsWinNT - Returns true if we're running Windows NT.
'===============================================================================
    
    If OSVersion.dwOSVersionInfoSize = 0 Then               ' this is our first time making this call
        OSVersion.dwOSVersionInfoSize = Len(OSVersion)      ' initialize so API knows which version being used
        GetVersionEx OSVersion                              ' make the call once & then save/re-use it
    End If

    IsWinNt = OSVersion.dwPlatformId = VER_PLATFORM_WIN32_NT ' return the result
    
End Function

Public Function WinMajor() As Long

    If OSVersion.dwOSVersionInfoSize = 0 Then               ' this is our first time making this call
        OSVersion.dwOSVersionInfoSize = Len(OSVersion)      ' initialize so API knows which version being used
        GetVersionEx OSVersion                              ' make the call once & then save/re-use it
    End If

    WinMajor = OSVersion.dwMajorVersion                     ' return the result

End Function

Public Function ProperSubDir(ByVal Path As String) As String
'===============================================================================
'   ProperSubDir - Ensures that the supplied path (drive/directory) end in a backslash.
'
'   Path        The path to test for a trailing "\".
'
'   RETURNS     "Properly" formatted path with trailing path delimiter.
'===============================================================================
Dim Delimiter   As String   ' segmented path delimiter

    If InStr(Path, "://") > 0 Then          ' it's a URL path
        Delimiter = "/"                     ' use URL path delimiter
    Else                                    ' it's a disk based path
        Delimiter = "\"                     ' use disk based path delimiter
    End If
    
    Select Case Right$(Path, 1)             ' whats last character in path?
    Case "/", "\"                           ' it's one of the valid delimiters
        ProperSubDir = Path                 ' use the supplied path
    Case Else                               ' needs a trailing path delimiter
        ProperSubDir = Path & Delimiter     ' append it
    End Select
    
End Function

Public Function FileExists(ByVal Path As String, Optional Strict As Boolean = True) As Boolean
'===============================================================================
'   FileExists - Returns True if the supplied file exists.
'
'   Path        The path to the file to test.
'
'   RETURNS     True if the file exists (literally). If it's a directory, then
'               the function returns false. If access is denied, then
'===============================================================================
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_FILE_NOT_FOUND = 2&
Dim Attributes      As Long     ' Resulting file's attributes
Dim Pos             As Long     ' Ordinal delimeter
Dim SecondPass      As Boolean  ' True if second pass
Dim SavedError      As Long     ' error saved from first pass

    SetLastError 0                                          ' clear LastDLLError property
    
    If Len(Path) = 0 Then                                   ' no path supplied
        ' NOTE: Some API's use the current directory if a null string is passed.
        SetLastError ERROR_FILE_NOT_FOUND                   ' so caller knows file doesnt exist
        Exit Function                                       ' we're done
    End If

TryAgain:
    If IsWinNt Then                                         ' Running NT
        Attributes = GetFileAttributesW(StrPtr(Path))       ' Fetch the file's attributes
    Else                                                    ' Running Windows
        Attributes = GetFileAttributesA(Path)               ' Fetch the file's attributes
    End If

    ' NOTE: When we get an ERROR_ACCESS_DENIED on the file,
    ' we will attempt to verify that the parent dir exists.
    ' If we cannot, then we throw an error. Otherwise the
    ' file exists.
    
    If Attributes <> &HFFFFFFFF Then                        ' succeeded if not -1 returned
        If SecondPass And SavedError <> ERROR_ACCESS_DENIED Then
            ' We got an ERROR_ACCESS_DENIED on the FILE, and the
            ' parent directory exists. Therefore, we can assume
            ' the the file does exists [just no permission to
            ' access the file].
            FileExists = True
        Else
            If Strict Then                                  ' Only files allowed; no folders
                FileExists = ((Attributes And vbDirectory) = 0) ' return success if not a directory
            Else
                FileExists = True
            End If
            
        End If
            
    ElseIf SecondPass Then                                  ' call failed; Drive not ready, Device unavailable, etc.
        Call ApiRaise(SavedError)                           ' throw saved error if problem with root path
    Else
        
        Select Case err.LastDllError
        Case ERROR_ACCESS_DENIED
            
            SavedError = err.LastDllError                   ' save error before gets overwritten
            Do While Len(Path) > 3                          ' "3" is the magic # of the smallest path "A:\"
                ' Strip the file portion from the path.
                If Right$(Path, 1) = "\" Then
                    Path = Right$(Path, Len(Path) - 1)
                    Exit Do
                End If
            Loop
            
            SecondPass = True                               ' so we know not to try again
            GoTo TryAgain                                   ' give it one more shot
        
        Case ERROR_FILE_NOT_FOUND
            ' exit a failure
            
        Case Else
            Call ApiRaise                                   ' throw error if other problem
        End Select
        
    End If
    
End Function

Public Function DateToPackedJulian(ByVal TheDate As Date) As Long
'===============================================================================
'   DateToPackedJulian - Converts a date to a Julian day in the form of a packed DWORD.
'   The function works with dates based on the Gregorian (modern) calendar. Y2K COMPLAINT.
'
'   TheDate         The date to convert.
'   RETURNS         A DWORD whose LoWord is the year and the HiWord is the Julian day.
'===============================================================================
    TheDate = DateValue(TheDate)                ' strip the time part, if any
    
    ' split the date into two 16-bit numbers
    ' then pack into one 32-bit number:
    DateToPackedJulian = MakeDWord(Year(TheDate), _
                         (TheDate - DateSerial(Year(TheDate) - 1, 12, 31)))
End Function

Public Function DateFromPackedJulian(ByVal Value As Long) As Date
'===============================================================================
'   DateFromPackedJulian - Converts a packed Julian date (using DateToPackedJulian) to
'   a date data type.
'
'   Value           The julian value to convert.
'   RETURNS         A date representation of the julian date supplied.
'===============================================================================
    DateFromPackedJulian = DateSerial(LoWord(Value), 1, 1)  ' create date from Julian year
    DateFromPackedJulian = DateFromPackedJulian + _
        DateAdd("d", HiWord(Value) - 1, #12:00:00 AM#)      ' add in the # of days
End Function

Public Function SelectContents(Optional Control As TextBox) As Boolean
'===============================================================================
'   SelectContents - Selects the contents of a TextBox control.
'
'   Control     Optional. The subject TextBox control to select.
'               ActiveControl is assumed if not supplied.
'
'   RETURNS     True on success. Errors not thrown.
'===============================================================================

    On Error GoTo ErrHandler                ' in case disabled, not visible, whatever
    If Control Is Nothing Then              ' nothing supplied
        Set Control = Screen.ActiveControl  ' use the currently selected control
    End If
    
    Control.SelStart = 0                    ' start of selection
    Control.SelLength = Len(Control.Text)   ' select entire contents
    
Exit Function
Resume
ErrHandler:
    err.Clear
    Exit Function
    
End Function

Public Sub ChangeInputState(Control As Control, AllowInput As Boolean)
'===============================================================================
'   ChangeInputState - Sets the input state of the control. For a TextBox
'   or other intrinsic input control, sets the Locked property, otherwise
'   sets the Enabled property.
'
'   Control     The subject control to change state.
'   AllowInput  True to allow user input.
'===============================================================================

    On Error GoTo ErrHandler                        ' in case no Enabled property
    
    Select Case TypeName(Control)                   ' what kind of control?
    Case "Label"                                    ' do nothing
    Case "TextBox", "ComboBox"                      ' allows textual input
        Control.Locked = Not AllowInput             ' enable/disable input
        If AllowInput Then
            Control.BackColor = vbWindowBackground  ' use disabled color
        Else
            Control.BackColor = vbButtonFace        ' use enabled color
        End If
    
    Case Else                                       ' no user input
        Control.Enabled = AllowInput                ' just disable the control
    End Select
    
    
Exit Sub
Resume
ErrHandler:
    err.Clear
    Exit Sub
    
End Sub

Public Function TrimNulls(InString As String) As String
'===============================================================================
'   TrimNulls - Trims null terminators from a string.
'
'   InString        The string to strip nulls from.
'
'   RETURNS         The resulting string.
'===============================================================================
Dim Pos     As Long     ' location of null terminator

    Pos = InStr(InString, Chr$(0))              ' find null terminator
    If Pos > 0 Then                             ' terminator located
        TrimNulls = Left$(InString, Pos - 1)    ' strip terminator
    Else                                        ' no terminator found
        TrimNulls = InString                    ' return the supplied input string
    End If
    
End Function

Public Function ListFindString(ByVal SearchString As String, _
                               Control As Control, _
                               Optional StartFrom As Long = 0, _
                               Optional SelectIt As Boolean = False, _
                               Optional PartialMatch As Boolean = False, _
                               Optional HighlightIt As Boolean = True) As Long
'===============================================================================
'   ListFindString - Locates the index of SearchString within a ListBox
'   or a ComboBox.
'
'   SearchString    The string to search for in the control
'   Control         Control of type ListBox or ComboBox.
'   SelectIt        Optional. True to select the item if found.
'   PartialMatch    Optional. True if a partial matches allowed.
'   HighlightIt     Optional. True perform a "Quicken" style selection.
'                   NOTE: PartialMatch must also be True.
'   RETURNS         The index of the string within the ListBox/ComboBox,
'                   or -1 if not found.
'===============================================================================
Dim Msg_FINDSTRING  As Long     ' the "search" message sent to the control
Dim Msg_SETCURSEL   As Long     ' the "select" message sent to the control
Dim Msg_GETTEXTLEN  As Long     ' the "text length" message sent to the control
Dim Msg_GETTEXT     As Long     ' the "text" message sent to the control
Dim sContents       As String   ' text of item in the text portion of control
Dim sItem           As String   ' text of matching item
    ListFindString = -1                                 ' by default, return "Not Found" result
    
    ' ---Validate inputs, etc.
    If (Len(SearchString) = 0) Then                     ' nothing to look for
        Exit Function                                   ' we're done
    End If
    If StartFrom < 0 Then StartFrom = 0                 ' make sure we start from the bottom if out of range
    
    Select Case UCase$(TypeName(Control))               ' what type of control supplied
    
    Case "LISTBOX"                                      ' valid control
        Msg_FINDSTRING = IIf(PartialMatch, _
            LB_FINDSTRING, LB_FINDSTRINGEXACT)          ' use the appropriate message
        Msg_SETCURSEL = LB_SETCURSEL                    ' use the appropriate message
        Msg_GETTEXT = LB_GETTEXT                        ' use the appropriate message
        Msg_GETTEXTLEN = LB_GETTEXTLEN                  ' use the appropriate message
    Case "COMBOBOX"                                     ' valid control
        Msg_FINDSTRING = IIf(PartialMatch, _
            CB_FINDSTRING, CB_FINDSTRINGEXACT)          ' use the appropriate message
        Msg_SETCURSEL = CB_SETCURSEL                    ' use the appropriate message
        Msg_GETTEXT = CB_GETLBTEXT                      ' use the appropriate message
        Msg_GETTEXTLEN = CB_GETLBTEXTLEN                ' use the appropriate message
    
    Case Else                                           ' invalid control
        err.Raise 5, Module & ".ListFindString", "An invalid control type was supplied."
    End Select
    
    ' ---Issue the search command
    If IsWinNt Then                                     ' on NT, use WIDE calls
        ListFindString = SendMessageW(Control.hWnd, Msg_FINDSTRING, 0&, _
            ByVal StrPtr(SearchString & Chr$(0)))       ' find first match
    Else                                                ' on Windows, use ANSI calls
        ListFindString = SendMessageA(Control.hWnd, _
            Msg_FINDSTRING, 0&, ByVal SearchString)     ' find first match
    End If

    ' ---Perform "Quicken" style selection, if requested.
    If ListFindString > -1 _
    And HighlightIt _
    And PartialMatch _
    And UCase$(Control.Text) = UCase$(SearchString) Then
        On Error Resume Next
        sItem = String(SendMessageA(Control.hWnd, _
            Msg_GETTEXTLEN, ListFindString, 0&) + 2, 0) ' size buffer to size of string
        If IsWinNt Then                                 ' on NT, use WIDE calls
            SendMessageW Control.hWnd, Msg_GETTEXT, _
                ListFindString, ByVal StrPtr(sItem)     ' fetch the string
        Else                                            ' on Windows, use ANSI calls
            SendMessageA Control.hWnd, Msg_GETTEXT, _
                ListFindString, ByVal sItem             ' fetch the string
        End If
        sContents = Control.Text                        ' save contents before it changes
        sItem = TrimNulls(sItem)                        ' strip null terminators
        Control.Text = Control.Text & Right$(sItem, Len(sItem) - Len(sContents))
        Control.SelStart = Len(SearchString) 'sContents)               ' move cursor to end of contents
        Control.SelLength = Len(sItem) - Len(sContents) ' select remaining text
    End If
    
    If SelectIt Then                                    ' we need to select the contents of control
        ' NOTE: We use the API here in order to avoid
        ' the edit portion of the control from being
        ' changed.
        SendMessageA Control.hWnd, Msg_SETCURSEL, _
            ListFindString, 0&                          ' select item in the list, if any
    End If
 
End Function

Public Function InputRespond(KeyCode As Integer, ParamArray Overrides()) As Boolean
'===============================================================================
'   InputRespond - Use this function when you need to respond to input
'   via the KeyUp/Down event (instead of the KeyPress event). Take,
'   for example the "Start | Run" MRU input, or a Quicken style ComboBox
'   or to perform a quiet TextBox validation.
'
'   KeyCode     The key being struck.
'   Overrides() Array of keys to override from the "Select Case" list.
'
'   RETURNS     True if caller should respond to the user input.
'===============================================================================
Dim Override    As Variant

    For Each Override In Overrides          ' loop thru each override
        If KeyCode = Override Then          ' we have a valid input keycode
            InputRespond = True             ' return success
            Exit Function                   ' we're done
        End If
    Next Override

    ' NOTE: The following keys are not "usually" responded to.
    Select Case KeyCode
    Case vbKeyDelete, vbKeyTab, vbKeyShift, vbKeyNumlock, vbKeyPause
    Case vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyCapital
    Case vbKeyInsert, vbKeyPrint, vbKeyScrollLock, vbKeyPause
    Case vbKeyF1 To vbKeyF16, vbKeyEscape, vbKeyControl
    Case vbKeyTab, vbKeyExecute, vbKeyLeft To vbKeyDown, vbKeyMenu
    Case vbKeySnapshot, vbKeySeparator, vbKeySelect
    Case vbKeyReturn, vbKeyHelp, vbKeyCapital, vbKeyCancel, vbKeyBack
    Case Else                               ' appears to be an "ASNI" character
        InputRespond = True                 ' assume caller should respond
    End Select

End Function

Public Function GetUserId() As String
'===============================================================================
'   GetUserId
'
'   RETURNS     On success, the user name currently logged on to the network. If not
'               logged on to the network, returns the user id logged on to the system.
'               On failure, an empty string.
'===============================================================================
Static UserID   As String       ' The user currently logged on
Dim ErrorCode   As Long         ' result
Dim Success     As Boolean      ' result
Dim Buffer      As String       ' output buffer
Dim BufferLen   As Long         ' size of the output buffer

    On Error GoTo ErrHandler                        ' trap all errors

    If Len(UserID) > 0 Then                         ' we already know this info
        GoTo ExitLabel                              ' we're done
    End If
    
    Buffer = String(255, 0)                         ' pad buffer with nulls
    BufferLen = Len(Buffer)                         ' remember size of output buffer

    ' Fetch id user logged on to network [if logged on]
    ErrorCode = WNetGetUser(vbNullString, Buffer, BufferLen)

    If ErrorCode = 0 Then                           ' user is logged on to the network
        
        UserID = TrimNulls(Buffer)                  ' kill null terminators
    
    Else                                            ' not logged on to the network
        
        Success = getusername(Buffer, BufferLen)    ' get Windows' current user value
        If Success Then
            UserID = TrimNulls(Buffer)              ' kill null terminators
        End If
        
    End If

ExitLabel:
    GetUserId = UserID                              ' return the requested info

Exit Function
Resume
ErrHandler:
    Exit Function                                   ' just exit & return null string

End Function

Public Function ValidSSNFormat(SSN As Variant, _
                               Optional Output As Boolean = False, _
                               Optional StripDelimiters As Boolean = False, _
                               Optional EmptyIsValid As Boolean = False) As Boolean
'===============================================================================
'   ValidSSNFormat - Ensures that the SSN field is valid. Optionally, notifies the user
'   if invalid.
'
'   SSN             The social security number to verify. Variant to make flexible.
'   Output          True to output the supplied SSN in a properly formatted manner.
'   StripDelimiters True to remove the embedded delimiters like "-".
'                   Output must also be True.
'   EmptyIsValid    True if an empty SSN is valid.
'
'   RETURNS         True if SSN is valid, or if a default was used.
'===============================================================================
Dim WorkSSN     As String       ' work version of SSN
Dim WorkSSN2    As String       ' work version of SSN
Dim sChar       As String       ' work value
Dim i           As Integer      ' work

'   On Error GoTo ErrHandler                ' trap all errors
    
    WorkSSN = Trim$(SSN)                    ' trim of blanks
    
    If Len(WorkSSN) = 0 Then                ' empty ssn supplied
    
        ValidSSNFormat = EmptyIsValid
        If Output Then SSN = WorkSSN        ' return null string
        GoTo ExitLabel                      ' we're done
    
    End If
    
    For i = 1 To Len(SSN)                   ' loop thru each character in SSN
        
        sChar = Mid$(WorkSSN, i, 1)         ' fetch substring
        If IsNumeric(sChar) Then            ' its a number
            WorkSSN2 = WorkSSN2 & sChar     ' append to substring
        End If
    
    Next i
    
    If Len(WorkSSN2) = 9 Then               ' a valid SSN length supplied
        
        ValidSSNFormat = True               ' it's a valid SSN supplied
        If Output Then
            If StripDelimiters Then         ' need to output raw value
                SSN = WorkSSN2
            Else                            ' need to output formatted value
                SSN = Format(WorkSSN2, "###-##-####")
            End If
        End If
    
        ValidSSNFormat = True
    
   'Else invalid SSN/length supplied
    End If
    
    
ExitLabel:

Exit Function
ErrHandler:
'   ErrAttach           ' save error info
    Debug.Assert 0
'   ErrRaise            ' throw error back to caller

End Function

Public Function ValidTelFormat(Telephone As Variant, _
                               Optional Output As Boolean = False, _
                               Optional StripDelimiters As Boolean = False, _
                               Optional AlphaToNumeric As Boolean = False, _
                               Optional OutputFormat As String = "###-###-####", _
                               Optional EmptyIsValid As Boolean = True) As Boolean
'===============================================================================
'   ValidTelFormat - Ensures that the Telephone field is valid. Optionally,
'   notifies the user if invalid. The caller should submit a maskbased on
'   locale specific information.
'
'   Telephone       The telephone number to verify. Variant to make flexible.
'   Output          True to output the supplied Telephone in a properly formatted manner.
'   StripDelimiters True to remove the embedded delimiters like "-".
'                   Output must also be True.
'   AlphaToNumeric  True to convert valid dial pad letters to numeric equiv.
'   OutputFormat    Format to apply to output string.
'   EmptyIsValid    True if an empty Telephone is valid.
'
'   RETURNS         True if Telephone is valid, or if a default was used.
'                   NOTE: Non-USA countries not supported. Also, area code required.
'===============================================================================
Dim WorkTel     As String       ' work version of Telephone
Dim WorkTel2    As String       ' work version of Telephone
Dim sChar       As String       ' work value
Dim i           As Integer      ' work

'   On Error GoTo ErrHandler                        ' trap all errors
    
    WorkTel = Trim$(Telephone)                      ' trim of blanks
    If Left$(WorkTel, 1) = "1" Then                 ' strip leading "1"
        WorkTel = Right$(WorkTel, Len(WorkTel) - 1)
    End If
    
    If Len(WorkTel) = 0 Then                        ' empty Telephone supplied
    
        ValidTelFormat = EmptyIsValid
        If Output Then Telephone = WorkTel          ' return null string
        GoTo ExitLabel                              ' we're done
    
    End If
    
    For i = 1 To Len(Telephone)                     ' loop thru each character in Telephone
        
        sChar = Mid$(WorkTel, i, 1)                 ' fetch substring
        
        If IsNumeric(sChar) Then                    ' its a number
            WorkTel2 = WorkTel2 & sChar             ' append to substring
        Else
            Select Case UCase$(sChar)               ' attempt to convert characters to numbers
            Case "A", "B", "C": If AlphaToNumeric Then sChar = "2"
            Case "D", "E", "F": If AlphaToNumeric Then sChar = "3"
            Case "G", "H", "I": If AlphaToNumeric Then sChar = "4"
            Case "J", "K", "L": If AlphaToNumeric Then sChar = "5"
            Case "M", "N", "O": If AlphaToNumeric Then sChar = "6"
            Case "P", "R", "S": If AlphaToNumeric Then sChar = "7"
            Case "T", "U", "V": If AlphaToNumeric Then sChar = "8"
            Case "W", "X", "Y": If AlphaToNumeric Then sChar = "9"
            Case Else                               ' invalid digit is skipped
                GoTo ForNext
            End Select
            WorkTel2 = WorkTel2 & sChar             ' append to substring
        'Else                                       ' invalid digit is skipped
        End If
ForNext:
    Next i
    
    If Len(WorkTel2) = 10 Then                      ' a valid Telephone length supplied
        
        ValidTelFormat = True                    ' it's a valid Telephone supplied
        If Output Then
            If StripDelimiters Then                 ' need to output raw value
                Telephone = WorkTel2
            Else                                    ' need to output formatted value
                Telephone = Format$(WorkTel2, OutputFormat)
            End If
        End If
        
        ValidTelFormat = True
    
   'Else invalid Telephone/length supplied
    End If
    
ExitLabel:

Exit Function
ErrHandler:
'   ErrAttach           ' save error info
    Debug.Assert 0
'   ErrRaise            ' throw error back to caller

End Function

Public Function ValidZipFormat(ZipCode As Variant, _
                               Optional Output As Boolean = False, _
                               Optional StripPlus4 As Boolean = True, _
                               Optional EmptyIsValid As Boolean = True) As Boolean
'===============================================================================
'   ValidZipFormat - Ensures that the Zipcode field is valid. Optionally, notifies the user
'   if invalid.
'
'   ZipCode         The zip code to verify. Variant to make flexible.
'   Output          True to output the supplied zip code in a properly formatted manner.
'   StripPlus4      Removes trailing digits.
'   EmptyIsValid    True if an empty ZipCode is valid.
'
'   RETURNS         True if ZipCode is valid.
'===============================================================================
Dim WorkZip     As String       ' work version of ZipCode
Dim WorkZip2    As String       ' work version of ZipCode
Dim sChar       As String       ' work value
Dim i           As Integer      ' work

'   On Error GoTo ErrHandler                        ' trap all errors
    
    WorkZip = Trim$(ZipCode)                        ' trim of blanks
    
    If Len(WorkZip) = 0 Then                        ' empty ZipCode supplied
    
        ValidZipFormat = EmptyIsValid
        If Output Then ZipCode = WorkZip            ' return null string
        GoTo ExitLabel                              ' we're done
    
    End If
    
    For i = 1 To Len(ZipCode)                       ' loop thru each character in ZipCode
        
        sChar = Mid$(WorkZip, i, 1)                 ' fetch substring
        
        If IsNumeric(sChar) Then                    ' its a number
            WorkZip2 = WorkZip2 & sChar             ' append to substring
        'Else                                       ' invalid digit is skipped
        End If
    Next i
    
    Select Case Len(WorkZip2)
    Case 5                                          ' a valid ZipCode length supplied
        ' do nothing
    Case 9                                          ' MAYBE a valid ZipCode length supplied
        If Not StripPlus4 Then
            WorkZip2 = Left$(WorkZip2, 5)           ' strip of trailing digits
        End If
    Case Else                                       ' invalid ZipCode/length supplied
        GoTo ExitLabel
    End Select
    
ReturnSuccess:
    If Output Then
        ZipCode = WorkZip2
    End If

    ValidZipFormat = True
    
ExitLabel:

    Exit Function
    Resume
ErrHandler:
    Debug.Assert 0

End Function

Public Function IsCompiled() As Boolean
'============================================================================
'   IsCompiled - Identifies if code is executing under the development
'   environment or from a precompiled project.
'
'   4/6/2000    Many thanks to Thomas Leirer who pointed out a much better way
'               to implement IsCompiled. The original code is left behind for
'               laughs. Thanks again Thomas.
'
'   RETURNS:    True if code is NOT running from within the VB development
'               environment.
'============================================================================
'Static BeenHere As Boolean      ' static value saved for future calls
'Dim RetValue    As Boolean      ' static value saved for future calls
'Dim vArray      As Variant
'Dim sModule     As String
'Dim nChars      As Long

    On Error GoTo ExitLabel
    ' Produce a divide by zero error if running from code.
    Debug.Print 1 / 0

    IsCompiled = True
    
ExitLabel:

'    ' test if we have already executed function
'    If Not BeenHere Then
'        ' return previously saved value and exit
'        IsCompiled = RetValue
'        GoTo ExitLabel
'    End If
'
'    ' So we know to skip this part on the next call
'    BeenHere = True
'
'    ' Size buffer & initialize it to nulls
'    sModule = String(255, 0)
'    ' get the name of the EXE that is currently running
'    nChars = GetModuleFileName(VB.App.hInstance, sModule, 255)
'    sModule = Left$(sModule, nChars)
'
'    ' strip off the left-most part of the path
'    vArray = Split(sModule, "\")                   ' ie: c:\program files\pchute.exe
'    If Not IsArray(vArray) Then Exit Function
'
'    sModule = vArray(UBound(vArray))              '     xxxxxxxxxxxxxxxxx__________ [fell asleep here]
'
'    ' Compare the file name part.
'    IsCompiled = (UCase$(sModule) = Left$(UCase$(App.EXEName), Len(sModule)))
'ExitLabel:
End Function



Public Function LastDllErrorMsg(Optional ByVal Number As Long, _
                                Optional ByVal NewLineCharacter As String = vbCrLf) As String
'===================================================================================
'   LastDllErrorMsg - Returns the description of the supplied error number.
'
'   Number              Optional. The number to decode if supplied, otherwise the
'                       Err.LastDllError is used.
'   NewLineCharacter    Optional. Use to specificy an alternate character for CR/LFs
'===================================================================================
    
    Const ERROR_INSUFFICIENT_Message = 122&
    Const FORMAT_MESSAGE_FROM_HMODULE = &H800&
    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000&
    Const ERROR_EXTENDED_ERROR = 1208&
    Dim Message         As String   ' Place where error description will be nSize to.
    Dim Provider        As String   ' Place where error next provider output to
    Dim nSize           As Long     ' Number of bytes nSize to Message
    Dim Result          As Long     ' general return code
    Dim Success         As Boolean
    Dim ErrorCode       As Long
    Dim PreservedDllErr As Long     ' saved error number; since error functions w/be calling this


    On Error GoTo ExitLabel                     ' There shouldnt be any errors here.
    PreservedDllErr = err.LastDllError          ' save LastDllError as it will be restored on exit
    
    If Number = 0 Then                          ' no error code supplied
        Number = err.LastDllError               ' use the VB last known API error code
    ElseIf (Number And &H8000&) = 0 Then        ' 'possible' HRESULT
        ' NOTE: We must often contend with errors returned by VB
        ' and other interfaces with a description of "OLE Automation Error".
        ' These HRESULTS can OFTEN be translated to a Windows' system error
        ' by it's low word. This is typical when trying to create an object
        ' (perhaps via DCOM) or when accessing an object via a system
        ' interface such as IShellFolder.
        Number = Number And &HFFFF&             ' Essentially, we strip vbObjectError
    End If
    
    ' init output Message & clear LastDllError
    SetLastError 0
    nSize = 0
    
    '  Translate generic error descriptions
    Select Case Number
    Case ERROR_EXTENDED_ERROR
        ' server has extended error info
        
        ' Initialize output buffers
        Message = String(512, 0)
        Provider = String(128, 0)
        
        ' Get the error info from the server
        Result = WNetGetLastErrorA(ErrorCode, Message, Len(Message), Provider, Len(Provider))
        If Result = 0 Then
            ' call succeeded
            
            ' Strip NULLs from the strings
            Message = TrimNulls(Message)
            Provider = TrimNulls(Provider)
        
        End If
        
        ' Build friendly message
        Message = Provider & " returned error #" & ErrorCode & " " & Message & "."
        nSize = Len(Message)

    Case 600 To 799
        ' RAS specific codes have an error base of 600
        
        ' Initialize our Message and translate the error code
        Message = String(512, 0)
        nSize = Len(Message)
        Result = RasGetErrorStringA(Number, Message, nSize)
        If Result = 0 Then
            ' call succeeded
            
            ' Fetch message size
            nSize = InStr(Message, vbNullChar) - 1
            ' Sanity check
            If nSize < 0 Then nSize = 0
        End If
        
    Case 12003 'ERROR_INTERNET_EXTENDED_ERROR
        ' wininet extended errors
        
        ' Fetch Message size needed
        nSize = 0
        Success = InternetGetLastResponseInfoA(Result, Message, nSize)
        
        ' Size Message to fit
        Message = String(nSize + 1, 0)
        Success = InternetGetLastResponseInfoA(Result, Message, (nSize))

        ' BTW: Parsing this string may allow us to re-map
        '      to a real error. This is where we end up when
        '      the wininet parser fails (even on simple things
        '      that might not represent an error condition).

    Case 12000 To 12999
        ' wininet specific errors
        
        ' This functions assumes that wininet.dll is already mapped into memory
        Result = GetModuleHandleA("wininet.dll")
        
        If Result <> 0 Then
            
            ' Initialize our Message and translate the error code
            Message = Space(256)
            nSize = _
                FormatMessageA(FORMAT_MESSAGE_FROM_HMODULE, _
                               ByVal Result, Number, 0&, Message, 255, ByVal 0&)
        
        ' Else: we could be nice and temporarily load the library
        End If


    Case Else
        ' All other Windows errors
        
        ' Initialize our Message and translate the error code
        Message = Space(256)
        nSize = _
            FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM, _
                           ByVal 0&, Number, 0&, Message, 255&, ByVal 0&)
    End Select

    ' return the error description
    LastDllErrorMsg = Left$(Message, nSize)

    ' If the caller wants to translate line breaks then do so
    If StrComp(NewLineCharacter, vbCrLf, vbBinaryCompare) <> 0 Then
        LastDllErrorMsg = Replace(LastDllErrorMsg, vbCrLf, Left$(NewLineCharacter, 1))
    End If
    
    ' Strip ending CR/LF's as many API messages like to do.
    If Right$(LastDllErrorMsg, 2) = vbCrLf Then
        LastDllErrorMsg = Left$(LastDllErrorMsg, Len(LastDllErrorMsg) - 2)
    End If

    ' Terminate description with a period if not.
    Select Case Right$(LastDllErrorMsg, 1)
    Case ".", "?", "!"
        ' this case not very locale friendly
    Case Else
        LastDllErrorMsg = LastDllErrorMsg & "."
    End Select

ExitLabel:
    ' Restore Err.LastDllError value is it was on entry
    SetLastError PreservedDllErr
    
    Debug.Assert (err.Number = 0)
    
End Function

Public Function IsValidHandle(ByVal Handle As Long) As Boolean

    Const INVALID_HANDLE_VALUE = -1
    
    Select Case Handle
    Case 0, INVALID_HANDLE_VALUE
    Case Else
        IsValidHandle = True
    End Select
    
End Function

Public Function FolderFromPath(ByVal Path As String) As String
'============================================================================
'   FolderFromPath - Returns the directory portion of the supplied path.
'   This function does not access the file system, therefore everything
'   to the right of the directory seperator "\" or "/" is considered to
'   be the path's file part and is stripped of of the resulting string.
'============================================================================
    
    Dim DirParts    As Variant
    Dim Delimiter   As String
    Dim i           As Long
    
    Path = Trim$(Path)
    If Len(Path) > 2 Then
        For i = Len(Path) To 1 Step -1
            Select Case Mid$(Path, i, 1)
            Case "/", "\"
                FolderFromPath = Left$(Path, i - 1)
                Exit Function
            End Select
        Next i
    End If
    
    FolderFromPath = Path
    
End Function

Public Function BytesToStr(Buffer() As Byte, _
                           Optional IsAnsi As Boolean = False, _
                           Optional IsUnicode As Boolean = False) As String
'===============================================================================
'   BytesToStr - Converts a bytes array to a friendly string. If neither IsAnsi
'   or IsUnicode is specified, then a determination will be made based on the
'   operating system.
'
'   Buffer()        The byte array to convert.
'   IsAnsi          Optional. True if buffer is known to be Ansi.
'   IsUnicode       Optional. True if buffer is known to be Unicode.
'
'   RETURNS         A BSTR given a byte array.
'===============================================================================

    Dim Unspecified     As Boolean
    
    ' Identify if string format specified
    Unspecified = (Abs(IsAnsi) + Abs(IsUnicode)) = 0
    
    If (IsWinNt And Unspecified) Or IsUnicode Then
        ' Return the WIDE string
        BytesToStr = Buffer
    Else
        ' convert to unicode from an ANSI string
        BytesToStr = StrConv(Buffer, vbUnicode)
    End If
    
End Function

Public Function PtrToStr(lpString As Long, Optional nBytes As Long = -1) As String
'===============================================================================
'   PtrToStr - Wrapper function for PtrToStrA and PtrToStrW based on the current
'   operating system.
'===============================================================================

    If IsWinNt Then
        PtrToStr = PtrToStrW(lpString, nBytes)
    Else
        PtrToStr = PtrToStrA(lpString, nBytes)
    End If
    
End Function

Public Function PtrToStrA(ByVal lpString As Long, Optional ByVal nBytes As Long = -1) As String
'===============================================================================
'   PtrToStrA - Copies the ANSI string from the pointer and returns it as
'   a BSTR (string).
'
'   lpString        Pointer to an ANSI string.
'   nBytes          The size of the string in bytes. If -1, then lpString is
'                   assumed to be null terminated.
'
'   RETURNS         The string as pointed to by lpString, or vbNullString
'                   if lpString is zero.
'===============================================================================

    Dim Buffer()    As Byte     ' size of buffer

    PtrToStrA = vbNullString                        ' return null pointer by default
    
    Select Case lpString                            ' validate the pointer
    Case 0, INVALID_HANDLE_VALUE                    ' bogus pointer supplied
        Exit Function                               ' return an empty string for null address
    End Select
    
    If nBytes = 0 Then                              ' zero length string
        Exit Function                               ' we're done
    ElseIf nBytes = -1 Then
        nBytes = lstrlenA(ByVal lpString)            ' how big is the string?
    End If
    If nBytes <= 0 Then Exit Function               ' exit if empty
    
    ReDim Buffer(1 To nBytes)                       ' size array to fit
    CopyMemory VarPtr(Buffer(1)), lpString, nBytes  ' copy string to byte array work area
    
    PtrToStrA = StrConv(Buffer, vbUnicode)          ' convert to Unicode
    
End Function

Public Function PtrToStrW(ByVal lpString As Long, Optional ByVal nBytes As Long = -1) As String
'===============================================================================
'   PtrToStrW - Copies the WIDE string from the pointer and returns it as
'   a BSTR (string).
'
'   lpString        Pointer to a WIDE string.
'   nBytes          The size of the string in bytes. If -1, then lpString is
'                   assumed to be null terminated.
'
'   RETURNS         The string as pointed to by lpString, or vbNullString
'                   if lpString is zero.
'===============================================================================
    
    Dim Buffer()    As Byte     ' size of buffer

    PtrToStrW = vbNullString                        ' return null pointer by default
    
    Select Case lpString                            ' validate the pointer
    Case 0, INVALID_HANDLE_VALUE                    ' bogus pointer supplied
        Exit Function                               ' return an empty string for null address
    End Select
    
    If nBytes = 0 Then                              ' zero length string
        Exit Function                               ' we're done
    ElseIf nBytes = -1 Then
        nBytes = lstrlenW(lpString) * 2             ' how big is the string?
    End If
    If nBytes <= 0 Then Exit Function               ' exit if empty
    
    ReDim Buffer(1 To nBytes)                       ' size array to fit
    CopyMemory VarPtr(Buffer(1)), _
        lpString, nBytes                            ' copy string to byte array work area
        
    PtrToStrW = Buffer                              ' Return the Unicode string

End Function


Public Function PtrToData(ByVal lpData As Long, ByVal DataType As VbVarType, _
                           Optional ByVal Length As Long = -1) As Variant
'===============================================================================
'   PtrToData - This functions takes a long pointer to a data element and
'   converts it to a variant of subtype DataType.
'
'   lpData          Pointer to the data element to convert.
'   DataType        Type of data being converted.
'   Length          Optional. Size of data [vbString and vbByte+vbArray only].
'                   However, in some cases we allow this member to be set.
'                   For instance, a caller may only need to copy the first 8
'                   bytes in the case of a Decimal data type which would be
'                   typical when working with API's and 64 bit values split
'                   into DWORDs.
'
'   RETURNS         The data element as a variant.
'===============================================================================

    On Error GoTo ErrHandler
    
    ' This SELECT CASE identifies the data type and initializes VARTYPE.
    ' The actual data is copied after the switch, except for the exceptions
    ' such as byte array, currency and string data types which are handled
    ' in-line.
    Select Case DataType
        
    Case vbByte
        
        PtrToData = CByte(PtrToData)
        Length = 1
        
    Case vbBoolean
        
        PtrToData = CBool(PtrToData)
        If Length = -1 Then Length = 2
        
    Case vbInteger
        
        PtrToData = CInt(PtrToData)
        If Length = -1 Then Length = 2
        
    Case vbLong
        
        PtrToData = CLng(PtrToData)
        If Length = -1 Then Length = 4
        
    Case vbSingle
        
        PtrToData = CSng(PtrToData)
        Length = 4
        
    Case vbDate
        
        PtrToData = CDate(PtrToData)
        Length = 8
        
    Case vbDouble
        
        PtrToData = CDbl(PtrToData)
        Length = 8
        
    Case vbDecimal
        
        PtrToData = CDec(PtrToData)
        If Length = -1 Then Length = 12
        
        ' NOTE: The Decimal data type uses the all of the VARIANT
        ' structure except for the 2 lead VARTYPE bytes. It's funny
        ' that if you look up the Decimal data type in the MSDN and
        ' VB documentation, you will find some conflict in terms of
        ' how many bytes the data type actually occupies. The Decimal
        ' data type is a 12 byte value.
        
        ' Copy the data into return value
        CopyMemory ByVal VarPtr(PtrToData) + 2, ByVal lpData, Length
        GoTo ExitLabel
        
    Case vbObject
        
        ' TODO: This is untested
        Set PtrToData = Nothing
        Length = 4
        
    Case (vbByte Or vbArray)
        
        Dim bArray()    As Byte
        
        ' Size temp array to fit
        ReDim bArray(1 To Length)
        ' Copy the data into the array
        CopyMemory ByVal VarPtr(bArray(1)), ByVal lpData, Length
        
        ' return the array
        PtrToData = bArray
        GoTo ExitLabel
    
    Case vbCurrency
    
        ' This case handled out of the switch because of scaling
        PtrToData = CCur(PtrToData)
        Length = 8
        
        ' Copy the data into return value
        CopyMemory ByVal VarPtr(PtrToData) + 8, ByVal lpData, Length
        
        ' Handle the scale
        PtrToData = PtrToData * 10000
        GoTo ExitLabel
    
    Case vbString
        
        ' Assumed to be a VB native WIDE string
        
        ' Convert to unicode string
        PtrToData = PtrToStrW(lpData, Length)
        GoTo ExitLabel
    
    Case Else
    
        ' we should never be here
        Debug.Assert 0
        
        ' Throw error to caller
        On Error GoTo 0
        err.Raise 5
    
    End Select
    
    ' NOTE: A Variant's descriptor, padding & union take up 8 bytes.
    ' Therefore, the data (or pointer to data) is located at bytes 9-12 or 9-16
    
    ' Copy the data into return value
    CopyMemory ByVal VarPtr(PtrToData) + 8, ByVal lpData, Length
        
ExitLabel:
    
    Exit Function
    Resume
ErrHandler:
    Debug.Assert 0
    Resume ExitLabel
    
End Function

Public Function TranslateColor(ByVal Color As OLE_COLOR, _
                               Optional hPalette As Long = 0) As OLE_COLOR
'===============================================================================
'   Converts system defined colors to RGB colors as well as an RGB
'   colorref to a color supported by the supplied/default palette.
'
'   Color       The subject color to translate.
'   hPalette    Palette used by the device, or zero for default palette.
'
'   RETURNS     RGB color value.
'===============================================================================
    
    ' -- test if the supplied color is a VB system color constant
    
'    If (&H80000000 And Color) = &H80000000 Then
'        ' it's a system color
'
'        ' the low order bits define the system color index
'        TranslateColor = GetSysColor(LoWord(Color))
'        If TranslateColor = 0 Then
'            TranslateColor = CLR_INVALID
'            ' Caller can check Err.LastDllError
'        End If
'
'    Else
'        ' it's an RGB color
        
        ' convert color if needed
        If OleTranslateColor(Color, hPalette, TranslateColor) <> 0 Then
            TranslateColor = CLR_INVALID
        End If

'    End If
    
End Function

Public Function ShellWait(ByVal PathName As String, _
                          Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, _
                          Optional ByVal TimeoutValue As Long = 0, _
                          Optional TimedOut As Boolean) As Long
'===============================================================================
'   ShellWait - Shells the supplied program and waits until the process
'   is completed.
'
'   Parameters      Identical in name and functionality to that of the
'                   VBA.Shell function, except for:
'   TimeOutValue    Input. Number of milliseconds to time out the wait process.
'   TimedOut        Output. Set to True if the wait timed out.
'
'   RETURNS         A long value that defines how the shell application
'                   exited [ignored]. The SDK says:
'
'   ERRORS          Reflected back and thrown to caller.
'
'   If the process has terminated, the termination status returned may be
'   one of the following:
'   * The exit value specified in the ExitProcess or TerminateProcess
'     function.
'   * The return value from the main or WinMain function of the process.
'   * The exception value for an unhandled exception that caused the
'     process to terminate.
'
'===============================================================================

    Dim hInstance   As Long     ' hInstance of shelled application
    Dim hProcess    As Long     ' hProcess of shelled application
    Dim nRetVal     As Long     ' generic return value; debugging only
    Dim nErr        As Long     ' saved error number
    Dim sErr        As Long     ' saved error description
    Dim Flags       As Long     ' OpenProces flags
    
    On Error GoTo ErrHandler    ' trap all errors for open process
    
    ' Shell the command; an error raises if fails
    hInstance = Shell(PathName, WindowStyle)
    
    ' clear the LastDLLError [system error] value
    err.Clear
    
    ' open the process of the just shelled application
    ' TODO: Or in SYNCHRONIZE for NT so we don't have to
    ' loop with a DoEvents, which is much better on the CPU.
    Flags = PROCESS_QUERY_INFORMATION
    If IsWinNt() Then
        Flags = Flags Or SYNCHRONIZE
    End If
    hProcess = OpenProcess(Flags, 1&, hInstance)
  
    Select Case hProcess
    Case 0, INVALID_HANDLE_VALUE
    
        ' Failure is most likely access denied to the process
        Select Case err.LastDllError
        Case ERROR_ACCESS_DENIED
            err.Raise 70  ' VB's version of "Access Denied"
            ' Or: Err.Raise -ERROR_ACCESS_DENIED, , LastDllErrorMsg
        Case Else
            err.Raise 5, Description:="There was an unknown error opening the shelled process. " & _
                                    "System error #" & err.LastDllError & "."
        End Select
    'Case Else: Is a valid handle to a process
    End Select
  
    If IsWinNt() Then
        
        ' Under NT we "wait" on the process handle for better CPU utilization,
        ' and only execute a DoEvents when there are message to be processed.
        ShellWait = WaitOnEvent(hProcess, TimeoutValue, QS_ALLINPUT)

    Else

        ' loop until we get an exit code for the process
        Do

            nRetVal = GetExitCodeProcess(hProcess, ShellWait)

            ' We should never stop here
            Debug.Assert CBool(nRetVal)

            ' TODO: Use WaitForSingleObjectEx if on NT
            ' because it uses less CPU cycles.
            DoEvents

        Loop While ShellWait = STILL_ACTIVE

    End If

ExitLabel:
    ' Close open process, if any
    nRetVal = CloseHandle(hProcess)
    
    Exit Function
    Resume
ErrHandler:
    Debug.Assert 0
    
    ' Save error info
    nErr = err.Number: sErr = err.Description
    
    ' Close open process, if any
    nRetVal = CloseHandle(hProcess)
    
    ' prepare to throw error back
    On Error GoTo 0
    ' Throw it back
    err.Raise nErr, , sErr
    
End Function

Public Sub Pause(ByVal Milliseconds As Long, _
                 Optional ByVal WakeMask As QueueStatusFlags = QS_ALLINPUT, _
                 Optional lpAbort As Long)

    Dim hEvent      As Long

    ' Sanity check.
    If Milliseconds <= 0 Then Exit Sub

    ' Create a bogus event to wait on. It will never be siganaled.
    hEvent = CreateEvent(ByVal 0, 1, 0, vbNullString)
    If IsValidHandle(hEvent) Then
        ' Wait until timeout expires
        WaitOnEvent hEvent, Milliseconds, WakeMask, lpAbort:=lpAbort
        ' Destroy the event
        CloseHandle hEvent
    Else
        ApiRaise err.LastDllError, Module, "Could not create event."
    End If
    
End Sub

Public Function WaitOnEvent(ByVal hEvent As Long, _
                            Optional ByVal TimeoutValue As Long = 15000, _
                            Optional ByVal WakeMask As QueueStatusFlags = QS_ALLINPUT, _
                            Optional WaitValue As Long, _
                            Optional ByVal lpAbort As Long) As Long
    
    ' Wait on the single event
    WaitOnEvent = WaitOnEvents(1, _
        VarPtr(hEvent), _
        False, _
        TimeoutValue, _
        WakeMask, _
        WaitValue, _
        lpAbort)

End Function

Public Function WaitOnEvents(ByVal nEvents As Long, _
                             ByVal lpEvents As Long, _
                             Optional ByVal WaitOnAll As Boolean = False, _
                             Optional ByVal TimeoutValue As Long = 15000, _
                             Optional ByVal WakeMask As QueueStatusFlags = QS_ALLINPUT, _
                             Optional WaitValue As Long, _
                             Optional ByVal lpAbort As Long) As Long
    
    Const INFINITE = &HFFFFFFFF         ' Infinite timeout
    Const WAIT_FAILED = -1&
    Const WAIT_TIMEOUT = 258&           ' Wait timed out
    Const WAIT_ABANDONED_0 = &H80&
    Const WAIT_OBJECT_0 = 0
    'onst WAIT_IO_COMPLETION = &HC0&    ' MsgWaitForMultipleObjectsEx

    Dim EndTime         As Long
    Dim TimeoutTicks    As Long
    Dim AbortVal        As Long
    
    ' Calculate time out value, if any supplied
    If TimeoutValue > 0 Then
        EndTime = timeGetTime() + TimeoutValue
    Else
        TimeoutTicks = INFINITE
    End If

    Do

        ' If an abort variable supplied...
        If lpAbort <> 0 Then
            ' Copy contents of abort variable to local variable
            CopyMemory ByVal VarPtr(AbortVal), ByVal lpAbort, 2
            ' Exit if the abort variable is set.
            If AbortVal <> 0 Then Exit Do
        End If

        ' Calculate # of milliseconds left before we must time out.
        If TimeoutValue > 0 Then
            TimeoutTicks = TickDiff(timeGetTime(), EndTime)
        End If

        ' Wait on the supplied event
        WaitValue = MsgWaitForMultipleObjects _
            (nEvents, ByVal lpEvents, Abs(WaitOnAll), TimeoutTicks, WakeMask)

        ' Process the result
        Select Case WaitValue
        Case WAIT_OBJECT_0 To (WAIT_OBJECT_0 + nEvents - 1), _
             WAIT_ABANDONED_0 To (WAIT_ABANDONED_0 + nEvents - 1)
            
            ' Return a one-based index identifying the event that signaled
            WaitOnEvents = WaitValue + 1
            Exit Do
        
        Case (WAIT_OBJECT_0 + nEvents)
        
            ' We need to process the msg queue. Refer to the SDK for more info.
            'WaitOnEvents = nEvents
            'Exit Do
            
        Case WAIT_TIMEOUT
            
            ' We've timed out. Exit function.
            Exit Do
        
        Case WAIT_FAILED
            
            ' There was an error. Handle probably invalid.
            ApiRaise err.LastDllError, , "Wait event(s) failed."
            
        Case Else
                    
            Debug.Assert 0
        
        End Select
        
        DoEvents

    Loop
    
End Function

Public Sub ApiRaise(Optional ByVal Number As Long, _
                    Optional ByVal Source As String, _
                    Optional ByVal DescPrefix As String, _
                    Optional ByVal HelpFile As String, _
                    Optional ByVal HelpContext As Long, _
                    Optional ByVal NewLineCharacter As String = vbCrLf)
'===============================================================================
'   ApiRaise - Builds and throws an error returned from a Win32 API call.
'
'   NewLineCharacter    Allows the caller to convert linefeeds to a space or
'                       other character. This helps when displaying log entries
'                       in a List that cannot display multiple lines.
'
'   All other parameters are the same as Err.Raise method.
'
'   NOTE: vbObjectError will be added to Number so as not to confict with VB/VBA.
'===============================================================================

    ' Use LastDllError if no Number supplied
    If Number = 0 Then
        Number = err.LastDllError
    End If
    
    ' Use EXE name if no Source supplied
    If Len(Source) = 0 Then
        Source = App.EXEName
    End If
    
    ' If a msg prefix is used then massage the message
    If Len(DescPrefix) > 0 Then
        
        ' Terminate prexif with a period if not.
        DescPrefix = Trim$(DescPrefix)
        Select Case Right$(DescPrefix, 1)
        Case ".", "?", "!"
            ' this case not very locale friendly
        Case Else
            DescPrefix = DescPrefix & "."
        End Select
        
        ' Build error message and throw it
        err.Raise Number + vbObjectError, Source, _
            DescPrefix & NewLineCharacter & LastDllErrorMsg(Number, NewLineCharacter), _
            HelpFile, HelpContext
    Else
        
        ' Build error message and throw it
        err.Raise Number + vbObjectError, Source, _
            LastDllErrorMsg(Number, NewLineCharacter), _
            HelpFile, HelpContext
    
    End If
    
End Sub

Public Sub ScreenToWindow(hWnd As Long, PtX As Long, PtY As Long)
'===========================================================================
'   ScreenToWindow - Converts the coordinates relative to the screen to
'   coordinates relative to the supplied window and returns it as a POINT.
'===========================================================================
    
    Dim RcWnd   As RECT
  
    GetWindowRect hWnd, RcWnd
    PtX = PtX - RcWnd.Left
    PtY = PtY - RcWnd.Top
  
End Sub

Public Sub WindowToScreen(hWnd As Long, PtX As Long, PtY As Long)
'===========================================================================
'   WindowToScreen - Converts the coordinates relative to the window to
'   screen coordinates and returns it as a POINT.
'===========================================================================
    
    Dim RcWnd   As RECT
  
    GetWindowRect hWnd, RcWnd
    PtX = PtX + RcWnd.Left
    PtY = PtY + RcWnd.Top
  
End Sub

Public Sub ClientToWindow(hWnd As Long, PtX As Long, PtY As Long)
'===========================================================================
'   ClientToWindow - Converts the coordinates relative to the supplied
'   window's client area to coordinates relative to the window and returns it
'   as a POINT.
'===========================================================================
    
    Dim RcWnd       As RECT
    Dim RcClient    As RECT
    Dim Pt          As POINTAPI
    
    ' Fetch supplied window's bounding  rectangle
    Call GetWindowRect(hWnd, RcWnd)
    
    ' Fetch client area of supplied window
    Call GetClientRect(hWnd, RcClient)
    
    ' Convert client coordinates relative to screen
    Pt.X = RcClient.Left: Pt.y = RcClient.Top
    Call ClientToScreen(hWnd, Pt)
    
    ' Convert the client coordinates relative to the window's origin
    PtX = PtX + (Pt.X - RcWnd.Left)
    PtY = PtY + (Pt.y - RcWnd.Top)
  
End Sub

Public Function ImageToIPicture(hImage As Long, ByVal PictureType As PictureTypeConstants, _
                                 Optional ByVal hPalette As Long, _
                                 Optional ByVal Width As Long, _
                                 Optional ByVal Height As Long) As IPicture
'===============================================================================
'   ImageToIPicture - Generic function to convert various image types to a
'   IPicture object. Icons, bitmaps, meta & enhanced meta files supported. For
'   cursors, use vbPicTypeIcon.
'
'   hImage          The image to convert.
'   PictureType     The kind of image to convert.
'   hPalette        Optional. BMP only - uses specified palette. Not required.
'   Width/Height    Optional. WMF only - specifies extents of image. Required.
'
'   RETURNS         The IPicture object for supplied image.
'===============================================================================

    Dim IID_IPicture    As GUID
    Dim PicDef          As PICTDESC
    Dim ErrorCode       As Long
    
    ' Ensure picture type is supported, and reference correct structure
    Select Case PictureType
    Case vbPicTypeBitmap
        PicDef.cbSizeofstruct = 12  ' members: cbSizeofStruct, picType, hImage
        PicDef.xExt_Or_hPal = hPalette
    
    Case vbPicTypeIcon
        PicDef.cbSizeofstruct = 16  ' members: cbSizeofStruct, picType, hImage, hPalette
    
    Case vbPicTypeMetafile
        PicDef.xExt_Or_hPal = Width
        PicDef.yExt = Height
        PicDef.cbSizeofstruct = 20  ' members: cbSizeofStruct, picType, hImage, xExt, yExt
    
    Case vbPicTypeEMetafile
        PicDef.cbSizeofstruct = 12  ' members: cbSizeofStruct, picType, hImage
    
    Case Else
        err.Raise 5, App.ProductName, "Invalid picture type."
    End Select
    
    ' FYI IPictureDisp GUID: "{7BF80981-BF32-101A-8BBB-00AA00300CAB}" [see OCIDL.H]
    ' Convert GUID string for IPicture interface to binary
    ErrorCode = CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), VarPtr(IID_IPicture))
    If ErrorCode <> 0 Then
        ApiRaise ErrorCode, App.ProductName, "Could not convert image to Picture."
    End If
    
    ' Fill remainder of picture description
    PicDef.picType = PictureType
    PicDef.hImage = hImage

    ' Convert it to a IPicture object. The image is destroyed when object is destroyed.
    Call OleCreatePictureIndirect(PicDef, IID_IPicture, 1&, ImageToIPicture)

End Function

Public Function hFontToStdFont(hFont As Long) As StdFont
'===============================================================================
'   hFontToStdFont - Convert a font handle to a StdFont object.
'
'   hFont           The font handle to convert.
'
'   RETURNS         The IPicture object for supplied image.
'===============================================================================

    Dim IID_IFont       As GUID
    Dim FntDef          As FONTDESC
    Dim LogFnt          As LOGFONT
    Dim WorkFont        As IFont
    Dim RetValue        As Long
    Dim ErrorCode       As Long

    ' Sanity check
    If Not IsValidHandle(hFont) Then
        err.Raise 5, App.ProductName, "Invalid picture type."
    End If
    
    ' Get information about the referenced font
    If IsWinNt() Then
        RetValue = GetObjectApiW(hFont, Len(LogFnt), ByVal VarPtr(LogFnt))
    Else
        RetValue = GetObjectApiA(hFont, Len(LogFnt), LogFnt)
    End If
    
    ' Throw error on failure
    If RetValue = 0 Then
        ApiRaise err.LastDllError, Module
    End If
    
    ' Fill FONTDESC with information needed to create the Font object
    With FntDef
        .cbSizeofstruct = Len(FntDef)
        .lpstrName = PtrToStr(VarPtr(LogFnt.lfFaceName(LBound(LogFnt.lfFaceName))))
        .cySize = LogFnt.lfHeight
        .sWeight = LogFnt.lfWeight
        .sCharset = LogFnt.lfCharSet
        .fItalic = LogFnt.lfItalic
        .fUnderline = LogFnt.lfUnderline
        .fStrikethrough = LogFnt.lfStrikeOut
    End With
    
    ' FYI IFontDisp GUID: "{BEF6E003-A874-101A-8BBA-00AA00300CAB}" [see OCIDL.H]
    ' Convert GUID string for IFont interface to binary
    ErrorCode = CLSIDFromString(StrPtr("{BEF6E002-A874-101A-8BBA-00AA00300CAB}"), VarPtr(IID_IFont))
    If ErrorCode <> 0 Then
        ApiRaise ErrorCode, App.ProductName, "Could not convert hFont to Font object."
    End If

    ' Convert it to a IFont object. The font is destroyed when object is destroyed.
    Call OleCreateFontIndirect(FntDef, IID_IFont, WorkFont)

    ' Return the result as a StdFont
    Set hFontToStdFont = WorkFont
    
End Function

Public Function HiMetricToPixels(HiMetricX As Long, HiMetricY As Long, PixelsX As Long, PixelsY As Long) As Boolean
'=========================================================================
'   HiMetricToPixels - Converts hi-metric units to pixels. Hi-metric is
'   units expressed via the varios OLE Picture objects.
'
'   NOTE: You can also call the LogicalToDeviceX/Y function as well.
'
'   HiMetricX           The X unit to convert.
'   HiMetricY           The Y unit to convert.
'   PixelsX             Output. The converted X unit.
'   PixelsY             Output. The converted Y unit.
'=========================================================================
    
    Const MM_HIMETRIC   As Long = 3
    Dim MapPt           As POINTAPI         ' Stores a point (x,y) to help convert from HIMETRIC to pixels.
    Dim SavedMapMode    As Integer          ' Stores the old map mode.
    Dim WidthOut        As Long             ' Holds picture width in pixels.
    Dim HgtOut          As Long             ' Holds picture height in pixels.
    Dim ReturnValue     As Long             ' Holds API return values.
    Dim DC              As Long             ' Screen DC
    
    ' get the screen device
    DC = GetDC(0&)
    
    ' Set the mapmode of the hdc to HIMETRIC
    ' and calculate the size in pixels
    SavedMapMode = SetMapMode(DC, MM_HIMETRIC)
    
    ' Save hi-metric values
    MapPt.X = HiMetricX
    MapPt.y = HiMetricY
    
    ' LPtoDP will convert the width and height to pixels.
    ReturnValue = LPtoDP(DC, MapPt, 1)
    If ReturnValue = 0 Then
       ' Restore SavedMapMode and exit
       GoTo ExitLabel
    End If

    ' Return the result
    PixelsX = Abs(MapPt.X): PixelsY = Abs(MapPt.y)

ExitLabel:

    ' Restore the mapping mode as it was
    ' & release the device context
    Select Case SavedMapMode
    Case 0, -1 'invalid_handle_value
    Case Else
        ReturnValue = SetMapMode(DC, SavedMapMode)
        ReleaseDC 0&, DC
    End Select

End Function

Public Property Get ComboDropState(Combo As ComboBox) As Boolean
'===============================================================================
'   ComboDropState property - Gets/Sets the width of the drop down portion of the control.
'===============================================================================
    ComboDropState = (SendMessage(Combo.hWnd, CB_GETDROPPEDSTATE, 0&, 0&) > 0)
End Property
Public Property Let ComboDropState(Combo As ComboBox, ByVal NewValue As Boolean)
    If ComboDropState(Combo) <> NewValue Then
        ' If we're here, then the ComboBox is not in the desired state.
        ' Windows' normal behavior given this case is just to close-up
        ' the list portion, and nothing more, no click event should fire.
        
        SendMessage Combo.hWnd, CB_SHOWDROPDOWN, Abs(NewValue), 0&
    End If
End Property

Public Property Get ComboDropWidth(Combo As ComboBox) As Single
'===============================================================================
'   ComboDropWidth property - Gets/Sets the width of the drop down portion of the control.
'===============================================================================
Dim ReturnVal As Single

    ' convert to the desired scale
    ReturnVal = SendMessage(Combo.hWnd, CB_GETDROPPEDWIDTH, 0&, 0&)
    ComboDropWidth = Combo.Parent.ScaleX(ReturnVal, vbPixels, Combo.Parent.ScaleMode)

End Property
Public Property Let ComboDropWidth(Combo As ComboBox, ByVal NewValue As Single)
    
    ' convert to the desired scale
    NewValue = Combo.Parent.ScaleX(NewValue, Combo.Parent.ScaleMode, vbPixels)
    SendMessage Combo.hWnd, CB_SETDROPPEDWIDTH, CLng(Abs(NewValue)), 0&

End Property

Public Function LogicalToDeviceX(ByVal Width As Single, ByVal ScaleFrom As ScaleModeConstants, Optional ByVal hdc As Long) As Single
'===============================================================================
'   LogicalToDeviceX - Converts the supplied horizontal logical unit to pixels.
'
'   Width           The value to convert to pixels.
'   ScaleFrom       The unit of measurement to scale from.
'   hDC             Optional. The device context that Width relates to.
'
'   RETURNS         The value scaled to pixels.
'===============================================================================
    
    Dim LOGPIXELS       As Long
    
    ' No error handler
    On Error GoTo 0

    ' Use current meta device as a basic for the scale, if any,
    ' otherwise use the screen device.
    If IsValidHandle(hdc) Then
        ' Get # of pixels per logical  inch
        LOGPIXELS = GetDeviceCaps(hdc, LOGPIXELSX)
    Else
        ' Get the screen device
        hdc = GetDC(0&)
        ' Get # of pixels per logical  inch
        LOGPIXELS = GetDeviceCaps(hdc, LOGPIXELSX)
        ' Release the screen
        ReleaseDC 0, hdc
    End If
    
    ' Now for perform the conversion
    Select Case ScaleFrom
    Case vbCentimeters
        ' Convert from metric; 2540 = himetric units per inch
        LogicalToDeviceX = (Width * LOGPIXELS) / 2.54
        
    Case vbMillimeters
        ' Convert from metric; 2540 = himetric units per inch
        LogicalToDeviceX = (Width * LOGPIXELS) / 25.4
        
    Case vbHimetric
    
        ' Convert from himetric; 2540 = himetric units per inch
        LogicalToDeviceX = (Width * LOGPIXELS) / 2540
    
    Case vbTwips
        ' Convert to twips; 1440 = twips per inch
        LogicalToDeviceX = (Width * LOGPIXELS) / 1440
    
    Case vbCharacters
        ' Convert to twips; 1 horizontal character = 120 twips per unit; vertical = 240 twips per unit
        LogicalToDeviceX = ((Width * LOGPIXELS) * 120) / 1440
    
    Case vbInches
        ' Convert to inches; 1440 = twips per inch
        LogicalToDeviceX = Width * LOGPIXELS
    
    Case vbPoints
        ' Convert to points; 72 = points per inch
        LogicalToDeviceX = ((Width * LOGPIXELS) * 20) / 1440
    
    Case vbPixels
        ' No conversion needed for pixels
        LogicalToDeviceX = Width
    
    Case Else
    
        err.Raise 5, Module
    
    End Select
    

ExitLabel:
    
    Exit Function

End Function


Public Function LogicalToDeviceY(ByVal Height As Single, ByVal ScaleFrom As ScaleModeConstants, Optional ByVal hdc As Long) As Single
'===============================================================================
'   LogicalToDeviceY - Converts the supplied vertical logical unit to pixels.
'
'   Height          The value to convert to pixels.
'   ScaleFrom       The unit of measurement to scale from.
'   hDC             Optional. The device context that Width relates to.
'
'   RETURNS         The value scaled to pixels.
'===============================================================================
    
    Dim LOGPIXELS       As Long
    
    ' No error handler
    On Error GoTo 0

    ' Use current meta device as a basic for the scale, if any,
    ' otherwise use the screen device.
    If IsValidHandle(hdc) Then
        ' Get # of pixels per logical  inch
        LOGPIXELS = GetDeviceCaps(hdc, LOGPIXELSY)
    Else
        ' Get the screen device
        hdc = GetDC(0&)
        ' Get # of pixels per logical  inch
        LOGPIXELS = GetDeviceCaps(hdc, LOGPIXELSY)
        ' Release the screen
        ReleaseDC 0, hdc
    End If
    
    ' Now for perform the conversion
    Select Case ScaleFrom
    Case vbCentimeters
        ' Convert from himetric; 2540 = himetric units per inch
        LogicalToDeviceY = (Height * LOGPIXELS) / 2.54
        
    Case vbMillimeters
        ' Convert from himetric; 2540 = himetric units per inch
        LogicalToDeviceY = (Height * LOGPIXELS) / 25.4
    
    Case vbHimetric
    
        ' Convert from himetric; 2540 = himetric units per inch
        LogicalToDeviceY = (Height * LOGPIXELS) / 2540

    Case vbTwips
        ' Convert to twips; 1440 = twips per inch
        LogicalToDeviceY = (Height * LOGPIXELS) / 1440

    Case vbCharacters
        ' Convert to twips; 1 horizontal character = 120 twips per unit; vertical = 240 twips per unit
        LogicalToDeviceY = ((Height * LOGPIXELS) * 240) / 1440

    Case vbInches
        ' Convert to inches; 1440 = twips per inch
        LogicalToDeviceY = Height * LOGPIXELS

    Case vbPoints
        ' Convert to points; 72 = points per inch
        LogicalToDeviceY = ((Height * LOGPIXELS) * 20) / 1440

    Case vbPixels
        ' No conversion needed for pixels
        LogicalToDeviceY = Height

    Case Else
    
        err.Raise 5, Module
    
    End Select
    

ExitLabel:
    
    Exit Function

End Function

Public Function DeviceToLogicalX(ByVal Width As Single, ByVal ScaleTo As ScaleModeConstants, Optional ByVal hdc As Long) As Single
'===============================================================================
'   DeviceToLogicalX - Converts the supplied Width to the desired logical unit.
'
'   Width           The value to convert from pixels.
'   ScaleFrom       The unit of measurement to scale from.
'   hDC             Optional. The device context that Width relates to.
'
'   RETURNS         The value scaled to pixels.
'===============================================================================

    Dim LOGPIXELS       As Long
    
    ' No error handler
    On Error GoTo 0

    ' Use current meta device as a basic for the scale, if any,
    ' otherwise use the screen device.
    If IsValidHandle(hdc) Then
        ' Get # of pixels per logical  inch
        LOGPIXELS = GetDeviceCaps(hdc, LOGPIXELSX)
    Else
        ' Get the screen device
        hdc = GetDC(0&)
        ' Get # of pixels per logical  inch
        LOGPIXELS = GetDeviceCaps(hdc, LOGPIXELSX)
        ' Release the screen
        ReleaseDC 0, hdc
    End If
    
    ' Now for perform the conversion
    Select Case ScaleTo
    Case vbCentimeters
        ' Convert to centimeters; 2540 = himetric units per inch
        DeviceToLogicalX = (Width * 2.54) / LOGPIXELS
        
    Case vbMillimeters
        ' Convert to millimeters; 2540 = himetric units per inch
        DeviceToLogicalX = (Width * 25.4) / LOGPIXELS
    
    Case vbHimetric
    
        ' Convert to himetric; 2540 = himetric units per inch
        DeviceToLogicalX = (Width * 2540) / LOGPIXELS
    
    Case vbTwips
        ' Convert from twips; 1440 = twips per inch
        DeviceToLogicalX = (Width * 1440) / LOGPIXELS

    Case vbCharacters
        ' Convert from twips; 1 horizontal character = 120 twips per unit; vertical = 240 twips per unit
        DeviceToLogicalX = ((Width * 1440) / LOGPIXELS) / 120
    
    Case vbInches
        ' Convert from twips; 1440 = twips per inch
        DeviceToLogicalX = Width / LOGPIXELS
    
    Case vbPoints
        ' Convert from twips; 72 = points per inch
        DeviceToLogicalX = ((Width * 1440) / LOGPIXELS) / 20
    
    Case vbPixels
        ' No conversion needed for pixels
        DeviceToLogicalX = Width
    
    Case Else
    
        err.Raise 5, Module
    
    End Select

ExitLabel:
    
    Exit Function

End Function

Public Function DeviceToLogicalY(ByVal Height As Single, ByVal ScaleTo As ScaleModeConstants, Optional ByVal hdc As Long) As Single
'===============================================================================
'   DeviceToLogicalY - Converts the supplied Height to the desired logical unit.
'
'   Height          The value to convert from pixels.
'   ScaleFrom       The unit of measurement to scale from.
'   hDC             Optional. The device context that Width relates to.
'
'   RETURNS         The value scaled to pixels.
'===============================================================================

    Dim LOGPIXELS       As Long
    
    ' No error handler
    On Error GoTo 0

    ' Use current meta device as a basic for the scale, if any,
    ' otherwise use the screen device.
    If IsValidHandle(hdc) Then
        ' Get # of pixels per logical  inch
        LOGPIXELS = GetDeviceCaps(hdc, LOGPIXELSY)
    Else
        ' Get the screen device
        hdc = GetDC(0&)
        ' Get # of pixels per logical  inch
        LOGPIXELS = GetDeviceCaps(hdc, LOGPIXELSY)
        ' Release the screen
        ReleaseDC 0, hdc
    End If

    ' Now for perform the conversion
    Select Case ScaleTo
    Case vbCentimeters
        ' Convert to centimeters; 2540 = himetric units per inch
        DeviceToLogicalY = (Height * 2.54) / LOGPIXELS

    Case vbMillimeters
        ' Convert to millimeters; 2540 = himetric units per inch
        DeviceToLogicalY = (Height * 25.4) / LOGPIXELS

    Case vbHimetric
    
        ' Convert to himetric; 2540 = himetric units per inch
        DeviceToLogicalY = ((Height * 2540) / LOGPIXELS)

    Case vbTwips
        ' Convert from twips; 1440 = twips per inch
        DeviceToLogicalY = (Height * 1440) / LOGPIXELS

    Case vbCharacters
        ' Convert from twips; 1 horizontal character = 120 twips per unit; vertical = 240 twips per unit
        DeviceToLogicalY = ((Height * 1440) / LOGPIXELS) / 240

    Case vbInches
        ' Convert from twips; 1440 = twips per inch
        DeviceToLogicalY = Height / LOGPIXELS

    Case vbPoints
        ' Convert from twips; 72 = points per inch
        DeviceToLogicalY = ((Height * 1440) / LOGPIXELS) / 20

    Case vbPixels
        ' No conversion needed for pixels
        DeviceToLogicalY = Height

    Case Else

        err.Raise 5, Module

    End Select

ExitLabel:

    Exit Function

End Function

Public Function MakeQuad(HighWord As Long, LowWord As Long) As Variant
'===============================================================================
'   MakeQuad - Convert two 32 bit integers into a 64 bit value integer.
'===============================================================================

    ' This might seem a little strange, but it is correct
    ' (maybe not politically correct, however). Since the
    ' data is placed on the stack from right to left, we
    ' merely point to the right most word (LowWord) and
    ' copy the eight bytes into a Variant's Decimal type
    ' data area. Pretty simple.
    MakeQuad = PtrToData(VarPtr(LowWord), vbDecimal, 8)

End Function

Public Function HiDWord(ByVal vQuad As Variant) As Long
'===============================================================================
'   HiDWord - Returns the upper 32 bits of a 64 bit value [expressed as a
'   Decimal or Currency within a Variant].
'
'   NOTE: This function is often called when working with file system based API
'   calls. For that reason we force the "value" as a Decimal/Variant. This is to
'   avoid the hassle and confusion of scaling when using a Currency data type.
'   The technique of using the Currency data type in place of structures like
'   FILETIME and LARGE_INTEGER, et al., can be very tricky and is sometimes hard
'   to tell when and why they need to be scaled and de-scaled.
'
'   vQuad       The 64 bit value.
'
'   RETURNS     The high order DWORD of the supplied 64 bit value.
'===============================================================================

    ' Ensure it's of the Decimal data type
    Select Case VarType(vQuad)
    Case vbDecimal
        ' correct data type
    Case vbCurrency
        ' a valid 64 bit integer
        vQuad = CDec(vQuad)
    Case Else
        ' anything else would be uncivilized
        err.Raise 5
    End Select
    
    ' NOTE: A Variant's descriptor, padding & union take up 8 bytes.
    CopyMemory ByVal VarPtr(HiDWord), ByVal VarPtr(vQuad) + 8 + 4, 4

End Function

Public Function LoDWord(ByVal vQuad As Variant) As Long
'===============================================================================
'   HiDWord - Returns the lower 32 bits of a 64 bit value [expressed as a
'   Decimal or Currency within a Variant].
'
'   NOTE: This function is often called when working with file system based API
'   calls. For that reason we force the "value" as a Decimal/Variant. This is to
'   avoid the hassle and confusion of scaling when using a Currency data type.
'   The technique of using the Currency data type in place of structures like
'   FILETIME and LARGE_INTEGER, et al., can be very tricky and is sometimes hard
'   to tell when and why they need to be scaled and de-scaled.
'
'   vQuad       The 64 bit value.
'
'   RETURNS     The low order DWORD of the supplied 64 bit value.
'===============================================================================
    
    ' Ensure it's of the Decimal data type
    Select Case VarType(vQuad)
    Case vbDecimal
        ' correct data type
    Case vbCurrency
        ' a valid 64 bit integer
        vQuad = CDec(vQuad)
    Case Else
        ' anything else would be uncivilized
        err.Raise 5
    End Select
    
    ' NOTE: A Variant's descriptor, padding & union take up 8 bytes.
    CopyMemory ByVal VarPtr(LoDWord), ByVal VarPtr(vQuad) + 8, 4

End Function

Public Function GetSafeArrayInfo(TheArray As Variant, ArrayInfo As SAFEARRAY) As Boolean
'===============================================================================
'   GetSafeArrayInfo - Fills a SAFEARRAY structure for the supplied array. The
'   information contained in the SAFEARRAY structure allows the caller to
'   identify the number of dimensions and the number of elements for each
'   dimension (among other things). Element information for each dimension is
'   stored in a one-based sub-array of SAFEARRAYBOUND structures (rgsabound).
'
'   TheArray        The array to get information on.
'   ArrayInfo       The output SAFEARRAY structure.
'
'   RETURNS         True if the array is instantiated.
'===============================================================================

    Dim lpData      As Long         ' Pointer to the variants data item
    Dim VType       As Integer      ' the VARTYPE member of the VARIANT structure

    ' Exit if no array supplied
    If Not IsArray(TheArray) Then Exit Function
    
    With ArrayInfo
    
        ' Get the VARTYPE value from the first 2 bytes of the VARIANT structure
        CopyMemory ByVal VarPtr(VType), ByVal VarPtr(TheArray), 2
        
        ' Get the pointer to the array descriptor (SAFEARRAY structure)
        ' NOTE: A Variant's descriptor, padding & union take up 8 bytes.
        CopyMemory ByVal VarPtr(lpData), ByVal (VarPtr(TheArray) + 8), 4

        ' Test if lpData is a pointer or a pointer to a pointer.
        If (VType And VT_BYREF) <> 0 Then

            ' Get real pointer to the array descriptor (SAFEARRAY structure)
            CopyMemory ByVal VarPtr(lpData), ByVal lpData, 4
            
            ' This will be zero if array not dimensioned yet
            If lpData = 0 Then Exit Function
            
        End If

        ' Fill the SAFEARRAY structure with the array info
        ' NOTE: The fixed part of the SAFEARRAY structure is 16 bytes.
        CopyMemory ByVal VarPtr(ArrayInfo.cDims), ByVal lpData, 16

        ' Ensure the array has been dimensioned before getting SAFEARRAYBOUND information
        If ArrayInfo.cDims > 0 Then

            ' Size the array to fit the # of bounds
            ReDim .rgsabound(1 To .cDims)

            ' Fill the SAFEARRAYBOUND structure with the array info
            CopyMemory ByVal VarPtr(.rgsabound(1)), ByVal lpData + 16, ArrayInfo.cDims * Len(.rgsabound(1))

            ' So caller knows there is information available for the array in output SAFEARRAY
            GetSafeArrayInfo = True
            
        End If

    End With

End Function


Public Function MakeBLOB(ByVal lpBuffer As Long, ParamArray Params()) As Long
'===============================================================================
'   MakeBLOB - Creates a BLOB from a list of variables.
'
'   This functions helps to convert UDT's to BLOBs. The BLOBs can then be used
'   to save to disk or transported over a socket.
'
'   NOTE:   The function assumes that the caller has allocated enough memory
'   pointed to in lpBuffer. The caller is also expected to know the format of
'   the BLOB and how to restore it.
'
'   lpBuffer        Long pointer to the data buffer to receive the BLOB.
'                   A value of zero tells the function NOT to perform any
'                   copying. This allows the function to be called a first
'                   time to determine the total # of bytes needed.
'   Params          Parameter array of variables that will be written to the
'                   BLOB. Presently only intrinsic data types and byte arrays
'                   are supported. The math is quite simple to add support for
'                   arrays of other types, and recursion can be used for multi-
'                   dimensional arrays, which is probably beyond the scope of
'                   it's usage. For sophisticated structured BLOBs consider
'                   using the DataBag class instead.
'===============================================================================

    Dim lpData      As Long         ' Pointer to the variants data item
    Dim Length      As Long         ' Size of a Param
    Dim VType       As Long         ' VARTYPE member of VARIANT structure
    Dim i           As Long         ' work
    
    On Error GoTo ErrHandler
    
    ' NOTE: There is an undocumented bug in VB6 when using For/Each
    ' iterations on a ParamArray that results in the enumerated item
    ' being filled with a pointer to the data item instead of the
    ' actual data item. For instance, Params(1) would equate to StrPtr(Param)
    ' instead of Param's value if it were of type vbString. My testing
    ' found it not to be a problem from a compiled project. Therefore,
    ' I use a For/Next.

    For i = LBound(Params) To UBound(Params)
    
        ' Get the VARTYPE value from the first 2 bytes of the VARIANT structure
        CopyMemory ByVal VarPtr(VType), ByVal VarPtr(Params(i)), 2
        
        ' Get the pointer to the Variant's data item
        ' NOTE: A Variant's descriptor, padding & union take up 8 bytes.
        CopyMemory ByVal VarPtr(lpData), ByVal VarPtr(Params(i)) + 8, 4

        ' Test if lpData is a pointer or a pointer to a pointer.
        If (VType And VT_BYREF) <> 0 Then
            ' Get real pointer to the array descriptor (SAFEARRAY structure)
            CopyMemory ByVal VarPtr(lpData), ByVal lpData, 4
        End If
        
        ' This SELECT CASE identifies the data type and initializes VARTYPE.
        ' The actual data is copied after the switch, except for the exceptions.
        Select Case VarType(Params(i))
            
        Case vbByte
            
            ' Copy the variant's data value into the BLOB
            Length = 1
            If lpBuffer <> 0 Then _
                CopyMemory ByVal lpBuffer, ByVal VarPtr(Params(i)) + 8, Length
            
        Case vbBoolean, vbInteger
            
            ' Copy the variant's data value into the BLOB
            Length = 2
            If lpBuffer <> 0 Then _
                CopyMemory ByVal lpBuffer, ByVal VarPtr(Params(i)) + 8, Length
            
        Case vbLong, vbSingle
            
            ' Copy the variant's data value into the BLOB
            Length = 4
            If lpBuffer <> 0 Then _
                CopyMemory ByVal lpBuffer, ByVal VarPtr(Params(i)) + 8, Length
            
        Case vbDate, vbDouble, vbCurrency
            
            ' Copy the variant's data value into the BLOB
            Length = 8
            If lpBuffer <> 0 Then _
                CopyMemory ByVal lpBuffer, ByVal VarPtr(Params(i)) + 8, Length
            
        Case vbDecimal
            
            Length = 12
            If lpBuffer <> 0 Then _
                CopyMemory ByVal lpBuffer, ByVal VarPtr(Params(i)) + 2, Length
            
        Case (vbByte Or vbArray)
            
            ' An error will raise if the array is not dimensioned.
            ' This is OK, we'll just process the next parameter.
            On Error GoTo ForNext

            ' Size to # of elements in the byte array
            Length = Abs((LBound(Params(i))) - (UBound(Params(i)))) + 1
            
            ' NOTE: The pointer to array data (pvData) has a
            ' 12 byte offset in the SAFEARRAY structure.
            CopyMemory ByVal VarPtr(lpData), ByVal lpData + 12, 4
            
            ' Copy the data into the BLOB
            If lpBuffer <> 0 Then _
                CopyMemory ByVal lpBuffer, ByVal lpData, Length
        
        Case vbString
            
            ' Size to a unicode string
            Length = Len(Params(i)) * 2
            
            ' Copy the unicode string into the BLOB
            If lpBuffer <> 0 Then _
                CopyMemory ByVal lpBuffer, ByVal lpData, Length
            
        Case Else
        
            ' we should never be here
            Debug.Assert 0
            On Error GoTo 0
            err.Raise 5
        
        End Select
        
        ' Move current pointer in the buffer
        If lpBuffer <> 0 Then _
            lpBuffer = lpBuffer + Length
            
        ' So caller knows how many bytes needed/copied
        MakeBLOB = MakeBLOB + Length

ForNext:
        ' restore error handler in case cleared from above.
        On Error GoTo ErrHandler
    
    Next i
    
ExitLabel:
    
    Exit Function
    Resume
ErrHandler:
    Debug.Assert 0
    Resume ExitLabel
    
End Function

Public Sub SwapStrings(String1 As String, String2 As String)
'===============================================================================
'   SwapStrings - Swaps two strings by copying their underlying BSTR pointers,
'   without copying any of the underlying string data. Great for sorting
'   functions.
'
'   String1/String2     The strings to swap.
'===============================================================================

    Dim temp As String
      
    ' Copy into swap temp variable
    CopyMemory ByVal VarPtr(temp), ByVal VarPtr(String1), 4
    
    ' Swap the BSTR pointers
    CopyMemory ByVal VarPtr(String1), ByVal VarPtr(String2), 4
    CopyMemory ByVal VarPtr(String2), ByVal VarPtr(temp), 4

End Sub

Public Function HexBytes(ByteArray) As String
'===============================================================================
'   HexByte - Returns a hex string representation of the supplied bite.
'===============================================================================

    Dim i As Long
    
    ' Sanity check input array
    If VarType(ByteArray) <> vbByte + vbArray Then
        err.Raise 13, Module
    End If
    
    For i = LBound(ByteArray) To UBound(ByteArray)
        If Len(Hex$(ByteArray(i))) = 1 Then
            HexBytes = HexBytes & "0" & LCase$(Hex$(ByteArray(i))) & " "
        Else
            HexBytes = HexBytes & LCase$(Hex$(ByteArray(i))) & " "
        End If
    Next i
    
End Function


Public Function HexWORD(ByVal WORD As Integer, Optional Prefix As String = "0x") As String
'===============================================================================
'   HexDWORD - Returns a hex string representation of a DWORD.
'===============================================================================

    Dim Bytes(1 To 2)   As Byte
    Dim i               As Long
    
    CopyMemory ByVal VarPtr(Bytes(1)), ByVal VarPtr(WORD), 4

    HexWORD = Prefix
    For i = UBound(Bytes) To LBound(Bytes) Step -1
        If Len(Hex$(Bytes(i))) = 1 Then
            HexWORD = HexWORD & "0" & LCase$(Hex$(Bytes(i)))
        Else
            HexWORD = HexWORD & LCase$(Hex$(Bytes(i)))
        End If
    Next i
    
End Function


Public Function HexDWORD(ByVal DWORD As Long, Optional Prefix As String = "0x") As String
'===============================================================================
'   HexDWORD - Returns a hex string representation of a DWORD.
'===============================================================================

    Dim Bytes(1 To 4)   As Byte
    Dim i               As Long
    
    CopyMemory ByVal VarPtr(Bytes(1)), ByVal VarPtr(DWORD), 4

    HexDWORD = Prefix
    For i = UBound(Bytes) To LBound(Bytes) Step -1
        If Len(Hex$(Bytes(i))) = 1 Then
            HexDWORD = HexDWORD & "0" & LCase$(Hex$(Bytes(i)))
        Else
            HexDWORD = HexDWORD & LCase$(Hex$(Bytes(i)))
        End If
    Next i
    
End Function

Public Function HexQUAD(Quad, Optional Prefix As String = "0x") As String
'===============================================================================
'   HexQUAD - Returns a hex string representation of a QUAD/QWORD.
'===============================================================================

    Dim Bytes(1 To 8)   As Byte
    Dim i               As Long

    ' Sanity check input array
    If VarType(Quad) <> vbDecimal Then
        err.Raise 13, Module
    End If

    CopyMemory ByVal VarPtr(Bytes(1)), ByVal VarPtr(LoDWord(Quad)), 4
    CopyMemory ByVal VarPtr(Bytes(1)) + 4, ByVal VarPtr(HiDWord(Quad)), 4
    

    HexQUAD = "0x"
    For i = UBound(Bytes) To LBound(Bytes) Step -1
        If Len(Hex$(Bytes(i))) = 1 Then
            HexQUAD = HexQUAD & "0" & LCase$(Hex$(Bytes(i)))
        Else
            HexQUAD = HexQUAD & LCase$(Hex$(Bytes(i)))
        End If
    Next i
    
End Function

Public Function EncodeData(Data, DisplayFormat As EncodeFormats, Optional Prefix As String) As String
'===============================================================================
'   EncodeData - Returns a hex string representation of the supplied binary data.
'===============================================================================

    Dim Bytes()     As Byte         ' "String" varsion of "Data" converted to this
    Dim iStep       As Long         ' iterator step
    Dim Value       As Variant      ' Integer, Lond or Decimal data
    Dim nElements   As Long         ' # of elements in the byte array
    Dim EncodedStr  As String       ' Encoded string like "0x00000000"
    Dim lpValue     As Long
    Dim ccValue     As Long         ' # of characters
    Dim i           As Long
    Dim j           As Long
    
    ' Sanity check input array
    Select Case VarType(Data)
    Case vbByte + vbArray
        ' default data type to encode
    Case vbString
        Bytes = Data
        ' NOTE: Caller should consider converting the string to ANSI
        EncodeData = EncodeData(Bytes, DisplayFormat)
        Exit Function
    Case Else
        err.Raise 13, Module
    End Select
    
    Select Case DisplayFormat
    Case efHexBytes
        ' Don't loop if returning "byte" encoded values, just
        ' hand off the entire processing to another function.
        EncodeData = HexBytes(Data)
        Exit Function
    Case efHexWORDS
        iStep = 2
    Case efHexDWORDs
        iStep = 4
    Case efHexQWORDs
        iStep = 8
    Case Else
        err.Raise 5, Module
    End Select
    
    ' Count # of elements in the array and ensure
    ' the array is fixed to the size of the output
    ' format.
    nElements = Abs((LBound(Data)) - (UBound(Data))) + 1
    If nElements Mod iStep <> 0 Then
        err.Raise 5, Module, "Size of input data does not match output format."
    End If
    
    ' Size the output buffer to fit what will be returned.
    ' We do this because we don't want to perform any string
    ' concatenation which will kill our performance.
    ccValue = iStep + Len(Prefix) + 1   ' the "+ 1" is the (space) delimiter
    EncodeData = String((ccValue * nElements) - 1, " ")
    
    ' Loop thru entire array in 2, 4, or 8 byte increments.
    For i = LBound(Data) To UBound(Data) Step iStep
        
        ' Track # of actual conversions
        j = j + 1
        
        Select Case DisplayFormat
        Case efHexWORDS
            
            Value = PtrToData(VarPtr(Data(i)), vbInteger)
            EncodedStr = HexWORD(Value, Prefix)
            
        Case efHexDWORDs
            
            lpValue = VarPtr(Data(i))
            Value = PtrToData(lpValue, vbLong)
            EncodedStr = HexDWORD(Value, Prefix)
        
        Case efHexQWORDs
            
            ' NOTE: We only want the 8 bytes; not all 12
            Value = PtrToData(VarPtr(Data(i)), vbDecimal, 8)
            EncodedStr = HexQUAD(Value, Prefix)
        
        End Select
        
        ' Copy ONLY the bytes that represent the encoded value.
        ' The "* 2" is to account for Unicode.
        CopyMemory ByVal StrPtr(EncodeData) + (ccValue * 2) * (j - 1), _
            ByVal StrPtr(EncodedStr), Len(EncodedStr) * 2
        
    Next i

End Function



Function GetBlueColor(ColorValue As Long) As Long
'===============================================================================
'   GetBlueColor - Gets the blue color from a color value.
'===============================================================================

    GetBlueColor = ColorValue And &HFF&

End Function

Function GetGreenColor(ColorValue As Long) As Long
'===============================================================================
'   GetGreenColor - Gets the green color from a color value.
'===============================================================================

    GetGreenColor = (ColorValue And &HFF00&) \ 256

End Function

Function GetRedColor(ColorValue As Long) As Long
'===============================================================================
'   GetRedColor - Gets the red color from a color value.
'===============================================================================

    GetRedColor = (ColorValue And &HFF0000) \ 65536

End Function


Public Function HtmlHexColor(Color As Long) As String
'===============================================================================
'   HtmlHexColor - Converts a color value to a string suitable for html.
'   For example: #63FFFF
'===============================================================================

    If Len(Hex$(GetBlueColor(Color))) = 1 Then
        HtmlHexColor = "#" & "0" & Hex$(GetBlueColor(Color))
    Else
        HtmlHexColor = "#" & Hex$(GetBlueColor(Color))
    End If

    If Len(Hex$(GetGreenColor(Color))) = 1 Then
        HtmlHexColor = HtmlHexColor & "0" & Hex$(GetGreenColor(Color))
    Else
        HtmlHexColor = HtmlHexColor & Hex$(GetGreenColor(Color))
    End If

    If Len(Hex$(GetRedColor(Color))) = 1 Then
        HtmlHexColor = HtmlHexColor & "0" & Hex$(GetRedColor(Color))
    Else
        HtmlHexColor = HtmlHexColor & Hex$(GetRedColor(Color))
    End If

End Function

Public Function ProperCase(ByVal InputString As String, _
                           Optional ByVal ProperCaseRule As ProperCaseRules = pcDelimitSpaces) As String
'===============================================================================
'   ProperCase - Fixes what "StrConv(InputString, vbProperCase)" SHOULD do.
'   Each character is upper/lower cased as appropriate.
'===============================================================================

    Dim CapIt           As Boolean
    Dim iPos            As Long
    
    ' Initialize the loop
    CapIt = True
    
    ' Loop thru each character and change case as needed
    For iPos = 1 To Len(InputString)
        
        Select Case Asc(UCase$(Mid$(InputString, iPos)))
        Case Asc("A") To Asc("Z")
            ' Capitalize the letter if it's the first
            ' alpha character in the segment, otherwise
            ' lower case the character.
            If CapIt Then
                Mid$(InputString, iPos) = UCase$(Mid$(InputString, iPos))
                CapIt = False
            Else
                Mid$(InputString, iPos) = LCase$(Mid$(InputString, iPos))
            End If
        
        Case Asc(" ")
            ' TODO: We should actual query the system for this character
            CapIt = True
            
        Case Else
            
            ' So we know to UCase& next alpha character
            CapIt = (ProperCaseRule <> pcDelimitSpaces)
            
        End Select
    
    Next iPos

    ' Return the result
    ProperCase = InputString
    
End Function


Public Function TickCompare(ByVal TickStart As Currency, ByVal TickEnd As Currency) As Long
'===============================================================================
'   TickCompare - Compares the result of two GetTickCount/timeGetTime calls and
'   returns the difference. Two's complement and wrapping is handled.
'===============================================================================

    ' Handle two's complement for values larger than 2147483647&
    If TickStart < 0 Then
        TickStart = TickStart + CCur(2 ^ 32)
    End If

    ' Handle two's complement AND the case where
    ' timeGetTime/GetTickCount wraps at (2 ^ 32 or ~49.7 days):
    If TickEnd < 0 _
    Or TickEnd < TickStart Then
        TickEnd = TickEnd + CCur(2 ^ 32)
    End If

    ' Return the result
    Select Case TickStart
    Case TickEnd
        ' Return an EQUAL result
        TickCompare = 0
    Case Is < TickEnd
        ' Return a LESS THAN result
        TickCompare = -1
    Case Else
        ' Return a GREATER THAN result
        TickCompare = 1
    End Select

End Function

Public Function TickDiff(ByVal TickStart As Currency, ByVal TickEnd As Currency) As Long
'===============================================================================
'   TickDiff - Compares the result of two GetTickCount/timeGetTime calls and
'   returns the difference. Refer to the VBA.StrComp function for implementation.
'===============================================================================

    ' Handle two's complement for values larger than 2147483647&
    If TickStart < 0 Then
        TickStart = TickStart + CCur(2 ^ 32)
    End If

    ' Handle two's complement AND the case where
    ' timeGetTime/GetTickCount wraps at (2 ^ 32 or ~49.7 days):
    If TickEnd < 0 _
    Or TickEnd < TickStart Then
        TickEnd = TickEnd + CCur(2 ^ 32)
    End If

    ' Return the result
    TickDiff = TickEnd - TickStart

End Function


Public Function RandomString(Optional ByVal MinLength As Long = 20, _
                             Optional ByVal MaxLength As Long = 29, _
                             Optional ByVal ExclusionCharacters As String = " ", _
                             Optional ByVal RandomOption As RandomStringOptions = rsoAlphaNumericChars) As String
'===============================================================================
'   RandomString - Generates a random string. This would normally be done to
'   generate a random password. The odds of generating a string that can be
'   used in a dictionary attack are improbable, not impossible. Critical apps
'   may consider use a different random number generator.
'
'   Max/MinLength       Determines the minimum and maximum size of the string.
'   ExclusionCharacters Characters that cannot appear in the random string.
'   RandomOption        Defines special options used to define additional rules.
'
'   RETURNS             The random string.
'===============================================================================

    ' This value tell the Rnd to give random integers between 0-255 only
    Dim RndInt      As Integer      ' A random integer
    Dim PwLen       As String       ' Length of random string
    Dim Buffer()    As Byte         ' Where random string is built
    Dim iCharLo     As Integer      ' The lower range of the the ASCII/ANSI table;
    Dim iCharHi     As Integer      ' Either 128 or 256 (for extended ASCII/ANSI chars)
    Dim i           As Long
    Dim j           As Long

    ' Sanity check
    If MinLength < 1 Or MaxLength < MinLength Then
        err.Raise 5, Module
    End If
    
    ' Fire up the random number generator
    Randomize

    ' Create a random number between the desired string length range
    RndInt = Int((MaxLength - MinLength + 1) * Rnd + MinLength)

    ' Default lower bound character
    iCharLo = 0
    
    If RandomOption = rsoKeyboardChars Then
        ' -- only keyboard characters are supported
        
        ' Characters 32 thru 126 are keyboard characters
        iCharLo = 32: iCharHi = 126
            
    ElseIf RandomOption = rsoAlphaNumericChars Then
        
        ' This range included entire alpha/numeric characters
        iCharLo = 48: iCharHi = 122
        
    ElseIf RandomOption = rsoAllCharsExtended Then
        ' -- we can use the entire "standard" ascii character set
        iCharHi = 127
        
    Else ' RandomOption = rsoAllChars
        ' -- we can use the entire character set, including extended characters
        iCharHi = 255
    End If
    
    ' Size the buffer to fit
    ReDim Buffer(RndInt)

    ' Loop thru the output buffer
    For i = LBound(Buffer) To UBound(Buffer)
        
        If RandomOption = rsoAlphaNumericChars Then
            Do
                ' Get a random character in the character set range
                Buffer(i) = Int((iCharHi - iCharLo + 1) * Rnd + iCharLo)
                Select Case Buffer(i)
                Case 48 To 57, 65 To 90, 97 To 122
                    ' within the alpha numeric range of characters
                    Exit Do
                Case Else
                    ' just keep on looping until alpha
                    ' numeric character generated.
                End Select
            Loop
        Else
            ' Get a random character in the character set range
            Buffer(i) = Int((iCharHi - iCharLo + 1) * Rnd + iCharLo)
        End If
        
        ' Loop thru the unsupported exclusion characters.
TryAgain:
        For j = 1 To Len(ExclusionCharacters)
            
            ' If an unsupported character is identified then
            ' generate another random character and redrive
            ' the check for unsupported characters.
            If Buffer(i) = Asc(Mid$(ExclusionCharacters, 1)) Then
                Buffer(i) = Int((iCharHi - iCharLo + 1) * Rnd + iCharLo)
                GoTo TryAgain
            End If
            
        Next j

    Next i

    ' Return the resulting string
    RandomString = StrConv(Buffer, vbUnicode)
    
End Function

Public Function ParseURL(ByVal InputURL As String, _
                         Optional OutProtocol As String = vbNullString, _
                         Optional OutServer As String = vbNullString, _
                         Optional OutSubPath As String = vbNullString) As Boolean
'===============================================================================
'   ParseURL - Parses an internet address and dumps each path segment into their
'   respective output variable.
'
'   InputURL        The address should be in the form of:
'                       <protocol>://<ServerDnsOrIp>/<subpath>
'   OutProtocol     The protocol portion of the URL.
'   OutServer       The server name or IP address portion of the URL.
'   OutSubpath      The subpath portion of the URL.
'
'   RETURNS         True if string "appears" to be a valid URL.
'===============================================================================

    Const sProtocol     As String = "://"
    Dim iPos            As Long
    Dim iSeg            As Long
    Dim sSegment        As String
    
    ' Normalize the path
    InputURL = VBA.Replace(InputURL, "\", "/")
    
    ' Find protocol delimiter
    iPos = InStr(InputURL, sProtocol)
    If iPos = 0 Then Exit Function
    
    ' Dump the protocol string if output supplied
    If StrPtr(OutProtocol) <> 0 Then
        OutProtocol = Mid$(InputURL, 1, iPos + Len(sProtocol) - 1)
    End If
    
    ' Find next starting position
    iSeg = iPos + Len(sProtocol)
    
    ' Find the server/subpath delimeter
    iPos = InStr(iSeg, InputURL, "/")
    If iPos = 0 Then
        'OutServer = Mid$(InputURL, iSeg, Len(InputURL) - iSeg + 1)
        OutServer = Right$(InputURL, Len(InputURL) - iPos + 1)
    Else
        OutServer = Mid$(InputURL, iSeg, iPos - iSeg)
    End If
    
    ' Exit a failure if not server supplied
    If Len(OutServer) = 0 Then
        Exit Function
    End If
    
    ' Fill subpath only if an output was supplied
    If StrPtr(OutSubPath) <> 0 Then
        ' Fill subpath info preserving a leading "/" which
        ' is wininet will expect, especially when attempting
        ' to change to the root directory.
        OutSubPath = Right$(InputURL, Len(InputURL) - iPos + 1)
        If Right$(OutSubPath, 1) = "/" Then
            ' Strip trailing "/"; wininet doesn't like it.
            OutSubPath = Left$(OutSubPath, Len(OutSubPath) - 1)
        End If
    End If

    ' Return a mild success
    ParseURL = True

End Function

Public Function OpenURL(ByVal Url As String, _
                        Optional ByVal TimeOut As Long = 30, _
                        Optional ByVal Flags As Long) As String
'===============================================================================
'   OpenURL - Downloads the contents of an URL. Only IE's INTERNET_OPEN_TYPE_PRECONFIG
'   connection mode is supported for this function.
'
'   URL             The URL to download.
'   TimeOut         RESERVED. The number of seconds to ellapse before timing out.
'   Flags           See the SDK for InternetOpenUrl.
'
'   RETURNS         The contents of the downloaded URL.
'
'   NOTE:   This function uses memory objects for each downloaded chunk in order
'   to optimize the download process. We then "gather" the chunks of memory together
'   after the download is complete. We do it this way because we have no way of
'   telling how large the download is in advance. This approach allows us to avoid
'   performing incremental string concatenation on each chunk (the VB way).
'===============================================================================
    
    Const INTERNET_OPEN_TYPE_PRECONFIG = 0
    Const PageSize      As Long = 2048      ' block size of URL reads
    Const ChunkSize     As Long = 10        ' block sie of array dimensions
    Dim hInternet       As Long             ' Connection to the internet
    Dim hConnection     As Long             ' Connection to the server and URL
    Dim cbBuffer        As Long             ' # of downloaded bytes [ANSI]
    Dim lpPos           As Long             ' Position in the ouput buffer [on return]
    Dim MemObjs()       As TMemRecord       ' one memory object for each 2K page downloaded
    Dim nMemObjects     As Long
    Dim nBytesRead      As Long             ' # of bytes ACTUALLY read from URL [output]
    Dim sBuffer         As String
    Dim Success         As Boolean
    Dim ErrNum          As Long             ' Filled by error handler
    Dim ErrMsg          As String           ' Filled by error handler
    Dim i               As Long
    
    
    ' Trap all errors for resource cleanup
    On Error GoTo ErrHandler

    ' Open inet connection
    hInternet = InternetOpen(App.ProductName, _
                             INTERNET_OPEN_TYPE_PRECONFIG, _
                             vbNullString, _
                             vbNullString, _
                             0)
    
    ' Notify caller if failed
    If Not IsValidHandle(hInternet) Then
        ApiRaise err.LastDllError, Module, "Could not connect to the internet."
    End If

    ' Read the next chunk
    hConnection = InternetOpenUrl(hInternet, Url, ByVal 0&, 0&, Flags, 0&)
    
    ' Notify caller if failed
    If Not IsValidHandle(hConnection) Then
        ApiRaise err.LastDllError, Module, "Could not connect to the internet."
    End If

    ' Initialize array of memory objects
    ReDim MemObjs(1 To ChunkSize)
    
    ' Begin reading the page
    Success = True
    Do While Success
    
        ' Make sure there is room for another memory object
        If nMemObjects < UBound(MemObjs) Then
            ' make room for more memory object pointers
            ReDim Preserve MemObjs(1 To UBound(MemObjs) + ChunkSize)
        End If
        
        nMemObjects = nMemObjects + 1
        With MemObjs(nMemObjects)

            ' Fetch another 2K page.
            .hMemObject = GlobalAlloc(GPTR, PageSize)
            
            ' Read the next chunk. Keep running sum of bytes downloaded
            Success = InternetReadFile(hConnection, ByVal .hMemObject, PageSize, .cbMemObject)
            If .cbMemObject = 0 Then Exit Do

            cbBuffer = cbBuffer + .cbMemObject
            
        End With
        
    Loop
    
ExitLabel:
    
    ' Size buffer to (exactly) fit the return value
    If ErrNum = 0 Then
        OpenURL = String(cbBuffer, 0)
        lpPos = StrPtr(OpenURL)
    End If
    
    ' gather the output buffer from the scattered pointers
    For i = 1 To nMemObjects
        
        With MemObjs(i)
            
            ' Copy the downloaded chunk into our output buffer
            If ErrNum = 0 And .cbMemObject <> 0 Then
                sBuffer = PtrToStrA(.hMemObject, .cbMemObject)
                CopyMemory ByVal lpPos, ByVal StrPtr(sBuffer), .cbMemObject * 2
                lpPos = lpPos + (.cbMemObject * 2)
            End If

            ' Free the memory object
            .hMemObject = GlobalFree(.hMemObject)

        End With
        
    Next i
    
    
    ' Close connections. Bottom up since InternetCloseHandle will recurse.
    If IsValidHandle(hConnection) Then
        InternetCloseHandle hConnection
    End If
    If IsValidHandle(hInternet) Then
        InternetCloseHandle hInternet
    End If
    
    ' Throw error to caller if any
    If ErrNum <> 0 Then
        On Error GoTo 0
        err.Raise ErrNum, Module, ErrMsg
    End If
    
    Exit Function
    Resume
ErrHandler:
    Debug.Assert 0
    
    ' Save error info and resume
    ErrNum = err.Number
    ErrMsg = err.Description
    Resume ExitLabel
    
End Function

















'===============================================================================
'   ANSI/WIDE API Wrappers
'===============================================================================

Public Function SendMessage(hWnd As Long, wMsg As Long, wParam As Long, lParam As Long) As Long
'===============================================================================
'   SendMessage - ANSI/WIDE wrapper around API function.
'===============================================================================

    If IsWinNt Then                                             ' on NT, use WIDE calls
        SendMessage = SendMessageW(hWnd, wMsg, wParam, ByVal lParam)  ' make the call
    Else                                                        ' on Windows, use ANSI calls
        SendMessage = SendMessageA(hWnd, wMsg, wParam, ByVal lParam)
    End If
    
End Function

Public Function WNetGetUser(ByVal lpName As String, lpUserName As String, lpnLength As Long) As Long
'===============================================================================
'   WNetGetUser - ANSI/WIDE wrapper. Refer to the WNet SDK.
'===============================================================================
    If Len(lpName) = 0 Then lpName = vbNullString   ' ensure blanks made NULL
    
    If IsWinNt Then                     ' running in NT, use WIDE calls
        WNetGetUser = WNetGetUserW(StrPtr(lpName), StrPtr(lpUserName), lpnLength)
    Else                                ' running on Win9X, use ANSI calls
        WNetGetUser = WNetGetUserA(lpName, lpUserName, lpnLength)
    End If
    
End Function

Public Function getusername(lpBuffer As String, nSize As Long) As Long
'===============================================================================
'   GetUserName - ANSI/WIDE wrapper. Refer to the SDK.
'===============================================================================
    If Len(lpBuffer) = 0 Then lpBuffer = vbNullString   ' ensure blanks made NULL
    
    If IsWinNt Then                     ' running in NT, use WIDE calls
        getusername = GetUserNameW(StrPtr(lpBuffer), nSize)
    Else                                ' running on Win9X, use ANSI calls
        getusername = GetUserNameA(lpBuffer, nSize)
    End If
    
End Function

Public Function GetComputerName(lpBuffer As String, nSize As Long) As Long
'===============================================================================
'   GetComputerName - ANSI/WIDE wrapper. Refer to the SDK.
'===============================================================================
    If Len(lpBuffer) = 0 Then lpBuffer = vbNullString   ' ensure blanks made NULL
    
    If IsWinNt Then                     ' running in NT, use WIDE calls
        GetComputerName = GetComputerNameW(StrPtr(lpBuffer), nSize)
    Else                                ' running on Win9X, use ANSI calls
        GetComputerName = GetComputerNameA(lpBuffer, nSize)
    End If
    
End Function

Public Function GetModuleFileName(ByVal hModule As Long, lpFilename As String, ByVal nSize As Long) As Long
'===============================================================================
'   GetModuleFileName - ANSI/WIDE wrapper. Refer to the SDK.
'===============================================================================
    
    If Len(lpFilename) = 0 Then lpFilename = vbNullString ' ensure blanks made NULL
    
    If IsWinNt Then                                     ' running in NT, use WIDE calls
        GetModuleFileName = GetModuleFileNameW(hModule, StrPtr(lpFilename), nSize)
    Else                                                ' running on Win9X, use ANSI calls
        GetModuleFileName = GetModuleFileNameA(hModule, lpFilename, nSize)
    End If
    
End Function

Public Function SHGetPathFromIDList(ByVal Pidl As Long, ByRef pszPath As String) As Long
'===============================================================================
'   SHGetPathFromIDList - ANSI/WIDE wrapper. Refer to the SDK.
'===============================================================================

    If Len(pszPath) = 0 Then pszPath = vbNullString     ' ensure blanks made NULL
    
    If IsWinNt Then                                     ' running in NT, use WIDE calls
        SHGetPathFromIDList = SHGetPathFromIDListW(Pidl, StrPtr(pszPath))
    Else                                                ' running on Win9X, use ANSI calls
        SHGetPathFromIDList = SHGetPathFromIDListA(Pidl, pszPath)
    End If

End Function

Public Function WNetGetLastError(lpError As Long, lpErrorBuf As String, ByVal nErrorBufSize As Long, _
                                 lpNameBuf As String, ByVal nNameBufSize As Long) As Long
'===============================================================================
'   WNetGetLastError - ANSI/WIDE wrapper. Refer to the SDK.
'===============================================================================

    If IsWinNt Then                                     ' running in NT, use WIDE calls
        WNetGetLastError = WNetGetLastErrorW(lpError, StrPtr(lpErrorBuf), nErrorBufSize, StrPtr(lpNameBuf), nNameBufSize)
    Else                                                ' running on Win9X, use ANSI calls
        WNetGetLastError = WNetGetLastErrorA(lpError, lpErrorBuf, nErrorBufSize, lpNameBuf, nNameBufSize)
    End If

End Function

Public Function RasGetErrorString(ByVal uErrorValue As Long, ByVal lpszErrorString As String, ByVal cBufSize As Long) As Long

    If IsWinNt Then                             ' running on WinNT use WIDE calls
        RasGetErrorString = RasGetErrorStringW(uErrorValue, StrPtr(lpszErrorString), cBufSize)
    Else                                        ' running on Win95 use ANSI calls
        RasGetErrorString = RasGetErrorStringA(uErrorValue, lpszErrorString, cBufSize)
    End If
    
End Function

Public Function FormatMessage(ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

    If IsWinNt Then                             ' running on WinNT use WIDE calls
        FormatMessage = FormatMessageW(dwFlags, ByVal lpSource, dwMessageId, dwLanguageId, StrPtr(lpBuffer), nSize, Arguments)
    Else                                        ' running on Win95 use ANSI calls
        FormatMessage = FormatMessageA(dwFlags, ByVal lpSource, dwMessageId, dwLanguageId, lpBuffer, nSize, Arguments)
    End If

End Function



Public Function MapFtpResponse(ByVal Response As String, _
                               Optional ThrowError As Boolean = False) As Long
'===============================================================================
'   MapFtpResponse - Attempts to map responses from FTP servers to Win32
'   errors.  The WinInet.DLL fails to map certain errors/responses because
'   different messages are used by different servers. Such responses are
'   usually seen as ERROR_INTERNET_EXTENDED_ERROR -- in fact this error
'   is often displayed in IE as "The server returned extened information."
'   (You would think someone on the Internet Explorer team would supply
'   the extended error information by now).
'
'   For instance, the FTP response "550 Permission Denied." would be
'   mapped to the error ERROR_ACCESS_DENIED.
'
'   Refer to KB Article #Q168492 for more info about
'   ERROR_INTERNET_EXTENDED_ERROR.
'
'   Response    Response from the FTP server. Use the KillerVB function
'               LastDllErrorMsg(Err.LastDllError) which will detect this
'               condition and call the function InternetGetLastResponse
'               automatically.
'
'   ThrowError  Optional. True to throw the error back to caller.
'
'   RETURNS     ERROR_INTERNET_EXTENDED_ERROR, or a remapped Win32 error
'               code.
'
'   ERRORS      Thrown if ThrowError is True.
'
'===============================================================================

    Const ERROR_ACCESS_DENIED = 5&
    Const ERROR_PATH_NOT_FOUND = 3&
    Const ERROR_FILE_NOT_FOUND = 2&
    Const ERROR_INTERNET_CONNECTION_RESET = 12031&
    
    ' An extended error was returned from the server.
    ' This is typically a string or buffer containing
    ' a verbose error message. Call InternetGetLastResponseInfo
    ' to retrieve the error text:
    Const ERROR_INTERNET_EXTENDED_ERROR = 12003&

    Dim ErrNum      As Long
    Dim ErrMsg      As String
    Dim sResponse   As String

    On Error GoTo ErrHandler
    
    ' The parsing algorithm requires a space on both sides
    sResponse = " " & UCase$(Response) & " "

    If InStr(sResponse, "550 ") > 0 Then
    
        ' A 550 sResponse recieved.
        
        ' Some responses seen thus far:
        ' 550 <path>: The system cannot find the file specified.
        ' 550 <path>: No such file or directory.
        
        If (InStr(sResponse, "ACCESS") > 0 Or InStr(sResponse, "PERMISSION") > 0) _
        And InStr(sResponse, "DENIED") > 0 Then
            
            ' Would appear to be access denied to file/folder
            ApiRaise ERROR_ACCESS_DENIED
        
        ElseIf InStr(sResponse, "NOT FOUND") > 0 _
        Or InStr(sResponse, "NO SUCH") > 0 _
        Or InStr(sResponse, "CANNOT FIND") > 0 Then

            ' Would appear that a file/folder cannot be found

            If InStr(sResponse, "FILE") > 0 Then

                ' File cannot be found
                ApiRaise ERROR_FILE_NOT_FOUND

            ElseIf InStr(sResponse, "PATH") > 0 _
            Or InStr(sResponse, "DIRECTORY") > 0 Then

                ' Folder cannot be found
                ApiRaise ERROR_PATH_NOT_FOUND

            End If
        
        End If

    ElseIf InStr(sResponse, "421 ") > 0 _
    Or InStr(sResponse, "426 ") > 0 Then
            
        ' 421 Service not available, closing control connection.
        ' 426 Connection closed; transfer aborted.
        
        ' Map to connection reset
        ApiRaise ERROR_INTERNET_CONNECTION_RESET, DescPrefix:=Response
            
    End If

    ' If we could not map the error then we just
    ' return the supplied information.
    ApiRaise ERROR_INTERNET_EXTENDED_ERROR, DescPrefix:=Response

ExitLabel:

    Exit Function
    Resume

ErrHandler:
    
    ' Save error info
    ErrNum = err.Number: ErrMsg = err.Description
    
    If ThrowError Then
        ' Throw the remapped error.
        ApiRaise ERROR_INTERNET_CONNECTION_RESET, DescPrefix:=Response
    Else
        ' Throw the remapped error.
        MapFtpResponse = ErrNum
    End If

'Taken from: http://www.internic.net/std/std9.txt
'
'Postel & Reynolds                                              [Page 39]
'
'RFC 959                                                     October 1985
'File Transfer Protocol
'
'      4.2.2 Numeric  Order List of Reply Codes
'
'         110 Restart marker reply.
'             In this case, the text is exact and not left to the
'             particular implementation; it must read:
'                  MARK yyyy = mmmm
'             Where yyyy is User-process data stream marker, and mmmm
'             server 's equivalent marker (note the spaces between markers
'             and "=").
'        *119 Terminal not available, will try mailbox. [taken from Microsoft document not RFC 959]
'         120 Service ready in nnn minutes.
'         125 Data connection already open; transfer starting.
'         150 File status okay; about to open data connection.
'         202 Command not implemented, superfluous at this site.
'         211 System status, or system help reply.
'         212 Directory status.
'         213 File status.
'         214 Help message.
'             On how to use the server or the meaning of a particular
'             non-standard command.  This reply is useful only to the
'             human user.
'         215 NAME system type.
'             Where NAME is an official system name from the list in the
'             Assigned Numbers document.
'         220 Service ready for new user.
'         221 Service closing control connection.
'             Logged out if appropriate.
'         225 Data connection open; no transfer in progress.
'         226 Closing data connection.
'             Requested file action successful (for example, file
'             transfer or file abort).
'         227 Entering Passive Mode (h1,h2,h3,h4,p1,p2).
'         230 User logged in, proceed.
'         250 Requested file action okay, completed.
'         257 "PATHNAME" created.
'         331 User name okay, need password.
'         332 Need account for login.
'         350 Requested file action pending further information.
'         421 Service not available, closing control connection.
'             This may be a reply to any command if the service knows it
'             must shut down.
'         425 Can't open data connection.
'         426 Connection closed; transfer aborted.
'         450 Requested file action not taken.
'             File unavailable (e.g., file busy).
'         451 Requested action aborted: local error in processing.
'         452 Requested action not taken.
'             Insufficient storage space in system.
'         500 Syntax error, command unrecognized.
'             This may include errors such as command line too long.
'         501 Syntax error in parameters or arguments.
'         502 Command not implemented.
'         503 Bad sequence of commands.
'         504 Command not implemented for that parameter.
'         530 Not logged in.
'         532 Need account for storing files.
'         550 Requested action not taken.
'             File unavailable (e.g., file not found, no access).
'         551 Requested action aborted: page type unknown.
'         552 Requested file action aborted.
'             Exceeded storage allocation (for current directory or
'             dataset).
'         553 Requested action not taken.
'             File name not allowed.
End Function


Public Function SetDataMemory(Value As Variant, OldSize As Long) As Long
'===============================================================================
'   SetData - Updates or adds a new data element to the memory pool. This could
'   be a column name, a column value, whatever.
'===============================================================================

    Dim Size        As Long     ' Old buffer size
    Dim lpData As Long
    On Error GoTo ErrHandler
    
    Select Case VarType(Value)
    Case (vbByte Or vbArray)
        
        ' Size buffer to the # of elements in the byte array
        Size = Abs((LBound(Value)) - (UBound(Value))) + 1
    
    Case vbByte
        
        ' Size buffer to data type
        Size = 1           ' NOTE: We don't care about alignment
    
    Case vbBoolean, vbInteger
         
        ' Size buffer to data type
        Size = 2           ' NOTE: We don't care about alignment
    
    Case vbLong, vbSingle
        
        ' Size buffer to data type
        Size = 4
    
    Case vbCurrency, vbDate, vbDouble
        
        ' Size buffer to data type
        Size = 8
    
    Case vbDecimal

        ' Size buffer to data type
        Size = 12
    
    Case vbString
        
        ' Size buffer to # of Unicode characters
        Size = Len(Value) * 2

         ' Size buffer to fit if larger than previous buffer, if any
        lpData = GlobalAlloc(GMEM_ZEROINIT, Size)

        ' copy data into the buffer
        CopyMemory ByVal lpData, ByVal StrPtr(Value), Size
        
        ' we're done; bypass default processing
        GoTo ExitLabel
        
    Case Else
        Stop
        'Err.Raise 5
    End Select

     ' Size buffer to fit if larger than previous buffer, if any
    lpData = GlobalAlloc(GMEM_ZEROINIT, Size)

    ' copy data into the buffer
    MakeBLOB lpData, Value

ExitLabel:
    
    ' Return pointer to data element
    OldSize = Size
    SetDataMemory = lpData

    Exit Function
    Resume
ErrHandler:
    'Debug.Assert 0
    Resume ExitLabel
    
End Function


