Attribute VB_Name = "basBinary"
Option Explicit
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

'Used for Hooking into the key board.
Public Const WH_KEYBOARD = 2
Public Const CB_FINDSTRING = &H14C
'Used for the windows hook.
Public hHook As Long
'Used to track the last key pressed.
Public LastKeyPressed As Long

'Memory Constants

Public Const GMEM_ZEROINIT         As Long = &H40&
Public Const GENERIC_READ          As Long = &H80000000
Public Const GENERIC_WRITE         As Long = &H40000000

'text constants
Public Const EM_SETREADONLY = &HCF

Public Enum taciGramType
    msg_Recordset = 1
    msg_Message = 2
    msg_WorkQueue = 3
    msg_Reminder = 4
    msg_StationID = 5
    msg_Voice = 6
End Enum

'DataServer Constants
Public Const msg_MoveFirst = 1
Public Const msg_MoveNext = 2
Public Const msg_MovePrevious = 3
Public Const msg_MoveLast = 4
Public Const msg_BOF = 5
Public Const msg_EOF = 6
Public Const msg_OpenRecordset = 7

Public Type StringData
    Size As Long
    lpData As Long
End Type

Public Type taci_DataGram
    DataType As taciGramType
    BlobSize As Long
    Blob() As Byte
End Type

Public Type taci_Message
    WhoFrom As StringData
    WhoTo As StringData
    Subject As StringData
    Message As StringData
    Attachment() As Byte
End Type
    
Public Type taci_WorkQueue
    TimeArrived As Date
    SQLSting As String
End Type

Public Type taci_StationID
    HostName As StringData
    UserID As StringData
    RealName As StringData
End Type

Public Type taci_Connections
    SocketInstance As Long
    requestID As Long
    HostName As String
    UserName As String
End Type

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Public Declare Function GlobalReAlloc Lib "KERNEL32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessageC Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public AgentLogin As String
Public Conn() As taci_Connections
Public gSockInstance As Long

Public taciGram As taci_DataGram
Public Destbt() As Byte
Public bTotalBytesRecieved As Boolean
Public TotalBytesRecieved As Long
Public LastByteSize As Long
Public WritePosition As Long



Public Function PrepareAStationID() As Byte()
Dim vtBytes As Variant
Dim taciSTID As taci_StationID
Dim bt() As Byte
Dim taciGram As taci_DataGram
Dim RetVal As Variant

taciSTID.HostName.lpData = SetDataMemory(computername, taciSTID.HostName.Size)
taciSTID.RealName.lpData = SetDataMemory(getRealName(computername, localUserName), taciSTID.RealName.Size)
taciSTID.UserID.lpData = SetDataMemory(localUserName, taciSTID.UserID.Size)
'taciSTID.lpData = SetDataMemory(strMess, taciMsg.Message.Size)

taciGram.DataType = msg_StationID
taciGram.Blob = PrepareToSend(VarPtr(taciSTID), LenB(taciSTID), msg_StationID)
taciGram.BlobSize = UBound(taciGram.Blob)
ReDim Preserve bt(LenB(taciGram.BlobSize) + LenB(taciGram.DataType) + UBound(taciGram.Blob)) As Byte
CopyMemory ByVal VarPtr(bt(0)), ByVal VarPtr(taciGram), LenB(taciGram)
CopyMemory ByVal VarPtr(bt(LenB(taciGram.BlobSize) + LenB(taciGram.DataType))), ByVal VarPtr(taciGram.Blob(0)), UBound(taciGram.Blob)
PrepareAStationID = bt
End Function

Public Function AttachFromSend(bt() As Byte, lpsize As Long, inType As taciGramType) As Long

Select Case inType
    Case msg_Message
        Dim taciMsg As taci_Message
        Dim Subject, Towhom, FromWhom, strMess As String
        
        'WhoFrom Section
        lpsize = LenB(taciMsg.WhoFrom.Size)
        CopyMemory ByVal VarPtr(taciMsg.WhoFrom.Size), ByVal VarPtr(bt(LBound(bt))), LenB(taciMsg.WhoFrom.Size)
        FromWhom = PtrToData(VarPtr(bt(LBound(bt))) + lpsize, vbString, taciMsg.WhoFrom.Size)
        taciMsg.WhoFrom.lpData = SetDataMemory(FromWhom, LenB(FromWhom))
        
        'Whoto Section
        lpsize = lpsize + LenB(FromWhom)
        CopyMemory ByVal VarPtr(taciMsg.WhoTo.Size), ByVal VarPtr(bt(LBound(bt))) + lpsize, LenB(taciMsg.WhoTo.Size)
        lpsize = lpsize + LenB(taciMsg.WhoTo.Size)
        Towhom = PtrToData(VarPtr(bt(LBound(bt))) + lpsize, vbString, taciMsg.WhoTo.Size)
        taciMsg.WhoTo.lpData = SetDataMemory(Towhom, LenB(Towhom))
        
        'Subject Section
        lpsize = lpsize + LenB(Towhom)
        CopyMemory ByVal VarPtr(taciMsg.Subject.Size), ByVal VarPtr(bt(LBound(bt))) + lpsize, LenB(taciMsg.Subject.Size)
        lpsize = lpsize + LenB(taciMsg.Message.Size)
        Subject = PtrToData(VarPtr(bt(LBound(bt))) + lpsize, vbString, taciMsg.Subject.Size)
        taciMsg.Subject.lpData = SetDataMemory(Subject, LenB(Subject))
        
        'Message section
        lpsize = lpsize + LenB(Subject)
        CopyMemory ByVal VarPtr(taciMsg.Message.Size), ByVal VarPtr(bt(LBound(bt))) + lpsize, LenB(taciMsg.Message.Size)
        lpsize = lpsize + LenB(taciMsg.Message.Size)
        strMess = PtrToData(VarPtr(bt(LBound(bt))) + lpsize, vbString, taciMsg.Message.Size)
        taciMsg.Message.lpData = SetDataMemory(strMess, LenB(strMess))
        'Attach SQL code to insert the message into the database
        'With deTaci.rsMessages
        '    .AddNew
        '    .Fields("message_time") = Now()
        '    .Fields("from") = FromWhom
        '    .Fields("to") = Towhom
        '    .Fields("subject") = Subject
        '    .Fields("message") = strMess
        '    .Update
        'End With
        
        AttachFromSend = VarPtr(taciMsg)
    Case msg_WorkQueue
    
End Select
End Function

Public Function PrepareToSend(inType As Long, TypeSize As Long, taciType As taciGramType) As Byte()
Dim bt() As Byte
Dim lpsize As Long
Select Case taciType
    Case msg_WorkQueue
    
    Case msg_Message
        Dim taciMess As taci_Message
        CopyMemory ByVal VarPtr(taciMess), ByVal inType, TypeSize
        
        With taciMess
            ReDim bt(0 To (.Subject.Size + LenB(.Subject.Size) _
                + .Message.Size + LenB(.Message.Size) _
                + .WhoTo.Size + LenB(.WhoTo.Size) _
                + .WhoFrom.Size + LenB(.WhoFrom.Size))) As Byte
                
            CopyMemory ByVal VarPtr(bt(0)), ByVal VarPtr(.WhoFrom.Size), LenB(.WhoFrom.Size)
            lpsize = LenB(.WhoFrom.Size)
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal .WhoFrom.lpData, .WhoFrom.Size
            lpsize = lpsize + .WhoFrom.Size
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal VarPtr(.WhoTo.Size), LenB(.WhoTo.Size)
            lpsize = lpsize + LenB(.WhoTo.Size)
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal .WhoTo.lpData, .WhoTo.Size
            lpsize = lpsize + .WhoTo.Size
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal VarPtr(.Subject.Size), LenB(.Subject.Size)
            lpsize = lpsize + LenB(.Subject.Size)
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal .Subject.lpData, .Subject.Size
            lpsize = lpsize + .Subject.Size
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal VarPtr(.Message.Size), LenB(.Message.Size)
            lpsize = lpsize + LenB(.Message.Size)
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal .Message.lpData, .Message.Size
            lpsize = lpsize + .Message.Size
            
        End With
    Case msg_StationID
        Dim taciSTID As taci_StationID
        CopyMemory ByVal VarPtr(taciSTID), ByVal inType, TypeSize
        
        With taciSTID
            ReDim bt(0 To (.HostName.Size + LenB(.HostName.Size) _
                + .RealName.Size + LenB(.RealName.Size) _
                + .UserID.Size + LenB(.UserID.Size))) As Byte
                
            CopyMemory ByVal VarPtr(bt(0)), ByVal VarPtr(.HostName.Size), LenB(.HostName.Size)
            lpsize = LenB(.HostName.Size)
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal .HostName.lpData, .HostName.Size
            lpsize = lpsize + .HostName.Size
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal VarPtr(.RealName.Size), LenB(.RealName.Size)
            lpsize = lpsize + LenB(.RealName.Size)
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal .RealName.lpData, .RealName.Size
            lpsize = lpsize + .RealName.Size
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal VarPtr(.UserID.Size), LenB(.UserID.Size)
            lpsize = lpsize + LenB(.UserID.Size)
            CopyMemory ByVal VarPtr(bt(0)) + lpsize, ByVal .UserID.lpData, .UserID.Size
            lpsize = lpsize + .UserID.Size
            
        End With
        
End Select
PrepareToSend = bt

End Function

Public Function GetMemory(ByVal NewSize As Long) As Long
'===============================================================================
'   GetMemory - Helper function for SetData. Given the supplied inputs, it
'   evaluates whether a new memory block is needed, or if existing memory needs
'   to be resized, or if the existing memory area can be recycled.
'===============================================================================

    ' Size buffer to fit if larger than previous buffer, if any
    If NewSize <> 0 Then
        ' Create a new memory buffer
        GetMemory = GlobalAlloc(GMEM_ZEROINIT, NewSize)
    
    End If
    
End Function

Public Function PrepareAMessage(Subject As String, Towhom As String, FromWhom As String, strMess As String) As Byte()
Dim vtBytes As Variant
Dim taciMsg As taci_Message
Dim bt() As Byte
Dim taciGram As taci_DataGram
Dim RetVal As Variant

taciMsg.Subject.lpData = SetDataMemory(Subject, taciMsg.Subject.Size)
taciMsg.WhoTo.lpData = SetDataMemory(Towhom, taciMsg.WhoTo.Size)
taciMsg.WhoFrom.lpData = SetDataMemory(FromWhom, taciMsg.WhoFrom.Size)
taciMsg.Message.lpData = SetDataMemory(strMess, taciMsg.Message.Size)

taciGram.DataType = msg_Message
taciGram.Blob = PrepareToSend(VarPtr(taciMsg), LenB(taciMsg), msg_Message)
taciGram.BlobSize = UBound(taciGram.Blob)
ReDim Preserve bt(LenB(taciGram.BlobSize) + LenB(taciGram.DataType) + UBound(taciGram.Blob)) As Byte
CopyMemory ByVal VarPtr(bt(0)), ByVal VarPtr(taciGram), LenB(taciGram)
CopyMemory ByVal VarPtr(bt(LenB(taciGram.BlobSize) + LenB(taciGram.DataType))), ByVal VarPtr(taciGram.Blob(0)), UBound(taciGram.Blob)
PrepareAMessage = bt
End Function
Public Function PrepareAVoice(bt() As Byte) As Byte()
Dim retbt() As Byte
Dim taciGram As taci_DataGram
Dim RetVal As Variant
taciGram.DataType = msg_Voice
taciGram.Blob = bt
taciGram.BlobSize = UBound(taciGram.Blob)
ReDim Preserve retbt(LenB(taciGram.BlobSize) + LenB(taciGram.DataType) + UBound(taciGram.Blob)) As Byte
CopyMemory ByVal VarPtr(retbt(0)), ByVal VarPtr(taciGram), LenB(taciGram)
CopyMemory ByVal VarPtr(retbt(LenB(taciGram.BlobSize) + LenB(taciGram.DataType))), ByVal VarPtr(taciGram.Blob(0)), UBound(taciGram.Blob)
PrepareAVoice = retbt
End Function

