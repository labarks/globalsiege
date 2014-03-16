Attribute VB_Name = "modNetwork"
Option Explicit

'Remote UDP broadcast address used by clients when searching for a host.
Public Const cgBroadcastAddress = "255.255.255.255"

'The default known GlobalSiege port number.
Public Const gcDefaultPortNumber As Long = 4813

'A marker to identify remote terminals. Prepended to the terminal names.
Public Const gcRemoteTerminalMarker = ":"

Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function GetHostName Lib "ws2_32.dll" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Public Declare Function Netbios Lib "netapi32.dll" (pncb As NET_CONTROL_BLOCK) As Byte
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Function GetProcessHeap Lib "kernel32" () As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

Public Const INADDR_NONE = &HFFFF
Public Const SOCKET_ERROR = -1
Public Const WSABASEERR = 10000
Public Const WSAEFAULT = (WSABASEERR + 14)
Public Const WSAEINVAL = (WSABASEERR + 22)
Public Const WSAEINPROGRESS = (WSABASEERR + 50)
Public Const WSAENETDOWN = (WSABASEERR + 50)
Public Const WSASYSNOTREADY = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Public Const WSANOTINITIALISED = (WSABASEERR + 93)
Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004

Public Const NCBASTAT As Long = &H33
Public Const NCBNAMSZ As Long = 16
Public Const HEAP_ZERO_MEMORY As Long = &H8
Public Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Public Const NCBRESET As Long = &H32

Public Const MaxConnections As Long = 20       'Maximum number of connections allowed. Causes errors when dynamically re-dimensioned.
Public Const cNetNone As Byte = 0
Public Const cNetHost As Byte = 1
Public Const cNetClient As Byte = 2
Public Const cBannedListFile As String = "\BannedList.cpt"

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

'Used in NetRegDataType
Public Type AcknowledgeType
    RetryCount      As Long
    RemoteCommand   As Byte
    WinType         As Byte
    BytBuf()        As Byte
End Type

Public Type NetRegDataType
    RegCode         As String
    HostIP          As String
    HostName        As String
    ValidPassword   As Boolean
    PasswordTrys    As Long
    AppVersion      As String
    HostID          As String
    VotesAgainst    As String
    LeKey           As Long
    LeSlot          As Long
    LeSlotSpin      As Long
    LeType          As Long
    AcknowledgeMessage  As AcknowledgeType
End Type

Public Type WSAData
    wVersion       As Integer
    wHighVersion   As Integer
    szDescription  As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpVendorInfo   As Long
End Type

Public Type HOSTENT
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type

Public Type NET_CONTROL_BLOCK  'NCB
    ncb_command    As Byte
    ncb_retcode    As Byte
    ncb_lsn        As Byte
    ncb_num        As Byte
    ncb_buffer     As Long
    ncb_length     As Integer
    ncb_callname   As String * NCBNAMSZ
    ncb_name       As String * NCBNAMSZ
    ncb_rto        As Byte
    ncb_sto        As Byte
    ncb_post       As Long
    ncb_lana_num   As Byte
    ncb_cmd_cplt   As Byte
    ncb_reserve(9) As Byte ' Reserved, must be 0
    ncb_event      As Long
End Type

Public Type ADAPTER_STATUS
    adapter_address(5) As Byte
    rev_major         As Byte
    reserved0         As Byte
    adapter_type      As Byte
    rev_minor         As Byte
    duration          As Integer
    frmr_recv         As Integer
    frmr_xmit         As Integer
    iframe_recv_err   As Integer
    xmit_aborts       As Integer
    xmit_success      As Long
    recv_success      As Long
    iframe_xmit_err   As Integer
    recv_buff_unavail As Integer
    t1_timeouts       As Integer
    ti_timeouts       As Integer
    Reserved1         As Long
    free_ncbs         As Integer
    max_cfg_ncbs      As Integer
    max_ncbs          As Integer
    xmit_buf_unavail  As Integer
    max_dgram_size    As Integer
    pending_sess      As Integer
    max_cfg_sess      As Integer
    max_sess          As Integer
    max_sess_pkt_size As Integer
    name_count        As Integer
End Type
   
Public Type NAME_BUFFER
    name        As String * NCBNAMSZ
    name_num    As Integer
    name_flags  As Integer
End Type

Public Type ASTAT
    adapt          As ADAPTER_STATUS
    NameBuff(30)   As NAME_BUFFER
End Type

Public Type networkSettings
    ClientName(MaxConnections) As String            'Client name.
    setupControlChange As Boolean                   'True if setup values have changed
    playerOwner(5) As Byte                          'Which terminal controlls this player
    Controller(5) As Byte                           'Determine if human controlled for stats. 0 for human.
    pctInfoByt() As Byte                            'Coded text in pctInfo
    changeList() As Byte                            'List of changes made during the game
    RolledAttackDice(cMaxNumberOfDice - 1) As Byte  'Last dice roll.
    RolledDefenceDice(cMaxNumberOfDice - 1) As Byte
    xmitInfTxt As Boolean                           'Xmit info box text if true
    madeUpdate As Boolean                           'True if update has been made
    highestPriority As Long                         'Highest priority of update
    cardsPickedPos(2) As Byte                       'Position of each card picked
    redrawCards As Boolean                          'Need to redraw cards at next update
    reshuffleCards As Boolean                       'True if reshuffle cards
    LockControls As Boolean                         'Lock controls to prevent callbacks.
End Type

Public gHomeWebPage As String
Public gHelpWebPage As String
Public gDownloadWebPage As String
Public gIndexServerUrl As String
Public MyNetRegData As NetRegDataType
Public RemoteNetRegData(MaxConnections) As NetRegDataType           'List of reg codes of players
Public net As networkSettings
Public netWorkSituation As Byte             '0=no connections, 1=host, 2=client
Public myTerminalNumber As Byte                     'Network player number
Public sndComplete(MaxConnections) As Byte

'Return the remote terminal marker which defaults to a colon ":". The default
'is set in the global constant gcRemoteTerminalMarker and can be overridden by
'the registry value at key "RemoteTerminalMarker". The remote terminal marker
'is displayed in the Network Admin Panel and in the chat box.
Public Function GetRemoteTerminalMarker() As String
    GetRemoteTerminalMarker = GetSetting(gcApplicationName, _
                            "settings", "RemoteTerminalMarker", gcRemoteTerminalMarker)
End Function

'Return the IP that looks valid. Priority going to pIP1. This is used to overcome a bug
'with the winsock control which occasionally returns either a partial IP address or no
'IP address at all.
Public Function ChooseValidIP(pIP1 As String, Optional pIP2 As String, Optional pIP3 As String)
    If IsValidIP(pIP1) Then
        ChooseValidIP = pIP1
    ElseIf IsValidIP(pIP2) Then
        ChooseValidIP = pIP2
    ElseIf IsValidIP(pIP3) Then
        ChooseValidIP = pIP3
    Else
        ChooseValidIP = pIP1
        Call LogError("", "ChooseValidIP(""" _
                & pIP1 & """,""" _
                & pIP2 & """,""" _
                & pIP3 & """): Warning returned " _
                & ChooseValidIP)
    End If
End Function

'Return true if the passed IP address is a valid IP address format (#.#.#.#).
Public Function IsValidIP(pIP As String) As Boolean
    Dim i As Long
    Dim vOcts() As String
    
    On Error GoTo ErrHand
    
    'Split octets on dots (#.#.#.#)
    vOcts = Split(pIP, ".")
    
    'Check that there are 4 octets.
    If UBound(vOcts) <> 3 Then
        IsValidIP = False
        Exit Function
    End If
    
    'Check octet is numeric and less than or equal to 255.
    For i = 0 To 3
        If Not IsNumeric(vOcts(i)) Then
            IsValidIP = False
            Exit Function
        ElseIf CLng(vOcts(i)) > 255 Then
            IsValidIP = False
            Exit Function
        End If
    Next
    
    'Fall through to here means all is good.
    IsValidIP = True
    
    Exit Function
ErrHand:
    IsValidIP = False
    Exit Function
End Function

'Return Time Stamp in seconds acurate to 3 decimal places.
'Return format: <seconds>.<milli seconds>
'EG: 751066.578
Public Function GetTimeStamp() As String
    GetTimeStamp = CStr(CDbl(GetTickCount) / 1000)
End Function

'The function takes a Double containing a value in the 
'range of an unsigned Long and returns a Long that you 
'can pass to an API that requires an unsigned Long.
Public Function UnsignedToLong(Value As Double) As Long
    If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If
End Function

'The function takes an unsigned Long from an API and 
'converts it to a Double for display or arithmetic purposes.
Public Function LongToUnsigned(Value As Long) As Double
    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If
End Function

'The function takes a Long containing a value in the range 
'of an unsigned Integer and returns an Integer that you 
'can pass to an API that requires an unsigned Integer.
Public Function UnsignedToInteger(Value As Long) As Integer
    If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If
End Function

'The function takes an unsigned Integer from and API and 
'converts it to a Long for display or arithmetic purposes.
Public Function IntegerToUnsigned(Value As Integer) As Long
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
End Function

'Convert a string pointer to a string.
Public Function StringFromPointer(ByVal lPointer As Long) As String
    Dim strTemp As String
    Dim lRetVal As Long
    strTemp = String$(lstrlen(ByVal lPointer), 0)
    lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)
    If lRetVal Then StringFromPointer = strTemp
End Function

'Return host name of this terminal.
Public Function GetLocalHostName() As String
    Dim udtWinsockData  As WSAData
    Dim lngPtrToHOSTENT As Long
    Dim udtHostent      As HOSTENT
    Dim lngPtrToIP      As Long
    Dim arrIpAddress()  As Byte
    Dim strIpAddress    As String
    Dim strHostName     As String * 256
    Dim lngRetVal       As Long
    Dim i               As Long
    
    lngRetVal = GetHostName(strHostName, 256)
    If lngRetVal = SOCKET_ERROR Then
        netMain.WriteText Err.LastDllError, True
        Exit Function
    End If
    
    GetLocalHostName = Trim$(Left(strHostName, InStr(1, strHostName, Chr(0)) - 1))
End Function

'Return comma delimeted list of IP addresses.
'"x.x.x.x,y.y.y.y,z.z.z.z"
'Use the following to convert into an array of IP addresses -
'   Dim IPAdress() As String
'   IPAdress = Split(modNetwork.GetLocalHostIP, ",")
Public Function GetLocalHostIP() As String
    Dim udtWinsockData  As WSAData
    Dim lngPtrToHOSTENT As Long
    Dim udtHostent      As HOSTENT
    Dim lngPtrToIP      As Long
    Dim arrIpAddress()  As Byte
    Dim strIpAddress    As String
    Dim strHostName     As String * 256
    Dim lngRetVal       As Long
    Dim i               As Long
    
    'start up winsock service
    lngRetVal = WSAStartup(&H101, udtWinsockData)
    If lngRetVal <> 0 Then
        netMain.WriteText "Socket Error", True
        Exit Function
    End If

    'Get the local host name
    lngRetVal = GetHostName(strHostName, 256)
    If lngRetVal = SOCKET_ERROR Then
        netMain.WriteText Err.LastDllError, True
        Exit Function
    End If
    
    'Call the gethostbyname Winsock API function
    GetLocalHostIP = ""
    lngPtrToHOSTENT = gethostbyname(Trim$(Left(strHostName, InStr(1, strHostName, Chr(0)) - 1)))
    If lngPtrToHOSTENT = 0 Then
        netMain.WriteText Err.LastDllError, True
    Else
        RtlMoveMemory udtHostent, lngPtrToHOSTENT, LenB(udtHostent)
        RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
        Do Until lngPtrToIP = 0
            ReDim arrIpAddress(udtHostent.hLength - 1)
            RtlMoveMemory arrIpAddress(0), lngPtrToIP, udtHostent.hLength
            For i = 0 To udtHostent.hLength - 1
                strIpAddress = strIpAddress & arrIpAddress(i) & "."
            Next
            strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
            GetLocalHostIP = GetLocalHostIP & strIpAddress & ","
            strIpAddress = ""
            udtHostent.hAddrList = udtHostent.hAddrList + LenB(udtHostent.hAddrList)
            RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
         Loop
         
         'Remove last comma.
         GetLocalHostIP = Left(GetLocalHostIP, Len(GetLocalHostIP) - 1)
    End If
    
    Call WSACleanup
End Function


