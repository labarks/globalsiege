Attribute VB_Name = "modNetIxServer"
Option Explicit

' The format of the command line arguments passwd by the URL to the Indexing Server.
'define("arg_COMMAND",0);
'define("arg_SESION_ID",1);
'define("arg_HOST_IP",2);
'define("arg_TCP_PORT",3);
'define("arg_UDP_PORT",4);
'define("arg_CRYPT_KEYS",5);
'define("arg_LOCKED", 6);
'define("arg_HIDDEN", 7);
'define("arg_SES_NAME",8);
'define("arg_TERM_NAME",9);
'define("arg_GS_VERSION",10);
'define("arg_UNIQ_ID",11);
'define("arg_GS_CHECKSUM",12);
'define("arg_GS_TIME_STAMP",13);
'define("arg_GS_USERNAME",14);
'define("arg_GS_HASHCODE",15);
'define("arg_LOCAL_SES_ID",16);
'define("arg_COMMENT",17);

'Enumerate the commands that can be executed by the Indexing Server. These
'commands are passed as a number at the first position in the arg list.
Public Enum eIxCommand
    DO_NOTHING = 0
    MY_IP = 1
    NEW_SESSION = 2
    CLOSE_SESSION = 3
    CLEANUP_SESSION = 4
    KEEP_SESS_ALIVE = 5
    OPTIONS_CHANGED = 6
    LIST_SESSIONS = 7
    COMMENT = 8
    BANNED_PLAYER = 9
    AGREEMENT_ACK = 10
    JOINED_SESSION = 11
    LEFT_SESSION = 12
    UPLOAD_STATS = 13
End Enum

Public Type InetSessionType
    IxServerClientUrl  As String
    ID              As String
    RemoteSessID    As String
    LocalSessID     As String
    Tstamp          As String
    IP              As String
    ReportedIP      As String
    IPList          As String
    LocalTcpPort    As String
    LocalUdpPort    As String
    RemoteTcpPort   As String
    RemoteUdpPort   As String
    RemoteIP        As String
    SesName         As String
    HostName        As String
    RebroadcastCntr As Long
End Type

Public Type InternetType
    ServerIndexAddr     As String
    LastCheckTime       As Single
    Locked              As Boolean
    Connected           As Boolean
    SessionName         As String
    SessionID           As String
    LastAction          As Long
    BannedList          As String
    KilledList          As String
    IxServerPingTime    As String
End Type

Public InetSes As InetSessionType
Public Inet As InternetType
Const vQueueGetPostDelimiter As String = "|| Post part -->>"

'------------------------------------------------------------------------------------------------------
'       Public functions
'------------------------------------------------------------------------------------------------------

'Check the website for any GlobalSiege news such as a new version released etc.
'If it fails, do not try again until the next time GlobalSiege is started.
Public Function IxServerCheckGsNews(Optional pSilent As Boolean = True) As Boolean
    Dim vPageText As String
    
    On Error Resume Next
    
    If Not netInetGateway.IsStillExecuting Then
        vPageText = IxServerComunication(gcNewsServerURL, "", pSilent)
        vPageText = IxLedServerDirectives(vPageText)
        Call IxProcessServerDirectives(vPageText)
    End If
End Function

'Contact the Indexing Server and get my IP address and check my login.
Public Function IxServerCheckLogin(Optional pSilent As Boolean = False) As Boolean
    Dim vListIpResponse As String
    Dim vHashCode As String
    
    On Error GoTo ErrHand
    IxServerCheckLogin = False
    
    'Contact the server and get IP addresses comma delimited for this machine as seen by the web.
    vListIpResponse = IxServerExecuteCommand(eIxCommand.DO_NOTHING)
                
    'Exit if there is an error connecting to the server.
    If InStr(1, UCase(netInetGateway.GetErrorText), "ERROR") > 0 _
    Or InStr(1, UCase(IxGetServerDirective(vListIpResponse, "STATUS")), "ERROR") > 0 Then
        LogError "IxServerPostNewSession", "Error connecting to the IX server LIST_MY_IP: " _
                    & netInetGateway.GetErrorText & ": " & vListIpResponse
        netMain.WriteText "Unable to contact the Indexing Server. " & netInetGateway.GetErrorText, True
        netMain.WriteText "Try again in a few minutes.", True
        IxServerCheckLogin = False
        Exit Function
    End If
    
    vHashCode = gGsLeUtils.LE6d(IxGetServerDirective(vListIpResponse, "HASH"))
    
    'Check if the command was queued or not.
    If IxGetServerDirective(vListIpResponse, "QUEUED_AT") <> "" Then
        
        'Command queued, will finish the session posting later.
        IxServerCheckLogin = True
        LogError "IxServerCheckLogin", "Request to IxServer has been queued."
    
    ElseIf vHashCode = "" Then
        
        'If the hash code is blank.
        IxServerCheckLogin = False
    
    Else
        'All good, return TRUE.
        IxServerCheckLogin = True
    End If
    
    Call IxProcessServerDirectives(vListIpResponse, pSilent)
    
    Exit Function
ErrHand:
    netMain.WriteText "Error - " & Err.Description, True
    LogError "IxServerCheckLogin", " Error: " & Err.Number & " " & Err.Description & " - " & vListIpResponse
    IxServerCheckLogin = False
    Exit Function
End Function

'Contact the Indexing Server and get my IP address
'and notify that the T&C has been agreed to.
Public Function IxServerTcAgreed(Optional pSilent As Boolean = False) As Boolean
    Dim vListIpResponse As String
    
    On Error GoTo ErrHand
    IxServerTcAgreed = False
    
    'Contact the server and get IP addresses comma delimited for this machine as seen by the web.
    vListIpResponse = IxServerExecuteCommand(eIxCommand.AGREEMENT_ACK)
                
    'Exit if there is an error connecting to the server.
    If InStr(1, UCase(netInetGateway.GetErrorText), "ERROR") > 0 _
    Or InStr(1, UCase(IxGetServerDirective(vListIpResponse, "STATUS")), "ERROR") > 0 Then
        LogError "IxServerTcAgreed", "Error connecting to the IX server LIST_MY_IP: " _
                    & netInetGateway.GetErrorText & ": " & vListIpResponse
        netMain.WriteText "Unable to contact the Indexing Server. " & netInetGateway.GetErrorText, True
        netMain.WriteText "Try again in a few minutes.", True
        IxServerTcAgreed = False
        Exit Function
    End If
    
    'Check if the command was queued or not.
    If IxGetServerDirective(vListIpResponse, "QUEUED_AT") <> "" Then
        
        'Command queued, will finish the session posting later.
        IxServerTcAgreed = True
        LogError "IxServerTcAgreed", "Request to IxServer has been queued."
    
    Else
        'All good, return TRUE.
        IxServerTcAgreed = True
    End If
    
    Call IxProcessServerDirectives(vListIpResponse, pSilent)
    
    Exit Function
ErrHand:
    netMain.WriteText "Error - " & Err.Description, True
    LogError "IxServerTcAgreed", " Error: " & Err.Number _
                & " " & Err.Description & " - " & vListIpResponse
    IxServerTcAgreed = False
    Exit Function
End Function

'If I am a Session Host, upload end of war stats.
'The format of pIxStatsReport should be:
'gGsLeUtils.LE6(Controlling Terminal),vCountriesIBeat,vCountriesILost,vUnitsIBeat,vUnitsILost,vScore <CR>
'gGsLeUtils.LE6(Controlling Terminal),vCountriesIBeat,vCountriesILost,vUnitsIBeat,vUnitsILost,vScore <CR>
' ...
Public Sub IxServerUploadStats(pIxStatsReport)
    Dim vResponse As String
    Dim vPost As String
    
    If netWorkSituation = cNetHost And netMain.optInet.Value Then
        vPost = "session_stats=" & Trim(pIxStatsReport)
        'vResponse = IxServerExecuteCommand(eIxCommand.UPLOAD_STATS, InetSes.ID, , , , , , , , , vPost, True)
        vResponse = IxServerExecuteCommand(eIxCommand.UPLOAD_STATS _
                    , InetSes.ID _
                    , InetSes.IP _
                    , InetSes.LocalTcpPort _
                    , InetSes.LocalUdpPort _
                    , _
                    , netMain.chkPasswordSession.Value _
                    , netMain.chkHideSession.Value _
                    , InetSes.SesName _
                    , _
                    , vPost)
    End If
End Sub

'Tell the Indexing Server that I have joined a session. Force to queue
'because there is no rush with this data and it is not important enough
'to notify the player if it fails.
Public Sub IxServerJoinSession(pSessionID As String)
    Dim vResponse As String
    Dim vPost As String
    
    If netMain.optInet.Value Then
        vPost = "session_id=" & Trim(pSessionID)
        vResponse = IxServerExecuteCommand(eIxCommand.JOINED_SESSION, , , , , , , , , , vPost, True)
    End If
End Sub

'Tell the Indexing Server that I have left a session. Force to queue
'because there is no rush with this data and it is not important enough
'to notify the player if it fails.
Public Sub IxServerLeftSession(pSessionID As String)
    Dim vResponse As String
    Dim vPost As String
    
    If netMain.optInet.Value Then
        vPost = "session_id=" & Trim(pSessionID)
        vResponse = IxServerExecuteCommand(eIxCommand.LEFT_SESSION, , , , , , , , , , vPost, False)
    End If
End Sub

'Try to post a new session on the Session Indexing Server. Return TRUE if successful.
'This is the forst part, request the IP address as seen by the Indexing Server.
Public Function IxServerPostNewSession(Optional pSilent As Boolean = False) As Boolean
    Dim vListIpResponse As String
    Dim vDisplayName As String
    
    On Error GoTo ErrHand
    IxServerPostNewSession = False
    
    'Make sure the control change option timer is off.
    netMain.tmrControlChange.Enabled = False
    
    InetSes.LocalTcpPort = netMain.sckTCP(0).LocalPort
    InetSes.LocalUdpPort = netFindHosts.sckUDP.LocalPort
    InetSes.IP = netMain.sckTCP(0).LocalIP             'Not internet IP!!
    InetSes.HostName = netMain.sckTCP(0).LocalHostName
    
    'Ensure there is a session name chosen.
    netMain.txtSesName.Text = Trim(netMain.txtSesName.Text)
    If Len(netMain.txtSesName.Text) = 0 Then
        netMain.txtSesName.Text = netMain.PickRandomSesname
    End If
    InetSes.SesName = netMain.txtSesName.Text
    DoEvents
    
    'Contact the server and get IP addresses comma delimited for this machine as seen by the web.
    If Not pSilent Then
        netMain.WriteText "Connecting to the war server...", True
    End If
    vListIpResponse = IxServerExecuteCommand(eIxCommand.MY_IP)
     
    'Exit if there is an error connecting to the server.
    If InStr(1, UCase(netInetGateway.GetErrorText), "ERROR") > 0 _
    Or InStr(1, UCase(IxGetServerDirective(vListIpResponse, "STATUS")), "ERROR") > 0 Then
        LogError "IxServerPostNewSession", "Error connecting to the IX server LIST_MY_IP: " _
                    & netInetGateway.GetErrorText & ": " & vListIpResponse
        netMain.WriteText "Unable to contact the Indexing Server. " & netInetGateway.GetErrorText, True
        netMain.WriteText "Try again in a few minutes.", True
        IxServerPostNewSession = False
        Exit Function
    End If
    
    'Check if the command was queued or not.
    If IxGetServerDirective(vListIpResponse, "QUEUED_AT") <> "" Then
        
        'Command queued, will finish the session posting later.
        IxServerPostNewSession = True
        LogError "IxServerPostNewSession", "Request to IxServer has been queued."
        
    ElseIf IxGetServerDirective(vListIpResponse, "STATUS") = "LoginFail" Then
        LogError "IxServerPostNewSession", "Login failed for user " & Trim(netMain.txtUserName.Text)
        netMain.WriteText "Please login."
        IxServerPostNewSession = False
        netIxServerLogin.SetCallback = "PostInetHostDetails"
        netIxServerLogin.Show , netMain
        
    'Check if the T&C have been signed.
    ElseIf IxGetServerDirective(vListIpResponse, "TC_NEEDS_ACK") <> "" Then
        LogError "IxServerPostNewSession", "T&C not signed by user " & Trim(netMain.txtUserName.Text)
        IxServerPostNewSession = False
        netIxServerAgreement.SetCallback = "PostInetHostDetails"
        netIxServerAgreement.SetAgreementUrl = IxGetServerDirective(vListIpResponse, "TC_NEEDS_ACK")
        netMain.WriteText "Please agree to the Terms and Conditions for online use.", False
        
        'Show the agreement form and set the owner form to one that is visible.
        If netMain.Visible Then
            netIxServerAgreement.Show , netMain
        Else
            netIxServerAgreement.Show , TheMainForm
        End If
    
    Else
        
        'All good, go to the next stage.
        vDisplayName = Trim(gGsLeUtils.LE6d(IxGetServerDirective(vListIpResponse, "DISPLAY_NAME")))
        If vDisplayName <> "" Then
            netMain.txtTerminalName.Text = vDisplayName
        End If
        IxServerPostNewSession = IxServerPostNewHaveIp(vListIpResponse, pSilent, True)
    End If
    
    Exit Function
ErrHand:
    LogError "IxServerPostNewSession", "Error: " & Err.Number & " " & Err.Description & " - " & vListIpResponse
    netMain.WriteText "Error - " & Err.Description, True
    IxServerPostNewSession = False
    Exit Function
End Function

'Create an IP list by adding the local IP to the IP address from the Indexing Server.
'IP addresses are delimited by "x" instead of by comma. The IX IP address must go
'first incase udp responses are blocked.
Private Function IxServerCreateIpList(pIpList As String) As String
    Dim vFormatIpList As String
    
    'Get list of IP addresses from the Indexing Server.
    vFormatIpList = IxGetServerDirective(pIpList, "YOUR_IP")
    
    'Add the local IP address.
    vFormatIpList = vFormatIpList & "x" & modNetwork.GetLocalHostIP
    
    'Replace commas with 'x' because the Listbox index uses commas to seperate fields.
     vFormatIpList = Replace(vFormatIpList, ",", "x")
    
    'Clean up the IP list by removeing duplicates.
    IxServerCreateIpList = CleanList(vFormatIpList, "x")
End Function

'The second part of posting a new war. We should have the IP address asreported
'by the Indexing Server. Now try to post the session. pProcessDirectives is to
'stop circular calls if this was from the Index Server Command Processor.
Private Function IxServerPostNewHaveIp(pListIpResponse As String, _
Optional pSilent As Boolean = False, _
Optional pProcessDirectives As Boolean = False) As Boolean
    Dim vNewSesResponse As String
    Dim vLines() As String
    
    'Make sure the control change option timer is off.
    netMain.tmrControlChange.Enabled = False
    
    'Check that the Index Server has given the correct response.
    If IxGetServerDirective(pListIpResponse, "STATUS") <> "Ready_For_Post" Then
        If Not pSilent Then
            Call netMain.WriteText("Warning: The response from the Index Server is garbled.")
            LogError "IxServerPostNewHaveIp", _
                    "Warning: The response from the Index Server is garbled - " _
                    & pListIpResponse
        End If
        
    End If
    
    'Get a list of IP addresses from the Indexing Server.
    InetSes.IP = IxServerCreateIpList(pListIpResponse)
    
    If Not pSilent Then
        netMain.WriteText "War server contacted. Starting a new session...", True
    End If
    
    'Post the new session on the Indexing Server.
    With MyNetRegData
    vNewSesResponse = IxServerExecuteCommand(eIxCommand.NEW_SESSION _
                , _
                , InetSes.IP _
                , InetSes.LocalTcpPort _
                , InetSes.LocalUdpPort _
                , .LeType _
                    & ":" & .LeKey _
                    & ":" & .LeSlot _
                    & ":" & .LeSlotSpin _
                , netMain.chkPasswordSession.Value _
                , netMain.chkHideSession.Value _
                , InetSes.SesName _
                , _
                , "Welcome_Msg=" & EncodeNonAscii(netMain.txtWelcomeMsg.Text))
    End With
    
    'Exit if error connecting to the site.
    If InStr(1, netInetGateway.GetErrorText, "Error") > 0 _
    Or InStr(1, UCase(IxGetServerDirective(vNewSesResponse, "STATUS")), "ERROR") > 0 Then
        LogError "IxServerPostNewHaveIp", "Error connecting to the server - " & vNewSesResponse
        netMain.WriteText netInetGateway.GetErrorText, True
        netMain.WriteText "Try again in a few minutes.", True
        IxServerPostNewHaveIp = False
        Exit Function
    End If
    
    'Check if the command was queued or not.
    If IxGetServerDirective(vNewSesResponse, "QUEUED_AT") <> "" Then
        
        'Command queued, will finish the session posting later.
        IxServerPostNewHaveIp = True
        LogError "IxServerPostNewHaveIp", "Request to IxServer has been queued."
        
    Else
        
        'All good, go to the next stage.
        IxServerPostNewHaveIp = IxServerPostNewSesPosted(vNewSesResponse, pSilent, True)
    End If
    
    Exit Function
ErrHand:
    LogError "IxServerPostNewHaveIp", "Error: " & Err.Number & " " & Err.Description & " - " & vNewSesResponse
    netMain.WriteText "Error - " & Err.Description, True
    IxServerPostNewHaveIp = False
    Exit Function
End Function

'The third part of posting a new war. The session should have been posted now.
'Let the user know and get the session ID. pProcessDirectives is to
'stop circular calls if this was from the Index Server Command Processor.
Private Function IxServerPostNewSesPosted(pNewSesResponse As String, _
Optional pSilent As Boolean = False, _
Optional pProcessDirectives As Boolean = False) As Boolean
    Dim vDirective As String
    
    InetSes.ID = IxGetServerDirective(pNewSesResponse, "SESSION_ID")
    InetSes.Tstamp = DecodeNonAscii(IxGetServerDirective(pNewSesResponse, "GS_TIME"))
    
    'Was the session posted?
    If Len(InetSes.ID) > 0 And Len(InetSes.Tstamp) > 0 And Len(InetSes.IP) > 0 Then
        
        'Session was posted.
        IxServerPostNewSesPosted = True
        netMain.chkHideSession.Enabled = True
        netMain.chkHideSession.Value = vbUnchecked
    Else
        
        'Something went wrong.
        LogError "IxServerPostNewSesPosted", "NEW_SESSION response is badly formatted - " & pNewSesResponse
        netMain.WriteText "Error posting session: server error.", True
        netMain.WriteText "Try again in a few minutes.", True
        IxServerPostNewSesPosted = False
    End If
    
    'Execute any extra directives from the host.
    If pProcessDirectives Then
        Call IxProcessServerDirectives(pNewSesResponse, pSilent)
    End If
    Exit Function
ErrHand:
    LogError "IxServerPostNewSesPosted", "Error: " & Err.Number & " " & Err.Description & " - " & pNewSesResponse
    netMain.WriteText "Error - " & Err.Description, True
    IxServerPostNewSesPosted = False
    Exit Function
End Function

'Actions to to disconnect. If I am the host then close all.
Public Function IxServerCloseSession(Optional pReason As eIxCommand = eIxCommand.CLOSE_SESSION, _
Optional pAskedBefore As Boolean = False) As Boolean
    Dim vSocketIndex As Long
    Dim vResponse As String
    
    'Do you want to close all connections?
    If netWorkSituation <> cNetNone _
    And Not pAskedBefore _
    And Not gServerMode _
    And Not gHeadlessMode Then
        If MsgBox(Phrase(239), vbYesNo, Phrase(242)) = vbNo Then
            Exit Function
        End If
    End If
    
    On Error GoTo ErrHand
    
    'Make sure the control change option timer is off.
    netMain.tmrControlChange.Enabled = False
    
    'Ensure the broadcasting time is off, which which
    'should already be the case for a host.
    netFindHosts.tmrBroadcast.Enabled = False
    
    'Internet host must close the posted session.
    If netMain.optHost.Value And netMain.optInet.Value Then
        netMain.tmrKeepAlive.Enabled = False
        
        'If there is no session ID then there mustn't be an active session to close.
        If Trim(InetSes.ID) <> "" Then
            netMain.WriteText "Closing session...", True
            
            'Remove the session from the Indexing Server.
            vResponse = IxServerExecuteCommand(pReason _
                    , InetSes.ID _
                    , InetSes.IP _
                    , InetSes.LocalTcpPort _
                    , InetSes.LocalUdpPort _
                    , _
                    , netMain.chkPasswordSession.Value _
                    , netMain.chkHideSession.Value _
                    , InetSes.SesName)
            
            'Check the results.
            If IxGetServerDirective(vResponse, "SESSION_REMOVED") = InetSes.ID Then
                
                'Session successfully removed.
                InetSes.ID = ""
                netMain.tmrKeepAlive.Tag = ""
                IxServerCloseSession = True
            ElseIf InStr(1, netInetGateway.GetErrorText, "Error") > 0 Then
                
                'Some error suffered by the Internet control.
                LogError "IxServerCloseSession", "Error connecting to the IX server. " & pReason & ": " & vResponse
                IxServerCloseSession = False
            ElseIf IxGetServerDirective(vResponse, "QUEUED_AT") <> "" Then
                
                'netInetGateway was busy, the command has been queued. netInetGateway.tmrInet1Queue is
                'responsible for processing the command queue.
                netMain.WriteText "The Internet gateway is busy, close request has been queued.", True
                LogError "IxServerCloseSession", "Request to the IxServer has been queued."
                IxServerCloseSession = False
            Else
                
                'Unexpected response.
                LogError "IxServerCloseSession", "Unexpected response from the IX server. " _
                        & pReason & ": " & vResponse
                IxServerCloseSession = False
            End If
            
            'netMain.WriteText "Session closed.", True
            Call IxProcessServerDirectives(vResponse, False)
        End If
        
        'Reset the Hide checkbox.
        netMain.chkHideSession.Enabled = False
        netMain.chkHideSession.Value = vbUnchecked
    
    'Internet clients should notify the Indexing Server that they have left a session.
    ElseIf netMain.optJoin.Value _
    And netMain.optInet.Value _
    And netMain.sckTCP(0).State <> sckClosed Then
        netMain.WriteText "Disconnecting..."
        Call IxServerLeftSession(InetSes.RemoteSessID)
        netMain.WriteText "Connection closed."
    End If
    
    'Save the IP address if one was entered manually to connect.
    InetSes.RemoteIP = ""
    
    'Reset players in the setup screen.
    Call TheMainForm.resetPlayerOwners
    netMain.sckTCP(0).Close
    netFindHosts.sckUDP.Close
    netMain.sckTCP(0).RemoteHost = ""
    sndComplete(0) = 0
    netMain.optJoin.Value = True
    TheMainForm.mnuNetDisconnect.Enabled = False
    
    netMain.txtSesName.Text = Trim(GetSetting(gcApplicationName, "settings", "LastSesName", ""))
    
    'Rename buttons etc and place into join mode.
    Call netMain.OptJoinClick
    
    DoEvents
    
    'Change the encryption keys.
    MyNetRegData.LeKey = Int(GenRandom4 * &H7D) + 1
    MyNetRegData.LeSlot = Int(GenRandom4 * &H7D) + 1
    MyNetRegData.LeSlotSpin = Int(GenRandom4 * &H7D) + 1
    
    If netMain.CountTerminals = 0 Then
        'netMain.WriteText Phrase(246), False  'Connection closed.
        Call netMain.enableButtons(True)
        Call netMain.EnableInternetOptions(True)
        netWorkSituation = cNetNone
        
        'Ensure "Remote player" in the setup screen player list.
        Call TheMainForm.ResetPlayerList
        Call TheMainForm.EnableSetupControls(True)
        Call TheMainForm.resetPlayerOwners
        Exit Function
    End If
    
    'I am host, I still have more stuff to close.
    For vSocketIndex = 1 To MaxConnections
        If (netMain.sckTCP(vSocketIndex).State <> sckClosed) Then
            netMain.sckTCP(vSocketIndex).Close
            netMain.sckTCP(vSocketIndex).Tag = ""
            netMain.sckTCP(vSocketIndex).RemoteHost = ""
        End If
        sndComplete(vSocketIndex) = 0
        DoEvents
        'Unload sckTCP(vSocketIndex)
        RemoteNetRegData(vSocketIndex).RegCode = ""
        RemoteNetRegData(vSocketIndex).HostIP = ""
        RemoteNetRegData(vSocketIndex).HostName = ""
        RemoteNetRegData(vSocketIndex).ValidPassword = False
        RemoteNetRegData(vSocketIndex).PasswordTrys = 0
        RemoteNetRegData(vSocketIndex).AppVersion = ""
        RemoteNetRegData(vSocketIndex).HostID = ""
        RemoteNetRegData(vSocketIndex).VotesAgainst = ""
        
    Next
    netWorkSituation = cNetNone
    Call TheMainForm.ResetPlayerList
    Call TheMainForm.EnableSetupControls(True)
    Call TheMainForm.resetPlayerOwners
    Call netMain.WriteText(Phrase(247) & vbCrLf, False)  'All connections closed.
    Call netMain.enableButtons(True)
    Call netMain.EnableInternetOptions(True)
    Exit Function
ErrHand:
    Resume Next
End Function

'Commands sent to the Index Server while netInetGateway is busy are queued to
'netInetGateway.gGatewayQueue using function IxServerQueueCommand. The timer fires
'every second and will call function IxServerProcessCommandQueue as soon as
'netInetGateway becomes available. The format of the queue is:
'<URL>?<GET>:<POST>
Private Sub IxServerQueueCommand(pUrlAndGet As String, pPost As String)
    With netInetGateway
    .gGatewayQueue = .gGatewayQueue & vbCrLf & pUrlAndGet & vQueueGetPostDelimiter & pPost
    .gGatewayQueue = CleanList(.gGatewayQueue, vbCrLf)
    .tmrInet1Queue.Enabled = True
    End With
End Sub

'Clear the queue if there were problems connecting to the Index Server.
Private Sub IxServerFlushQueue()
    netInetGateway.gGatewayQueue = ""
    netInetGateway.tmrInet1Queue.Enabled = False
End Sub

'If Inet1 was busy during an attempt to contact the Index Server, the URL string
'gets added to netInetGateway.gGatewayQueue and will attempt to send at a later time.
'This gets called from netInetGateway.tmrInet1Queue_timer() event. The reason the
'timer is there is because you cannot have a timer attached to a module.
'The format of the queue is:
'<URL>?<GET>:<POST>
'<URL>?<GET>:<POST>
'...
'<URL>?<GET>:<POST>
Public Sub IxServerCheckCommandQueue()
    Dim vCommandQueue As String
    
    On Error Resume Next
    
    With netInetGateway
    
    'Check if the Internet control is ready.
    If Not .IsStillExecuting Then
        
        'We are clear to send data.
        .tmrInet1Queue.Enabled = False
        
        'Save the queue in a local variable and empty the queue. This is to
        'help catch newly queued commands while using the Internet control.
        vCommandQueue = .gGatewayQueue
        .gGatewayQueue = ""
        
        'Process the queued commands and place before any new commands that
        'arrived in the queue while the Internet control was busy.
        .gGatewayQueue = CleanList(Trim(IxServerProcessCommandQueue(vCommandQueue) _
                        & vbCrLf & .gGatewayQueue), vbCrLf)
        
        'Reset the queue timer if there are still commands in the queue.
        If .gGatewayQueue <> "" Then
            .tmrInet1Queue.Enabled = True
        End If
    End If
    End With
End Sub

'Process commmands waiting in the IxServer command queue. Return any commands
'that were not executed so that they can be placed back in the command queue.
'The format of the queue is:
'<URL>?<GET>:<POST>
'<URL>?<GET>:<POST>
'...
'<URL>?<GET>:<POST>
Public Function IxServerProcessCommandQueue(pCommandQueue As String) As String
    Dim vCommands() As String
    Dim vUrlAndGet As String
    Dim vPost As String
    Dim vReturn As String
    Dim vIndex As Long
    
    On Error Resume Next
    
    'Split the command queue by carrage return.
    vCommands = Split(pCommandQueue, vbCrLf)
    If UBound(vCommands) >= 0 Then
        
        'For each command in the command queue.
        For vIndex = 0 To UBound(vCommands)
            If Trim(vCommands(vIndex)) <> "" Then
                
                'Separate the <URL>?<Get> part from the <POST> part which
                'are delimited by a colon.
                vUrlAndGet = GetListElement(vCommands(vIndex), 0, vQueueGetPostDelimiter)
                vPost = GetListElement(vCommands(vIndex), 1, vQueueGetPostDelimiter)
                
                'Execute the queued command.
                vReturn = IxServerComunication(vUrlAndGet, vPost)
                
                'Decrypt the returned directives from the Indexing Server.
                vReturn = IxLedServerDirectives(vReturn)
                
                'Process the returned directives.
                Call IxProcessServerDirectives(vReturn, False, True)
                
                'Remove the command from the queue. If the command was queued again
                'it will be requeued at the front of netInetGateway.gGatewayQueue.
                vCommands(vIndex) = ""
                
                'Return the other commands in the queue that were not processed.
                IxServerProcessCommandQueue = CleanList(Join(vCommands, vbCrLf), vbCrLf)
                Debug.Print vCommands(vIndex)
                Debug.Print Replace(vReturn, "<br />", "<br />" & vbCrLf)
                Exit For
            End If
        Next
    End If
End Function

'Send a refresh (heartbeat) signal to the Indexing Server with a eIxCommand.KEEP_SESS_ALIVE signal
'to stop the session from going stale and being removed. This function is called
'every minute from netMain.tmrKeepAlive and oly sends a keepalive signal every 9
'minutes. If pReset is set to true, the timer is reset to zero and is called when
'the session is locked or hidden because they update the session time stamp on the
'Index Server as well.
Public Sub IxServerRefreshSession()
    Static vMinutes As Long
    Dim vResponse As String
    Dim vConnectedPlayerList As String
    
    'Make sure listening if posted as host on the internet.
    If netMain.optInet.Value _
    And Trim(InetSes.ID) <> "" _
    And netMain.optHost.Value _
    And netMain.sckTCP(0).State = sckListening Then
        vMinutes = vMinutes + 1
        If vMinutes >= 9 Then
            vMinutes = 0
            
            'Create a list of connected terminals.
            vConnectedPlayerList = "connected_players=" & netMain.ListConnectedTermNames(True)
            
            'Contact the Indexing Server.
            vResponse = IxServerExecuteCommand(eIxCommand.KEEP_SESS_ALIVE _
                    , InetSes.ID _
                    , InetSes.IP _
                    , InetSes.LocalTcpPort _
                    , InetSes.LocalUdpPort _
                    , _
                    , netMain.chkPasswordSession.Value _
                    , netMain.chkHideSession.Value _
                    , InetSes.SesName _
                    , _
                    , vConnectedPlayerList)
            
            'Check if the session was refreshed, start a new session if there was a problem.
            If IxGetServerDirective(vResponse, "SESSION_REFRESHED") = InetSes.ID Then
                
                'All good, session was refreshed.
                Call IxProcessServerDirectives(vResponse, True)
            Else
            
                'Session was not found on the Index Server, quietly start a new one.
                LogError "IxServerRefreshSession", "Session not found on the IxServer, starting a new one. " _
                    & vResponse
                Call IxServerPostNewSession(True)
            End If
        End If
    Else
        
        'Ensure the session is closed and stop the timer.
        Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, True)
        netMain.tmrKeepAlive.Enabled = False
    End If
End Sub

'Notify the Indexing Server of session lock/un-lock, hide/un_hide etc.
'Called from netMain.tmrControlChange. Force to queue because trhere is no rush.
Public Sub IxServerOptionsChanged(Optional pCommand As eIxCommand = eIxCommand.OPTIONS_CHANGED)
    Dim vResponse As String
    Dim vConnectedPlayerList As String
    
    On Error Resume Next
    
    If netMain.optInet.Value _
    And Trim(InetSes.ID) <> "" _
    And netMain.optHost.Value _
    And netMain.sckTCP(0).State = sckListening Then
        
        'Create a list of connected terminals.
        vConnectedPlayerList = "connected_players=" & netMain.ListConnectedTermNames(True)
            
        vResponse = IxServerExecuteCommand( _
                    pCommand _
                    , InetSes.ID _
                    , InetSes.IP _
                    , InetSes.LocalTcpPort _
                    , InetSes.LocalUdpPort _
                    , _
                    , netMain.chkPasswordSession.Value _
                    , netMain.chkHideSession.Value _
                    , InetSes.SesName _
                    , _
                    , vConnectedPlayerList _
                    , True)
        
        'Check for the expected response.
        If IxGetServerDirective(vResponse, "QUEUED_AT") <> "" Then
            
            'Command queued, will finish the session posting later.
            LogError "IxServerOptionsChanged", "Request to IxServer has been queued."
        
        ElseIf IxGetServerDirective(vResponse, "OPTIONS_UPDATED") = "" Then
            
            'Expected response not found.
            LogError "IxServerOptionsChanged", "Cound not find the expected token in the IxServer response. " _
                    & vResponse
        Else
            
            'All good.
            Call IxProcessServerDirectives(vResponse, True)
        End If
        
    End If
End Sub

'Notify the Index Server that the passed player was banned. This information
'will be used to identify nuisance players.
Public Sub IxServerPlayerWasBanned(pBannedPlayer As String, Optional pReason As String = "")
    Dim vResponse As String
    Dim vPost As String
    
    On Error Resume Next
    
    If netMain.optInet.Value _
    And Trim(InetSes.ID) <> "" _
    And netMain.optHost.Value _
    And netMain.sckTCP(0).State = sckListening Then
        
        vPost = "player_banned=" & Trim(pBannedPlayer)
        
        vResponse = IxServerExecuteCommand( _
                    eIxCommand.BANNED_PLAYER _
                    , InetSes.ID _
                    , InetSes.IP _
                    , InetSes.LocalTcpPort _
                    , InetSes.LocalUdpPort _
                    , _
                    , netMain.chkPasswordSession.Value _
                    , netMain.chkHideSession.Value _
                    , InetSes.SesName _
                    , pReason _
                    , vPost _
                    , False)
        
        'Check for the expected response.
        If IxGetServerDirective(vResponse, "QUEUED_AT") <> "" Then
            
            'Command queued, will finish the session posting later.
            LogError "IxServerPlayerWasBanned", "Request to the IxServer has been queued."
        
        ElseIf IxGetServerDirective(vResponse, "BAN_RECORDED") = "" Then
            
            'Expected response not found.
            LogError "IxServerPlayerWasBanned", "Cound lot find the expected token in the IxServer response. " _
                    & vResponse
        Else
            
            'All good.
            Call IxProcessServerDirectives(vResponse, True)
        End If
    End If
End Sub

'Show list and return TRUE if host found.
'Put IP and Port number in text boxes.
'Translate and perform commands from server.
'If Silent is true, nothing is written to the output except error messages.
Public Function IxServerListSession(Optional pSilent As Boolean = False) As Boolean
    Dim vResponse As String
    
    On Error GoTo ErrHand
    
    IxServerListSession = False
    DoEvents
    
    'Not an Internet session.
    If Not netMain.optInet.Value Then
        Exit Function
    End If

    'Response lines "Session_ID, Time_Stamp, IP_Address, Main_Port, Host_UDP_Port, Session_Name, My_IP"
    If Not pSilent Then
        netMain.WriteText "Connecting to the war server...", False
    End If
    
    'Wxecute the LIST command on the Indexing Server for a list of war sessions.
    vResponse = IxServerExecuteCommand(eIxCommand.LIST_SESSIONS)
    
    'Exit if error connecting to site.
    If InStr(1, netInetGateway.GetErrorText, "Error") > 0 Then
        netMain.WriteText netInetGateway.GetErrorText, True
        netMain.WriteText "Try again in a few minutes.", True
        LogError "IxServerListSession", "Error: " & netInetGateway.GetErrorText
        IxServerListSession = False
        Exit Function
    End If
    
    'Check if the command was queued or not.
    If IxGetServerDirective(vResponse, "QUEUED_AT") <> "" Then
        
        'Command queued, will finish the session posting later.
        LogError "IxServerListSessions", "Request to the IxServer has been queued."
    
    ElseIf IxGetServerDirective(vResponse, "STATUS") = "LoginFail" Then
        LogError "IxServerListSession", "Login failed for user " & Trim(netMain.txtUserName.Text)
        netMain.WriteText "Please log in."
        IxServerListSession = False
        netIxServerLogin.SetCallback = "IxServerListSession"
        netIxServerLogin.Show , netMain
        
        
    'Check if the T&C have been signed.
    ElseIf IxGetServerDirective(vResponse, "TC_NEEDS_ACK") <> "" Then
        LogError "IxServerPostNewHaveIp", "T&C not signed by user " & Trim(netMain.txtUserName.Text)
        IxServerListSession = False
        netIxServerAgreement.SetCallback = "IxServerListSession"
        netIxServerAgreement.SetAgreementUrl = IxGetServerDirective(vResponse, "TC_NEEDS_ACK")
        netMain.WriteText "Please agree to the Terms and Conditions for online use.", False
        
        'Show the agreement form and set the owner form to one that is visible.
        If netMain.Visible Then
            netIxServerAgreement.Show , netMain
        Else
            netIxServerAgreement.Show , TheMainForm
        End If
        
    Else
        
        'All good, go to the next stage.
        If Not pSilent Then
            netMain.WriteText "Connected.", False
        End If
        Call netFindHosts.AddInternetSessions(vResponse, pSilent)
        Call IxProcessServerDirectives(vResponse, pSilent)
    End If
    
    Exit Function
ErrHand:
    LogError "IxServerListSession", "Error: " & Err.Number & " " & Err.Description
    netMain.WriteText "Error - " & Err.Description, True
End Function

'------------------------------------------------------------------------------------------------------
'       Private functions
'------------------------------------------------------------------------------------------------------

'Create a string correctly formated for use as the URL query for the Indexing Server
'and query the Indexing server returning its response. Any changes here must be reflected
'in the gs-indexing-server plugin at the constant defenitions section. If pForceToQueue
'is set to TRUE, the command is sent straight to the queue instead of being processes
'right away.
'The format of the command line arguments passed to the Indexing Server by the URL:
'define("arg_COMMAND",0);
'define("arg_SESION_ID",1);
'define("arg_HOST_IP",2);
'define("arg_TCP_PORT",3);
'define("arg_UDP_PORT",4);
'define("arg_CRYPT_KEYS",5);
'define("arg_LOCKED", 6);
'define("arg_HIDDEN", 7);
'define("arg_SES_NAME",8);
'define("arg_TERM_NAME",9);
'define("arg_GS_VERSION",10);
'define("arg_UNIQ_ID",11);
'define("arg_GS_CHECKSUM",12);
'define("arg_GS_TIME_STAMP",13);
'define("arg_GS_USERNAME",14);
'define("arg_GS_HASHCODE",15);
'define("arg_LOCAL_SES_ID",16);
'define("arg_COMMENT",17);
Private Function IxServerExecuteCommand(pCommand As eIxCommand, _
Optional pSESION_ID As String = "", _
Optional pHOST_IP As String = "", _
Optional pTCP_PORT As String = "", _
Optional pUDP_PORT As String = "", _
Optional pCRYPT_KEYS As String = "", _
Optional pLocked As String = "", _
Optional pHidden As String = "", _
Optional pSES_NAME As String = "", _
Optional pComment As String = "", _
Optional pPostString As String = "", _
Optional pForceToQueue As Boolean = False) As String
    
    Dim vGetString As String
    Dim vPostString As String
    Dim vReturn As String
    Dim vTimeBefore As String
    Dim vTimeDiff As Double
    
    On Error Resume Next
    
    
    'Build the URL with the query strings for comms with the Indexing Server.
    vGetString = gIndexServerUrl _
                & "?" & gGsLeUtils.LE6(CStr(pCommand) _
                & "," & pSESION_ID)
                    
    'If post data was passed, put a "&" symbol on the end so that the rest
    'of the query string can be appended.
    If Trim(pPostString) <> "" Then
        vPostString = Trim(pPostString) & "&"
    End If
    
    'Build the rest of the query string. This bit will get sent by POST.
    vPostString = vPostString & "arg=" & gGsLeUtils.LE6(pHOST_IP _
                    & "," & pTCP_PORT _
                    & "," & pUDP_PORT _
                    & "," & pCRYPT_KEYS _
                    & "," & pLocked _
                    & "," & pHidden _
                    & "," & EncodeNonAscii(pSES_NAME) _
                    & "," & IxGetCommonUrlTail _
                    & "," & pComment)
    
    'Check if we should queue the request.
    If Not netInetGateway.IsStillExecuting _
    And Not pForceToQueue Then
        
        'Record the time before contacting the Index Server.
        vTimeBefore = GetTimeStamp
        
        'Communicate with the Index Server.
        vReturn = IxServerComunication(vGetString, vPostString)
        
        'Work out the Index Server response time.
        vTimeDiff = CDbl(GetTimeStamp) - CDbl(vTimeBefore)
        Inet.IxServerPingTime = Format(vTimeDiff, "#0.00")
        Call LogInfo("IxServerExecuteCommand", "Ix Server ping time " _
                    & Inet.IxServerPingTime & "s", 5)
        
        'Decrypt the response from the Index Server.
        IxServerExecuteCommand = IxLedServerDirectives(vReturn)
    
    Else
        
        'Queue the request and notify the calling function
        'by returning "QUEUED_AT,<timestamp>" in the reply string.
        Call IxServerQueueCommand(vGetString, vPostString)
        IxServerExecuteCommand = "QUEUED_AT," & GetTimeStamp
        
    End If
    
    'Debug.Print vGetString & "---," & vPostString
    'Debug.Print IxServerExecuteCommand
End Function

'Gather info for common parts of the URL query. Used by both hosts and clients.
'Return "<Terminal Name>,<GlobalSiege Version>,<Unique ID>,<File_Checksum>,<GS_Time_Stamp>,<User_Name>,<Hash_Code>"
Public Function IxGetCommonUrlTail() As String
    Dim vUserName As String
    Dim vHashCode As String
    
    'vUserName = "baggins"
    'vPassword = "abc123"
    'vHashCode = "$P$BiL44qgI7yZ3yNS82a2p0cx6DRrVFa1"
    
    IxGetCommonUrlTail = EncodeNonAscii(netMain.txtTerminalName.Text) _
                & "," & EncodeNonAscii(GetVersionInfo) _
                & "," & EncodeNonAscii(GetUniqueId()) _
                & "," & evalChk.fileCS _
                & "," & EncodeNonAscii(GetTimeStamp) _
                & "," & gGsLeUtils.LE6(netMain.txtUserName.Text) _
                & "," & gGsLeUtils.LE6(netMain.txtUserName.Tag) _
                & "," & InetSes.LocalSessID
End Function

'Attempt to connect to the Indexing Server with the passed URL and return the
'page contens if successful. Try 4 times. Log errors as they pop up.
Public Function IxServerComunication( _
pAddress As String, _
Optional pPostData As String = "", _
Optional pSilent As Boolean = False) As String
    Dim vConnectAttempts As Long
    Dim vBeginPos As Long
    Dim vEndPos As Long
    Dim vUrlBodyText As String
    
    On Error GoTo ErrHand
    
    'Try 4 times.
    For vConnectAttempts = 0 To 2
        
        'This will be overwritten. If not, there was an error.
        IxServerComunication = "STATUS,ERROR - Could not contact the Indexing Server.<br />"
        If Not pSilent Then
            IxServerComunication = IxServerComunication & "SHOW_CHAT_BOX,ERROR - Could not contact the Indexing Server.<br />"
        End If
        
        'netInetGateway.GetErrorText = ""
        
        'Contact the Indexing Server via netInetGateway.
        vUrlBodyText = Trim(netInetGateway.ExecuteRequest(pAddress, pPostData))
        
        'Write the web page to the log directory if in dev mode.
        If gcAppDevelopMode Then
            SaveDataFile "IndexServerResponse.html", vUrlBodyText
        End If
        
        DoEvents
        'netInetGateway.GetErrorText
        'Check for any errors.
        If InStr(1, UCase(Trim(netInetGateway.GetErrorText)), "ERROR") <= 0 Then
            
            'Check header and log the error text if there is a problem.
            'Continue regardless of the header status. Errors in the body
            'will be found later. This is just to help log errors for
            'later analysis.
            If InStr(1, UCase(IxGetHeaderStatus(netInetGateway.GetHeadderText)), "OK") = 0 Then
                LogError "IxServerComunication", "IxGetHeaderStatus() returned " _
                        & IxGetHeaderStatus(netInetGateway.GetHeadderText)
            End If
            
            'Find BEGIN and END tokens and check if their positions make sense.
            vBeginPos = InStr(1, UCase(vUrlBodyText), "BEGIN_IX_SERVER_RESPONSE")
            vEndPos = InStr(1, UCase(vUrlBodyText), "END_IX_SERVER_RESPONSE")
            
            'Check BEGIN and END tokens make sense.
            If IxGetBodyTokensStatus(vBeginPos, vEndPos) = "OK" Then
                
                'All good so far. Exit the loop and return the body text between the
                'BEGIN and END tokens replacing the error message from above.
                IxServerComunication = Mid(vUrlBodyText, vBeginPos, _
                                    (vEndPos - vBeginPos) + Len("END_IX_SERVER_RESPONSE"))
                Exit For
            Else
                
                'Begin and end positions do not make sense.
                LogError "IxServerComunication", "IxGetBodyTokensStatus() returned " _
                        & IxGetBodyTokensStatus(vBeginPos, vEndPos)
            End If
        Else
            
            'netInetGateway.GetErrorText contains some error text. Log it and try again.
            LogError "IxServerComunication", "Error in netInetGateway.GetErrorText: " & netInetGateway.GetErrorText
        End If
        
        'Wait 2 seconds then try again up to 3 times.
        'Call pause(2000, True)
        Call Sleep(2000)
        If vConnectAttempts < 2 Then
            LogError "IxServerComunication", "Trying to connect again..."
            netMain.WriteText "Trying again...", False
        Else
            LogError "IxServerComunication", "Giving up."
            netMain.WriteText "Giving up. Cannot connect to the game server. " _
                                & "Please check your Internet connection.", False
            
            'Flush the Inet queue.
            Call IxServerFlushQueue
        End If
    Next
    
    Exit Function
ErrHand:
    LogError "IxServerComunication", "Error: " & Err.Number & " " & Err.Description
    If Not pSilent Then
        netMain.WriteText "Error contacting the Indexing Server. " & Err.Description, True
    End If
    Resume Next
End Function

'Decrypt the passed directives returned from the Index Server.
'vEncryption defaults to "None" and is set by the "METHOD" directive. The encryption
'method selected applies to the directives following on and should be changed back
'to "none" before the "END_IX_SERVER_RESPONSE" token.
Public Function IxLedServerDirectives(pIxServerText As String) As String
    Dim vReturn As String
    Dim vGsIndexPageText() As String
    Dim vParts() As String
    Dim vIndex As Long
    Dim vIndex2 As Long
    Dim vEncryption As String
    Dim vValidDirectives As Boolean
    
    On Error Resume Next
    
    'Clean and split the directives text into a string array.
    vGsIndexPageText = IxSplitHtmlLineBreaks(pIxServerText)
    
    'Explicitly set the default encryption
    vEncryption = "none"
    vValidDirectives = False
    
    For vIndex = 0 To UBound(vGsIndexPageText)
        
        'The "END_IX_SERVER_RESPONSE" token is not always encrypted. This is to ensure that
        'this token is still picked up if we forget to turn encryption off.
        If InStr(1, vGsIndexPageText(vIndex), "END_IX_SERVER_RESPONSE") > 0 Then
            vParts = Split(Trim(vGsIndexPageText(vIndex)), ",")
        Else
            vParts = Split(IxLedUsingMethod(Trim(vGsIndexPageText(vIndex)), vEncryption), ",")
        End If
        
        'Ignore blank lines which return -1.
        If UBound(vParts) >= 0 Then
            Select Case UCase(Trim(vParts(0)))
            
            'Valid directives only between "begin" and "end"
            Case "BEGIN_IX_SERVER_RESPONSE"
                vValidDirectives = True
                vReturn = vReturn & Join(vParts, ",") & "<br />" & vbCrLf
            
            Case "END_IX_SERVER_RESPONSE"
                vValidDirectives = False
                vReturn = vReturn & Join(vParts, ",")
            
            'Turn encryption on or off. If there is a second part, use it
            'as the encryption method. At the moment, the only methd is "le5"
            'which is LE5 with the default keys.
            Case "METHOD"
                If vValidDirectives Then
                    If UBound(vParts) >= 1 Then
                        vEncryption = LCase(Trim(vParts(1)))
                    Else
                        vEncryption = "none"
                    End If
                End If
                
            'Rebuild the directive and add to the return list.
            Case Else
                vReturn = vReturn & Join(vParts, ",") & "<br />" & vbCrLf
                
            End Select
        End If
    Next
    
    IxLedServerDirectives = vReturn
End Function

'Process and execute the general directives returned from the Indexing Server.
'If pSilent is true then no messages are posted to the user.
Private Sub IxProcessServerDirectives(pIxServerDirectives As String, _
Optional pSilent As Boolean = False, _
Optional pQueued As Boolean = False)
    Dim vPageText As String
    Dim vGsIndexPageText() As String
    Dim vParts() As String
    Dim vIndex As Long
    Dim vIndex2 As Long
    Dim vRetVal As VbMsgBoxResult
    Dim vValidDirectives As Boolean
    
    On Error Resume Next
    
    'Clean and split the directives text into a string array.
    vGsIndexPageText = IxSplitHtmlLineBreaks(pIxServerDirectives)
    
    vValidDirectives = False
    
    For vIndex = 0 To UBound(vGsIndexPageText)
        
        'Split and replace all non ascii characters.
        vParts = Split(vGsIndexPageText(vIndex), ",")
        For vIndex2 = 0 To UBound(vParts)
            vParts(vIndex2) = Trim(DecodeNonAscii(vParts(vIndex2)))
        Next
        
        'Ignore blank lines which return -1.
        If UBound(vParts) >= 0 Then
            Select Case UCase(vParts(0))
            
            'Valid directives only between "begin" and "end"
            Case "BEGIN_IX_SERVER_RESPONSE"
                vValidDirectives = True
            
            Case "END_IX_SERVER_RESPONSE"
                vValidDirectives = False
            
            'Some status text was returned either from the news server
            'or by the functions trying to connect to the news server.
            'This does nothing yet but it planned for fhe future.
            Case "GS_DEBUG_PRINT"
                If vValidDirectives And UBound(vParts) >= 1 Then
                    Debug.Print vParts(1)
                    LogInfo "IxProcessServerDirectives", vParts(1), 5
                End If
            
            'Latest version number available.
            'Format "LATEST_GS_RELEASE,<MAJ.MIN.Rev>,<Download_Page>"
            'Example "LATEST_GS_RELEASE,00.09.0092,http://www.globalsiege.net/download/"
            Case "LATEST_GS_RELEASE"
                If vValidDirectives And UBound(vParts) >= 2 Then
                    Call IxDisplayNewVersion(vParts(1), vParts(2))
                End If
            
            'Home web page for GlobalSiege. Defaults to "http://www.globalsiege.net"
            Case "GS_HOME_PAGE"
                If vValidDirectives And UBound(vParts) >= 1 Then
                    gHomeWebPage = vParts(1)
                End If
            
            'Help file web page.
            Case "HELP_WEB_PAGE"
                If vValidDirectives And UBound(vParts) >= 1 Then
                    gHelpWebPage = vParts(1)
                End If
            
            'Download web page.
            Case "DOWNLOAD_PAGE"
                If vValidDirectives And UBound(vParts) >= 1 Then
                    gDownloadWebPage = vParts(1)
                End If
            
            'Indexing server web page for online hosts.
            Case "NET_HOST_INDEX"
                If vValidDirectives And UBound(vParts) >= 1 Then
                    gIndexServerUrl = vParts(1)
                End If
            
            'Show text in a modal message box.
            'Format "SHOW_MESSAGE_BOX,<Title>,<Body>"
            'Escape code is "$", eg "$2C" returns a comma.
            Case "SHOW_MESSAGE_BOX"
                If vValidDirectives And UBound(vParts) >= 2 And Not pSilent Then
                    vRetVal = MsgBox(vParts(2), vbOKOnly, vParts(1))
                End If
            
            'Show text in the chatter box.
            'Format "SHOW_CHAT_BOX,<Text>"
            'Escape code is "$", eg "$2C" returns a comma.
            Case "SHOW_CHAT_BOX"
                If vValidDirectives And UBound(vParts) >= 1 And Not pSilent Then
                    Call netMain.WriteText(vParts(1), True)
                End If
            
            'IP list recieved from the Index Server. Add it
            'to the IP addresses that are already know.
            Case "YOUR_IP"
                If vValidDirectives And UBound(vParts) >= 1 And Not pSilent Then
                    
                    'Get a list of IP addresses from the Indexing Server.
                    InetSes.ReportedIP = Trim(DecodeNonAscii(vParts(1)))
                    InetSes.IP = IxServerCreateIpList(Trim(DecodeNonAscii(vGsIndexPageText(vIndex))))
                End If
            
            'The user has successfully logged in by passing a valid password as the hash key.
            Case "LOGGED_IN"
                If vValidDirectives And UBound(vParts) >= 1 And Not pSilent Then
                    
                    'Notify the user and save the hash code returned by the Index Server
                    'in netMain.txtUserName.Tag. This hash code will be used from now on
                    'to save the Indexing Server from hashing the clear password on every
                    'interaction. This hash code will be encrypted and saved in the
                    'registry under "IxPasswordHash".
                    netMain.WriteText "Logged in as " & Trim(gGsLeUtils.LE6d(vParts(1))) & ".", True
                    netMain.txtUserName.Tag = Trim(gGsLeUtils.LE6d(IxGetServerDirective(pIxServerDirectives, "HASH")))
                    SaveSetting gcApplicationName, "settings", "IxPasswordHash", gGsLeUtils.LE6(netMain.txtUserName.Tag)
                End If
            
            'The IX Server has acknowledged that the user has accepted the T&C.
            Case "TC_ACCEPTED"
                If vValidDirectives And UBound(vParts) >= 1 And Not pSilent Then
                    
                    'Carry on from where we left off.
                    netMain.WriteText "T&C has been marked as being accepted on " & Trim(vParts(1)) & " GMT.", True
                    SaveSetting gcApplicationName, "settings", "IxTC_Accepted", Trim(vParts(1))
                End If
            
            
            'The user's display name returned by the Indexing Server. This can be changed
            'by the user by logging into the user's account at www.globalsiege.net.
            Case "DISPLAY_NAME"
                If vValidDirectives And UBound(vParts) >= 1 Then
                    
                    'User's display name could be different to the user name.
                    netMain.txtTerminalName.Text = Trim(gGsLeUtils.LE6d(vParts(1)))
                    'If netFindHosts.Visible Then
                    '    netFindHosts.txtTerminalName.Text = netMain.txtTerminalName.Text
                    'End If
                End If
            
            'Status message returned from the Index Server.
            Case "STATUS"
                If vValidDirectives And pQueued And UBound(vParts) >= 0 Then
                    
                    'This is the Ix Server's response to "LIST_MY_IP". Attempt to post
                    'the new session.
                    If vParts(1) = "Ready_For_Post" Then
                        If Not IxServerPostNewHaveIp(pIxServerDirectives) Then
                            
                            'Failed to post the session. Clean up and reset.
                            With netMain
                            .tmrKeepAlive.Enabled = False
                            'WriteText "Could not post host name.", True
                            Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, True)
                            .optLan.Value = True
                            .optInet.Value = True
                            .optHost.Value = True
                            .chkHideSession.Enabled = False
                            .chkHideSession.Value = vbUnchecked
                            End With
                        End If
                    End If
                End If
            
            'New war has been posted successfully, arg 1 is the session ID
            'returned from the Ix Server.
            Case "SESSION_ID"
                If vValidDirectives And pQueued And UBound(vParts) >= 0 Then
                    If IsNumeric(vParts(1)) Then
                        If Not IxServerPostNewSesPosted(pIxServerDirectives) Then
                            'Failed to post the session.
                            With netMain
                            .tmrKeepAlive.Enabled = False
                            'WriteText "Could not post host name.", True
                            Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, True)
                            .optLan.Value = True
                            .optInet.Value = True
                            .optHost.Value = True
                            .chkHideSession.Enabled = False
                            .chkHideSession.Value = vbUnchecked
                            End With
                        End If
                    End If
                End If
            
            'A list of sessions is contained within the return body.
            Case "SESSION_LIST"
                If vValidDirectives And pQueued Then
                    If Not pSilent Then
                        netMain.WriteText "Connected.", False
                    End If
                    Call netFindHosts.AddInternetSessions(pIxServerDirectives, pSilent)
                End If
            
            'No match.
            Case Else
                'Ignore directive that are not understood, they could be meaningful
                'to another part of this program.
            End Select
        End If
    Next
End Sub

'Remove some commin HTML tags. Quick n dirty for VB6.
Public Function IxRemoveSomeHtmlTags(ByVal pHtmlTest As String) As String
    pHtmlTest = Replace(pHtmlTest, "&#8220;", """")
    pHtmlTest = Replace(pHtmlTest, "&#8221;", """")
    pHtmlTest = Replace(pHtmlTest, "&#8217;", "'")
    pHtmlTest = Replace(pHtmlTest, "&#038;", "&")
    pHtmlTest = Replace(pHtmlTest, "&amp;", "&")
    pHtmlTest = Replace(pHtmlTest, "-->", "")
    pHtmlTest = Replace(pHtmlTest, "<!--", "")
    pHtmlTest = Replace(pHtmlTest, "BEGIN_IX_SERVER_RESPONSE", "")
    pHtmlTest = Replace(pHtmlTest, "END_IX_SERVER_RESPONSE", "")
    pHtmlTest = Replace(pHtmlTest, vbTab, "")
    IxRemoveSomeHtmlTags = Trim(pHtmlTest)
End Function

'Replace all line breaks consistent by replacing with "<br />"
Public Function IxConsistantLineBreaks(ByVal pHtmlText As String, _
Optional pLineBreak As String = "<br />") As String
    pHtmlText = Replace(pHtmlText, vbCrLf, "")
    pHtmlText = Replace(pHtmlText, vbCr, "")
    pHtmlText = Replace(pHtmlText, vbLf, "")
    pHtmlText = Replace(pHtmlText, "<p>", "")
    pHtmlText = Replace(pHtmlText, "<P>", "")
    pHtmlText = Replace(pHtmlText, "</p>", pLineBreak)
    pHtmlText = Replace(pHtmlText, "</P>", pLineBreak)
    pHtmlText = Replace(pHtmlText, "<BR>", pLineBreak)
    pHtmlText = Replace(pHtmlText, "<br>", pLineBreak)
    pHtmlText = Replace(pHtmlText, "<BR />", pLineBreak)
    pHtmlText = Replace(pHtmlText, "<br />", pLineBreak)
    IxConsistantLineBreaks = pHtmlText
End Function

'Split on HTML line breaks and return a string array.
Public Function IxSplitHtmlLineBreaks(pHtmlText As String) As String()
    IxSplitHtmlLineBreaks = Split(IxConsistantLineBreaks(pHtmlText), "<br />")
End Function

'Decrypt pEncryptedText using the method pEncryptMethod.
'Methods are:
'   "None"   - None
'   "le5"    - LE5d using default keys.
'   "le6"    - LE6d using default keys.
'Note that "END" can be either encrypted or in the clear.
Private Function IxLedUsingMethod(pEncryptedText As String, pEncryptMethod As String) As String
    If UCase(pEncryptedText) = "END_IX_SERVER_RESPONSE" Then
        IxLedUsingMethod = pEncryptedText
    Else
        Select Case Trim(LCase(pEncryptMethod))
            Case "le5"
                IxLedUsingMethod = gGsLeUtils.LE5d(pEncryptedText)
            Case "le6"
                IxLedUsingMethod = gGsLeUtils.LE6d(pEncryptedText)
            Case Else
                IxLedUsingMethod = pEncryptedText
        End Select
    End If
End Function

'Notify the user if there is a new version available. Update the version label
'on the Setup Screen.
Private Sub IxDisplayNewVersion(pNewVer As String, pDownloadPage As String)
    'Dim vMsgboxText As String
    Dim vNewVersion As Double
    Dim vCurrentVersion As Double
    Dim vNewVerParts() As String
    Dim vNewVerFormatted As String
    
    On Error Resume Next
    
    vNewVerParts = Split(pNewVer, ".")
    
    If UBound(vNewVerParts) = 2 Then
        vNewVersion = CDbl(Replace(pNewVer, ".", ""))
        vCurrentVersion = CDbl(GetVersionInfo(""))
        
        If vNewVersion > vCurrentVersion Then
            vNewVerParts(0) = Trim(CStr(CLng(vNewVerParts(0))))
            vNewVerParts(1) = Trim(CStr(CLng(vNewVerParts(1))))
            vNewVerParts(2) = Trim(Format(vNewVerParts(2), "0000"))
            vNewVerFormatted = Join(vNewVerParts, ".")
            With TheMainForm.lblVersion
            'TheMainForm.lblVersion.Caption = Phrase(197) & SubstituteStringTokens("<Var.Maj>.<Var.Min>.<Var.Rev>") _
                        & " - Version " & vNewVerFormatted & " is available for download."
            .Caption = "Version " & vNewVerFormatted _
                        & " is available. Click here for more information"
            .Tag = pDownloadPage
            .ForeColor = &HFF0000
            .FontUnderline = True
            End With
        End If
    End If
End Sub

'Return the http status text from the passed URL header.
Private Function IxGetHeaderStatus(pHeaderText As String) As String
    Dim vIndex As Long
    Dim vHeader() As String
    
    vHeader = Split(pHeaderText, vbCrLf)
    
    'Check the length of the header.
    If UBound(vHeader) = 0 Then
        IxGetHeaderStatus = "Warning - Null header."
    Else
        
        'Find "HTTP/1.1 200 OK" tag in the header.
        IxGetHeaderStatus = "Warning - no status found in the URL header."
        For vIndex = 0 To UBound(vHeader)
            If InStr(1, Trim(vHeader(vIndex)), "HTTP/") = 1 Then
                IxGetHeaderStatus = Trim(vHeader(vIndex))
                Exit For
            End If
        Next
    End If
End Function

'Check the passed begin and end token positions and return the status text.
'"OK" means all ok. Anything else means bad.
Private Function IxGetBodyTokensStatus(pBeginPos As Long, pEndPos As Long) As String
    If pBeginPos >= pEndPos Then
        IxGetBodyTokensStatus = "BEGIN and END tokens invalid. BEGIN pos=" & CStr(pBeginPos) _
                & " END pos=" & CStr(pEndPos)
    
    ElseIf pBeginPos = 0 Then
        IxGetBodyTokensStatus = "BEGIN token is invalid. BEGIN pos=" & CStr(pBeginPos)
    
    ElseIf pEndPos = 0 Then
        IxGetBodyTokensStatus = "END token is invalid. END pos=" & CStr(pEndPos)
    Else
        IxGetBodyTokensStatus = "OK"
    End If
    
End Function

'Return TRUE if the directive exists.
Private Function IxServerDirectiveExist(pDirectives As String, pKey As String) As Boolean
    IxServerDirectiveExist = InStr(1, pDirectives, pKey) > 0
End Function

'Return the value of the passed key from the passed directive list. If more
'than one line contains the key, the keys will be returned delimited by crlf.
Private Function IxGetServerDirective(pDirectives As String, pKey As String) As String
    Dim vIndex As Long
    Dim vLines() As String
    
    vLines() = IxSplitHtmlLineBreaks(pDirectives)
    For vIndex = 0 To UBound(vLines)
        If InStr(1, vLines(vIndex), pKey) > 0 Then
            
            'If the line contains the key by itself, return the key.
            If vLines(vIndex) = pKey Then
                IxGetServerDirective = IxGetServerDirective & vLines(vIndex) & vbCrLf
            
            'Else return the kay values.
            Else
                vLines(vIndex) = Trim(Replace(vLines(vIndex), pKey, ""))
                
                'Remove the leading comma if one exists.
                If Mid(vLines(vIndex), 1, 1) = "," Then
                    vLines(vIndex) = Trim(Mid(vLines(vIndex), 2))
                End If
                
                'Lop off the last comma.
                If Right(vLines(vIndex), 1) = "," Then
                    vLines(vIndex) = Trim(Mid(vLines(vIndex), 1, Len(vLines(vIndex)) - 1))
                End If
                IxGetServerDirective = IxGetServerDirective & vLines(vIndex) & vbCrLf
            End If
        End If
    Next
    
    'Lop off the last crlf.
    If InStr(1, IxGetServerDirective, vbCrLf) > 0 Then
        IxGetServerDirective = Mid(IxGetServerDirective, 1, Len(IxGetServerDirective) - Len(vbCrLf))
    End If
End Function




