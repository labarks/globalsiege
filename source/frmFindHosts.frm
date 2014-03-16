VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form netFindHosts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Session Locator"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5415
   Icon            =   "frmFindHosts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrBroadcast 
      Left            =   4800
      Top             =   2880
   End
   Begin VB.Timer tmrEnableButton 
      Interval        =   500
      Left            =   4440
      Top             =   2880
   End
   Begin VB.TextBox txtSessionDescription 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Session Name"
         Object.Width           =   5901
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Vacant"
         Object.Width           =   1817
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Ping"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "description"
         Text            =   "Description"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "best_ping"
         Text            =   "BestPing"
         Object.Width           =   0
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckUDP 
      Left            =   4080
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label lblDescription 
      Caption         =   "Session description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
End
Attribute VB_Name = "netFindHosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------
'Handle all UDP socket communications and Session Listbox transactions.
'---------------------------------------------------------------------------------

'All UDP encryption keys are set to "17, 19, 3"

'The order in which session data is passed back from the Internet Indexing Server.
'Set DataUbound to the highest element number.
Private Enum eInetSess
    vCommand = 0
    Session_ID = 1
    Host_IP = 2
    TCP_Port = 3
    UDP_Port = 4
    Le_Keys = 5
    IsLocked = 6
    Ses_Name = 7
    Welcome_Msg = 8
    DataUbound = 8
End Enum

'The order in which session data is sent by the session host via the UDP socket.
'Set DataUbound to the highest element number.
Private Enum eSessData
    vCommand = 0
    Host_IP = 1
    Host_Port = 2
    Session_Name = 3
    Player_Count = 4
    Session_ID = 5
    Host_UDP_Port = 6
    Time_Stamp = 7
    IP_Sent_To = 8
    LeType = 9
    LeKey = 10
    LeSlot = 11
    LeSlotSpin = 12
    IsLocked = 13
    DataUbound = 13
End Enum

'The order in which session data is saved in the session's list item key.
'Set DataUbound to the highest element number.
Private Enum eLvKeyData
    Session_ID = 0
    Host_IP = 1
    Host_Port = 2
    Host_UDP_Port = 3
    Marker = 4
    LeType = 5
    LeKey = 6
    LeSlot = 7
    LeSlotSpin = 8
    IsLocked = 9
    DataUbound = 9
End Enum

'Cancel button cliek event. Set focus back to the main form in an attempt to
'stop it from hiding behind outher applications. The form_unload event is triggered.
'by the Unload Me statement where everything is cleaned up before unloading.
Private Sub cmdCancel_Click()
    On Error Resume Next
    TheMainForm.SetFocus
    Unload Me
End Sub

'Extract connection info from the passed session ListItem. Used to connect to the session.
'Return the SessionID from the Listview key.
Private Function GetThisSessionInfo(ByVal Item As ListItem) As String
    Dim vKeyParts() As String
    
    On Error Resume Next
    
    'The data contained in the key is as follows:
    '(0)Session_ID, (1)Host_IP, (2)Host_Port, (3)Host_UDP_Port, (4)-, _
    '(5)LeType, (6)LeKey, (7)LeSlot, (8)LeSlotSpin
    vKeyParts = Split(Item.Key, ",")
    
    'Get the session's IP address.
    InetSes.RemoteIP = GetIpAddress(vKeyParts(eLvKeyData.Host_IP))
    
    'Get the remote port number.
    InetSes.RemoteTcpPort = vKeyParts(eLvKeyData.Host_Port)
    
    'Set the session name.
    netMain.txtSesName.Text = Trim(Item.Text)
    InetSes.SesName = Trim(Item.Text)
    
    'Set the session encryption keys.
    MyNetRegData.LeType = vKeyParts(eLvKeyData.LeType)
    MyNetRegData.LeKey = vKeyParts(eLvKeyData.LeKey)
    MyNetRegData.LeSlot = vKeyParts(eLvKeyData.LeSlot)
    MyNetRegData.LeSlotSpin = vKeyParts(eLvKeyData.LeSlotSpin)
    
    'Return the SessionID from the List Item's key.
    GetThisSessionInfo = vKeyParts(eLvKeyData.Session_ID)
End Function

'Click connect button on netMain.
Private Sub cmdConnect_Click()
    
    'Extrext connection info from the list box into global variables and tags.
    InetSes.RemoteSessID = GetThisSessionInfo(ListView1.SelectedItem)
    
    'Validate the port numbers on the Network Admin screen.
    If netMain.ValidateBeforeConnection Then
        
        'Port numbers are good, commence connecting.
        Call netMain.ConnectTo(InetSes.RemoteSessID)
        
    End If
    Unload Me
End Sub

'Return an IP address from the passed IP list. If it is a list, split it up and
'use the first one IP address because it probably has the fastest ping time if
'it has indeed answered a ping. The passed list is "x" delimited.
Private Function GetIpAddress(pIpList As String) As String
    Dim vIPs() As String
    
    On Error GoTo ErrHand
    
    'Split the passed IP list into an array.
    vIPs = Split(pIpList, "x")
    
    'Return the first element in the array.
    If UBound(vIPs) >= 0 Then
        GetIpAddress = vIPs(0)
    Else
        GetIpAddress = pIpList
    End If
    
    Exit Function
ErrHand:
    GetIpAddress = pIpList
    Exit Function
End Function

'Form load event. Disable the Connect button and clear the list view, which
'we don't really need to do.
Private Sub Form_Load()
    Dim vListItem As ListItem
    cmdConnect.Enabled = False
    ListView1.ListItems.Clear
End Sub

'Form unload event. Switch off the timer and set the sellected session IP
'address to en empty string "".
Private Sub Form_Unload(Cancel As Integer)
    tmrBroadcast.Enabled = False
    InetSes.RemoteIP = ""
    TheMainForm.SetFocus
End Sub

'Update the session description text when the selection changes in the list view.
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtSessionDescription.Text = ListView1.SelectedItem.SubItems(3)
End Sub

'Sort listview by clicked subitem.
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
End Sub

'Connect if dbl clicked.
Private Sub ListView1_DblClick()
    If cmdConnect.Enabled Then
        Call cmdConnect_Click
    Else
        Call cmdCancel_Click
    End If
End Sub

Private Sub ListView1_GotFocus()
    cmdConnect.Enabled = (ListView1.ListItems.Count > 0)
    'ListView1.SelectedItem.Key
End Sub

'Ensure cmdConnect is enables only if there are sessions being displayed.
Private Sub tmrEnableButton_Timer()
    If ListView1.ListItems.Count > 0 Then
        cmdConnect.Enabled = True
        txtSessionDescription.Text = ListView1.SelectedItem.SubItems(3)
    Else
        cmdConnect.Enabled = False
        txtSessionDescription.Text = ""
    End If
End Sub

'I am a client.Resend the host finder broadcast every four seconds, ping all Internet
'hosts every other four seconds and requery the Internet Indexing server every 40 seconds.
'sckUDP.SendData tends to throw errors every second time it is executed if no hosts
'have replied. This seems to be a Windows problem of some sort. Another random problem is
'that hosts sometimes don't answer.
Private Sub tmrBroadcast_Timer()
    On Error Resume Next
    
    'Turn off the timer so that it doesn't call here before finishing.
    tmrBroadcast.Enabled = False
    
    'Check time stamp of all sessions and place a dash in the ping column
    'of any hosts that hasn't responded in more than 20 seconds.
    Call DashOutThenRemoveOldSessions
    
    'Select a different course of action each time through this sub. InetSes.RebroadcastCntr
    'is a global variable so that it can be reset to zero externally.
    If InetSes.RebroadcastCntr Mod 20 = 19 Then
        
        'Requery Internet Indexing Server every 20 seconds.
        If netMain.optInet.Value And netMain.optJoin.Value Then
            Call IxServerListSession(True)
        End If
        
        'Reset the counter and begin again.
        InetSes.RebroadcastCntr = 0
        
    ElseIf InetSes.RebroadcastCntr Mod 2 = 1 Then
        
        'Ping the internet hosts.
        Call PingInternetHosts
        
    Else
        
        'Broadcast for LAN hosts.
        Call SendClientUdpRequest(cgBroadcastAddress, InetSes.LocalUdpPort)
    
    End If
    
    'Increase for the next run through.
    InetSes.RebroadcastCntr = InetSes.RebroadcastCntr + 1
    DoEvents
    
    'Turn the timer back on.
    tmrBroadcast.Enabled = True
    
End Sub

'Choose a colour for the session in the Listview depending on if it is locked,
'has answered a ping etc.
Private Function ChooseSessionListviewColor(pListItem As ListItem)
    Dim vPingTime As String
    Dim vIsLocked As Boolean
    
    'Get the ping time. Could be a number or a dash if more than 20 seconds old.
    vPingTime = pListItem.SubItems(2)
    
    'Find if the session is locked. Could be a "1" or a "0". Convert to a boolean.
    vIsLocked = (Trim(GetListElement(pListItem.Key, eLvKeyData.IsLocked)) = 1)
    
    'Work out the possible combinations.
    If Not vIsLocked And IsNumeric(vPingTime) Then
        
        'Session not locked and has a recent ping reply.
        pListItem.ForeColor = RGB(0, 192, 96)
        pListItem.Bold = True
        pListItem.ToolTipText = "" & pListItem.Text & " - alive"
    
    ElseIf Not vIsLocked And Not IsNumeric(vPingTime) Then
        
        'Session not locked and no ping reply.
        pListItem.ForeColor = vbRed
        pListItem.Bold = False
        pListItem.ToolTipText = "" & pListItem.Text & " - status unknown"
        
    ElseIf vIsLocked And IsNumeric(vPingTime) Then
        
        'Session is locked and has a recent ping reply.
        pListItem.ForeColor = vbBlue
        pListItem.Bold = True
        pListItem.ToolTipText = "" & pListItem.Text & " - alive [password protected]"
        
    ElseIf vIsLocked And Not IsNumeric(vPingTime) Then
        
        'Session is locked and no ping reply.
        pListItem.ForeColor = vbRed
        pListItem.Bold = False
        pListItem.ToolTipText = pListItem.Text & " - status unknown [password protected]"
        
    End If
End Function

'Display the name and session details choosing the correct colour and if locked,
'put "<Password protected>" in the description.
Private Sub FormatSessionListviewDisplay(pListItem As ListItem, _
pSessionName As String, _
pSessionDescription As String)
    Dim vSessionDescription As String
    
    'Remove the "<Password protected>" phrase from the description that is already there.
    vSessionDescription = Replace(pListItem.SubItems(3), _
                        "<Password protected>" & vbCrLf, "", , 1)
    
    'Work out if a session description was passed.
    If Trim(pSessionDescription) <> "" Then
        
        'Session description was passed. Use it.
        vSessionDescription = pSessionDescription
        
    End If
    
    'Choose a colour for the session's text in the Listview.
    Call ChooseSessionListviewColor(pListItem)
    
    'Set the session name.
    pListItem.Text = pSessionName
    
    'Is the session is locked?
    If Trim(GetListElement(pListItem.Key, eLvKeyData.IsLocked)) = 1 Then
        
        'Description from the Welcome_Msg.
        pListItem.SubItems(3) = "<Password protected>" & vbCrLf _
                                & Replace(vSessionDescription, "<Password protected>" & vbCrLf, "", , 1)
    
    Else
        
        'Description from the Welcome_Msg.
        pListItem.SubItems(3) = vSessionDescription
        
    End If
End Sub

'Check if the passed directive from the Indexing Server is an INternet session and
'add or update the passed Internet session to the list box. Return TRUE if successful.
'A session key is created and added to the list box withwith ping and player count
'dashed out.
'Input format expected:
'"SESSION,<Session_ID>,<Host_IPs>,<TCP_Port>,<UDP_Port>,<Le_Keys>,<Locked>,<Ses_Name>,<Welcome_Msg>,"
'Session key format:
'"Session_ID,Host_IPs,Host_Port,Host_UDP_Port,'X',LeType,LeKey,LeSlot,LeSlotSpin,IsLocked"
Private Function AddThisInternetSession(pSessionData As String) As Boolean
    Dim vDataParts() As String
    Dim vKeyParts() As String
    Dim vListItem As ListItem
    Dim vLindex As Long
    Dim vLvItemKey As String
    
    'Split by comma.
    vDataParts = Split(pSessionData, ",")
    
    'Look for the "SESSION" token and validate the session listing string.
    'There are eInetSess.DataUbound + 1 elemnts because of the trailing comma
    'sent by the Indexing Server. This trailing comma is needed to get around
    'a PHP string bug which seems to lop off the last element.
    With ListView1
    If UBound(vDataParts) > 0 Then
        If vDataParts(0) = "SESSION" _
        And UBound(vDataParts) >= eInetSess.DataUbound Then
            
            'Session name.
            vDataParts(eInetSess.Ses_Name) = DecodeNonAscii(vDataParts(eInetSess.Ses_Name))
            
            'List of IP addresses.
            vDataParts(eInetSess.Host_IP) = CleanList(Replace(vDataParts(eInetSess.Host_IP), "_", "."))
            
            'Session description and welcome message.
            vDataParts(eInetSess.Welcome_Msg) = DecodeNonAscii(vDataParts(eInetSess.Welcome_Msg))
            
            'Create the item key.
            '"Session_ID,Host_IP,Host_Port,Host_UDP_Port,'X',
            'LeType,LeKey,LeSlot,LeSlotSpin,IsLocked"
            vLvItemKey = vDataParts(eInetSess.Session_ID) _
                    & "," & vDataParts(eInetSess.Host_IP) _
                    & "," & vDataParts(eInetSess.TCP_Port) _
                    & "," & vDataParts(eInetSess.UDP_Port) _
                    & ",X" _
                    & "," & GetListElement(vDataParts(eInetSess.Le_Keys), 0, ":") _
                    & "," & GetListElement(vDataParts(eInetSess.Le_Keys), 1, ":") _
                    & "," & GetListElement(vDataParts(eInetSess.Le_Keys), 2, ":") _
                    & "," & GetListElement(vDataParts(eInetSess.Le_Keys), 3, ":") _
                    & "," & vDataParts(eInetSess.IsLocked)
            
            'Check if the session ID is already in the list box.
            vLindex = GetSessionListindex(vDataParts(eInetSess.Session_ID))
            If vLindex = -1 Then
                
                'Not already in the list, add to the list box.
                'Add a new item with the key and session name.
                Set vListItem = .ListItems.Add(, vLvItemKey, vDataParts(eInetSess.Ses_Name))
                
                'Stop the first item in the list from hilighted.
                vListItem.Selected = False
                
                'Set the Player Vacant column to a dash.
                vListItem.SubItems(1) = "-"
                
                'Set the Ping column. to a dash.
                vListItem.SubItems(2) = "-"
            Else
                
                'Already in the list box. Mark with an 'X' so it doesn't get deleted.
                Set vListItem = .ListItems(vLindex)
                
                vKeyParts = Split(vListItem.Key, ",")
                vKeyParts(eLvKeyData.Marker) = "X"
                vKeyParts(eLvKeyData.IsLocked) = vDataParts(eInetSess.IsLocked)
                vListItem.Key = Join(vKeyParts, ",")
                
            End If
            
            'Display the name and session details choosing the correct colour and if locked,
            'put a "<Password protected>" in the description.
            Call FormatSessionListviewDisplay(vListItem, _
                            vDataParts(eInetSess.Ses_Name), _
                            vDataParts(eInetSess.Welcome_Msg))
            
            'Notify that a session was either added or updated.
            AddThisInternetSession = True
        Else
            
            'No valid session found in the passed session data.
            AddThisInternetSession = False
        End If
    End If
    End With
End Function

'Populate the Sessions Found listbox with the list from the Index Server.
'Called from modNetIxServer.IxProcessServerDirectives() and
'modNetIxServer.IxProcessServerDirectives()
'Format: "SESSION,<Session_ID>,<Host_IP>,<TCP_Port>,<UDP_Port>,<Le_Keys>,
'<Locked>,<Ses_Name>,<Welcome_Msg>"
Public Function AddInternetSessions(pServerCommands As String, _
Optional pSilent As Boolean = False) As Boolean
    Dim vResponseArray() As String
    Dim vIndex As Long
    Dim vSessionCount As Long
    
    'Split the passed directives by line break.
    vResponseArray = IxSplitHtmlLineBreaks(pServerCommands)
    
    'Only if there are lines in the array.
    If UBound(vResponseArray) >= 0 Then
        
        'Mark Internet sessions for timeout.
        Call MarkAllListedInetSessions
        
        'For each line in the command text.
        vSessionCount = 0
        With ListView1
        For vIndex = 0 To UBound(vResponseArray)
            
            'Increase the count if the session was added or updated.
            If AddThisInternetSession(vResponseArray(vIndex)) Then
                vSessionCount = vSessionCount + 1
            End If
        
        Next
        
        Call RemoveDeadListedInetSessions

        'Tell the player how many sessions were found.
        If Not pSilent Then
            If vSessionCount = 0 Then
                
                'No sessions were found.
                netMain.WriteText "No Internet sessions found. Start your own session or try again later." _
                                    , Me.Visible
                AddInternetSessions = False
            
            ElseIf vSessionCount = 1 Then
                
                'Only one session found. Must wordify it rightly.
                netMain.WriteText "1 Internet session found.", False
                AddInternetSessions = True
                DoEvents
                
                'Begin pinging the IP addresses of the found host.
                Call PingInternetHosts
            
            Else
                
                'Two or more sessions found.
                netMain.WriteText CStr(vSessionCount) & " Internet sessions found.", Me.Visible
                AddInternetSessions = True
                DoEvents
                
                'Befin pinging all the addresses of all the hosts.
                Call PingInternetHosts
            
            End If
        End If
        End With
    Else
        
        'There was no lines at all in the paddes directives list.
        If Not pSilent Then
            netMain.WriteText "No Internet sessions found. Start your own session or try again later." _
                                , Me.Visible
        End If
    End If
    Exit Function
ErrHand:
    LogError "AddInternetSessions", "Error: " & Err.Number & " " & Err.Description
    netMain.WriteText "Error - " & Err.Description, False
End Function

'Return the time in seconds since the time of the passed timestamp.
'Return "-" if the time is more is than 20 seconds or if there is an error.
'The pTimeStamp is in seconds accurate to three decimal places. The return
'time is formatted into seconds accurate to two decimal places.
Private Function GetPingTime(pTimeStamp As String) As String
    Dim vTimeDiff As Double
    
    On Error GoTo ErrHand
    
    'Find the length of time since the passed time stamp.
    vTimeDiff = CDbl(GetTimeStamp) - CDbl(DecodeNonAscii(pTimeStamp))
    
    'Check the time difference is more than zero and less than 20 seconds.
    If vTimeDiff >= 0 And vTimeDiff < 20 Then
        
        'Between 0 and 20 seconds. Format and return the time difference.
        GetPingTime = Format(vTimeDiff, "#0.00")
    
    'If the time diff doesn't make sense, set it to random seconds. This is a bodgy
    'work-around for strange GS Timestamps being returned from the host.
    ElseIf vTimeDiff > 1000 Or vTimeDiff < -1000 Then
        GetPingTime = Format(Rnd * 7 + 3, "#0.00")
        
    Else
        
        'Lett than  0 or more than 20. Return a dash '-'.
        GetPingTime = "-"
    
    End If
    Exit Function
ErrHand:
    
    'Return a dash if there is any error.
    GetPingTime = "-"
    Exit Function
End Function

'Mark all Internet sessions with a dash '-' in the key to indicate that it is
'already in the list. When we get a new list from the Indexing Server, the
'dash will be overwritten with an 'X' if the session is in the list. Any seesions
'that still have a dash in the key will be deleted by function RemoveDeadListedInetSessions().
Private Sub MarkAllListedInetSessions()
    Dim vListItem As ListItem
    Dim vIndex As Long
    Dim vLvKeyParts() As String
    
    'For each session in the list.
    With ListView1
    For vIndex = 1 To .ListItems.Count
        
        'Set the object variable.
        Set vListItem = .ListItems(vIndex)
        
        'Split the key into vLvKeyParts.
        vLvKeyParts = Split(vListItem.Key, ",")
        
        'Check the number of elements in the key to prevent errors.
        If UBound(vLvKeyParts) = eLvKeyData.DataUbound Then
            
            'Ignore LAN sessions. They get deleted by
            'function DashOutThenRemoveOldSessions().
            If Mid(vLvKeyParts(eLvKeyData.Session_ID), 1, 3) <> "LAN" Then
                
                'Make the relevent element with a dash and rejoin the key.
                vLvKeyParts(eLvKeyData.Marker) = "-"
                vListItem.Key = Join(vLvKeyParts, ",")
                
            End If
        End If
    Next
    End With
End Sub

'Check time stamp of listitems which is stored in tag and place a dash in the
'ping time of hosts who haven't responded in more than 20 seconds. We need to
'bail out of the loop early sometimes because if an item is removed from the list
'and we continue itterating that list, we will miss the next session because it has
'moved up the list to fill the removed session's place and we risk going past the end.
Private Sub DashOutThenRemoveOldSessions()
    Dim vListItem As ListItem
    Dim vIndex As Long
    Dim vHostTS As String
    Dim vPingTime As String
    Dim vStartAgain As Boolean
    
    On Error Resume Next
    
    With ListView1
    
    'Keep looping until there are no sessions to delete. We need
    'to do it this way because when a session is deleted, we miss the
    'next one that moves up into its place. We also risk running past
    'the end of the collection.
    Do
        vStartAgain = False
        
        'For each session in the LIstView.
        For vIndex = 1 To .ListItems.Count
            
            'Reference the session item.
            Set vListItem = .ListItems(vIndex)
            
            'Get the last response timestamp from the item's tag.
            vHostTS = vListItem.Tag
            
            'If the timestamp is a number, not blank or already a dash "-".
            If IsNumeric(vHostTS) Then
                
                'Find how long it has been since the last reply from this host.
                vPingTime = GetPingTime(vHostTS)
                
                'If the time is more that 20 seconds, which is signified by a dash.
                If vPingTime = "-" Then
                
                    'Ping time > 20 seconds. If the session is a LAN session.
                    If Mid(vListItem.Key, 1, 3) = "LAN" Then
                        
                        'Remove if the session is a LAN session.
                        .ListItems.Remove vIndex
                        
                        'Set to run the whole loop again.
                        vStartAgain = True
                        Exit For
                        
                    Else
                        
                        'Put a dash in the ping time for Internet sessions but
                        'do not remove. They will be removed during the next
                        'query to the Indexing Server.
                        vListItem.SubItems(2) = "-"
                        
                    End If
                End If
            Else
                
                'The session time stamp in the tag is not numeric.
                vListItem.SubItems(2) = "-"
            End If
            
            Call ChooseSessionListviewColor(vListItem)
        Next
    
    'Start again if a session was removed.
    Loop While vStartAgain
        
    End With
End Sub

'Remove Internet sessions that are no longer listed by the Index Server. We need to
'bail out of the loop early sometimes because if an item is removed from the list
'and we continue itterating that list, we will miss the next session because it has
'moved up the list to fill the removed session's place and we risk going past the end.
Private Sub RemoveDeadListedInetSessions()
    Dim vListItem As ListItem
    Dim vIndex As Long
    Dim vLvKeyParts() As String
    Dim vStartAgain As Boolean
    
    On Error Resume Next
    
    With ListView1
    
    'Keep looping until there are no sessions to delete. We need
    'to do it this way because when a session is deleted, we miss the
    'next one that moves up into its place. We also risk running past
    'the end of the collection.
    Do
        
        vStartAgain = False
        
        'For each session in the list.
        For vIndex = 1 To .ListItems.Count
            
            'Point to the list item.
            Set vListItem = .ListItems(vIndex)
            
            'Split the session from the Key.
            vLvKeyParts = Split(vListItem.Key, ",")
            
            'Check there are enough elements to prevent errors.
            If UBound(vLvKeyParts) = eLvKeyData.DataUbound Then
                
                'If the session is an Internet session. This is can be determined by the
                'fact that LAN only sessions have a unique ID that startw with "LAN".
                'This comes from InetSes.LocalSessID on the session host.
                If Mid(vLvKeyParts(eLvKeyData.Session_ID), 1, 3) <> "LAN" Then
                    
                    'Remove the session if it has a dash in the marker position.
                    'vLvKeyParts(eLvKeyData.Marker) will be 'X' if the session is newly listed.
                    If vLvKeyParts(eLvKeyData.Marker) = "-" Then
                        
                        'Remove the session.
                        .ListItems.Remove vIndex
                        
                        'Set to run the whole loop again.
                        vStartAgain = True
                        Exit For
                        
                    End If
                End If
            End If
        Next
    
    'Start again if a session was removed.
    Loop While vStartAgain
    
    End With
End Sub

'Return List Index of the passed session ID in the listbox, -1 if not in there.
Private Function GetSessionListindex(HostID As String) As Integer
    Dim vListItem As ListItem
    Dim vIndex As Long
    Dim vLvKeyParts() As String
    
    'Assume the session is not there.
    GetSessionListindex = -1
    
    With ListView1
    
    'For each session in the Listview.
    For vIndex = 1 To .ListItems.Count
        
        'Set the session to the object variable.
        Set vListItem = .ListItems(vIndex)
        
        'Split the session key.
        vLvKeyParts = Split(vListItem.Key, ",")
        
        'Compare the host ID.
        If vLvKeyParts(eLvKeyData.Session_ID) = HostID Then
            
            'Session found. Return the list index.
            GetSessionListindex = vIndex
            Exit For
            
        End If
    Next
    End With
End Function

'I am a host. A client has sent a session info request. Validate the request and
'format and sent a reply to that client containing session info.
Private Function ReplyToClientBroadcast(pRequestParts() As String) As String
    Dim vReplyString As String
    Dim vSessionID As String
    
    On Error GoTo ErrHand
    
    With MyNetRegData
    
    'Verify that this is a valid message sent by a client and
    'the session is not hidden.
    If pRequestParts(0) = "RequestDetail" _
    And UBound(pRequestParts) = 4 _
    And netMain.chkHideSession = vbUnchecked Then
        
        'Internet client request recieved.
        
        'Internet client request format:
        '"(0)RequestDetail, (1)Client_IP, (2)Client_Port, (3)Time_Stamp, (4)IP_Sent_To"
        
        'Reply to the client:
        '"(0)SesDetail, (1)Host_IP, (2)Host_Port, (3)Session_Name, (4)Player_Count,
        '(5)**Session_ID, (6)Host_UDP_Port, (7)Time_Stamp, (8)IP_Sent_To, (9)LeType,
        '(10)LeKey, (11)LeSlot, (12)LeSlotSpin, (13)IsLocked, (14)Local_Ses_ID"
        '**Don't like this, should keep the session ID from the Index Server a secret to thwart game hacking.
        
        'If it was a broadcast, the IP_Sent_To will be an 'X'. Send my (host's) local IP address.
        If pRequestParts(4) = "X" Then
            pRequestParts(4) = ChooseValidIP(netMain.sckTCP(0).LocalIP, netMain.sckTCP(0).LocalHostName)
        End If
        
        'Get the session ID. This is assigned by the Indexing Server if this is an Internet
        'session, however if this is a LAN only session, this ID will not be set so we
        'need to use another number which is unique to this session.
        If InetSes.ID = "" Then
            vSessionID = "LAN-" & InetSes.LocalSessID
        Else
            vSessionID = InetSes.LocalSessID 'InetSes.ID
        End If
        
        'Build the reply string.
        vReplyString = "SesDetail," _
                        & InetSes.IP & "," _
                        & netMain.sckTCP(0).LocalPort & "," _
                        & EncodeNonAscii(Trim(netMain.txtSesName.Text)) & "," _
                        & TheMainForm.CountClaimedPlayers & "," _
                        & vSessionID & "," _
                        & InetSes.LocalUdpPort & "," _
                        & pRequestParts(3) & "," _
                        & pRequestParts(4) & "," _
                        & .LeType & "," _
                        & .LeKey & "," _
                        & .LeSlot & "," _
                        & .LeSlotSpin & "," _
                        & CStr(netMain.chkPasswordSession.Value)
        
    End If
    
    ReplyToClientBroadcast = vReplyString
    
    End With
    Exit Function
ErrHand:
    Debug.Print "ReplyToClientBroadcast() Error: " & Err.Number & " " _
                & Err.Description, sckUDP.State, sckUDP.RemoteHost, sckUDP.RemoteHostIP
    LogError "ReplyToClientBroadcast", "Error: " & Err.Number & " " & Err.Description
    Resume Next
    Exit Function
End Function

'Check the passed ping time and set to 20 seconds if not numeric.
'This would mean that it hasn't been filled in yet or could be dashed out.
'Convert to a single data type.
Private Function ValidatePingTime(pPingTime As String) As Single
    If IsNumeric(pPingTime) Then
        ValidatePingTime = CSng(pPingTime)
    Else
        ValidatePingTime = 20
    End If
End Function

'Check the best ping time in the passed session list item. If the passed IP address
'is the best time, set it as the first IP address in the session's IP list.
Private Sub ArrangeBestIpPingTime(pListView As ListItem, pIpAddress As String, pPingTime As String)
    Dim vPingTime As Single
    Dim vBestPing As Single
    Dim vKeyParts() As String
    
    On Error Resume Next
    
    'Check the passed ping time and best ping time and set to 20 seconds if not numeric.
    'This would mean that it hasn't been filled in yet or could be dashed out.
    vPingTime = ValidatePingTime(pPingTime)
    vBestPing = ValidatePingTime(pListView.SubItems(4))
    
    'If the new ping time is ledd than the best ping time, set the IP address
    'as the first in the IP list and update the best ping time.
    If vPingTime < vBestPing Then
    
        'Extract the IP list from the Listbox key.
        vKeyParts = Split(pListView.Key, ",")
        
        'Add the IP address to the front of the list and remove all duplicates
        'which in effect moves it to the front of the list.
        vKeyParts(eLvKeyData.Host_IP) = CleanList(pIpAddress & "x" & vKeyParts(eLvKeyData.Host_IP), "x")
        
        'Reassemble the Listview key.
        pListView.Key = Join(vKeyParts, ",")
        
        'Update the best ping time.
        pListView.SubItems(4) = CStr(vPingTime)
    End If
End Sub

'I am a client. A reply to my broadcast has been recieved. Put into or
'update the sessions as required.
Private Function BroadcastReplyRecieved(pRequestParts() As String) As String
    Dim vLvItemIndex As Long
    Dim vListItem As ListItem
    Dim vPingTime As String
    Dim vIpList As String
    
    On Error GoTo ErrHand
    
    'Host response format:
    '"(0)GS#HostInetDetails1.0, (1)Host_IP, (2)Host_Port, (3)Session_Name, (4)Player_Count,
    '(5)Session_ID, (6)Host_UDP_Port, (7)Time_Stamp, (8)IP_Sent_To, (10)LeKey,
    '(10)LeKey, (11)LeSlot, (12)LeSlotSpin, (13)IsLocked, (14)Local_Ses_ID"
        
    With ListView1
    
    'Validate the response.
    If pRequestParts(eSessData.vCommand) = "SesDetail" _
    And UBound(pRequestParts) = eSessData.DataUbound Then
        
        'Find the index of the session in the Listview.
        vLvItemIndex = IPInetListIndex(pRequestParts(eSessData.Session_ID))
        
        'Replace non ascii characters in the session's name.
        pRequestParts(eSessData.Session_Name) = DecodeNonAscii(pRequestParts(eSessData.Session_Name))
        
        'Is the session already in the listview.
        If vLvItemIndex = -1 Then
            
            'Not in the Listview, add it with a random key. This key
            'will be updated when the key is assembled below.
            Set vListItem = .ListItems.Add(, "ABC" & CStr(Rnd), CStr(pRequestParts(eSessData.Session_Name)))
            
            'Stop the first item in the list from hilighted.
            vListItem.Selected = False
            
            'Set the contact IP address to the original IP address that
            'found the host.
            vIpList = pRequestParts(eSessData.IP_Sent_To)
        Else
            
            'Session found in the Listview.
            Set vListItem = .ListItems(vLvItemIndex)
            
            'Keep the list of IP addresses. The order will be changed if
            'the IP address that found the host has a quickest ping time.
            vIpList = GetListElement(vListItem.Key, 1)
        End If
        
        'Listitem key:
        '"(0)Session_ID, (1)Host_IP, (2)Host_Port, (3)Host_UDP_Port, (4)-, _
        '(5)LeType, (6)LeKey, (7)LeSlot, (8)LeSlotSpin"
        vListItem.Key = pRequestParts(eSessData.Session_ID) & "," _
            & vIpList & "," _
            & pRequestParts(eSessData.Host_Port) & "," _
            & pRequestParts(eSessData.Host_UDP_Port) & "," _
            & "-," _
            & pRequestParts(eSessData.LeType) & "," _
            & pRequestParts(eSessData.LeKey) & "," _
            & pRequestParts(eSessData.LeSlot) & "," _
            & pRequestParts(eSessData.LeSlotSpin) & "," _
            & pRequestParts(eSessData.IsLocked)
        
        'Display the number of vacant players.
        vListItem.SubItems(1) = pRequestParts(eSessData.Player_Count)
        
        'Display the ping time.
        vPingTime = GetPingTime(DecodeNonAscii(pRequestParts(eSessData.Time_Stamp)))
        vListItem.SubItems(2) = vPingTime
        
        'Display the name and session details choosing the correct colour and if locked,
        'put a "[P]" for private in the name and "<Password protected>" in the description.
        Call FormatSessionListviewDisplay(vListItem, pRequestParts(eSessData.Session_Name), "")
        
        'Arrange the sessions's IP list so that the fastest one is the first one.
        Call ArrangeBestIpPingTime(vListItem, pRequestParts(eSessData.IP_Sent_To), vPingTime)
        
        'Put the timestamp of the last ping response in the Listitem's tag.
        vListItem.Tag = GetTimeStamp
        
    End If
    End With
    Exit Function
ErrHand:
    LogError "BroadcastReplyRecieved", "Error: " & Err.Number & " " & Err.Description
    Debug.Print "BroadcastReplyRecieved() Error: " & Err.Number & " " _
                & Err.Description, sckUDP.State, sckUDP.RemoteHost, sckUDP.RemoteHostIP
    Resume Next
    Exit Function
End Function


'Send data via the UDP port. This cannot be used for the client request at this
'stage because of the way it needs to be done in SendClientUdpRequest().
Private Function SendUdpData(pRemoteHostIP As String, _
pRemotePort As String, _
pMessageText As String) As Boolean
    Dim vByteArray() As Byte
    
    On Error GoTo ErrHand
    
    'Encrypt and convert to a hex encoded byte array.
    Call HexStringToByteArray(vByteArray, gGsLeUtils.LE6(pMessageText, 17, 19, 3))
    
    'Set up the UDP socket for to send the passed
    'text to the passed IP and Port numbers.
    sckUDP.RemoteHost = pRemoteHostIP
    sckUDP.RemotePort = pRemotePort
    
    'Send the passed text.
    sckUDP.SendData vByteArray
    
    'Sleep 5
    DoEvents
    
    LogInfo "SendUdpData", "Rport: " & sckUDP.RemotePort _
                            & " Lport: " & sckUDP.LocalPort _
                            & " Sent: " & pMessageText, 5, True
    Debug.Print "SendUdpData() SEND: ", sckUDP.RemoteHost, sckUDP.RemotePort, pMessageText
        
    Exit Function
ErrHand:
    Debug.Print "SendUdpData() Error: " & Err.Number & " " & Err.Description, _
                "State: " & sckUDP.State _
                & " RemoteIP: " & sckUDP.RemoteHost _
                & " LocalIP: " & sckUDP.LocalIP _
                & " Pemote Port: " & sckUDP.RemotePort _
                & " Local Port: " & sckUDP.LocalPort, pMessageText
    LogError "SendUdpData", "Error: " & Err.Number & " " & Err.Description
    Resume Next
    Exit Function
End Function

'Broadcast message has arrived.
'TODO: Refactor, comment and document.
Private Sub sckUDP_DataArrival(ByVal bytesTotal As Long)
    Dim vBytBuf()        As Byte
    Dim vUdpData         As String
    Dim vDataParts()    As String
    Dim vRemoteHostIP    As String
    Dim vReplyString    As String

    On Error GoTo ErrHand
    
    'Work around the Winsock bug, try to find a valid
    'looking IP address for the remote host.
    vRemoteHostIP = ChooseValidIP(sckUDP.RemoteHostIP, sckUDP.RemoteHost)
    
    'Don't talk to banned players no matter if I am a host or client.
    If Len(vRemoteHostIP) > 0 _
    And netMain.IsBanned(vRemoteHostIP) Then
        On Error Resume Next
        
        'Flush the data and bail out.
        sckUDP.GetData vBytBuf(), vbByte, 1
        DoEvents
        Exit Sub
        
    End If
    
    On Error GoTo ErrHand
    
    'Get the data recieved by the UDP socket.
    sckUDP.GetData vBytBuf()
    
    Debug.Print ">> " & ByteArrayToHexString(vBytBuf)
    
    'Decrypt and convert to a string.
    vUdpData = gGsLeUtils.LE6d(ByteArrayToHexString(vBytBuf), 17, 19, 3)
    
    'Write to the log at medium priority.
    LogInfo "sckUDP_DataArrival", "Rport: " & sckUDP.RemotePort _
                                & " Lport: " & sckUDP.LocalPort _
                                & " Recieved: " & vUdpData, 5, True
    Debug.Print "sckUDP_DataArrival():", sckUDP.State, sckUDP.RemoteHost, sckUDP.RemoteHostIP, vUdpData
    
    'Split the UDP data string on commas.
    vDataParts = Split(vUdpData, ",")
    
    'Bail out here if no data to prevent errors below.
    If UBound(vDataParts) <= 0 Then
        Exit Sub
    End If
    
    'Sort out the WinSock bug. It sometimes does not return a valid IP address,
    'they are either completley missing or missing some octets.
    vRemoteHostIP = ChooseValidIP(vRemoteHostIP, vDataParts(eSessData.Host_IP))
    
    If netMain.optHost.Value Then
        
        'I am a host. Client sent a request for session info.
        'Format the session details into a string.
        vReplyString = ReplyToClientBroadcast(vDataParts)
        
        'Send the session details to the client.
        Call SendUdpData(vRemoteHostIP, vDataParts(eSessData.Host_Port), vReplyString)
        
    Else
    
        'I am a client. Received data from a host. Add to the session locator.
        Call BroadcastReplyRecieved(vDataParts)
        
    End If
    Exit Sub
ErrHand:
    Debug.Print "sckUDP_DataArrival() Error: " & Err.Number & " " & Err.Description, _
                "State: " & sckUDP.State _
                & " RemoteIP: " & sckUDP.RemoteHost _
                & " LocalIP: " & sckUDP.LocalIP _
                & " Pemote Port: " & sckUDP.RemotePort _
                & " Local Port: " & sckUDP.LocalPort, vUdpData
    LogError "sckUDP_DataArrival", "Error: " & Err.Number & " " & Err.Description
    Resume Next
    Exit Sub
End Sub

'A UDP error was detected. Notify the user and log the error.
Private Sub sckUDP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    
    'Not an error so ifnore and bail.
    If Number = sckSuccess Then
        Exit Sub
    End If
    
    'Notify the user and log to the error log.
    netMain.WriteText "Broadcast error " & Number & " " & Description & ": Source" & Source, False
    LogError "sckUDP_Error", Number & ": " & Description & ": Source" & Source
End Sub

'Return List Index of the passed session ID, -1 if not in found in the session listview.
Private Function IPInetListIndex(pHostID As String) As Long
    Dim vListItem As ListItem
    Dim vItemIndex As Long
    Dim vLvKeyParts() As String
    
    'Assume that the session is not in the list.
    IPInetListIndex = -1
    
    With ListView1
    
    'For each session in the list.
    For vItemIndex = 1 To .ListItems.Count
        
        'Set the session to the object variable.
        Set vListItem = .ListItems(vItemIndex)
        
        'Split the key elements.
        vLvKeyParts = Split(vListItem.Key, ",")
        
        'Check to ensure no errors will be thrown.
        If UBound(vLvKeyParts) = eLvKeyData.DataUbound Then
            
            'Compare the host ID.
            If vLvKeyParts(eLvKeyData.Session_ID) = pHostID Then
                
                'Found the session in the list. Return the item index.
                IPInetListIndex = vItemIndex
                Exit For
                
            End If
        End If
    Next
    End With
End Function

'Ping each session host IP and port in the passed ping list. The
'list is a CR delimited list of IP and Port numbers seperated by a
'comma. The order of the passed list is randomised using function
'ShuffleString(). This seems to vastly improve the UDP responses
'of the session hosts.
Private Sub ProcessUdpPingList(pPingList As String)
    Dim vPingList As String
    Dim vPingArray() As String
    Dim vIndex As Long
    
    On Error Resume Next
    
    'Shuffle the order of the passed ping list. and create an array.
    vPingList = ShuffleString(pPingList, vbCrLf)
    vPingArray = Split(vPingList, vbCrLf)
    
    'For each entry in the list.
    For vIndex = 0 To UBound(vPingArray)
        
        'Send a UDP ping to the comma delimited IP and port number.
        Call SendClientUdpRequest(GetListElement(vPingArray(vIndex), 0), _
                                    GetListElement(vPingArray(vIndex), 1))
        DoEvents
        
    Next
    
    
End Sub

'Try to contact every IP address of every Internet session session hosts in the Listview.
'Sessions usually have multiple IP addresses with the IP address with the fastest ping time
'at the front of the list. This is done by function ArrangeBestIpPingTime(). The order is
'however randomised because that seems to vastly improve the ping response of all contactable
'session hosts.
Private Sub PingInternetHosts()
    Dim vListItem As ListItem
    Dim vIpIndex As Long
    Dim vWaitIndex As Long
    Dim vSessionIndex As Long
    Dim vKeyParts() As String
    Dim vKeyIPs() As String
    Dim vPingList As String
    
    On Error GoTo ErrHand
    
    With ListView1
    vPingList = ""
    
    'For every session listed in the Listview.
    For vSessionIndex = 1 To .ListItems.Count
        
        'Set the Listview session to the object variable.
        Set vListItem = .ListItems(vSessionIndex)
        
        'Split the session key.
        vKeyParts = Split(vListItem.Key, ",")
        
        'Only ping Internet sessions because LAN sessions are pinged by a broadcast.
        If UBound(vKeyParts) = eLvKeyData.DataUbound _
        And vKeyParts(eLvKeyData.Session_ID) <> "LAN" Then
            
            'Get the IP list.
            vKeyIPs = Split(vKeyParts(eLvKeyData.Host_IP), "x")
            
            'For each IP.
            For vIpIndex = 0 To UBound(vKeyIPs)
                
                'Add the address and port to the destination ping list.
                vPingList = vPingList & vKeyIPs(vIpIndex) & "," _
                            & vKeyParts(eLvKeyData.Host_UDP_Port) & vbCrLf
                
            Next
        End If
    Next
    
    'Process the destination ping list.
    Call ProcessUdpPingList(vPingList)
    
    End With
    
    Exit Sub
ErrHand:
    Debug.Print "PingInternetHosts(): ERROR:" & sckUDP.State, Err.Number, Err.Description
    LogError "PingInternetHosts", "Error: " & Err.Number & " " & Err.Description
    Resume Next
End Sub

'I am a client. Send a broadcast across the LAN looking for hosts. Can
'also be used to send a UDP query to a specific host. The order of operations
'with the UDP socket seem to work best. A bit of Voodoo coding.
Private Sub SendClientUdpRequest( _
Optional pRemoteIP As String = cgBroadcastAddress, _
Optional pRemotePort As String = "0")
    Static sHackCounter As Long
    Dim vPort As Long
    Dim vWaitIndex As Long
    Dim vPacketData As String
    Dim vByteArray() As Byte
    Dim vLocalIp As String
    
    On Error Resume Next
    
    'Get the starting local port number.
    vPort = 0
    
    'Close and reopen the UDP port.
    DoEvents
    
    vLocalIp = ChooseValidIP(InetSes.ReportedIP, sckUDP.LocalIP)
    
    'Set the remote address and local & remote ports.
    sckUDP.RemoteHost = pRemoteIP
    sckUDP.RemotePort = CLng(pRemotePort)
    'sckUDP.LocalPort = InetSes.LocalUdpPort
    
    If sckUDP.LocalPort <> InetSes.LocalUdpPort Then
        sckUDP.Close
        sckUDP.LocalPort = InetSes.LocalUdpPort
    End If
    
    'Format the UDP text.
    '"RequestDetail, Local_IP, Listening_Port, Time_Stamp, 'X'"
    vPacketData = "RequestDetail," _
                    & vLocalIp & "," _
                    & sckUDP.LocalPort _
                    & "," & EncodeNonAscii(GetTimeStamp) & "," _
                    & "X"
    
    'Encrypt and convert to a hex encoded byte array.
    Call HexStringToByteArray(vByteArray, gGsLeUtils.LE6(vPacketData, 17, 19, 3))
    
    'Send the UDP text.
    sckUDP.SendData vByteArray
    
    LogInfo "SendClientUdpRequest", "Rport: " & sckUDP.RemotePort _
                                & " Lport: " & sckUDP.LocalPort _
                                & " Sent: " & vPacketData, 5, True
    
    Debug.Print " -- "; vLocalIp, sckUDP.LocalPort
    Debug.Print "SendClientUdpRequest(): Sent: ", sckUDP.RemoteHost, sckUDP.RemotePort, _
                vPacketData
    Debug.Print ">> " & gGsLeUtils.LE6d(ByteArrayToHexString(vByteArray), 17, 19, 3)
    Exit Sub
ErrHand:
    
    'Some error has occured. Log it and try to remediate. Could be the
    'UDP local port setting.
    Debug.Print "SendClientUdpRequest(): sckUDP state: " & sckUDP.State, Err.Number, Err.Description
    
    'Local port may be in use, try to use another port.
    'The error numbers no longer seem to work so we will have to hack
    'this for win 7. sHack counter is to stop endless loops.
    sHackCounter = sHackCounter + 1
    If sHackCounter Mod 10 <> 9 Then
        
        'Try to reopen the broadcast port to a random number selected by the OS.
        sckUDP.Close
        DoEvents
        sckUDP.LocalPort = 0
        LogError "SendClientUdpRequest", _
                    "Resetting the broadcast local port: " _
                    & Err.Number & " " & Err.Description
        Resume
        
    Else
        
        'We have tried 9 times, give up.
        sHackCounter = 0
        netMain.WriteText "Could not open the broadcast port after multiple attempts. " _
                    & Err.Description, True
        LogError "SendClientUdpRequest", _
                    "Error: Could not open the broadcast port after multiple attempts: " _
                    & Err.Number & " " & Err.Description
        
        Exit Sub
    End If
End Sub

'I am a client. "Find War" button has been clicked. Show the Session Locator (this form)
'and begin broadcasting for a host on 255.255.255.255.
'Called from netMain.cmdConnect() event only.
Public Sub BroadcastFindHost()
    On Error GoTo ErrHand
    
    'Display the sesion locator.
    Me.Show , TheMainForm
    DoEvents
    sckUDP.LocalPort = InetSes.LocalUdpPort
    sckUDP.Bind
    
    'Send a broadcast to any hosts on the LAN.
    Call SendClientUdpRequest(cgBroadcastAddress, InetSes.LocalUdpPort)
    
    'Resend broadcast message every 2 seconds.
    InetSes.RebroadcastCntr = 0
    tmrBroadcast.Interval = 2000
    tmrBroadcast.Enabled = True
    Exit Sub
ErrHand:
    Debug.Print "BroadcastFindHost() Error: " & Err.Number & " " & Err.Description, _
                "State: " & sckUDP.State _
                & " RemoteIP: " & sckUDP.RemoteHost _
                & " LocalIP: " & sckUDP.LocalIP _
                & " Pemote Port: " & sckUDP.RemotePort _
                & " Local Port: " & sckUDP.LocalPort
    LogError "BroadcastFindHost", "Error: " & Err.Number & " " & Err.Description
    Resume Next
    Exit Sub
End Sub

'I am a host. Open the UDP port and listen for client broadcasts.
'Called from netMain.cmdConnect_Click().
Public Sub BroadcastListen()
    Dim vPort As Long
    Dim vErrorText As String
    
    On Error GoTo ErrHand
    'MsgBox "LeType: " & MyNetRegData.LeType & vbCrLf _
            & "LeKey: " & MyNetRegData.LeKey & vbCrLf _
            & "LeSlot: " & MyNetRegData.LeSlot & vbCrLf _
            & "LeSlotSpin: " & MyNetRegData.LeSlotSpin
    
    'Close the UDP socket and set it up in listening mode.
    sckUDP.Close
    DoEvents
    
    'Set the port number.
    vPort = InetSes.LocalUdpPort
    sckUDP.LocalPort = vPort
    
    'Put the port into listening mode.
    sckUDP.Bind
    
    'Check if the port was changed in the error handler and warn
    'the user if it has and this is a LAN only war.
    If sckUDP.LocalPort <> CLng(netMain.txtUdpPort.Text) Then
        
        'The port number is different.
        vErrorText = "** Warning: Port " & netMain.txtUdpPort.Text _
                & " in use. Now broadcasting on port " & CStr(sckUDP.LocalPort) _
                & ". Ensure this port is NATed correctly. LAN clients will need " _
                & "to enter this number into their network settings."
        netMain.WriteText vErrorText & vbCrLf, True
        LogError "BroadcastListen", "Port number was changed by the error handler."
        If Not gHeadlessMode Then
            MsgBox vErrorText, vbExclamation, "Broadcast Port Warning"
        End If
        
        'Update the global variables wit the new port number.
        InetSes.LocalUdpPort = sckUDP.LocalPort
        netMain.txtUdpPort.Text = Trim(CStr(sckUDP.LocalPort))
    
    End If
    
    'Notify the user that the broadcast port is open.
    netMain.WriteText "Broadcast port open.", True
    
    Exit Sub
ErrHand:
    'LogError "BroadcastListen", "Error: " & Err.Number & " " & Err.Description
    
    'If 'Address in use' error, try to change the port number.
    If Err.Number = 10048 Then
        
        'Port in use, try to use another port.
        'Random port number selected, try another random number.
        sckUDP.Close
        DoEvents
        sckUDP.LocalPort = 0
        LogError "BroadcastListen", "UDP port reset to a random value."
        Resume
            
    Else
        netMain.WriteText "Could not open the broadcast port. " & Err.Description, True
        LogError "BroadcastListen", "Error: " & Err.Number & " " & Err.Description
        Exit Sub
    End If
End Sub
