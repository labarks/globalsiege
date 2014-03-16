VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form netInetGateway 
   BorderStyle     =   0  'None
   Caption         =   "netInetGateway"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrInet1Queue 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   120
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   1440
      Top             =   120
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   20
   End
End
Attribute VB_Name = "netInetGateway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------------------------------------
'netInetGateway     12/11/2011
'Handle all communications to the Internet Indexing Server. Also
'contains the Internet Indexing command queue and timer.
'-------------------------------------------------------------------------------------------------------------

'The list of queued commands and requests for the Indexing Server.
Public gGatewayQueue As String

'Private variables.
'Headder returned from the web page.
Private gHeadderText As String

'Body of the web page.
Private gBodyText As String

'Error text from attempting to connect to the web page.
Private gErrorText As String

'TRUE if waiting for a web page to load.
Private gStillExecuting As Boolean

'Return the headder text from the last request.
Public Property Get GetHeadderText() As String
    GetHeadderText = gHeadderText
End Property

'Return the error text from the last request.
Public Property Get GetErrorText() As String
    GetErrorText = gErrorText
End Property

'Return TRUE if the Internet control is busy and cannot be used.
Public Property Get IsStillExecuting() As Boolean
    IsStillExecuting = gStillExecuting
End Property

'Form Load event handler.
Private Sub Form_Load()
    tmrTimeout.Enabled = False
End Sub

'If Inet1 was busy during an attempt to contact the Index Server, the URL string
'gets added to tmrInet1Queue's tag and will attempt to send at a later time.
Private Sub tmrInet1Queue_Timer()
    On Error Resume Next
    
    Call IxServerCheckCommandQueue
End Sub

'http://www.w3.org/TR/html401/interact/forms.html#h-17.13.4.1
'application/x-www-form-urlencoded
'This is the default content type. Forms submitted with this content type must be encoded as follows:
'Control names and values are escaped. Space characters are replaced by `+', and then reserved characters are escaped as described in [RFC1738], section 2.2: Non-alphanumeric characters are replaced by `%HH', a percent sign and two hexadecimal digits representing the ASCII code of the character. Line breaks are represented as "CR LF" pairs (i.e., `%0D%0A').
'The control names/values are listed in the order they appear in the document. The name is separated from the value by `=' and name/value pairs are separated from each other by `&'.
'
' netInetGateway.Test1
Public Function Test1() As String
    Dim vQuery As String
    Dim i As Long
    
    
    vQuery = "a=0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789" & vbCrLf
    For i = 0 To 10
        vQuery = vQuery & CStr(i) & " 0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789" & vbCrLf
    Next
    
    Call ExecuteRequest("http://globalsiege.net/xtest1/", vQuery)
    'Call ExecuteRequest("http://10.1.1.102/gs/xtest1/", vQuery)
    
    Debug.Print vbCrLf & "Headder: " & vbCrLf & gHeadderText
    Debug.Print vbCrLf & "Body: " & gBodyText
    Debug.Print vbCrLf & "Error: " & gErrorText
End Function

'netInetGateway.ExecuteRequest()
'
'Test string: ?netinetgateway.ExecuteRequest("http://10.1.1.102/gs/v00090200/indexserver/?460CD69808BA5E0043942F3C91CDF8D6AD3B5E81431AB7C008F4E70B49946E609538E1FADD3D9D7119AEB30507FA2334B17584703719EFFA1BB1B51C40842D2FA5C622A1B609913FE8902033854367ADDEE2822A4C824D2699F4CDA270459B4E2A92C527B344776DAD07DBED17B8B6566489483893C017A9510655703334ECEDC18360729A2AD1EF363D870E4AB51FC09E3F51877A045CCAF382185A7C22C983E2C48E627F963426F0C40ADB597A9DC81BE8C827884B248CEBD9BA1251A021D9E406526B7D46A909C69D3A4577674687DCED99215B7B491392FCFD93556B79563A","Welcome_Msg=abc123")
'
'Connect to the Internet Indexing Server application web page with the passad URL
'and send the string passed by pPostText as a POST request. Wait for the reply and
'return the whole body text. Headder text and error text are storred in the local
'global variables gHeadderText and gErrorText. These read only properties that can
'be accessed externally by GetHeadderText() and GetErrorText(). Private global
'gStillExecuting is set to TRUE while the Internet control is busy and can be accessed
'externally using property IsStillExecuting().
Public Function ExecuteRequest(pUrl As String, Optional pPostText As String = "") As String
    Dim vMethod As String
    Dim vContentType As String
    
    'Bail if the Internet control is busy. This should not
    'ever happen because because IsStillExecuting() should
    'be checked before calling this function.
    If Not gStillExecuting And Not Inet1.StillExecuting Then
        
        'Set the connection method and the content type
        'depending on if there is any post text.
        If Trim(pPostText) = "" Then
            vMethod = "GET"
            vContentType = ""
        Else
            vMethod = "POST"
            vContentType = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
        End If
        
        'No the Internet control is set to busy.
        gStillExecuting = True
        
        'Clear all results and prepare for use.
        gErrorText = ""
        gHeadderText = ""
        gBodyText = ""
        
        'Set up the Internet control.
        Inet1.Protocol = icHTTP
        Inet1.URL = pUrl
        
        'Connect to the passed URL and send the passed post text via POST.
        Inet1.Execute , "POST", pPostText, vContentType
        
        'Timer to kill can go here. No evidence of any need has been found yet.
        'Wait for the Internet control to finish.
        Do While Inet1.StillExecuting
            
            'Sleep stopps the CPU from being hammered with a dead-loop.
            Call Sleep(50)
            DoEvents
        
        Loop
        
        'The body text is filled out in the Inet1_StateChanged() event handler.
        ExecuteRequest = gBodyText
        
        'Notify that the Internet control is ready for the next command.
        gStillExecuting = False
        
    
    End If
End Function

'This is where the web page body text, headder text and error text are filled in.
'-------------------------------------------------------------------------------------------------------------
'Inet1 states:
'
'Value Constant                 Description
' 0  icNone                  'No state to report.
' 1  icHostResolvingHost     'The control is looking up the IP address of the specified host computer.
' 2  icHostResolved          'The control successfully found the IP address of the specified host computer.
' 3  icConnecting            'The control is connecting to the host computer.
' 4  icConnected             'The control successfully connected to the host computer.
' 5  icRequesting            'The control is sending a request to the host computer.
' 6  icRequestSent           'The control successfully sent the request.
' 7  icReceivingResponse     'The control is receiving a response from the host computer.
' 8  icResponseReceived      'The control successfully received a response from the host computer.
' 9  icDisconnecting         'The control is disconnecting from the host computer.
' 10 icDisconnected          'The control successfully disconnected from the host computer.
' 11 icError                 'An error occurred in communicating with the host computer.
' 12 icResponseCompleted     'The request has completed and all data has been received.
'-------------------------------------------------------------------------------------------------------------
Private Sub Inet1_StateChanged(ByVal State As Integer)
    Dim vBodyTextChunk As Variant
    Dim vInet1State As String
    
    On Error Resume Next
    
    Select Case State
    Case icNone
    '0 No state to report.
        vInet1State = "No state to report."
    Case 1  'icHostResolvingHost <- the constant doesn't actually exist so we have to use the number.
    '1 Looking up the IP address of the specified host computer.
        vInet1State = "1 Looking up the IP address of the specified host computer."
        
    Case icHostResolved
    '2 Successfully found the IP address of the specified host computer.
        vInet1State = "2 Successfully found the IP address of the specified host computer."
        
    Case icConnecting
    '3 Connecting to the host computer.
        vInet1State = "3 Connecting to the host computer."
        
    Case icConnected
    '4 Successfully connected to the host computer.
        vInet1State = "4 Successfully connected to the host computer."
        
    Case icRequesting
    '5 Sending a request to the host computer.
        vInet1State = "5 Sending a request to the host computer."
        
    Case icRequestSent
    '6 Successfully sent the request.
        vInet1State = "6 Successfully sent the request."
        
    Case icReceivingResponse
    '7 Receiving a response from the host computer.
        vInet1State = "7 Receiving a response from the host computer."
        
    Case icResponseReceived
    '8 Successfully received a response from the host computer.
        vInet1State = "8 Successfully received a response from the host computer."
        
    Case icDisconnecting
    '9 Disconnecting from the host computer.
        vInet1State = "9 Disconnecting from the host computer."
        
    Case icDisconnected
    '10 Successfully disconnected from the host computer.
        vInet1State = "10 Successfully disconnected from the host computer."
        
    Case icError
    '11 An error occurred in communicating with the host computer.
        vInet1State = "11 An error occurred in communicating with the host computer: " & Inet1.ResponseInfo
        
        'Enter the error text into the global error variable and log the error.
        'Also report the error to the end user.
        gErrorText = "Error connecting to the game server: " & Inet1.ResponseInfo
        LogError "Inet1_StateChanged", "icError: " & gErrorText
        netMain.WriteText gErrorText, False
        Debug.Print "Error headder: " & vbCrLf & Inet1.GetHeader
        
    Case icResponseCompleted
    '12 The request has completed and all data has been received.
        vInet1State = "12 The request has completed and all data has been received."
        
        'Get the headder text. This contains the flags 200 OK or 401 error etc.
        gHeadderText = Inet1.GetHeader
        
        'Build the body text from scratch.
        gBodyText = ""
        
        'Get first chunk of the body text.
        vBodyTextChunk = Inet1.GetChunk(1024, icString)
        
        DoEvents
        
        'Get the rest of the chunks of the body text.
        Do While Len(vBodyTextChunk) > 0
            gBodyText = gBodyText & vBodyTextChunk
            
            DoEvents
            
            'Get the next chunk.
            vBodyTextChunk = Inet1.GetChunk(1024, icString)
        Loop
        
        'Debug.Print gHeadderText
        'Debug.Print gBodyText
    End Select
    'Debug.Print vInet1State
    LogInfo "Inet1_StateChanged", "State: " & vInet1State, 5
End Sub
