VERSION 5.00
Begin VB.Form netIxServerAgreement 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "<Var.ExeName> Online Terms & Agreement"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrActivateButons 
      Interval        =   5000
      Left            =   2400
      Top             =   5880
   End
   Begin VB.CommandButton cmdViewOnline 
      Caption         =   "View online"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdDisagree 
      Cancel          =   -1  'True
      Caption         =   "I do not agree"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdAgree 
      Caption         =   "I agree"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox txtAgreement 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "netIxServerAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Where to return after the login has been validated.
Private gCallBackFunction As String

'URL for the online agreement text.
Private gAgreementUrl As String

'Set the return location after a successful login.
Property Let SetCallback(pCallbackFunction)
    gCallBackFunction = pCallbackFunction
End Property

'Set the URL for the online agreement text.
Property Let SetAgreementUrl(pAgreementUrl As String)
    Dim vAgreementText As String
    
    'Set the global variable.
    gAgreementUrl = pAgreementUrl
    
    'Stop the broadcast timer to stop this page from being
    'reloaded while the player is still reading it after a
    'refresh is requested from the Ix Server.
    netFindHosts.tmrBroadcast.Enabled = False
    
    'Get the agreement text from the web page.
    vAgreementText = IxServerComunication(pAgreementUrl, , True)
    
    'Remove HTML tags and format.
    vAgreementText = IxConsistantLineBreaks(vAgreementText, vbCrLf & vbCrLf)
    vAgreementText = IxRemoveSomeHtmlTags(vAgreementText)
    
    'Send the text to the text box.
    txtAgreement.Text = vAgreementText
End Property

'"I agree" clicked.
Private Sub cmdAgree_Click()
    On Error Resume Next
    
    Me.Hide
    
    'Notify the Ix Server and check if there is a callback to execute.
    If IxServerTcAgreed Then
        If gCallBackFunction = "PostInetHostDetails" Then
            Call netMain.ConnectDisconnectBeginSession
        ElseIf gCallBackFunction = "IxServerListSession" Then
            Call netMain.DisplaySessionLocator
        End If
    End If
    Unload Me
End Sub

'"I do not agree" clicked.
Private Sub cmdDisagree_Click()
    On Error Resume Next
    
    netMain.optLan.Value = True
    netMain.WriteText "You do not agree to the Terms and Conditions for online use. No matter, the computer players will keep you company."
    Unload Me
End Sub

'Show the agreement text in a web page.
Private Sub cmdViewOnline_Click()
    On Error Resume Next
    Call TheMainForm.OpenWebPage(gAgreementUrl)
End Sub

'Set up the caption.
Private Sub Form_Load()
    Me.Caption = SubstituteStringTokens(Me.Caption)
    
    'Stop the session locator hanging around.
    If netFindHosts.Visible Then
        Unload netFindHosts
    End If
    
    'Activate the buttons after a short wait.
    cmdDisagree.Enabled = False
    cmdAgree.Enabled = False
    tmrActivateButons.Enabled = True
End Sub

'Activate the agree and disagree buttons after a delay. This is to stop players
'agreeing straight away and causing the "Please wait" message to come up when
'reconnecting to the IxServer in netMain.cmdConnect_Click().
Private Sub tmrActivateButons_Timer()
    cmdDisagree.Enabled = True
    cmdAgree.Enabled = True
End Sub
