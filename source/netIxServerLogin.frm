VERSION 5.00
Begin VB.Form netIxServerLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "<Var.ExeName> Login"
   ClientHeight    =   2775
   ClientLeft      =   2835
   ClientTop       =   3360
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1639.562
   ScaleMode       =   0  'User
   ScaleWidth      =   3887.236
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLostPasssword 
      Caption         =   "L&ost Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Retrieve your password"
      Top             =   2400
      Width           =   1860
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Create a new free GlobalSiege account"
      Top             =   2160
      Width           =   1860
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Top             =   1335
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Log In"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "Log into your GlobalSiege account"
      Top             =   2160
      Width           =   1980
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1725
      Width           =   2325
   End
   Begin VB.Label lblLogin 
      Caption         =   $"netIxServerLogin.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   1740
      Width           =   855
   End
End
Attribute VB_Name = "netIxServerLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Where to return after the login has been validated.
Private gCallBackFunction As String

'Set the return location after a successful login.
Property Let SetCallback(pCallbackFunction)
    gCallBackFunction = pCallbackFunction
End Property

'Register a new account. Open the globalsiege.net account web page.
Private Sub cmdRegister_Click()
    'http://www.globalsiege.net/wp-login.php?action=lostpassword
    On Error Resume Next
    Call TheMainForm.OpenWebPage(gRegisterAccountWebPage)
    Me.Hide
End Sub

'Forgot password. Open the globalsiege.net account web page.
Private Sub cmdLostPasssword_Click()
    'http://www.globalsiege.net/wp-login.php?action=lostpassword
    On Error Resume Next
    Call TheMainForm.OpenWebPage(gLostPasswordWebPage)
    Me.Hide
End Sub

'Actions for the OK click event.
Private Sub cmdOK_Click()
    On Error Resume Next
    
    'The number of attempts to log in.
    Static vCountAttempts As Long
    
    'Save the user name and password in form.netMain
    netFindHosts.tmrBroadcast.Enabled = False
    netMain.txtUserName.Text = txtUserName.Text
    netMain.txtUserName.Tag = txtPassword.Text
    SaveSetting gcApplicationName, "settings", "IxPasswordHash", gGsLeUtils.LE6(netMain.txtUserName.Tag)
    SaveSetting gcApplicationName, "settings", "IxAccountName", txtUserName.Text
    netMain.WriteText "Validating login...", True
    
    'Hide the form from view.
    Me.Hide
    
    'Check the passed login details.
    If IxServerCheckLogin Then
        
        'Login was successful. Check if there is a callback to execute.
        If gCallBackFunction = "PostInetHostDetails" Then
            Call netMain.ConnectDisconnectBeginSession
        ElseIf gCallBackFunction = "IxServerListSession" Then
            Call netMain.DisplaySessionLocator
        End If
        
        Unload Me
    
    Else
        
        'Login attempt failed. Update login count and bail if too many.
        vCountAttempts = vCountAttempts + 1
        If vCountAttempts >= 3 Then
            netMain.WriteText "Invalid login.", True
            Unload Me
        Else
            Me.Show
            netMain.WriteText "Invalid username or password, try again!", True
            MsgBox "Invalid username or password, try again.", , "Login"
            txtPassword.SetFocus
            Call SelectAndHighlightText(txtPassword)
        End If
        
    End If
End Sub

'Set up the caption and hilight the username.
Private Sub Form_Load()
    Me.Caption = SubstituteStringTokens(Me.Caption)
    lblLogin.Caption = SubstituteStringTokens(lblLogin.Caption)
    
    'Stop the session locator hanging around.
    If netFindHosts.Visible Then
        Unload netFindHosts
    End If
    
    'Set selected text.
    If Trim(netMain.txtUserName.Text) <> "" Then
        txtUserName.Text = Trim(netMain.txtUserName.Text)
        Me.Show
        txtPassword.SetFocus
        Call SelectAndHighlightText(txtPassword)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TheMainForm.SetFocus
End Sub

'Select all username text.
Private Sub txtUserName_GotFocus()
    Call SelectAndHighlightText(txtUserName)
End Sub
