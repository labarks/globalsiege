VERSION 5.00
Begin VB.Form netWiz 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Host Multiplayer War"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10290
   Icon            =   "netWiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmScreen 
      Caption         =   "Begin Session"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   4
      Left            =   4320
      TabIndex        =   22
      Top             =   7320
      Width           =   4455
      Begin VB.Label Label4 
         Caption         =   $"netWiz.frx":030A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame frmScreen 
      Caption         =   "Refresh Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Width           =   4455
      Begin VB.ComboBox comboRefresh 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   $"netWiz.frx":03FC
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame frmScreen 
      Caption         =   "Terminal  Name (Nick Name)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   5160
      Width           =   4455
      Begin VB.TextBox txtHostTermName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Text            =   "Terminal Name"
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label Label8 
         Caption         =   $"netWiz.frx":04FF
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame frmScreen 
      Caption         =   "Find Multiplayer Session"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   7
      Left            =   4800
      TabIndex        =   14
      Top             =   4440
      Width           =   4455
      Begin VB.Label Label6 
         Caption         =   $"netWiz.frx":05BA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame frmScreen 
      Caption         =   "Terminal Name (Nick Name)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   6
      Left            =   4800
      TabIndex        =   11
      Top             =   2400
      Width           =   4455
      Begin VB.TextBox txtTermName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Text            =   "Terminal Name"
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label Label7 
         Caption         =   $"netWiz.frx":0676
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame frmScreen 
      Caption         =   "Session Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   5
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox comboSesClient 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Choose to find either an Internet war or a Local Area Network (LAN) war. "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame frmScreen 
      Caption         =   "Session Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   4455
      Begin VB.TextBox txtSes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Text            =   "Session Name"
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Enter a Session Name. The name that you choose here will be used by other players to identify and connect to your new session."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< Back"
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
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame frmScreen 
      Caption         =   "Session Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox comboSes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   $"netWiz.frx":0731
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Step >>"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "netWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ScreenNo As Long
Dim Lockout As Boolean

Private Sub Form_Load()
    Lockout = True
    Me.Width = 4770
    Me.Height = 3570
    
    'Prepare session screen.
    comboSes.AddItem "LAN"
    comboSes.AddItem "Internet"
    If netMain.optLan.Value Then
        comboSes.ListIndex = 0
    Else
        comboSes.ListIndex = 1
    End If
    
    'Prepare Refresh Rate screen.
    comboRefresh.AddItem "High"
    comboRefresh.AddItem "Medium"
    comboRefresh.AddItem "Low"
    With netMain
    comboRefresh.ListIndex = Abs(CInt(.optRefresh(1).Value + .optRefresh(2).Value * 2))
    End With
    
    'Prepare Client Session screen.
    comboSesClient.AddItem "LAN"
    comboSesClient.AddItem "Internet"
    
    If netMain.optLan.Value Then
        comboSesClient.ListIndex = 0
    Else
        comboSesClient.ListIndex = 1
    End If
    
    txtTermName.Text = netMain.txtTerminalName.Text
    txtHostTermName.Text = netMain.txtTerminalName.Text
    
    Call DisplayScreen
    Lockout = False
    DoEvents
End Sub

'Display the screen number denoted by the global variable "ScreenNo"
Private Sub DisplayScreen()
    Dim i As Long
    
    'Clear all screens.
    For i = 0 To frmScreen.Count - 1
        frmScreen(i).Visible = False
    Next
    
    'Display current screen.
    Label6.Caption = SubstituteStringTokens(Label6.Caption)
    frmScreen(ScreenNo).Left = 120
    frmScreen(ScreenNo).Top = 120
    frmScreen(ScreenNo).Visible = True
    
    Select Case ScreenNo
    
    'Session Type.
    Case 0
        'Set as host.
        netMain.optHost.Value = True
        
        Me.Caption = "Host Multiplayer War - Step 1"
        cmdNext.Visible = True
        cmdNext.Caption = "Next Step >>"
        cmdBack.Visible = False
        
    'Session Name.
    Case 1
        Me.Caption = "Host Multiplayer War - Step 2"
        cmdNext.Visible = True
        cmdNext.Caption = "Next Step >>"
        cmdBack.Visible = True
        cmdBack.Caption = "<< Back"
        
        txtSes.Text = netMain.txtSesName.Text
    
    'Host Terminal Name.
    Case 2
        Me.Caption = "Host Multiplayer War - Step 2.5"
        cmdNext.Visible = True
        cmdNext.Caption = "Next Step >>"
        cmdBack.Visible = True
        cmdBack.Caption = "<< Back"
    
    'Refresh Rate.
    Case 3
        Me.Caption = "Host Multiplayer War - Step 3"
        cmdNext.Visible = True
        cmdNext.Caption = "Next Step >>"
        cmdBack.Visible = True
        cmdBack.Caption = "<< Back"
        
        netMain.optRefresh(comboRefresh.ListIndex).Value = True
    
    'Begin.
    Case 4
        Me.Caption = "Host Multiplayer War - Step 4"
        cmdNext.Visible = True
        cmdNext.Caption = "Begin"
        cmdBack.Visible = True
        cmdBack.Caption = "<< Back"
        
    'Client Session Type.
    Case 5
        'Set as host.
        netMain.optJoin.Value = True
        
        Me.Caption = "Find Multiplayer War - Step 1"
        cmdNext.Visible = True
        cmdNext.Caption = "Next Step >>"
        cmdBack.Visible = False
    
    'List sessions.
    Case 6
        Me.Caption = "Find Multiplayer War - Step 1.5"
        cmdNext.Visible = True
        cmdNext.Caption = "Next Step >>"
        cmdBack.Visible = True
        cmdBack.Caption = "<< Back"
    
    'List sessions.
    Case 7
        Me.Caption = "Find Multiplayer War - Step 2"
        cmdNext.Visible = True
        cmdNext.Caption = "List Sessions"
        cmdBack.Visible = True
        cmdBack.Caption = "<< Back"
        
    End Select
End Sub

'Next button pressed.
Private Sub cmdNext_Click()
    
    Select Case ScreenNo
    'Session Type.
    Case 0
        
    'Session Name.
    Case 1
        netMain.txtSesName.Text = Trim(txtSes.Text)
    
    'Host Terminal Name.
    Case 2
        netMain.txtTerminalName.Text = txtHostTermName.Text
    
    'Refresh Rate.
    Case 3
        netMain.optRefresh(comboRefresh.ListIndex).Value = True
        
    'Begin.
    Case 4
        Call netMain.cmdConnect_Click
        Me.Hide
        Unload Me
        Exit Sub
    
    'Client Session Type.
    Case 5
        netMain.optRefresh(comboSesClient.ListIndex).Value = True
    
    'Terminal Name.
    Case 6
        netMain.txtTerminalName.Text = txtTermName.Text
    
    'Find and list available sessions.
    Case 7
        netMain.DisplaySessionLocator
        Me.Hide
        Unload Me
        Exit Sub
        
    End Select
    
    'Display next screen.
    ScreenNo = ScreenNo + 1
    
    'Skip the terminal name screen if internet session.
    If (ScreenNo = 2 Or ScreenNo = 6) And netMain.optInet.Value Then
        ScreenNo = ScreenNo + 1
    End If
    DisplayScreen
End Sub

'Back button pressed.
Private Sub cmdBack_Click()

    'Display previous screen.
    ScreenNo = ScreenNo - 1
    
    'Skip the terminal name screen if internet session.
    If (ScreenNo = 2 Or ScreenNo = 6) And netMain.optInet.Value Then
        ScreenNo = ScreenNo - 1
    End If
    DisplayScreen
End Sub

'Session Type - Screen 0.
Private Sub comboSes_Click()
    If Lockout Then
        Exit Sub
    End If
    
    If comboSes.ListIndex = 0 Then
        netMain.optLan.Value = True
    Else
        netMain.optInet.Value = True
    End If
End Sub

'Session Type - Screen 0.
Private Sub comboSes_Validate(Cancel As Boolean)
    If comboSes.ListIndex = 0 Then
        netMain.optLan.Value = True
    Else
        netMain.optInet.Value = True
    End If
End Sub

'Refresh Rate - Screen 2.
Private Sub comboRefresh_Click()
    If Lockout Then
        Exit Sub
    End If
    
    netMain.optRefresh(comboRefresh.ListIndex).Value = True
End Sub

'Refresh Rate - Screen 2.
Private Sub comboRefresh_Validate(Cancel As Boolean)
    netMain.optRefresh(comboRefresh.ListIndex).Value = True
End Sub

'Session Type Client - Screen 5.
Private Sub comboSesClient_Click()
    If Lockout Then
        Exit Sub
    End If
    
    If comboSesClient.ListIndex = 0 Then
        netMain.optLan.Value = True
    Else
        netMain.optInet.Value = True
    End If
    
End Sub

'Session Type Client - Screen 5.
Private Sub comboSesClient_Validate(Cancel As Boolean)
    If comboSesClient.ListIndex = 0 Then
        netMain.optLan.Value = True
    Else
        netMain.optInet.Value = True
    End If
End Sub

