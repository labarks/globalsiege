VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form netMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Network Administration Panel"
   ClientHeight    =   11175
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   11175
   ScaleWidth      =   15720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrAcknowledge 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   120
      Top             =   7200
   End
   Begin VB.Frame frameInfo 
      BorderStyle     =   0  'None
      Height          =   2775
      Index           =   3
      Left            =   2040
      TabIndex        =   57
      Top             =   6240
      Width           =   4095
      Begin RichTextLib.RichTextBox rtIpConfig 
         Height          =   2655
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4683
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Rmote.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer tmrControlChange 
      Interval        =   10000
      Left            =   600
      Top             =   6720
   End
   Begin VB.Timer tmrCheckGSNews 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   6720
   End
   Begin VB.Timer tmrForfeitTurn 
      Interval        =   60000
      Left            =   1080
      Top             =   6720
   End
   Begin VB.Timer tmrFillInfo 
      Interval        =   2500
      Left            =   1080
      Top             =   6240
   End
   Begin VB.Frame frameInfo 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   2
      Left            =   6480
      TabIndex        =   38
      Top             =   7320
      Width           =   4095
      Begin VB.PictureBox pctNetConnections 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   100
         ScaleHeight     =   3015
         ScaleWidth      =   3975
         TabIndex        =   52
         Top             =   120
         Width           =   3975
         Begin VB.Frame frmHostOptions 
            Caption         =   "Actions"
            Height          =   735
            Left            =   0
            TabIndex        =   53
            Top             =   2160
            Width           =   3855
            Begin VB.PictureBox pctNetVoteActions 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   100
               ScaleHeight     =   375
               ScaleWidth      =   3735
               TabIndex        =   54
               Top             =   240
               Width           =   3735
               Begin VB.CommandButton cmdBan 
                  Caption         =   "Kick && Ban"
                  Enabled         =   0   'False
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
                  Left            =   2520
                  TabIndex        =   31
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.CommandButton cmdKill 
                  Caption         =   "Kick Connection"
                  Enabled         =   0   'False
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
                  Left            =   1200
                  TabIndex        =   30
                  Top             =   0
                  Width           =   1335
               End
               Begin VB.CommandButton cmfForfeit 
                  Caption         =   "Forfeit Turn"
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
                  Left            =   0
                  TabIndex        =   29
                  Top             =   0
                  Width           =   1095
               End
            End
         End
         Begin MSComctlLib.ListView lsvConnections 
            Height          =   2055
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Connection"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Terminal Name"
               Object.Width           =   3069
            EndProperty
         End
      End
   End
   Begin VB.Frame frameInfo 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   1
      Left            =   6480
      TabIndex        =   37
      Top             =   3840
      Width           =   4095
      Begin VB.PictureBox pctNetHostOptions 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   100
         ScaleHeight     =   3015
         ScaleWidth      =   3975
         TabIndex        =   49
         Top             =   120
         Width           =   3975
         Begin VB.VScrollBar vscrollTimeLimit 
            Height          =   375
            LargeChange     =   30
            Left            =   480
            Max             =   0
            Min             =   990
            SmallChange     =   30
            TabIndex        =   26
            Top             =   2230
            Value           =   180
            Width           =   255
         End
         Begin VB.TextBox txtTimeLimit 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "180"
            ToolTipText     =   "Restrict the number of connections from the same IP address"
            Top             =   2280
            Width           =   480
         End
         Begin VB.VScrollBar vscrollMaxConnections 
            Height          =   375
            Left            =   240
            Max             =   0
            Min             =   9
            TabIndex        =   20
            Top             =   500
            Value           =   6
            Width           =   255
         End
         Begin VB.VScrollBar vscrollMaxPlayers 
            Height          =   375
            Left            =   240
            Max             =   0
            Min             =   6
            TabIndex        =   18
            Top             =   20
            Value           =   6
            Width           =   255
         End
         Begin VB.TextBox txtPassword 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   22
            ToolTipText     =   "Password required for players to connect"
            Top             =   960
            Width           =   2775
         End
         Begin VB.CheckBox chkPasswordSession 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   21
            ToolTipText     =   "Clients can only connect if they enter a password"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtMaxPlayers 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "3"
            ToolTipText     =   "Restrict the number of players each terminal can claim"
            Top             =   40
            Width           =   250
         End
         Begin VB.TextBox txtMaxConnections 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "1"
            ToolTipText     =   "Restrict the number of connections from the same IP address"
            Top             =   520
            Width           =   250
         End
         Begin VB.CheckBox chkHideSession 
            Caption         =   "Hide this session from the Internet"
            Enabled         =   0   'False
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
            Left            =   0
            TabIndex        =   27
            ToolTipText     =   "Hide this session from the Internet preventing new players from connecting."
            Top             =   2760
            Width           =   3375
         End
         Begin VB.TextBox txtWelcomeMsg 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   1080
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            ToolTipText     =   "Welcome message displayed when a client connects"
            Top             =   1240
            Width           =   2775
         End
         Begin VB.CheckBox chkWlcmMsg 
            Caption         =   "Welcome Message"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   23
            ToolTipText     =   "Display welcome message when a client connects"
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblTimeLimit 
            Caption         =   "Idle time limit per turn (seconds)"
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
            Left            =   840
            TabIndex        =   55
            Top             =   2280
            Width           =   3015
         End
         Begin VB.Label lblMaxClaim 
            Caption         =   "Maximum number of armies each terminal can claim."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   51
            ToolTipText     =   "Restrict the number of players each terminal can claim"
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label lblMaxCon 
            Caption         =   "Maximum number of connections per IP address."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   50
            ToolTipText     =   "Restrict the number of connections from the same IP address"
            Top             =   480
            Width           =   2535
         End
      End
   End
   Begin VB.Frame frameInfo 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   0
      Left            =   6480
      TabIndex        =   36
      Top             =   120
      Width           =   4095
      Begin VB.PictureBox pctNetSettings 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   100
         ScaleHeight     =   3015
         ScaleWidth      =   3975
         TabIndex        =   44
         Top             =   120
         Width           =   3975
         Begin VB.CommandButton cmdChangeLogin 
            Caption         =   "Account Settings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   13
            ToolTipText     =   "Change your log in details"
            Top             =   1800
            Width           =   2295
         End
         Begin VB.TextBox txtUserName 
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
            Height          =   315
            Left            =   0
            TabIndex        =   12
            Text            =   "User name"
            ToolTipText     =   "Your GlobalSiege user name"
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CheckBox chkNetLog 
            Caption         =   "Log network activity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1200
            TabIndex        =   10
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtUdpPort 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   15
            Text            =   "4813"
            ToolTipText     =   "All terminals must use the same port number"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtSesName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            TabIndex        =   11
            Text            =   "Name of War"
            ToolTipText     =   "Enter the name of the session"
            Top             =   840
            Width           =   3855
         End
         Begin VB.TextBox txtTerminalName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            TabIndex        =   9
            Text            =   "Terminal Name"
            ToolTipText     =   "Enter a name for this terminal to be used in this session."
            Top             =   240
            Width           =   3855
         End
         Begin VB.TextBox txtTcpPort 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2400
            TabIndex        =   14
            Text            =   "4813"
            ToolTipText     =   "All terminals must use the same port number"
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton cmdRestore 
            Caption         =   "Restore Defaults"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   16
            ToolTipText     =   "Restore default network and Internet settings. "
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label lblUserName 
            AutoSize        =   -1  'True
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   56
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label Label9 
            Caption         =   "Broadcast"
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
            Left            =   3240
            TabIndex        =   48
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Session Name"
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
            Left            =   0
            TabIndex        =   47
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label lblDisplayName 
            Caption         =   "Display Name"
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
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label txtPortNmbr 
            Caption         =   "Main Port"
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
            Left            =   2400
            TabIndex        =   45
            Top             =   1200
            Width           =   735
         End
      End
   End
   Begin VB.Timer tmrKeepAlive 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   6240
   End
   Begin VB.Timer tmrFlushQueue 
      Enabled         =   0   'False
      Left            =   600
      Top             =   6240
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox pctNetMainButtons 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   4440
         ScaleHeight     =   975
         ScaleWidth      =   1815
         TabIndex        =   43
         Top             =   4440
         Width           =   1815
         Begin VB.CommandButton cmdConnect 
            Caption         =   "&Find Sessions"
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
            Left            =   0
            TabIndex        =   0
            ToolTipText     =   "Connect with these settings"
            Top             =   0
            Width           =   1695
         End
         Begin VB.CommandButton cmdOK 
            Cancel          =   -1  'True
            Caption         =   "&Hide"
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
            Left            =   0
            TabIndex        =   1
            ToolTipText     =   "Hide the network setup dialog box without disconnecting"
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.TextBox txtChat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Lists connected terminals, TCP errors, etc..."
         Top             =   3960
         Width           =   4215
      End
      Begin VB.Frame frameConnectionOptions 
         Height          =   4215
         Left            =   4440
         TabIndex        =   35
         ToolTipText     =   "Network war must have 1 host"
         Top             =   120
         Width           =   1695
         Begin VB.PictureBox pctNetConnectionOptions3 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   1455
            TabIndex        =   42
            Top             =   3000
            Width           =   1455
            Begin VB.OptionButton optRefresh 
               Caption         =   "Slow"
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
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   8
               Top             =   720
               Width           =   1335
            End
            Begin VB.OptionButton optRefresh 
               Caption         =   "Medium"
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
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   7
               Top             =   360
               Width           =   1335
            End
            Begin VB.OptionButton optRefresh 
               Caption         =   "Fast"
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
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.PictureBox pctNetConnectionOptions2 
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   1455
            TabIndex        =   41
            Top             =   1680
            Width           =   1455
            Begin VB.OptionButton optInet 
               Caption         =   "Internet"
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
               Left            =   0
               TabIndex        =   5
               Top             =   480
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton optLan 
               Caption         =   "LAN"
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
               Left            =   0
               TabIndex        =   4
               Top             =   0
               Width           =   1455
            End
         End
         Begin VB.PictureBox pctNetConnectionOptions1 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   1455
            TabIndex        =   39
            Top             =   360
            Width           =   1455
            Begin VB.OptionButton optHost 
               Caption         =   "Host"
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
               Left            =   0
               TabIndex        =   3
               ToolTipText     =   "Become a host and listen for connections"
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton optJoin 
               Caption         =   "Join"
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
               Left            =   0
               TabIndex        =   2
               ToolTipText     =   "Connect to a listening host"
               Top             =   480
               Value           =   -1  'True
               Width           =   1455
            End
         End
         Begin VB.Label lblNetType 
            AutoSize        =   -1  'True
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   120
            TabIndex        =   60
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lblRefreshRate 
            AutoSize        =   -1  'True
            Caption         =   "Refresh"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   120
            TabIndex        =   59
            Top             =   2640
            Width           =   660
         End
         Begin VB.Label lblNetConnectionOptions 
            AutoSize        =   -1  'True
            Caption         =   " Options "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   840
         End
      End
      Begin MSComctlLib.TabStrip tabInfo 
         Height          =   3735
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   6588
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Settings"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Host Options"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Connections"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "IP Config"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSWinsockLib.Winsock sckTCP 
      Index           =   0
      Left            =   120
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "netMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PacketType
    bPacket()    As Byte
    CutOff       As Long
    TerminalTurn As Long
End Type

Const dfltSpeed As Byte = 0
Const xmitMilliSeconds As Long = 10      'Milli seconds between Xmits to same port

Const cNetWelcomeMessageFile = "NetWelcomeMessage.txt"

Private MyTermNo As String
Private connectionSpeed As Byte             '0,1,2 = slowest..fastest
Private lastSendTime(MaxConnections) As Long             'Last time msg was sent to this terminal
Private Packets(MaxConnections) As PacketType
Private ConnectSoundFile As String

Private Sub cmdChangeLogin_Click()
    On Error Resume Next
    netIxServerLogin.Show , Me
End Sub

'Check the website for any GlobalSiege news such as a new version released etc.
Private Sub tmrCheckGSNews_Timer()
    On Error Resume Next
    
    tmrCheckGSNews.Enabled = False
    If TheMainForm.hlpCheckForUpdates.Checked Then
        Call IxServerCheckGsNews(True)
    End If
End Sub

'Write text to session history.
'If PostMesage is true the message will be posted
'using chat box if net text box not visible.
Public Sub WriteText(pText As String, Optional PostMessage As Boolean = True)
    If Len(txtChat.Text) = 0 Then
        txtChat.Text = pText
    Else
        txtChat.Text = txtChat.Text & vbCrLf & pText
    End If
    txtChat.SelStart = Len(txtChat)
    If PostMessage Then
        Call TheMainForm.PostMessage(pText)
    End If
End Sub

'Stop network from stalling.
'Counterpart MyTerminalKicked().
'TODO: Refactor, comment and document.
Public Sub KickNextTerminal(terminal As Long)
    Dim BytBuf() As Byte
    
    ReDim BytBuf(2) As Byte
    
    If netWorkSituation = cNetClient Then
        Call XmitBytes(0, 15, CByte(terminal), BytBuf)
    ElseIf netWorkSituation = cNetHost Then
        Call XmitBytes(CByte(terminal), 15, CByte(terminal), BytBuf)
        '**TODO: Host needs to do something here if he gets a kick.
    End If
    
    LogError "KickNextTerminal", "Terminal " & terminal
End Sub

'Last player has kicked this terminal to prevent network stall.
'Counterpart KickNextTerminal().
'TODO: Refactor, comment and document.
Public Sub MyTerminalKicked(pTermFrom, pLastPlayer As Long, BytBuf() As Byte)
    
    LogError "MyTerminalKicked", "From " & CStr(pTermFrom) & " LastPlayer: " & CStr(pLastPlayer)
    If TheMainForm.SetupScreen.Visible Then
        Exit Sub
    End If
    If gPlayerID(gPlayerTurn).playerWho = 0 Then
        Exit Sub
    End If
    If net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber Then
        If netWorkSituation = cNetHost Then
            Call netMain.requestRefreshHost(pLastPlayer)
        Else
            Call netMain.requestRefresh
        End If
    Else
        If netWorkSituation = cNetHost Then            'If host
            Call netMain.KickNextTerminal(CLng(pTermFrom))
        End If
    End If
End Sub

'I am exchanging these cards
Public Sub ChangeCards(cardsPickedPos() As Byte, pPlayerTurn As Byte)
    Dim BytBuf() As Byte
    Dim i As Long
    
    ReDim BytBuf(4) As Byte
    For i = 0 To 2
        BytBuf(i + 2) = cardsPickedPos(i)
    Next
    If netWorkSituation = cNetClient Then
        Call XmitBytes(0, 12, pPlayerTurn, BytBuf)
    ElseIf netWorkSituation = cNetHost Then
        Call XmitBytesAll(0, 12, pPlayerTurn, BytBuf)
    End If
End Sub

'I am host and have pressed setup Cancle
'TODO: Look into this, consider terminals who have connected while in setup.
Public Sub CancleSetup()
    Dim BytBuf() As Byte
    
    Call TheMainForm.PackForNetRefresh(BytBuf)
    Call XmitBytesAll(0, 11, 0, BytBuf)
    'TheMainForm.TimerWatch.Enabled = True
    '---------------------------------------
    'Dim byteMes() As Byte
    
    'Call TheMainForm.packWarSettings(byteMes)
    'Call XmitBytesAll(0, 14, 0, byteMes)
    
    'Call TheMainForm.PackForNetRefresh(byteMes)
    'Call XmitBytesAll(0, 6, 0, byteMes)
    TheMainForm.TimerWatch.Enabled = True
End Sub

'----------------------------------------------------------------------------------------
'Send the passed message and if pRequireAck is set to TRUE then set up the acknowledge timer.
'If I am a client terminal, send the win message to the host. If I am the host, send the message
'to all the connected terminals except for the source terminal which is passed in pSender. The
'message will be saved in the RemoteNetRegData structure for the relevent remote terminals and
'the acknowledge timer will be set if pRequireAck is set to TRUE. When the acknowledge timer,
'tmrAcknowledge is triggered, it will check if the RemoteNetRegData structure has been reset by
'an acknowledge response (command 25). If not, the message will be sent again for up to three
'times before giving up.
Private Sub SendMessage(pCommand As Byte, _
pBytBuf() As Byte, _
Optional pWinType As Byte = 0, _
Optional pSender As Integer = 0, _
Optional pRequireAck As Boolean = False)
    Dim vBytBuf() As Byte
    
    'Copy the packet in case the original changes while waiting for the ack.
    ReDim vBytBuf(0) As Byte
    Call CopyBytes(vBytBuf, pBytBuf, 0)
    
    'If I am a client.
    If netWorkSituation = cNetClient And pSender <> 0 Then
        
        'Only send to the host.
        Call SendMessageToHost(pCommand, vBytBuf, pWinType, pSender)
    
    'If I am the host.
    ElseIf netWorkSituation = cNetHost Then
        
        'Send to all connected clients except for the sender.
        Call SendMessageToAllClients(pCommand, vBytBuf, pWinType, pSender)
        
    End If
    
End Sub

'Send the passed message to the host terminal and setup the
'acknowledge timer to resend the message if it doesn't get acknowledged.
Private Sub SendMessageToHost(pCommand As Byte, _
pBytBuf() As Byte, _
Optional pWinType As Byte = 0, _
Optional pSender As Integer = 0, _
Optional pRequireAck As Boolean = False)
    Dim vIndex As Long
    
    'Send the message to the host.
    Call XmitBytes(0, pCommand, pWinType, pBytBuf)
    
    'Set up the acknowledge timer. Only applies if the remote
    'terminal is a GS versions greater than 0.9.297
    If pRequireAck And RemoteNetRegData(0).AppVersion > "00090297" Then
        RemoteNetRegData(0).AcknowledgeMessage.RetryCount = 0
        RemoteNetRegData(0).AcknowledgeMessage.RemoteCommand = pCommand
        RemoteNetRegData(0).AcknowledgeMessage.WinType = CStr(pWinType)
        RemoteNetRegData(0).AcknowledgeMessage.BytBuf = pBytBuf
    End If
    
    'Set the ack timer.
    tmrAcknowledge.Enabled = True

End Sub

'Send the passed message to all connected terminals and setup the
'acknowledge timer to resend the message if it doesn't get acknowledged.
Private Sub SendMessageToAllClients(pCommand As Byte, _
pBytBuf() As Byte, _
Optional pWinType As Byte = 0, _
Optional pSender As Integer = 0, _
Optional pRequireAck As Boolean = False)
    Dim vIndex As Long
    Dim vSetAckTimer As Boolean
    
    'Call XmitBytesAll(0, pCommand, pWinType, pBytBuf)
    
    vSetAckTimer = False
    
    'Send the passed message to all connected clients.
    For vIndex = 1 To MaxConnections
        
        'Similar to function XmitBytesAll except for the ack part.
        If vIndex <> CLng(pSender) And sckTCP(vIndex).State = sckConnected Then
            
            'Send the message.
            Call XmitBytes(vIndex, pCommand, pWinType, pBytBuf)
            
            'Set up the acknowledge timer. Only applies if the remote
            'terminal is a GS versions greater than 0.9.297
            If pRequireAck And RemoteNetRegData(vIndex).AppVersion > "00090297" Then
                RemoteNetRegData(vIndex).AcknowledgeMessage.RetryCount = 0
                RemoteNetRegData(vIndex).AcknowledgeMessage.RemoteCommand = pCommand
                RemoteNetRegData(vIndex).AcknowledgeMessage.WinType = CStr(pWinType)
                RemoteNetRegData(vIndex).AcknowledgeMessage.BytBuf = pBytBuf
                vSetAckTimer = True
            End If
        End If
    Next
    
    'Set the ack timer as required.
    tmrAcknowledge.Enabled = vSetAckTimer
End Sub

'To stop a hang when certain messages are not received, these messages need to be acknowledged.
'This timer checks if messages have been acknowledged.
Private Sub tmrAcknowledge_Timer()
    Dim BytBuf() As Byte
    Dim vIndex As Long
    Dim vSetAckTimer As Boolean
    
    On Error Resume Next
    
    vSetAckTimer = False
    
    'If I am a client.
    If netWorkSituation = cNetClient Then
        vSetAckTimer = MessageNotAcknowledged(0)
        
    'If I am the host.
    Else
        'For all potential connections.
        For vIndex = 1 To MaxConnections
            vSetAckTimer = vSetAckTimer Or MessageNotAcknowledged(vIndex)
        Next
    End If

    
    'If IHaveWon was resent.
    If vSetAckTimer Then
        
        'Log it to the error log.
        LogError "tmrAcknowledge_Timer", _
                "Message acknowledgement not recieved."
    'Not sent.
    Else
        
        'Reset all AcknowledgeMessage tags if the timer is to be turned off.
        Call ResetAcknowledgeTags
        
    End If
    
    'Set the timer as required.
    tmrAcknowledge.Enabled = vSetAckTimer
    
End Sub

'Reset all AcknowledgeMessage tags.
Private Sub ResetAcknowledgeTags()
    Dim vIndex As Long
    
    'For each possible terminal.
    For vIndex = 0 To MaxConnections
        With RemoteNetRegData(vIndex).AcknowledgeMessage
        .WinType = 0
        .RemoteCommand = 0
        .RetryCount = 0
        End With
    Next
End Sub

'Resend the message to the passed connection index. Return TRUE if sent.
Private Function MessageNotAcknowledged(pConnectionIndex As Long) As Boolean
    With RemoteNetRegData(pConnectionIndex).AcknowledgeMessage
    
    'Only if it is a valid connection.
    If sckTCP(pConnectionIndex).State = sckConnected Then
    
        'Only resend if this is less than the third retry count.
        If .RemoteCommand > 0 And .RetryCount < 4 Then
        
            'Resend the win message.
            Call XmitBytes(pConnectionIndex, .RemoteCommand, CByte(.WinType), .BytBuf)
            
            'Increment the retry count.
            .RetryCount = .RetryCount + 1
            
            'Reset the retry timer.
            MessageNotAcknowledged = True
        
        'More that 3 attempts, no more resends to this connection.
        Else
            
            .RetryCount = 0
            .RemoteCommand = 0
            .WinType = 0
            MessageNotAcknowledged = False
            
        End If
    End If
    
    End With
End Function


'Sent from the remote terminal by command 25 to acknowledge that this terminal
'has won the war. The timer gets reset in function tmrAcknowledge_Timer() if
'it is not required. No need to reset it here.
Private Sub MessageAcknowledged(pConnectionIndex As Integer)
    
    'Clear the tag.
    With RemoteNetRegData(pConnectionIndex).AcknowledgeMessage
    .RemoteCommand = 0
    .WinType = 0
    .RetryCount = 0
    End With
    
End Sub
'----------------------------------------------------------------------------------------

'I have won the game, send which checkWin to run.
'TODO: Comment and document.
Public Sub IHaveWon(whichWin As Byte)
    Dim BytBuf() As Byte
    
    'Pack a refresh message.
    Call TheMainForm.PackForNetRefresh(BytBuf)
    
    Call SendMessage(10, BytBuf, whichWin, CByte(myTerminalNumber), True)
    
    TheMainForm.TimerWatch.Enabled = False
    TheMainForm.TimerWatch.Interval = 5000
End Sub

'Return the speed of the terminal: 0,1,2 = fastest..slowest
'TODO: Refactor, comment and document.
Public Function terminalSpeed() As Byte
    terminalSpeed = connectionSpeed
End Function

'Compose update, send depending on priority/slowest speed
'Priority 0,1,2 = lowest... highest
'TODO: Refactor, comment and document.
Public Sub handleUpdate(priority As Byte)
    Dim slowestSpd As Byte
    Dim dvCode As Long
    
    If (priority) >= connectionSpeed Then
        Call sendUpdate
    End If
End Sub

    'Compose and send update information to all terminals if host
    'To host if terminal
Public Sub sendUpdate()
    Dim byteMess() As Byte
    
    If Not TheMainForm.ComposeNetUpdate(byteMess) Then Exit Sub
    If UBound(byteMess) < 3 Then Exit Sub
    DoEvents
    If netWorkSituation = cNetHost Then
        Call XmitBytesAll(0, 8, 0, byteMess)
        Call UpdateForfeitTimer(CByte(gPlayerTurn - 1))
    ElseIf netWorkSituation = cNetClient Then
        Call XmitBytes(0, 8, myTerminalNumber, byteMess)
    End If
End Sub

     'Terminal requested refresh, send to terminal only
Public Sub SendRefreshRequested(terminal As Integer)
    Dim byteMes() As Byte
    
    Call TheMainForm.PackForNetRefresh(byteMes)
    Call XmitBytes(CByte(terminal), 6, 0, byteMes)
    TheMainForm.TimerWatch.Enabled = True
End Sub

'Send current situation to remote terminals. This is sent at the end of this terminal's turn.
Public Sub SendRefresh()
    Dim vBytBuf() As Byte
    
    'Pack the current situation.
    Call TheMainForm.PackForNetRefresh(vBytBuf)
    
    'Send to remote terminals.
    Call SendMessage(6, vBytBuf, 0, CByte(myTerminalNumber), True)
    
    'Update the forfeit timer if I am the session host.
    If netWorkSituation = cNetHost Then
        Call UpdateForfeitTimer(CByte(gPlayerTurn - 1))
    End If
    
    'Set the watchdog timer.
    TheMainForm.TimerWatch.Enabled = True
    
End Sub

    'Send war settings to terminal as requested
Private Sub sendWar(terminal As Integer)
    Dim byteMes() As Byte
    
    Call TheMainForm.packWarSettings(byteMes)
    Call XmitBytes(CByte(terminal), 14, 0, byteMes)
    TheMainForm.TimerWatch.Enabled = True
End Sub

    'I (client) request game so that I can join war
Public Sub requestWar()
    Dim byteMes() As Byte
    
    ReDim byteMes(2) As Byte
    Call XmitBytes(0, 13, 0, byteMes)
End Sub

    'I (client) requests refresh
Public Sub requestRefresh()
    Dim byteMes() As Byte
    ReDim byteMes(2) As Byte
    Call XmitBytes(0, 7, 0, byteMes)
End Sub

    'I host request refresh from terminal
Public Sub requestRefreshHost(whichTerminal As Long)
    Dim byteMes() As Byte
    ReDim byteMes(2) As Byte
    Call XmitBytes(whichTerminal, 7, 0, byteMes)
End Sub

'Client selected this player. Parameter pPlayerType should be
'set to 0 for human and anything else for computer. This is
'received by DibsOnPlayer() via command 5.
Public Sub ClaimPlayer(PlayerNo As Byte, pPlayerType As Long)
    Dim byteMsg() As Byte
    
    If Len(Trim(txtTerminalName.Text)) = 0 Then
        txtTerminalName.Text = "localhost"
    End If
    
    ReDim byteMsg(Len(txtTerminalName.Text) + 4) As Byte
    byteMsg(2) = myTerminalNumber
    
    'Human player = 0, computer player or remote is anything else.
    byteMsg(3) = CByte(pPlayerType)
    
    Call CopyBytes(byteMsg, StrConv(Trim(txtTerminalName.Text), vbFromUnicode), 4)
    Call XmitBytes(0, 5, PlayerNo, byteMsg)
End Sub

'Tell terminal to release this player.
Public Sub releasePlayer(BytMes() As Byte, whichPlayer As Byte, actualOwner As Byte)
    Dim NewBytMes() As Byte
    ReDim NewBytMes(Len(net.ClientName(actualOwner)) + 3) As Byte
    NewBytMes(2) = BytMes(2)
    NewBytMes(3) = actualOwner
    Call CopyBytes(NewBytMes, StrConv(net.ClientName(actualOwner), vbFromUnicode), 4)
    Call XmitBytes(CLng(NewBytMes(2)), 3, whichPlayer, BytMes)
End Sub

'Client has disclaimed the passed player.
Public Sub DisClaimPlayer(PlayerNo As Byte)
    Dim byteMsg() As Byte
    
    ReDim byteMsg(3) As Byte
    byteMsg(2) = 0
    byteMsg(3) = CByte(remoteIndex)
    Call XmitBytes(0, 5, PlayerNo, byteMsg)
End Sub

'Notify other players that this player is now a remote player.
Public Sub NotifyDisclamedPlayer(PlayerNo As Byte)
    Dim byteMsg() As Byte
    
    ReDim byteMsg(3) As Byte
    byteMsg(2) = 0
    byteMsg(3) = CByte(remoteIndex)
    Call XmitBytesAll(0, 5, PlayerNo, byteMsg)
End Sub

'I am client, request list of connected players from host.
Public Sub RequestConnectedList()
    Dim byteMsg() As Byte
    
    ReDim byteMsg(2) As Byte
    byteMsg(2) = 0
    Call XmitBytes(0, 18, CByte(myTerminalNumber), byteMsg)
End Sub

    'Host: send the setup settings and file name to all players
Public Sub sendSetupScreen()
    Static timeLastSent As Double
    Dim tmp As Double
    Dim vByteBuf() As Byte
    
    tmp = Timer
    'If tmp < (timeLastSent + xmitDelay) Then Exit Sub
    timeLastSent = tmp
    
    Call TheMainForm.PackSetupScreen(vByteBuf)
    Call XmitBytesAll(0, 2, 0, vByteBuf)
End Sub

'Host: send the setup settings and file name to 1 player.
Public Sub SendSetupScreenTo(pRemoteTerminal As Byte)
    Dim setUpByte() As Byte
    Call TheMainForm.PackSetupScreen(setUpByte)
    Call XmitBytes(CLng(pRemoteTerminal), 2, 0, setUpByte)
End Sub

    'Send starting situation of war
Public Sub sendOwnerScoreOrder(ByteMessage() As Byte)
    DoEvents
    Call XmitBytesAll(0, 4, 0, ByteMessage)
End Sub

'TODO: Comment and document.
Private Sub chkPasswordSession_Click()
    On Error Resume Next
    If chkPasswordSession.Value And chkPasswordSession.Enabled Then
        txtPassword.Enabled = True
        txtPassword.BackColor = &H80000005
    Else
        txtPassword.Enabled = False
        txtPassword.BackColor = &H8000000F
    End If
End Sub

Private Sub chkPasswordSession_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    tmrControlChange.Enabled = True
End Sub

Private Sub chkHideSession_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    tmrControlChange.Enabled = True
End Sub

Private Sub chkWlcmMsg_Click()
    If chkWlcmMsg.Value And chkWlcmMsg.Enabled Then
        txtWelcomeMsg.Enabled = True
        txtWelcomeMsg.BackColor = &H80000005
    Else
        txtWelcomeMsg.Enabled = False
        txtWelcomeMsg.BackColor = &H8000000F
    End If
End Sub

'Display the Session Locator.
Public Sub DisplaySessionLocator()
    On Error Resume Next
    
    'Network war.
    If sckTCP(0).State <> sckClosed Then
        Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, False)
        Exit Sub
    ElseIf optJoin.Value Then
        'Client.
        'Update the user.
        netMain.WriteText "Finding hosts...", False
        Call netFindHosts.BroadcastFindHost
        If optInet.Value Then
            'Find hosts on internet.
            If Not IxServerListSession(False) Then
                Exit Sub
            End If
        End If
    End If
End Sub

'Validate and return the TCP port chosen by the user from the
'Network Admin Panel. Return the default if not valid.
Private Function GetChosenTcpPort() As Long
    
    'Check if numeric.
    If IsNumeric(txtTcpPort.Text) Then
        
        'If chosen port number is within the port range.
        If CLng(txtTcpPort.Text) >= 0 And CLng(txtTcpPort.Text) < 65536 Then
            
            'Return the chosen port number.
            GetChosenTcpPort = CLng(txtTcpPort.Text)
        Else
            
            'If not, return the default port number.
            GetChosenTcpPort = gcDefaultPortNumber
            txtTcpPort.Text = CStr(gcDefaultPortNumber)
            
        End If
    End If
End Function

'Validate and return the UDP port chosen by the user from the
'Network Admin Panel. Return the default if not valid.
Private Function GetChosenUdpPort() As Long
    
    'Check if numeric.
    If IsNumeric(txtUdpPort.Text) Then
        
        'If chosen port number is within the port range.
        If CLng(txtUdpPort.Text) >= 0 And CLng(txtUdpPort.Text) < 65536 Then
            
            'Return the chosen port number.
            GetChosenUdpPort = CLng(txtUdpPort.Text)
        Else
            
            'If not, return the default port number.
            GetChosenUdpPort = gcDefaultPortNumber
            txtUdpPort.Text = CStr(gcDefaultPortNumber)
            
        End If
    End If
End Function

Private Sub cmdRestore_Click()
    InetSes.LocalUdpPort = gcDefaultPortNumber
    txtTcpPort.Text = CStr(gcDefaultPortNumber)
    txtUdpPort.Text = CStr(gcDefaultPortNumber)
    'txtTerminalName.Text = sckTCP(0).LocalHostName
    If Len(Trim(txtTerminalName.Text)) = 0 Then
        txtTerminalName.Text = "localhost"
    End If
End Sub

Private Sub Form_Activate()
    Call ListInfo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Pause button.
    If KeyCode = 19 Then
        Call TheMainForm.ActivatePauseMode
    End If
End Sub

Private Sub optInet_Click()
    On Error Resume Next
    Call EnableInternetOptions(True)
End Sub

Private Sub optLan_Click()
    On Error Resume Next
    Call EnableInternetOptions(False)
End Sub

' Enable or disable Internet options as required.
Public Sub EnableInternetOptions(pEnabled As Boolean)
    Dim vBackColour As Long
    
    cmdChangeLogin.Enabled = pEnabled
    
    txtTerminalName.Enabled = optLan.Value And optLan.Enabled
End Sub

Private Sub optRefresh_Click(Index As Integer)
    connectionSpeed = Index
End Sub

'Count the number of active terminals.
Public Function CountTerminals() As Long
    Dim i As Long
    
    On Error Resume Next
    CountTerminals = 0
    For i = 1 To MaxConnections
        If sckTCP(i).State <> sckClosed Then
            CountTerminals = CountTerminals + 1
        End If
    Next
End Function

'Return the next available terminal number.
'Return 0 if none found.
Private Function GetNextTerminal() As Long
    Dim i As Long
    
    GetNextTerminal = 0
    For i = 1 To MaxConnections
        If sckTCP(i).State = sckClosed Then
            GetNextTerminal = i
            Exit Function
        End If
    Next
End Function

'TODO: Refactor, comment and document.
Private Sub sckTCP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim setUpByte()     As Byte
    Dim NextTerminal    As Long
    Dim vMessage As String
    Dim vPort As Long
    
    On Error GoTo ErrHand
    
    If Index = 0 Then
        NextTerminal = GetNextTerminal
        
        'Reject if too many connection requests.
        If NextTerminal = 0 Then
            WriteText "Connection attempt refused. Too many clients.", True
            Exit Sub
        End If
         
        'Don't talk to banned players.
        If IsBanned(ChooseValidIP(sckTCP(Index).RemoteHostIP, sckTCP(Index).RemoteHost)) Then
            Exit Sub
        End If
        
        Packets(NextTerminal).bPacket = ""
        Packets(NextTerminal).CutOff = 0
        lastSendTime(NextTerminal) = 0
        RemoteNetRegData(NextTerminal).RegCode = ""
        RemoteNetRegData(NextTerminal).HostIP = ""
        RemoteNetRegData(NextTerminal).HostName = ""
        RemoteNetRegData(NextTerminal).ValidPassword = False
        RemoteNetRegData(NextTerminal).PasswordTrys = 0
        RemoteNetRegData(NextTerminal).AppVersion = ""
        RemoteNetRegData(NextTerminal).HostID = ""
        RemoteNetRegData(NextTerminal).VotesAgainst = ""
        sndComplete(NextTerminal) = 0
        sckTCP(NextTerminal).Tag = ""
        

        vPort = 0

        
        sckTCP(NextTerminal).LocalPort = vPort
        sckTCP(NextTerminal).Accept requestID
        DoEvents
        
        'Keep track of all clients
        net.ClientName(NextTerminal) = Phrase(233) + str(NextTerminal)   'Terminal nn
        vMessage = net.ClientName(NextTerminal) 'Replace(net.ClientName(NextTerminal), ",", ".")
        
        'Is password required?
        If chkPasswordSession.Value = vbChecked Then
            vMessage = vMessage & "," & "Validate"
            sckTCP(NextTerminal).Tag = "Validate"
        End If
        Call XmitString(NextTerminal, 0, CByte(NextTerminal), vMessage & vbCrLf)
        DoEvents
        If chkWlcmMsg.Value = vbChecked And Trim(txtWelcomeMsg.Text) <> "" Then
            Call XmitString(NextTerminal, 1, CByte(NextTerminal), _
                              txtWelcomeMsg.Text & vbCrLf)
        End If
    End If
    Exit Sub
ErrHand:
    If Err.Number = 10048 Then
        
        'Port in use, try to use another port.
        sckTCP(NextTerminal).Close
        DoEvents
        sckTCP(NextTerminal).LocalPort = 0
        LogError "sckTCP_ConnectionRequest", "Local port in use, trying another port number."
        Resume
        
    Else
        WriteText "Could not open client port. " & Err.Description, True
        LogError "sckTCP_ConnectionRequest", "Error: " & Err.Number & " " & Err.Description
        Exit Sub
    End If
    
End Sub

'Return TRUE if it is currently terminal(Index)'s turn.
Private Function IsTerminalsTurn(Index As Integer) As Boolean
    Dim vWinTest As Boolean
    
    On Error Resume Next
    
    vWinTest = (netWorkSituation = cNetHost) And Not (gCurrentMode = 13 Or gCurrentMode = 18)
    IsTerminalsTurn = vWinTest And (Index = net.playerOwner(gPlayerTurn - 1))
End Function

'Process data in bytBuf, retransmit to other clients if required.
'Needs refactoring badly.
'TODO: Refactor, comment and document.
Public Sub ProcessData(Index As Integer, ID As Long, BytBuf() As Byte, BytFramed() As Byte)
    Dim strText As String
    Dim BytHold() As Byte
    Dim x As Long
    Dim y As Byte
    Dim commandByte As Byte
    Dim PlayerByte As Byte
    Dim vTemp As Long
    Dim vTempStr() As String
    Dim vCounter As Long
    
    On Error GoTo ErrHand
    
    'commandByte = bytBuf(0)
    PlayerByte = BytBuf(1)
    BytHold = BytBuf
    
    'Ignore if disconnection pending.
    If InStr(1, sckTCP(Index).Tag, "silent") > 0 Then
        commandByte = 255
    ElseIf InStr(1, sckTCP(Index).Tag, "quiet") > 0 _
    And BytBuf(0) <> 9 Then
        commandByte = 255
    ElseIf InStr(1, sckTCP(Index).Tag, "Validate") > 0 _
    And BytBuf(0) <> 9 Then
        commandByte = 255
    Else
        commandByte = BytBuf(0)
    End If
    
    'Log packet.
    LogInfo "ProcessData", "Term=" & Format(Index, "00") _
                        & " PlrByt=" & Format(PlayerByte, "00") _
                        & " ID=" & Format(ID, "00000") _
                        & " Cmd=" & Format(commandByte, "00") _
                        & " Pkg=" & ToHexStr(BytHold), 5, True
    Debug.Print "ProcessData command"; commandByte; " From"; Index; "PlayerByte"; PlayerByte; "Length"; UBound(BytBuf)
    
    Select Case commandByte
    
    Case 0
        'Assign name recieved from host
        gPlayerTurn = 1
        Packets(Index).CutOff = ID
        TheMainForm.cmdEnd.Tag = "Recently Connected"     'Prevent audit after connecting.
        Call getMyTermNo(BytBuf, PlayerByte)
        Call requestSetupScreen
        
    Case 1
        'Public message arrived.
        If netWorkSituation = cNetHost Then
            Call XmitPacketAll(CLng(Index), commandByte, PlayerByte, BytFramed)
            
        End If
        Call netChatterBox.printMessageByte(BytBuf, PlayerByte)
        
    Case 2
        'Setup settings. Only host can send these.
        If Index = 0 Then
            Packets(Index).CutOff = ID
            Call TheMainForm.UnpackSetupScreen(BytBuf)
        End If
    Case 3
        'Release player. Only host can send these.
        If Index = 0 Then
            Call TheMainForm.releaseMyPlayer(BytBuf, PlayerByte)
        End If
        
    Case 4
        'New war information.
        Packets(Index).CutOff = ID
        Call TheMainForm.startNewWar(BytBuf)
        
        'Stats are valid if starting from the begining of the war.
        Call TheMainForm.ValidateAllStats(True)
        
    Case 5
        'Client has claimed a player. If I am host, forward to every connected terminal.
        If TheMainForm.DibsOnPlayer(BytHold, PlayerByte, Index) Then
            Call XmitPacketAll(CLng(Index), commandByte, PlayerByte, BytFramed)
        End If
        
    Case 6
        'Refresh map. Sent at the end of each turn.
        'Accept if sent by the host or it is the terminal's turn.
        If Index = 0 Or IsTerminalsTurn(Index) Then
            Packets(Index).CutOff = ID
            'Call XmitPacketAll(CLng(Index), commandByte, PlayerByte, BytFramed)
            
            'Resend to all except the sender. Set up the acknowledgement timer.
            Call SendMessage(6, BytHold, PlayerByte, Index, True)
            
            'Set to the current situation.
            Call TheMainForm.UnpackNetRefresh(BytHold, Index)
            
            'TODO: Figure out and rationalise.
            Packets(Index).TerminalTurn = net.playerOwner(gPlayerTurn - 1)
            
            'Keep turning the cheat codes off just in case. This helps to make hacking a bit tougher.
            Call TheMainForm.TurnOffCheatCodes
            
            'Reset tmrForfeitTurn. Set tag to player's terminal no and 0 minute timer.
            Call UpdateForfeitTimer(CByte(gPlayerTurn - 1))
            
            'Acknowledge the refresh message has been recieved.
            Call XmitString(CLng(Index), 25, PlayerByte, "ack")
            
        End If
        
        TheMainForm.TimerWatch = False
        
    Case 7
        'Request refresh.
        Call SendRefreshRequested(Index)
        
    Case 8
        'Update information.
        'If I am host, check that the data is in order and
        'it was sent by the terminal who's turn it is.
        'If I am client, trust data sent by host.
        'Only accept updates that have a message  ID greater than the
        'cutoff ID for thie connected terminal. This is to prevent
        'updates that were sent before a refresh, win etc that arrive
        'after the refresh or win etc has been applied.
        If Index = 0 Or _
        IsTerminalsTurn(Index) _
        And (ID > Packets(Index).CutOff Or ID + 2000 < Packets(Index).CutOff) Then
            Packets(Index).CutOff = ID
            Call XmitPacketAll(CLng(Index), commandByte, PlayerByte, BytFramed)
            Call TheMainForm.UnpackNetUpdate(BytHold)
            Call UpdateForfeitTimer(CByte(gPlayerTurn - 1))
            TheMainForm.pctInfoBox.Refresh
            TheMainForm.TimerWatch.Enabled = False
            TheMainForm.TimerWatch.Interval = 5000
        End If
        
        'Debug.Print "Update: "; net.playerOwner(gPlayerTurn - 1)
    Case 9
        'Remote requested setup info. Check reg code.
        Packets(Index).CutOff = ID
        If IsBanned(ChooseValidIP(sckTCP(Index).RemoteHostIP, sckTCP(Index).RemoteHost), _
        net.ClientName(Index), RemoteNetRegData(Index).HostID) Then
        
            'This IP has been banned.
            sckTCP(Index).Tag = "silent"
            Call XmitString(CLng(Index), 20, 0, "Connection refused.")
            WriteText net.ClientName(Index) & " is in your banned list. Connection terminated.", False
            
        ElseIf CountConnectionsFromIP(Index) > CLng(txtMaxConnections.Text) Then
        
            'Connection limit for this IP reached.
            sckTCP(Index).Tag = "silent"
            Call XmitString(CLng(Index), 20, 0, "Connection refused. Connection limit from this IP exceeded.")
            WriteText net.ClientName(Index) & " attempted to connect. Connection limit exceded for IP.", False
        
        'Function ListReg() is where the name is passed from the client.
        ElseIf Not listReg(Index, BytBuf) Then
        
            'Clients on different machines are using the same reg code.
            sckTCP(Index).Tag = "silent"
            Call XmitBytes(CLng(Index), 16, 0, BytBuf)
        
        ElseIf Not CheckPassword(Index, BytBuf) Then
            
            'Client must send a password to connect to this machine. 3 tries only.
            sckTCP(Index).Tag = "quiet"
            RemoteNetRegData(Index).PasswordTrys = RemoteNetRegData(Index).PasswordTrys + 1
            If RemoteNetRegData(Index).PasswordTrys >= 3 Then
                Call XmitString(CLng(Index), 20, 0, "Incorrect password.")
                WriteText net.ClientName(Index) & " attempted to connect. Invalid password.", False
            Else
                Call XmitBytes(CLng(Index), 22, 0, BytBuf)
            End If
            
        Else
            
            'All OK.
            sckTCP(Index).Tag = ""
            Call TheMainForm.PackSetupScreen(BytHold)
            Call XmitBytes(CLng(Index), 2, 0, BytHold)
            WriteText net.ClientName(Index) & Phrase(231), True '<Term Name> has connected.
            Call XmitStringAll(CLng(Index), 1, 0, net.ClientName(Index) & Phrase(231) & vbCrLf)
            
            'Send rules. **Not actually sent**
            If vscrollTimeLimit.Value <> 0 Then
                strText = "Forfeit turn after " & CStr(vscrollTimeLimit.Value) & " seconds" & vbCrLf
            Else
                strText = ""
            End If
            
            'Send version info and encryption key.
            Call XmitString(CLng(Index), 23, 0, GetVersionInfo _
                        & "," & GetUniqueId _
                        & "," & CStr(MyNetRegData.LeKey) _
                        & "," & CStr(MyNetRegData.LeSlot) _
                        & "," & CStr(MyNetRegData.LeSlotSpin))
            PlaySoundFromFile ConnectSoundFile
            
        End If
        
    Case 10
        'We have a winner.
        If Index = 0 Or IsTerminalsTurn(Index) Then
            Packets(Index).CutOff = ID
            y = PlayerByte
            'Call XmitPacketAll(CLng(Index), commandByte, PlayerByte, BytFramed)
                
            'Resend to all except the sender
            Call SendMessage(10, BytHold, y, Index, True)
            
            'Set the current situation to that sent.
            Call TheMainForm.UnpackNetRefresh(BytHold, Index)
        
            'Show win and stats.
            Call TheMainForm.gameWon(y)
            
            'Acknowledge the win message has been recieved.
            Call XmitString(CLng(Index), 25, PlayerByte, "ack")
            
            TheMainForm.TimerWatch.Enabled = False
            TheMainForm.TimerWatch.Interval = 5000
        End If
            
    Case 11
        'Setup cancel.
        Call TheMainForm.UnpackNetRefresh(BytHold, Index)
        TheMainForm.SetupScreen.Visible = False
        Call EnableMissionOptions
        TheMainForm.TimerWatch.Enabled = False
        TheMainForm.TimerWatch.Interval = 5000

    Case 12
        'Cards are being exchanged.
        If Index = 0 Or IsTerminalsTurn(Index) Then
            Call XmitPacketAll(CLng(Index), commandByte, PlayerByte, BytFramed)
            Call TheMainForm.exchangeTheseCards(PlayerByte, BytHold)
            Call UpdateForfeitTimer(CByte(gPlayerTurn - 1))
            TheMainForm.TimerWatch.Enabled = False
            TheMainForm.TimerWatch.Interval = 5000
            'TheMainForm.pctDice.Refresh
        End If
        
    Case 13
        'Remote requested to join war.
        Call sendWar(Index)
        
    Case 14
        'Recieved requested war info as to join.
        Call TheMainForm.unpackWarSettings(BytHold)
        
    Case 15
        'Last terminal has kicked this dog - prevent stall.
        Call MyTerminalKicked(PlayerByte, CLng(Index), BytHold)
    
    Case 16
        ReDim bBuff(12) As Byte
        Call XmitString(0, 9, 0, "          ")
    
    Case 17
        'Private message.
        Call XmitPacketAll(CLng(Index), commandByte, PlayerByte, BytFramed)
        Call netChatterBox.privateMessage(BytBuf, PlayerByte)
    
    Case 18
        'I am host, client request me to send a list of connected terminals.
        'Clients will send this request every 5 seconds.
        Call XmitString(CByte(PlayerByte), 19, PlayerByte, ListConnectedTerminals)
        
    Case 19
        'I am client, a list of connected terminals has arrived from the host.
        Call DisplayConnectedTerminals(StrConv(MidB(BytBuf, 3), vbUnicode))
        
    Case 20
        'Host rejected connection because this IP has been banned or limit reached.
        'Call WriteText("Connection refused by host.")
        Call WriteText(StrConv(MidB(BytBuf, 3), vbUnicode))
        Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, True)
    
    Case 21
        'Forfeit turn. Host has told me to press pass.
        Call TheMainForm.ForfeitTurn
    
    Case 22
        'Host has requested a password to enter this game.
        TheMainForm.SetupScreen.Visible = True
        Call RequestPassword
    
    Case 23
        'Remote machine as sent their version number (host and client).
        On Error Resume Next
        vTempStr = Split(StrConv(MidB(BytBuf, 3), vbUnicode), ",")
        If UBound(vTempStr) > 0 Then
            RemoteNetRegData(Index).HostID = vTempStr(1)
        End If
        If UBound(vTempStr) > 3 Then
            'RemoteNetRegData(Index).LeKey = CLng(vTempStr(2))
            'RemoteNetRegData(Index).LeSlot = CLng(vTempStr(3))
            'RemoteNetRegData(Index).LeSlotSpin = CLng(vTempStr(4))
        End If
        
        If IsBanned(ChooseValidIP(sckTCP(Index).RemoteHostIP, sckTCP(Index).RemoteHost), _
                    net.ClientName(Index), RemoteNetRegData(Index).HostID) Then
        
            'This IP has been banned.
            sckTCP(Index).Tag = "silent"
            Call XmitString(CLng(Index), 20, 0, "Connection refused.")
            WriteText net.ClientName(Index) & " is in your banned list. Connection terminated.", False
        
        ElseIf netWorkSituation = cNetClient Then
            'Send version number if I am a client.
            Call XmitString(CLng(Index), 23, 0, GetVersionInfo & "," _
                            & GetUniqueId _
                            & "," & CStr(MyNetRegData.LeKey) _
                            & "," & CStr(MyNetRegData.LeSlot) _
                            & "," & CStr(MyNetRegData.LeSlotSpin))
        End If
        'Set encryption type depending on version.
        RemoteNetRegData(Index).AppVersion = Format(vTempStr(0), "00000000")
        'If RemoteNetRegData(Index).AppVersion > "03020037" Then
        '    RemoteNetRegData(Index).LeType = 2
        'Else
        '    RemoteNetRegData(Index).LeType = 0
        'End If
        RemoteNetRegData(Index).LeType = 2
        On Error GoTo ErrHand
    
    Case 24
        'Remote has sent a vote.
        Call HandlePlayerVote(Index, BytBuf)
    
    Case 25
        'A win acknowledgement has been recieved.
        Call MessageAcknowledged(Index)
    
    Case 255
        'Ignore if disconnection pending.
    Case Else
        Debug.Print "Error in natMain.ProcessData - commandByte = " & CStr(commandByte)
        LogError "ProcessData", "Case Else: " & CStr(commandByte), True
    End Select
    Exit Sub
ErrHand:
    LogError "ProcessData", "cmd=" & CStr(commandByte) & " " & Err.Number & " " & Err.Description, True
    Exit Sub
End Sub

'Add vote from remote terminal to the player's list.
'Xmit format: Command 24 "<Target terminal>,<kill|forfeit>"
'Save format: <date time stamp>,<kill|forfeit>,<source terminal> <CRLF>
Private Sub HandlePlayerVote(pTermFromIndex As Integer, BytBuf() As Byte)
    Dim vParts() As String
    
    vParts = Split(Mid(StrConv(BytBuf, vbUnicode), 3), ",")
    
    'Check for correct format.
    If UBound(vParts) <> 1 Then
        Exit Sub
    End If
    
    'This check must be done second.
    If Not IsNumeric(vParts(0)) Then
        Exit Sub
    End If
    
    If CountPlayersTermOwns(pTermFromIndex) > 0 Then
        'Add vote to the tally.
        RemoteNetRegData(vParts(0)).VotesAgainst _
            = RemoteNetRegData(vParts(0)).VotesAgainst _
            & CStr(CDbl(Now)) & "," _
            & CStr(vParts(1)) & "," _
            & CStr(pTermFromIndex) _
            & vbCrLf
        
        'Notify other players that this vote has arrived.
        If vParts(1) = "FORFEIT" Then
            Call XmitStringAll(0, 1, 0, "Forfeit Turn vote received." & vbCrLf)
            Call WriteText("Forfeit Turn vote received.")
        ElseIf vParts(1) = "KILL" Then
            Call XmitStringAll(0, 1, 0, "Kick " & net.ClientName(CLng(vParts(0))) & " vote received." & vbCrLf)
            Call WriteText("Kick " & net.ClientName(CLng(vParts(0))) & " vote received.")
        End If
        
        'Tally votes and take action if required.
        Call TallyPlayerVotes(CInt(vParts(0)))
        
    End If
End Sub

'Tally player votes against passed terminal and take action.
Private Sub TallyPlayerVotes(pTargetTerm As Integer)
    Dim vParts() As String
    Dim vLine() As String
    Dim vLineIndex As Long
    Dim vDateTimeLimit As Double
    Dim vVotedKill As Long
    Dim vVotedForfeit As Long
    Dim vTotalTerminals As Long
    Dim vTermVotedKill(MaxConnections) As Boolean
    Dim vTermVotedForfeit(MaxConnections) As Boolean
    Dim BytBuf() As Byte
    
    On Error Resume Next
    
    'Set expiry to 5 minutes.
    vDateTimeLimit = CDbl(DateAdd("n", -5, Now))
    
    
    If RemoteNetRegData(pTargetTerm).VotesAgainst <> "" Then
        vLine = Split(RemoteNetRegData(pTargetTerm).VotesAgainst, vbCrLf)
        
        'Count the votes.
        'For each line in the votelist.
        For vLineIndex = 0 To UBound(vLine) - 1
            vParts = Split(vLine(vLineIndex), ",")
            If UBound(vParts) = 2 Then
                
                'If less than 5 minutes old.
                If CDbl(vParts(0)) >= vDateTimeLimit Then
                    If UCase(vParts(1)) = "KILL" _
                    And Not vTermVotedKill(CLng(vParts(2))) Then
                        'Kill vote.
                        vVotedKill = vVotedKill + 1
                        vTermVotedKill(CLng(vParts(2))) = True
                    ElseIf UCase(vParts(1)) = "FORFEIT" _
                    And Not vTermVotedForfeit(CLng(vParts(2))) Then
                        'Forfeit vote.
                        vVotedForfeit = vVotedForfeit + 1
                        vTermVotedForfeit(CLng(vParts(2))) = True
                    End If
                End If
            End If
        Next
        
        'Tally the votes.
        vTotalTerminals = CountTerminalsWithPlayers
        'If net.playerOwner(gPlayerTurn - 1) = myTerminalNumber Then
        '    'I am the host and this vote was against my player. Add me as an
        '    'eligible voter so that one terminal can't boss me around.
        '    vTotalTerminals = vTotalTerminals + 1
        'End If
        
        'Count kill votes.
        If vVotedKill > vTotalTerminals * 0.5 Then
            cmdKill.Tag = sckTCP(pTargetTerm).RemoteHostIP & "," _
                        & EncodeNonAscii(net.ClientName(pTargetTerm)) & "," _
                        & CStr(pTargetTerm)
            
            'Notify the Index Server server of banning.
            If optInet.Value Then
                Call IxServerPlayerWasBanned(sckTCP(pTargetTerm).RemoteHostIP _
                                            & "," & gGsLeUtils.LE6(net.ClientName(pTargetTerm)) _
                                            , "Kicked by majority vote")
            End If
            
            Call KillConnection
            
        'Count forfeit votes.
        ElseIf vVotedForfeit > vTotalTerminals * 0.5 _
        And net.playerOwner(gPlayerTurn - 1) = pTargetTerm Then
            If net.playerOwner(gPlayerTurn - 1) = myTerminalNumber Then
                'Oops, I, the host own this player.
                Call TheMainForm.ForfeitTurn
            Else
                Call XmitString(CLng(net.playerOwner(gPlayerTurn - 1)), 1, 0, _
                            "\beep Your turn has been forfeited by majority vote." & vbCrLf)
                Call XmitStringAll(CLng(net.playerOwner(gPlayerTurn - 1)), 1, 0, "Turn forfeited by majority vote." & vbCrLf)
                ReDim BytBuf(2) As Byte
                Call XmitBytes(CLng(net.playerOwner(gPlayerTurn - 1)), 21, 0, BytBuf)
            End If
            
            'Remove forfeit votes from the tag so that 1 person cannot force another forfeit
            'on this player's next turn.
            Call ClearForfeitVotes
        End If
    End If
End Sub

'Clear all forfeit votes at the end of the turn.
Public Sub ClearForfeitVotes()
    Dim vTerminal As Long
    
    For vTerminal = 0 To MaxConnections
        RemoteNetRegData(vTerminal).VotesAgainst _
                = Replace(RemoteNetRegData(vTerminal).VotesAgainst, "FORFEIT", "XFORFEIT")
    Next
End Sub

'Count the number of active terminals that have claimed players.
Public Function CountTerminalsWithPlayers() As Long
    Dim vSocketIndex As Long
    Dim vPlayerIndex As Long
    
    On Error Resume Next
    
    CountTerminalsWithPlayers = 0
    
    For vSocketIndex = 1 To MaxConnections
        If sckTCP(vSocketIndex).State <> sckClosed Then
            For vPlayerIndex = 0 To 5
                If net.playerOwner(vPlayerIndex) = vSocketIndex Then
                    CountTerminalsWithPlayers = CountTerminalsWithPlayers + 1
                    'Only count once no matter how many players this terminal has.
                    Exit For
                End If
            Next
        End If
    Next
End Function

'Ask host to tell me my name and terminal number and setup settings.
'Send reg info: "RegCode, LocalIP, LocalHostName, Password".
'TODO: Refactor, comment and document.
Private Sub RequestPassword()
    Dim bBuff() As Byte
    Dim vPassword As String
    
    If Len(txtTerminalName.Text) = 0 Then
        txtTerminalName.Text = "localhost"
    End If
    
    vPassword = InputBox("This session is password protected. Please enter the password.", _
                "This session is password protected")
    
    'Is this needed?
    If Len(Trim(vPassword)) = 0 Then
        Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, True)
    End If

    Call XmitString(0, 9, 0, EncodeNonAscii(txtTerminalName.Text) & "," _
                      & Replace(vPassword, ",", "."))
End Sub

'Update the forfeit turn timer. This is called initial by sendRefresh() which resets
'the counter for a new player turn and by ProcessData() to intercept valid updates
'made by the current player.
Private Sub UpdateForfeitTimer(pPlayer As Byte)
    Dim vForfeitInterval As String

    'Check tmrForfeitTurn. Set tag to player's terminal number and reset timer.
    If netWorkSituation = cNetHost _
    And vscrollTimeLimit.Value <> 0 _
    And Not TheMainForm.SetupScreen.Visible _
    And pPlayer = gPlayerTurn - 1 Then
        tmrForfeitTurn.Enabled = False
        tmrForfeitTurn.Interval = 1000
        vForfeitInterval = CStr(vscrollTimeLimit.Value)
        tmrForfeitTurn.Tag = CStr(net.playerOwner(gPlayerTurn - 1)) & "," & vForfeitInterval '180
        tmrForfeitTurn.Enabled = True
        'Debug.Print "Forfeit timer reset."
    Else
        tmrForfeitTurn.Enabled = False
        tmrForfeitTurn.Tag = "0,0"
    End If
End Sub

'Update and check the time left for the current player before forfeit.
'Give a 10 second warning beforehand. The timer control is set to a 1 second
'interval and the actual timer is kep in the timer's tag. The format of the
'timer is "<Player turn>,<seconds left>"
Private Sub tmrForfeitTurn_Timer()
    Dim vTimeLeft As Long
    Dim vParts() As String
    Dim BytBuf() As Byte
    
    On Error GoTo ErrHand
    
    vParts = Split(tmrForfeitTurn.Tag, ",")
    If vscrollTimeLimit.Value <> 0 _
    And UBound(vParts) = 1 _
    And gCurrentMode <> 13 _
    And gCurrentMode <> 18 _
    And Not TheMainForm.SetupScreen.Visible Then
        
        If vParts(0) = net.playerOwner(gPlayerTurn - 1) Then
        
            vTimeLeft = CLng(vParts(1)) - 1
            
            If vTimeLeft < -5 Then
                
                'Host forces a disconnect by killing the connection as a last resort.
                Call sckTCP_Close(CInt(net.playerOwner(gPlayerTurn - 1)))
                
            ElseIf vTimeLeft = 0 Then
                
                'Host tells client to disconnect.
                If netWorkSituation = cNetHost And CLng(vParts(0)) = 0 Then
                    Call WriteText("Time is up. Your turn has been forfeited.")
                    Call TheMainForm.ForfeitTurn
                Else
                    Call XmitString(CLng(net.playerOwner(gPlayerTurn - 1)), 1, 0, _
                        "\beep Time is up. Your turn has been forfeited." & vbCrLf)
                    ReDim BytBuf(2) As Byte
                    Call XmitBytes(CLng(net.playerOwner(gPlayerTurn - 1)), 21, 0, BytBuf)
                End If
                
            ElseIf vTimeLeft = 10 Then
                
                'Host sends a 10 second warning.
                If netWorkSituation = cNetHost And CLng(vParts(0)) = 0 Then
                    Call WriteText("10 seconds left...")
                Else
                    Call XmitString(CLng(net.playerOwner(gPlayerTurn - 1)), 1, 0, _
                        "\beep 10 seconds left..." & vbCrLf)
                    'Debug.Print "10 seconds left."
                End If
            End If
            tmrForfeitTurn.Tag = vParts(0) & "," & CStr(vTimeLeft)
        End If
    Else
        tmrForfeitTurn.Enabled = False
    End If
    Exit Sub
ErrHand:
    Exit Sub
End Sub

'Count connections from the passed IP address.
Private Function CountConnectionsFromIP(pIndex As Integer) As Long
    Dim vIP As String
    Dim i As Long
    
    vIP = ChooseValidIP(sckTCP(pIndex).RemoteHostIP, sckTCP(pIndex).RemoteHost)
    
    CountConnectionsFromIP = 0
    For i = 1 To MaxConnections
        If sckTCP(i).State <> sckClosed _
        And ChooseValidIP(sckTCP(i).RemoteHostIP, sckTCP(i).RemoteHost) = sckTCP(pIndex).RemoteHostIP Then
            CountConnectionsFromIP = CountConnectionsFromIP + 1
        End If
    Next
End Function

'I am host, return a list of connected terminals. This string is used
'by function DisplayConnectedTerminals() to display connected terminals
'in the local listinfo box and is also sent to remote players as part of
'command 19 which is put into clients' listinfo boxes.
'Format: <IP>,<Termname>,<SocketNumber><cr>
Public Function ListConnectedTerminals() As String
    Dim PlayerList As String
    Dim i As Long
    
    PlayerList = ChooseValidIP(sckTCP(0).LocalIP, sckTCP(0).LocalHostName) _
                & "," & EncodeNonAscii(netMain.txtTerminalName.Text) _
                & "," & 0 & vbCrLf
    
    If netWorkSituation = cNetHost Then
        For i = 1 To MaxConnections
            If netMain.sckTCP(i).State <> sckClosed Then
                PlayerList = PlayerList & FormatConnectedDetails(i) & vbCrLf
            End If
        Next
    End If
    ListConnectedTerminals = PlayerList
End Function

'I am the host. Return a list of connected terminal owners' display names.
'If pEncrypt is TRUE, encrypt each line using gGsLeUtils.LE6. If not, encode all
'non ascii characters.
Public Function ListConnectedTermNames(Optional pEncrypt As Boolean = False) As String
    Dim PlayerList As String
    Dim vIndex As Long
    
    'Host only.
    If netWorkSituation = cNetHost Then
        
        'For all connected clients (0 is the host).
        For vIndex = 1 To MaxConnections
            
            'If the client socket is open.
            If netMain.sckTCP(vIndex).State <> sckClosed Then
                
                If pEncrypt Then
                    'Encrypt and add the client name to the list.
                    ListConnectedTermNames = ListConnectedTermNames _
                                    & gGsLeUtils.LE6(net.ClientName(vIndex)) & ","
                Else
                
                    'Add the client name to the list.
                    ListConnectedTermNames = ListConnectedTermNames _
                                    & EncodeNonAscii(net.ClientName(vIndex)) & ","
                End If
            End If
        Next
    End If
    
    'Remove blanks and end commas.
    ListConnectedTermNames = CleanList(ListConnectedTermNames)
End Function


'I host. Format details about connected terminals.
'Format: <IP>,<Termname>,<SocketNumber<cr>>
Private Function FormatConnectedDetails(pElementNo As Long) As String
    FormatConnectedDetails = _
            ChooseValidIP(netMain.sckTCP(pElementNo).RemoteHostIP, netMain.sckTCP(pElementNo).RemoteHost) _
            & "," & EncodeNonAscii(net.ClientName(pElementNo)) _
            & "," & CStr(pElementNo)
End Function

'Kill clients if they appear in passed list.
'*** Not used ***
Public Sub KillClients(pClientDetails As String, pBan As Boolean)
    Dim i As Long
    For i = 1 To MaxConnections
        If netMain.sckTCP(i).State <> sckClosed Then
            If InStr(1, pClientDetails, FormatConnectedDetails(i)) > 0 Then
                If pBan Then
                    'Ban this IP.
                    'net.BanList = net.BanList & netMain.sckTCP(i).RemoteHostIP & vbCrLf
                End If
                
                sckTCP_Close CInt(i)
                
            End If
        End If
    Next
End Sub

'Return list of connected players, IPs and controled armies.
'Format "Army<crlf>Terminal Name<crlf>IP<crlf>"
'Get info from host if required.
Public Function Whois() As String
    Dim bytBuff() As Byte
    Dim i As Long
    Dim j As Long
    
    'Am I host?
    If myTerminalNumber = 0 Then
        For i = 0 To 5
            Whois = Whois _
                    & CStr(i) & vbCrLf _
                    & net.ClientName(i) & vbCrLf _
                    & ChooseValidIP(sckTCP(i).RemoteHostIP, sckTCP(i).RemoteHost) & vbCrLf
        Next
    Else
    
    End If
End Function

'Return byte array containing whois info (host).
Private Sub AnswerWhois()
    
End Sub

'Check password if required. If not already entered for this client then
'check if it is in the passed byte array and return false if not.
'TODO: Refactor, comment and document.
Private Function CheckPassword(Index As Integer, BytBuf() As Byte) As Boolean
    Dim i As Long
    Dim rCode As String
    Dim rParts() As String
    Dim rPasswd As String
    
    On Error Resume Next
    
    'Is password required or already entered?
    If chkPasswordSession.Value = vbUnchecked _
    Or RemoteNetRegData(Index).ValidPassword Then
        CheckPassword = True
        Exit Function
    End If
    
    'Check if it was passed in byte array.
    rCode = StrConv(BytBuf(), vbUnicode)
    rCode = Mid(rCode, 3)
    
    If Len(Trim(rCode)) = 0 Then
        CheckPassword = False
        Exit Function
    End If
    
    rParts = Split(rCode, ",")
    
    If UBound(rParts) = 0 Then
        CheckPassword = False
        Exit Function
    End If
    
    'Find password. Quick and nasty.
    If UBound(rParts) = 1 Then
        rPasswd = rParts(1)
    ElseIf UBound(rParts) = 4 Then
        rPasswd = rParts(4)
    Else
        CheckPassword = False
        Exit Function
    End If
    
    'Validate the password.
    If rPasswd = Replace(txtPassword.Text, ",", ".") Then
        
        'Pass.
        RemoteNetRegData(Index).ValidPassword = True
        CheckPassword = True
        Exit Function
        
    Else
        
        'Fail.
        RemoteNetRegData(Index).ValidPassword = False
        CheckPassword = False
        Exit Function
        
    End If

End Function

'return false if player reg code is in list and on a different terminal (Host).
'"RegCode, LocalIP, LocalHostName, TerminalName, Password".
'This is where the client's terminal name is accepted, net.ClientName.
' **TODO 'TODO: Refactor, comment and document.
Private Function listReg(Index As Integer, BytBuf() As Byte) As Boolean
    Dim i As Long
    Dim rCode As String
    Dim rParts() As String
    Dim vValidIP As String
    Dim vCountClientName As String            'Used to modify the client name is it is already being used.
    
    On Error Resume Next
    listReg = True
    
    rCode = StrConv(BytBuf(), vbUnicode)
    rCode = Mid(rCode, 3)
    
    If Len(Trim(rCode)) = 0 Then
        Exit Function
    End If
    
    rParts = Split(rCode, ",")
    
    vValidIP = ChooseValidIP(rParts(1), sckTCP(Index).RemoteHostIP, sckTCP(Index).RemoteHost)
    vValidIP = ChooseValidIP(vValidIP, sckTCP(0).RemoteHostIP)
    
    If UBound(rParts) < 3 Then
        
        'No reg code. This should be filled with the host name. >> TEST <<
        net.ClientName(Index) = DecodeNonAscii(Trim(rParts(0)))
        Exit Function
    
    ElseIf UBound(rParts) >= 3 Then
        
        If Trim(rParts(0)) = "" Then
            'No reg code.
            RemoteNetRegData(Index).HostIP = vValidIP
            RemoteNetRegData(Index).HostName = DecodeNonAscii(rParts(2))
            net.ClientName(Index) = DecodeNonAscii(rParts(3))
            Exit Function
        End If
        
        'Process reg code and stuff.
        vCountClientName = ""
        For i = 1 To MaxConnections
            If CInt(i) <> Index Then
                If rParts(0) = RemoteNetRegData(i).RegCode _
                And IsValidIP(vValidIP) And (vValidIP <> RemoteNetRegData(i).HostIP _
                Or False) Then
                    listReg = False
                    Exit Function
                End If
                
                'Append a "-02" to the client name if that name is already being used.
                'Take appropriate actions to keep it unique.
                If DecodeNonAscii(rParts(3)) & vCountClientName = net.ClientName(i) And sckTCP(i).State = sckConnected Then
                    If vCountClientName = "" Then
                        vCountClientName = "-02"
                    Else
                        vCountClientName = Format(CLng(vCountClientName) - 1, "00")
                    End If
                End If
            End If
        Next
        
        Debug.Print vValidIP & " connected (sub listReg)"
        RemoteNetRegData(Index).RegCode = Trim(rParts(0))
        RemoteNetRegData(Index).HostIP = vValidIP
        RemoteNetRegData(Index).HostName = DecodeNonAscii(rParts(3)) & vCountClientName ' rParts(2)
        net.ClientName(Index) = DecodeNonAscii(rParts(3)) & vCountClientName
    End If
End Function

'TODO: Refactor, comment.
Private Sub getMyTermNo(BytBuf() As Byte, PlayerByte As Byte)
    Dim strText As String
    
    strText = StrConv(BytBuf(), vbUnicode)
    strText = Replace(Mid(strText, 3), vbCrLf, "")
    If InStr(1, strText, ",") > 1 Then
        strText = Mid(strText, 1, InStr(1, strText, ",") - 1)
    End If
    WriteText Phrase(236) & strText, True   '>> Your are
    MyTermNo = strText
    myTerminalNumber = PlayerByte
End Sub

'Validate user inputs before connecting.
Public Function ValidateBeforeConnection() As Boolean
    Dim vMsgBox As VbMsgBoxResult
    
    On Error Resume Next
    
    ValidateBeforeConnection = False
    
    'Check the TCP port is numeric.
    If Not IsNumeric(txtTcpPort.Text) Then
        
        vMsgBox = MsgBox("Port number must be between 1 and 65535.", vbOKOnly)
        txtTcpPort.SetFocus
    
    'Check the TCP port is between 0 and 65535.
    ElseIf CLng(txtTcpPort.Text) < 0 Or CLng(txtTcpPort.Text) > 65535 Then
        
        vMsgBox = MsgBox("Port number must be between 0 and 65535.", vbOKOnly)
        txtTcpPort.SetFocus
    
    'Check the UDP port is numeric.
    ElseIf Not IsNumeric(txtUdpPort.Text) Then
        
        vMsgBox = MsgBox("Port number must be between 1 and 65535.", vbOKOnly)
        txtUdpPort.SetFocus
    
    'Check the UDP port is between 0 and 65535.
    ElseIf CLng(txtUdpPort.Text) < 0 Or CLng(txtTcpPort.Text) > 65535 Then
        
        vMsgBox = MsgBox("Port number must be between 0 and 65535.", vbOKOnly)
        txtUdpPort.SetFocus
    
    Else
        
        'Validated.
         ValidateBeforeConnection = True
    End If
End Function

'Handle the click event for the FindWar or Begin or Disconnect button.
Public Sub cmdConnect_Click()
    Static vLastTime As Date
    
    On Error Resume Next
    
    'Wait 3 seconds before reactivating.
    If Abs(dateDiff("s", vLastTime, Now)) < 2 Then
        
        WriteText "Please wait one moment..."
    
    'Make sure the Internet Gateway is not busy.
    ElseIf netInetGateway.IsStillExecuting Then
        
        WriteText "The Internet gateway is busy. Please wait one moment..."
    
    'All good. Find hosts or begin.
    Else
    
        'Remember the last time here.
        vLastTime = Now
        
        'Find hosts or Begin or Disconnect depending on the situation.
        Call ConnectDisconnectBeginSession
        
    End If
End Sub

'Find Session or Begin Session or Disconnect from a Session.
Public Sub ConnectDisconnectBeginSession()
    
    On Error Resume Next
    
    'If the TCP socket is closed.
    If sckTCP(0).State = sckClosed Then
        
        'Ensure all port numbers and IP addresses are valid. This is
        'applicable for both Hosts and Clients.
        If ValidateBeforeConnection Then
            
            'Set TCP and UDP port numbers.
            InetSes.LocalTcpPort = GetChosenTcpPort
            InetSes.LocalUdpPort = CLng(txtUdpPort.Text)
            
        Else
            
            'Port or IP numbers are invalid.
            chkHideSession.Enabled = False
            chkHideSession.Value = vbUnchecked
            Exit Sub
        End If
    End If
    
    'No cheats at all during network wars.
    Call TheMainForm.TurnOffCheatCodes
    
    'If the main TCP socket is open, this is a "Disconnect" command.
    If sckTCP(0).State <> sckClosed Then
        
        'Close all connections (host and client) and remove the session
        'from the Indexing Server if required.
        Call IxServerCloseSession(eIxCommand.CLOSE_SESSION, False)
        Exit Sub
        
    'If I am a client, this is a "Connect" command.
    ElseIf optJoin.Value Then
        
        Call DisplaySessionLocator
        
        'Connect to the session.
        'chkHideSession.Enabled = False
        'chkHideSession.Value = vbUnchecked
        'Call ConnectTo
    
    'Else I must be a host, this must be "Begin" command.
    Else
        
        'Begin a new session.
        Call BeginNewSession
        
    End If
End Sub

'I will shortly be a Host. Start a new Session.
Private Sub BeginNewSession()
    On Error Resume Next
    
    'Start a new session.
    'Check the terminal name is not blank.
    If Trim(txtTerminalName.Text) = "" Then
        
        'If blank, use the Winsock localhost name.
        txtTerminalName.Text = sckTCP(0).LocalHostName
    End If
    
    'Check the terminal name is not blank again. Winsock is buggy.
    If Trim(txtTerminalName.Text) = "" Then
        
        'If still blank, Winsock failed. Use "localhost" as the terminal name.
        txtTerminalName.Text = "localhost"
        
    End If
    
    'Check a Session Name has been chosen.
    If Len(Trim(txtSesName.Text)) = 0 Then
        txtSesName.Text = PickRandomSesname
    End If
    
    'Save the session name in the registry.
    SaveSetting gcApplicationName, "settings", "LastSesName", Trim(txtSesName.Text)
    
    'Warn if the broadcast port is different from the well known broadcast port.
    If CLng(txtUdpPort.Text) <> gcDefaultPortNumber _
    And optLan.Value = vbChecked Then
        WriteText "**Warning** Broadcast port is not " & str(gcDefaultPortNumber) & "."
    End If
    
    'Set TCP socket to listening mode.
    Call ListenFor
    
    'Set UDP socket to listen for client broadcast messages.
    Call netFindHosts.BroadcastListen
    
    'If an Internet session.
    If optInet.Value Then
        
        'Post host details to the Indexing Server.
        Call PostInetHostDetails
    Else
        
        'LAN only session.
        chkHideSession.Enabled = False
        chkHideSession.Value = vbUnchecked
        WriteText "Ready.", True
    End If
End Sub

'Post host details to internet.
Public Sub PostInetHostDetails()
    If optInet.Value Then
        'Post host details to internet.
        If IxServerPostNewSession Then
            tmrKeepAlive.Enabled = True
        Else
            tmrKeepAlive.Enabled = False
            'WriteText "Could not post host name.", True
            Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, True)
            optLan.Value = True
            optInet.Value = True
            optHost.Value = True
            chkHideSession.Enabled = False
            chkHideSession.Value = vbUnchecked
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

'TODO: Refactor, comment and document.
Private Sub Form_Load()
    Dim tmp As Integer
    Dim i As Long
    Dim f1 As Integer
    Dim vLine As String
    Dim ownersName As String
    Dim vForfeitInterval As String
    
    On Error Resume Next
    
    'Show the IP Info tab only if we are in develop mode.
    If Not gcAppDevelopMode Then
        tabInfo.Tabs.Remove 4
    End If
    
    TheMainForm.mnuNetCntr.Enabled = True
    
    vForfeitInterval = GetSetting(gcApplicationName, "settings", "NetForfeit", "XX")
    If Not IsNumeric(vForfeitInterval) Then
        vForfeitInterval = 180
        SaveSetting gcApplicationName, "settings", "NetForfeit", vForfeitInterval
    End If
    vscrollTimeLimit.Value = vForfeitInterval
    
    gHomeWebPage = gcDefaultHomePageClearURL
    gHelpWebPage = gcDefaultHelpPageURL
    gDownloadWebPage = gcDefaultDownloadPageURL
    gIndexServerUrl = gcIndexServerURL
    
    'Create a (hopefully) unique local session ID. The number is 7 digits long.
    'Remove decimal and comma symbol for international compatability.
    InetSes.LocalSessID = Replace(CStr(GenRandom4 * 100), ".", "")
    InetSes.LocalSessID = Replace(InetSes.LocalSessID, ",", "")
    
    lastSendTime(0) = 0
    Packets(0).bPacket = ""
    With RemoteNetRegData(0)
    .RegCode = ""
    .HostIP = ""
    .HostName = ""
    .ValidPassword = False
    .PasswordTrys = 0
    .AppVersion = ""
    .HostID = ""
    .VotesAgainst = ""
    '.LeKey = 0
    '.LeSlot = 0
    '.LeSlotSpin = 0
    '.LeType = 0
    End With
    With RemoteNetRegData(1)
    .RegCode = ""
    .HostIP = ""
    .HostName = ""
    .ValidPassword = False
    .PasswordTrys = 0
    .AppVersion = ""
    .HostID = ""
    .VotesAgainst = ""
    '.LeKey = 0
    '.LeSlot = 0
    '.LeSlotSpin = 0
    '.LeType = 0
    End With
    With MyNetRegData
    .RegCode = evalChk.RegCode
    .HostIP = sckTCP(0).LocalIP
    .HostName = sckTCP(0).LocalHostName
    .ValidPassword = False
    .PasswordTrys = 0
    .AppVersion = GetVersionInfo
    .HostID = GetUniqueId
    .VotesAgainst = ""
    .LeType = 0
    .LeKey = Int(GenRandom4 * &H7D) + 1
    .LeSlot = Int(GenRandom4 * &H7D) + 1
    .LeSlotSpin = Int(GenRandom4 * &H7D) + 1
    'Debug.Print Int(GenRandom4 * &HFF) + 1
    End With
    
    txtTcpPort.Text = Trim(GetSetting(gcApplicationName, "settings", "TcpPort", CStr(gcDefaultPortNumber)))
    txtUdpPort.Text = Trim(GetSetting(gcApplicationName, "settings", "UdpPort", CStr(gcDefaultPortNumber)))
    InetSes.LocalUdpPort = Trim(CLng(txtUdpPort.Text))
    
    txtTerminalName.Text = GetSetting(gcApplicationName, "settings", "TermName", sckTCP(0).LocalHostName)
    If Len(Trim(txtTerminalName.Text)) = 0 Then
        txtTerminalName.Text = "localhost"
    End If
    myTerminalNumber = 0
    connectionSpeed = dfltSpeed                 'Slowest for now
    sndComplete(0) = 0                          '0 when clear to send
    Call TheMainForm.resetPlayerOwners
    txtChat.Text = ""
    tmrFillInfo.Tag = ""
    
    tmp = CInt(GetSetting(gcApplicationName, "settings", "rfrsRateV29", "0"))
    
    optRefresh(tmp).Value = True
    netMain.Height = 6005
    netMain.Width = 6405
    Inet.Locked = False
    txtSesName.Text = GetSetting(gcApplicationName, "settings", "LastSesName", Trim(warSit.filename))
    'If Len(Trim(txtSesName.Text)) = 0 Then
    '    txtSesName.Text = PickRandomSesname
    'End If
    'netFindHosts.tmrBroadcast.Enabled = False
    
    ConnectSoundFile = GetSetting(gcApplicationName, "settings", "ConnectSoundFile", App.Path & "\ding.wav")
    
    txtMaxPlayers.Text = GetSetting(gcApplicationName, "settings", "NetMaxPlayerClaim", txtMaxPlayers.Text)
    vscrollMaxPlayers.Value = CInt(txtMaxPlayers.Text)
    txtMaxConnections.Text = GetSetting(gcApplicationName, "settings", "NetMaxConFromIP", txtMaxConnections.Text)
    vscrollMaxConnections.Value = CInt(txtMaxConnections.Text)
    chkPasswordSession.Value = GetSetting(gcApplicationName, "settings", "NetEnforcePassword", chkPasswordSession.Value)
    chkWlcmMsg.Value = GetSetting(gcApplicationName, "settings", "NetShowWelcomeMessage", chkWlcmMsg.Value)
    
    'Setup tabbed box.
    For i = 0 To frameInfo.Count - 1
        frameInfo(i).Move tabInfo.ClientLeft, tabInfo.ClientTop, _
                    tabInfo.ClientWidth, tabInfo.ClientHeight
    Next
    frameInfo(tabInfo.SelectedItem.Index - 1).ZOrder 0
    
    'Load TCP Sockets.
    For i = 1 To MaxConnections + 1
        Load sckTCP(i)
    Next
    
    txtWelcomeMsg.Text = LoadConfigFile(cNetWelcomeMessageFile)
    Call chkWlcmMsg_Click

    Call chkPasswordSession_Click
    
    txtPassword.Text = GetSetting(gcApplicationName, "settings", "NetPassword", "war")
    txtUserName.Text = GetSetting(gcApplicationName, "settings", "IxAccountName", "")
    txtUserName.Tag = gGsLeUtils.LE6d(GetSetting(gcApplicationName, "settings", "IxPasswordHash", ""))
    
    lsvConnections.ColumnHeaders.Clear
    lsvConnections.ColumnHeaders.Add , , "Connection", (lsvConnections.Width / 10) * 3
    lsvConnections.ColumnHeaders.Add , , "Display Name", (lsvConnections.Width / 10) * 7 - 80
    
    frmHostOptions.Visible = True
    frmHostOptions.Top = pctNetConnections.Height - frmHostOptions.Height
    
    'Retrieve banned list from file if it exists.
    If Len(Dir(GetConfigDataDir & cBannedListFile)) > 0 Then
        f1 = FreeFile
        Open GetConfigDataDir & cBannedListFile For Input As f1
        Do While Not EOF(f1)
            Line Input #f1, vLine
            If InStr(1, Inet.BannedList, vLine) = 0 Then
                Inet.BannedList = Inet.BannedList & vLine & vbCrLf
            End If
        Loop
        Inet.BannedList = Replace(Inet.BannedList, vbCrLf & vbCrLf, vbCrLf)
        Close f1
    End If
    chkHideSession.Enabled = False
    chkHideSession.Value = vbUnchecked
    Call setLanguage
    Call optJoin_Click
    Call EnableInternetOptions(optInet.Value)
    tmrCheckGSNews.Enabled = True
End Sub

Private Sub Form_Resize()
    Exit Sub
    If Me.WindowState = 1 Then Exit Sub
    If Me.ScaleWidth < 500 Then Exit Sub
    If Me.ScaleHeight < 500 Then Exit Sub
    txtChat.Width = Me.ScaleWidth
    txtChat.Height = Me.ScaleHeight - 200
End Sub

'TODO: Refactor, comment and document.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim tmp
    Dim f1 As Integer
    
    On Error GoTo ErrHand
    
    'Save last port numbers.
    SaveSetting gcApplicationName, "settings", "TcpPort", Trim(txtTcpPort.Text)
    SaveSetting gcApplicationName, "settings", "UdpPort", Trim(txtUdpPort.Text)
    
    SaveSetting gcApplicationName, "settings", "TermName", txtTerminalName.Text

    SaveSetting gcApplicationName, "settings", "NetPassword", txtPassword.Text
    SaveSetting gcApplicationName, "settings", "IxAccountName", txtUserName.Text
    
    SaveSetting gcApplicationName, "settings", "ConnectSoundFile", ConnectSoundFile
    
    SaveSetting gcApplicationName, "settings", "NetMaxPlayerClaim", txtMaxPlayers.Text
    SaveSetting gcApplicationName, "settings", "NetMaxConFromIP", txtMaxConnections.Text
    SaveSetting gcApplicationName, "settings", "NetForfeit", vscrollTimeLimit.Value
    SaveSetting gcApplicationName, "settings", "NetEnforcePassword", chkPasswordSession.Value
    SaveSetting gcApplicationName, "settings", "NetShowWelcomeMessage", chkWlcmMsg.Value
    
    'Save the welcome message to a file.
    SaveConfigFile cNetWelcomeMessageFile, txtWelcomeMsg.Text
    
    If TheMainForm.Visible Then
        TheMainForm.SetFocus
    End If
    
    'If the close request came from the form control menu then just hide.
    If UnloadMode = vbFormControlMenu Then
        Me.Hide
        Cancel = -1
        Exit Sub
    
    'If close request came from elsewhere and the main TCP socket is not
    'closed or if there is an Internet session posted, ask for confirmation.
    ElseIf Not gServerMode _
    And Not gHeadlessMode _
    And (sckTCP(0).State <> sckClosed Or Trim(InetSes.ID) <> "") Then
        
        'Do you want to close all connections?
        If MsgBox(Phrase(239), vbYesNo, gcApplicationName) = vbNo Then
            
            Me.Hide
            
            'Cancel TheMainForm unload.
            TheMainForm.gCancelUnload = -1
            Cancel = -1
            Exit Sub
            
        Else
            
            'Tell TheMainForm that the user has already been asked about
            'closing all connections. Do not ask if they really want to
            'end the war. That would be annoying and most uncool.
            TheMainForm.gCancelUnload = -2
            
        End If
    End If
    
    'This causes GlobalSiege to hang when killed from the Task Scheduler,
    'so don't call it at all if in server mode. TODO: fix.
    If Not gHeadlessMode Then
        Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, True)
    End If
    
    netWorkSituation = cNetNone
    Call EnableMissionOptions
    TheMainForm.mnuNetCntr.Enabled = False
    Inet.LastCheckTime = 0
    If TheMainForm.Visible Then
        TheMainForm.SetFocus
    End If
    
    Exit Sub
ErrHand:
    Resume Next
End Sub

'I am client and request connection to remote host.
'TODO: Refactor, comment and document.
Public Sub ConnectTo(Optional pSessionID As String = "LAN")
    Dim vPort As Long
    
    On Error Resume Next
    
    Call TheMainForm.TurnOffCheatCodes
    
    netFindHosts.tmrBroadcast.Enabled = False
    
    'Notify the user that we are connecting to the session.
    WriteText Phrase(243) & InetSes.SesName & "...", True 'Connecting to <Sess Name>...
    
    sckTCP(0).Close
    sckTCP(0).RemoteHost = Trim(InetSes.RemoteIP)
    sckTCP(0).RemotePort = InetSes.RemoteTcpPort
    

    vPort = 0
    
    'Use a random local port. No good using a specfic port number
    'because it might be in use and could cause problems. Ports
    'take ages to close and if the same local port number is used
    'again too soon, an error is generated.
    'sckTCP(0).LocalPort = GetChosenTcpPort
    sckTCP(0).LocalPort = 0
    
    sckTCP(0).Connect
    
    'Use DoEvents here so that any TCP errors will come through from
    'sckTCP_Error() if it is an Address In Use error.
    DoEvents
    
    'If an error was detected, choose another port and try again.
    'This might be superflouse because it only happens when a specific
    'local port is chosen. By setting the local port to zero, the OS
    'chooses a free port for us. This should not affec router forwarding
    'because the client initiates the connection.
    If sckTCP(0).State = sckError Then
        sckTCP(0).Close
        sckTCP(0).LocalPort = 0
        sckTCP(0).Connect
        DoEvents
    End If
    
    'Change the button's caption to "Disconnect".
    cmdConnect.Caption = Phrase(242)   'Disconnect
    TheMainForm.mnuNetDisconnect.Enabled = True
    
    'Enable and disable various controls as required.
    Call enableButtons(False)
    Call EnableInternetOptions(False)
    
    'Set the network situation global as a client.
    netWorkSituation = cNetClient
    Call TheMainForm.ResetPlayerList
    Call TheMainForm.EnableSetupControls(False)
    Call IxServerJoinSession(pSessionID)
    Exit Sub
End Sub

    'Enable or disable all buttons
Public Sub enableButtons(OnOrOff As Boolean)
    Dim bClor As Long
    
    optJoin.Enabled = OnOrOff
    optHost.Enabled = OnOrOff
    optInet.Enabled = OnOrOff
    optLan.Enabled = OnOrOff
    TheMainForm.mnuNetHost.Enabled = OnOrOff
    TheMainForm.mnuNetClient.Enabled = OnOrOff
    TheMainForm.Toolbar1.Buttons(11).Enabled = OnOrOff
    
    txtSesName.Enabled = OnOrOff
    txtTerminalName.Enabled = OnOrOff And optLan.Value
    
    txtTcpPort.Enabled = OnOrOff
    txtUdpPort.Enabled = OnOrOff
    cmdRestore.Enabled = OnOrOff
End Sub

'Pick a random session name.
'TODO: Refactor, comment and document.
Public Function PickRandomSesname() As String
    Dim RandName(15) As String
    
    'Randomize
    
    '"Un-Named War"
    RandName(0) = "Pi(I\h`_R\m"
    '"WAR!!"
    RandName(1) = "R<M"
    '"way too cool for a name"
    RandName(2) = "r\tojj^jjgajm\i\h`"
    '"No Name Session"
    RandName(3) = "`gnFmpaajrjnc`\mSS"""
    '"War Baby YEAH"
    RandName(4) = "R\m=\]tT@<C"
    '"Some Random Session Name"
    RandName(5) = "Njh`M\i_jhN`nndjiI\h`"
    '"Real War"
    RandName(6) = "M`\gR\m"
    '"Insert name here"
    RandName(7) = "Njh`M\i_jhN`nndjiI\h`"
    '"Pyep..."
    RandName(8) = "Kt`k)))"
    '"I'm the best. Try to beat me."
    RandName(9) = "D""hoc`]`no)Omtoj]`\oh`)"
    '"Try to beat the best."
    RandName(10) = "Omtoj]`\ooc`]`no)"
    '"WE WANT YOU"
    RandName(11) = "R@R<IOTJP"
    '"Just Another War"
    RandName(12) = "Epno<ijoc`mR\m"
    '"Want a shot at the title?"
    RandName(13) = "R\io\ncjo\ooc`odog`:"
    '"THE MOTHER OF ALL BATTLES"
    RandName(14) = "OC@HJOC@MJA<GG=<OOG@N"
    '"do you have the guts for battle?"
    RandName(15) = "_jtjpc\q`oc`bponajm]\oog`:"
    
    PickRandomSesname = gGsLeUtils.LE4d(RandName(CInt(GenRandom4 * (UBound(RandName))))) '0 to 15
End Function

'Test function PickRandomSesname (above).
Private Sub xxx()
    Dim i As Long
    For i = 0 To 5000
        Debug.Print PickRandomSesname
    Next
End Sub

'I am a host, put me into listen mode.
Private Sub ListenFor()
    
    On Error GoTo ErrHand
    If Len(txtTerminalName.Text) = 0 Then
        txtTerminalName.Text = "localhost"
    End If
    net.ClientName(0) = txtTerminalName.Text
    Call TheMainForm.resetPlayerOwners
    
    sckTCP(0).Close
    DoEvents
    sckTCP(0).LocalPort = InetSes.LocalTcpPort
    sckTCP(0).Listen
    
    If sckTCP(0).LocalPort <> GetChosenTcpPort _
    And GetChosenTcpPort <> 0 Then
        WriteText "**Warning** Port " & CStr(GetChosenTcpPort) _
                    & " is in use. Now using " _
                    & CStr(sckTCP(0).LocalPort) _
                    & ".", True
    End If
    InetSes.LocalTcpPort = sckTCP(0).LocalPort
    txtTcpPort.Text = CStr(sckTCP(0).LocalPort)
    
    WriteText Phrase(248), True                 'Listening for connections.
    
    'MyTermNo = Phrase(235)                     'Host
    MyTermNo = txtTerminalName.Text
    cmdConnect.Caption = Phrase(242)            'Disconnect
    'cmdConnect.Caption = Phrase(336)             'Close
    TheMainForm.mnuNetDisconnect.Enabled = True
    netWorkSituation = cNetHost
    myTerminalNumber = 0
    Call TheMainForm.ResetPlayerList
    Call TheMainForm.EnableSetupControls(True)
    Call TheMainForm.resetPlayerOwners
    Call enableButtons(False)
    Call EnableInternetOptions(False)
    
    Exit Sub
ErrHand:
    If Err.Number = 10048 Then
        
        'Port in use, try to use another port.
        sckTCP(0).Close
        DoEvents
        sckTCP(0).LocalPort = 0
        'WriteText "**Warning** Port " & txtTcpPort.Text & " is not available."
        LogError "ListenFor", "Trying another port: " & Err.Number & " " & Err.Description
        Resume

    Else
        WriteText "Could not open the main port. " & Err.Description, True
        LogError "ListenFor", "Error: " & Err.Number & " " & Err.Description
        Exit Sub
    End If
End Sub

'TODO: Refactor, comment and document.
Private Sub optHost_Click()
    Static sAskedBefore As Boolean
    
    On Error Resume Next
    cmdConnect.Caption = Phrase(251)   'Begin
    'txtSesName.Text = GetSetting(gcApplicationName, "settings", "LastSesName", Trim(warSit.filename))
    txtSesName.Enabled = True
    optRefresh(0).Enabled = True
    optRefresh(1).Enabled = True
    optRefresh(2).Enabled = True
    'TheMainForm.mnuNetKill.Enabled = True
    'TheMainForm.mnuNetKillBan.Enabled = True
    
    'Only print IP info etc once, the first time host is clicked.
    If Not sAskedBefore Then
        WriteText Phrase(237) & Chr(34) & sckTCP(0).LocalHostName & Chr(34), False 'Terminal Host Name is
        'WriteText Phrase(421) & sckTCP(0).LocalIP, False                           'Terminal IP address is
        WriteText Phrase(421) & Replace(modNetwork.GetLocalHostIP, ",", ", "), False
        'WriteText Phrase(238) & vbCrLf, False                                      'Click <IP info...> for more information.
        sAskedBefore = True
    End If
    
    If Me.Visible Then
        cmdConnect.Default = True
        cmdConnect.SetFocus
    End If
    
    pctNetHostOptions.Enabled = True
    txtMaxPlayers.Enabled = True
    txtMaxConnections.Enabled = True
    lblMaxClaim.Enabled = True
    lblMaxCon.Enabled = True
    
    txtTimeLimit.Enabled = True
    vscrollTimeLimit.Enabled = True
    lblTimeLimit.Enabled = True
    
    TheMainForm.mnuAutoRestart.Enabled = True
    chkPasswordSession.Enabled = True
    chkWlcmMsg.Enabled = True
    chkHideSession.Enabled = False
    chkHideSession.Value = vbUnchecked
    vscrollMaxPlayers.Enabled = True
    vscrollMaxConnections.Enabled = True
    Call chkPasswordSession_Click
    Call chkWlcmMsg_Click
    
End Sub

'TODO: Refactor, comment and document.
Public Sub OptJoinClick()
    Dim vCtrl As Control
    
    On Error Resume Next
    'cmdConnect.Caption = Phrase(252)  'Connect
    cmdConnect.Caption = "&Find Sessions"
    txtSesName.Enabled = False
    'txtSesName.BackColor = &H8000000F
    'txtSesName.Text = ""
    optRefresh(0).Enabled = False
    optRefresh(1).Enabled = False
    optRefresh(2).Enabled = False
    If Me.Visible Then
        cmdConnect.Default = True
        cmdConnect.SetFocus
    End If
    
    pctNetHostOptions.Enabled = True
    txtMaxPlayers.Enabled = False
    txtMaxConnections.Enabled = False
    lblMaxClaim.Enabled = False
    lblMaxCon.Enabled = False
    
    txtTimeLimit.Enabled = False
    vscrollTimeLimit.Enabled = False
    lblTimeLimit.Enabled = False
    
    TheMainForm.mnuAutoRestart.Enabled = False
    chkPasswordSession.Enabled = False
    chkWlcmMsg.Enabled = False
    chkHideSession.Enabled = False
    chkHideSession.Value = vbUnchecked
    vscrollMaxPlayers.Enabled = False
    vscrollMaxConnections.Enabled = False
    Call chkPasswordSession_Click
    Call chkWlcmMsg_Click
End Sub

Private Sub optJoin_Click()
    Call OptJoinClick
End Sub

'When connection closed by remote machine
'TODO: Refactor, comment and document.
Public Sub sckTCP_Close(Index As Integer)
    On Error GoTo ErrHand
    sckTCP(Index).Close
    sndComplete(Index) = 0
    
    TheMainForm.cmdSetupOk.Tag = "sckTCP_Close."
    
    If Index = 0 Then
        'cmdConnect.Caption = Phrase(252)  'Connect
        cmdConnect.Caption = "&Find Sessions"
        enableButtons True
        EnableInternetOptions True
        netWorkSituation = cNetNone
        Call TheMainForm.ResetPlayerList
        Call TheMainForm.EnableSetupControls(True)
        Call TheMainForm.resetPlayerOwners
        WriteText Phrase(253), True   'Connection has been lost.
        Call netMain.DisplayConnectedTerminals("")
    Else
        If InStr(1, sckTCP(Index).Tag, "silent") > 0 _
        Or InStr(1, sckTCP(Index).Tag, "quiet") > 0 Then
            'Disconnect silently.
            sckTCP(Index).Tag = ""
        Else
            'Inform everyone about disconnection.
            WriteText net.ClientName(Index) & Phrase(254), True    'Terminal <name> has disconnected.
            Call XmitStringAll(CLng(Index), 1, 0, net.ClientName(Index) & Phrase(254) & vbCrLf)
        End If
        Call TheMainForm.lostPlayerOwner(CByte(Index))
        RemoteNetRegData(Index).RegCode = ""
        RemoteNetRegData(Index).HostIP = ""
        RemoteNetRegData(Index).HostName = ""
        RemoteNetRegData(Index).ValidPassword = False
        RemoteNetRegData(Index).PasswordTrys = 0
        RemoteNetRegData(Index).AppVersion = ""
        RemoteNetRegData(Index).HostID = ""
        RemoteNetRegData(Index).VotesAgainst = ""
        'RemoteNetRegData(Index).LeKey = 0
        'RemoteNetRegData(Index).LeSlot = 0
        'RemoteNetRegData(Index).LeSlotSpin = 0
        'RemoteNetRegData(Index).LeType = 0
        net.ClientName(Index) = ""
        
    End If
    Exit Sub
ErrHand:
    WriteText "Error - " & Err.Description, True
    LogError "sckTCP_Close", "Error: " & Err.Number & " " & Err.Description
    Resume Next
End Sub

'TODO: Refactor, comment and document.
Private Sub sckTCP_Connect(Index As Integer)
    'Don't talk to banned players.
    If IsBanned(ChooseValidIP(sckTCP(Index).RemoteHostIP, sckTCP(Index).RemoteHost)) Then
        On Error Resume Next

        If netWorkSituation = cNetClient Then
            WriteText "Can't connect. This host is in your banned list."
            Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, True)
        Else
            sckTCP(Index).Close
        End If
        Exit Sub
    End If
    WriteText Phrase(256), True  'Connection established.
End Sub

'Ws2 bug sometimes returns an incomplete IP address. If the local IP address returned
'from sckTCP(0) is not a complete IP address, try to find its closest match. What a bodge!
'TODO: Document.
Private Function GetValidLocalIP() As String
    Dim vIpIndex
    Dim vLocalIps() As String
    
    GetValidLocalIP = sckTCP(0).LocalIP
    
    If Not IsValidIP(GetValidLocalIP) Then
        
        'Assume the last octet was knocked off, try to find a close match.
        'This is real bodgy. Thanks heaps Microsoft.
        vLocalIps = Split(modNetwork.GetLocalHostIP, ",")
        For vIpIndex = 0 To UBound(vLocalIps)
            If InStr(1, vLocalIps(vIpIndex), GetValidLocalIP) Then
                GetValidLocalIP = vLocalIps(vIpIndex)
                Exit For
            End If
        Next
        
        'Local IP address is still not valid, just use the first IP address returned.
        If Not IsValidIP(GetValidLocalIP) Then
            GetValidLocalIP = vLocalIps(0)
        End If
    End If
End Function

'Ask host to tell me my name and terminal number and setup settings.
'Send reg info: "RegCode, LocalIP, LocalHostName".
Private Sub requestSetupScreen()
    Dim bBuff() As Byte
    Dim vValidIP As String
    
    If Len(txtTerminalName.Text) = 0 Then
        txtTerminalName.Text = "localhost"
    End If

    DoEvents
    Call XmitString(0, 9, 0, evalChk.RegCode & "," _
                  & Trim(GetValidLocalIP) & "," _
                  & EncodeNonAscii(CStr(sckTCP(0).LocalHostName)) & "," _
                  & EncodeNonAscii(txtTerminalName.Text))
    
End Sub

'Make room and insert command bytes and send byte array to remote ID.
'TODO: Refactor, comment and document.
Private Sub XmitData(pRemoteID As Long, rmoteCommand As Byte, rmotePlayer As Byte, bMessage() As Byte)
    Dim bytSend() As Byte
    Dim cntr As Long
    
    ReDim bytSend(UBound(bMessage) - LBound(bMessage) + 2) As Byte
    
    For cntr = LBound(bMessage) + 2 To UBound(bMessage)
        bytSend(cntr + 2) = bMessage(cntr)
    Next
    
    Call XmitBytes(pRemoteID, rmoteCommand, rmotePlayer, bytSend)
End Sub

'Make room and insert command bytes and send byte array to all not ID.
'TODO: Refactor, comment and document.
Private Sub XmitDataAll(myPortNmbr As Long, rmoteCommand As Byte, rmotePlayer As Byte, bMessage() As Byte)
    Dim bytSend() As Byte
    Dim cntr As Long
    
    ReDim bytSend(UBound(bMessage) - LBound(bMessage) + 2) As Byte
    
    For cntr = LBound(bMessage) + 2 To UBound(bMessage)
        bytSend(cntr + 2) = bMessage(cntr)
    Next
    
    Call XmitBytesAll(myPortNmbr, rmoteCommand, rmotePlayer, bytSend)
End Sub

'Send byte array to remote ID. First 2 bytes are reserved.
'Parameter pRemoteID used by host to send one particular
'terminal. Clients should pass 0.
'If boolean IsAlreadyPacked is true, data is already framed (packeted)
'and will be sent as is.
'Return false if failed.
'TODO: Refactor, comment and document.
Public Function XmitBytes(pRemoteID As Long, rmoteCommand As Byte, _
                           rmotePlayer As Byte, bMessage() As Byte, _
                           Optional IsAlreadyPacked As Boolean = False) As Boolean
    Dim vIndex As Long
    Dim vTimer1Hold As Boolean
    Dim vTimer2Hold As Boolean
    
    On Error GoTo ErrHand
    XmitBytes = False
    
    If netWorkSituation <> cNetNone Then
        
        'Error fixer. Yucky but works. This is needed particularly
        'when new players have joined and the socket hasn't been
        'loaded properly yet.
        For vIndex = 0 To 100000
            If UBound(sndComplete) >= MaxConnections Then
                Exit For
            End If
            DoEvents
        Next
    End If
    
    'Remember the timer settings on the main form.
    vTimer1Hold = TheMainForm.Timer1.Enabled
    vTimer2Hold = TheMainForm.Timer2.Enabled
    
    'Set a short xmit delay of 10 milliseconds (default) before
    'sending data on this socket again.
    lastSendTime(pRemoteID) = lastSendTime(pRemoteID) + xmitDelay
    
    'Stop the clock if the xmit delay has not been achieved. This could be
    'the case when the computer players are going really really fast.
    If GetTickCount < lastSendTime(pRemoteID) Then
        TheMainForm.Timer1.Enabled = False
        TheMainForm.Timer2.Enabled = False
    End If
    
    'Give time for last message to clear. Also account for GetTickCount
    'resetting which it does every 49th day. GetTickCount returns how many
    'milliseconds the system has been up.
    For vIndex = 0 To 100000
        If GetTickCount > lastSendTime(pRemoteID) Or (GetTickCount + 60000) < lastSendTime(pRemoteID) Then
            Exit For
        End If
        DoEvents
    Next
    
    'Remember the last time data was sent on this socket.
    lastSendTime(pRemoteID) = GetTickCount
    
    'Resume the computer players etc.
    TheMainForm.Timer1.Enabled = vTimer1Hold
    TheMainForm.Timer2.Enabled = vTimer2Hold
    
    'Check to be certain that the socket is connected.
    If sckTCP(pRemoteID).State = sckConnected Then
        
        sndComplete(pRemoteID) = 1
        
        'Help prevent unwanted callbacks.
        If Not net.LockControls Then
            Debug.Print "XmitBytes Command"; rmoteCommand; "pRemoteID"; pRemoteID; "rmotePlayer"; rmotePlayer; "Length"; UBound(bMessage) '", Time; "; Timer"
            LogInfo "XmitBytes", "Term=" & Format(pRemoteID, "00") _
                                & " PlrByt=" & Format(rmotePlayer, "00") _
                                & " ID=" & Format(pRemoteID, "00000") _
                                & " Cmd=" & Format(rmoteCommand, "00") _
                                & " Pkg=" & ToHexStr(bMessage), 5, True
            
            'Check if the data has already been packed.
            If IsAlreadyPacked Then
                
                'Data is already packed so send it directly.
                sckTCP(pRemoteID).SendData bMessage()
            Else
                
                'Data is not packed. Pack and send.
                bMessage(0) = rmoteCommand
                bMessage(1) = rmotePlayer
                SendBytes pRemoteID, bMessage()
            End If
        End If
        
        XmitBytes = True
    End If
    DoEvents
    Exit Function
ErrHand:
    
    LogError "XmitBytes", Err.Description, True
    Debug.Print "XmitBytes() Error: " & Err.Description
    Exit Function
    'Resume Next
End Function

'Flush data from network queues.
'TODO: Refactor, comment and document.
Private Sub tmrFlushQueue_Timer()
    Dim bPeek() As Byte
    Dim cntr As Integer
    
    On Error Resume Next
    
    tmrFlushQueue.Enabled = False
    For cntr = 0 To MaxConnections
        If cntr <> myTerminalNumber And sckTCP(cntr).State = sckConnected Then
            sckTCP(cntr).PeekData bPeek()
            If UBound(bPeek) > 0 Then
                Call sckTCP_DataArrival(cntr, 0)
            End If
        End If
    Next
End Sub

'Process data queues. Data is packed in packages (frames). Extract
'data then process.
'TODO: Refactor, comment and document.
Private Sub sckTCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim BytBuf() As Byte
    Dim BytHold() As Byte
    Dim BytFramed() As Byte
    Dim bTemp(1) As Byte
    Dim Length As Long
    Dim MarkerLen As Long
    Dim PackLen As Long
    Dim EndMarker As Long
    Dim CSum As Long
    Dim ID As Long
    Dim vRemoteHostIP As String
    Static InHere As Boolean
    
    On Error GoTo ErrHand
    
    vRemoteHostIP = ChooseValidIP(RemoteNetRegData(Index).HostIP, sckTCP(Index).RemoteHostIP, sckTCP(Index).RemoteHost)
    
    'Don't talk to banned players.
    If IsBanned(vRemoteHostIP, net.ClientName(Index), RemoteNetRegData(Index).HostID) Then
        If sckTCP(Index).Tag = "" Then
            sckTCP(Index).Tag = "silent"
        End If
        LogInfo "sckTCP_DataArrival", "IsBanned = TRUE - " _
                & vRemoteHostIP & ", " & net.ClientName(Index) _
                & ", " & RemoteNetRegData(Index).HostID, 0, True
        Exit Sub
    End If
    
    'Lock-------------
    If InHere Then
        tmrFlushQueue.Interval = 1000
        tmrFlushQueue.Enabled = True
        LogInfo "sckTCP_DataArrival", "Static variable InHere = TRUE."
        Exit Sub
    End If
    
    InHere = True
    '-----------------
    
    'Get data
    sckTCP(Index).GetData BytBuf()
    
    With Packets(Index)
    Call CopyBytes(.bPacket, BytBuf, UBound(.bPacket) + 1)
    'DoEvents
    
    Do While UBound(.bPacket) > 0
        PackLen = UBound(.bPacket)
        EndMarker = ByteToInt(.bPacket, 0)
        Length = ByteToInt(.bPacket, 2)
        bTemp(0) = .bPacket(0)
        bTemp(1) = .bPacket(1)
        MarkerLen = WhereInArray(.bPacket, bTemp, 2) - 9
        
        'Make sure packet is not junk.
        If MarkerLen > 0 And Length <> MarkerLen Then
            .bPacket = ""
            InHere = False
            LogInfo "sckTCP_DataArrival", "MarkerLen > 0 And Length <> MarkerLen - " _
                    & "MarkerLen = " & CStr(MarkerLen) & ", Length = " & CStr(Length)
            Exit Sub
        End If
        
        CSum = ByteToInt(.bPacket, 4)
        ID = ByteToInt(.bPacket, 6)
        If PackLen >= Length + 10 Then
            If .bPacket(Length + 9) = .bPacket(0) And .bPacket(Length + 10) = .bPacket(1) Then
                
                'Valid packet found.
                BytHold = MidB(.bPacket, 9, Length + 1)
                BytFramed = MidB(.bPacket, 1, Length + 11)
                .bPacket = MidB(.bPacket, Length + 12)
                'Debug.Print "Arrived: "; ToHexStr(BytHold)
                Call LedByteArray(BytHold, Index)
                Call UnCompress(BytHold)
                Call ProcessData(Index, ID, BytHold, BytFramed)
            Else
                'Junk.
                .bPacket = ""
                LogInfo "sckTCP_DataArrival", "Packet marked as junk (1)."
            End If
        ElseIf PackLen > 10000 Then
            'Junk
            .bPacket = ""
            LogInfo "sckTCP_DataArrival", "Packet marked as junk (2)."
        Else
            'Rest of the packet is still on its way..
            InHere = False
            LogInfo "sckTCP_DataArrival", "Rest of the packet is still on its way.", 5
            Exit Sub
        End If
    Loop
    End With
    InHere = False
    Exit Sub
ErrHand:
    LogError "sckTCP_DataArrival", "Error: " & CStr(Err.Number) & ", " & Err.Description, True
    Resume Next
End Sub

'Compress, package (frame) and send data.
'Allows chunk of data to be sent in different packets and reassembled when received.
'Also allows different chunks sent in same packed to be correctly extracted.
'Frame format:
'| End_Marker | Length | Checksum |   ID   | Payload .... | End_Marker |
'TODO: Refactor, comment and document.
Private Function SendBytes(pRemoteID As Long, bMessage() As Byte) As Boolean
    Dim bHold() As Byte
    Dim bPacket() As Byte
    Dim bPayload() As Byte
    Dim Length As Long
    Dim ID As Long
    Dim EndMarker As Long
    Dim bTmp() As Byte
    
    bPayload = bMessage
    
    'Temp-----------------
    'Dim x1, x2, x3
    'Dim xBefore As String
    'Dim xAfter As String
    'bTmp = bPayload
    'x1 = UBound(bTmp)
    'xBefore = ToHexStr(bTmp)
    'Debug.Print "Raw          : "; ToHexStr(bTmp)
    'Call Compress(bTmp)
    'x2 = UBound(bTmp)
    'Debug.Print "Compressed   : "; ToHexStr(bTmp)
    'Call UnCompress(bTmp)
    'xAfter = ToHexStr(bTmp)
    'Debug.Print "UnCompressed : "; ToHexStr(bTmp)
    'Debug.Print "Size         : "; CLng(x2 / x1 * 100); "%"
    'Debug.Print "Status       : ";
    'If xBefore = xAfter Then
    '    Debug.Print "OK"
    'Else
    '    Debug.Print "***** FAIL ******"
    'End If
    'End Temp-------------
    
    Call Compress(bPayload)
    Call LeByteArray(bPayload, CInt(pRemoteID))
    
    Length = UBound(bPayload)
    ReDim bPacket(Length + 10) As Byte
    Call IntToByte(Length, bHold)
    Call CopyBytes(bPacket, bHold, 2)
    
    Call CopyBytes(bPacket, bPayload, 8)
    
    ID = GetNextID
    Call IntToByte(ID, bHold)
    Call CopyBytes(bPacket, bHold, 6)
    
    bHold(0) = &H13
    bHold(1) = &H13
    Call CreateEndMarker(bPacket, bHold)
    Call CopyBytes(bPacket, bHold, 0)
    Call CopyBytes(bPacket, bHold, Length + 9)
    
    'Debug.Print ToHexStr(bPacket)
    sckTCP(pRemoteID).SendData bPacket()
End Function

'Encrypt payload.
Private Sub LeByteArray(ByRef pBytes() As Byte, pIndex As Integer)
    'Exit Sub
    With MyNetRegData
    '.LeKey = 5
    '.LeSlot = 6
    '.LeSlotSpin = 7
    Select Case .LeType
    Case 0
        'If .AppVersion > "03020037" Then
        '    Call HexStringToByteArray(pBytes, gGsLeUtils.LE5(ByteArrayToHexString(pBytes)))
        'End If
        Call HexStringToByteArray(pBytes, gGsLeUtils.LE6(ByteArrayToHexString(pBytes), _
            .LeKey, .LeSlot, .LeSlotSpin))
    End Select
    End With
End Sub

'Decrypt payload.
Private Sub LedByteArray(ByRef pBytes() As Byte, pIndex As Integer)
    'Exit Sub
    With MyNetRegData
    '.LeKey = 5
    '.LeSlot = 6
    '.LeSlotSpin = 7
    'Debug.Print pBytes
    Select Case .LeType
    Case 0
        'If RemoteNetRegData(pIndex).AppVersion > "03020037" Then
        '    Call HexStringToByteArray(pBytes, gGsLeUtils.LE5d(ByteArrayToHexString(pBytes)))
        'End If
        Call HexStringToByteArray(pBytes, gGsLeUtils.LE6d(ByteArrayToHexString(pBytes), _
            .LeKey, .LeSlot, .LeSlotSpin))
    End Select
    'Debug.Print pBytes
    End With
End Sub

'Compress the passed byte array.
'TODO: Refactor, comment and document.
Private Sub Compress(ByRef Bytes() As Byte)
    Dim cntr As Long
    Dim RptCnt As Long
    Dim NibCntr As Long
    Dim NewBytes() As Byte
    Dim NewPos As Long
    Dim LastByte As Byte
    
    'Debug.Print ToHexStr(Bytes)
    ReDim NewBytes(UBound(Bytes) * 2) As Byte
    NewPos = 0
    
    'Insert escape bytes.
    For cntr = 0 To UBound(Bytes)
        LastByte = Bytes(cntr)
        If (LastByte And &HF0) = &HE0 Then
            'EF is escape byte.
            NewBytes(NewPos) = &HEF
            NewBytes(NewPos + 1) = LastByte - &H10
            NewPos = NewPos + 2
        Else
            NewBytes(NewPos) = LastByte
            NewPos = NewPos + 1
        End If
    Next
    
    ReDim Preserve NewBytes(NewPos - 1) As Byte
    'Debug.Print ToHexStr(NewBytes)
    Bytes = NewBytes
    NewPos = 0
    
    'Find repeating bytes.
    For cntr = 0 To UBound(Bytes)
        LastByte = Bytes(cntr)
        RptCnt = 0
        
        'Find repeating.
        If cntr < UBound(Bytes) - 4 Then
            Do While LastByte = Bytes(cntr + RptCnt + 1) _
            And RptCnt < 250 _
            And RptCnt + cntr < UBound(Bytes) - 1
                RptCnt = RptCnt + 1
            Loop
        End If
        
        'Encode.
        If RptCnt > 4 Then
            NewBytes(NewPos) = &HEE
            NewBytes(NewPos + 1) = CByte(RptCnt)
            NewBytes(NewPos + 2) = LastByte
            NewPos = NewPos + 3
            cntr = cntr + RptCnt
        Else
            NewBytes(NewPos) = LastByte
            NewPos = NewPos + 1
        End If
    Next
    
    ReDim Preserve NewBytes(NewPos - 1) As Byte
    'Debug.Print ToHexStr(NewBytes)
    Bytes = NewBytes
    NewPos = 0
    
    'Find repeating bytes with the same upper nibble.
    For cntr = 0 To UBound(Bytes)
        LastByte = Bytes(cntr) And &HF0
        RptCnt = 0
        
        'Find repeating.
        If cntr < UBound(Bytes) - 4 And LastByte < &HE0 Then
            Do While LastByte = (Bytes(cntr + RptCnt + 1) And &HF0) _
            And RptCnt < 250 _
            And RptCnt + cntr < UBound(Bytes) - 1
                RptCnt = RptCnt + 1
            Loop
        End If
        
        'Encode.
        If RptCnt > 4 Then
            NewBytes(NewPos) = &HE0 + (LastByte \ &H10)
            NewBytes(NewPos + 1) = CByte(RptCnt)
            NewPos = NewPos + 2
            
            For NibCntr = cntr To cntr + RptCnt
                If (NibCntr - cntr) Mod 2 = 0 Then
                    NewBytes(NewPos) = (Bytes(NibCntr) And &HF) * &H10
                Else
                    NewBytes(NewPos) = NewBytes(NewPos) + (Bytes(NibCntr) And &HF)
                    NewPos = NewPos + 1
                End If
            Next
            If (NibCntr - cntr) Mod 2 = 1 Then
                NewPos = NewPos + 1
            End If
            cntr = cntr + RptCnt
        Else
            NewBytes(NewPos) = Bytes(cntr)
            NewPos = NewPos + 1
        End If
    Next
    
    ReDim Preserve NewBytes(NewPos - 1) As Byte
    'Debug.Print ToHexStr(NewBytes)
    Bytes = NewBytes
End Sub

'Uncompress the passed byte array.
'TODO: Refactor, comment and document.
Private Sub UnCompress(ByRef Bytes() As Byte)
    Dim cntr As Long
    Dim RptCnt As Long
    Dim RepeatTotal As Long
    Dim NewBytes() As Byte
    Dim NewPos As Long
    Dim LastByte As Byte
    Dim UpperByte As Byte
    
    'Debug.Print ToHexStr(Bytes)
    ReDim NewBytes(UBound(Bytes)) As Byte
    NewPos = 0
    
    'Extract repeating bytes with same upper nibble.
    For cntr = 0 To UBound(Bytes)
        
        'Find repeating marker (En, < EE).
        LastByte = Bytes(cntr)
        If (LastByte And &HF0) = &HE0 And LastByte < &HEE Then
        
            'Get upper byte.
            UpperByte = (Bytes(cntr) And &HF) * &H10
            
            'Get repeat count.
            RepeatTotal = Bytes(cntr + 1)
            
            'Make sure there is enough space.
            ReDim Preserve NewBytes(UBound(NewBytes) + CLng(Bytes(cntr + 1)) + 1) As Byte
            
            'Point to first data byte.
            cntr = cntr + 2
            
            For RptCnt = 0 To RepeatTotal
                If RptCnt Mod 2 = 0 Then
                    NewBytes(NewPos) = UpperByte + ((Bytes(cntr + (RptCnt \ 2)) And &HF0) \ &H10)
                    NewPos = NewPos + 1
                Else
                    NewBytes(NewPos) = UpperByte + (Bytes(cntr + (RptCnt \ 2)) And &HF)
                    NewPos = NewPos + 1
                End If
            Next
            cntr = cntr + (RepeatTotal \ 2)
        Else
            NewBytes(NewPos) = Bytes(cntr)
            NewPos = NewPos + 1
        End If
    Next
    
    ReDim Preserve NewBytes(NewPos - 1) As Byte
    Bytes = NewBytes
    NewPos = 0
    
    'Extract repeating.
    For cntr = 0 To UBound(Bytes)
        
        'Find repeating marker (EE).
        LastByte = Bytes(cntr)
        If LastByte = &HEE Then
            LastByte = Bytes(cntr + 2)
            ReDim Preserve NewBytes(UBound(NewBytes) + CLng(Bytes(cntr + 1)) + 1) As Byte
            For RptCnt = 0 To Bytes(cntr + 1)
                NewBytes(NewPos) = LastByte
                NewPos = NewPos + 1
            Next
            cntr = cntr + 2
        Else
            NewBytes(NewPos) = Bytes(cntr)
            NewPos = NewPos + 1
        End If
    Next
    
    ReDim Preserve NewBytes(NewPos - 1) As Byte
    Bytes = NewBytes
    NewPos = 0
    
    'Extract escaped bytes (EF).
    For cntr = 0 To UBound(Bytes)
        LastByte = Bytes(cntr)
        If LastByte = &HEF Then
            NewBytes(NewPos) = Bytes(cntr + 1) + &H10
            cntr = cntr + 1
            NewPos = NewPos + 1
        Else
            NewBytes(NewPos) = Bytes(cntr)
            NewPos = NewPos + 1
        End If
    Next
    
    ReDim Preserve NewBytes(NewPos - 1) As Byte
    Bytes = NewBytes
End Sub

'Create random two bytes that are not found in bPacket array
'to be used as an end marker.
Private Sub CreateEndMarker(bPacket() As Byte, bUnique() As Byte)
    Dim bUniqueNum(1) As Byte
    Dim vIX As Long
    
    For vIX = 0 To 10000
        If IsInArray(bPacket, bUnique) Then
            bUnique(0) = CByte(GenRandom4 * &HFF)
            bUnique(1) = CByte(GenRandom4 * &HFF)
        Else
            Exit For
        End If
    Next
End Sub

'Return unique number. Thread issues are extremely unlikely.
'TODO: Refactor, comment and document.
Private Function GetNextID() As Long
    Static ID As Long
    GetNextID = ID
    ID = ID + 1
    If ID > 65530 Then
        ID = 0
    End If
End Function

    'Send string to pRemoteID
    'TODO: Refactor, comment and document.
Public Sub XmitString(pRemoteID As Long, rmoteCommand As Byte, rmotePlayer As Byte, sMessage As String)
    Call XmitBytes(pRemoteID, rmoteCommand, rmotePlayer, StrConv("AA" & sMessage, vbFromUnicode))
End Sub

'Send byte array to all except me. First 2 bytes are reserved.
'No action is taken if I am a client. Parameter myPortNmbr is
'used by the host to pass data to clients, but not back to the
'client who originally sent the data to the host.
Private Sub XmitBytesAll(myPortNmbr As Long, rmoteCommand As Byte, _
rmotePlayer As Byte, bMessage() As Byte)
    Dim cntr As Long
    
    On Error Resume Next
    
    If CountTerminals = 0 Then
        Exit Sub
    End If
    
    For cntr = 1 To MaxConnections
        If (cntr <> myPortNmbr) And sckTCP(cntr).State = sckConnected Then
            Call XmitBytes(cntr, rmoteCommand, rmotePlayer, bMessage)
        End If
    Next
End Sub

'Xmit bytes to all depending on the client version number allowing seperation for
'backward compatability with older versions.
'If pVersionHigher is true, xmit to only clients with a version equal to or higher
'than pVersionLimit. If false, only xmit to versions lower than pVersionLimit.
Private Sub XmitBytesAllVersion(myPortNmbr As Long, rmoteCommand As Byte, _
rmotePlayer As Byte, pVersionHigher As Boolean, pVersionLimit As Long, bMessage() As Byte)
    Dim vTerminalIX As Long
    
    On Error Resume Next
    
    If CountTerminals > 0 Then
        For vTerminalIX = 1 To MaxConnections
            If (vTerminalIX <> myPortNmbr) And sckTCP(vTerminalIX).State = sckConnected Then
                If pVersionHigher And RemoteNetRegData(vTerminalIX).AppVersion >= pVersionLimit _
                Or Not pVersionHigher And RemoteNetRegData(vTerminalIX).AppVersion < pVersionLimit Then
                    Call XmitBytes(vTerminalIX, rmoteCommand, rmotePlayer, bMessage)
                End If
            End If
        Next
    End If
End Sub

'Same as XmitBytesAll above except data is already framed by remote
'client. Used to forward data directly from client to all other clients
'by host.
Private Sub XmitPacketAll(myPortNmbr As Long, rmoteCommand As Byte, _
                          rmotePlayer As Byte, bMessage() As Byte)
    Dim cntr As Long
    
    If CountTerminals = 0 Then
        Exit Sub
    End If
    
    For cntr = 1 To MaxConnections
        If (cntr <> myPortNmbr) And sckTCP(cntr).State = sckConnected Then
            Call XmitBytes(cntr, rmoteCommand, rmotePlayer, bMessage, True)
        End If
    Next
End Sub

    'Send string to all except me
Public Sub XmitStringAll(myPortNmbr As Long, rmoteCommand As Byte, rmotePlayer As Byte, sMessage As String)
    Call XmitBytesAll(myPortNmbr, rmoteCommand, rmotePlayer, StrConv("AA" & sMessage, vbFromUnicode))
End Sub

Public Sub XmitStringAllVersion(myPortNmbr As Long, rmoteCommand As Byte, _
rmotePlayer As Byte, pVersionHigher As Boolean, pVersionLimit As Long, sMessage As String)
    Call XmitBytesAllVersion(myPortNmbr, rmoteCommand, rmotePlayer, pVersionHigher, _
                            pVersionLimit, StrConv("AA" & sMessage, vbFromUnicode))
End Sub

'Display error information.
'TODO: Refactor, comment and document.
Private Sub sckTCP_Error(Index As Integer, ByVal Number As Integer, _
Description As String, ByVal Scode As Long, _
ByVal Source As String, ByVal HelpFile As String, _
ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    Dim vPort As Long
    
    On Error Resume Next
    
    ' Not an error.
    If Number = sckSuccess Then
        Exit Sub
    End If
    
    'Address in use error.
    If Number = 10048 And Index = 0 Then
        sckTCP(Index).LocalPort = 0
        Exit Sub
    End If
    
    sndComplete(Index) = 0
    WriteText "Terminal " & str(Index) & " Error " & Number & " " & Description, True
    
    sckTCP(Index).Close
    sckTCP(Index).Tag = ""
    sndComplete(Index) = 0
    
    If Index = 0 Then
        'cmdConnect.Caption = Phrase(252)  'Connect
        cmdConnect.Caption = "&Find Sessions"
        enableButtons True
        EnableInternetOptions True
        netWorkSituation = cNetNone
        Call TheMainForm.ResetPlayerList
        Call TheMainForm.EnableSetupControls(True)
        Call TheMainForm.resetPlayerOwners
        'WriteText Phrase(253), True   'Connection has been lost.
    Else
        WriteText net.ClientName(Index) & Phrase(254), True    'Terminal <name> has disconnected.
        Call TheMainForm.lostPlayerOwner(CByte(Index))
        RemoteNetRegData(Index).RegCode = ""
        RemoteNetRegData(Index).HostIP = ""
        RemoteNetRegData(Index).HostName = ""
        RemoteNetRegData(Index).ValidPassword = False
        RemoteNetRegData(Index).PasswordTrys = 0
        RemoteNetRegData(Index).AppVersion = ""
        RemoteNetRegData(Index).HostID = ""
        RemoteNetRegData(Index).VotesAgainst = ""
    End If
    Exit Sub
End Sub

Private Sub sckTCP_SendComplete(Index As Integer)
    sndComplete(Index) = 0              'Clear to send
End Sub

'Update host posting on internet every 10 minutes.
'TODO: comment and document.
Private Sub tmrKeepAlive_Timer()
    On Error Resume Next
    Call IxServerRefreshSession
End Sub

'The user has changed either chkPasswordSession or chkHideSession.
'TODO: comment and document.
Private Sub tmrControlChange_Timer()
    On Error Resume Next
    tmrControlChange.Enabled = False
    Call IxServerOptionsChanged
End Sub

Private Sub txtUdpPort_GotFocus()
    Call SelectAndHighlightText(txtUdpPort)
End Sub

Private Sub txtPassword_GotFocus()
    Call SelectAndHighlightText(txtPassword)
End Sub

Private Sub txtTcpPort_GotFocus()
    Call SelectAndHighlightText(txtTcpPort)
End Sub

    'Keep track of key presses, if CR then xmit string
    'TODO: Refactor, comment and document.
Private Sub txtChat_KeyPress(KeyAscii As Integer)
    Dim x As Long
    Static strSend As String
    
    If CountTerminals = 0 Then
        If sckTCP(0).State <> sckConnected Then
            Exit Sub
        End If
    End If
    
    If KeyAscii = Asc(vbCr) Then
        If CountTerminals = 0 Then
            Call XmitString(0, 1, myTerminalNumber, strSend & vbCrLf)
            strSend = ""
            Exit Sub
        End If
        
        Call XmitStringAll(0, 1, myTerminalNumber, strSend & vbCrLf)
        strSend = ""
    ElseIf (KeyAscii = vbKeyBack) Or (KeyAscii = vbKeyDelete) Then
        strSend = Left(strSend, Len(strSend) - 1)
    Else
        strSend = strSend & Chr(KeyAscii)
    End If
End Sub

    'Return the delay time required between xmits to the same port.
    'TODO: Refactor, comment and document.
Private Function xmitDelay() As Long
    xmitDelay = xmitMilliSeconds _
        * Abs(optRefresh(0).Value _
        + optRefresh(1).Value * 2 _
        + optRefresh(2).Value * 4)
End Function

'------------------------------------------------------------------------------------
' Tabbed box functions.
Private Sub tabInfo_Click()
    frameInfo(tabInfo.SelectedItem.Index - 1).ZOrder 0
End Sub

'Returns the number of players terminal pTerminalIndex owns.
Public Function CountPlayersTermOwns(pTerminalIndex As Integer) As Integer
    Dim vPlayerIndex As Long
    
    CountPlayersTermOwns = 0
    
    For vPlayerIndex = 0 To 5
        If net.playerOwner(vPlayerIndex) = pTerminalIndex Then
            CountPlayersTermOwns = CountPlayersTermOwns + 1
        End If
    Next
End Function

'Find any disconnected terminals that own players and terminate properly.
Private Sub FindDeadPlayerOwners()
    Dim vTerminalIndex As Integer
    
    If netWorkSituation = cNetHost Then
        For vTerminalIndex = 1 To MaxConnections
            If sckTCP(vTerminalIndex).State = sckClosed _
            And CountPlayersTermOwns(vTerminalIndex) > 0 Then
                Call sckTCP_Close(vTerminalIndex)
            End If
        Next
    End If
End Sub

'Refresh tabbed boxes. Check for unwanted connections.
'TODO: Refactor, document.
Private Sub tmrFillInfo_Timer()
    Dim vTerminalIndex As Long
    
    On Error Resume Next
    
    'Re enable Vote button for the next player.
    If IsNumeric(cmfForfeit.Tag) Then
        If cmfForfeit.Tag <> CStr(gPlayerTurn) Then
            cmfForfeit.Tag = ""
            cmfForfeit.Enabled = CountPlayersTermOwns(CInt(myTerminalNumber)) > 0 _
                                Or netWorkSituation <> cNetClient
        End If
    End If
    
    'Update the connections list box.
    Call ListInfo
    
    'Actions for client terminals
    If netWorkSituation = cNetClient Then
        'Ensure cheat codes are off.
        Call TheMainForm.TurnOffCheatCodes
        
        'Change Connections Action frame's caption to "Vote".
        frmHostOptions.Caption = "Vote"
        netChatterBox.cmdVote.Caption = "Vote"
        
    'Actions for Host terminals.
    Else
    
        'Change Connections Action frame's caption to "Actions".
        frmHostOptions.Caption = "Actions"
        netChatterBox.cmdVote.Caption = "&Take Action"
        
        'Find any dead connections that own players.
        FindDeadPlayerOwners
    End If
    
    'Disable the vote button on the chat box if I don't actually own any players.
    netChatterBox.cmdVote.Enabled = netWorkSituation = cNetHost _
                                    Or CountPlayersTermOwns(CInt(myTerminalNumber)) > 0
    
    'Enable / disable the hide sessions check box depending on connection status.
    If optInet.Value And netWorkSituation = cNetHost Then
        chkHideSession.Enabled = True
    Else
        chkHideSession.Enabled = False
        chkHideSession.Value = vbUnchecked
    End If
    
    On Error Resume Next
    
    'Give some time before killing connection un silent mode.
    'This is to make it more difficult for wood be hackers.
    For vTerminalIndex = 1 To MaxConnections
        If sckTCP(vTerminalIndex).Tag = "silent" Then
            sckTCP(vTerminalIndex).Tag = "silent kill"
        ElseIf sckTCP(vTerminalIndex).Tag = "silent kill" Then
            'sckTCP(vTerminalIndex).Close
            sckTCP(vTerminalIndex).Tag = ""
            Call sckTCP_Close(CInt(vTerminalIndex))
        End If
    Next
    
    'Check Action/Vote buttons
    Call CheckActionVoteButtons
    
    'tmrFillInfo.Tag contains client audit results to send to the host after a time delay.
    If tmrFillInfo.Tag <> "" Then
        Call netMain.XmitString(0, 1, 0, tmrFillInfo.Tag)
        Call TheMainForm.PostMessage(tmrFillInfo.Tag)
        tmrFillInfo.Tag = ""
    End If
    
    'Ensure controls are enabled/disabled as required.
    'Call EnableInternetOptions(optInet.Value)
End Sub

'Fill out all tabbed boxes. Called from form load and timer tmrFillInfo_Timer()
'every 10 seconds. Also calledfrom Form_Activate().
Private Sub ListInfo()
    Dim vIpList As String
    Dim IPAdr() As String
    Dim vListHold As String
    Dim i As Long
    Dim vText As String
    
    On Error Resume Next
    
    'Text for the IP Config tab.
    vText = ""
    
    'Get the local hostname.
    vText = vText & " Local Hostname" & vbTab & GetLocalHostName & vbCrLf
    
    'Get the local IP address and if known, the IP address as seen from the Index Server.
    vIpList = modNetwork.GetLocalHostIP & "," & InetSes.IP
    vIpList = Replace(vIpList, "x", ",")
    vIpList = CleanList(vIpList, ",")
    'Debug.Print vIpList
    
    IPAdr = Split(vIpList, ",")
    For i = 0 To UBound(IPAdr)
        vText = vText & " Local Host IP" & vbTab & IPAdr(i) & vbCrLf
    Next
    
    'Broadcast Port.
    vText = vText & " Bradcast Port" & vbTab & "Loc: " & netFindHosts.sckUDP.LocalPort _
                    & ", Rmt: " & netFindHosts.sckUDP.RemotePort & vbCrLf
    
    vText = vText & " Main Port" & vbTab & vbTab & "Loc: " & netMain.sckTCP(0).LocalPort _
                    & ", Rmt: " & netMain.sckTCP(0).RemotePort & vbCrLf
    
    If sckTCP.Count > 1 Then
        For i = 1 To sckTCP.Count - 1
            If sckTCP(i).State <> sckClosed Then
                vText = vText & " Terminal " & CStr(i) & vbTab & "Loc: " _
                                    & sckTCP(i).LocalPort & ", Rmt: " & sckTCP(i).RemotePort & vbCrLf
            End If
        Next
    End If
    
    If vText <> rtIpConfig.Tag Then
        rtIpConfig.Text = ""
        rtIpConfig.SelFontSize = 8.25
        rtIpConfig.SelText = vText
        rtIpConfig.Tag = vText
    End If
    
    'If I am the host, check terminal connections and send to connected
    'clients if there has been a change.
    If netWorkSituation = cNetHost Or netWorkSituation = cNetNone Then
        Call DisplayConnectedTerminals(ListConnectedTerminals)
    End If
    
End Sub

'pTerminalList format is IP<tab>Hostname<tab>Termname<cr>.
'If I am the host, this gets called from ListInfo().
'If I am a client, this gets called from ProcessData Command 19.
'The terminal list is saved in the listinfo's tag. If the passed terminal list is different
'to the contents of the tag or the network status, the list is updated. If I am the host,
'the list is sent to all connected clients.
'Static vNetStatus helps detect changes in the networking status.
'TODO: Refactor, comment and document.
Public Sub DisplayConnectedTerminals(pTerminalList As String)
    Static vNetStatus As Long
    Dim i As Long
    Dim vLines() As String
    Dim vParts() As String
    Dim L As ListItem
    Dim vStatus As String
    Dim vIndex As Long
    
    On Error Resume Next

    If lsvConnections.Tag <> pTerminalList Or vNetStatus <> netWorkSituation Then
        vNetStatus = netWorkSituation
        
        With lsvConnections
            If .ListItems.Count > 0 Then
                vIndex = .SelectedItem.Index
            Else
                vIndex = 0
            End If
            .ListItems.Clear
            
            vLines = Split(pTerminalList, vbCrLf)
            
            For i = 0 To UBound(vLines) - 1
                vParts = Split(vLines(i), ",")
                If UBound(vParts) = 2 Then
                    
                    'Determine status of connection (host/client)
                    If netWorkSituation = cNetNone Then
                        vStatus = "Localhost"
                    ElseIf netWorkSituation = cNetClient And vParts(2) = "0" Then
                        vStatus = "Host"
                    ElseIf netWorkSituation = cNetHost And vParts(2) = "0" Then
                        vStatus = "Host"
                    Else
                        vStatus = "Terminal " & vParts(2)
                    End If
                    
                    Set L = .ListItems.Add(, vParts(0) & vParts(2), vStatus)
                    L.SubItems(1) = DecodeNonAscii(vParts(1))
                    L.Tag = vLines(i)
                End If
            Next
            .Tag = pTerminalList
            If vIndex > 0 And vIndex <= .ListItems.Count Then
                .ListItems(vIndex).EnsureVisible
            End If
            
            'Notify all connected terminals if I am the host.
            If netWorkSituation = cNetHost Then
                Call XmitStringAll(0, 19, 0, pTerminalList)
            End If
        End With
    End If
End Sub

'Select connection to kill.
Private Sub lsvConnections_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Call EnableActionVoteButtons(Item.Tag)
End Sub

'Enable vote/action buttons as required of the selected item.
'Regularly called from the timer
Private Sub CheckActionVoteButtons()
    If Not lsvConnections.SelectedItem Is Nothing Then
        Call EnableActionVoteButtons(lsvConnections.SelectedItem.Tag)
    End If
End Sub

'Enable vote/action buttons as required from the selected item's tag
'and put the item's tag into the kill button's tag.
Private Sub EnableActionVoteButtons(SelectedItemTag As String)
    Dim vParts() As String
    Dim vPlayersMyTermOwns As Integer
    
    vParts = Split(SelectedItemTag, ",")
    vPlayersMyTermOwns = CountPlayersTermOwns(CInt(myTerminalNumber))
    
    'Cannot kill or vote against the host, but can vote against your self.
    If CLng(vParts(2)) <> 0 Then
        cmdKill.Tag = SelectedItemTag
        cmdKill.Enabled = netWorkSituation <> cNetClient _
                        Or vPlayersMyTermOwns > 0
        cmdBan.Enabled = netWorkSituation <> cNetClient
    Else
        cmdKill.Enabled = False
        cmdBan.Enabled = False
        cmdKill.Tag = ""
        'cmfForfeit.Enabled = False
    End If
    If IsNumeric(cmfForfeit.Tag) Then
        If cmfForfeit.Tag <> CStr(gPlayerTurn) Then
            cmfForfeit.Tag = ""
            cmfForfeit.Enabled = vPlayersMyTermOwns > 0 Or netWorkSituation <> cNetClient
        End If
    End If
    'cmfForfeit.Enabled = vPlayersMyTermOwns > 0 Or netWorkSituation <> cNetClient
End Sub

'An item from lsvConnections has been programatically selected and
'the action buttons neet to be enabled/disabled as required.
Public Sub lsvConnectionsItemSelected(pListItem As Long)
    Call lsvConnections_ItemClick(lsvConnections.ListItems(pListItem))
End Sub

'If I am the host, force the immediate forfeit of the current player's turn.
'If I am a client, vote to forfeit the current player's turn.
'Remember that this terminal has voted for the current player and disable the
'command button until the it is the next player's turn. The player memory
'is in cmfForfeit.Tag and the timer tmrFillInfo event updates the button state.
Private Sub cmfForfeit_Click()
    Dim vConfirm As VbMsgBoxResult
    Dim BytBuf() As Byte
    
    If netWorkSituation = cNetClient Then
        'I am a client.
        
        vConfirm = MsgBox("Vote the current player to forfeit their turn?", _
                    vbYesNo, "Confirm Vote")
        If vConfirm = vbYes Then
            If net.playerOwner(gPlayerTurn - 1) = myTerminalNumber Then
                'I am a client and have forfeited my own player. Well, OK!
                Call TheMainForm.ForfeitTurn
            Else
                'Vote actions.
                Call netMain.XmitString(0, 24, 0, CStr(net.playerOwner(gPlayerTurn - 1)) & ",FORFEIT")
                cmfForfeit.Enabled = False
                cmfForfeit.Tag = CStr(gPlayerTurn)
            End If
        End If
    Else
        'I am a host.
        vConfirm = MsgBox("Force the current player to forfeit their turn?", _
                    vbYesNo, "Confirm Action")
        If vConfirm = vbYes Then
            If net.playerOwner(gPlayerTurn - 1) = myTerminalNumber Then
                'Oops, I, the host own this player and I have forfeited.
                Call TheMainForm.ForfeitTurn
            Else
                Call netMain.XmitString(CLng(net.playerOwner(gPlayerTurn - 1)), 1, 0, _
                    "\beep Your turn has been forfeited." & vbCrLf)
                ReDim BytBuf(2) As Byte
                Call netMain.XmitBytes(CLng(net.playerOwner(gPlayerTurn - 1)), 21, 0, BytBuf)
            End If
        End If
    End If
End Sub

'Kill the player. If I am a client, send a vote to kill the player.
'Show confirmation box first.
Private Sub cmdKill_Click()
    Dim vConfirm As VbMsgBoxResult
    
    If netWorkSituation = cNetClient Then
        
        'I am a client.
        vConfirm = MsgBox("Vote to kill the selected terminal and ban from this session?", _
                    vbYesNo, "Confirm Vote")
        If vConfirm = vbYes Then
            Call VoteToKillConnection
        End If
    Else
        
        'I am a host.
        vConfirm = MsgBox("Kill the selected terminal and ban from this session?", _
                    vbYesNo, "Confirm Action")
        If vConfirm = vbYes Then
            Call KillConnection
        End If
    End If
End Sub

'Vote to kill the connection and ban them from this session.
Public Sub VoteToKillConnection()
    Dim vParts() As String
    
    vParts = Split(cmdKill.Tag, ",")
    If UBound(vParts) >= 2 Then
        Call netMain.XmitString(0, 24, 0, vParts(2) & ",KILL")
    End If
End Sub

'Kill the connection placed in cmdKill.Tag and ban for the life of
'this session. The connection will not be placed in the permanent ban list file.
Public Sub KillConnection()
    Dim vParts() As String
    Dim vNewBan As String
    Dim vRemoteID As String
    
    On Error Resume Next
    
    vParts = Split(cmdKill.Tag, ",")
    If UBound(vParts) >= 2 Then
        vRemoteID = RemoteNetRegData(CLng(vParts(2))).HostID
        
        'Kill connection.
        If netWorkSituation = cNetHost Then
            Call sckTCP_Close(CInt(vParts(2)))
        Else
            Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, True)
        End If
        
        'Put in the session banned list.
        vNewBan = gGsLeUtils.LE2(vParts(0), &HD7) & "," & vParts(1) & "," & vRemoteID
        Inet.KilledList = Inet.KilledList & vNewBan & vbCrLf
    End If
    cmdKill.Enabled = False
    cmdBan.Enabled = False
End Sub

'Return TRUE if this IP is banned. Check unique ID too and add to banned list if match.
Public Function IsBanned(ByVal pIP As String, _
Optional pClientName As String = "", _
Optional pUniqueID As String = "") As Boolean
    Dim vClassC As String
    Dim vList As String
    Dim vNewBan As String
    Dim vIP As String
    On Error Resume Next
    
    vClassC = Mid(pIP, 1, InStrRev(pIP, ".")) & "*"
    vClassC = gGsLeUtils.LE2(vClassC, &HD7)
    vIP = gGsLeUtils.LE2(pIP, &HD7)
    If InStr(1, Inet.BannedList, vIP & ",") > 0 _
    Or InStr(1, Inet.BannedList, vClassC & ",") > 0 _
    Or InStr(1, Inet.KilledList, vIP & ",") > 0 _
    Or (Trim(pUniqueID) <> "" And ((InStr(1, Inet.BannedList, pUniqueID) > 0) _
                                Or InStr(1, Inet.KilledList, pUniqueID) > 0)) Then
        If IsValidIP(pIP) Then
            IsBanned = True
        End If
    Else
        IsBanned = False
    End If
    
    'If unique ID found but IP not in list then add IP to list.
    If Trim(pUniqueID) <> "" Then
        If InStr(1, Inet.BannedList, pUniqueID) > 0 _
        And InStr(1, Inet.BannedList, vIP) = 0 Then
            vNewBan = vIP & "," & pClientName & "," & pUniqueID
            Inet.BannedList = Inet.BannedList & vNewBan & vbCrLf
            IsBanned = True
            Call SaveBannedList
        End If
    End If
    
    'Log it if this client is banned.
    If IsBanned Then
        Call LogInfo("", "IsBanned(""" & vIP & """,""" & pClientName & """,""" & pUniqueID & """): " _
                & " attempted to connect and is in the banned list.")
    End If
End Function

'Kill and ban player.
'TODO: Refactor, comment and document.
Private Sub cmdBan_Click()
    Dim vParts() As String
    Dim vResponse As VbMsgBoxResult
    Dim vClassC As String
    Dim vNewBan As String
    Dim i As Long
    Dim vReturn As String
    Dim vRemoteID As String
    
    On Error Resume Next
    
    'Only the host can ban clients for ever..
    If netWorkSituation = cNetClient Then
        Call cmdKill_Click
        Exit Sub
    End If
    
    vParts = Split(cmdKill.Tag, ",")
    If UBound(vParts) >= 2 Then
        vRemoteID = RemoteNetRegData(CLng(vParts(2))).HostID
        
        vResponse = MsgBox("Kill the selected terminal and ban from this host for ever?", _
                    vbYesNo, "Confirm Action")
        If vResponse = vbNo Then
            Exit Sub
        End If
        
        'Kill the connection in cmdKill.Tag.
        Call KillConnection
        
        'Put in banned list.
        vClassC = Mid(vParts(0), 1, InStrRev(vParts(0), ".")) & "*"
        vResponse = MsgBox("Would you like to ban the player's domain?" & vbCrLf & vbCrLf _
                        & "Select yes to ban the player's whole class C domain (" & vClassC & ")." & vbCrLf _
                        & "Select no to ban the player's IP address only (" & vParts(0) & ").", _
                    vbYesNoCancel)
        If vResponse = vbCancel Then
            Exit Sub
        ElseIf vResponse = vbYes Then
            vNewBan = gGsLeUtils.LE2(vClassC, &HD7) & "," & vParts(1) & "," & vRemoteID
        Else
            vNewBan = gGsLeUtils.LE2(vParts(0), &HD7) & "," & vParts(1) & "," & vRemoteID
        End If
        Inet.BannedList = Inet.BannedList & vNewBan & vbCrLf
        
        If optInet.Value Then
            
            'Notify the Index Server server of banning.
            Call IxServerPlayerWasBanned(vParts(0) & "," & gGsLeUtils.LE6(vParts(1)), "Banned by the host")
        End If
        Call SaveBannedList
    End If
End Sub

'Save banned list to file.
'TODO: Refactor, comment and document.
Private Sub SaveBannedList()
    Dim f1 As Integer
    
    f1 = FreeFile
    Open GetConfigDataDir & cBannedListFile For Output As f1
    Print #f1, Inet.BannedList
    Close #f1
End Sub

'------------------------------------------------------------------------------------

'Convert string to hex string.
'TODO: Refactor, comment and document.
Public Function BytToHex(Pkt As String) As String
    Dim cntr As Long
    Dim vChar As Long
    
    For cntr = 1 To Len(Pkt)
        vChar = CLng(Asc(Mid(Pkt, cntr, 1)))
        BytToHex = BytToHex & Hex(vChar)
    Next
End Function

'Convert hex string into string.
'TODO: Refactor, comment and document.
Public Function HexToByte(pHex As String) As String
    Dim i As Long
    'HexToByte = ""
    For i = 0 To Len(pHex) \ 2 - 1
        HexToByte = HexToByte & Chr(Format("&h" & Mid(pHex, i * 2 + 1, 2)))
    Next
End Function

'Return decimal ACSII code.
'TODO: Refactor, comment and document.
Private Function ToDecimal(Hx As String) As Long
    ToDecimal = Asc(UCase(Hx))
    If ToDecimal >= 65 Then
        ToDecimal = (ToDecimal And 7) + 9
    Else
        ToDecimal = ToDecimal And 15
    End If
End Function

'Convert byte array into hex string.
'TODO: Refactor, comment and document.
Private Function ToHexStr(Pkt() As Byte) As String
    Dim cntr As Long
    Dim vChar As Long
    Dim vByt As String
    
    For cntr = 0 To UBound(Pkt)
        vByt = Hex(Pkt(cntr))
        If Len(vByt) < 2 Then
            vByt = "0" & vByt
        End If
        ToHexStr = ToHexStr & vByt & ":"
    Next
End Function

'TODO: Refactor, comment and document.
'** TODO Request a refresh if client after language has changed.
Public Sub setLanguage()
    With netMain
        .Caption = Phrase(293)  'Global Siege network setup
        .lblNetConnectionOptions.Caption = " " & Phrase(294) & " " 'Options
        '.frameConnectType.Caption = Phrase(295) 'Connection type
        '.frmRefresh.Caption = Phrase(296) 'Refresh rate
        '.frameSettings(0).Caption = Phrase(297)    'Settings
        '.lblStatus.Caption = Phrase(298) 'Session history
        '.lblRemName.Caption = Phrase(299) 'Name or IP address of host
        .txtPortNmbr.Caption = Phrase(300)  'Port number
        .optJoin.Caption = Phrase(301)  'Join a war
        .optHost.Caption = Phrase(302)  'Host a war
        .optRefresh(0).Caption = Phrase(306)    'High
        .optRefresh(1).Caption = Phrase(307)    'Medium
        .optRefresh(2).Caption = Phrase(308)   'Low
        '.cmdIpInfo.Caption = Phrase(309) '&IP Info...
        '.cmdConnect.Caption = Phrase(310) '&Connect
        .cmdConnect.Caption = "&Find Sessions"
        '.cmdOK.Caption = Phrase(311) '&OK
        .cmdOK.Caption = Phrase(334) '&Hide
        
        .frameConnectionOptions.ToolTipText = Phrase(312) 'Network war must have 1 host
        lblNetConnectionOptions.ToolTipText = Phrase(312) 'Network war must have 1 host
        .pctNetConnectionOptions3.ToolTipText = Phrase(314) 'The frequency at which remote players are updated
        .optRefresh(0).ToolTipText = Phrase(314) 'The frequency at which remote players are updated
        .optRefresh(1).ToolTipText = Phrase(314) 'The frequency at which remote players are updated
        .optRefresh(2).ToolTipText = Phrase(314) 'The frequency at which remote players are updated
        '.lblRemName.ToolTipText = Phrase(315) 'Enter the name of the host terminal
        .txtPortNmbr.ToolTipText = Phrase(316) 'All terminals must use the same port number
        .optJoin.ToolTipText = Phrase(317)  'Connect to a listening host
        .optHost.ToolTipText = Phrase(318)  'Become a host and listen for connections
        '.cmdIpInfo.ToolTipText = Phrase(319) 'Display host name and IP configuration for this terminal.
        .cmdConnect.ToolTipText = "Open the Session Locator" 'Phrase(320) 'Connect with these settings
        .cmdOK.ToolTipText = Phrase(321) 'Hide the network setup dialog box without disconnecting
        .txtChat.ToolTipText = Phrase(322)  'Lists connected terminals, TCP errors, etc...
    End With
End Sub

Private Sub txtSesName_GotFocus()
    Call SelectAndHighlightText(txtSesName)
End Sub

Private Sub txtTerminalName_GotFocus()
    Call SelectAndHighlightText(txtTerminalName)
End Sub

'TODO: Refactor, comment and document.
Private Sub vscrollMaxConnections_Change()
    On Error Resume Next
    txtMaxConnections.Text = CStr(vscrollMaxConnections.Value)
End Sub
'TODO: Refactor, comment and document.
Private Sub vscrollMaxPlayers_Change()
    On Error Resume Next
    txtMaxPlayers.Text = CStr(vscrollMaxPlayers.Value)
End Sub
'TODO: Refactor, comment and document.
Private Sub vscrollTimeLimit_Change()
    On Error Resume Next
    txtTimeLimit.Text = CStr(vscrollTimeLimit.Value)
    If gPlayerTurn > 0 Then
        Call UpdateForfeitTimer(CByte(gPlayerTurn - 1))
    End If
End Sub
