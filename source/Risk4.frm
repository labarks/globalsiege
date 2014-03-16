VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form TheMainForm 
   AutoRedraw      =   -1  'True
   Caption         =   "<Var.ExeName>"
   ClientHeight    =   10020
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14085
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Risk4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   Palette         =   "Risk4.frx":7612
   PaletteMode     =   2  'Custom
   ScaleHeight     =   668
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   939
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrDrawWin 
      Enabled         =   0   'False
      Left            =   0
      Top             =   3720
   End
   Begin VB.Timer tmrFlashInfoBox 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   4200
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   1005
      ButtonWidth     =   1085
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   21
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Setup"
            Key             =   "toolNew"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset"
            Key             =   "toolReset"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "toolOpen"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "toolSave"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fast"
            Key             =   "toolFast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Undo"
            Key             =   "toolUndo"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Net"
            Key             =   "toolNet"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Key             =   "toolFind"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Chat"
            Key             =   "toolChat"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Mission"
            Key             =   "toolMission"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Conts"
            Key             =   "toolConts"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.Frame frmTst 
         Caption         =   "                        SB     CMod PTurn  T1      T2         Checkpoint"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   6735
         Begin VB.CommandButton cmdWorkoutOdds 
            Caption         =   "Estimate Odds"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   138
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox txtTst 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   26
            Text            =   "CM"
            Top             =   135
            Width           =   375
         End
         Begin VB.TextBox txtTst 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   25
            Text            =   "PT"
            Top             =   135
            Width           =   375
         End
         Begin VB.TextBox txtTst 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   24
            Text            =   "T1"
            Top             =   135
            Width           =   375
         End
         Begin VB.TextBox txtTst 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   23
            Text            =   "T2"
            Top             =   135
            Width           =   375
         End
         Begin VB.TextBox txtTst 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   22
            Text            =   "CP"
            Top             =   135
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3081
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   21
            Text            =   "4"
            Top             =   135
            Width           =   255
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Mask"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   135
            Width           =   615
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrWatchDog 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   4680
   End
   Begin VB.Timer tmrFindCPUspeed 
      Enabled         =   0   'False
      Left            =   0
      Top             =   3240
   End
   Begin VB.Timer TimerWatch 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   2760
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00BB6302&
      FillStyle       =   0  'Solid
      Height          =   9750
      Left            =   360
      ScaleHeight     =   646
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1236
      TabIndex        =   11
      Top             =   600
      Width           =   18600
      Begin VB.CommandButton cmdEnglish 
         Appearance      =   0  'Flat
         Caption         =   "EN"
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
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox SetupScreen 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9135
         Left            =   360
         ScaleHeight     =   9135
         ScaleWidth      =   19305
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   19305
         Begin VB.Frame frameSetupControls 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8535
            Left            =   360
            TabIndex        =   15
            Top             =   240
            Width           =   18855
            Begin VB.Frame frameSetup 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5430
               Index           =   4
               Left            =   2880
               TabIndex        =   150
               Top             =   5880
               Width           =   7335
               Begin VB.Frame frmContValues 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1935
                  Left            =   840
                  TabIndex        =   157
                  ToolTipText     =   "The War options box."
                  Top             =   2520
                  Width           =   5655
                  Begin VB.PictureBox pctContValues 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1455
                     Left            =   120
                     ScaleHeight     =   1455
                     ScaleWidth      =   5355
                     TabIndex        =   158
                     Top             =   360
                     Width           =   5355
                     Begin VB.VScrollBar udContVal 
                        Height          =   365
                        Index           =   5
                        Left            =   3600
                        Max             =   0
                        Min             =   50
                        TabIndex        =   176
                        Top             =   1080
                        Value           =   3
                        Width           =   255
                     End
                     Begin VB.TextBox txtContVal 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   5
                        Left            =   3120
                        Locked          =   -1  'True
                        TabIndex        =   174
                        Text            =   "3"
                        Top             =   1080
                        Width           =   495
                     End
                     Begin VB.VScrollBar udContVal 
                        Height          =   365
                        Index           =   4
                        Left            =   3600
                        Max             =   0
                        Min             =   50
                        TabIndex        =   173
                        Top             =   600
                        Value           =   8
                        Width           =   255
                     End
                     Begin VB.TextBox txtContVal 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   4
                        Left            =   3120
                        Locked          =   -1  'True
                        TabIndex        =   171
                        Text            =   "8"
                        Top             =   600
                        Width           =   495
                     End
                     Begin VB.VScrollBar udContVal 
                        Height          =   365
                        Index           =   3
                        Left            =   3600
                        Max             =   0
                        Min             =   50
                        TabIndex        =   170
                        Top             =   120
                        Value           =   4
                        Width           =   255
                     End
                     Begin VB.TextBox txtContVal 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   3
                        Left            =   3120
                        Locked          =   -1  'True
                        TabIndex        =   168
                        Text            =   "4"
                        Top             =   120
                        Width           =   495
                     End
                     Begin VB.VScrollBar udContVal 
                        Height          =   365
                        Index           =   2
                        Left            =   600
                        Max             =   0
                        Min             =   50
                        TabIndex        =   167
                        Top             =   1080
                        Value           =   6
                        Width           =   255
                     End
                     Begin VB.TextBox txtContVal 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   2
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   165
                        Text            =   "6"
                        Top             =   1080
                        Width           =   495
                     End
                     Begin VB.VScrollBar udContVal 
                        Height          =   365
                        Index           =   1
                        Left            =   600
                        Max             =   0
                        Min             =   50
                        TabIndex        =   164
                        Top             =   600
                        Value           =   3
                        Width           =   255
                     End
                     Begin VB.TextBox txtContVal 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   1
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   162
                        Text            =   "3"
                        Top             =   600
                        Width           =   495
                     End
                     Begin VB.VScrollBar udContVal 
                        Height          =   365
                        Index           =   0
                        Left            =   600
                        Max             =   0
                        Min             =   50
                        TabIndex        =   161
                        Top             =   120
                        Value           =   6
                        Width           =   255
                     End
                     Begin VB.TextBox txtContVal 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   0
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   159
                        Text            =   "6"
                        Top             =   120
                        Width           =   495
                     End
                     Begin VB.Label lblContVal 
                        AutoSize        =   -1  'True
                        Caption         =   "Australia"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   5
                        Left            =   3960
                        TabIndex        =   175
                        Top             =   1080
                        Width           =   915
                     End
                     Begin VB.Label lblContVal 
                        AutoSize        =   -1  'True
                        Caption         =   "Asia"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   4
                        Left            =   3960
                        TabIndex        =   172
                        Top             =   600
                        Width           =   480
                     End
                     Begin VB.Label lblContVal 
                        AutoSize        =   -1  'True
                        Caption         =   "Africa"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   3
                        Left            =   3960
                        TabIndex        =   169
                        Top             =   120
                        Width           =   615
                     End
                     Begin VB.Label lblContVal 
                        AutoSize        =   -1  'True
                        Caption         =   "Europe"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   2
                        Left            =   960
                        TabIndex        =   166
                        Top             =   1080
                        Width           =   765
                     End
                     Begin VB.Label lblContVal 
                        AutoSize        =   -1  'True
                        Caption         =   "South America"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   1
                        Left            =   960
                        TabIndex        =   163
                        Top             =   600
                        Width           =   1545
                     End
                     Begin VB.Label lblContVal 
                        AutoSize        =   -1  'True
                        Caption         =   "North America"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   0
                        Left            =   960
                        TabIndex        =   160
                        Top             =   120
                        Width           =   1500
                     End
                  End
                  Begin VB.Label lblContValues 
                     AutoSize        =   -1  'True
                     Caption         =   " Battalion conscriptions from occupied continents "
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
                     Height          =   270
                     Left            =   120
                     TabIndex        =   211
                     Top             =   0
                     Width           =   5130
                  End
               End
               Begin VB.Frame frmNewUnits 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2415
                  Left            =   840
                  TabIndex        =   151
                  ToolTipText     =   "The War options box."
                  Top             =   0
                  Width           =   5655
                  Begin VB.PictureBox pctNewUnits 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   2055
                     Left            =   120
                     ScaleHeight     =   2055
                     ScaleWidth      =   5475
                     TabIndex        =   152
                     Top             =   240
                     Width           =   5475
                     Begin VB.VScrollBar udNewUnitClac 
                        Height          =   365
                        Index           =   0
                        Left            =   1680
                        Max             =   1
                        Min             =   42
                        TabIndex        =   194
                        Top             =   120
                        Value           =   5
                        Width           =   255
                     End
                     Begin VB.TextBox txtNewUnitClac 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H8000000F&
                        BorderStyle     =   0  'None
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   0
                        Left            =   1320
                        Locked          =   -1  'True
                        TabIndex        =   193
                        Text            =   "5"
                        Top             =   120
                        Width           =   375
                     End
                     Begin VB.VScrollBar udNewUnitClac 
                        Height          =   365
                        Index           =   2
                        Left            =   2160
                        Max             =   0
                        Min             =   50
                        TabIndex        =   192
                        Top             =   840
                        Value           =   3
                        Width           =   255
                     End
                     Begin VB.VScrollBar udNewUnitClac 
                        Height          =   365
                        Index           =   1
                        Left            =   360
                        Max             =   0
                        Min             =   50
                        TabIndex        =   191
                        Top             =   480
                        Value           =   1
                        Width           =   255
                     End
                     Begin VB.TextBox txtNewUnitClac 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H8000000F&
                        BorderStyle     =   0  'None
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   2
                        Left            =   1800
                        Locked          =   -1  'True
                        TabIndex        =   189
                        Text            =   "3"
                        Top             =   840
                        Width           =   375
                     End
                     Begin VB.TextBox txtNewUnitClac 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H8000000F&
                        BorderStyle     =   0  'None
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   1
                        Left            =   0
                        Locked          =   -1  'True
                        TabIndex        =   187
                        Text            =   "1"
                        Top             =   480
                        Width           =   375
                     End
                     Begin VB.Label lblNewUnitClac 
                        Caption         =   "mobilized at the start of each player's turn."
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   5
                        Left            =   240
                        TabIndex        =   210
                        Top             =   1200
                        Width           =   4635
                        WordWrap        =   -1  'True
                     End
                     Begin VB.Label lblNewUnitClac 
                        AutoSize        =   -1  'True
                        Caption         =   "battalion(s)"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   4
                        Left            =   2460
                        TabIndex        =   195
                        Top             =   840
                        Width           =   1155
                     End
                     Begin VB.Label lblNewUnitClac 
                        AutoSize        =   -1  'True
                        Caption         =   "a minimum of"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   3
                        Left            =   240
                        TabIndex        =   190
                        Top             =   840
                        Width           =   1395
                     End
                     Begin VB.Label lblNewUnitClac 
                        AutoSize        =   -1  'True
                        Caption         =   "battalion(s) will be drafted with"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   2
                        Left            =   660
                        TabIndex        =   188
                        Top             =   480
                        Width           =   3120
                     End
                     Begin VB.Label lblNewUnitClac 
                        AutoSize        =   -1  'True
                        Caption         =   "country(s) occupied"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   1
                        Left            =   1980
                        TabIndex        =   186
                        Top             =   120
                        Width           =   2025
                     End
                     Begin VB.Label lblNewUnitClac 
                        AutoSize        =   -1  'True
                        Caption         =   "For every"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   0
                        Left            =   240
                        TabIndex        =   156
                        Top             =   120
                        Width           =   975
                     End
                  End
                  Begin VB.Label lblNewUnits 
                     AutoSize        =   -1  'True
                     Caption         =   " New recruit calculation "
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
                     Height          =   270
                     Left            =   120
                     TabIndex        =   153
                     Top             =   0
                     Width           =   2445
                  End
               End
            End
            Begin VB.Frame frameSetup 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4455
               Index           =   2
               Left            =   9000
               TabIndex        =   96
               Top             =   0
               Width           =   7335
               Begin VB.Frame frmWorkoutOdds 
                  BeginProperty Font 
                     Name            =   "MS Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2775
                  Left            =   4080
                  TabIndex        =   128
                  Top             =   1680
                  Width           =   3135
                  Begin VB.PictureBox pctWorkoutOdds 
                     AutoRedraw      =   -1  'True
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "Comic Sans MS"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   2295
                     Left            =   120
                     ScaleHeight     =   2295
                     ScaleWidth      =   2895
                     TabIndex        =   129
                     Top             =   360
                     Width           =   2895
                     Begin VB.Label Label1 
                        AutoSize        =   -1  'True
                        Caption         =   "The probability of winning attacks with the chosen dice settings:"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   11.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   765
                        Left            =   240
                        TabIndex        =   140
                        Top             =   120
                        Width           =   2400
                        WordWrap        =   -1  'True
                     End
                     Begin VB.Label lblAttackProb 
                        AutoSize        =   -1  'True
                        Caption         =   "50%"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   48
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   1095
                        Left            =   600
                        TabIndex        =   139
                        Top             =   1080
                        Width           =   1755
                        WordWrap        =   -1  'True
                     End
                  End
                  Begin VB.Label lblOdds 
                     AutoSize        =   -1  'True
                     Caption         =   " Attack probability"
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
                     Left            =   240
                     TabIndex        =   130
                     Top             =   0
                     Width           =   1695
                  End
               End
               Begin VB.Frame frmDiceRules 
                  BeginProperty Font 
                     Name            =   "MS Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   4335
                  Left            =   0
                  TabIndex        =   123
                  Top             =   120
                  Width           =   3735
                  Begin VB.PictureBox Picture2 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   3975
                     Left            =   120
                     ScaleHeight     =   3975
                     ScaleWidth      =   3495
                     TabIndex        =   124
                     Top             =   240
                     Width           =   3495
                     Begin VB.OptionButton optDiceRules 
                        Caption         =   "No dice"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   11.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   0
                        Left            =   0
                        TabIndex        =   209
                        Top             =   120
                        Width           =   3200
                     End
                     Begin VB.OptionButton optDiceRules 
                        Caption         =   "Attacker wins draw"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   11.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   1
                        Left            =   0
                        TabIndex        =   208
                        Top             =   360
                        Width           =   3200
                     End
                     Begin VB.OptionButton optDiceRules 
                        Caption         =   "Defender wins draw"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   11.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   2
                        Left            =   0
                        TabIndex        =   207
                        Top             =   600
                        Value           =   -1  'True
                        Width           =   3200
                     End
                     Begin VB.OptionButton optDiceRules 
                        Caption         =   "Both retreat when draw"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   11.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   3
                        Left            =   0
                        TabIndex        =   206
                        Top             =   840
                        Width           =   3200
                     End
                     Begin VB.OptionButton optDiceRules 
                        Caption         =   "Both lose when draw"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   11.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   4
                        Left            =   0
                        TabIndex        =   205
                        Top             =   1080
                        Width           =   3200
                     End
                     Begin VB.CheckBox chkSortDice 
                        Caption         =   "Sort dice before battle"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   11.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Left            =   0
                        TabIndex        =   204
                        ToolTipText     =   "Sort dice before comparing."
                        Top             =   1800
                        Value           =   1  'Checked
                        Width           =   3200
                     End
                     Begin VB.PictureBox pctSameDice 
                        BorderStyle     =   0  'None
                        BeginProperty Font 
                           Name            =   "MS Serif"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   1095
                        Left            =   0
                        ScaleHeight     =   1095
                        ScaleWidth      =   3495
                        TabIndex        =   126
                        Top             =   2280
                        Width           =   3495
                        Begin VB.OptionButton optDiceSame 
                           Caption         =   "No difference"
                           BeginProperty Font 
                              Name            =   "Arial"
                              Size            =   11.25
                              Charset         =   0
                              Weight          =   400
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   270
                           Index           =   0
                           Left            =   0
                           TabIndex        =   203
                           Top             =   360
                           Value           =   -1  'True
                           Width           =   3200
                        End
                        Begin VB.OptionButton optDiceSame 
                           Caption         =   "Win battle"
                           BeginProperty Font 
                              Name            =   "Arial"
                              Size            =   11.25
                              Charset         =   0
                              Weight          =   400
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   270
                           Index           =   1
                           Left            =   0
                           TabIndex        =   202
                           ToolTipText     =   "Instantly win the dice roll."
                           Top             =   600
                           Width           =   3200
                        End
                        Begin VB.OptionButton optDiceSame 
                           Caption         =   "Lose battle"
                           BeginProperty Font 
                              Name            =   "Arial"
                              Size            =   11.25
                              Charset         =   0
                              Weight          =   400
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   270
                           Index           =   2
                           Left            =   0
                           TabIndex        =   201
                           ToolTipText     =   "Instantly loose the dice roll."
                           Top             =   840
                           Width           =   3200
                        End
                        Begin VB.Label lblSameDice 
                           AutoSize        =   -1  'True
                           Caption         =   "All same dice thrown"
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
                           Height          =   270
                           Left            =   0
                           TabIndex        =   127
                           Top             =   0
                           Width           =   2145
                        End
                     End
                  End
                  Begin VB.Label lblDiceRules 
                     AutoSize        =   -1  'True
                     Caption         =   " Rules "
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
                     Height          =   270
                     Left            =   120
                     TabIndex        =   125
                     Top             =   0
                     Width           =   705
                  End
               End
               Begin VB.Frame frmDiceThrows 
                  BeginProperty Font 
                     Name            =   "MS Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1335
                  Left            =   4080
                  TabIndex        =   117
                  Top             =   120
                  Width           =   3135
                  Begin VB.PictureBox pctDiceThrown 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   975
                     Left            =   120
                     ScaleHeight     =   975
                     ScaleWidth      =   2655
                     TabIndex        =   118
                     Top             =   240
                     Width           =   2655
                     Begin VB.VScrollBar udDiceThrown 
                        Height          =   365
                        Index           =   1
                        Left            =   600
                        Max             =   1
                        Min             =   5
                        TabIndex        =   155
                        Top             =   600
                        Value           =   2
                        Width           =   255
                     End
                     Begin VB.VScrollBar udDiceThrown 
                        Height          =   365
                        Index           =   0
                        Left            =   600
                        Max             =   1
                        Min             =   5
                        TabIndex        =   154
                        Top             =   120
                        Value           =   3
                        Width           =   255
                     End
                     Begin VB.TextBox txtDiceThrown 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   1
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   132
                        Text            =   "2"
                        Top             =   600
                        Width           =   495
                     End
                     Begin VB.TextBox txtDiceThrown 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   0
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   119
                        Text            =   "3"
                        Top             =   120
                        Width           =   495
                     End
                     Begin VB.Label lblDiceThrown 
                        AutoSize        =   -1  'True
                        Caption         =   "Attack"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   0
                        Left            =   960
                        TabIndex        =   121
                        Top             =   120
                        Width           =   660
                     End
                     Begin VB.Label lblDiceThrown 
                        AutoSize        =   -1  'True
                        Caption         =   "Defence"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   1
                        Left            =   960
                        TabIndex        =   120
                        Top             =   600
                        Width           =   885
                     End
                  End
                  Begin VB.Label lblDiceThrow 
                     AutoSize        =   -1  'True
                     Caption         =   " Number thrown "
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
                     Height          =   270
                     Left            =   120
                     TabIndex        =   122
                     Top             =   0
                     Width           =   1680
                  End
               End
            End
            Begin VB.Frame frameSetup 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4455
               Index           =   1
               Left            =   6840
               TabIndex        =   95
               Top             =   5160
               Width           =   7335
               Begin VB.Frame fmrMissionList 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2295
                  Left            =   0
                  TabIndex        =   146
                  ToolTipText     =   "The War options box."
                  Top             =   2160
                  Width           =   7335
                  Begin VB.PictureBox pctMissionList 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1935
                     Left            =   120
                     ScaleHeight     =   1935
                     ScaleWidth      =   7155
                     TabIndex        =   147
                     Top             =   240
                     Width           =   7155
                     Begin VB.ListBox lstMissionList 
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   1740
                        ItemData        =   "Risk4.frx":88C3
                        Left            =   0
                        List            =   "Risk4.frx":88C5
                        TabIndex        =   149
                        Top             =   0
                        Width           =   7095
                     End
                  End
                  Begin VB.Label lblMissionList 
                     AutoSize        =   -1  'True
                     Caption         =   " Mission list "
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
                     TabIndex        =   148
                     Top             =   0
                     Width           =   1125
                  End
               End
               Begin VB.Frame frmMissionOptions 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2055
                  Left            =   0
                  TabIndex        =   141
                  ToolTipText     =   "The War options box."
                  Top             =   0
                  Width           =   4455
                  Begin VB.PictureBox pctMissionOptions 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1600
                     Left            =   120
                     ScaleHeight     =   1605
                     ScaleWidth      =   4275
                     TabIndex        =   142
                     Top             =   360
                     Width           =   4275
                     Begin VB.CheckBox chkMsnArmyWipeout 
                        Caption         =   "Army wipeout"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Left            =   0
                        TabIndex        =   200
                        Top             =   360
                        Value           =   1  'Checked
                        Width           =   4215
                     End
                     Begin VB.CheckBox chkMsnConquerHold 
                        Caption         =   "Conquer and hold"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Left            =   0
                        TabIndex        =   199
                        Top             =   600
                        Value           =   1  'Checked
                        Width           =   4215
                     End
                     Begin VB.CheckBox chkMsnMustComplete 
                        Caption         =   "Must complete your own mission"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Left            =   0
                        TabIndex        =   198
                        Top             =   840
                        Width           =   4215
                     End
                     Begin VB.CheckBox chkMsnWinImmediate 
                        Caption         =   "Win immediately on completion"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Left            =   0
                        TabIndex        =   197
                        Top             =   1080
                        Value           =   1  'Checked
                        Width           =   4215
                     End
                     Begin VB.CheckBox chkMsnAreUnique 
                        Caption         =   "Missions are unique"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Left            =   0
                        TabIndex        =   145
                        Top             =   1320
                        Value           =   1  'Checked
                        Width           =   4215
                     End
                     Begin VB.CheckBox chkMsnMissionsOn 
                        Caption         =   "Missions"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   240
                        Left            =   0
                        TabIndex        =   144
                        ToolTipText     =   "All players are issued with secret missions."
                        Top             =   0
                        Width           =   4215
                     End
                  End
                  Begin VB.Label lblMissionOptions 
                     AutoSize        =   -1  'True
                     Caption         =   " Mission options "
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
                     Height          =   270
                     Left            =   120
                     TabIndex        =   143
                     Top             =   0
                     Width           =   1755
                  End
               End
            End
            Begin VB.Frame frameSetup 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4455
               Index           =   3
               Left            =   0
               TabIndex        =   86
               Top             =   4680
               Width           =   7335
               Begin VB.Frame frmFixedValues 
                  BeginProperty Font 
                     Name            =   "MS Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2295
                  Left            =   4800
                  TabIndex        =   106
                  Top             =   0
                  Width           =   2535
                  Begin VB.PictureBox pctFixedCards 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1935
                     Left            =   120
                     ScaleHeight     =   1935
                     ScaleWidth      =   2295
                     TabIndex        =   107
                     Top             =   240
                     Width           =   2295
                     Begin VB.VScrollBar udFixedValues 
                        Height          =   365
                        Index           =   3
                        Left            =   600
                        Max             =   0
                        Min             =   50
                        TabIndex        =   184
                        Top             =   1560
                        Value           =   10
                        Width           =   255
                     End
                     Begin VB.VScrollBar udFixedValues 
                        Height          =   365
                        Index           =   2
                        Left            =   600
                        Max             =   0
                        Min             =   50
                        TabIndex        =   183
                        Top             =   1080
                        Value           =   8
                        Width           =   255
                     End
                     Begin VB.VScrollBar udFixedValues 
                        Height          =   365
                        Index           =   1
                        Left            =   600
                        Max             =   0
                        Min             =   50
                        TabIndex        =   182
                        Top             =   600
                        Value           =   6
                        Width           =   255
                     End
                     Begin VB.VScrollBar udFixedValues 
                        Height          =   365
                        Index           =   0
                        Left            =   600
                        Max             =   0
                        Min             =   50
                        TabIndex        =   181
                        Top             =   120
                        Value           =   4
                        Width           =   255
                     End
                     Begin VB.TextBox txtFixedValues 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   3
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   135
                        Text            =   "10"
                        Top             =   1560
                        Width           =   495
                     End
                     Begin VB.TextBox txtFixedValues 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   2
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   134
                        Text            =   "8"
                        Top             =   1080
                        Width           =   495
                     End
                     Begin VB.TextBox txtFixedValues 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   1
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   133
                        Text            =   "6"
                        Top             =   600
                        Width           =   495
                     End
                     Begin VB.TextBox txtFixedValues 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   0
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   108
                        Text            =   "4"
                        Top             =   120
                        Width           =   495
                     End
                     Begin VB.Label lblFixedValues 
                        AutoSize        =   -1  'True
                        Caption         =   "3 Artillery"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   0
                        Left            =   960
                        TabIndex        =   112
                        Top             =   120
                        Width           =   960
                     End
                     Begin VB.Label lblFixedValues 
                        AutoSize        =   -1  'True
                        Caption         =   "3 Infantry"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   1
                        Left            =   960
                        TabIndex        =   111
                        Top             =   600
                        Width           =   915
                     End
                     Begin VB.Label lblFixedValues 
                        AutoSize        =   -1  'True
                        Caption         =   "3 Cavalry"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   2
                        Left            =   960
                        TabIndex        =   110
                        Top             =   1080
                        Width           =   975
                     End
                     Begin VB.Label lblFixedValues 
                        AutoSize        =   -1  'True
                        Caption         =   "1 of Each"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   3
                        Left            =   960
                        TabIndex        =   109
                        Top             =   1560
                        Width           =   990
                     End
                  End
                  Begin VB.Label lbCardValues 
                     AutoSize        =   -1  'True
                     Caption         =   " Fixed card values "
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
                     Height          =   270
                     Left            =   120
                     TabIndex        =   113
                     Top             =   0
                     Width           =   1950
                  End
               End
               Begin VB.Frame frmTheDeck 
                  BeginProperty Font 
                     Name            =   "MS Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2295
                  Left            =   2160
                  TabIndex        =   97
                  Top             =   0
                  Width           =   2535
                  Begin VB.PictureBox pctTheDeck 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1935
                     Left            =   120
                     ScaleHeight     =   1935
                     ScaleWidth      =   2295
                     TabIndex        =   98
                     Top             =   240
                     Width           =   2295
                     Begin VB.VScrollBar udCardDeck 
                        Height          =   365
                        Index           =   3
                        Left            =   600
                        Max             =   0
                        Min             =   15
                        TabIndex        =   180
                        Top             =   1560
                        Value           =   2
                        Width           =   255
                     End
                     Begin VB.VScrollBar udCardDeck 
                        Height          =   365
                        Index           =   2
                        Left            =   600
                        Max             =   0
                        Min             =   15
                        TabIndex        =   179
                        Top             =   1080
                        Value           =   14
                        Width           =   255
                     End
                     Begin VB.VScrollBar udCardDeck 
                        Height          =   365
                        Index           =   1
                        Left            =   600
                        Max             =   0
                        Min             =   15
                        TabIndex        =   178
                        Top             =   600
                        Value           =   14
                        Width           =   255
                     End
                     Begin VB.VScrollBar udCardDeck 
                        Height          =   365
                        Index           =   0
                        Left            =   600
                        Max             =   0
                        Min             =   15
                        TabIndex        =   177
                        Top             =   120
                        Value           =   14
                        Width           =   255
                     End
                     Begin VB.TextBox txtCardDeck 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   0
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   131
                        Text            =   "14"
                        Top             =   120
                        Width           =   495
                     End
                     Begin VB.TextBox txtCardDeck 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   3
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   104
                        Text            =   "2"
                        Top             =   1560
                        Width           =   495
                     End
                     Begin VB.TextBox txtCardDeck 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   2
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   102
                        Text            =   "14"
                        Top             =   1080
                        Width           =   495
                     End
                     Begin VB.TextBox txtCardDeck 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   1
                        Left            =   120
                        Locked          =   -1  'True
                        TabIndex        =   100
                        Text            =   "14"
                        Top             =   600
                        Width           =   495
                     End
                     Begin VB.Label lblCardDeck 
                        AutoSize        =   -1  'True
                        Caption         =   "Wild"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   3
                        Left            =   960
                        TabIndex        =   105
                        Top             =   1560
                        Width           =   465
                     End
                     Begin VB.Label lblCardDeck 
                        AutoSize        =   -1  'True
                        Caption         =   "Cavalry"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   2
                        Left            =   960
                        TabIndex        =   103
                        Top             =   1080
                        Width           =   780
                     End
                     Begin VB.Label lblCardDeck 
                        AutoSize        =   -1  'True
                        Caption         =   "Infantry"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   1
                        Left            =   960
                        TabIndex        =   101
                        Top             =   600
                        Width           =   720
                     End
                     Begin VB.Label lblCardDeck 
                        AutoSize        =   -1  'True
                        Caption         =   "Artillery"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   12
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   270
                        Index           =   0
                        Left            =   960
                        TabIndex        =   99
                        Top             =   120
                        Width           =   765
                     End
                  End
                  Begin VB.Label lblTheDeck 
                     AutoSize        =   -1  'True
                     Caption         =   " The card deck "
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
                     Height          =   270
                     Left            =   240
                     TabIndex        =   114
                     Top             =   0
                     Width           =   1605
                  End
               End
               Begin VB.Frame frmSetupCards 
                  BeginProperty Font 
                     Name            =   "MS Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3135
                  Left            =   0
                  TabIndex        =   87
                  Top             =   0
                  Width           =   2055
                  Begin VB.PictureBox pctSetupCards 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   2775
                     Left            =   120
                     ScaleHeight     =   2775
                     ScaleWidth      =   1815
                     TabIndex        =   88
                     Top             =   240
                     Width           =   1815
                     Begin VB.VScrollBar udMaximumCardValue 
                        Height          =   365
                        LargeChange     =   5
                        Left            =   840
                        Max             =   1
                        Min             =   200
                        TabIndex        =   185
                        Top             =   1200
                        Value           =   20
                        Width           =   255
                     End
                     Begin VB.OptionButton optCardMode 
                        Caption         =   "Increasing"
                        Height          =   255
                        Index           =   2
                        Left            =   0
                        TabIndex        =   137
                        ToolTipText     =   "Cards are not issued."
                        Top             =   600
                        Width           =   1215
                     End
                     Begin VB.OptionButton optCardMode 
                        Caption         =   "Fixed"
                        Height          =   255
                        Index           =   1
                        Left            =   0
                        TabIndex        =   136
                        ToolTipText     =   "Cards are not issued."
                        Top             =   360
                        Value           =   -1  'True
                        Width           =   1215
                     End
                     Begin VB.OptionButton optCardMode 
                        Caption         =   "None"
                        Height          =   255
                        Index           =   0
                        Left            =   0
                        TabIndex        =   92
                        ToolTipText     =   "Cards are not issued."
                        Top             =   120
                        Width           =   1215
                     End
                     Begin VB.CheckBox chkCardsVulture 
                        Caption         =   "Capture"
                        Height          =   375
                        Left            =   0
                        TabIndex        =   91
                        ToolTipText     =   "Capture players' cards when you wipe out their last unit."
                        Top             =   2280
                        Width           =   1215
                     End
                     Begin VB.CheckBox chkCardsHidden 
                        Caption         =   "Hidden"
                        Height          =   255
                        Left            =   0
                        TabIndex        =   90
                        Top             =   2040
                        Value           =   1  'Checked
                        Width           =   1215
                     End
                     Begin VB.TextBox txtMaximumCardValue 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Left            =   240
                        Locked          =   -1  'True
                        TabIndex        =   89
                        Text            =   "20"
                        Top             =   1200
                        Width           =   615
                     End
                     Begin VB.Label lblSetupMaxCardValue 
                        AutoSize        =   -1  'True
                        Caption         =   "Maximum value"
                        ForeColor       =   &H00FF0000&
                        Height          =   240
                        Left            =   240
                        TabIndex        =   93
                        Top             =   960
                        Width           =   1350
                     End
                  End
                  Begin VB.Label lblSetupCards 
                     AutoSize        =   -1  'True
                     Caption         =   " Cards "
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
                     Height          =   270
                     Left            =   120
                     TabIndex        =   94
                     Top             =   0
                     Width           =   765
                  End
               End
            End
            Begin VB.Frame frameSetup 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4455
               Index           =   0
               Left            =   240
               TabIndex        =   31
               Top             =   600
               Width           =   7335
               Begin VB.Frame fSetupWarOptions 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1335
                  Left            =   4680
                  TabIndex        =   80
                  ToolTipText     =   "The War options box."
                  Top             =   3120
                  Width           =   2535
                  Begin VB.PictureBox pctSetupWarOptions 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Left            =   120
                     ScaleHeight     =   735
                     ScaleWidth      =   2235
                     TabIndex        =   81
                     Top             =   360
                     Width           =   2235
                     Begin VB.OptionButton optSupplyLines 
                        Caption         =   "Unlimited"
                        Height          =   240
                        Left            =   0
                        TabIndex        =   84
                        ToolTipText     =   "No restrictions on army movements."
                        Top             =   0
                        Width           =   3135
                     End
                     Begin VB.OptionButton optLimitSupply 
                        Caption         =   "Limited"
                        Height          =   240
                        Left            =   0
                        TabIndex        =   83
                        ToolTipText     =   "Limit army movements."
                        Top             =   240
                        Value           =   -1  'True
                        Width           =   3135
                     End
                     Begin VB.OptionButton optNoSupply 
                        Caption         =   "No supply lines"
                        Height          =   240
                        Left            =   0
                        TabIndex        =   82
                        ToolTipText     =   "Restrict army movements to 1 adjacent country only."
                        Top             =   480
                        Width           =   3135
                     End
                  End
                  Begin VB.Label lblSetupBattleOptions 
                     AutoSize        =   -1  'True
                     Caption         =   " Supply lines "
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
                     Height          =   270
                     Left            =   120
                     TabIndex        =   85
                     Top             =   0
                     Width           =   1365
                  End
               End
               Begin VB.Frame fSetupPlayerNumber 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2895
                  Left            =   4680
                  TabIndex        =   69
                  ToolTipText     =   "Change the number of starting players."
                  Top             =   0
                  Width           =   2535
                  Begin VB.PictureBox pctSetupFirstPlayer 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   2415
                     Left            =   120
                     ScaleHeight     =   2415
                     ScaleWidth      =   2355
                     TabIndex        =   70
                     Top             =   360
                     Width           =   2355
                     Begin VB.VScrollBar udStartingArmies 
                        Height          =   365
                        Left            =   600
                        Max             =   2
                        Min             =   6
                        TabIndex        =   77
                        Top             =   240
                        Value           =   6
                        Width           =   255
                     End
                     Begin VB.TextBox txtStartingArmies 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Left            =   240
                        Locked          =   -1  'True
                        TabIndex        =   76
                        Text            =   "6"
                        Top             =   240
                        Width           =   375
                     End
                     Begin VB.VScrollBar udExtraStartingUnits 
                        Height          =   365
                        LargeChange     =   5
                        Left            =   720
                        Max             =   1
                        Min             =   200
                        TabIndex        =   75
                        Top             =   1080
                        Value           =   20
                        Width           =   255
                     End
                     Begin VB.TextBox txtExtraStartingUnits 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Left            =   240
                        Locked          =   -1  'True
                        TabIndex        =   74
                        Text            =   "20"
                        Top             =   1080
                        Width           =   495
                     End
                     Begin VB.CheckBox chkExtraStartingUnits 
                        Caption         =   "Extra starting battalions"
                        Height          =   255
                        Left            =   0
                        TabIndex        =   73
                        ToolTipText     =   "Extra starting units are randomly distributed over each players' territories."
                        Top             =   840
                        Width           =   2415
                     End
                     Begin VB.OptionButton optRandomFirstPlayer 
                        Caption         =   "Random"
                        Height          =   255
                        Left            =   0
                        TabIndex        =   72
                        ToolTipText     =   "Random selection of the first player."
                        Top             =   1920
                        Value           =   -1  'True
                        Width           =   2295
                     End
                     Begin VB.OptionButton optPlr1FirstPlayer 
                        Caption         =   "The Red Army"
                        Height          =   255
                        Left            =   0
                        TabIndex        =   71
                        ToolTipText     =   "Player 1 - the Red Army is the first player."
                        Top             =   2160
                        Width           =   2175
                     End
                     Begin VB.Label lblStartingArmies 
                        AutoSize        =   -1  'True
                        Caption         =   " Starting Players "
                        Height          =   240
                        Left            =   0
                        TabIndex        =   79
                        Top             =   0
                        Width           =   1515
                     End
                     Begin VB.Label lblSetupFirst 
                        AutoSize        =   -1  'True
                        Caption         =   " First battalion "
                        ForeColor       =   &H00FF0000&
                        Height          =   240
                        Left            =   0
                        TabIndex        =   78
                        Top             =   1680
                        Width           =   1305
                     End
                  End
                  Begin VB.Label lblSetupStart 
                     AutoSize        =   -1  'True
                     Caption         =   " Start options "
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
                     Height          =   270
                     Left            =   120
                     TabIndex        =   115
                     Top             =   0
                     Width           =   1440
                  End
               End
               Begin VB.Frame plrOpt 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000F&
                  Height          =   4455
                  Index           =   0
                  Left            =   120
                  TabIndex        =   32
                  ToolTipText     =   "Player option box for the Red Army."
                  Top             =   0
                  Width           =   4455
                  Begin VB.PictureBox pctSetupPlayerContainer5 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Left            =   120
                     ScaleHeight     =   375
                     ScaleWidth      =   4215
                     TabIndex        =   63
                     Top             =   3960
                     Width           =   4215
                     Begin VB.VScrollBar vscrollPlayerSelect 
                        Height          =   365
                        Index           =   5
                        Left            =   3960
                        Max             =   0
                        Min             =   3
                        TabIndex        =   68
                        Top             =   0
                        Value           =   3
                        Width           =   255
                     End
                     Begin VB.VScrollBar udPlayerStartCountries 
                        Height          =   365
                        Index           =   5
                        Left            =   1080
                        Max             =   0
                        Min             =   41
                        TabIndex        =   67
                        Top             =   0
                        Value           =   7
                        Width           =   255
                     End
                     Begin VB.TextBox txtPlayerStartCountries 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   5
                        Left            =   600
                        Locked          =   -1  'True
                        TabIndex        =   66
                        Text            =   "7"
                        Top             =   0
                        Width           =   495
                     End
                     Begin VB.ComboBox PlayerSelect 
                        Height          =   360
                        Index           =   5
                        ItemData        =   "Risk4.frx":88C7
                        Left            =   1440
                        List            =   "Risk4.frx":88D4
                        Locked          =   -1  'True
                        Style           =   1  'Simple Combo
                        TabIndex        =   65
                        TabStop         =   0   'False
                        Text            =   "PlayerSelect"
                        Top             =   20
                        Width           =   2535
                     End
                     Begin VB.PictureBox pctClr 
                        AutoRedraw      =   -1  'True
                        BackColor       =   &H00C0C0C0&
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   13.5
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   420
                        Index           =   5
                        Left            =   0
                        ScaleHeight     =   360
                        ScaleWidth      =   435
                        TabIndex        =   64
                        TabStop         =   0   'False
                        ToolTipText     =   "Player option box for the Gray Army."
                        Top             =   0
                        Width           =   495
                     End
                  End
                  Begin VB.PictureBox pctSetupPlayerContainer4 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   495
                     Left            =   120
                     ScaleHeight     =   495
                     ScaleWidth      =   4215
                     TabIndex        =   57
                     Top             =   3240
                     Width           =   4215
                     Begin VB.VScrollBar vscrollPlayerSelect 
                        Height          =   365
                        Index           =   4
                        Left            =   3960
                        Max             =   0
                        Min             =   3
                        TabIndex        =   62
                        Top             =   0
                        Value           =   3
                        Width           =   255
                     End
                     Begin VB.VScrollBar udPlayerStartCountries 
                        Height          =   365
                        Index           =   4
                        Left            =   1080
                        Max             =   0
                        Min             =   41
                        TabIndex        =   61
                        Top             =   0
                        Value           =   7
                        Width           =   255
                     End
                     Begin VB.TextBox txtPlayerStartCountries 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   4
                        Left            =   600
                        Locked          =   -1  'True
                        TabIndex        =   60
                        Text            =   "7"
                        Top             =   0
                        Width           =   495
                     End
                     Begin VB.ComboBox PlayerSelect 
                        Height          =   360
                        Index           =   4
                        ItemData        =   "Risk4.frx":8916
                        Left            =   1440
                        List            =   "Risk4.frx":8923
                        Locked          =   -1  'True
                        Style           =   1  'Simple Combo
                        TabIndex        =   59
                        TabStop         =   0   'False
                        Text            =   "PlayerSelect"
                        Top             =   20
                        Width           =   2535
                     End
                     Begin VB.PictureBox pctClr 
                        AutoRedraw      =   -1  'True
                        BackColor       =   &H00FF00FF&
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   13.5
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   420
                        Index           =   4
                        Left            =   0
                        ScaleHeight     =   360
                        ScaleWidth      =   435
                        TabIndex        =   58
                        TabStop         =   0   'False
                        ToolTipText     =   "Player option box for the Purple Army."
                        Top             =   0
                        Width           =   495
                     End
                  End
                  Begin VB.PictureBox pctSetupPlayerContainer3 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   495
                     Left            =   120
                     ScaleHeight     =   495
                     ScaleWidth      =   4215
                     TabIndex        =   51
                     Top             =   2520
                     Width           =   4215
                     Begin VB.VScrollBar vscrollPlayerSelect 
                        Height          =   365
                        Index           =   3
                        Left            =   3960
                        Max             =   0
                        Min             =   3
                        TabIndex        =   56
                        Top             =   0
                        Value           =   3
                        Width           =   255
                     End
                     Begin VB.VScrollBar udPlayerStartCountries 
                        Height          =   365
                        Index           =   3
                        Left            =   1080
                        Max             =   0
                        Min             =   41
                        TabIndex        =   55
                        Top             =   0
                        Value           =   7
                        Width           =   255
                     End
                     Begin VB.TextBox txtPlayerStartCountries 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   3
                        Left            =   600
                        Locked          =   -1  'True
                        TabIndex        =   54
                        Text            =   "7"
                        Top             =   0
                        Width           =   495
                     End
                     Begin VB.ComboBox PlayerSelect 
                        Height          =   360
                        Index           =   3
                        ItemData        =   "Risk4.frx":8965
                        Left            =   1440
                        List            =   "Risk4.frx":8972
                        Locked          =   -1  'True
                        Style           =   1  'Simple Combo
                        TabIndex        =   53
                        TabStop         =   0   'False
                        Text            =   "PlayerSelect"
                        Top             =   20
                        Width           =   2535
                     End
                     Begin VB.PictureBox pctClr 
                        AutoRedraw      =   -1  'True
                        BackColor       =   &H0000FFFF&
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   13.5
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   420
                        Index           =   3
                        Left            =   0
                        ScaleHeight     =   360
                        ScaleWidth      =   435
                        TabIndex        =   52
                        TabStop         =   0   'False
                        ToolTipText     =   "Player option box for the Yellow Army."
                        Top             =   0
                        Width           =   495
                     End
                  End
                  Begin VB.PictureBox pctSetupPlayerContainer2 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   495
                     Left            =   120
                     ScaleHeight     =   495
                     ScaleWidth      =   4215
                     TabIndex        =   45
                     Top             =   1800
                     Width           =   4215
                     Begin VB.VScrollBar vscrollPlayerSelect 
                        Height          =   365
                        Index           =   2
                        Left            =   3960
                        Max             =   0
                        Min             =   3
                        TabIndex        =   50
                        Top             =   0
                        Value           =   3
                        Width           =   255
                     End
                     Begin VB.VScrollBar udPlayerStartCountries 
                        Height          =   365
                        Index           =   2
                        Left            =   1080
                        Max             =   0
                        Min             =   41
                        TabIndex        =   49
                        Top             =   0
                        Value           =   7
                        Width           =   255
                     End
                     Begin VB.TextBox txtPlayerStartCountries 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   2
                        Left            =   600
                        Locked          =   -1  'True
                        TabIndex        =   48
                        Text            =   "7"
                        Top             =   0
                        Width           =   495
                     End
                     Begin VB.ComboBox PlayerSelect 
                        Height          =   360
                        Index           =   2
                        ItemData        =   "Risk4.frx":89B4
                        Left            =   1440
                        List            =   "Risk4.frx":89C1
                        Locked          =   -1  'True
                        Style           =   1  'Simple Combo
                        TabIndex        =   47
                        TabStop         =   0   'False
                        Text            =   "PlayerSelect"
                        Top             =   20
                        Width           =   2535
                     End
                     Begin VB.PictureBox pctClr 
                        AutoRedraw      =   -1  'True
                        BackColor       =   &H00FFFF00&
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   13.5
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   420
                        Index           =   2
                        Left            =   0
                        ScaleHeight     =   360
                        ScaleWidth      =   435
                        TabIndex        =   46
                        TabStop         =   0   'False
                        ToolTipText     =   "Player option box for the Blue Army."
                        Top             =   0
                        Width           =   495
                     End
                  End
                  Begin VB.PictureBox pctSetupPlayerContainer1 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   495
                     Left            =   120
                     ScaleHeight     =   495
                     ScaleWidth      =   4215
                     TabIndex        =   39
                     Top             =   1080
                     Width           =   4215
                     Begin VB.VScrollBar udPlayerStartCountries 
                        Height          =   365
                        Index           =   1
                        Left            =   1080
                        Max             =   0
                        Min             =   41
                        TabIndex        =   44
                        Top             =   0
                        Value           =   7
                        Width           =   255
                     End
                     Begin VB.TextBox txtPlayerStartCountries 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   1
                        Left            =   600
                        Locked          =   -1  'True
                        TabIndex        =   43
                        Text            =   "7"
                        Top             =   0
                        Width           =   495
                     End
                     Begin VB.VScrollBar vscrollPlayerSelect 
                        Height          =   365
                        Index           =   1
                        Left            =   3960
                        Max             =   0
                        Min             =   3
                        TabIndex        =   42
                        Top             =   0
                        Value           =   3
                        Width           =   255
                     End
                     Begin VB.ComboBox PlayerSelect 
                        Height          =   360
                        Index           =   1
                        ItemData        =   "Risk4.frx":8A03
                        Left            =   1440
                        List            =   "Risk4.frx":8A10
                        Locked          =   -1  'True
                        Style           =   1  'Simple Combo
                        TabIndex        =   41
                        TabStop         =   0   'False
                        Text            =   "PlayerSelect"
                        Top             =   20
                        Width           =   2535
                     End
                     Begin VB.PictureBox pctClr 
                        AutoRedraw      =   -1  'True
                        BackColor       =   &H0000FF00&
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   13.5
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   420
                        Index           =   1
                        Left            =   0
                        ScaleHeight     =   360
                        ScaleWidth      =   435
                        TabIndex        =   40
                        TabStop         =   0   'False
                        ToolTipText     =   "Player option box for the Green Army."
                        Top             =   0
                        Width           =   495
                     End
                  End
                  Begin VB.PictureBox pctSetupPlayerContainer0 
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   495
                     Left            =   120
                     ScaleHeight     =   495
                     ScaleWidth      =   4215
                     TabIndex        =   33
                     Top             =   360
                     Width           =   4215
                     Begin VB.VScrollBar udPlayerStartCountries 
                        Height          =   365
                        Index           =   0
                        Left            =   1080
                        Max             =   0
                        Min             =   41
                        TabIndex        =   38
                        Top             =   0
                        Value           =   7
                        Width           =   255
                     End
                     Begin VB.TextBox txtPlayerStartCountries 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Index           =   0
                        Left            =   600
                        Locked          =   -1  'True
                        TabIndex        =   37
                        Text            =   "7"
                        Top             =   0
                        Width           =   495
                     End
                     Begin VB.VScrollBar vscrollPlayerSelect 
                        Height          =   365
                        Index           =   0
                        Left            =   3960
                        Max             =   0
                        Min             =   3
                        TabIndex        =   36
                        Top             =   0
                        Value           =   3
                        Width           =   255
                     End
                     Begin VB.ComboBox PlayerSelect 
                        Height          =   360
                        Index           =   0
                        ItemData        =   "Risk4.frx":8A52
                        Left            =   1440
                        List            =   "Risk4.frx":8A5F
                        Locked          =   -1  'True
                        Style           =   1  'Simple Combo
                        TabIndex        =   35
                        TabStop         =   0   'False
                        Text            =   "PlayerSelect"
                        Top             =   20
                        Width           =   2535
                     End
                     Begin VB.PictureBox pctClr 
                        AutoRedraw      =   -1  'True
                        BackColor       =   &H000000FF&
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   13.5
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   420
                        Index           =   0
                        Left            =   0
                        ScaleHeight     =   360
                        ScaleWidth      =   435
                        TabIndex        =   34
                        TabStop         =   0   'False
                        ToolTipText     =   "Player option box for the Red Army."
                        Top             =   0
                        Width           =   495
                     End
                  End
                  Begin VB.Label lblPlayerOptions 
                     AutoSize        =   -1  'True
                     Caption         =   " Allocator "
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
                     Height          =   270
                     Left            =   120
                     TabIndex        =   116
                     Top             =   0
                     Width           =   1035
                  End
               End
            End
            Begin VB.PictureBox pctSetopDeclareCancel 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   7680
               ScaleHeight     =   1095
               ScaleWidth      =   1215
               TabIndex        =   27
               Top             =   4080
               Width           =   1215
               Begin VB.CommandButton cmdSUPcncl 
                  Cancel          =   -1  'True
                  Caption         =   "&Cancel Set Up"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   0
                  MaskColor       =   &H8000000F&
                  TabIndex        =   29
                  ToolTipText     =   "Exit setup."
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.CommandButton cmdSetupOk 
                  Caption         =   "Declare &War"
                  Default         =   -1  'True
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   0
                  MaskColor       =   &H8000000F&
                  TabIndex        =   28
                  ToolTipText     =   "Start new game with these settings."
                  Top             =   0
                  Width           =   1215
               End
            End
            Begin MSComctlLib.TabStrip tabSetup 
               Height          =   4935
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   7455
               _ExtentX        =   13150
               _ExtentY        =   8705
               MultiRow        =   -1  'True
               _Version        =   393216
               BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
                  NumTabs         =   5
                  BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Army"
                     Key             =   "Army"
                     ImageVarType    =   2
                  EndProperty
                  BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Missions"
                     Key             =   "Missions"
                     ImageVarType    =   2
                  EndProperty
                  BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Dice"
                     Key             =   "Dice"
                     ImageVarType    =   2
                  EndProperty
                  BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Cards"
                     Key             =   "Cards"
                     ImageVarType    =   2
                  EndProperty
                  BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Recruits"
                     Key             =   "Recruits"
                     ImageVarType    =   2
                  EndProperty
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label lblVersion 
               AutoSize        =   -1  'True
               Caption         =   "Version x.x.x"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000011&
               Height          =   210
               Left            =   120
               TabIndex        =   196
               Top             =   5160
               Width           =   975
            End
         End
      End
      Begin VB.PictureBox pctInfoBox 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1785
         Left            =   120
         ScaleHeight     =   115
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   189
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   6960
         Width           =   2895
      End
      Begin VB.CommandButton cmdExchange 
         BackColor       =   &H00BB6302&
         Caption         =   "Exchange"
         Height          =   375
         Left            =   4440
         MaskColor       =   &H8000000F&
         TabIndex        =   13
         Top             =   7560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCardCncl 
         BackColor       =   &H00BB6302&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4440
         MaskColor       =   &H8000000F&
         TabIndex        =   12
         Top             =   7920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAttack 
         BackColor       =   &H00BB6302&
         Caption         =   "&Attack"
         Height          =   495
         Left            =   720
         MaskColor       =   &H8000000F&
         TabIndex        =   0
         Top             =   5400
         Width           =   825
      End
      Begin VB.CommandButton cmdMove 
         BackColor       =   &H00BB6302&
         Caption         =   "&Move"
         Height          =   495
         Left            =   720
         MaskColor       =   &H8000000F&
         TabIndex        =   1
         Top             =   5880
         Width           =   825
      End
      Begin VB.CommandButton cmdEnd 
         BackColor       =   &H00BB6302&
         Caption         =   "&Pass"
         Height          =   495
         Left            =   720
         MaskColor       =   &H8000000F&
         TabIndex        =   2
         Top             =   6360
         Width           =   825
      End
      Begin VB.PictureBox pctTransfer 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00BB6302&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   615
         TabIndex        =   3
         Top             =   5160
         Width           =   615
         Begin VB.OptionButton tfRateAll 
            BackColor       =   &H00BB6302&
            Caption         =   "&999"
            Height          =   195
            Left            =   0
            MaskColor       =   &H8000000F&
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1440
            Width           =   615
         End
         Begin VB.OptionButton tfRate50 
            BackColor       =   &H00BB6302&
            Caption         =   "50"
            Height          =   225
            Left            =   0
            MaskColor       =   &H8000000F&
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1200
            Width           =   615
         End
         Begin VB.OptionButton tfRate20 
            BackColor       =   &H00BB6302&
            Caption         =   "20"
            Height          =   225
            Left            =   0
            MaskColor       =   &H8000000F&
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   960
            Width           =   615
         End
         Begin VB.OptionButton tfRate10 
            BackColor       =   &H00BB6302&
            Caption         =   "10"
            Height          =   225
            Left            =   0
            MaskColor       =   &H8000000F&
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton tfRate5 
            BackColor       =   &H00BB6302&
            Caption         =   " &5"
            Height          =   225
            Left            =   0
            MaskColor       =   &H8000000F&
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton tfRate2 
            BackColor       =   &H00BB6302&
            Caption         =   " &2"
            Height          =   225
            Left            =   0
            MaskColor       =   &H8000000F&
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton tfRate1 
            BackColor       =   &H00BB6302&
            Caption         =   " &1"
            Height          =   225
            Left            =   0
            MaskColor       =   &H8000000F&
            TabIndex        =   4
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   0
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   1320
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":8AA1
            Key             =   ""
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":8DF3
            Key             =   ""
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":9145
            Key             =   ""
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":9497
            Key             =   ""
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":97E9
            Key             =   ""
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":9B3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":9E8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":A1DF
            Key             =   ""
            Object.Tag             =   "Fast"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":A531
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":A883
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":A995
            Key             =   ""
            Object.Tag             =   "Net"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":AAA7
            Key             =   ""
            Object.Tag             =   "Reset"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":AE6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":B333
            Key             =   ""
            Object.Tag             =   "Chat"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":B785
            Key             =   ""
            Object.Tag             =   "Conts"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":BBFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":C051
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":C4A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":C8F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":D247
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":D8C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Risk4.frx":DFD3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New war..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileReset 
         Caption         =   "&Reset war"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuAutoRestart 
         Caption         =   "&Automatically restart war"
      End
      Begin VB.Menu mnuViewStep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLoadWar 
         Caption         =   "&Open war..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveWar 
         Caption         =   "&Save war"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileWarAs 
         Caption         =   "Save war as..."
      End
      Begin VB.Menu mnuFileStep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptUndo 
         Caption         =   "Start turn again"
      End
      Begin VB.Menu mnuFileBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuNet 
      Caption         =   "M&ultiplayer"
      Begin VB.Menu mnuNetHost 
         Caption         =   "Host Multiplayer War"
         Begin VB.Menu mnuHostQueue 
            Caption         =   "Queue Cards..."
         End
         Begin VB.Menu mnuNetHostLan 
            Caption         =   "Host LAN War..."
         End
         Begin VB.Menu mnuNetHostInet 
            Caption         =   "Host Internet War..."
         End
      End
      Begin VB.Menu mnuNetClient 
         Caption         =   "Find Multiplayer War"
         Begin VB.Menu mnuClientQue 
            Caption         =   "Queue Cards..."
         End
         Begin VB.Menu mnuClientLan 
            Caption         =   "Find LAN War..."
         End
         Begin VB.Menu mnuNetClientInternet 
            Caption         =   "Find Internet War..."
            Shortcut        =   ^I
         End
      End
      Begin VB.Menu mnuNetAdvanced 
         Caption         =   "Admin Panel..."
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuNetBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNetDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuNetBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNetCntr 
         Caption         =   "View/Activate Counter"
      End
      Begin VB.Menu mnuNetChat 
         Caption         =   "Compose Message..."
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu optnFastWar 
         Caption         =   "Fast &war"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuOptStep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullScreen 
         Caption         =   "Full Screen"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuViewQualityDisplay 
         Caption         =   "Smooth display"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu3Ddisplay 
         Caption         =   "3D map"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBorder 
         Caption         =   "Border around map"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFlashInfoBox 
         Caption         =   "Flashing Infobox"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptStep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLFont 
         Caption         =   "Font..."
      End
      Begin VB.Menu mnuOptToolbox 
         Caption         =   "&Tool bar"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuOptReport 
         Caption         =   "&Reshuffle notice"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptStats 
         Caption         =   "&End of war stats"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMisSeeReminder 
         Caption         =   "&Mission reminder"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptLanguage 
         Caption         =   "Language..."
      End
   End
   Begin VB.Menu mnuMission 
      Caption         =   "Mi&ssion"
      Begin VB.Menu mnuMissionSee 
         Caption         =   "&See mission"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Help..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu hlpMRhome 
         Caption         =   "<Var.ExeName> Home &Page..."
      End
      Begin VB.Menu hlpContMap 
         Caption         =   "&Continent map"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuHlpStep2 
         Caption         =   "-"
      End
      Begin VB.Menu hlpCheckForUpdates 
         Caption         =   "Check for updates"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuHlpStep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuMainBrk1 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuMainBrk2 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuAttack 
      Caption         =   "&Attack"
   End
   Begin VB.Menu mnuMove 
      Caption         =   "&Move"
   End
   Begin VB.Menu mnuPass 
      Caption         =   "&Pass"
   End
   Begin VB.Menu mnuDeclareWar 
      Caption         =   "&Declare War"
   End
   Begin VB.Menu mnuCancelSetup 
      Caption         =   "Cancel Setup"
   End
End
Attribute VB_Name = "TheMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants
Const remoteIndex As Long = 3           '"Remote player" list index
Const A3Index As Long = 4
Const playFast As Integer = 12          'Fast clock speed
Const playSlow As Integer = 100         'Slow clock speed
Const diceFast As Integer = 10           'Fast dice
Const diceSlow As Integer = 300          'Slow dice speed

Public gCancelUnload As Integer         'Used to cancel form unload when unloading other forms.

'Graphics variables
Public gPictureMaskRatioX As Double
Public gPictureMaskRatioY As Double
Public gSyncViewportNeeded As Boolean      'Set to true if the map has changed and the viewport needs resynching.
Public gDeviceCaps12 As Long            'Global device caps for Windows colour depth.
Dim gLockStartPlayerCount As Boolean    'Stop cascading changes when players are set to have 0 countries.
Dim gCtryTextColor(49) As Long      'Color of text on that country.
Dim gCtryCurrentColor(49) As Long   'Current colour of each country.
Dim ctryIsBlack As Integer              'This country is currently under attack and is black. Set during blitting.
Public gLittleCardFontSize As Single        'Font size to use with little card text.
Dim gBorderWidth As Long                'Divide the form width by this number to get the border width. Default 100.
Dim gPlayerFlashColor(5) As Long      'Flash this colour in the info box when player's turn.
Public gCurrentMousePosX As Integer
Public gCurrentMousePosY As Integer
Dim CountryUnderAttack As Integer       'This country is under attack
Dim EasyAttackConts(6) As Boolean       'Easy attack marker.
Dim IsEasyAttackOn As Boolean           'Commit all units if True.
Dim clickLock As Boolean                'Lock map click
Dim gPreviousSettings As PreviousSettingsType        'Remember setup settings incase of cancel
Dim warBak As WarControlType                'Recover from error
Public transferNmbr As Integer             'Number transfered at a time
Public nmbrOfPlayers As Integer            'Number of players
Dim gCtryOrder(49) As Integer       'Order countries were placed
Dim timerCounter As Integer
Dim retreatArmies As Integer
Public gPlayerValue As Integer              'Player gets at start of go
Dim notHitMove As Boolean               'False if touched key
Dim gMoveLimit As Integer                'Moves limited mode
Dim gMoveTimes(42) As Integer       'Record of moves made
Dim gMovedIn(42) As Integer         'Record of armies moved
Public AutoCountry As Integer              'Computer selects this country
Public playSpeed As Integer                'Timer speed
Public gComputerPressed As Boolean          'Computer pressed button
Dim flashingBorder As Boolean           'Border shows player color if true
Public diceSpeed As Integer                'Speed of dice
Dim tmpTimer1 As Boolean                'Temp storage for timer values
Dim tmpTimer2 As Boolean
'Dim optimizeDice As Boolean             'Throw gDiceArray(5) if gDiceArray(2) > 4
Dim boolDrawnWin As Boolean             'Only draw "VICTORY" once, True if drawn
Public boolIssueCard As Boolean            'Player gets a card at end of go.
Dim AutoSecondMove As Boolean           'Another move after this one
Dim nextCountryMove As Integer          'Move to 2 different places
Dim prefercont As Integer               'Attack this cont
Public sWinMessage As String               'Message to print to info screen when won
Public gComputerAquiredCards As Boolean             'Tell computer players about cards
Dim cardstmp As Integer                 'Temp storage for setup cancle
Dim AtkForCard As Attack                'Attack and retreat details
Dim L As Integer                        'Language, 0=English
Dim modeBeforeExchange As Integer       'gCurrentMode before cards were exchanged
Dim playerDefence(6) As Single
Dim Plr(6) As Long
Dim CPUspeedTimer As Boolean            'Used to test cpu speed
Public gMapSetupLock As Boolean             'Stop interferance of map setup
Public gWarRestartLock As Boolean              'Prevent DrawWin being interupted
Dim gNewWarPlayerTurn As Integer        'Remember playerturn for cancel.
Dim Audit(6) As AuditPlayerType         'Auditing to prevent remote terminal cheating.
Dim AuditCardDeck As Long               'Remember value of the card deck.
Dim AuditShadow As String               'Shadow file for auditing system. Used to guard against fiddling.
Dim gCurrentWarPath As String           'File path and name of the current war.

'Old bubble sort. Still used for cards.
'TODO - fix cards.
Public Sub SortDice(firstDice As Integer, lastDice As Integer)

    Dim hold As Integer
    Dim pass As Integer, cntr2 As Integer

    For pass = firstDice To lastDice
        For cntr2 = firstDice To lastDice - 1
            If gDiceArray(cntr2 - 1) < gDiceArray(cntr2) Then
                hold = gDiceArray(cntr2)
                gDiceArray(cntr2) = gDiceArray(cntr2 - 1)
                gDiceArray(cntr2 - 1) = hold
                cntr2 = firstDice
            End If
        Next cntr2
    Next pass
End Sub

'Update debug viewer for testing purposes.
'Set test points at entry and exit of sus procedures.
'Log to file "MRLog_<date stamp>.log"
Public Sub updateTestViewer(checkPoint As String)
    frmTst.Visible = gCheatMode.testing
    txtTst(0).Text = gCurrentMode
    txtTst(1).Text = gPlayerTurn
    txtTst(2).Text = Timer1.Enabled
    txtTst(3).Text = Timer2.Enabled
    If Trim(checkPoint) <> "" Then
        txtTst(4).Text = checkPoint
    End If
    If gCheatMode.testing Then
        'Call Module1.pause(200, True)
        Call AppendToMRLogs(checkPoint)
    End If
    DoEvents
End Sub

'Append to logs directory. If directory doesn't exist
'then create it in GetLogDataDir.
'Log to file "MRLog_<date stamp>.log"
Public Sub AppendToMRLogs(pRecord As String)
    Dim filename As String
    Dim vFileNo As Integer
    
    On Error Resume Next
    If gCheatMode.testing Then
        vFileNo = FreeFile
        Open GetLogDataDir & "\MRLog_" & Format(Date, "yyyymmdd") & ".log" For Append As vFileNo
        Print #vFileNo, Format(Now, "hh:mm:ss") & " - " & pRecord
        Close vFileNo
    End If
End Sub

'Record total player unit tally and card value for auditing.
'Remember card deck value as well.
'vAuditShadow tracks and verifies the auditing system by recording changes as they happen.
'vAuditShadow format:
'"TermNumber,Audit(1).Player,Audit(1).PlayerCard,Audit(2).Player,Audit(2).PlayerCard,
'...,Audit(6).Player,Audit(6).PlayerCard,Player Number,[u|d]Score Change,Player Number,
'[u|d]Score Change,...,Player Number,[u|d]Score Change,
'[f]Final Score Player 1,Final Cards PLayer 1,[f]Final Score Player 2,Final Cards PLayer 1,...
'Where "u" means score increase, "d" means score decrease, "c" means card added, "r" means card removed.
'An example of a complete shadow record ready to send to the host. Form terminal 1, The Green Army
'traded cards (value 6 = 10 points: art,inf,cav) then attacked the red army and won, recieving a card
'at the end of the turn.
'1,18,3,19,8,17,5,15,3,19,4,2,8,2,u5,2,u10,2,r6,1,d1,2,c3,17,f3,34,f5,17,f5,15,f3,19,f4,2,f8,
Private Sub AuditPlayerRecord()
    Dim cntr As Long
    Dim cntr2 As Long
    Dim vAuditShadow As String
    Dim vTest As String
    ''Call updateTestViewer("AuditPlayerRecord")
    
    If netWorkSituation = cNetClient And net.playerOwner(gPlayerTurn - 1) = myTerminalNumber Then  'cmdEnd.Tag = "OK" Then
    vAuditShadow = CStr(myTerminalNumber) & ","
    
    'Remember player
    For cntr = 1 To 6
        Audit(cntr).Player = CountPlayerUnits(CInt(cntr))
        Audit(cntr).PlayerMax = Audit(cntr).Player
        'Debug.Print playerID(cntr).strColor & " " & Audit(cntr).PlayerMax
        vTest = vTest & "Plr " & CStr(cntr) & " p" & CStr(Audit(cntr).Player) & " m" & CStr(Audit(cntr).PlayerMax) & " "
        
        Audit(cntr).PlayerCard = 0
        For cntr2 = 0 To 10
            Audit(cntr).PlayerCard = Audit(cntr).PlayerCard + gPlayerID(cntr).card(cntr2)
        Next
        vAuditShadow = vAuditShadow & CStr(Audit(cntr).Player) & "," & Audit(cntr).PlayerCard & ","
    Next
    LogInfo "AuditPlayerRecord", vTest, 5, True
    AuditShadow = gGsLeUtils.LE5(vAuditShadow)
    
    'Remember card deck value.
    AuditCardDeck = AuditGetDeckValue
    End If
End Sub

'Append the end of turn scores before sending to the host as part of the refresh.
Public Sub AuditShadowAppend()
    Dim cntr As Long
    Dim CardValue As Integer
    Dim cntr2 As Long
    Dim vAuditShadow As String
    
    vAuditShadow = gGsLeUtils.LE5d(AuditShadow)
    
    'Check if already appended.
    If InStr(13, vAuditShadow, "f") > 1 Then
        Exit Sub
    End If
    For cntr = 1 To 6
        'Total units.
        vAuditShadow = vAuditShadow & CStr(CountPlayerUnits(CInt(cntr))) & ","
        'Total cards.
        CardValue = 0
        If Audit(cntr).Player > 0 Then
            For cntr2 = 0 To 10
                CardValue = CardValue + gPlayerID(cntr).card(cntr2)
            Next
        End If
        vAuditShadow = vAuditShadow & "f" & CStr(CardValue) & ","
    Next
    AuditShadow = gGsLeUtils.LE5(vAuditShadow)
End Sub

'Return the total value of cards left in deck.
Private Function AuditGetDeckValue() As Long
    Dim cntr As Long
    
    AuditGetDeckValue = 0
    For cntr = 1 To 4
        AuditGetDeckValue = AuditGetDeckValue + (gCardDeck(cntr - 1) * cntr)
    Next
End Function

'Return "" if player passes the audit and return "Error" and reason if not.
'This is performed by the client terminal and sent to the host after a time delay.
Public Sub AuditPlayerCompare()
    Dim cntr As Long
    Dim cntr2 As Long
    Dim CardValue As Long
    Dim MissingDeckValue As Long
    Dim vResults As String
    Dim vPassed As Boolean
    Dim vTest As String
    
    If netWorkSituation = cNetClient And cmdEnd.Tag = "OK" Then
        vPassed = True
        vResults = ""
        MissingDeckValue = AuditCardDeck - AuditGetDeckValue
        
        For cntr = 1 To 6
            vTest = vTest & "Plr " & CStr(cntr) & " p" & CStr(Audit(cntr).Player) & " m" & CStr(Audit(cntr).PlayerMax) & " "
        
            'Audit total player units for each player.
            If Audit(cntr).Player <> CountPlayerUnits(CInt(cntr)) Then
                vPassed = False
                vResults = vResults _
                            & "Unit counts for " & gPlayerID(cntr).strColor _
                            & "do not match." _
                            & " Expected " & CStr(Audit(cntr).Player) _
                            & ", detected " & CStr(CountPlayerUnits(CInt(cntr))) & "." '& vbCrLf
            ElseIf CountPlayerUnits(CInt(cntr)) > Audit(cntr).PlayerMax Then
                vPassed = False
                vResults = vResults _
                            & "Unit counts for " & gPlayerID(cntr).strColor _
                            & "are beyond the total issued." _
                            & " Expected " & CStr(Audit(cntr).PlayerMax) _
                            & ", detected " & CStr(CountPlayerUnits(CInt(cntr))) & "." '& vbCrLf
            End If
            
            'Audit total card value for player.
            If Audit(cntr).Player > 0 Then
                CardValue = 0
                For cntr2 = 0 To 10
                    CardValue = CardValue + gPlayerID(cntr).card(cntr2)
                Next
                If Audit(cntr).PlayerCard <> CardValue And gCurrentMode <> 13 Then
                    If Audit(cntr).PlayerCard <> CardValue - MissingDeckValue And MissingDeckValue >= 0 Then
                        vPassed = False
                        vResults = vResults _
                                & "Card value for " & gPlayerID(cntr).strColor _
                                & "do not match." _
                                & " Expected " & CStr(Audit(cntr).PlayerCard) _
                                & ", detected " & CStr(CardValue) & "." '& vbCrLf
                    End If
                End If
            End If
        Next
        
        If Not vPassed Then
            LogInfo "AuditPlayerCompare", "Inventory Mismatch: " & netMain.txtTerminalName.Text & " - " & vResults
            LogInfo "AuditPlayerCompare", vTest, 5, True
            vResults = "\beep \clr=255  - Inventory Mismatch -\clr=16777215 " & vbCrLf _
                    & "Terminal name: " & netChatterBox.CreateColorCodeForTerminal _
                    & netMain.txtTerminalName.Text & " \clr=" & CStr(vbWhite) & " " & vbCrLf _
                    & vResults '& "\clr=255       -------------------------\clr=16777215 "
                    
            netMain.tmrFillInfo.Tag = vResults & vbCrLf
        End If
    End If
    
    'Call AuditShadowCompare 'Remove after testing.
    cmdEnd.Tag = "OK"
End Sub

'I am host, check the passed shadow file and whinge if required.
'Convert passed byte buffer and return length of this part. 255 marks the end.
Private Function AuditShadowBytes(pBytBuff() As Byte, pIndex As Long) As Long
    Dim i As Long
    Dim vAuditShadow As String
    Dim vResults As String
    
    If pIndex >= UBound(pBytBuff) Then
        Exit Function
    End If
    
    'Find the length of the bits we want.
    For i = pIndex To UBound(pBytBuff)
        If pBytBuff(i) = 255 Then
            AuditShadowBytes = i - pIndex
            Exit For
        End If
    Next
    
    'Exit here if I am not the host, return the pointer to the next lot of data.
    If netWorkSituation <> cNetHost Then
        Exit Function
    End If
    
    'Convert
    vAuditShadow = gGsLeUtils.LE5d(StrConv(MidB(pBytBuff, pIndex + 1, AuditShadowBytes), vbUnicode))
    
    'Judge.
    vResults = AuditShadowCompare(vAuditShadow)
    If vResults <> "" Then
        Call netMain.XmitStringAll(0, 1, 0, vResults & vbCrLf)
        Call netChatterBox.printMessageString(vResults & vbCrLf)
        LogInfo "AuditShadowBytes", vResults
    End If
    
    'String to byte() - bytBuf() = StrConv("AA" & sMessage, vbFromUnicode)
    'Byte() to string - strText = StrConv(bytBuf(), vbUnicode)
End Function

'Print total score of all players in the message box.
Private Sub AuditPrintClemsCounter()
    Static vWasWinLastTime As Boolean
    Dim vCountActivePlayers As Integer
    Dim vPlayerUnits As Integer
    Dim vPlayer As Integer
    Dim vCounterText As String
    
    'Only if "mr#audit" is active
    If netWorkSituation <> cNetHost Or Not mnuNetCntr.Checked Then
        Exit Sub
    End If
    
    'Header selects font and size.
    vCounterText = "\fname=Courier_New \fsize=6 \fbold=1 \DestWindow=Counter "
    
    'Count total units for each platyer.
    vCountActivePlayers = 0
    For vPlayer = 1 To 6
        vPlayerUnits = CountPlayerUnits(vPlayer)
        vCounterText = vCounterText & "\clr=" & CStr(gPlayerID(vPlayer).bkgndColor) _
                    & " " & Format(vPlayerUnits, "@@@@") & ""
        If vPlayerUnits > 0 Then
            vCountActivePlayers = vCountActivePlayers + 1
        End If
    Next
    
    vCounterText = Replace(vCounterText, " 0", " X")
    
    'Check if last time in this function was a win to stop multiple printings.
    If vCountActivePlayers <= 1 Then
        If vWasWinLastTime Then
            Exit Sub
        Else
            vWasWinLastTime = True
            vCounterText = vCounterText & vbCrLf & "  ---------------------"
        End If
    Else
        vWasWinLastTime = False
    End If
    
    Call netMain.XmitStringAll(0, 1, 0, vCounterText & vbCrLf)
    Call netChatterBox.printMessageString(vCounterText)
End Sub

'Read shadow file to ensure the auditing system integrity.
'Host only.
Private Function AuditShadowCompare(pAuditShadow As String) As String
    Dim vSplit() As String
    Dim cntr As Long
    Dim cntr2 As Long
    Dim vAudit(6) As AuditPlayerType
    Dim vPlayer As Long
    Dim vChange As Integer
    Dim CardValue As Integer
    Dim vMark As Long
    On Error GoTo ErrHand
    
    AuditShadowCompare = ""
    
    'Host only.
    If netWorkSituation <> cNetHost Then
        Exit Function
    End If
    
    'Make sure terminal hasn't recently claimed or released player.
    If cmdSetupOk.Tag <> "" Then
        cmdSetupOk.Tag = ""
        Exit Function
    End If
    
    'Make sure the audit shadow has not been invalidated by being too long.
    If InStr(1, pAuditShadow, "x") > 1 And Len(pAuditShadow) > 5000 Then
        Exit Function
    End If
    
    vSplit = Split(pAuditShadow, ",")
    If Not IsNumeric(vSplit(0)) Then
        AuditShadowCompare = "Corrupted audit returned from the terminal."
        Exit Function
    End If
    
    'Get units and cards at the start of the turn.
    For cntr = 0 To 5
        vAudit(cntr + 1).Player = CInt(vSplit(cntr * 2 + 1))
        vAudit(cntr + 1).PlayerCard = CInt(vSplit(cntr * 2 + 2))
    Next
    
    'Follow how the points changed during the turn.
    vMark = 0
    For cntr = 13 To UBound(vSplit) - 1 Step 2
        vPlayer = CLng(vSplit(cntr))
        vChange = CInt(Mid(vSplit(cntr + 1), 2))
        Select Case Mid(vSplit(cntr + 1), 1, 1)
        Case "u"                'Total units increase.
            vAudit(vPlayer).Player = vAudit(vPlayer).Player + vChange
        Case "d"                'Total units decrease.
            vAudit(vPlayer).Player = vAudit(vPlayer).Player - vChange
        Case "c"                'Total cards increase.
            vAudit(vPlayer).PlayerCard = vAudit(vPlayer).PlayerCard + vChange
        Case "r"                'Total cards decrease.
            vAudit(vPlayer).PlayerCard = vAudit(vPlayer).PlayerCard - vChange
        Case "f"                'End of turn score.
            If vAudit(vMark + 1).Player <> CInt(vSplit(cntr)) Then
                AuditShadowCompare = "Country scores from Terminal " & CStr(vSplit(0)) & " do not match the host's inventory."
                Exit For
            End If
            If vAudit(vMark + 1).PlayerCard <> CInt(Mid(vSplit(cntr + 1), 2)) _
            And vAudit(vMark + 1).Player > 0 Then
                AuditShadowCompare = "Cards from Terminal " & CStr(vSplit(0)) & " do not match the host's inventory."
                Exit For
            End If
            vMark = vMark + 1
        End Select
    Next
    
    If vMark = 0 Then
        AuditShadowCompare = "Corrupted audit returned from terminal " & CStr(vSplit(0)) & "."
    End If
    
    Exit Function
ErrHand:
    AuditShadowCompare = "Corrupted audit returned from terminal " & CStr(vSplit(0)) & "."
    Exit Function
End Function

'This card has been traded by the remote client.
Public Sub AuditTradeCard(pPlayer As Integer, pCard As Integer)
    Dim vSubscript As String
    Dim vAuditShadow As String
    
    If netWorkSituation <> cNetClient Then
        Exit Sub
    End If
    
    vAuditShadow = gGsLeUtils.LE5d(AuditShadow)
    
    Audit(pPlayer).PlayerCard = Audit(pPlayer).PlayerCard - pCard
    
    'Bail out if the audit shadow has been invalidated due to length.
    If InStr(1, vAuditShadow, "x") = 0 Then
    
        'Add to shadow.
        If pCard > 0 Then
            vSubscript = "r"
        Else
            vSubscript = "c"
        End If
        vAuditShadow = vAuditShadow & CStr(pPlayer) & "," & vSubscript & CStr(Abs(pCard)) & ","
        
        AuditShadow = gGsLeUtils.LE5(vAuditShadow)
    End If
End Sub

'Add or subtract units from player.
Private Sub AuditUpdateScore(pPlayer As Integer, pUnits As Integer)
    Dim vSubscript As String
    Dim vAuditShadow As String
    
    If netWorkSituation <> cNetClient Then
        Exit Sub
    End If
    
    vAuditShadow = gGsLeUtils.LE5d(AuditShadow)
    
    Audit(pPlayer).Player = Audit(pPlayer).Player - pUnits
    If Audit(pPlayer).Player < 0 Then
        Audit(pPlayer).Player = 0
    End If
    
    'Bail out if the audit shadow has been invalidated due to length.
    If InStr(1, vAuditShadow, "x") = 0 Then
        'Debug.Print Len(vAuditShadow)
        'Add to shadow.
        If pUnits > 0 Then
            vSubscript = "d"
        Else
            vSubscript = "u"
        End If
        vAuditShadow = vAuditShadow & CStr(pPlayer) & "," & vSubscript & CStr(pUnits) & ","
        
        'Invalidate if too long.
        If Len(vAuditShadow) > 5000 Then
            vAuditShadow = vAuditShadow & "x,"
        End If
        
        AuditShadow = gGsLeUtils.LE5(vAuditShadow)
    End If
End Sub

'Add issued points to max possible player value.
Public Sub AuditAddCardsIssued(pPlayer As Integer, pCardsIssued As Integer)
    Dim vAuditShadow As String
    
    If netWorkSituation <> cNetClient Then
        Exit Sub
    End If
    
    vAuditShadow = gGsLeUtils.LE5d(AuditShadow)
    vAuditShadow = vAuditShadow & CStr(pPlayer) & ",c" & CStr(pCardsIssued) & ","
    AuditShadow = gGsLeUtils.LE5(vAuditShadow)
End Sub

'Add issued points to max possible player value.
Public Sub AuditAddPointsIssued(pPlayer As Integer, pPointsIssued As Integer)
    Dim vAuditShadow As String
    
    If netWorkSituation <> cNetClient Then
        Exit Sub
    End If
    
    vAuditShadow = gGsLeUtils.LE5d(AuditShadow)
    Audit(pPlayer).PlayerMax = Audit(pPlayer).PlayerMax + pPointsIssued
    Audit(pPlayer).Player = Audit(pPlayer).Player + pPointsIssued
    
    'Bail out if the audit shadow has been invalidated due to length.
    If InStr(1, vAuditShadow, "x") = 0 Then
        vAuditShadow = vAuditShadow & CStr(pPlayer) & ",u" & CStr(pPointsIssued) & ","
        AuditShadow = gGsLeUtils.LE5(vAuditShadow)
    End If
End Sub

'Count countries held by current player.
Private Function CountPlayerUnits(pPlayerWho As Integer, Optional pDefeatedCtry As Integer = 0) As Integer
    Dim cntr1 As Long
    
    CountPlayerUnits = 0
    For cntr1 = 1 To 42
        If gCountryOwner(cntr1) = pPlayerWho And cntr1 <> pDefeatedCtry Then
            CountPlayerUnits = CountPlayerUnits + gCtryScore(cntr1)
        End If
    Next cntr1
End Function

Public Sub TurnOffCheatCodes()
    Dim i As Long
    
    gCheatMode.cheatActive = False
    gCheatMode.createMap = False
    gCheatMode.seeCards = False
    gCheatMode.seeMissions = False
    gCheatMode.undoEnabled = False
End Sub

'Get name of player controller (1-6).
Public Function GetArmyOrControllerName(pPlayerController As Byte) As String
    If (playerSelect_getIndex(CInt(pPlayerController - 1)) = 0) Then
        GetArmyOrControllerName = netMain.txtTerminalName
    Else
        GetArmyOrControllerName = PlayerSelect(pPlayerController - 1).Text
    End If
End Function

    ' Return True if whichPlayer is a computer controlled by this terminal
Public Function IsComputerPlayer(whichPlayer As Integer) As Boolean
    Dim plrType As Integer
    
    plrType = gPlayerID(whichPlayer).playerWho
    If plrType = 1 Or plrType = 2 Or plrType = A3Index Or unclaimedPlayer(whichPlayer) Then
        IsComputerPlayer = True
    Else
        IsComputerPlayer = False
    End If
End Function

'Validate or invalidate all stats.
Public Sub ValidateAllStats(pValid As Boolean)
    Dim vPlayer As Long
    
    For vPlayer = 1 To 6
        gPlayerStats(vPlayer).IsValid = pValid
        If Not pValid Then
            gPlayerStats(vPlayer).InvalidatedReason = "Not present for" & vbCrLf & "the full war."
        Else
            gPlayerStats(vPlayer).InvalidatedReason = "OK."
        End If
    Next
End Sub

'These cards have been picked by remote player
Public Sub exchangeTheseCards(PlayerByte As Byte, BytHold() As Byte)
    Dim i As Long
    Dim pTurn As Long
    Dim sep As Long
    Dim tmp2 As Integer
    Dim vSelectCards(2) As Integer
    
    If net.redrawCards Then
        tmp2 = gCurrentMode
        gCurrentMode = 2
        Call DrawAllCards
        gCurrentMode = tmp2
        net.redrawCards = False
    End If
    
    'Display these cards as being picked.
    pTurn = CLng(PlayerByte)
    For i = 0 To 2
        vSelectCards(i) = gPlayerID(pTurn).card(CLng(BytHold(i + 2)) - 1)
        gCurrentMode = 5
        Call DrawBigCard(vSelectCards(i), CInt(BytHold(i + 2)), False)
        gPlayerID(pTurn).pickedCards(CLng(BytHold(i + 2)) - 1) = False
    Next
    
    'Remove picked cards from hand and copy other cards up to fill the gap.
    sep = 0
    For i = 1 To 10
        If gPlayerID(gPlayerTurn).card(i - 1) = 0 Then
            Exit For
        End If
        gPlayerID(gPlayerTurn).card(i - sep - 1) = gPlayerID(gPlayerTurn).card(i - 1)
        If gPlayerID(gPlayerTurn).pickedCards(i - 1) = False Then
            sep = sep + 1
            gPlayerID(gPlayerTurn).pickedCards(i - 1) = True
            net.redrawCards = True
        End If
    Next
    
    'Remove cards cards that have been copied.
    For i = i - sep To i - 1
        gPlayerID(gPlayerTurn).card(i - 1) = 0
    Next
    
    Call SyncForgroundMap("exchangeTheseCards")
End Sub

    'Remote player has won the game
Public Sub gameWon(PlayerByte As Byte)
    gCurrentMode = 1
    InfoBoxPrint 0
    If PlayerByte = 1 Then
        Call CheckWinDuringTurn(gPlayerTurn, True)
    Else
        Call SetBit(True, CLng(gPlayerTurn), gWinMemoryBits)
        Call CheckWinStartOfTurn(gPlayerTurn, True)
    End If

    net.madeUpdate = True
    
    'Make sure the viewport gets refreshed.
    gSyncViewportNeeded = True
    Call SyncForgroundMap("gameWon")
End Sub

    'Transmit info text if true
Public Sub xmitInfoText(TorF As Boolean)
    net.xmitInfTxt = TorF
End Sub

'Set player owners to 0, none claimed yet and enable player selector controls.
Public Sub resetPlayerOwners()
    Dim i As Integer
    
    cmdSetupOk.Tag = "resetPlayerOwners."
    For i = 0 To 5
        net.playerOwner(i) = 0
        net.Controller(i) = 0
        RenameRemote CByte(i), Phrase(28)
        Call PlayerSelect_Enable(i, True)
    Next
End Sub

'Return the number of claimed players over total players.
Public Function CountClaimedPlayers() As String
    Dim i As Integer
    Dim Count As Integer
    Dim vTotal As Integer
    Dim vResult As Integer
    
    Count = 0
    For i = 0 To 5
        
        'Count claimed players.
        If net.playerOwner(i) <> 0 Then
            Count = Count + 1
        ElseIf playerSelect_getIndex(i) <> remoteIndex Then
            Count = Count + 1
        End If
        
        'Count total players.
        If CInt(txtPlayerStartCountries(i).Text) <> 0 Then
            vTotal = vTotal + 1
        End If
    Next
    vResult = vTotal - Count
    If vResult < 0 Then
        vResult = 0
    End If
    CountClaimedPlayers = CStr(vResult)
End Function

'Terminal has claimed this player. Not actually used.
Public Sub setPlayerOwner(whichPlayer As Byte, terminal As Byte)
    net.playerOwner(whichPlayer) = terminal
    If terminal = 0 Then
        RenameRemote whichPlayer, Phrase(28)
        Call PlayerSelect_Enable(CInt(whichPlayer), True)
    Else
        RenameRemote whichPlayer, net.ClientName(whichPlayer)
        Call PlayerSelect_Enable(CInt(whichPlayer), False)
    End If
End Sub

'Lost remote terminal, free claimed players (I am host).
'If in the middle of a war then change to a computer player, press reset.
Public Sub lostPlayerOwner(terminal As Byte)
    Dim i As Integer
    Dim tmp2 As Byte
    Dim PressEnd As Boolean
    Dim vsetupControlChange As Boolean
    
    On Error GoTo ErrHand
    
    net.LockControls = True
    vsetupControlChange = net.setupControlChange
    tmp2 = netWorkSituation
    netWorkSituation = cNetNone
    For i = 0 To 5
        If net.playerOwner(i) = terminal Then
            net.playerOwner(i) = 0
            net.Controller(i) = 0
            RenameRemote CByte(i), Phrase(28)                  'Remote Player.
            Call PlayerSelect_Enable(i, True)
            net.LockControls = False
            Call netMain.NotifyDisclamedPlayer(CByte(i))        'Notify other terminals.
            net.LockControls = True
            If Not SetupScreen.Visible Then
                gPlayerID(i + 1).playerWho = remoteIndex   'Make computer player
                If (i + 1) = gPlayerTurn Then            'Take over turn
                    PressEnd = True
                End If
            End If
        End If
    Next
    'Call netMain.sendSetupScreen
    netWorkSituation = tmp2
    If PressEnd Then
        gComputerPressed = True
        Call EndClicked
        gComputerPressed = False
    End If
    net.setupControlChange = vsetupControlChange
    net.LockControls = False
    Exit Sub
ErrHand:
    Resume Next
End Sub

'Change name of Remote Player in Setup screen. Used for networking.
Public Sub RenameRemote(pPlayerID As Byte, pNewRemoteName As String)
    Dim tmp As Long
    Dim cntr As Long
    Dim tmp2 As Byte
    
    On Error Resume Next
    
    tmp = PlayerSelect(pPlayerID).ListIndex
    PlayerSelect(pPlayerID).Clear
    PlayerSelect(pPlayerID).AddItem Phrase(7), 0         'Human
    If net.playerOwner(pPlayerID) <> myTerminalNumber Then
        PlayerSelect(pPlayerID).AddItem GetRemoteTerminalMarker & pNewRemoteName, 1      ':Remote
    Else
        PlayerSelect(pPlayerID).AddItem pNewRemoteName, 1      'Remote
    End If
    PlayerSelect(pPlayerID).AddItem Phrase(8), 2         'Average computer
    PlayerSelect(pPlayerID).AddItem Phrase(9), 3         'Smart computer
    PlayerSelect(pPlayerID).AddItem Phrase(36), 4        'Intelligent computer player
    PlayerSelect(pPlayerID).ListIndex = tmp
    'If netWorkSituation <> cNetNone Then
    '    If netMain.Visible Then netMain.SetFocus
    'End If
End Sub

'Some one has claimed this player. Return True if successful.
'Countepart to ClaimPlayer() via command 5.
'bytMes(2) = 0 - Disclaim player(whichPlayer).
'bytMes(2) > 0 - Terminal number claims player(whichPlayer).
'bytMes(3) - Player controller. 0 for human anything else is remote or computer.
'bytMes(4) - Terminal name.
Public Function DibsOnPlayer(BytMes() As Byte, whichPlayer As Byte, pTerminal As Integer) As Boolean
    Dim tst As Boolean
    Dim i As Long
    Dim vPlayerCount As Long
    Dim tmp As String
    
    On Error GoTo ErrHand
    
    DibsOnPlayer = False
    
    'Check and make sure client can have this player.
    If (net.playerOwner(whichPlayer) <> 0 _
    Or playerSelect_getIndex(CInt(whichPlayer)) <> remoteIndex) _
    And BytMes(2) <> 0 And netWorkSituation = cNetHost Then
        'You can't have him.
        Call netMain.SendSetupScreenTo(CByte(pTerminal))
        Exit Function
    End If
    
    'Checks if I am the host.
    If netWorkSituation = cNetHost Then
        'Attempting to disclaim player make sure this terminal owns it (I am host).
        If BytMes(2) = 0 And net.playerOwner(whichPlayer) <> pTerminal Then
            'You can't disclaim him.
            Call netMain.SendSetupScreenTo(CByte(pTerminal))
            Exit Function
        End If
        
        'Attempting to claim player, make sure it is free (I am host).
        If BytMes(2) > 0 And net.playerOwner(whichPlayer) <> 0 Then
            'You can't claim him. Already taken.
            Call netMain.SendSetupScreenTo(CByte(pTerminal))
            Exit Function
        End If
    'Else
        'You could put a check here to make sure this was sent by the host but
        'if I am a client, I can only be connected to the host.
    End If
    
    'Check number of players already owned by this terminal.
    'Host can own any number of players.
    vPlayerCount = 0
    If BytMes(2) <> 0 And netWorkSituation = cNetHost Then
        If CLng(netMain.txtMaxPlayers.Text) = 0 Then
            'Nup, try again.
            Call netMain.SendSetupScreenTo(BytMes(2))
            Call netMain.XmitString(CLng(BytMes(2)), 1, 0, "Spectators only in this war." & vbCrLf)
            Exit Function
        End If
        
        For i = 1 To 6
            If net.playerOwner(i - 1) = BytMes(2) Then
                vPlayerCount = vPlayerCount + 1
                If vPlayerCount >= CLng(netMain.txtMaxPlayers.Text) Then
                    'No, too many players claimed. Keep trying.
                    Call netMain.SendSetupScreenTo(BytMes(2))
                    If Trim(netMain.txtMaxPlayers.Text) = "1" Then
                        tmp = "player"
                    Else
                        tmp = "players"
                    End If
                    Call netMain.XmitString(CLng(BytMes(2)), 1, 0, "War limited to " _
                                    & Trim(netMain.txtMaxPlayers.Text) & " " & tmp & " per terminal." & vbCrLf)
                    Exit Function
                End If
            End If
        Next
    End If
    
    'Cleared to claim or disclaim player.
    net.LockControls = True
    cmdSetupOk.Tag = "Dibs on player."
    
    If gPlayerTurn = 0 Then
        gPlayerTurn = 1
    End If
    
    tst = (whichPlayer = CByte(gPlayerTurn - 1)) _
        And (netWorkSituation = cNetHost) _
        And (Not SetupScreen.Visible) And myTerminalNumber = 0
        
    net.playerOwner(whichPlayer) = BytMes(2)
    net.Controller(whichPlayer) = BytMes(3)
    DibsOnPlayer = True
    
    'Invalidate stats.
    gPlayerStats(whichPlayer + 1).IsValid = False
    gPlayerStats(whichPlayer + 1).InvalidatedReason = "Changed controller" _
                                                & vbCrLf & "during the war."
    
    If BytMes(2) = 0 Then
        'Un claimed by client
        RenameRemote whichPlayer, Phrase(28)            'Remote Player.
        Call PlayerSelect_Enable(CInt(whichPlayer), True)
        
        'End turn if it is lost client's turn.
        If tst Then
            gComputerPressed = True
            Call EndClicked
            gComputerPressed = False
        End If
    Else
        RenameRemote whichPlayer, StrConv(MidB(BytMes, 5), vbUnicode)
        Call PlayerSelect_Enable(CInt(whichPlayer), False)
    End If
    net.setupControlChange = False
    net.LockControls = False
    Exit Function
ErrHand:
    Resume Next
End Function

'Recieved order to release this player to another terminal.
Public Sub releaseMyPlayer(BytMes() As Byte, whichPlayer As Byte)
    On Error GoTo ErrHand
    
    net.LockControls = True
    cmdSetupOk.Tag = "releaseMyPlayer."
    net.playerOwner(whichPlayer) = BytMes(3)
    net.Controller(whichPlayer) = remoteIndex
    Call playerSelect_showIndex(CInt(whichPlayer), remoteIndex)
    RenameRemote whichPlayer, StrConv(MidB(BytMes, 4), vbUnicode)
    Call PlayerSelect_Enable(CInt(whichPlayer), False)
    net.LockControls = False
    Exit Sub
ErrHand:
    Resume Next
End Sub

'Show message if NetMain is not visible.
Public Sub PostMessage(messageText As String)
    On Error Resume Next
    If Not netMain.Visible Then
        'netChatterBox.Show , Me
        Call netChatterBox.printMessageString(messageText)
    End If
End Sub

'Test new setupscreen locking system.
Private Sub LockSetupControls(Optional pUnlock As Boolean = False)
    Dim Ctrl As Control
    For Each Ctrl In TheMainForm.Controls
        If Ctrl.name = "PlayerSelect" Then
            Debug.Print Ctrl.name
        End If
    Next Ctrl
End Sub

'Enables or disable setup controls. Used by the networking system to set up
'the terminal as either a host (pEnabled = TRUE) or a client (pEnabled = FALSE).
Public Sub EnableSetupControls(pEnabled As Boolean)
    Dim vIndex As Integer
    
    On Error Resume Next
    
    'Start options box.
    lblSetupStart.Enabled = pEnabled
    lblStartingArmies.Enabled = pEnabled
    'lblExtraStartingUnits.Enabled = pEnabled
    lblSetupBattleOptions.Enabled = pEnabled
    lblSetupFirst.Enabled = pEnabled
    fSetupPlayerNumber.Enabled = pEnabled
    fSetupWarOptions.Enabled = pEnabled
    cmdSUPcncl.Enabled = pEnabled
    mnuCancelSetup.Enabled = cmdSUPcncl.Enabled
    mnuFileReset.Enabled = pEnabled
    mnuFileLoadWar.Enabled = pEnabled
    lblPlayerOptions.Enabled = pEnabled
    
    'Lock player starting countries.
    For vIndex = 0 To txtPlayerStartCountries.Count - 1
        txtPlayerStartCountries(vIndex).Enabled = pEnabled
        udPlayerStartCountries(vIndex).Enabled = pEnabled
    Next
    udStartingArmies.Enabled = pEnabled
    txtStartingArmies.Enabled = pEnabled
    chkExtraStartingUnits.Enabled = pEnabled
    udExtraStartingUnits.Enabled = pEnabled And (chkExtraStartingUnits.Value = vbChecked)
    txtExtraStartingUnits.Enabled = pEnabled And (chkExtraStartingUnits.Value = vbChecked)
    optRandomFirstPlayer.Enabled = pEnabled
    optPlr1FirstPlayer.Enabled = pEnabled
    chkMsnMissionsOn.Enabled = pEnabled
    optSupplyLines.Enabled = pEnabled
    optLimitSupply.Enabled = pEnabled
    optNoSupply.Enabled = pEnabled
    
    'Cards
    frmSetupCards.Enabled = pEnabled
    frmTheDeck.Enabled = pEnabled
    frmFixedValues.Enabled = pEnabled
    lblSetupCards.Enabled = pEnabled
    For vIndex = 0 To optCardMode.Count - 1
        optCardMode(vIndex).Enabled = pEnabled
    Next
    chkCardsHidden.Enabled = pEnabled And (optCardMode(0).Value = vbUnchecked)
    chkCardsVulture.Enabled = pEnabled And (optCardMode(0).Value = vbUnchecked)
    udMaximumCardValue.Enabled = pEnabled And (optCardMode(2).Value = vbChecked)
    txtMaximumCardValue.Enabled = pEnabled And (optCardMode(2).Value = vbChecked)
    lblSetupMaxCardValue.Enabled = pEnabled And (optCardMode(2).Value = vbChecked)
    
    'The Deck
    lblTheDeck.Enabled = pEnabled
    For vIndex = 0 To txtCardDeck.Count - 1
        txtCardDeck(vIndex).Enabled = pEnabled
        udCardDeck(vIndex).Enabled = pEnabled
        lblCardDeck(vIndex).Enabled = pEnabled
    Next
    
    'Fixed values.
    lbCardValues.Enabled = pEnabled
    For vIndex = 0 To txtFixedValues.Count - 1
        txtFixedValues(vIndex).Enabled = pEnabled
        udFixedValues(vIndex).Enabled = pEnabled
        lblFixedValues(vIndex).Enabled = pEnabled
    Next
    
    'Dice
    frmDiceRules.Enabled = pEnabled
    frmDiceThrows.Enabled = pEnabled
    lblDiceRules.Enabled = pEnabled
    lblDiceThrow.Enabled = pEnabled
    For vIndex = 0 To optDiceRules.Count - 1
        optDiceRules(vIndex).Enabled = pEnabled
    Next
    
    'Reinforcments tab.
    For vIndex = 0 To udNewUnitClac.Count - 1
        udNewUnitClac(vIndex).Enabled = pEnabled
        txtNewUnitClac(vIndex).Enabled = pEnabled
    Next
    
    For vIndex = 0 To lblNewUnitClac.Count - 1
        lblNewUnitClac(vIndex).Enabled = pEnabled
    Next
    
    For vIndex = 0 To udContVal.Count - 1
        udContVal(vIndex).Enabled = pEnabled
        txtContVal(vIndex).Enabled = pEnabled
        lblContVal(vIndex).Enabled = pEnabled
    Next
    
    Call EnableDiceOptions(pEnabled)
    
    Call EnableMissionOptions(pEnabled)
    
    If gCurrentMode = 100 Then
        cmdSetupOk.Caption = Phrase(29) ' "Join &War"
        cmdSetupOk.Enabled = True
    Else
        cmdSetupOk.Caption = Phrase(26)
        cmdSetupOk.Enabled = pEnabled
    End If
    
    mnuDeclareWar.Caption = cmdSetupOk.Caption
    mnuDeclareWar.Enabled = cmdSetupOk.Enabled
    
    If netWorkSituation <> cNetNone Then
        If netMain.Visible Then
            netMain.SetFocus
        End If
    End If
End Sub

Private Sub EnableDiceOptions(pEnabled As Boolean)
    Dim vIndex As Long
    
    chkSortDice.Enabled = pEnabled And Not optDiceRules(0).Value
    lblSameDice.Enabled = pEnabled And Not optDiceRules(0).Value
    For vIndex = 0 To optDiceSame.Count - 1
        optDiceSame(vIndex).Enabled = pEnabled And Not optDiceRules(0).Value
    Next
    lblDiceThrow.Enabled = pEnabled And Not optDiceRules(0).Value
    For vIndex = 0 To txtDiceThrown.Count - 1
        txtDiceThrown(vIndex).Enabled = pEnabled And Not optDiceRules(0).Value
        udDiceThrown(vIndex).Enabled = pEnabled And Not optDiceRules(0).Value
        lblDiceThrown(vIndex).Enabled = pEnabled And Not optDiceRules(0).Value
    Next
End Sub

    'Return the name of the current war
Public Function GetWarName() As String
    GetWarName = Trim(warSit.filename)
End Function

    'Make this the name of the war
Public Sub makeWarName(newFileName As String)
    warSit.filename = newFileName
    If SetupScreen.Visible Then
        Call ChangeTitlebarText(Phrase(34) + Trim(warSit.filename))      'Global Siege Set Up...
    Else
        Call ChangeTitlebarText(Phrase(35) + Trim(warSit.filename))      'Global Siege -
    End If
End Sub

    'Add "Remote player" to list
Public Sub ResetPlayerList()
    Dim vHoldPlayerIx As Long
    Dim cntr As Integer
    Dim vHoldNwSitu As Byte
    
    On Error Resume Next
    
    vHoldNwSitu = netWorkSituation
    netWorkSituation = cNetNone
    For cntr = 0 To 5
        Call PlayerSelect_Enable(cntr, True)
        
        vHoldPlayerIx = playerSelect_getIndex(CInt(cntr)) 'PlayerSelect(cntr).ListIndex
        PlayerSelect(cntr).Clear
        PlayerSelect(cntr).AddItem Phrase(7), 0     'Human
        'PlayerSelect(cntr).AddItem Phrase(28), 1    'Remote
        If net.playerOwner(cntr) <> myTerminalNumber Then
            PlayerSelect(cntr).AddItem GetRemoteTerminalMarker & Phrase(28), 1      ':Remote
        Else
            PlayerSelect(cntr).AddItem Phrase(28), 1      'Remote
        End If
        PlayerSelect(cntr).AddItem Phrase(8), 2     'Average computer
        PlayerSelect(cntr).AddItem Phrase(9), 3    'Smart computer
        PlayerSelect(cntr).AddItem Phrase(36), 4 'Intelligent computer player
        'PlayerSelect(cntr).ListIndex = vHoldPlayerIx
        Call playerSelect_showIndex(CInt(cntr), CInt(vHoldPlayerIx))
    Next
    netWorkSituation = vHoldNwSitu
    If netWorkSituation <> cNetNone Then
        If netMain.Visible Then
            netMain.SetFocus
        End If
    End If
End Sub

    'Command to send update info
Private Sub handleUpdate()
    If netWorkSituation <> cNetNone Then
        Call netMain.handleUpdate(CByte(net.highestPriority))
    End If
End Sub

'Update game from composed list received from remote (or host).
'Called every move. Packed by function ComposeNetUpdate()
Public Sub UnpackNetUpdate(byteMess() As Byte)
    Dim vIndex As Long
    Dim vInfoTextLength As Long
    Dim vPlayer As Byte
    Dim vCardIndex As Byte
    Dim vHold1 As Integer
    Dim vCurrentMode As Integer
    Dim vDiceHold As Long
    Dim vIndex2 As Long
    Dim vNewScore As Integer
    Dim vAttackDice(cMaxNumberOfDice - 1) As Integer
    Dim vDefenceDice(cMaxNumberOfDice - 1) As Integer
    
    If frmMissions.Visible Then
        frmMissions.Hide
    End If
    
    gNewWarPlayerTurn = 0
    
    If gCurrentMode = 100 And SetupScreen.Visible Then
        Exit Sub
    End If
    
    'Only I make updates for players on my terminal
    If net.playerOwner(gPlayerTurn - 1) = myTerminalNumber Then
        Exit Sub
    End If
    
    'Redraw cards if cards have recently been exchanged.
    If net.redrawCards Then
        vHold1 = gCurrentMode
        gCurrentMode = 2
        Call DrawAllCards
        gCurrentMode = vHold1
        net.redrawCards = False
    End If
    
    'Get changed country details.
    For vIndex = 0 To ((UBound(byteMess) - 2 - 4) \ 4)
        If byteMess(vIndex * 4 + 2) = 0 Then
            'Error
            Exit Sub
        ElseIf byteMess(vIndex * 4 + 2) = 255 Then     'End of the first part
            Exit For
        ElseIf byteMess(vIndex * 4 + 2) < 50 Then
            vHold1 = gCountryOwner(CLng(byteMess(vIndex * 4 + 2)))
            
            vNewScore = CInt(byteMess(vIndex * 4 + 2 + 2) + byteMess(vIndex * 4 + 2 + 3) * 256)
            
            'If the below are different, country was defeated.
            gCountryOwner(CLng(byteMess(vIndex * 4 + 2))) = CInt(byteMess(vIndex * 4 + 2 + 1))
            Call printNewScore(CLng(byteMess(vIndex * 4 + 2)), vNewScore)
            Call ColorCountry(CLng(byteMess(vIndex * 4 + 2)), _
            gPlayerID(byteMess(vIndex * 4 + 2 + 1)).lngColor)
            
            'Also means defeat if below are equal.
            If CInt(byteMess(vIndex * 4 + 2)) = CountryUnderAttack Then
                'gCountryIsBlack = 0
                CountryUnderAttack = 0
            End If
            
            'Vulture cards?
            If vHold1 <> CInt(byteMess(vIndex * 4 + 2 + 1)) Then
                vCurrentMode = gCurrentMode
                gCurrentMode = 2
                Call CheckVultureCard(vHold1)
                gCurrentMode = vCurrentMode
            End If
                    
        ElseIf byteMess(vIndex * 4 + 2) < 110 Then
            'Card list?
            vPlayer = (byteMess(vIndex * 4 + 2) - 50) \ 10 + 1
            vCardIndex = ((byteMess(vIndex * 4 + 2) - 50) Mod 10) + 1
            gPlayerID(CLng(vPlayer)).card(CLng(vCardIndex) - 1) = CInt(byteMess(vIndex * 4 + 2 + 1))
            
        ElseIf byteMess(vIndex * 4 + 2) <= 172 Then
            'Country being attacked.
            vHold1 = gCountryOwner(CLng(byteMess(vIndex * 4 + 2) - 130))
            vNewScore = CInt(byteMess(vIndex * 4 + 2 + 2) + byteMess(vIndex * 4 + 2 + 3) * 256)
            
            gCountryOwner(CLng(byteMess(vIndex * 4 + 2) - 130)) = CInt(byteMess(vIndex * 4 + 2 + 1))
            Call ChangeScoreUnderAttack(CLng(byteMess(vIndex * 4 + 2) - 130), vNewScore)
            Call ColorCountryUnderAttack(CLng(byteMess(vIndex * 4 + 2) - 130))
            CountryUnderAttack = CInt(byteMess(vIndex * 4 + 2) - 130)
            
            If vHold1 <> CInt(byteMess(vIndex * 4 + 2 + 1)) Then
                vCurrentMode = gCurrentMode
                gCurrentMode = 2
                Call CheckVultureCard(vHold1)
                gCurrentMode = vCurrentMode
            End If
        End If
    Next
    
    gPlayerTurn = CInt(byteMess(vIndex * 4 + 3) And 7)    'Player turn
    gCurrentCardValue = CInt(byteMess(vIndex * 4 + 4))    'Card value
    
    'Get attack dice info.
    vDiceHold = CLng(byteMess(vIndex * 4 + 5)) + (CLng(byteMess(vIndex * 4 + 6)) * &H100)
    vAttackDice(0) = CInt(vDiceHold And &H7)
    vAttackDice(1) = CInt((vDiceHold \ &H8) And &H7)
    vAttackDice(2) = CInt((vDiceHold \ &H40) And &H7)
    vAttackDice(3) = CInt((vDiceHold \ &H200) And &H7)
    vAttackDice(4) = CInt((vDiceHold \ &H1000) And &H7)
    
    'Get defence dice info.
    vDiceHold = CLng(byteMess(vIndex * 4 + 7)) + (CLng(byteMess(vIndex * 4 + 8)) * &H100)
    vDefenceDice(0) = CInt(vDiceHold And &H7)
    vDefenceDice(1) = CInt((vDiceHold \ &H8) And &H7)
    vDefenceDice(2) = CInt((vDiceHold \ &H40) And &H7)
    vDefenceDice(3) = CInt((vDiceHold \ &H200) And &H7)
    vDefenceDice(4) = CInt((vDiceHold \ &H1000) And &H7)
    
    TheMainForm.pctInfoBox.BackColor = gPlayerID(gPlayerTurn).bkgndColor
    If flashingBorder And (gPlayerID(gPlayerTurn).playerWho = 0) Then
        If netWorkSituation <> cNetNone Then       'Flash color if human on this terminal
            If (net.playerOwner(gPlayerTurn - 1) = myTerminalNumber) Then
                TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
            Else
                TheMainForm.BackColor = &H8000000F
            End If
        Else
            TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
        End If
    Else
        TheMainForm.BackColor = &H8000000F
    End If
    vInfoTextLength = InfoBoxPrintBytes(byteMess, vIndex * 4 + 10)
    
    'Show dice.
    If vDiceHold > 0 Then
        Call DisplayDiceOnBoard(vAttackDice, vDefenceDice)
    End If
    
    Call DrawLittleCards
    Call SyncForgroundMap("UnpackNetUpdate")
    Call resetChangeList
End Sub

'Compose a game from change list ready to send
'Return false if failed. Gets called every move.
'Received by function UnpackNetUpdate()
Public Function ComposeNetUpdate(pByteMess() As Byte) As Boolean
    Dim vIndex As Long
    Dim vIndex2 As Long
    Dim DiceHold As Long
    
    'No need to send.
    If (UBound(net.changeList) = 0) _
    And (net.changeList(0) = 0) _
    Or net.madeUpdate Then
        ComposeNetUpdate = False
        Exit Function
    End If
    
    net.madeUpdate = True
    net.highestPriority = 0
    
    'Bail out if the current player is not on this terminal.
    If net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber Then
        ComposeNetUpdate = False
        Exit Function
    End If
    
    'Resize array to accomodate the change list, player turn,
    'current card value, dice, and pctInfo text.
    ReDim pByteMess(UBound(net.changeList) * 4 + 4 + 9) As Byte
    
    For vIndex = 0 To UBound(net.changeList)
        If net.changeList(vIndex) <= 42 Then         'Countries.
            pByteMess(vIndex * 4 + 2) = net.changeList(vIndex)
            pByteMess(vIndex * 4 + 2 + 1) = CByte(gCountryOwner(net.changeList(vIndex)))
            pByteMess(vIndex * 4 + 2 + 2) = getLowerByte(gCtryScore(net.changeList(vIndex)))
            pByteMess(vIndex * 4 + 2 + 3) = getUpperByte(gCtryScore(net.changeList(vIndex)))
        ElseIf net.changeList(vIndex) < 110 Then     'Cards.
            pByteMess(vIndex * 4 + 2) = net.changeList(vIndex)
            pByteMess(vIndex * 4 + 2 + 1) = _
            gPlayerID((net.changeList(vIndex) - 50) \ 10 + 1).card(((net.changeList(vIndex) - 50) Mod 10))
        ElseIf net.changeList(vIndex) <= 172 Then    '?
            pByteMess(vIndex * 4 + 2) = net.changeList(vIndex)
            pByteMess(vIndex * 4 + 2 + 1) = CByte(gCountryOwner(net.changeList(vIndex) - 130))
            pByteMess(vIndex * 4 + 2 + 2) = getLowerByte(gCtryScore(net.changeList(vIndex) - 130))
            pByteMess(vIndex * 4 + 2 + 3) = getUpperByte(gCtryScore(net.changeList(vIndex) - 130))
        End If
    Next
    
    pByteMess(vIndex * 4 + 2) = 255                   'End of first part
    pByteMess(vIndex * 4 + 3) = CByte(gPlayerTurn)
    pByteMess(vIndex * 4 + 4) = CByte(gCurrentCardValue)
    
    'Only send dice info if refresh rate set to high (0).
    'Dice info is sent as octals in two bytes.
    If netMain.terminalSpeed = 0 Then
        
        'Attack dice.
        DiceHold = net.RolledAttackDice(0) _
                 + net.RolledAttackDice(1) * &H8 _
                 + net.RolledAttackDice(2) * &H40 _
                 + net.RolledAttackDice(3) * &H200 _
                 + net.RolledAttackDice(4) * &H1000
        pByteMess(vIndex * 4 + 5) = CByte(DiceHold And &HFF)
        pByteMess(vIndex * 4 + 6) = CByte((DiceHold And &HFF00) \ &H100)
        
        'Defence dice.
        DiceHold = net.RolledDefenceDice(0) _
                 + net.RolledDefenceDice(1) * &H8 _
                 + net.RolledDefenceDice(2) * &H40 _
                 + net.RolledDefenceDice(3) * &H200 _
                 + net.RolledDefenceDice(4) * &H1000
        pByteMess(vIndex * 4 + 7) = CByte(DiceHold And &HFF)
        pByteMess(vIndex * 4 + 8) = CByte((DiceHold And &HFF00) \ &H100)
        
    Else
        pByteMess(vIndex * 4 + 5) = 0
        pByteMess(vIndex * 4 + 6) = 0
        pByteMess(vIndex * 4 + 7) = 0
        pByteMess(vIndex * 4 + 8) = 0
    End If
    
    'Clear Dice
    For vIndex2 = 0 To cMaxNumberOfDice - 1
        net.RolledAttackDice(vIndex2) = 0
        net.RolledDefenceDice(vIndex2) = 0
    Next
    
    pByteMess(vIndex * 4 + 9) = 0                     'Spare

    Call appendByteArray(pByteMess, net.pctInfoByt)      'Ends with 255
    
    Call resetChangeList
    ComposeNetUpdate = True
End Function

'Add a change to the list and set highest priority.
'1   - 42  = countries.
'50  - 109 = cards (player-1)*10 + 50.
'120 - 130 = picked cards.
'131 - 172 = country under attack.
Private Sub addChangeToList(whatChanged As Integer, cardOwner As Integer, changePriority As Long)
    Dim i As Long

    If netWorkSituation = cNetNone Then Exit Sub
    If gPlayerTurn = 0 Then Exit Sub
    If net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber Then Exit Sub
    
    If net.highestPriority < changePriority Then
        net.highestPriority = changePriority
    End If
    
    For i = 0 To UBound(net.changeList)
        If checkChangeList(i, whatChanged, cardOwner) Then
            Call putChange(whatChanged, cardOwner, i)
            Exit Sub
        End If
    Next
    ReDim Preserve net.changeList(i) As Byte
    Call putChange(whatChanged, cardOwner, i)
End Sub

    'Return true if change is already in list
Private Function checkChangeList(Indx As Long, whatChanged As Integer, cardOwner As Integer) As Boolean
    If net.changeList(Indx) < 50 Then
        checkChangeList = ((net.changeList(Indx) = whatChanged) _
                            Or (net.changeList(Indx) + 130 = whatChanged) _
                            Or (net.changeList(Indx) = 0))
    ElseIf net.changeList(Indx) <= 130 Then
        checkChangeList = _
        (net.changeList(Indx) = (cardOwner - 1) * 10 + 50 + (whatChanged - 1))
    Else
        checkChangeList = ((net.changeList(Indx) = whatChanged) _
                            Or (net.changeList(Indx) - 130 = whatChanged) _
                            Or (net.changeList(Indx) = 0))
    End If
End Function

    'Put change to list at index
Private Sub putChange(whatChanged As Integer, cardOwner As Integer, Indx As Long)
    If whatChanged < 50 Then
        net.changeList(Indx) = CByte(whatChanged)
    ElseIf whatChanged <= 130 Then
        net.changeList(Indx) = _
        (cardOwner - 1) * 10 + 50 + (whatChanged - 1)
    Else
        net.changeList(Indx) = CByte(whatChanged)
    End If
End Sub

    'Clear change list
Private Sub resetChangeList()
    ReDim net.changeList(0) As Byte
End Sub

'Pack all scores on this terminal. All scores if I am host.
Private Sub PackTerminalScore(ByteMessage() As Byte)
    Dim bScores() As Byte
    Dim bHold() As Byte
    Dim cntr As Long
    Dim lPosition As Long
    
    'First byte contains count of players reported.
    ReDim Preserve bScores(1) As Byte
    
    lPosition = 1
    
    For cntr = 0 To 5
        'If I control this player ot if I am host.
        'playerID(cntr + 1).startWith > 0
        'If net.playerOwner(cntr) = myTerminalNumber _
        'Or myTerminalNumber = 0 Then
            
        'Increment number of players.
        bScores(0) = bScores(0) + 1
        
        'cntr controlled this player.
        ReDim Preserve bScores(17 * (cntr + 1) + 1) As Byte
        bScores(lPosition) = cntr + 1
        lPosition = lPosition + 1
        
        'Stats.
        Call IntToByte(gPlayerStats(cntr + 1).CardsTraded, bHold())
        Call CopyBytes(bScores, bHold, lPosition)
        lPosition = lPosition + 2
        
        Call IntToByte(gPlayerStats(cntr + 1).CountriesDefeated, bHold())
        Call CopyBytes(bScores, bHold, lPosition)
        lPosition = lPosition + 2
        
        Call IntToByte(gPlayerStats(cntr + 1).CountriesLost, bHold())
        Call CopyBytes(bScores, bHold, lPosition)
        lPosition = lPosition + 2
        
        Call IntToByte(gPlayerStats(cntr + 1).PlayersWipedOut, bHold())
        Call CopyBytes(bScores, bHold, lPosition)
        lPosition = lPosition + 2
        
        'If net.playerOwner(Plr - 1) = myTerminalNumber Then
        Call IntToByte(gPlayerStats(cntr + 1).PlrController, bHold())
        Call CopyBytes(bScores, bHold, lPosition)
        lPosition = lPosition + 2
        
        Call IntToByte(gPlayerStats(cntr + 1).UnitsBeaten, bHold())
        Call CopyBytes(bScores, bHold, lPosition)
        lPosition = lPosition + 2
        
        Call IntToByte(gPlayerStats(cntr + 1).UnitsFromCards, bHold())
        Call CopyBytes(bScores, bHold, lPosition)
        lPosition = lPosition + 2
        
        Call IntToByte(gPlayerStats(cntr + 1).UnitsIssued, bHold())
        Call CopyBytes(bScores, bHold, lPosition)
        lPosition = lPosition + 2
        
        Call IntToByte(gPlayerStats(cntr + 1).UnitsLost, bHold())
        Call CopyBytes(bScores, bHold, lPosition)
        lPosition = lPosition + 2
            
        'End If
    Next
    Call CopyBytes(ByteMessage, bScores, UBound(ByteMessage) + 1)
End Sub

'Unpack stats.
Private Function UnpackTerminalScore(ByteMessage() As Byte, Indx As Long) As Long
    Dim bHold() As Byte
    Dim cntr As Long
    Dim lPosition As Long
    Dim NoStats As Long
    Dim Plr As Long
    
    lPosition = Indx
    
    'Get number of Stats.
    NoStats = ByteMessage(lPosition)
    lPosition = lPosition + 1
    
    For cntr = 0 To NoStats - 1
        
        'Get player number.
        Plr = CLng(ByteMessage(lPosition))
        lPosition = lPosition + 1
        
        gPlayerStats(Plr).CardsTraded = ByteToInt(ByteMessage, lPosition)
        lPosition = lPosition + 2
        
        gPlayerStats(Plr).CountriesDefeated = ByteToInt(ByteMessage, lPosition)
        lPosition = lPosition + 2
        
        gPlayerStats(Plr).CountriesLost = ByteToInt(ByteMessage, lPosition)
        lPosition = lPosition + 2
        
        gPlayerStats(Plr).PlayersWipedOut = ByteToInt(ByteMessage, lPosition)
        lPosition = lPosition + 2
        
        If net.playerOwner(Plr - 1) <> myTerminalNumber Then
            gPlayerStats(Plr).PlrController = ByteToInt(ByteMessage, lPosition)
        End If
        lPosition = lPosition + 2
        
        gPlayerStats(Plr).UnitsBeaten = ByteToInt(ByteMessage, lPosition)
        lPosition = lPosition + 2
        
        gPlayerStats(Plr).UnitsFromCards = ByteToInt(ByteMessage, lPosition)
        lPosition = lPosition + 2
        
        gPlayerStats(Plr).UnitsIssued = ByteToInt(ByteMessage, lPosition)
        lPosition = lPosition + 2
        
        gPlayerStats(Plr).UnitsLost = ByteToInt(ByteMessage, lPosition)
        lPosition = lPosition + 2
    Next
    
    UnpackTerminalScore = lPosition
End Function

'List player controllers, only modifying those on this terminal.
'Allowes players to be claimed mid-way through a game.
Public Sub ListPlayerControllers()
    Dim cntr As Long
    
    For cntr = 0 To 5
        If net.playerOwner(cntr) = myTerminalNumber Then
            gPlayerStats(cntr + 1).PlrController = gPlayerID(cntr + 1).playerWho
        End If
    Next
End Sub

'Pack current situation for screen refresh
'Called at the start of the turn.
Public Sub PackForNetRefresh(pByteMessage() As Byte)
    Dim vByteBuf() As Byte
    Dim vIndex As Long
    Dim vIndex2 As Long
    Dim vIndex3 As Integer
    Dim DiceHold As Long
    Dim BMsize As Long
    
    'Size of the first section of the buffer. Dynamic info such as
    'stats, pctInfo text and the shadow audit.
    BMsize = 170
    
    net.madeUpdate = True
    ReDim pByteMessage(BMsize) As Byte
    
    'Countries
    For vIndex = 0 To 41
        pByteMessage(vIndex + 2) = CByte(gCountryOwner(vIndex + 1))
        pByteMessage(vIndex + 44) = getLowerByte(gCtryScore(vIndex + 1))
        pByteMessage(vIndex + 86) = getUpperByte(gCtryScore(vIndex + 1))
    Next
    
    'Each players' cards.
    For vIndex = 0 To 5
        For vIndex2 = 0 To 4
            pByteMessage(128 + (vIndex * 5) + vIndex2) = _
            CByte((gPlayerID(vIndex + 1).card(vIndex2 * 2)) + _
            (gPlayerID(vIndex + 1).card(vIndex2 * 2 + 1) * 16))
        Next
    Next
    
    pByteMessage(158) = CByte(gPlayerTurn)
    
    'Reshuffle notice bit 4
    If net.reshuffleCards Then
        pByteMessage(158) = pByteMessage(158) Or 8
        net.reshuffleCards = False
    End If
    pByteMessage(159) = CByte(gCurrentMode)
    pByteMessage(160) = CByte(gTargetCtry)
    pByteMessage(161) = CByte(gSourceCtry)
    
    pByteMessage(162) = CByte(Abs(netMain.optRefresh(1).Value) _
        + Abs(netMain.optRefresh(2).Value) * 2)
        
    'Players with mission to hold 18 countries
    For vIndex3 = 1 To 6
        Call SetBit((gPlayerID(vIndex3).mission = 14), vIndex3 + 1, pByteMessage(162))
    Next
    
    'Card deck.
    pByteMessage(163) = CByte(gCurrentCardValue)
    pByteMessage(164) = CByte(gCardDeck(0) + gCardDeck(1) * 16)
    pByteMessage(165) = CByte(gCardDeck(2) + gCardDeck(3) * 16)
    pByteMessage(166) = gWinMemoryBits
    
    'Only send dice info if refresh rate set to high (0).
    'Dice info is sent as octals in two bytes.
    If netMain.terminalSpeed = 0 Then
        'Attack dice.
        DiceHold = net.RolledAttackDice(0) _
                 + net.RolledAttackDice(1) * &H8 _
                 + net.RolledAttackDice(2) * &H40 _
                 + net.RolledAttackDice(3) * &H200 _
                 + net.RolledAttackDice(4) * &H1000
        pByteMessage(167) = CByte(DiceHold And &HFF)
        pByteMessage(168) = CByte((DiceHold And &HFF00) \ &H100)
        
        'Defence dice.
        DiceHold = net.RolledDefenceDice(0) _
                 + net.RolledDefenceDice(1) * &H8 _
                 + net.RolledDefenceDice(2) * &H40 _
                 + net.RolledDefenceDice(3) * &H200 _
                 + net.RolledDefenceDice(4) * &H1000
        pByteMessage(169) = CByte(DiceHold And &HFF)
        pByteMessage(170) = CByte((DiceHold And &HFF00) \ &H100)
    Else
        pByteMessage(167) = 0
        pByteMessage(168) = 0
        pByteMessage(169) = 0
        pByteMessage(170) = 0
    End If
    
    'Update stats with player controller.
    Call ListPlayerControllers
    
    'Pack stats.
    Call PackTerminalScore(pByteMessage)
    
    'Pack pctInto text. Ends with 255
    Call appendByteArray(pByteMessage, net.pctInfoByt)
    
    'Append audit shadow string. Ends with "X"
    Call AuditShadowAppend                                  'Call just in case.
    Call appendByteArray(pByteMessage, StrConv(AuditShadow & "X", vbFromUnicode))
    pByteMessage(UBound(pByteMessage)) = 255
    
    'Send player stats valid or not for each player.
    ReDim vByteBuf(6) As Byte
    For vIndex = 0 To 5
        vByteBuf(vIndex) = CByte(gPlayerStats(vIndex + 1).IsValid) And 128
    Next
    Call appendByteArray(pByteMessage, vByteBuf)
    pByteMessage(UBound(pByteMessage)) = 255                                  'Must end with 255
    
    '------------------------------------------------------------------------
    'This bit is not needed as cards are processed above however it shows
    'clearly how to append extra stuff to the refresh. Just uncomment below
    'and it works. It is backwards compatible with earlier versions of MR
    'because it is ignored.
    
    'ReDim vByteBuf(3) As Byte
    'vByteBuf(0) = CByte(gCardDeck(0))
    'vByteBuf(1) = CByte(gCardDeck(1))
    'vByteBuf(2) = CByte(gCardDeck(2))
    'vByteBuf(3) = CByte(gCardDeck(3))
    'Call appendByteArray(pByteMessage, vByteBuf)
    '------------------------------------------------------------------------
    Call AuditPrintClemsCounter
    Call netMain.ClearForfeitVotes
End Sub

'Unpack current situation to screen, refresh only if required.
'Called at the start of the turn. Packed by function PackForNetRefresh()
Public Sub UnpackNetRefresh(ByteMessage() As Byte, pTerminalIndex As Integer)
    Dim vIndex As Long, x As Long
    Dim vNeedRefresh As Boolean
    Dim vTemp As Integer
    Dim vIndex2 As Long
    Dim vPlayer As Integer
    Dim DiceHold As Long
    Dim BMsize As Long
    Dim vAttackDice(cMaxNumberOfDice - 1) As Integer
    Dim vDefenceDice(cMaxNumberOfDice - 1) As Integer
    
    'Size of the first section of the buffer. Dynamic info such as
    'stats, pctInfo text and the shadow audit.
    BMsize = 170
    
    'Clean up
    If frmMissions.Visible Then
        frmMissions.Hide
    End If
    
    'Do not accept the refresh if the current player is owned by this terminal.
    'If net.playerOwner(gPlayerTurn - 1) = myTerminalNumber And Not SetupScreen.Visible Then
    '    Exit Sub
    'End If
    
    'Country owners and scores.
    vNeedRefresh = False
    For vIndex = 0 To 41
        vTemp = CInt(ByteMessage(vIndex + 2))
        If gCountryOwner(vIndex + 1) <> vTemp Then
            gCountryOwner(vIndex + 1) = vTemp
            vNeedRefresh = True
        End If
        vTemp = CInt(ByteMessage(vIndex + 44) + ByteMessage(vIndex + 86) * 256)
        If gCtryScore(vIndex + 1) <> vTemp Then
            gCtryScore(vIndex + 1) = vTemp
            vNeedRefresh = True
        End If
    Next
    
    'Each players' cards.
    For vIndex = 0 To 5
        For x = 0 To 4
            vTemp = CInt(ByteMessage(128 + (vIndex * 5) + x) And 15) 'Odd cards
            If gPlayerID(vIndex + 1).card(x * 2) <> vTemp Then
                gPlayerID(vIndex + 1).card(x * 2) = vTemp
                vNeedRefresh = True
            End If
            vTemp = CInt(ByteMessage(128 + (vIndex * 5) + x) \ 16) 'Even cards
            If gPlayerID(vIndex + 1).card(x * 2 + 1) <> vTemp Then
                gPlayerID(vIndex + 1).card(x * 2 + 1) = vTemp
                vNeedRefresh = True
            End If
        Next
    Next
    
    'Player turn, current mode, target country, source country.
    gPlayerTurn = CInt(ByteMessage(158) And 7)
    
    If gCurrentMode <> 100 Then
        gCurrentMode = CInt(ByteMessage(159))
    End If
    
    gTargetCtry = CInt(ByteMessage(160))
    gSourceCtry = CInt(ByteMessage(161))
    
    'Set network refresh rate assigned by the host (bit 0-1).
    If netWorkSituation = cNetClient Then
        Select Case (ByteMessage(162) And 3)
        Case 0
            netMain.optRefresh(0).Value = 1
        Case 1
            netMain.optRefresh(1).Value = 1
        Case 2
            netMain.optRefresh(2).Value = 1
        End Select
    End If
    
    'Players with mission to hold 18 countries (bit 2-7).
    For vPlayer = 1 To 6
        If GetBit(vPlayer + 1, ByteMessage(162)) Then
            gPlayerID(vPlayer).mission = 14
        End If
    Next
    
    'Cards.
    gCurrentCardValue = CInt(ByteMessage(163))
    gCardDeck(0) = CInt(ByteMessage(164) And 15)
    gCardDeck(1) = CInt(ByteMessage(164) \ 16)
    gCardDeck(2) = CInt(ByteMessage(165) And 15)
    gCardDeck(3) = CInt(ByteMessage(165) \ 16)
    gWinMemoryBits = ByteMessage(166)
    
    'Attack dice.
    DiceHold = CLng(ByteMessage(167)) + (CLng(ByteMessage(168)) * &H100)
    vAttackDice(0) = CByte(DiceHold And &H7)
    vAttackDice(1) = CByte((DiceHold \ &H8) And &H7)
    vAttackDice(2) = CByte((DiceHold \ &H40) And &H7)
    vAttackDice(3) = CByte((DiceHold \ &H200) And &H7)
    vAttackDice(4) = CByte((DiceHold \ &H1000) And &H7)
    
    'Defence dice.
    DiceHold = CLng(ByteMessage(169)) + (CLng(ByteMessage(170)) * &H100)
    vDefenceDice(0) = CByte(DiceHold And &H7)
    vDefenceDice(1) = CByte((DiceHold \ &H8) And &H7)
    vDefenceDice(2) = CByte((DiceHold \ &H40) And &H7)
    vDefenceDice(3) = CByte((DiceHold \ &H200) And &H7)
    vDefenceDice(4) = CByte((DiceHold \ &H1000) And &H7)
    
    'Show dice.
    If DiceHold > 0 Then
        Call DisplayDiceOnBoard(vAttackDice, vDefenceDice)
    End If
    
    'ByteMessage(167) contains count of player stats.
    'Player stats take 36 bytes per player + 1 byte for count.
    BMsize = UnpackTerminalScore(ByteMessage, BMsize + 1) - 1
    
    'Top bit is set to reshuffle notice.
    If ByteMessage(158) > 7 Then
        Call PutCardsBack
    End If
    
    If SetupScreen.Visible And netWorkSituation = cNetClient Then
        Call AssignPlayerIDs
    End If
    
    'Flash color if human on this terminal.
    TheMainForm.pctInfoBox.BackColor = gPlayerID(gPlayerTurn).bkgndColor
    If flashingBorder And (gPlayerID(gPlayerTurn).playerWho = 0) Then
        If netWorkSituation <> cNetNone Then
            If (net.playerOwner(gPlayerTurn - 1) = myTerminalNumber) Then
                TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
                tmrFlashInfoBox.Enabled = True
                ColorCountryUnderAttack 0
            Else
                TheMainForm.BackColor = &H8000000F
            End If
        Else
            TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
            tmrFlashInfoBox.Enabled = True
        End If
    Else
        TheMainForm.BackColor = &H8000000F
    End If
    
    'pctInfo text.
    InfoBoxPrint 0
    BMsize = BMsize + InfoBoxPrintBytes(ByteMessage, BMsize + 1) + 1
    
    'Shadow audit.
    BMsize = BMsize + AuditShadowBytes(ByteMessage, BMsize + 1) + 1
    
    If UBound(ByteMessage) >= BMsize + 6 Then
        For vIndex = 0 To 5
        
            'Get player stats valid or not from the host for each player except for me.
            If netWorkSituation = cNetClient Then
                If net.playerOwner(vIndex) <> myTerminalNumber Then
                    gPlayerStats(vIndex + 1).IsValid = CBool(ByteMessage(BMsize + vIndex + 1))
                    If gPlayerStats(vIndex + 1).IsValid Then
                        gPlayerStats(vIndex + 1).InvalidatedReason = "OK."
                    Else
                        gPlayerStats(vIndex + 1).InvalidatedReason = vbCrLf & "Stats invalidated."
                    End If
                End If
            
            'Get player stats valid or not from client if they own that player.
            Else
                If net.playerOwner(vIndex) = pTerminalIndex And gPlayerStats(vIndex + 1).IsValid Then
                    gPlayerStats(vIndex + 1).IsValid = CBool(ByteMessage(BMsize + vIndex + 1))
                    If Not gPlayerStats(vIndex + 1).IsValid Then
                        gPlayerStats(vIndex + 1).InvalidatedReason = vbCrLf & "Stats invalidated."
                    End If
                End If
            End If
        Next
        
        BMsize = BMsize + 6
    End If
    
    '------------------------------------------------------------------------
    'This bit is not needed as cards are processed above however it shows
    'clearly how to append extra stuff to the refresh. Just uncomment below
    'and it works. It is backwards compatible with earlier versions of MR
    'because it is ignored.
    
    'Get card info if available.
    'If UBound(ByteMessage) <> BMsize + 1 Then
    '    gCardDeck(0) = CInt(ByteMessage(BMsize + 1))
    '    gCardDeck(1) = CInt(ByteMessage(BMsize + 2))
    '    gCardDeck(2) = CInt(ByteMessage(BMsize + 3))
    '    gCardDeck(3) = CInt(ByteMessage(BMsize + 4))
    'End If
    '------------------------------------------------------------------------
    
    If vNeedRefresh Then
        Call refreshMap
    Else
        Call DrawLittleCards
    End If
    Call CheckCards
    Call resetChangeList
    Call CheckToSeeMission(gPlayerTurn)
    tfRate1.Value = True
    transferNmbr = 1
    gPlayerValue = GetPlayerValue(gPlayerTurn)
    Call AutoPlayerSelect
    If gPlayerTurn = 0 Then gPlayerTurn = 1
    gPlayerTurn = gPlayerTurn - 1
    Call SaveCheckpoint
    gPlayerTurn = gPlayerTurn + 1
    mnuOptUndo.Enabled = gCheatMode.undoEnabled
    Toolbar1.Buttons(8).Enabled = gCheatMode.undoEnabled
    Call SyncForgroundMap("UnpackNetRefresh")
    
    'Audit record snapshot at start of turn.
    Call AuditPlayerRecord
    Call AuditAddPointsIssued(gPlayerTurn, gPlayerValue)
    Call AuditPrintClemsCounter
    Call netMain.ClearForfeitVotes
End Sub

'Decode and print the text directives from the passed byte array starting at
'the passed index. Return the index of the last code. The end directive is &H7F.
Private Function InfoBoxPrintBytes(pEncodedText() As Byte, pStartIndex As Long) As Long
    Dim vIndex As Long
    Dim vDirectiveLength As Long
    
    For vIndex = pStartIndex To UBound(pEncodedText)
        vDirectiveLength = InfoBoxGetDirective(pEncodedText, vIndex)
        If vDirectiveLength = 0 Then
            InfoBoxPrintBytes = vIndex - pStartIndex
            Exit Function
        End If
        vIndex = vIndex + vDirectiveLength - 1
    Next
    DoEvents
End Function

'Stop array bound overflow towards end of print code.
Private Function InfoBoxGetDirective(pEncodedText() As Byte, pIndex As Long) As Long
    Dim vMaxDirectiveLength As Long
    
    vMaxDirectiveLength = UBound(pEncodedText) - pIndex
    
    If vMaxDirectiveLength = 0 Then
        InfoBoxGetDirective = InfoBoxPrintDirective(pEncodedText(pIndex))
    ElseIf vMaxDirectiveLength = 1 Then
        InfoBoxGetDirective = InfoBoxPrintDirective(pEncodedText(pIndex), _
                            pEncodedText(pIndex + 1))
    Else
        InfoBoxGetDirective = InfoBoxPrintDirective(pEncodedText(pIndex), _
                            pEncodedText(pIndex + 1), _
                            pEncodedText(pIndex + 2))
    End If
End Function

'Print text to the center of the passed picture box.
Private Sub CenterPrintText(pDestPctBox As PictureBox, pText As String)
    Dim vPrintY As Long
    Dim vPrintX As Long
    Dim vTextLine() As String
    Dim vIndex As Long
    
    On Error Resume Next
    
    'Knock off the last CrLf.
    If Right(pText, 2) = vbCrLf Then
        pText = Mid(pText, 1, Len(pText) - 2)
    End If
    
    vTextLine = Split(pText, vbCrLf)
    
    vPrintY = ((pDestPctBox.Height - pDestPctBox.TextHeight(pText)) / 2) - (pDestPctBox.Height / 10)
    pDestPctBox.CurrentY = vPrintY
    For vIndex = 0 To UBound(vTextLine)
        vPrintX = (pDestPctBox.Width - pDestPctBox.TextWidth(vTextLine(vIndex))) / 2
        pDestPctBox.CurrentX = vPrintX
        pDestPctBox.Print Trim(vTextLine(vIndex))
    Next
End Sub

'Decode the passed directive and print as required. Return the length of the directive.
'This is used to format and print text to the info box and is called directly for local
'wars and from the networking system. Commands that have the top bit set (&H80) have a
'crlf appended after being printed.
'Directives:
'0  -   Clear screen
'1  -   Print phrase(pArg1)
'2  -   Print the number pArg1 low byte and pArg2 high byte
'3  -   Space pArg1 times
'4  -   Period "."
'5  -   Set bold text for the next directive
'6  -   Set normal text for the next directive
'7  -   Carriage return
'8  -   Print country name(pArg1)
'9  -   Print army name(pArg1)
'10 -   Save current font size in the static variable sFontSize and set the font size * 1.5
'       and set to print at center
'11 -   Restore font size saved in the static variable sFontSize and set to left justify
'127 -  End of all directives
'And &H80 - Top bit set will cause a carriage return to be appended the end of the directive
Private Function InfoBoxPrintDirective(pCommand As Byte, _
Optional pArg1 As Byte = 0, Optional pArg2 As Byte = 0) As Long
    Dim vIndex As Long
    Dim vAddCrLf As Boolean
    Static sFontSize As Single
    Static sPrintAtCenter As Boolean
    
    InfoBoxPrintDirective = 0
    
    'Strip of the crlf directive and remember for action after
    'the directive has been actioned.
    vAddCrLf = (pCommand > &H7F)
    pCommand = pCommand And &H7F
    
    Select Case pCommand
    
    'Clear screen.
    Case 0
        TheMainForm.pctInfoBox.Cls
        TheMainForm.pctInfoBox.Print "";
        TheMainForm.pctInfoBox.Tag = ""
        InfoBoxPrintDirective = 1
    
    'Print Phrase(pArg1).
    Case 1
        If sPrintAtCenter Then
            pctInfoBox.Cls
            pctInfoBox.Print "";
            Call CenterPrintText(TheMainForm.pctInfoBox, Phrase(CLng(pArg1)))
        Else
            TheMainForm.pctInfoBox.Print Phrase(CLng(pArg1));
        End If
        TheMainForm.pctInfoBox.Tag = TheMainForm.pctInfoBox.Tag & Phrase(CLng(pArg1))
        InfoBoxPrintDirective = 2
    
    'Print number pArg1 low byte and pArg2 high byte.
    Case 2
        TheMainForm.pctInfoBox.Print Trim(str(pArg1 + pArg2 * 256));
        TheMainForm.pctInfoBox.Tag = TheMainForm.pctInfoBox.Tag & Trim(str(pArg1 + pArg2 * 256))
        InfoBoxPrintDirective = 3
    
    'Space pArg1 times.
    Case 3
        TheMainForm.pctInfoBox.Print Space(pArg1);
        TheMainForm.pctInfoBox.Tag = TheMainForm.pctInfoBox.Tag & Space(pArg1)
        InfoBoxPrintDirective = 2
    
    'Period.
    Case 4
        TheMainForm.pctInfoBox.Print ".";
        TheMainForm.pctInfoBox.Tag = TheMainForm.pctInfoBox.Tag & "."
        InfoBoxPrintDirective = 1
    
    'Set bold text for the next directive.
    Case 5
        TheMainForm.pctInfoBox.Font.Bold = True
        InfoBoxPrintDirective = 1
    
    'Set normal text for the next directive.
    Case 6
        TheMainForm.pctInfoBox.Font.Bold = False
        InfoBoxPrintDirective = 1
    
    'Carriage return.
    Case 7
        TheMainForm.pctInfoBox.Print
        TheMainForm.pctInfoBox.Tag = TheMainForm.pctInfoBox.Tag & vbCrLf
        InfoBoxPrintDirective = 1
    
    'Print country name(pArg1).
    Case 8
        TheMainForm.pctInfoBox.Print Trim(CountryID(CLng(pArg1)).ctryName);
        TheMainForm.pctInfoBox.Tag = TheMainForm.pctInfoBox.Tag & Trim(CountryID(CLng(pArg1)).ctryName)
        InfoBoxPrintDirective = 2
    
    'Print army name(pArg1).
    Case 9
        TheMainForm.pctInfoBox.Print Trim(gPlayerID(CLng(pArg1)).strColor);
        TheMainForm.pctInfoBox.Tag = TheMainForm.pctInfoBox.Tag & Trim(gPlayerID(CLng(pArg1)).strColor)
        InfoBoxPrintDirective = 2
    
    'Save current font size in the static variable sFontSize
    'and set the font size * 1.5 and set to print at center.
    Case 10
        sFontSize = TheMainForm.pctInfoBox.Font.Size
        TheMainForm.pctInfoBox.Font.Size = sFontSize * 1.5
        TheMainForm.pctInfoBox.Tag = TheMainForm.pctInfoBox.Tag & "<FONT15>"
        sPrintAtCenter = True
        InfoBoxPrintDirective = 1
    
    'Restore font size saved in the static variable sFontSize
    'and set to left justify.
    Case 11
        TheMainForm.pctInfoBox.Font.Size = sFontSize
        sPrintAtCenter = False
        InfoBoxPrintDirective = 1
    
    'End.
    Case 127
        InfoBoxPrintDirective = 0
    End Select
    
    'Add crlf as directed.
    If vAddCrLf Then
        TheMainForm.pctInfoBox.Print ""
        TheMainForm.pctInfoBox.Tag = TheMainForm.pctInfoBox.Tag & vbCrLf
        pCommand = pCommand + 128
    End If
End Function

'Encode and print infobox text with a carrage return appended.
Public Sub InfoBoxPrnCR(pCommand As Byte, Optional pArg As Integer = 0)
    Call InfoBoxPrint(pCommand + 128, pArg)
End Sub

'Encode and print infobox text and add to the network update byte array.
Public Sub InfoBoxPrint(pCommand As Byte, Optional pArg As Integer = 0)
    Dim vArg1 As Byte
    Dim vArg2 As Byte
    Dim vPointer As Long
    Dim vCmdDecoded As Byte
    
    On Error GoTo ErrHand
    vPointer = UBound(net.pctInfoByt)
    ReDim Preserve net.pctInfoByt(vPointer + 1) As Byte
    net.pctInfoByt(vPointer) = pCommand
    
    vCmdDecoded = pCommand And 127
    
    'Clear screen.
    If vCmdDecoded = 0 Then
        ReDim net.pctInfoByt(1) As Byte
        net.pctInfoByt(0) = 0
    
    'Phrase(pArg).
    ElseIf vCmdDecoded = 1 Then
        vArg1 = CByte(pArg)
        ReDim Preserve net.pctInfoByt(vPointer + 2) As Byte
        net.pctInfoByt(vPointer + 1) = vArg1
    
    'Numeral.
    ElseIf vCmdDecoded = 2 Then
        vArg1 = getLowerByte(pArg)
        vArg2 = getUpperByte(pArg)
        ReDim Preserve net.pctInfoByt(vPointer + 3) As Byte
        net.pctInfoByt(vPointer + 1) = vArg1
        net.pctInfoByt(vPointer + 2) = vArg2
    
    'Space(pArg).
    ElseIf vCmdDecoded = 3 Then
        vArg1 = CByte(pArg)
        ReDim Preserve net.pctInfoByt(vPointer + 2) As Byte
        net.pctInfoByt(vPointer + 1) = vArg1
    
    'Country(pArg).
    ElseIf vCmdDecoded = 8 Then
        vArg1 = CByte(pArg)
        ReDim Preserve net.pctInfoByt(vPointer + 2) As Byte
        net.pctInfoByt(vPointer + 1) = vArg1
    
    'Army name.
    ElseIf vCmdDecoded = 9 Then
        vArg1 = CByte(pArg)
        ReDim Preserve net.pctInfoByt(vPointer + 2) As Byte
        net.pctInfoByt(vPointer + 1) = vArg1
    End If
    
    net.pctInfoByt(UBound(net.pctInfoByt)) = 255
    Call InfoBoxPrintDirective(pCommand, vArg1, vArg2)
    Exit Sub
ErrHand:
    Resume Next
End Sub

    'Return lower byte
Private Function getLowerByte(lByte As Integer) As Byte
    getLowerByte = CByte(lByte And 255)
End Function

    'Return upper byte
Private Function getUpperByte(uByte As Integer) As Byte
    On Error GoTo ErrHand
    getUpperByte = CByte(uByte \ 256)
    Exit Function
ErrHand:
    getUpperByte = 255
    Exit Function
End Function

    'Return upper and lower bytes together
Private Function restoreInt(lByte As Byte, uByte As Byte) As Integer
    restoreInt = CInt(lByte + uByte * 256)
End Function

'I am the host. Append a list of player names to the passed byte array.
'Format: Length, Term_Number, Term_Name, Length, Term_Number, Term_Name,....
'GetPlayerOwners() is the counterpart.
'The keywords "PHRASE:" and "NAME:" are used to
'work out if the following is the phrase number for unclaimed
'or the actual name of the terminal. This needs to be done because
'if the host terminal is using a different language to the client
'terminals, they will see "Unclaimed" in the host's language.
Private Sub AppendPlayerOwners(ByteMessage() As Byte, Index As Long)
    Dim vIndex As Long
    Dim vTermName As String
    
    For vIndex = 0 To 5
    
        'Get the owner of this player's terminal name.
        If SetupScreen.Visible Then
            If playerSelect_getIndex(CInt(vIndex)) = remoteIndex _
            And net.playerOwner(vIndex) = myTerminalNumber Then
                
                'Player unclaimed. The keywords "PHRASE:" and "NAME:" are used for
                'internationalisation.
                vTermName = "PHRASE:28"         'Phrase(28)  '"Unclaimed"
            Else
                vTermName = "NAME:" & net.ClientName(net.playerOwner(vIndex))
            End If
        Else
            If gPlayerID(vIndex + 1).playerWho = remoteIndex _
            And net.playerOwner(vIndex) = myTerminalNumber Then
                
                'Player unclaimed.
                vTermName = "PHRASE:28"         'Phrase(28)  '"Unclaimed"
            Else
                vTermName = "NAME:" & net.ClientName(net.playerOwner(vIndex))
            End If
        End If
        
        'Append the terminal name.
        '**This can be done better outside of the loop.
        ReDim Preserve ByteMessage(Index + Len(vTermName) + 2)
        ByteMessage(Index) = CByte(Len(vTermName) + 2)
        ByteMessage(Index + 1) = CByte(net.playerOwner(vIndex))
        CopyBytes ByteMessage, StrConv(vTermName, vbFromUnicode), Index + 2
        Index = Index + Len(vTermName) + 2
    Next
End Sub

'Read and set player names from byte array.
'Format: Length, Term_Number, Term_Name, Length, Term_Number, Term_Name,....
'If I own that player, change remote name to "Remote Player".
'The keywords "PHRASE:" and "NAME:" are used to
'work out if the following is the phrase number for unclaimed
'or the actual name of the terminal. This needs to be done because
'if the host terminal is using a different language to the client
'terminals, they will see "Unclaimed" in the host's language.
Private Function GetPlayerOwners(ByteMessage() As Byte, ByRef Index As Long) As Long
    Dim vIndex As Long
    Dim vPlayerOwner As String
    
    On Error Resume Next
    
    For vIndex = 0 To 5
        If ByteMessage(Index + 1) <> myTerminalNumber Then
            vPlayerOwner = StrConv(MidB(ByteMessage, Index + 3, ByteMessage(Index) - 2), vbUnicode)
            If vPlayerOwner = "PHRASE:28" Then
                vPlayerOwner = Phrase(28)
            Else
                vPlayerOwner = Mid(vPlayerOwner, Len("NAME:") + 1)
            End If
            
            Call RenameRemote(CByte(vIndex), vPlayerOwner)
            net.playerOwner(vIndex) = ByteMessage(Index + 1)
        Else
            RenameRemote CByte(vIndex), Phrase(28)                'Remote Player.
        End If
        Index = Index + CLng(ByteMessage(Index))
    Next
    GetPlayerOwners = Index
End Function

'Pack sutup screen into the passed byte array for dispatch to remote
'terminals via command 9. Counterpart is UnpackSetupScreen().
'2 to 7 - Starting countries 'udPlayerStartCountries(0-5).Value'
'8 to 13 - Army controller 'playerSelect_getIndex(vIndex - 2)'
'14 - Bit encoded options
'15 - Bit encoded options
'16 - Bit encoded options
'17 - Maximum card value 'udMaximumCardValue.Value'
'18 - Number of starting armies 'udStartingArmies.Value'
'19 - Extra starting units 'udExtraStartingUnits.Value'
'20 - Bit encodec claimable army list
'21 - Bit encoded dice rules
'22 - Bit encoded dice rules
'23 to 24 - Number of dice to throw 'udDiceThrown(0-1).Value'
'25 to 28 - Number of cards in each deck 'udCardDeck(0-3)'
'29 to 32 - Fixed card values for each valid combination 'udFixedValues(0-3)'
'33 to 35 - New unit calculations 'udNewUnitClac(0-2)'
'36 to 41 - Continent values 'udContVal(0-5)'
'42 appended - Terminal names of claimed players.
'appended - War name.
Public Sub PackSetupScreen(ByteMessage() As Byte)
    Dim vIndex As Long
    
    ReDim ByteMessage(42) As Byte
    
    For vIndex = 2 To 7
        ByteMessage(vIndex) = CByte(udPlayerStartCountries(vIndex - 2).Value)
        ByteMessage(vIndex + 6) = CByte(playerSelect_getIndex(vIndex - 2))
        SetBit ((playerSelect_getIndex(vIndex - 2) <> 3) Or (net.playerOwner(vIndex - 2) <> 0)), vIndex - 2, ByteMessage(20)
    Next
    
    ByteMessage(14) = 0
    SetBit CBool(chkCardsVulture.Value), 0, ByteMessage(14)
    SetBit CBool(chkCardsHidden.Value), 1, ByteMessage(14)
    SetBit CBool(optPlr1FirstPlayer.Value), 2, ByteMessage(14)
    'setBit CBool(chkFastDice.Value), 3, ByteMessage(14)
    SetBit CBool(optnFastWar.Checked), 4, ByteMessage(14)
    SetBit CBool(chkMsnMissionsOn.Value), 5, ByteMessage(14)
    'setBit CBool(chkOptimizeDefenceDice.Value), 6, ByteMessage(14)
    SetBit CBool(optSupplyLines.Value), 7, ByteMessage(14)
    
    ByteMessage(15) = 0
    SetBit CBool(optLimitSupply.Value), 0, ByteMessage(15)
    SetBit CBool(optNoSupply.Value), 1, ByteMessage(15)
    'setBit CBool(chkBorder.Value), 2, ByteMessage(15)
    SetBit CBool(chkExtraStartingUnits.Value), 3, ByteMessage(15)
    SetBit CBool(optCardMode(0).Value), 4, ByteMessage(15)
    SetBit CBool(optCardMode(1).Value), 5, ByteMessage(15)
    SetBit CBool(optCardMode(2).Value), 6, ByteMessage(15)
    SetBit SetupScreen.Visible, 7, ByteMessage(15)
    
    ByteMessage(16) = 0 'CByte(gPreviousSettings.prevMode)
    SetBit chkMsnArmyWipeout.Value = vbChecked, 0, ByteMessage(16)
    SetBit chkMsnConquerHold.Value = vbChecked, 1, ByteMessage(16)
    SetBit chkMsnMustComplete.Value = vbChecked, 2, ByteMessage(16)
    SetBit chkMsnWinImmediate.Value = vbChecked, 3, ByteMessage(16)
    SetBit chkMsnAreUnique.Value = vbChecked, 4, ByteMessage(16)
    
    ByteMessage(17) = CByte(udMaximumCardValue.Value)
    ByteMessage(18) = CByte(udStartingArmies.Value)
    ByteMessage(19) = CByte(udExtraStartingUnits.Value)
    
    'Dice setup info.
    'Bit encoded dice rules.
    ByteMessage(21) = 0
    SetBit CBool(optDiceRules(0).Value), 0, ByteMessage(21)
    SetBit CBool(optDiceRules(1).Value), 1, ByteMessage(21)
    SetBit CBool(optDiceRules(2).Value), 2, ByteMessage(21)
    SetBit CBool(optDiceRules(3).Value), 3, ByteMessage(21)
    SetBit CBool(optDiceRules(4).Value), 4, ByteMessage(21)
    SetBit CBool(chkSortDice.Value), 5, ByteMessage(21)
    
    ByteMessage(22) = 0
    SetBit CBool(optDiceSame(0).Value), 0, ByteMessage(22)
    SetBit CBool(optDiceSame(1).Value), 1, ByteMessage(22)
    SetBit CBool(optDiceSame(2).Value), 2, ByteMessage(22)
    
    ByteMessage(23) = CByte(udDiceThrown(0).Value)
    ByteMessage(24) = CByte(udDiceThrown(1).Value)
    
    'Extra card details, deck suite numbers and fixed card values.
    '25 to 28 - Number of cards in each deck 'udCardDeck(0-3)'
    For vIndex = 0 To udCardDeck.Count - 1
        ByteMessage(vIndex + 25) = CByte(udCardDeck(vIndex).Value)
    Next
    
    '29 to 32 - Fixed card values for each valid combination 'udFixedValues(0-3)'
    For vIndex = 0 To udFixedValues.Count - 1
        ByteMessage(vIndex + 29) = CByte(udFixedValues(vIndex).Value)
    Next
    
    'Reinforcments tab.
    '33 to 35 - New unit calculations 'udNewUnitClac(0-2)'
    For vIndex = 0 To udNewUnitClac.Count - 1
        ByteMessage(vIndex + 33) = CByte(udNewUnitClac(vIndex).Value)
    Next
    
    '36 to 41 - Continent values 'udContVal(0-5)'
    For vIndex = 0 To udContVal.Count - 1
        ByteMessage(vIndex + 36) = CByte(udContVal(vIndex).Value)
    Next
    
    'Extra mission info can go here.
    'ByteMessage(42+) = CByte(udMissionSomething(0).Value)
    
    'Append the Terminal names of claimed players.
    'Format Length, Term_Number, Term_Name, Length, Term_Number, Term_Name,....
    Call AppendPlayerOwners(ByteMessage, 42)
    
    'Append the war name.
    Call appendByteArray(ByteMessage, StrConv(GetWarName, vbFromUnicode))
End Sub
    
'Unack sutup screen from passed byte array received from the host
'terminal via command 9. Counterpart is PackSetupScreen().
'2 to 7 - Starting countries 'udPlayerStartCountries(0-5).Value'
'8 to 13 - Army controller 'playerSelect_getIndex(vIndex - 2)'
'14 - Bit encoded options
'15 - Bit encoded options
'16 - Bit encoded options
'17 - Maximum card value 'udMaximumCardValue.Value'
'18 - Number of starting armies 'udStartingArmies.Value'
'19 - Extra starting units 'udExtraStartingUnits.Value'
'20 - Bit encodec claimable army list
'21 - Bit encoded dice rules
'22 - Bit encoded dice rules
'23 to 24 - Number of dice to throw 'udDiceThrown(0-1).Value'
'25 to 28 - Number of cards in each deck 'udCardDeck(0-3)'
'29 to 32 - Fixed card values for each valid combination 'udFixedValues(0-3)'
'33 to 35 - New unit calculations 'udNewUnitClac(0-2)'
'36 to 41 - Continent values 'udContVal(0-5)'
'42 appended - Terminal names of claimed players.
'appended - War name.
Public Sub UnpackSetupScreen(ByteMessage() As Byte)
    Dim vIndex As Long
    Dim vTmpNetSituation As Byte
    
    On Error Resume Next
    
    If frmMissions.Visible Then
        frmMissions.Hide
    End If
    
    tmpTimer1 = Timer1.Enabled
    tmpTimer2 = Timer2.Enabled
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    'Disable undo buttons.
    mnuOptUndo.Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    
    'Prevent resending changes when this function changes individual setup settings.
    vTmpNetSituation = netWorkSituation
    netWorkSituation = cNetNone
    
    'Starting countries and army controllers.
    For vIndex = 2 To 7
        'udPlayerStartCountries(vIndex - 2).Value = ByteMessage(vIndex)
        
        'Change starting country count without all the checking and callbacks.
        Call ChangPlayerStartingCtryCount(vIndex - 2, CInt(ByteMessage(vIndex)))
        
        'If I am a client terminal.
        If vTmpNetSituation = cNetClient Then
            
            'If this army is claimed by another terminal then disable
            'and mark as "Remote player".
            If GetBit(vIndex - 2, ByteMessage(20)) Then
                If net.playerOwner(vIndex - 2) <> myTerminalNumber Then
                    Call playerSelect_showIndex(vIndex - 2, remoteIndex)
                    Call PlayerSelect_Enable(vIndex - 2, False)
                Else
                    Call PlayerSelect_Enable(vIndex - 2, True)
                End If
            
            'Enable and mark as "Unclaimed".
            ElseIf ByteMessage(vIndex + 6) = remoteIndex Then
                Call PlayerSelect_Enable(vIndex - 2, True)
                Call playerSelect_showIndex(vIndex - 2, remoteIndex)
                net.playerOwner(vIndex - 2) = 0
            End If
        
        'If I am the host, remote terminal has disclaimed this player.
        '** needs checking.
        Else
            Call PlayerSelect_Enable(vIndex - 2, True)
            Call playerSelect_showIndex(vIndex - 2, CInt(ByteMessage(vIndex + 6)))
        End If
        
    Next
    
    'Sync the u/d control value with the player starting countries text boxes.
    Call SyncPlrStartingCountries
    
    chkCardsVulture.Value = Abs(GetBit(0, ByteMessage(14)))
    chkCardsHidden.Value = Abs(GetBit(1, ByteMessage(14)))
    optPlr1FirstPlayer.Value = Abs(GetBit(2, ByteMessage(14)))
    optRandomFirstPlayer.Value = Not optPlr1FirstPlayer.Value
    'chkFastDice.Value = Abs(getBit(3, ByteMessage(14)))
    optnFastWar.Checked = GetBit(4, ByteMessage(14))
    chkMsnMissionsOn.Value = Abs(GetBit(5, ByteMessage(14)))
    'chkOptimizeDefenceDice.Value = Abs(getBit(6, ByteMessage(14)))
    optSupplyLines.Value = Abs(GetBit(7, ByteMessage(14)))

    optLimitSupply.Value = Abs(GetBit(0, ByteMessage(15)))
    optNoSupply.Value = Abs(GetBit(1, ByteMessage(15)))
    'chkBorder.Value = Abs(getBit(2, ByteMessage(15)))
    chkExtraStartingUnits.Value = Abs(GetBit(3, ByteMessage(15)))
    optCardMode(0).Value = Abs(GetBit(4, ByteMessage(15)))
    optCardMode(1).Value = Abs(GetBit(5, ByteMessage(15)))
    optCardMode(2).Value = Abs(GetBit(6, ByteMessage(15)))
    
    chkMsnArmyWipeout.Value = Abs(CInt(GetBit(0, ByteMessage(16))))
    chkMsnConquerHold.Value = Abs(CInt(GetBit(1, ByteMessage(16))))
    chkMsnMustComplete.Value = Abs(CInt(GetBit(2, ByteMessage(16))))
    chkMsnWinImmediate.Value = Abs(CInt(GetBit(3, ByteMessage(16))))
    chkMsnAreUnique.Value = Abs(CInt(GetBit(4, ByteMessage(16))))
    
    'Max card value.
    udMaximumCardValue.Value = ByteMessage(17)
    'txtMaximumCardValue.Text = ByteMessage(17)
    
    'Number of starting armies.
    Call SetStartingArmyCount(CInt(ByteMessage(18)))
    
    'Distribute extra units.
    udExtraStartingUnits.Value = ByteMessage(19)
    
    'Bit encoded dice rules.
    optDiceRules(0).Value = GetBit(0, ByteMessage(21))
    optDiceRules(1).Value = GetBit(1, ByteMessage(21))
    optDiceRules(2).Value = GetBit(2, ByteMessage(21))
    optDiceRules(3).Value = GetBit(3, ByteMessage(21))
    optDiceRules(4).Value = GetBit(4, ByteMessage(21))
    chkSortDice.Value = Abs(GetBit(5, ByteMessage(21)))
    
    optDiceSame(0).Value = GetBit(0, ByteMessage(22))
    optDiceSame(1).Value = GetBit(1, ByteMessage(22))
    optDiceSame(2).Value = GetBit(2, ByteMessage(22))
    
    udDiceThrown(0).Value = ByteMessage(23)
    udDiceThrown(1).Value = ByteMessage(24)
    
    'Extra card details, deck suite numbers and fixed card values.
    udCardDeck(0).Value = ByteMessage(25)
    udCardDeck(1).Value = ByteMessage(26)
    udCardDeck(2).Value = ByteMessage(27)
    udCardDeck(3).Value = ByteMessage(28)
    udFixedValues(0).Value = ByteMessage(29)
    udFixedValues(1).Value = ByteMessage(30)
    udFixedValues(2).Value = ByteMessage(31)
    udFixedValues(3).Value = ByteMessage(32)
    
    'Extra card details, deck suite numbers and fixed card values.
    '25 to 28 - Number of cards in each deck 'udCardDeck(0-3)'
    For vIndex = 0 To udCardDeck.Count - 1
        udCardDeck(vIndex).Value = ByteMessage(vIndex + 25)
    Next
    
    '29 to 32 - Fixed card values for each valid combination 'udFixedValues(0-3)'
    For vIndex = 0 To udFixedValues.Count - 1
        udFixedValues(vIndex).Value = ByteMessage(vIndex + 29)
    Next
    
    'Reinforcments tab.
    '33 to 35 - New unit calculations 'udNewUnitClac(0-2)'
    For vIndex = 0 To udNewUnitClac.Count - 1
        udNewUnitClac(vIndex).Value = ByteMessage(vIndex + 33)
    Next
    
    '36 to 41 - Continent values 'udContVal(0-5)'
    For vIndex = 0 To udContVal.Count - 1
        udContVal(vIndex).Value = ByteMessage(vIndex + 36)
    Next
    
    'Extra mission info can go here.
    'ByteMessage(42+) = CByte(udMissionSomething(0).Value)
    
    netWorkSituation = vTmpNetSituation
    
    'If SetupScreen.Visible then hide the Setup Screen.
    If Not GetBit(7, ByteMessage(15)) Then
        gCurrentMode = 100
        Call EnableSetupControls(False)
        SetupScreen.Visible = True
        Call ShowMenuBar(SetupScreen.Visible)
    End If
    
    'Names of remote terminal player owners.
    vIndex = GetPlayerOwners(ByteMessage, 42)
    
    'Get the rest of the byte array.
    'Call GetRestOfByte(ByteMessage, vIndex + 1)
    
    'Set the war name.
    'Call makeWarName(StrConv(ByteMessage, vbUnicode))
    Call makeWarName(StrConv(MidB(ByteMessage, vIndex + 2), vbUnicode))
    
    'House keeping.
    SetupScreen.Visible = True
    Call ShowMenuBar(SetupScreen.Visible)
    TheMainForm.BackColor = &H8000000F
    net.setupControlChange = False
End Sub

    'Join a war which is in progress
Private Sub joinWar()
    Dim i As Long
    
    cmdSetupOk.Caption = Phrase(26)     'Declare war
    cmdSetupOk.Enabled = False
    mnuDeclareWar.Caption = cmdSetupOk.Caption
    mnuDeclareWar.Enabled = cmdSetupOk.Enabled
    SetupScreen.Visible = False
    Call ShowMenuBar(SetupScreen.Visible)
    gCurrentMode = 2
    'Call netMain.requestRefresh
    Call netMain.requestWar
    gWinMemoryBits = 0
    Call resetStats
    Call RememberStartingMissions
    Exit Sub
End Sub

    'I am client, host pressed Declare War
Public Sub startNewWar(ByteMessage() As Byte)
    Dim i As Long
    'Call updateTestViewer("startNewWar")
    
    'frmAdvanced.Hide
    gWinMemoryBits = 0
    For i = 1 To 42
        gCtryOrder(i) = ByteMessage(i + 1)
        gCtryScore(i) = ByteMessage(i + 43)
        gCountryOwner(i) = ByteMessage(i + 85)
        'MapColor(i) = playerID(gCountryOwner(i)).lngColor
    Next
    For i = 0 To 13     '128..141
        gMissions(i).IsActive = CBool(ByteMessage(i + 128))
    Next
    For i = 1 To 6      '142..147
        gPlayerID(i).mission = CInt(ByteMessage(i + 141))
    Next
    gPlayerTurn = CInt(ByteMessage(148))
    
    'Cards.
    gCardDeck(0) = CInt(ByteMessage(149) And 15)
    gCardDeck(1) = CInt(ByteMessage(149) \ 16)
    gCardDeck(2) = CInt(ByteMessage(150) And 15)
    gCardDeck(3) = CInt(ByteMessage(150) \ 16)
    
    Call GetRestOfByte(ByteMessage, 151)
    Call UnpackSetupScreen(ByteMessage)
    
    With A2opportunity
        .IsActive = False
        .pathPointer = 0
        .Path(0) = 0
    End With
    
    'cardmode = Abs(CInt(optCardsFixed.Value + (optCardsIncrease.Value * 2)))
    If GetCardMode = 2 Then
        gMaxCardValue = CInt(txtMaximumCardValue.Text)
    End If
    'If chkCardsHidden.Value = 1 Then
    '    CardsUp = False
    'Else
    '    CardsUp = True            'CardsUp = Not (chkCardsHidden.Value)
    'End If
    gCurrentCardValue = gcCardStartValue
    Call AssignPlayerIDs

    'optimizeDice = (chkOptimizeDefenceDice.Value = 1)
    
    If optnFastWar.Checked Then
        playSpeed = playFast
    Else
        playSpeed = playSlow
    End If
    
    'Missions on.
    If chkMsnMissionsOn.Value = vbChecked Then
        Call RememberStartingMissions
    Else
        Call ClearMissions
    End If

    InfoBoxPrint 0
    TheMainForm.pctInfoBox.BackColor = &HFFFFFF
    TheMainForm.BackColor = &H8000000F
    Call ClearMainCardsArea
    Call ClearDiceFromBoard
    InfoBoxPrnCR 1, 139
    
    gAskedToSeeMission = False
    boolIssueCard = False               'Hasn't got a card yet!
    SetupScreen.Visible = False
    Call ChangeTitlebarText(Phrase(35) + Trim(warSit.filename))         'Global Siege -
    Call ShowMenuBar(SetupScreen.Visible)
    gComputerPressed = False
    boolDrawnWin = False
    notHitMove = True
    gComputerAquiredCards = False
    
    Call ToglleCardKeys(False)
    Call resetMoveTimes                'Limited moves clear
    Call ResetCardsNewWar
    Call ShowNewMap                    'Normal
    net.setupControlChange = False
    Call AuditPlayerRecord
    Call AuditAddPointsIssued(gPlayerTurn, gPlayerValue)
    Call SyncForgroundMap("startNewWar")
End Sub

    'Returns true if current player is unclaimed remote
Private Function unclaimedPlayer(Optional whichPlayer As Integer = -1) As Boolean
    If whichPlayer = -1 Then
        whichPlayer = gPlayerTurn
    End If
    If gPlayerID(whichPlayer).playerWho <> remoteIndex Then
        unclaimedPlayer = False
        Exit Function
    End If
    unclaimedPlayer = (netWorkSituation = cNetNone) _
    Or (net.playerOwner(whichPlayer - 1) = myTerminalNumber)
End Function

'Return which player controlles the current player.
Public Function GetPlayerController(Optional pPlayer As Integer = -1) As Integer
    
    If pPlayer = -1 Then
        pPlayer = gPlayerTurn
    End If

    If unclaimedPlayer(pPlayer) Then
        GetPlayerController = 2
    Else
        GetPlayerController = gPlayerID(pPlayer).playerWho
    End If
End Function

    'I am host, send country owner, score and chosen order of each country
    '149+ contains war name
Private Sub sendOwnerScoreOrder()
    Dim ByteMessage() As Byte
    Dim Byt2() As Byte
    Dim i As Long
    
    ReDim ByteMessage(150) As Byte
    For i = 1 To 42     '2..127
        ByteMessage(i + 1) = CByte(gCtryOrder(i))
        ByteMessage(i + 43) = CByte(gCtryScore(i))
        ByteMessage(i + 85) = CByte(gCountryOwner(i))
    Next
    For i = 0 To 13     '128..141
        ByteMessage(i + 128) = CByte(gMissions(i).IsActive)
    Next
    For i = 1 To 6      '142..147
        ByteMessage(i + 141) = CByte(gPlayerID(i).mission)
    Next
    ByteMessage(148) = CByte(gPlayerTurn)
    ByteMessage(149) = CByte(gCardDeck(0) + gCardDeck(1) * 16)
    ByteMessage(150) = CByte(gCardDeck(2) + gCardDeck(3) * 16)
    
    Call PackSetupScreen(Byt2)
    Call appendByteArray(ByteMessage, Byt2)
    Call netMain.sendOwnerScoreOrder(ByteMessage)
    Call resetChangeList
End Sub

    'Pack game settings for a new player (I host)
Public Sub packWarSettings(ByteMessage() As Byte)
    Dim Byt2() As Byte
    Dim i As Long
    
    ReDim ByteMessage(152) As Byte
    For i = 1 To 42     '2..127
        ByteMessage(i + 1) = getLowerByte(gCtryScore(i))
        ByteMessage(i + 43) = getUpperByte(gCtryScore(i))
        ByteMessage(i + 85) = CByte(gCountryOwner(i))
    Next
    For i = 0 To 13     '128..141
        ByteMessage(i + 128) = CByte(gMissions(i).IsActive)
    Next
    For i = 1 To 6      '142..147
        ByteMessage(i + 141) = CByte(gPlayerID(i).mission)
    Next
    ByteMessage(148) = CByte(gPlayerTurn)
    ByteMessage(149) = CByte(gCardDeck(0) + gCardDeck(1) * 16)
    ByteMessage(150) = CByte(gCardDeck(2) + gCardDeck(3) * 16)
    ByteMessage(151) = CByte(gCurrentCardValue)
    ByteMessage(152) = CByte(gCurrentMode)
    
    Call PackSetupScreen(Byt2)
    Call appendByteArray(ByteMessage, Byt2)
End Sub

Public Sub unpackWarSettings(ByteMessage() As Byte)
    Dim i As Long
    'Call updateTestViewer("unpackWarSettings")
    
    gNewWarPlayerTurn = 0
    
    For i = 1 To 42     '2..127
        gCtryScore(i) = CInt(ByteMessage(i + 1) + (ByteMessage(i + 43) * 256))
        gCountryOwner(i) = CInt(ByteMessage(i + 85))
    Next
    For i = 0 To 13     '128..141
        gMissions(i).IsActive = CBool(ByteMessage(i + 128))
    Next
    For i = 1 To 6      '142..147
        gPlayerID(i).mission = CInt(ByteMessage(i + 141))
    Next
    
    gPlayerTurn = CInt(ByteMessage(148))
    gCardDeck(0) = CInt(ByteMessage(149) And 15)
    gCardDeck(1) = CInt(ByteMessage(149) \ 16)
    gCardDeck(2) = CInt(ByteMessage(150) And 15)
    gCardDeck(3) = CInt(ByteMessage(150) \ 16)
    gCurrentCardValue = CInt(ByteMessage(151))
    gCurrentMode = CInt(ByteMessage(152))
    
    Call GetRestOfByte(ByteMessage, 153)
    Call UnpackSetupScreen(ByteMessage)
    Call refreshMap
    Call AssignPlayerIDs
    
    'cardmode = Abs(CInt(optCardsFixed.Value + (optCardsIncrease.Value * 2)))
    'gateDefence = gateDefenseTable(cardmode)
    If GetCardMode = 2 Then
        gMaxCardValue = CInt(txtMaximumCardValue.Text)
    End If
    'If chkCardsHidden.Value = 1 Then
    '    CardsUp = False
    'Else
    '    CardsUp = True            'CardsUp = Not (chkCardsHidden.Value)
    'End If

    'optimizeDice = (chkOptimizeDefenceDice.Value = 1)
    
    If optnFastWar.Checked Then
        playSpeed = playFast
    Else
        playSpeed = playSlow
    End If
    
    If chkMsnMissionsOn.Value = vbChecked Then         'Missions on
        Call RememberStartingMissions
    Else
        Call ClearMissions
    End If
    
    gAskedToSeeMission = False
    boolIssueCard = False               'Hasn't got a card yet!
    SetupScreen.Visible = False
    Call ChangeTitlebarText((Phrase(35) + Trim(warSit.filename)))        'Global Siege -
    Call ShowMenuBar(Not mnuFullScreen.Checked Or SetupScreen.Visible)
    gComputerPressed = False
    boolDrawnWin = False
    notHitMove = True
    gComputerAquiredCards = False
    
    Call ToglleCardKeys(False)
    Call resetMoveTimes                'Limited moves clear
    net.setupControlChange = False
    
    tfRate1.Value = True
    transferNmbr = 1
    Call resetChangeList
    Timer1.Enabled = False
    gPlayerValue = GetPlayerValue(gPlayerTurn)
    Call AuditPlayerRecord
    Call AuditAddPointsIssued(gPlayerTurn, gPlayerValue)
    TheMainForm.pctInfoBox.BackColor = gPlayerID(gPlayerTurn).bkgndColor
    Call DrawLittleCards
    If flashingBorder And (GetPlayerController(gPlayerTurn) = 0) Then
        If netWorkSituation <> cNetNone Then       'Flash color if human on this terminal
            If (net.playerOwner(gPlayerTurn - 1) = myTerminalNumber) Then
                TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
                tmrFlashInfoBox.Enabled = True
            Else
                TheMainForm.BackColor = &H8000000F
            End If
        Else
            TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
            tmrFlashInfoBox.Enabled = True
        End If
    Else
        TheMainForm.BackColor = &H8000000F
    End If
    
    Call printPlaceUnits        ' You have xx units to place
    
    Call CheckToSeeMission(gPlayerTurn)
    Call ToggleKeys(gPlayerValue = 0)
    
    'gCurrentMode is passed from the host but setting to 2 seems to solve some problems.
    'TODO: Needs re-investigation.
    gCurrentMode = 2
    Call AutoPlayerSelect
    Call SyncForgroundMap("unpackWarSettings")
End Sub

'Report changes to setup screen to remote terminals if I am the host.
'Claim/disclaim armys if I am a client terminal.
Public Sub CheckSetupForChange()
    Dim vIndex As Byte
    
    'Check if networked or something has changed.
    If netWorkSituation = cNetNone Then
        net.setupControlChange = False
        Exit Sub
    ElseIf Not net.setupControlChange Then
        Exit Sub
    End If
    
    'If I am a client terminal.
    If netWorkSituation = cNetClient Then
        For vIndex = 0 To 5
            
            'Can only change player owner if it is unclaimed
            'or I have already claimed it.
            If net.playerOwner(vIndex) <> myTerminalNumber Then
                
                'If this army is unclaimed, then claim it.
                If (playerSelect_getIndex(CInt(vIndex)) <> remoteIndex) _
                And (vscrollPlayerSelect(vIndex).Enabled) Then
                    Call netMain.ClaimPlayer(vIndex, vscrollPlayerSelect(vIndex).Value)
                    net.playerOwner(vIndex) = myTerminalNumber
                End If
            Else
            
                'If I already own this army, unclaim it.
                If (playerSelect_getIndex(CInt(vIndex)) = remoteIndex) _
                And (vscrollPlayerSelect(vIndex).Enabled) Then
                    Call netMain.DisClaimPlayer(vIndex)
                    net.playerOwner(vIndex) = 0
                End If
            End If
            
            'If the player owner up/down button is disabled then it has
            'been claimed by another terminal or the host.
            If Not vscrollPlayerSelect(vIndex).Enabled Then
                Call playerSelect_showIndex(CInt(vIndex), remoteIndex)
            End If
        Next
        net.setupControlChange = False
    
    'If I am the host, send the entire setup screen.
    ElseIf (net.setupControlChange) And (netMain.CountTerminals <> 0) Then
        net.setupControlChange = False
        netMain.sendSetupScreen
    End If
End Sub

Private Sub chkMsnAreUnique_Click()
    On Error Resume Next
    
    Call PopulateMissionList
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub chkMsnArmyWipeout_Click()
    On Error Resume Next
    
    Call PopulateMissionList
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub chkMsnConquerHold_Click()
    On Error Resume Next
    
    Call PopulateMissionList
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub chkMsnMustComplete_Click()
    On Error Resume Next
    
    Call PopulateMissionList
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub chkMsnWinImmediate_Click()
    On Error Resume Next
    
    Call PopulateMissionList
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub chkSortDice_Click()
    Call PrintDiceOdds
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

'Switch to English.
Private Sub cmdEnglish_Click()
    On Error Resume Next
    frmLanguage.Show vbModal
    Call SetNewWords
    Call refreshMap
    Call SyncForgroundMap("cmdEnglish_Click")
End Sub

'Work out the odds of the chosen dice setup.
Private Sub cmdWorkoutOdds_Click()
    'Call WorkOutDiceOdds(1000000)
    Call CrunchDiceOdds
End Sub

 'Testing 122.
Private Sub Command1_Click()
    Dim dummy As Long
    Static vStyle As Boolean
    Dim Control As Control
    
    On Error Resume Next
    
    Mask4.Show

End Sub

Private Sub frameSetup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub frameSetupControls_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub frmSetupCards__MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub frmSetupCards_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub fSetupPlayerNumber_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub fSetupWarOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub hlpCheckForUpdates_Click()
    On Error Resume Next
    hlpCheckForUpdates.Checked = Not hlpCheckForUpdates.Checked
    netMain.tmrCheckGSNews.Interval = 5000
    netMain.tmrCheckGSNews.Enabled = True
End Sub

Private Sub lblVersion_Click()
    On Error Resume Next
    If Len(lblVersion.Tag) > 0 Then
        Call OpenWebPage(lblVersion.Tag)
    End If
End Sub

Private Sub mnuAttack_Click()
    On Error Resume Next
    Call AttackClicked
    Call SyncForgroundMap("mnuAttack_Click")
End Sub

Private Sub mnuAutoRestart_Click()
    On Error Resume Next
    mnuAutoRestart.Checked = Not mnuAutoRestart.Checked
    
    'Restart the war if the gCurrentMode is 18.
    If mnuAutoRestart.Checked And gCurrentMode = 18 Then
        mnuFileReset_Click
    End If
End Sub

Private Sub mnuCancelSetup_Click()
    Call cmdSUPcncl_Click
End Sub

Private Sub mnuClientLan_Click()
    On Error Resume Next
    Load netMain
    netMain.optLan.Value = True
    netMain.optJoin.Value = True
    netMain.DisplaySessionLocator
End Sub

Private Sub mnuClientQue_Click()
    On Error Resume Next
    Unload netWiz
    Load netMain
    netMain.txtSesName.Text = Trim(warSit.filename)
    netWiz.ScreenNo = 5
    'netWiz.Show vbModal, TheMainForm
    netWiz.Show , TheMainForm
End Sub

Private Sub mnuDeclareWar_Click()
    Call cmdSetupOk_Click
End Sub

Private Sub mnuFlashInfoBox_Click()
    mnuFlashInfoBox.Checked = Not mnuFlashInfoBox.Checked
End Sub

'Remove the title bar and hide the menu and fill the screen with war.
Private Sub mnuFullScreen_Click()
    Static sLastWindowsState As Integer
    Static sLastToolboxState As Boolean
    Static sLastBorderState As Boolean
    
    On Error Resume Next
    mnuFullScreen.Checked = Not mnuFullScreen.Checked
    
    If mnuFullScreen.Checked Then
        'Save windows size and toolbox state and maximize & hide.
        sLastToolboxState = mnuOptToolbox.Checked
        mnuOptToolbox.Checked = False
        Call ShowToolBar(False, False)
        sLastWindowsState = TheMainForm.WindowState
        Picture1.BorderStyle = 0
        TheMainForm.WindowState = vbMaximized
        sLastBorderState = mnuViewBorder.Checked
        mnuViewBorder.Checked = False
    Else
        'Restore windows size and toolbox state.
        TheMainForm.WindowState = sLastWindowsState
        mnuOptToolbox.Checked = sLastToolboxState
        mnuViewBorder.Checked = sLastBorderState
        Picture1.BorderStyle = 1
        Call ShowToolBar(mnuOptToolbox.Checked, False)
    End If
    
    'Save current title in the titlebar's tag. This must be done before
    'full screen because it causes the menu item to "slide" in and looks weird.
    If mnuFullScreen.Checked Then
        Call TagTitlebarText
    End If
    
    Call FlipWindowsTitleBar(TheMainForm.hWnd, Not mnuFullScreen.Checked)
    Call ShowMenuBar(SetupScreen.Visible)
    
    'Restore titlebar text from the titlebar's tag. This must be done
    'after normal screen has been restored because it makes the whole windo jump.
    If Not mnuFullScreen.Checked Then
        Call TagTitlebarText
    End If
End Sub

'Show or hide the menu bar when it is hidden in full screen mode.
'The menu bar should be visible when ever the setup screen is displayed.
Private Sub ShowMenuBar(pVisible As Boolean)
    Dim vShowNow As Boolean
    
    vShowNow = Not mnuFullScreen.Checked Or mnuViewBorder.Checked Or pVisible
    
    mnuFile.Visible = vShowNow
    mnuNet.Visible = vShowNow
    'mnuView.Visible = vShowNow
    mnuOptions.Visible = vShowNow
    mnuMission.Visible = vShowNow
    mnuHelp.Visible = vShowNow
    mnuMainBrk1.Visible = vShowNow
    mnuMainBrk2.Visible = vShowNow
    mnuAttack.Visible = vShowNow And Not cmdAttack.Visible
    mnuMove.Visible = vShowNow And Not cmdAttack.Visible
    mnuPass.Visible = vShowNow And Not cmdAttack.Visible
    
    'Make a call to the watchdog timer because it has certain functions
    'that handle small size windows that switch between buttons and
    'menu items.
    Call tmrWatchDog_Timer
    
End Sub

'Put titlebar text into the form's tag and change it to just gcApplicationName.
'This is to prevent the titlebar text from changing when in full screen mode
'causing the form to jump back into titlebar mode.
Private Sub TagTitlebarText()
    If mnuFullScreen.Checked Then
        'Full screen
        TheMainForm.Tag = TheMainForm.Caption
        TheMainForm.Caption = gcApplicationName
    Else
        'Normal window.
        If Trim(TheMainForm.Tag) = "" Then
            TheMainForm.Tag = gcApplicationName
        End If
        TheMainForm.Caption = TheMainForm.Tag
    End If
End Sub

'Change the titlebar text without disrupting the titlebar if in
'full screen mode. TheMainForm.tag is used in that case.
Private Sub ChangeTitlebarText(pNewText As String)
    If mnuFullScreen.Checked Then
        'Full screen, write to tag.
        TheMainForm.Tag = pNewText
    Else
        'Normal window, change titlebar text directly.
        TheMainForm.Caption = pNewText
    End If
End Sub

Private Sub mnuHostQueue_Click()
    On Error Resume Next
    Unload netWiz
    Load netMain
    netMain.txtSesName.Text = Trim(warSit.filename)
    netWiz.ScreenNo = 0
    'netWiz.Show vbModal, TheMainForm
    netWiz.Show , TheMainForm
End Sub

Private Sub mnuMisSeeReminder_Click()
    mnuMisSeeReminder.Checked = Not mnuMisSeeReminder.Checked
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print KeyCode
    
    'Ctrl or Alt key pressed, show menu bar and stop it being hidden again by a mouse move for a while.
    If KeyCode = 17 Or KeyCode = 18 Then
        mnuFile.Tag = CStr(CLng(Time * 100000))
        Call ShowMenuBar(True)
    End If
    
    'Pause button.
    If KeyCode = 19 Then
        Call ActivatePauseMode
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ctrl key released. Hide the menu.
    If KeyCode = 17 And mnuFullScreen.Checked Then
        mnuFile.Tag = ""
        Call ShowMenuBar(SetupScreen.Visible)
    End If
End Sub

    ' Press keys to change timer speed
Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim MultW As Single
    Dim MultD As Single
    
    On Error Resume Next
    
    'Debug.Print KeyAscii
    
    MultW = 1
    MultD = 1
    If (Chr(KeyAscii) = "-" Or Chr(KeyAscii) = "_") Then
        If playSpeed < 300 Then
            MultW = 1.5
        End If
        If diceSpeed < 300 Then
            MultD = 1.5
        End If
    ElseIf (Chr(KeyAscii) = "+" Or Chr(KeyAscii) = "=") Then
        If playSpeed > 3 Then
            MultW = 0.66666
        End If
        If diceSpeed > 3 Then
            MultD = 0.66666
        End If
    Else
        'Send other key presses to the chatter box.
        netChatterBox.Show , TheMainForm
        netChatterBox.txtWrite.SetFocus
        DoEvents
        SendKeys Chr(KeyAscii)
        Exit Sub
    End If
    playSpeed = playSpeed * MultW
    diceSpeed = diceSpeed * MultD
    
    Timer2.Interval = Timer2.Interval * MultW
    Timer1.Interval = Timer1.Interval * MultD
    
    If Timer2.Interval < 3 Then Timer2.Interval = 1
    If Timer1.Interval < 3 Then Timer1.Interval = 1
    If Timer2.Enabled Then
        Timer2.Enabled = False
        Timer2.Enabled = True
    End If
    If Timer1.Enabled Then
        Timer1.Enabled = False
        Timer1.Enabled = True
    End If
    
End Sub

    'Remember timer states and pause
Public Sub ActivatePauseMode(Optional StopNow As Boolean)
    Static t1 As Boolean
    Static t2 As Boolean
    
    If netWorkSituation <> cNetNone Then       'Only pause if player owned by this terminal
        If net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber Then
            Exit Sub
        End If
    End If
    
        'Safety trap
    If gPauseActive And (Timer1.Enabled Or Timer2.Enabled) Then
        gPauseActive = False
    End If
    
    If Not gPauseActive Then             '> STOP <
        gPauseActive = True
        t1 = Timer1.Enabled
        t2 = Timer2.Enabled
        Timer1.Enabled = False
        Timer2.Enabled = False
        InfoBoxPrnCR 0
        InfoBoxPrint 5                           'bold
        InfoBoxPrint 10                          'font size * 1.5
        InfoBoxPrnCR 1, 214                      '"PAUSE"
        InfoBoxPrint 11                          'reset font size
        InfoBoxPrint 6                           'normal
    Else                            '> GO <
        If StopNow = gPauseActive Then Exit Sub
        If gCurrentMode = 3 Then
            InfoBoxPrint 0
        End If
        gPauseActive = False
        Timer1.Enabled = t1
        Timer2.Enabled = t2
    End If
    
    Call handleUpdate
End Sub

Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub Frame9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub mnuMove_Click()
    On Error Resume Next
    Call MoveClicked
End Sub

Private Sub mnuNetAdvanced_Click()
    On Error Resume Next
    netMain.Show , TheMainForm
End Sub

Private Sub mnuNetChat_Click()
    On Error GoTo ErrHand
    netChatterBox.Show , TheMainForm
    Exit Sub
ErrHand:
    Exit Sub
End Sub

Private Sub mnuNetClientInternet_Click()
    On Error Resume Next
    Load netMain
    netMain.optInet.Value = True
    netMain.optJoin.Value = True
    netMain.DisplaySessionLocator
End Sub

Private Sub mnuNetCntr_Click()
    On Error Resume Next
    mnuNetCntr.Checked = Not mnuNetCntr.Checked
    
    If mnuNetCntr.Checked Then
        'Show the counter.
        frmCounter.Move TheMainForm.Left + TheMainForm.Width - frmCounter.Width, _
                        TheMainForm.Top + TheMainForm.Height - frmCounter.Height
        frmCounter.txtAudit.SelStart = 0
        frmCounter.txtAudit.SelLength = 9999
        frmCounter.txtAudit.SelText = ""
        frmCounter.txtAudit.Visible = True
        frmCounter.Show , TheMainForm
    Else
        'Hide the counter.
        frmCounter.Hide
    End If
End Sub

Private Sub mnuNetDisconnect_Click()
    On Error Resume Next
    'If half way through the draw win sequence or waiting to restart the war.
    tmrDrawWin.Tag = ""
    Unload netMain
End Sub

Private Sub mnuNetHostInet_Click()
    On Error Resume Next
    Load netMain
    netMain.optHost.Value = True
    netMain.optInet.Value = True
    netMain.txtSesName.Text = Trim(warSit.filename)
    netMain.Show , TheMainForm
End Sub

Private Sub mnuNetHostLan_Click()
    On Error Resume Next
    Load netMain
    netMain.optHost.Value = True
    netMain.optLan.Value = True
    netMain.txtSesName.Text = Trim(warSit.filename)
    netMain.Show , TheMainForm
End Sub

'Show the stats if gCurrentMode is 18 (Win message drawn).
Private Sub mnuOptStats_Click()
    On Error Resume Next
    mnuOptStats.Checked = Not mnuOptStats.Checked
    If gCurrentMode = 18 Then
        Call frmStats.ShowStats
    End If
End Sub

Private Sub mnuPass_Click()
    On Error Resume Next
    Call EndClicked
    Call SyncForgroundMap("mnuPass_Click")
End Sub

Private Sub mnuViewBorder_Click()
    On Error Resume Next
    mnuViewBorder.Checked = Not mnuViewBorder.Checked
    Call Form_Resize
End Sub

'Show font dialog box and change font.
Private Sub mnuViewLFont_Click()
    On Error Resume Next
    
    'Set up the Windows Font Selector dialog box.
    CommonDialog1.Flags = cdlCFScreenFonts _
                        Or cdlCFANSIOnly _
                        Or cdlCFForceFontExist _
                        Or cdlCFScalableOnly
    CommonDialog1.FontName = Picture1.Font.name
    CommonDialog1.FontBold = Picture1.Font.Bold
    CommonDialog1.FontItalic = Picture1.Font.Italic
    CommonDialog1.FontSize = Picture1.Font.Size
    'CommonDialog1.max = Picture1.font.Size
    'CommonDialog1.min = Picture1.font.Size
    
    'Show the Windows Font Selector.
    CommonDialog1.ShowFont
    
    'Main map.
    Picture1.Font.name = CommonDialog1.FontName
    Picture1.Font.Bold = CommonDialog1.FontBold
    Picture1.Font.Italic = CommonDialog1.FontItalic
    Mask4.Map1.Font.name = CommonDialog1.FontName
    Call ScaleViewportMapFont
    Call ScaleLittleCardFont       'Size Little Card font size.
    Call refreshMap
    
    Call SyncForgroundMap("mnuViewLFont_Click")
    
    'Info box.
    pctInfoBox.Font.name = CommonDialog1.FontName
    pctInfoBox.Font.Bold = False
    pctInfoBox.Font.Italic = CommonDialog1.FontItalic
    Call ScaleInfoBoxFont
    
    'Chat box.
    Call netChatterBox.ChooseChatBoxFont
End Sub

Private Sub mnuViewQualityDisplay_Click()
    On Error Resume Next
    mnuViewQualityDisplay.Checked = Not mnuViewQualityDisplay.Checked
    
    'Make sure the viewport gets refreshed.
    gSyncViewportNeeded = True
    Call SyncForgroundMap("mnuViewQualityDisplay_Click")
End Sub

Private Sub optDiceSame_Click(Index As Integer)
    Call PrintDiceOdds
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub pctFixedCards_Click()
    Call CheckSetupForChange
End Sub

Private Sub pctTheDeck_Click()
    Call CheckSetupForChange
End Sub

Private Sub pctWorkoutOdds_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub Picture2_Click()
    Call CheckSetupForChange
End Sub

'Select and print the attack winning odds.
Private Sub PrintDiceOdds()
    Dim vDiceOdds() As String
    
    vDiceOdds = Split(GetDiceOdds, ",")
    lblAttackProb.Caption = vDiceOdds((udDiceThrown(0).Value - 1) * 5 + (udDiceThrown(1).Value - 1)) & "%"
End Sub

Private Sub optDiceRules_Click(Index As Integer)
    Call EnableDiceOptions(frmDiceRules.Enabled)
    Call PrintDiceOdds
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub pctInfoBox_Click()
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
End Sub

Private Sub pctSetupFirstPlayer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Call CheckSetupForChange
End Sub

Private Sub pctSetupWarOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Call CheckSetupForChange
End Sub

Private Sub pctTransfer_Click()
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print "Picture1_KeyDown: ", KeyCode
    'F11 is 122
    If KeyCode = 122 And mnuFullScreen.Checked Then
        Call mnuFullScreen_Click
    End If
    
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    'Debug.Print "Picture1_KeyPress: ", KeyAscii
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    gCurrentMousePosX = Int(x)
    gCurrentMousePosY = Int(y)
    
    'Hide or show the menu bar. A bit long winded but the idea is to
    'only call the function as few times as required.
    If mnuFullScreen.Checked And Trim(mnuFile.Tag) = "" Then
        If y < 25 And Not mnuFile.Visible And Not Toolbar1.Visible Then
            Call ShowMenuBar(True)
            'Debug.Print "Show"
        ElseIf y > 25 And mnuFile.Visible And Not mnuViewBorder.Checked Then
            Call ShowMenuBar(False)
            'Debug.Print "Hide"
        End If
    End If
End Sub

Private Sub PlayerNumber_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub Picture2__MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub pctSetupCards_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub PlayerSelect_GotFocus(Index As Integer)
    On Error Resume Next
    If SetupScreen.Visible Then
        SetupScreen.SetFocus
    End If
End Sub

Private Sub plrOpt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
End Sub

Private Sub SetupScreen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 And mnuFullScreen.Checked Then
        Call mnuFullScreen_Click
    End If
End Sub

    'If setup controls have changed then send new values
    'This saves sending values for every change made
Private Sub SetupScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call CheckSetupForChange
    'If mnuFullScreen.Checked Then
    '    Call ShowMenuBar(Y < 25)
    'End If
End Sub

Private Sub chkCardsHidden_Click()
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub chkMsnMissionsOn_Click()
    Static vLock As Boolean
    
    If Not vLock Then
        vLock = True
        
        net.setupControlChange = True
        Call CheckSetupForChange
        Call EnableMissionOptions
        Call PopulateMissionList
        vLock = False
    End If
End Sub

Private Sub chkCardsVulture_Click()
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub optPlr1FirstPlayer_Click()
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub optRandomFirstPlayer_Click()
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub chkExtraStartingUnits_Click()
    txtExtraStartingUnits.Enabled = (chkExtraStartingUnits.Value = vbChecked And chkExtraStartingUnits.Enabled)
    udExtraStartingUnits.Enabled = (chkExtraStartingUnits.Value = vbChecked And chkExtraStartingUnits.Enabled)
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

    'Start attack mode
Public Sub cmdAttack_Click()
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    Call AttackClicked
    Call SyncForgroundMap("cmdAttack_Click")
End Sub

Private Sub AttackClicked()
    
    If gPauseActive Then Call ActivatePauseMode(False)
    If gCurrentMode = 13 Or gCurrentMode = 18 Or gCurrentMode = 6 Or gMapSetupLock Then
        Exit Sub
    End If
    
    If GetPlayerController(gPlayerTurn) <> 0 And Not gComputerPressed Then
        Exit Sub
    End If
    
    Call ColorCountryUnderAttack(0)
    
    With AtkForCard
        .from = 0
        .To = 0
        .On = False
    End With
    
    Call CheckWinDuringTurn(gPlayerTurn)
    If Not notHitMove Then                  'Hit move
        Exit Sub
    End If
    gCurrentMode = 20                    'Attack mode
    InfoBoxPrint 0
    InfoBoxPrint 9, gPlayerTurn
    InfoBoxPrint 3, 1
    InfoBoxPrint 5               'bold
    InfoBoxPrnCR 1, 64           'attacks
    InfoBoxPrnCR 6               'normal
    InfoBoxPrnCR 7
    If gPlayerID(gPlayerTurn).playerWho = 0 Then
        InfoBoxPrnCR 1, 65           '"<Click defending country>"
    End If
    
    Call addChangeToList(gTargetCtry, 0, 0)
End Sub

Private Sub cmdCardCncl_Click()
    Dim tstPress As Integer
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    If gPauseActive Then Call ActivatePauseMode(False)
    If gCurrentMode = 13 Then    'Game was won
        Exit Sub
    ElseIf gCurrentMode = 18 Then
        Exit Sub
    End If
    tstPress = (GetPlayerController(gPlayerTurn) <> 0) _
                And (Not gComputerPressed)
    If tstPress Then
        Exit Sub
    End If
    Call Mode5
    Call SyncForgroundMap("cmdCardCncl_Click")
End Sub

'Host has instructed me to pass my turn.
Public Sub ForfeitTurn()
    If net.playerOwner(gPlayerTurn - 1) = myTerminalNumber Then
        If gPlayerValue = 0 Then
            gPlayerValue = gPickedUpUnits
        End If
        Call AuditUpdateScore(gPlayerTurn, gPlayerValue)
        gComputerPressed = True
        Call EndClicked
        gComputerPressed = False
        
        'Make sure the viewport gets refreshed.
        gSyncViewportNeeded = True
        
        Call SyncForgroundMap("ForfeitTurn")
    End If
End Sub

    'Pass, get next player
Public Sub cmdEnd_Click()
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    Call EndClicked
    If gCurrentMode <> 18 Then
        'Make sure the viewport gets refreshed.
        gSyncViewportNeeded = True
        Call SyncForgroundMap("cmdEnd_Click")
    End If
End Sub

Private Sub EndClicked()
    Dim rslt As Boolean
    Dim tstPress As Integer
    
    On Error Resume Next
    
    If gPauseActive Then
        Call ActivatePauseMode(False)
    End If
    
    If gCurrentMode = 13 Or gCurrentMode = 18 Or gMapSetupLock Then
        Exit Sub
    End If

    If (GetPlayerController(gPlayerTurn) <> 0) And (Not gComputerPressed) Then
        Exit Sub
    End If
    
    'Reset some things for the computer players.
    With AtkForCard
        .from = 0
        .To = 0
        .On = False
    End With
    With A2opportunity
        .IsActive = False
        .pathPointer = 0
        .Path(0) = 0
    End With
    
    tstPress = (netWorkSituation <> cNetNone) _
    And (myTerminalNumber = net.playerOwner(gPlayerTurn - 1))
    
    If nmbrOfPlayers = 0 Then
        Exit Sub
    End If
    
    Call backUp
    warSit.SetupScreen = SetupScreen.Visible
    Call SaveCheckpoint              'Remember here
    Call CheckWinDuringTurn(gPlayerTurn)
    
    If boolIssueCard Then
        Call DealACard
    End If
    boolIssueCard = False
    notHitMove = True               'Hasn't clicked a key yet
    gCurrentMode = 1
    Call ToggleKeys(False)
    Call ToglleCardKeys(False)
    tfRate1.Value = True            'Make transfer rate 1 again
    transferNmbr = 1
    
    'Next player's turn from here on.
    gPlayerTurn = GetNextPlayer(gPlayerTurn)
    
    'Reset paths for new A3 turn.
    If GetPlayerController(gPlayerTurn) = A3Index Then
        A3.startOfNewTurn = True
    End If
    
    If GetPlayerController(gPlayerTurn) = 2 Then
        Call A3.randomizeNeigbor
    End If
    
    gPlayerValue = GetPlayerValue(gPlayerTurn)
    gCurrentMode = 2
    Call CheckToSeeMission(gPlayerTurn)
    
    mnuOptUndo.Enabled = gCheatMode.undoEnabled
    Toolbar1.Buttons(8).Enabled = mnuOptUndo.Enabled
    Call refreshMap
    
    If CheckWinStartOfTurn(gPlayerTurn) Then
        Exit Sub
    End If
    Call showPlayerInfo
    Call resetMoveTimes
    Call CardOutOfHand(gPlayerTurn)              'Put cards out of hand
    Call CheckCards
    Call DrawLittleCards
    
    If (netWorkSituation <> cNetNone) _
    And (net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber) Then
    
        'Hide dice if network game and refresh rate not showing dice when not my turn.
        If netMain.terminalSpeed > 0 Then
            Call ClearDiceFromBoard
        End If
    Else
        Call AutoPlayerSelect
    End If
    If tstPress Then
        Call AuditShadowAppend
        Call netMain.SendRefresh
        Call AuditPlayerCompare
        Call AuditPlayerRecord
        Call AuditAddPointsIssued(gPlayerTurn, gPlayerValue)
    End If
    If Not gComputerPressed Then
        net.madeUpdate = False
    End If
    If IsComputerPlayer(gPlayerTurn) Then
        Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
    End If
End Sub

'Rember the current war situation for save and undo functions.
'Called at the end of every thurn. Counterpart RevertToCheckpoint().
Private Sub SaveCheckpoint()
    Dim vIndex As Long
    
    With warSit
    
    'Save the owner and score for each country. Country 0 is not used.
    For vIndex = 1 To 49
        .sCtryOwner(vIndex) = gCountryOwner(vIndex)
        .sCtryScore(vIndex) = gCtryScore(vIndex)
    Next vIndex
    
    'Save the owner of each army.
    For vIndex = 0 To 6
        .sPlayerID(vIndex) = gPlayerID(vIndex)
    Next vIndex
    
    
    .sCurrentCardValue = gCurrentCardValue
    .sNmbrOfPlayers = nmbrOfPlayers
    .sPlayerTurn = gPlayerTurn
    .sBoolIssueCard = boolIssueCard
    For vIndex = 0 To 3
        .sCards(vIndex) = gCardDeck(vIndex)
    Next
    .kCardMode = GetCardMode
    .kCardsUp = (chkCardsHidden.Value = vbUnchecked)
    .kDiceSpeed = diceSpeed
    .kMaxCardValue = gMaxCardValue
    .kMissionsOn = (chkMsnMissionsOn.Value = vbChecked)
    .kMoveLimit = gMoveLimit
    .kPlaySpeed = playSpeed
    
    'Save version info to help keep this stuff backwards compatable
    'with earlier war file formats by allowing version detection.
    .GSVersion = GetVersionInfo
    
    End With
    
End Sub

'Retrieve saved situation for undo and open war file.
'Counterpart SaveCheckPoint().
Private Sub RevertToCheckpoint()
    Dim vIndex As Long
    Dim vBodgyHack As Boolean
    Dim vVersionOffset As Long
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    With warSit
    
    'Bail out if there is no save situation.
    If (.sNmbrOfPlayers = 0) Or (.sCtryOwner(1) = 0) Then
        Exit Sub
    End If
    
    'Earlier versions used a 1 based index which was changed to
    'a 0 based index in GlobalSiege 0.9.0092.
    If .GSVersion = "" Or .GSVersion < "00090095" Then
        vVersionOffset = 1
    Else
        vVersionOffset = 0
    End If
    
    For vIndex = 1 To 49
        gCountryOwner(vIndex) = .sCtryOwner(vIndex - vVersionOffset)
        gCtryScore(vIndex) = .sCtryScore(vIndex - vVersionOffset)
    Next vIndex

    For vIndex = 0 To 6
        gPlayerID(vIndex) = .sPlayerID(vIndex)
        If netWorkSituation <> cNetNone And vIndex > 0 Then
            If net.playerOwner(vIndex - 1) <> myTerminalNumber Then
                gPlayerID(vIndex).playerWho = remoteIndex
            End If
        End If
    Next vIndex

    gCurrentCardValue = .sCurrentCardValue
    nmbrOfPlayers = .sNmbrOfPlayers
    gPlayerTurn = .sPlayerTurn
    If gPlayerTurn = 0 Then
        gPlayerTurn = 1
        vBodgyHack = True
    End If
    boolIssueCard = .sBoolIssueCard
    For vIndex = 0 To 3
        gCardDeck(vIndex) = .sCards(vIndex)
    Next
    
    'CardsUp = .kCardsUp
    'Saved as a boolean, unable to change as it may affect saved games compatability.
    'Bit of a hack, but only change the setup screen controls if the setup
    'screen is hidden. The setup controls are loaded elsewhere.
    'TODO: This all needs to be simplified.
    If Not .SetupScreen Then
        If .kCardsUp Then
            chkCardsHidden.Value = vbUnchecked
        Else
            chkCardsHidden.Value = vbChecked
        End If
        
        chkMsnMissionsOn.Value = Abs(CInt(.kMissionsOn))
        Call SetCardMode(.kCardMode)
    End If
    
    gMaxCardValue = .kMaxCardValue
    gMoveLimit = .kMoveLimit
    
    End With
    
    TheMainForm.BackColor = &H8000000F
    gPickedUpUnits = 0
    boolDrawnWin = False
    Call CheckWinDuringTurn(gPlayerTurn)
    If boolIssueCard Then
        Call DealACard
    End If
    boolIssueCard = False
    notHitMove = True               'Hasn't clicked a key yet
    gCurrentMode = 1
    Call ToggleKeys(False)
    Call ToglleCardKeys(False)
    tfRate1.Value = True            'Make transfer rate 1 again
    transferNmbr = 1
    If vBodgyHack Then
        gPlayerTurn = gPlayerTurn - 1
    End If
    gPlayerTurn = GetNextPlayer(gPlayerTurn)
    
    'Reset paths for new A3 turn.
    If GetPlayerController(gPlayerTurn) = A3Index Then
        A3.startOfNewTurn = True
    End If
    
    If GetPlayerController(gPlayerTurn) = 2 Then
        Call A3.randomizeNeigbor
    End If
    
    gPlayerValue = GetPlayerValue(gPlayerTurn)
    gCurrentMode = 2
    Call CheckToSeeMission(gPlayerTurn)
    Call SetDisplayMode
        
    If CheckWinStartOfTurn(gPlayerTurn) Then
        'Call updateTestViewer("RevertToCheckpoint")
        Exit Sub
    End If
    Call showPlayerInfo
    Call resetMoveTimes
    Call CardOutOfHand(gPlayerTurn)              'Put cards out of hand
    Call CheckCards
    Call DrawLittleCards
    Call AutoPlayerSelect
End Sub

    'Put info on info screen
Private Sub showPlayerInfo()
    TheMainForm.pctInfoBox.BackColor = gPlayerID(gPlayerTurn).bkgndColor
    If flashingBorder _
    And (GetPlayerController(gPlayerTurn) = 0) _
    And (Not SetupScreen.Visible) Then
        If netWorkSituation <> cNetNone Then       'Flash color if human on this terminal
            If (net.playerOwner(gPlayerTurn - 1) = myTerminalNumber) Then
                TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
                tmrFlashInfoBox.Enabled = True
            Else
                TheMainForm.BackColor = &H8000000F
            End If
        Else
            TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
            tmrFlashInfoBox.Enabled = True
        End If
    Else
        TheMainForm.BackColor = &H8000000F
    End If
    
    Call printPlaceUnits        ' You have xx units to place
End Sub

'Return the next player if still in war.
Private Function GetNextPlayer(pPlayerTurn As Integer) As Integer
    Dim vIndex As Integer
    
    'Keep trying but not for ever. Twenty times to allow for an
    'increase in players in the future.
    For vIndex = 0 To 12
        pPlayerTurn = pPlayerTurn + 1
        If pPlayerTurn > 6 Then
            pPlayerTurn = 1
        End If
        If CountCountriesOwned(pPlayerTurn) > 0 Then
            Exit For
        End If
    Next
    
    GetNextPlayer = pPlayerTurn
End Function

'Calculate the passed player's start of turn value.
Public Function GetPlayerValue(pPlayer As Integer) As Integer
    
    'Work out the player's initial points
    'from the number of countries owned.
    GetPlayerValue = GetInitialPoints(pPlayer)
    If GetPlayerValue = 0 Then
        Exit Function
    End If
    
    'Add the total held continent values.
    GetPlayerValue = GetPlayerValue + GetOccupiedContinentValues(pPlayer)
End Function

'Calculate the value of all continents that the passed player completely occupies.
Private Function GetOccupiedContinentValues(pPlayer As Integer) As Integer
    Dim vIndex As Integer
    
    GetOccupiedContinentValues = 0
    For vIndex = 0 To 5
        If OwnContinent(vIndex + 1, pPlayer) Then
            GetOccupiedContinentValues = GetOccupiedContinentValues + udContVal(vIndex).Value
        End If
    Next
End Function

'Return True if the passed continent held by the passed player.
Public Function OwnContinent(whichCont As Integer, Player As Integer) As Boolean
    Dim cntr As Integer
    
    For cntr = Continents(whichCont - 1).FirstCountry To Continents(whichCont - 1).LastCountry
        If gCountryOwner(cntr) <> gPlayerTurn Then
            OwnContinent = False
            Exit Function
        End If
    Next cntr
    OwnContinent = True
End Function

'Return True if the passed continent is completley held by any enemy to the passed player.
Public Function ContHeldByEnemy(pContinent As Integer, pFriendlyPlayer As Integer) As Boolean
    Dim vIndex As Integer
    Dim vFirstCtryOwner As Integer
    
    ContHeldByEnemy = True
    
    'Who owns the first country in the passed continent.
    vFirstCtryOwner = gCountryOwner(Continents(pContinent - 1).FirstCountry)
    
    'Jump out if it is the passed player.
    If vFirstCtryOwner = pFriendlyPlayer Then
        ContHeldByEnemy = False
        Exit Function
    
    'Check if the first country owner ownes all countries in the continent.
    Else
        For vIndex = Continents(pContinent - 1).FirstCountry + 1 To Continents(pContinent - 1).LastCountry
            If gCountryOwner(vIndex) <> vFirstCtryOwner Then
                ContHeldByEnemy = False
                Exit For
            End If
        Next vIndex
    End If
End Function

'Calculate the passed player's initial value at the start of the turn.
'Continents values are not included here.
Public Function GetInitialPoints(pPlayer As Integer) As Integer
    Dim vCountriesOwned As Integer
    
    'Count the number of countries owned.
    vCountriesOwned = CountCountriesOwned(pPlayer)
    
    'Calculate initial player value. Default vCountriesOwned \ 8 + 3, minimum of 3
    If vCountriesOwned > 0 Then
        GetInitialPoints = ((vCountriesOwned \ udNewUnitClac(0).Value) * udNewUnitClac(1).Value)
        If GetInitialPoints < udNewUnitClac(2).Value Then
            GetInitialPoints = udNewUnitClac(2).Value
        End If
    End If
End Function

'Count countries occupied by the passed player.
Public Function CountCountriesOwned(pPlayer As Integer) As Integer
    Dim vIndex As Long
    
    CountCountriesOwned = 0
    For vIndex = 1 To 42
        If gCountryOwner(vIndex) = pPlayer Then
            CountCountriesOwned = CountCountriesOwned + 1
        End If
    Next
End Function

    'Move mode
Private Sub cmdMove_Click()
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    Call MoveClicked
    Call SyncForgroundMap("cmdMove_Click")
End Sub

Private Sub MoveClicked()
    Dim tstPress As Integer
    
    If gPauseActive Then Call ActivatePauseMode(False)
    If gCurrentMode = 13 Then    'Game was won
        Exit Sub
    ElseIf gCurrentMode = 18 Then
        Exit Sub
    ElseIf gMapSetupLock Then
        Exit Sub
    End If
    tstPress = (GetPlayerController(gPlayerTurn) <> 0) _
                And (Not gComputerPressed)
    If tstPress Then
        Exit Sub
    End If
    Call CheckWinDuringTurn(gPlayerTurn)
    notHitMove = False                  'Has hit a key
    gCurrentMode = 10
    cmdAttack.Enabled = False
    mnuAttack.Enabled = cmdAttack.Enabled
    Call ColorCountryUnderAttack(0)
        
    InfoBoxPrint 0                   'cls
    InfoBoxPrint 9, gPlayerTurn
    InfoBoxPrint 3, 1
    InfoBoxPrint 1, 73               'moves
    InfoBoxPrint 5
    InfoBoxPrint 1, 74               'to
    InfoBoxPrint 6
    InfoBoxPrnCR 3, 1
    InfoBoxPrnCR 7
    InfoBoxPrnCR 7
    If gPlayerID(gPlayerTurn).playerWho = 0 Then
        InfoBoxPrnCR 1, 72               '<Click destination country>
    End If
    
    Call addChangeToList(gTargetCtry, 0, 0)
End Sub

Private Sub cmdSUPcncl_Click()
    Dim dummy As Long
    Dim i As Long
    Dim cntr As Integer
    
    'If netWorkSituation = cNetClient And gPlayerTurn <> gNewWarPlayerTurn Then
    If netWorkSituation = cNetClient Then
        Call joinWar
        Exit Sub
    End If
    
    If frmStats.Visible Then
        frmStats.Hide
    End If
    
    If (nmbrOfPlayers = 0) Or (gCountryOwner(1) = 0) Then
        Exit Sub
    End If
    
    SetupScreen.Visible = False
    Call ChangeTitlebarText(Phrase(35) + Trim(warSit.filename))     'Global Siege -
    Call ShowMenuBar(Not mnuFullScreen.Checked Or SetupScreen.Visible)
    
    gCurrentMode = gPreviousSettings.PrevMode
    If gPreviousSettings.PrevBorder Then
        TheMainForm.BackColor = gPreviousSettings.PrevBorder
    End If
    If gCurrentMode <> 13 And gCurrentMode <> 18 Then
        If IsComputerPlayer(gPlayerTurn) Then
            Timer2.Enabled = True
            Timer1.Enabled = True
        End If
    End If
    chkCardsVulture.Value = Abs(cardstmp)
    With A2opportunity
        .IsActive = False
        .pathPointer = 0
        .Path(0) = 0
    End With
    
    For cntr = 0 To 5
        gPlayerID(cntr + 1).startWith = txtPlayerStartCountries(cntr).Text
        gPlayerID(cntr + 1).playerWho = playerSelect_getIndex(cntr)
    Next cntr
    
    warSit.SetupScreen = False
    'Call mnuOptUndo_Click
    If netWorkSituation = cNetHost Then
        Call netMain.CancleSetup
    End If
    'mnuMsnMissionOn.Enabled = (SetupScreen.Visible And netWorkSituation <> cNetClient)
    Call EnableMissionOptions

    Timer1.Enabled = True
    Timer2.Enabled = IsComputerPlayer(gPlayerTurn)
End Sub

    'Reprints the map
Public Sub refreshMap()
    Dim CountryNumber As Integer
    
    'Picture1.Cls
    'Picture1.Print ""
    Mask4.Map1.Cls
    Mask4.Map1.Print ""
    For CountryNumber = 1 To 42
        Call ColorCountry(CountryNumber, gPlayerID(gCountryOwner(CountryNumber)).lngColor)
    Next CountryNumber
    Call SnapEmptyCardArea
    Call SnapEmptyDiceArea
    Call DrawLittleCards
    Call resetChangeList
    If tmrFlashInfoBox.Tag <> CStr(gPlayerTurn) Then
        tmrFlashInfoBox.Tag = ""
    End If
    
    'Make sure the viewport gets refreshed.
    gSyncViewportNeeded = True
    
End Sub

'** Not actually used.
'Reprints the map and marks who is under attack.
Public Sub RefreshClientMap()
    Dim CountryNumber As Integer
    
    Picture1.Cls
    Picture1.Print ""
    For CountryNumber = 1 To 42
        If CountryNumber = CountryUnderAttack Then
            Call ColorCountryUnderAttack(CountryNumber)
        Else
            Call ColorCountry(CountryNumber, gPlayerID(gCountryOwner(CountryNumber)).lngColor)
        End If
    Next CountryNumber
    Call DrawLittleCards
    Call resetChangeList
End Sub

'Save the current war to file.
Private Sub SaveWar(pWarFilePath As String, pWarTitle As String, pWarDescription As String, pWarLocked As Boolean)
    Dim vIndex As Integer
    
    On Error GoTo writeError
    warSit.filename = pWarTitle
    warSit.fileDescription = pWarDescription
    warSit.Locked = pWarLocked
    
    For vIndex = 0 To 5
        warSit.playerStart(vIndex) = txtPlayerStartCountries(vIndex).Text
        warSit.playerCtrlr(vIndex) = playerSelect_getIndex(vIndex)
    Next
    
    With warSit
        .capture = chkCardsVulture.Value
        .cardMax = txtMaximumCardValue.Text
        .cardmode = GetCardMode
        .crdHidden = chkCardsHidden.Value
        .firstPlayer = optPlr1FirstPlayer.Value
        .warFastWar = optnFastWar.Checked
        .warMissions = chkMsnMissionsOn.Value * &H1 _
                    + chkMsnArmyWipeout.Value * &H2 _
                    + chkMsnConquerHold.Value * &H4 _
                    + chkMsnMustComplete.Value * &H8 _
                    + chkMsnWinImmediate.Value * &H10 _
                    + chkMsnAreUnique.Value * &H20
        .warSupply = optSupplyLines.Value
        .warSupplyLimit = optLimitSupply.Value
        .warSupplyNo = optNoSupply.Value
        .nmbrPlrs = txtStartingArmies.Text
        .SetupScreen = SetupScreen.Visible Or gCurrentMode = 3
        .tmpTimer1 = tmpTimer1
        .tmpTimer2 = tmpTimer2
        .cardstmp = cardstmp
        .PrevMode = gPreviousSettings.PrevMode
        .sGateDefence = 2 'gateDefence
        .chkExtraStartingUnits = chkExtraStartingUnits.Value
        .distUnits = CByte(CLng(txtExtraStartingUnits.Text) Mod 250)
        
        'Dice setup info.
        For vIndex = 0 To optDiceRules.Count - 1
            .optDiceRules(vIndex) = optDiceRules(vIndex).Value
        Next
        
        .chkSortDice = chkSortDice.Value
        
        For vIndex = 0 To optDiceSame.Count - 1
            .optDiceSame(vIndex) = optDiceSame(vIndex).Value
        Next
        
        For vIndex = 0 To udDiceThrown.Count - 1
            .udDiceThrown(vIndex) = udDiceThrown(vIndex).Value
        Next
        
        'Card setup info.
        For vIndex = 0 To udCardDeck.Count - 1
            .udCardDeck(vIndex) = udCardDeck(vIndex).Value
        Next
        
        For vIndex = 0 To udFixedValues.Count - 1
            .udFixedValues(vIndex) = udFixedValues(vIndex).Value
        Next
        
        'Reinforcments tab.
        For vIndex = 0 To udNewUnitClac.Count - 1
            .udNewUnitClac(vIndex) = udNewUnitClac(vIndex).Value
        Next
        
        For vIndex = 0 To udContVal.Count - 1
            .udContVal(vIndex) = udContVal(vIndex).Value
        Next
    End With
    
    If warSit.sCtryOwner(1) = 0 Then
        Call SaveCheckpoint
    End If
    
    Call SaveWarFile(pWarFilePath, warSit)
    Exit Sub
writeError:
    MsgBox Err.Description, vbCritical, "Error in SaveWar()"
    Resume Next
End Sub

'Get war from file and make it the current war. Backward compatibility is
'maintained with older file formats.
Private Function OpenWar(pWarFileName As String) As Boolean
    Dim vIndex As Integer
    
    If frmStats.Visible Then
        Unload frmStats
    End If
    
    If pWarFileName = "" Then
        OpenWar = False
        Exit Function
    End If
    If Dir(pWarFileName) = "" Then
        MsgBox pWarFileName & Phrase(44), vbCritical, Phrase(45)   'has been deleted or corrupted; Missing file
        OpenWar = False
        Exit Function
    End If
    
    OpenWar = LoadWarFile(pWarFileName, warSit)
    If Not OpenWar Then
        Exit Function
    End If
    
    With A2opportunity
        .IsActive = False
        .pathPointer = 0
        .Path(0) = 0
    End With
    
    Call resetStats
    
    With warSit
    
    For vIndex = 0 To 5
        Call ChangPlayerStartingCtryCount(vIndex, .playerStart(vIndex))
        If txtPlayerStartCountries(vIndex) < 1 Then
            pctClr(vIndex).Print " X"
        Else
            pctClr(vIndex).Cls
        End If

        Call playerSelect_showIndex(vIndex, Abs(.playerCtrlr(vIndex)))
        If net.playerOwner(vIndex) <> myTerminalNumber Then
            Call playerSelect_showIndex(vIndex, remoteIndex)
        End If
    Next
    
    'Sync the u/d control value with the player starting countries text boxes.
    Call SyncPlrStartingCountries
    
    Call RememberStartingMissions
    
    If .SetupScreen Then
        Call ChangeTitlebarText(Phrase(34) & Trim(.filename))     'Global Siege Set Up...
    Else
        Call ChangeTitlebarText(Phrase(35) & Trim(.filename))    'Global Siege -
    End If
    
    chkCardsVulture.Value = .capture
    udMaximumCardValue.Value = CInt(.cardMax)
    Call SetCardMode(Abs(.cardmode))
    txtMaximumCardValue.Enabled = (GetCardMode = 2)
    udMaximumCardValue.Enabled = (GetCardMode = 2)
    lblSetupMaxCardValue.Enabled = (GetCardMode = 2)
    chkCardsHidden.Value = .crdHidden
    optPlr1FirstPlayer.Value = .firstPlayer
    chkMsnMissionsOn.Value = .warMissions And 1
    optSupplyLines.Value = .warSupply
    optLimitSupply.Value = .warSupplyLimit
    optNoSupply.Value = .warSupplyNo
    
    'Starting armies.
    Call SetStartingArmyCount(CInt(.nmbrPlrs))
    SetupScreen.Visible = .SetupScreen
    Call ShowMenuBar(SetupScreen.Visible)
    chkExtraStartingUnits.Value = Abs(CInt(.chkExtraStartingUnits))
    udExtraStartingUnits.Value = CInt(.distUnits)
    
    'Check if extra card and dice options are valid by testing the version info
    'which is saved in the new file format. GSVersion is saved in function SaveCheckpoint().
    If .GSVersion <> "" Then
    
        For vIndex = 0 To optDiceRules.Count - 1
            optDiceRules(vIndex).Value = .optDiceRules(vIndex)
        Next
        
        chkSortDice.Value = .chkSortDice
        
        For vIndex = 0 To optDiceSame.Count - 1
            optDiceSame(vIndex).Value = .optDiceSame(vIndex)
        Next
        
        For vIndex = 0 To udDiceThrown.Count - 1
            udDiceThrown(vIndex).Value = .udDiceThrown(vIndex)
        Next
        
        'Card setup info.
        For vIndex = 0 To udCardDeck.Count - 1
            udCardDeck(vIndex).Value = .udCardDeck(vIndex)
        Next
        
        For vIndex = 0 To udFixedValues.Count - 1
            udFixedValues(vIndex).Value = .udFixedValues(vIndex)
        Next
        
        'Get extra mission options.
        chkMsnArmyWipeout.Value = Abs(CInt(GetBit(1, CByte(.warMissions))))
        chkMsnConquerHold.Value = Abs(CInt(GetBit(2, CByte(.warMissions))))
        chkMsnMustComplete.Value = Abs(CInt(GetBit(3, CByte(.warMissions))))
        chkMsnWinImmediate.Value = Abs(CInt(GetBit(4, CByte(.warMissions))))
        chkMsnAreUnique.Value = Abs(CInt(GetBit(5, CByte(.warMissions))))
        
        'Reinforcments tab. This is only valid for versions greater than 0.9.0085.
        If .GSVersion > "00090085" Then
            For vIndex = 0 To udNewUnitClac.Count - 1
                udNewUnitClac(vIndex).Value = .udNewUnitClac(vIndex)
            Next
            
            For vIndex = 0 To udContVal.Count - 1
                udContVal(vIndex).Value = .udContVal(vIndex)
            Next
        Else
            udNewUnitClac(0).Value = 3
            udNewUnitClac(1).Value = 1
            udNewUnitClac(2).Value = 3
            udContVal(0).Value = 5
            udContVal(1).Value = 2
            udContVal(2).Value = 5
            udContVal(3).Value = 3
            udContVal(4).Value = 7
            udContVal(5).Value = 2
        End If
    'Old file format. Use hard coded default settings.
    Else
        
        'Dice tab.
        optDiceRules(2).Value = True
        chkSortDice.Value = 1
        optDiceSame(0) = True
        udDiceThrown(0).Value = 3
        udDiceThrown(1).Value = 2
        
        'Card tab.
        udCardDeck(0).Value = 14
        udCardDeck(1).Value = 14
        udCardDeck(2).Value = 14
        udCardDeck(3).Value = 2
        
        udFixedValues(0).Value = 4
        udFixedValues(1).Value = 6
        udFixedValues(2).Value = 8
        udFixedValues(3).Value = 10
        
        'Reinforcments tab.
        udNewUnitClac(0).Value = 3
        udNewUnitClac(1).Value = 1
        udNewUnitClac(2).Value = 3
        udContVal(0).Value = 5
        udContVal(1).Value = 2
        udContVal(2).Value = 5
        udContVal(3).Value = 3
        udContVal(4).Value = 7
        udContVal(5).Value = 2
        
        'Mission tab.
        chkMsnArmyWipeout.Value = vbChecked
        chkMsnConquerHold.Value = vbChecked
        chkMsnMustComplete.Value = vbUnchecked
        chkMsnWinImmediate.Value = vbUnchecked
        chkMsnAreUnique.Value = vbChecked
        
    End If
    
    'mnuMsnMissionOn.Enabled = (SetupScreen.Visible And netWorkSituation <> cNetClient)
    If chkMsnMissionsOn.Value = vbUnchecked _
    And chkMsnArmyWipeout.Value = vbUnchecked _
    And chkMsnConquerHold.Value = vbUnchecked Then
        chkMsnArmyWipeout.Value = vbChecked
        chkMsnConquerHold.Value = vbChecked
    End If
    
    Call EnableMissionOptions
    
    If (.sCtryOwner(1) = 0) Or (.sNmbrOfPlayers = 0) Then
        Exit Function
    End If
    tmpTimer1 = .tmpTimer1
    tmpTimer2 = .tmpTimer2
    cardstmp = .cardstmp
    If .SetupScreen Then
        gPreviousSettings.PrevMode = 2
    End If
    End With
    Exit Function
fileError:
    OpenWar = False
    MsgBox pWarFileName + Phrase(48), vbCritical, Phrase(56) ' has been corrupted and cannot be opened.; File error
    Resume Next
End Function

    'Start new game - setup
Private Sub setupNewGame()
    Dim cntr As Long
    Dim i As Long
    
    'Global veriable gNewWarPlayerTurn is only used in resume war (cmdSUPcncl).
    gNewWarPlayerTurn = gPlayerTurn

    gPreviousSettings.PrevMode = gCurrentMode
    gPreviousSettings.PrevBorder = TheMainForm.BackColor
    If frmMissions.Visible Then
        frmMissions.Hide
    End If

    cardstmp = chkCardsVulture.Value
    If gCurrentMode <> 13 Then
            TheMainForm.BackColor = &H8000000F
    End If

    gCurrentMode = 3
    Call ChangeTitlebarText(Phrase(34) + Trim(warSit.filename))      'Global Siege Set Up...
    'tfRate1.Value = True
    transferNmbr = 1
    SetupScreen.Visible = True
    Call ShowMenuBar(SetupScreen.Visible)

    'mnuMsnMissionOn.Enabled = (SetupScreen.Visible And netWorkSituation <> cNetClient)
    Call EnableMissionOptions
    
        'Disable smart players if evaluation has expired
    For cntr = 0 To 5
        Call ChangPlayerStartingCtryCount(CInt(cntr), gPlayerID(cntr + 1).startWith)
    Next
    
    'Sync the u/d control value with the player starting countries text boxes.
    Call SyncPlrStartingCountries
    
    txtExtraStartingUnits.Enabled = chkExtraStartingUnits.Value
    udExtraStartingUnits.Enabled = chkExtraStartingUnits.Value
End Sub

    'Enable Attack, Moove, Pass keys if TRUE, disable if FALSE
Public Sub ToggleKeys(OnOrOff As Boolean)
    If gCurrentMode = 6 Or gCurrentMode = 13 Or gCurrentMode = 18 Then     'Getting bodgy!!
        OnOrOff = False
    End If
    If GetPlayerController(gPlayerTurn) <> 0 Then OnOrOff = False
    cmdAttack.Enabled = OnOrOff And notHitMove
    cmdMove.Enabled = OnOrOff
    cmdEnd.Enabled = OnOrOff
    mnuAttack.Enabled = cmdAttack.Enabled
    mnuMove.Enabled = cmdMove.Enabled
    mnuPass.Enabled = cmdEnd.Enabled
End Sub

'Show or hide the card buttons.
Public Sub ToglleCardKeys(OnOrOff As Boolean)
    If TheMainForm.GetPlayerController(gPlayerTurn) <> 0 Then
        OnOrOff = False
    End If
    cmdExchange.Visible = OnOrOff
    cmdCardCncl.Visible = OnOrOff
End Sub

'Show new countries consecutivelly in order of original placement.
'This sub clears the screen and sets up timer1 to colour countries individually.
Private Sub ShowNewMap()
    Picture1.Cls
    Picture1.Print ""
    Mask4.Map1.Cls
    Mask4.Map1.Print ""
    
    timerCounter = 1
    Timer2.Enabled = False
    Timer1.Enabled = True
    Timer1.Interval = playSpeed
    gCurrentMode = 3
End Sub

'Update the foreground map by syncing the background map.
'Arg "pMileStone" is only used during debugging. This sub should be called as few times as possible
'because it hammers the CPU when stretching a large image with the half_tone setting.
Public Sub SyncForgroundMap(Optional pMileStone As String)
    Dim vDummy As Long
    
    'Check to make sure the viewport needs to be refreshed.
    If Not gSyncViewportNeeded Or gHeadlessMode Then
        'Debug.Print "Exit SyncForgroundMap - gCurrentMode "; gCurrentMode, pMileStone
        Exit Sub
    End If
    
    If gCurrentMode <> 5 And gCurrentMode <> 8 And gCurrentMode <> 7 And gCurrentMode <> 18 _
    And netWorkSituation <> cNetClient Then
        Call DrawAllCards
    End If
    vDummy = DoBlt(TheMainForm.Picture1.hdc, 0, 0, TheMainForm.Picture1.ScaleWidth, TheMainForm.Picture1.ScaleHeight, _
            Mask4.Map1.hdc, 0, 0, Mask4.Map1.ScaleWidth, Mask4.Map1.ScaleHeight, vbSrcCopy)
    
    gSyncViewportNeeded = False
    
    If gCurrentMode <> 18 Then
        Call DrawLittleCardText
        Call PrintAllScores
    End If
    Call TheMainForm.Picture1.Refresh
End Sub

'Print all scores directly on the ViewPort immediately after a sync.
'Country score text that is white has a background shadow affect
'created by printing the score in black first slightly offset to
'the right by one pixel and down by one pixel and then printing the
'white score over the top in the correct position. This does not
'apply to scores that have a black foreground color.
Private Sub PrintAllScores()
    Dim vCountry As Integer
    Dim vTextOffsetX As Long
    Dim vTextOffsetY As Long
    
    If gCurrentMode = 3 Then
        Exit Sub
    End If
    
    'This situation may happen during startup.
    If gPlayerTurn = 0 Then
        gPlayerTurn = 1
    End If
    
    vTextOffsetY = Picture1.TextHeight("1234567890") / 2
    
    'Print the score of correct colour at the correct location for each country.
    For vCountry = 1 To 42
        'Print the text if the score is > 0.
        If gCtryScore(vCountry) > 0 Then
            'Figure out where to print from the left of the viewport.
            If InStr(1, CountryID(vCountry).PrintRules, "L") Then
                'Left most position.
                vTextOffsetX = 0
            ElseIf InStr(1, CountryID(vCountry).PrintRules, "R") Then
                'Right most position.
                vTextOffsetX = Picture1.TextWidth(CStr(gCtryScore(vCountry)))
            Else
                'Center score text.
                vTextOffsetX = Picture1.TextWidth(CStr(gCtryScore(vCountry))) / 2
            End If
            
            'If white foreground text, create text shadow effect
            'by printing black background text first slightly offset.
            If gCtryTextColor(vCountry) > RGB(80, 80, 80) Then
                Picture1.ForeColor = RGB(0, 0, 0)
                Picture1.CurrentX = ((CountryID(vCountry).printX) * gPictureMaskRatioX) - vTextOffsetX + 1
                Picture1.CurrentY = ((CountryID(vCountry).printY) * gPictureMaskRatioY) - vTextOffsetY + 1
                Picture1.Print gCtryScore(vCountry)
            End If
            
            'Print the foreground text.
            Picture1.ForeColor = gCtryTextColor(vCountry)
            Picture1.CurrentX = (CountryID(vCountry).printX * gPictureMaskRatioX) - vTextOffsetX
            Picture1.CurrentY = (CountryID(vCountry).printY * gPictureMaskRatioY) - vTextOffsetY
            Picture1.Print gCtryScore(vCountry)
        End If
    Next
End Sub

'Return X size conversion ratio of Picture1 / Mask4.Map1
Public Function GetPictureMaskRatioX() As Double
    GetPictureMaskRatioX = Picture1.ScaleWidth / Mask4.Map1.ScaleWidth
End Function

'Return Y size conversion ratio of Picture1 / Mask4.Map1
Public Function GetPictureMaskRatioY() As Double
    GetPictureMaskRatioY = Picture1.ScaleHeight / Mask4.Map1.ScaleHeight
End Function

'Doesn't work too well. Not used.
'Update a particular region of the map by synching that region with the background map.
'Call after coulourCountry.
Public Sub SyncCountry(pCtryNumber As Integer)
    Dim dummy As Long
    Dim vSrcHiddenX As Long
    Dim vSrcHiddenY As Long
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
    
        'The hidden map is now the source instead of being the destination.
        vSrcHiddenX = CInt(CountryID(pCtryNumber).destX)
        vSrcHiddenY = CInt(CountryID(pCtryNumber).destY)
        
        With CountryID(pCtryNumber)
        dummy = DoBlt(TheMainForm.Picture1.hdc, _
                vSrcHiddenX * gPictureMaskRatioX, vSrcHiddenY * gPictureMaskRatioY, CInt(.Width) * gPictureMaskRatioX, CInt(.Height) * gPictureMaskRatioY, _
                Mask4.Map1.hdc, _
                vSrcHiddenX, vSrcHiddenY, CInt(.Width), CInt(.Height), vbSrcCopy)
        End With
    End If
End Sub

'This country is ubder attack. Mark it.
'If 0 is passed, repaint the last country it's actual
'color and set the last country to 0.
Public Sub ColorCountryUnderAttack(CtryNumber As Integer)
    Dim dummy As Long
    Dim PrintColor As Long
    Static LastCountry As Integer
    
    'Colour the last country that was painted black to its current colour.
    If LastCountry <> CtryNumber And LastCountry <> 0 Then
        'Debug.Print "ColorCountryUnderAttack re-colour "; LastCountry, playerID(gCountryOwner(LastCountry)).lngColor
        'Debug.Print
        Call ColorCountry(LastCountry, gPlayerID(gCountryOwner(LastCountry)).lngColor)
        LastCountry = 0
    End If
    
    'If no country it to be painted black, set the last country to 0
    'so that it does not colour the last country again.
    LastCountry = CtryNumber
    If CtryNumber = 0 Then
        'Debug.Print "ColorCountryUnderAttack exit CtryNumber = 0"
        'Debug.Print
        Exit Sub
    End If
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
    
        'Black.
        Mask4.Map1.FillColor = RGB(0, 0, 0)
        
        'Make country black (DSna).
        dummy = BitBlt(Mask4.Map1.hdc, _
                CountryID(CtryNumber).destX, CountryID(CtryNumber).destY, _
                CountryID(CtryNumber).Width, CountryID(CtryNumber).Height, _
                Mask4.pctMaskArray(0).hdc, CountryID(CtryNumber).srcX, CountryID(CtryNumber).srcY, _
                &H220326)
        
        'Make sure the viewport gets refreshed.
        gSyncViewportNeeded = True
    End If
    
    'White text no matter where it is.
    gCtryTextColor(CtryNumber) = RGB(223, 223, 223)
    gCtryCurrentColor(CtryNumber) = RGB(0, 0, 0)
    ctryIsBlack = CtryNumber
    'Debug.Print "ColorCountryUnderAttack", CtryNumber, LastCountry, "EXIT OK"
End Sub

'Fill country(pCtryNumber) with color (pNewCtryColr) and update the change list
'unless pUpdateChangeList is set to false.
'Raster Operations - http://msdn.microsoft.com/en-us/library/aa932106.aspx
Public Sub ColorCountry(pCtryNumber As Integer, pNewCtryColr As Long, Optional pUpdateChangeList As Boolean = True)
    Static sLastCountryColored As Integer
    Static sLastColorPainted As Long
    Dim dummy As Long
    
    If gHeadlessMode _
    And pCtryNumber = sLastCountryColored _
    And sLastColorPainted = pNewCtryColr _
    And gCurrentMode > 10 _
    And gCtryCurrentColor(pCtryNumber) = pNewCtryColr Then
        'Does not need to be painted again.
        'Debug.Print gCurrentMode
        Exit Sub
    End If
    
    sLastCountryColored = pCtryNumber
    sLastColorPainted = pNewCtryColr
    
    Mask4.Map1.FillColor = pNewCtryColr
    Mask4.Map1.ForeColor = vbBlack
    
    'If gCurrentMode >= 10 Then
    '    Debug.Print pCtryNumber, pNewCtryColr, gCurrentMode
    'End If
    
    'Make country black (DSna).
    dummy = BitBlt(Mask4.Map1.hdc, _
            CountryID(pCtryNumber).destX, CountryID(pCtryNumber).destY, _
            CountryID(pCtryNumber).Width, CountryID(pCtryNumber).Height, _
            Mask4.pctMaskArray(0).hdc, CountryID(pCtryNumber).srcX, CountryID(pCtryNumber).srcY, _
            &H220326)
        
    'Mix the mask with the colour and bilt to the map (DPSao).
    dummy = BitBlt(Mask4.Map1.hdc, _
            CountryID(pCtryNumber).destX, CountryID(pCtryNumber).destY, _
            CountryID(pCtryNumber).Width, CountryID(pCtryNumber).Height, _
            Mask4.pctMaskArray(gPlayerID(gCountryOwner(pCtryNumber)).MaskIndex).hdc, _
            CountryID(pCtryNumber).srcX, CountryID(pCtryNumber).srcY, _
            &HEA02E9)
    
    'Mark how this country was blitted incase it needs to be re-drawn.
    'This is actually used for re-drawing Australia when under the vulture cards.
    gCtryCurrentColor(pCtryNumber) = pNewCtryColr
    If ctryIsBlack = pCtryNumber Then
        ctryIsBlack = 0
    End If
    
    'Make sure the viewport gets refreshed.
    gSyncViewportNeeded = True
    
    'Update score colour and add changed colour to the change list for network updating.
    Call printScore(pCtryNumber, pUpdateChangeList)
End Sub

'Print destination and source to debug screen
Private Sub PrintDestAndSource()
    Dim cntr As Long
    
    Debug.Print "Country Name  ", "Dest X", "Dest Y", "Width", "Height", "Src X", "Src Y"
    For cntr = 1 To 42
        Debug.Print CountryID(cntr).ctryName; CountryID(cntr).destX, CountryID(cntr).destY, _
                    CountryID(cntr).Width, CountryID(cntr).Height, CountryID(cntr).srcX, CountryID(cntr).srcY
    Next
End Sub

'Return the text colout the passed country should be.
Private Function GetCountryTextColor(pCountry As Integer) As Long
    If InStr(1, CountryID(pCountry).PrintRules, "W") Then
        GetCountryTextColor = RGB(223, 223, 223)
    Else
        GetCountryTextColor = gPlayerID(gCountryOwner(pCountry)).txtColor
    End If
End Function

'Set the correct text colour for the passed country (countryNumber) and
'add the score to the change list for networking games if pUpdateChangeList is true.
Private Sub printScore(CountryNumber As Integer, Optional pUpdateChangeList As Boolean = True)
    
    gCtryTextColor(CountryNumber) = GetCountryTextColor(CountryNumber)
    
    'Make sure the viewport gets refreshed.
    gSyncViewportNeeded = True
        
    'Add to change list.
    If pUpdateChangeList Then
        Call addChangeToList(CountryNumber, 0, 0)
    End If
End Sub

'Replace old score with background color if required
'and print new score. Puts new score in countryID() variable.
Public Sub ChangeScoreUnderAttack(CountryNumber As Integer, newScore As Integer)
    
    'Put the new score.
    gCtryScore(CountryNumber) = newScore
    'Add to change list.
    Call addChangeToList(CountryNumber + 130, 0, 0)
End Sub

'Prints a new score of a country and puts new score
'in country' array position gCtryScore() (Updates gCtryScore)
Private Sub printNewScore(CountryNumber As Integer, newScore As Integer)
    
    If gCtryScore(CountryNumber) <> newScore Then
        gCtryScore(CountryNumber) = newScore
        Call printScore(CountryNumber)
    End If
End Sub

    'Print country score with specified print color (newColor)
Private Sub putScore(CountryNumber As Integer, NewColor As Long)
    Mask4.Map1.ForeColor = NewColor
    Mask4.Map1.CurrentX = CountryID(CountryNumber).printX
    Mask4.Map1.CurrentY = CountryID(CountryNumber).printY
    Mask4.Map1.Print gCtryScore(CountryNumber)
    
    'Make sure the viewport gets refreshed.
    gSyncViewportNeeded = True
End Sub

    'Change the language; update controlls
Public Sub SetNewWords()
    Dim cntr As Long
    Dim tmp As Long
    Dim tmp2 As Byte
    
    On Error GoTo ErrHand
    L = CInt(GetSetting(gcApplicationName, "settings", "Lang", eLanguage.English)) And &HFF
    
    'Set the text for the English button that appears on the bottom right of the screen.
    With cmdEnglish
    Select Case (GetSystemDefaultLangID And &HFF)
    Case eLanguage.English
        .Caption = "EN"
    Case eLanguage.Italian
        .Caption = "IT"
    Case eLanguage.German
        .Caption = "DE"
    Case eLanguage.Spanish
        .Caption = "ES"
    Case eLanguage.French
        .Caption = "FR"
    Case eLanguage.Swedish
        .Caption = "SE"
    Case eLanguage.Norwegian
        .Caption = "NO"
    Case eLanguage.Danish
        .Caption = "DK"
    Case Else
        .Caption = "EN"
    End Select
    End With
    
    'Choose a font size that will fit.
    Call ScaleInfoBoxFont
    pctInfoBox.Cls
    Call ScaleLittleCardFont
    
    If L = eLanguage.German Or L = eLanguage.PhraseFile Then    'German, change size of buttons
        cmdAttack.Width = 65
        cmdMove.Width = 80
        cmdEnd.Width = 80
    ElseIf L = eLanguage.French Then
        cmdAttack.Width = 60
        cmdMove.Width = 60
        cmdEnd.Width = 60
    Else
        cmdAttack.Width = 55
        cmdMove.Width = 55
        cmdEnd.Width = 55
    End If
    
    tmp2 = netWorkSituation
    netWorkSituation = cNetNone
    For cntr = 0 To 5
        tmp = playerSelect_getIndex(CInt(cntr)) 'PlayerSelect(cntr).ListIndex
        PlayerSelect(cntr).Clear

        PlayerSelect(cntr).AddItem Phrase(7), 0     'Human
        PlayerSelect(cntr).AddItem Phrase(28), 1    'Remote
        PlayerSelect(cntr).AddItem Phrase(8), 2     'Average computer
        PlayerSelect(cntr).AddItem Phrase(9), 3    'Smart computer
        PlayerSelect(cntr).AddItem Phrase(36), 4
        
        Call playerSelect_showIndex(CInt(cntr), CInt(tmp))
        gPlayerID(cntr + 1).strColor = Phrase(cntr + 1)
    Next cntr
    netWorkSituation = tmp2

    optRandomFirstPlayer.Caption = Phrase(12)
    optPlr1FirstPlayer.Caption = Phrase(13)
    
    lblSetupCards.Caption = " " & Phrase(14) & " "     'Cards
    chkCardsHidden.Caption = Phrase(15)
    optCardMode(0).Caption = Phrase(16)
    optCardMode(1).Caption = Phrase(17)
    optCardMode(2).Caption = Phrase(177)
    chkCardsVulture.Caption = Phrase(178)
    lblSetupMaxCardValue.Caption = Phrase(18)     'Maximum
    
    lblSetupBattleOptions.Caption = " " & Phrase(19) & " "     'Supply lines
    chkMsnMissionsOn.Caption = Phrase(20)
    optSupplyLines.Caption = Phrase(30)
    optLimitSupply.Caption = Phrase(21)
    optNoSupply.Caption = Phrase(179)
    'chkOptimizeDefenceDice.Caption = Phrase(22)
    'chkExtraStartingUnits.Caption = Phrase(186)
    
    cmdSetupOk.Caption = Phrase(26)
    mnuDeclareWar.Caption = Phrase(26)
    cmdSUPcncl.Caption = Phrase(27)
    mnuCancelSetup.Caption = Replace(Phrase(27), vbCrLf, " ")
    
    mnuFile.Caption = Phrase(31)
    mnuFileNew.Caption = Phrase(32)
    mnuFileReset.Caption = Phrase(175)
    mnuFileLoadWar.Caption = Phrase(144)
    mnuFileSaveWar.Caption = Phrase(149)
    mnuFileWarAs.Caption = Phrase(147)
    mnuFileExit.Caption = Phrase(33)
    
    mnuOptions.Caption = Phrase(38)
    optnFastWar.Caption = Phrase(39)
    mnuMisSeeReminder.Caption = Phrase(360)
    mnuOptStats.Caption = Phrase(347)
    mnu3Ddisplay.Caption = Phrase(42)
    mnuOptReport.Caption = Phrase(172)
    mnuOptLanguage.Caption = Phrase(173)
    mnuOptUndo.Caption = Phrase(174)
    
    mnuMission.Caption = Phrase(46)
    mnuMissionSee.Caption = Phrase(47)
    
    mnuHelp.Caption = Phrase(49)
    mnuHelpContents.Caption = Phrase(49) 'Phrase(50)
    'mnuHelpIndex.Caption = Phrase(51)
    hlpMRhome.Caption = Phrase(52)
    
    pctTransfer.Cls
    cmdAttack.Caption = Phrase(59)
    cmdMove.Caption = Phrase(60)
    cmdEnd.Caption = Phrase(61)
    mnuAttack.Caption = cmdAttack.Caption
    mnuMove.Caption = cmdMove.Caption
    mnuPass.Caption = cmdEnd.Caption
    cmdExchange.Caption = Phrase(62)
    cmdCardCncl.Caption = Phrase(63)
    
    'Load missions with language
    gMissions(0).DescriptionText = Phrase(75)   '"You must wipe out all other players and conquer the world."
    
    'Kill army missions.
    For cntr = 1 To 6
        With gMissions(cntr)
        .DescriptionText = Phrase(76) & Phrase(cntr + 361) & "."
        .WinMessageText = Phrase(cntr) & "."
        End With
    Next cntr
    
    'Conquer and hold missions.
    For cntr = 7 To 13
        With gMissions(cntr)
        .DescriptionText = Phrase(70 + cntr)
        .WinMessageText = Phrase(216 + cntr)
        End With
    Next cntr
    
    'Refresh the mission list.
    Call PopulateMissionList
    
    'Recent changes: Menu items
    mnuNet.Caption = Phrase(420) 'Networ&k setup...
    mnuNetChat.Caption = Phrase(269) 'Compose message...
    mnuOptToolbox.Caption = Phrase(201) '&Tool bar
    hlpContMap.Caption = Phrase(270) 'C&ontinent map
    mnuHelpAbout.Caption = Phrase(271)  '&About
    
    Call frmContinents.UpdateContDetails
    Call netMain.setLanguage
    Call selectToolTips
    
    'Put player names at top of screen
    Call DrawLittleCards
    
    'Update the text in the Reinforcement tab of the setup screen.
    Call PluralReinforcementTabText
    
    Exit Sub
ErrHand:
    L = 0
    Resume Next
End Sub

    'Change tool tip text
Public Sub selectToolTips()
    Dim cntr As Long
    Dim tmp As String
    
    For cntr = 0 To 5
        'tmp = Phrase(150) + Phrase(cntr + 1)
        tmp = Phrase(cntr + 1)
        'lblPlayerName(cntr).ToolTipText = tmp
        pctClr(cntr).ToolTipText = tmp
        PlayerSelect(cntr).ToolTipText = tmp
        txtPlayerStartCountries(cntr).ToolTipText = Phrase(151)
        'udPlayerStartCountries(cntr).ToolTipText = Phrase(152)
    Next cntr
    txtStartingArmies.ToolTipText = Phrase(154)
    optRandomFirstPlayer.ToolTipText = Phrase(155)
    optPlr1FirstPlayer.ToolTipText = Phrase(156)
    lblSetupCards.ToolTipText = Phrase(157)
    chkCardsHidden.ToolTipText = Phrase(158)
    optCardMode(0).ToolTipText = Phrase(159)
    optCardMode(1).ToolTipText = Phrase(160)
    optCardMode(2).ToolTipText = Phrase(161)
    txtMaximumCardValue.ToolTipText = Phrase(162)
    'udMaximumCardValue.ToolTipText = Phrase(162)
    lblSetupMaxCardValue.ToolTipText = Phrase(162)
    
    fSetupWarOptions.ToolTipText = Phrase(163)
    chkMsnMissionsOn.ToolTipText = Phrase(164)
    optSupplyLines.ToolTipText = Phrase(180)
    optLimitSupply.ToolTipText = Phrase(165)
    optNoSupply.ToolTipText = Phrase(165)
    'chkOptimizeDefenceDice.ToolTipText = Phrase(166)
    'chkFast.ToolTipText = Phrase(167)
    'chkFastDice.ToolTipText = Phrase(168)
    'chkBorder.ToolTipText = Phrase(169)
    chkExtraStartingUnits.ToolTipText = Phrase(187)
    'Label3.ToolTipText = Phrase(187)
    txtExtraStartingUnits.ToolTipText = Phrase(187)
    
    cmdSetupOk.ToolTipText = Phrase(170)
    cmdSUPcncl.ToolTipText = Phrase(171)
    chkCardsVulture.ToolTipText = Phrase(176)
    
End Sub

Private Sub distUnits_GotFocus()
    On Error Resume Next
    SetupScreen.SetFocus
End Sub

'Initialise debug codes and cheat codes & responses.
Private Sub InitialiseCheatCodes()
    
    'Add 50 points
    gCheatMode.inCodes(0) = "reinforce"
    gCheatMode.responses(0) = "50 extra units are at your disposal."
    
    'See other cards
    gCheatMode.inCodes(1) = "iseeall"
    gCheatMode.responses(1) = "All cards are now visible."
    
    'Create map, undo enabled
    gCheatMode.inCodes(2) = "icre8"
    gCheatMode.responses(2) = "Right click any country that you wish to modify."
    
    'See missions
    gCheatMode.inCodes(3) = "ispy"
    gCheatMode.responses(3) = "All missions are now visible."
    
    'Activate/deactivate auto restart.
    gCheatMode.inCodes(4) = "onandon"
    gCheatMode.responses(4) = "Auto restart activated."
    
    'Put GlobalSiege into testing mode.
    gCheatMode.inCodes(5) = "testing122"
    gCheatMode.responses(5) = "<By your command>"
    
    'Change TheMainForm to a standard size for taking screenshots.
    gCheatMode.inCodes(6) = "model"
    gCheatMode.responses(6) = "You'r a FOX baby YEAH!"
    
    'Change TheMainForm to a standard size for taking screenshots.
    gCheatMode.inCodes(7) = "loglevel9"
    gCheatMode.responses(7) = "Logging like crazy."
End Sub

'Initialize Global Siege.
Private Sub Form_Load()
    Dim vIndex As Long
    Dim vEncodedBits As Byte
    Dim vLanguage As String

    On Error Resume Next
    
    'Initialise the random number generator and get it ready for its first use.
    Call InitialiseRandomNumberGenerator
    Randomize
    
    'Initialise the FileIO module.
    Call InitialiseFileIO
    
    LogInfo "TheMainForm.Form_Load", "Loading GlobalSiege."
    
    'Set the form's caption.
    Me.Caption = SubstituteStringTokens(Me.Caption)
    
    'Change the dropdown menu captions.
    hlpMRhome.Caption = SubstituteStringTokens(hlpMRhome.Caption)
    
    'Set the global language variable.
    'Set to -1 if the language is not already set in the registry.
    vLanguage = GetSetting(gcApplicationName, "settings", "Lang", "No language selected")
    If IsNumeric(vLanguage) Then
        gLanguage = CLng(vLanguage) And &HFF
        Call LoadPhrases
    Else
        'Show the language selector if there is no language set in the registry.
        frmLanguage.Show vbModal
    End If
    
    Call SetNewWords
    
    'Get a basic checksum of the executable file.
    evalChk.fileCS = CStr(GetFileCS(App.Path & "\" & App.EXEName & "." & "exe"))
    
    'Load various constants from Mask4.
    Call Mask4.LoadMaskConstants
    
    'Load the previousley calculated dice odds into the gDiceOdds() array.
    Call InitialiseDiceOddsArray
    
    mnuNetCntr.Enabled = False
    mnuNetCntr.Checked = False
    mnuOptUndo.Enabled = False
    clickLock = False
    gPauseActive = False
    net.highestPriority = 0
    ReDim net.changeList(0) As Byte
    ReDim net.pctInfoByt(0) As Byte
    net.xmitInfTxt = True
    netWorkSituation = cNetNone
    Call resetPlayerOwners
    
    'Set up cheat codes and responses and switch on if in testing mode.
    Call InitialiseCheatCodes
    gCheatMode.createMap = gcAppTestingMode
    gCheatMode.seeMissions = gcAppTestingMode
    gCheatMode.testing = gcAppTestingMode
    gCheatMode.seeCards = gcAppTestingMode
    gCheatMode.undoEnabled = gcAppTestingMode
    
    'All player stats are invalid unless a whole game is played.
    Call ValidateAllStats(False)
    net.setupControlChange = True
    
    gTargetCtry = 1
    gSourceCtry = 1
    
    On Error GoTo ErrHand
    
    'Set hidden pictures and masks to full size.
    Mask4.pctCardSource.AutoSize = True
    Mask4.Map1.AutoSize = True
    Mask4.pctLittleCards.Width = Mask4.Map1.Width
    For vIndex = 0 To Mask4.pctMaskArray.Count - 1
        Mask4.pctMaskArray(vIndex).AutoSize = True
    Next
    
    Me.WindowState = GetSetting(gcApplicationName, "settings", "state", "0")
    gBorderWidth = GetSetting(gcApplicationName, "settings", "BorderWidth", 100)
    mnuViewQualityDisplay.Checked = GetSetting(gcApplicationName, "settings", "SmoothDisplay", mnuViewQualityDisplay.Checked)
    mnuViewBorder.Checked = GetSetting(gcApplicationName, "settings", "Border", mnuViewBorder.Checked)
    Picture1.FontBold = GetSetting(gcApplicationName, "settings", "FontBold", Picture1.FontBold)
    Picture1.FontItalic = GetSetting(gcApplicationName, "settings", "FontItalic", Picture1.FontItalic)
    Picture1.FontName = GetSetting(gcApplicationName, "settings", "FontName", Picture1.FontName)
    mnu3Ddisplay.Checked = GetSetting(gcApplicationName, "settings", "3Ddisplay", mnu3Ddisplay.Checked)
    mnuAutoRestart.Checked = GetSetting(gcApplicationName, "settings", "AutoRestart", mnuAutoRestart.Checked)
    mnuFlashInfoBox.Checked = GetSetting(gcApplicationName, "settings", "FlashInfoBox", mnuFlashInfoBox.Checked)
    hlpCheckForUpdates.Checked = GetSetting(gcApplicationName, "settings", "CheckForUpdates", hlpCheckForUpdates.Checked)
    
    pctInfoBox.FontName = Picture1.FontName
    pctInfoBox.FontBold = False
    pctInfoBox.FontItalic = Picture1.FontItalic
    
    mnuOptToolbox.Checked = GetSetting(gcApplicationName, "settings", "toolBox", "True")
    Toolbar1.Visible = mnuOptToolbox.Checked
    
    SetupScreen.Top = 0
    SetupScreen.Left = 0
    
    If GetDeviceCaps(Picture1.hdc, 8) < 1000 Then
        TheMainForm.Left = 0
        TheMainForm.Top = 0
        TheMainForm.Height = TheMainForm.Height - 620
    End If
    Call CheckAppFontAndSize
    TheMainForm.BackColor = &H8000000F
    
    'Set computer players with different defence levels.
    For vIndex = 1 To 6
        playerDefence(vIndex) = GenRandom4 * 4.5 + 0.5
    Next
    
    gAskedToSeeMission = False
    Picture1.AutoRedraw = True
    
    'Set current mode (0=Test mode).
    gCurrentMode = 1
    
    'Nothing picked up yet.
    gPickedUpUnits = 0
    
    'Set the transfer rate
    transferNmbr = 1
    gMoveLimit = 1
    'optimizeDice = False
    
    'Set cards to fixed.
    Call SetCardMode(1)
    
    'Set the maximum value for cards.
    gMaxCardValue = CInt(txtMaximumCardValue)
    
    txtExtraStartingUnits.Enabled = (chkExtraStartingUnits.Value = vbChecked)
    udExtraStartingUnits.Enabled = (chkExtraStartingUnits.Value = vbChecked)
    Call ToggleKeys(False)
    Call ToglleCardKeys(False)
    Call getPlayers
    Call InitializeMissions
    notHitMove = True
    playSpeed = playSlow
    diceSpeed = diceSlow
    flashingBorder = True
    'optimizeDice = True
    mnuNetDisconnect.Enabled = False
    boolDrawnWin = True
    vEncodedBits = CByte(GetSetting(gcApplicationName, "settings", "supPersonal", 44))
    Call changeWarSpeed(GetBit(0, vEncodedBits))
    mnuOptStats.Checked = Not GetBit(4, vEncodedBits)   'Inverse to allow default on
    mnuMisSeeReminder.Checked = Not GetBit(6, vEncodedBits)
    mnuOptReport.Checked = GetSetting(gcApplicationName, "settings", "seeReshuf")
    
    vEncodedBits = CByte(GetSetting(gcApplicationName, "settings", "supAdvanced", 11))
    chkMsnArmyWipeout.Value = Abs(CInt(GetBit(0, vEncodedBits)))
    chkMsnConquerHold.Value = Abs(CInt(GetBit(1, vEncodedBits)))
    chkMsnMustComplete.Value = Abs(CInt(GetBit(2, vEncodedBits)))
    chkMsnWinImmediate.Value = Abs(CInt(GetBit(3, vEncodedBits)))
    chkMsnAreUnique.Value = Abs(CInt(GetBit(4, vEncodedBits)))
    mnuViewLFont.Checked = GetBit(7, vEncodedBits)
    
    On Error Resume Next
    
    'mnuMsnMissionOn.Enabled = (SetupScreen.Visible And netWorkSituation <> cNetClient)
    Call EnableMissionOptions
    
    'Initialize 3rd AI.
    Call setA3
    
    'Setup the tabbed Setup box.
    frameSetupControls.Width = pctSetopDeclareCancel.Left + pctSetopDeclareCancel.Width + 90 '9000
    frameSetupControls.Height = lblVersion.Top + lblVersion.Height + 17 '6495
    For vIndex = 0 To frameSetup.Count - 1
        frameSetup(vIndex).Move tabSetup.ClientLeft, tabSetup.ClientTop, _
                    tabSetup.ClientWidth, tabSetup.ClientHeight
    Next
    frameSetup(tabSetup.SelectedItem.Index - 1).ZOrder 0
    Call SyncForgroundMap("Form Load")
    lblVersion.Caption = Phrase(197) & SubstituteStringTokens("<Var.Maj>.<Var.Min>.<Var.Rev>")
    
    tmrWatchDog.Enabled = True
    
    'Fire up the networking system.
    Load netMain
    
    'Read any args from the command line if any.
    Call ReadPassedFile(Command)
    
    netMain.txtSesName.Text = GetSetting(gcApplicationName, "settings", "LastSesName", Trim(warSit.filename))
    
    'Only show if we are not running in headles mode.
    If Not gHeadlessMode Then
        Me.Show
    End If
    
    LogInfo "TheMainForm.Form_Load", "GlobalSiege loaded."
    
    Exit Sub

ErrHand:
    Call mnuFileNew_Click
    Resume Next
End Sub

'Try to open the file passed to the executable by pFileName line. If the file has
'a ".mrk" extension, it is a war file so open it. If the passed file ends with a
'".conf" or ".mrc" extension, it is a config file so read and follow it.
Private Sub ReadPassedFile(pFileName As String)
    Dim vFileName As String
    
    'Clean up any quotes around the file name.
    vFileName = Replace(pFileName, """", "")
    
    'If the passed file exists.
    If Len(vFileName) > 0 And Dir(vFileName) <> "" Then
        
        'If the passed file has a .conf extension
        If LCase(Right(vFileName, Len(".conf"))) = ".conf" _
        Or LCase(Right(vFileName, Len(".mrc"))) = ".mrc" Then
            
            'File extension ends in ".conf".
            Call ReadGsConfigFile(vFileName)
        Else
            
            'File extension is not ".conf", treat as a war file and open.
            gCurrentWarPath = vFileName
            Call OpenWarFile
        End If
        
    
    Else
        
        'No file was passed, set the current war file to that saved in the registry.
        gCurrentWarPath = GetSetting(gcApplicationName, "settings", "StartingWar", _
                                        App.Path & "\" & "Default war.mrk")
        Call OpenWarFile
    End If
    
End Sub

'Read and act on commands passed in the passed config file.
Private Sub ReadGsConfigFile(pFileName As String)
    Dim vConfigFile As String
    Dim vConfigLine() As String
    Dim vIndex As Long
    Dim vHold As String
    Dim vCommand As String
    
    On Error Resume Next
    
    LogInfo "ReadGsConfigFile", "Reading config file: " & pFileName
    
    'Get the file contents.
    vConfigFile = ReadTextFile(pFileName)
    
    'Split on cr.
    vConfigLine = Split(vConfigFile, vbCrLf)
    
    For vIndex = 0 To UBound(vConfigLine)
        
        'Remove all tebs and leading and trailing spaces.
        vConfigLine(vIndex) = Trim(Replace(vConfigLine(vIndex), vbTab, " "))
        
        'Ignore comments.
        If Mid(vConfigLine(vIndex), 1, 1) <> "#" Then
            vCommand = LCase(GetListElement(vConfigLine(vIndex), 0, "="))
            Select Case vCommand
            
            'Server_Mode: Set gServerMode global.
            Case "server_mode"
                gServerMode = LCase(GetListElement(vConfigLine(vIndex), 1, "=")) = "true"
                If gServerMode Then
                    mnuViewQualityDisplay.Checked = False
                    TheMainForm.Height = 7875
                    TheMainForm.Width = 9885
                    netMain.optHost.Value = True
                    netMain.optInet.Value = True
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
            
            'Headless
            Case "headless"
                gHeadlessMode = LCase(GetListElement(vConfigLine(vIndex), 1, "=")) = "true"
                If gHeadlessMode Then
                    TheMainForm.Visible = False
                    TheMainForm.Hide
                End If
            
            'Start_Delay
            Case "start_delay"
                vHold = GetListElement(vConfigLine(vIndex), 1, "=")
                If IsNumeric(vHold) Then
                    Call Sleep(CLng(vHold) * 1000)
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
                
            'Server_ID: Set the global variable gInstanceID to include in log file names.
            Case "server_id"
                vHold = GetListElement(vConfigLine(vIndex), 1, "=")
                If IsNumeric(vHold) Then
                    gInstanceID = "-" & vHold
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
            
            'Log_Level
            Case "log_level"
                vHold = GetListElement(vConfigLine(vIndex), 1, "=")
                If IsNumeric(vHold) Then
                    gAppLogLevel = CLng(vHold)
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
                
            'War_File: If the passed arg is a file, open it. If not, try to match
            'it with a war file by name in the app data directory and app path.
            Case "war_file"
                vHold = Trim(GetListElement(vConfigLine(vIndex), 1, "="))
                gCurrentWarPath = Trim(vHold)
                
                'If the arg is an existing file.
                If Dir(vHold) <> "" Then
                    
                    'If so then use it as the war file path.
                    gCurrentWarPath = vHold
                Else
                    
                    'If not then try to find a war file that matches the arg.
                    gCurrentWarPath = GetFilePathFromName(vHold)
                    
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
                Call OpenWarFile
                LogInfo "ReadGsConfigFile", "War file set to " & gCurrentWarPath
                
            'Auto_Restart
            Case "auto_restart"
                mnuAutoRestart.Checked = LCase(GetListElement(vConfigLine(vIndex), 1, "=")) = "true"
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
            
            'Session_Name
            Case "session_name"
                netMain.txtSesName.Text = DecodeNonAscii(GetListElement(vConfigLine(vIndex), 1, "="))
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
            
            'TCP_Port
            Case "tcp_port"
                vHold = GetListElement(vConfigLine(vIndex), 1, "=")
                If IsNumeric(vHold) Then
                    netMain.txtTcpPort.Text = vHold
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
            
            'UDP_Port
            Case "udp_port"
                vHold = GetListElement(vConfigLine(vIndex), 1, "=")
                If IsNumeric(vHold) Then
                    netMain.txtUdpPort.Text = vHold
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
            
            'Max_Armies_Per_Terminal
            Case "max_armies_per_terminal"
                vHold = GetListElement(vConfigLine(vIndex), 1, "=")
                If IsNumeric(vHold) Then
                    netMain.vscrollMaxPlayers.Value = CInt(vHold)
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
            
            'Max_Connections_Per_IP
            Case "max_connections_per_ip"
                vHold = GetListElement(vConfigLine(vIndex), 1, "=")
                If IsNumeric(vHold) Then
                    netMain.vscrollMaxConnections.Value = CInt(vHold)
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
            
            'Session_Password
            Case "session_password"
                netMain.txtPassword.Text = Trim(DecodeNonAscii(GetListElement(vConfigLine(vIndex), 1, "=")))
                If Len(netMain.txtPassword.Text) > 0 Then
                    netMain.chkPasswordSession.Value = vbChecked
                Else
                    netMain.chkPasswordSession.Value = vbUnchecked
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
                
            'Turn_Time_Limit
            Case "turn_time_limit"
                vHold = GetListElement(vConfigLine(vIndex), 1, "=")
                If IsNumeric(vHold) Then
                    netMain.vscrollTimeLimit.Value = CInt(vHold)
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
            
            'Welcome_Message
            Case "welcome_message"
                netMain.txtWelcomeMsg.Text = Trim(DecodeNonAscii(GetListElement(vConfigLine(vIndex), 1, "=")))
                If Len(netMain.txtWelcomeMsg.Text) > 0 Then
                    netMain.chkWlcmMsg.Value = vbChecked
                Else
                    netMain.chkWlcmMsg.Value = vbUnchecked
                End If
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
                
            'Fast_War
            Case "fast_war"
                Call changeWarSpeed(LCase(GetListElement(vConfigLine(vIndex), 1, "=")) = "true")
                LogInfo "ReadGsConfigFile", vConfigLine(vIndex)
            
            'Counter_On: Turn on Clem's counter.
            Case "counter_on"
                mnuNetCntr.Checked = LCase(GetListElement(vConfigLine(vIndex), 1, "=")) = "true"
                LogInfo "Counter_On", vConfigLine(vIndex)
                
            End Select
        End If
        
    Next
    
    'Begin the session if server mode was set to TRUE.
    If gServerMode Then
        
        'Wait between 1 and 10 seconds incase there are other sessions
        'automatically starting on the same server being triggered by
        'the same event. This is to stop flooding the web server.
        Sleep CLng(Rnd * 10000) + 1
        
        'Begin the Internet session.
        LogInfo "ReadGsConfigFile", "Connecting to the Indexing Server..."
        Call netMain.ConnectDisconnectBeginSession
        
        'Begin the war.
        Call cmdSetupOk_Click
    End If
End Sub

'Return the file path of the passed war file name.
Private Function GetFilePathFromName(pFileName As String) As String
    Dim vWars() As String
    Dim vFileName As String
    
    vFileName = pFileName 'EncodeNonAscii(pFileName)
    
    vWars = Split(ListFiles(App.Path, vFileName & gcWarFileExtension, gcWarFileExtension) _
            & ListFiles(GetWarDataDir, vFileName & gcWarFileExtension, gcWarFileExtension), vbCrLf)
    
    If UBound(vWars) > 0 Then
        GetFilePathFromName = GetListElement(vWars(0), 0, ",")
    Else
        GetFilePathFromName = ""
    End If
End Function

'Try to open the war file that is pointed to by gCurrentWarPath global. Called
'from function ReadPassedFile() above which is called from Form_Load event.
Private Sub OpenWarFile()
    
    If OpenWar(gCurrentWarPath) Then
        Call RevertToCheckpoint
    Else
        gCurrentWarPath = ""
        Call mnuFileNew_Click
    End If
    
    'Ensure the timers are off if the setup screen is visible.
    If SetupScreen.Visible Then
        Timer1.Enabled = False
        Timer2.Enabled = False
    End If
End Sub

    'Resets map and places players
Private Sub initializePlayers()
    Dim vIndex As Integer, cntr2 As Integer
    Dim rtrn As Integer, y1 As Integer, BigCntr As Integer
    
    A3.startOfNewTurn = True
    For BigCntr = 1 To 400              'Nmbr of times to try to get it right
        For vIndex = 1 To 42             'Clear the board
            gCtryOrder(vIndex) = 0
            gCountryOwner(vIndex) = 0
            'MapColor(vIndex) = &HFFFFFF
            gCtryScore(vIndex) = 0
        Next
        y1 = 1
        For vIndex = 1 To 6
            For cntr2 = 0 To 8      'Reset cards
                gPlayerID(vIndex).card(cntr2) = 0
            Next cntr2
            
            For cntr2 = 1 To gPlayerID(vIndex).startWith
                Do
                    rtrn = GenRandom4 * 41 + 1
                    If gCountryOwner(rtrn) = 0 Then
                        gCountryOwner(rtrn) = vIndex
                        'MapColor(rtrn) = playerID(vIndex).lngColor
                        gCtryScore(rtrn) = 1
                        gCtryOrder(y1) = rtrn
                        y1 = y1 + 1
                        Exit Do
                    End If
                Loop
            Next cntr2
        Next
        If isMapValid Then
            Exit For
        End If
    Next BigCntr
    Call distExtraUnits
End Sub

    'Randomly distribute extra units on starting countries
Private Sub distExtraUnits()
    Dim cntr As Long, bigCount As Long
    Dim extraUnits As Long
    Dim tmp As Integer
    Dim UnitsLeft(6) As Integer
    
    If chkExtraStartingUnits.Value = vbChecked Then
    
        extraUnits = CInt(txtExtraStartingUnits.Text)   'How many extra points to distibute
        For cntr = 1 To 6                   'reset extra points for each player
            If gPlayerID(cntr).startWith > 0 Then
                UnitsLeft(cntr) = extraUnits
            Else
                UnitsLeft(cntr) = 0
            End If
        Next cntr
        
        For cntr = 1 To 6
            Do While UnitsLeft(cntr) > 0
                tmp = GenRandom4 * 41 + 1
                If gCountryOwner(tmp) = cntr Then
                    gCtryScore(tmp) = gCtryScore(tmp) + 1
                    UnitsLeft(cntr) = UnitsLeft(cntr) - 1
                End If
            Loop
        Next cntr
    End If
End Sub

Private Function displayNewMap()
    Call refreshMap
End Function

    'Check map setup for unfair advantage
Private Function isMapValid() As Boolean
    Dim ctr1 As Integer
    Dim ctr2 As Integer
    Dim Plr As Integer
    For ctr1 = 0 To 5
        Plr = gCountryOwner(Continents(ctr1).FirstCountry)
        
        For ctr2 = Continents(ctr1).FirstCountry To Continents(ctr1).LastCountry
            isMapValid = False
            If Plr <> gCountryOwner(ctr2) Then
                isMapValid = True
                Exit For
            End If
        Next ctr2
        If isMapValid = False Then
            Exit Function
        End If
    Next ctr1
End Function

    'Set player data
Private Sub getPlayers()
    Dim cntr As Long
    
    'CardsUp = False                          'Show cards?
    nmbrOfPlayers = 6                       '2 to 6 players
    gCurrentCardValue = gcCardStartValue
    
    For cntr = 1 To 6
        gPlayerID(cntr).startWith = 7
    Next cntr
    
    Call SetDisplayMode
End Sub

Public Sub getMousePos(ByRef x As Single, ByRef y As Single)
    x = Me.ScaleX(CSng(gCurrentMousePosX), vbPixels, vbTwips)
    y = Me.ScaleY(CSng(gCurrentMousePosY), vbPixels, vbTwips)
End Sub

Private Sub hlpContMap_Click()
    'Dim tmpTim1 As Boolean
    'Dim tmpTim2 As Boolean
    
    'tmpTim1 = Timer1.Enabled
    'tmpTim2 = Timer2.Enabled
    'Timer1.Enabled = False
    'Timer2.Enabled = False
    On Error Resume Next
    frmContinents.Show , TheMainForm
    'Timer1.Enabled = tmpTim1
    'Timer2.Enabled = tmpTim2
End Sub

Private Sub mnu3Ddisplay_Click()
    Dim cntr As Long
    Dim vMessageBoxResult As VbMsgBoxResult
    
    On Error Resume Next
    mnu3Ddisplay.Checked = Not mnu3Ddisplay.Checked
    
    'Check if Windows can handle the high colours required for 3D mode. If not, display
    'a warning message. 2D display will be set in function SetDisplayMode().
    If mnu3Ddisplay.Checked And GetDeviceCaps(Picture1.hdc, 12) < 16 Then
        vMessageBoxResult = MsgBox(LimitTextWidth(Phrase(89), 50), vbOKOnly)
    End If
    
    Call SetDisplayMode
    
    If gCurrentMode = 3 Then
        Picture1.Cls
        Picture1.Print ""
        Mask4.Map1.Cls
        Mask4.Map1.Print ""
        Call ShowNewMap
    End If
    Call SyncForgroundMap("mnu3Ddisplay_Click")
End Sub

'Set colours for the different graphics modes available.
'Note that the glboal variable "gPlayerID(1-6)" cannot have its type changed in
'any way because it is used to save situations to file and if changed, will no longer
'be compatible with saved games.
Private Sub SetDisplayMode()
    Static sAlreadyAdked As Boolean
    Dim vMsgBoxResult As VbMsgBoxResult
    Dim vPlayer As Long
    Dim vMaskChoice As Long
    
    'Put the system colour depth into the global variable to save
    'asking Windows every time bitblt is used.
    gDeviceCaps12 = GetDeviceCaps(Picture1.hdc, 12)
    
    'Show a warning if the the display does not have high colours set and
    'set the 2D mask to the flat black and white blit stencil.
    If gDeviceCaps12 < 16 Then
        If Not sAlreadyAdked Then
            vMsgBoxResult = MsgBox(LimitTextWidth(Phrase(89), 50), vbOKOnly)
            sAlreadyAdked = True
        End If
        vMaskChoice = 0
        mnu3Ddisplay.Checked = False
    Else
        vMaskChoice = 4
    End If
    
    '3D display.
    If mnu3Ddisplay.Checked Then
        'The Red Army.
        With gPlayerID(1)
        .lngColor = RGB(255, 0, 0) 'RGB(255, 0, 0)
        .srfsLite = False               'Not actually used but could be used in the future.
        .MaskIndex = 1                  'Use blit mask #1
        .txtColor = RGB(223, 223, 223) '0
        .bkgndColor = RGB(255, 32, 32)
        gPlayerFlashColor(0) = RGB(255, 128, 128)
        End With
        
        'The Green Army.
        With gPlayerID(2)
        .lngColor = RGB(0, 255, 0)
        .MaskIndex = 1                  'Use mask 1
        .txtColor = 0
        .bkgndColor = RGB(88, 248, 88)
        gPlayerFlashColor(1) = RGB(192, 255, 192)
        End With
        
        'The Blue Army.
        With gPlayerID(3)
        .lngColor = RGB(255, 255, 255)  'Uses its own mask so set the fill colour white.
        .MaskIndex = 3                  'Use mask 3
        .txtColor = RGB(223, 223, 223)  'White score text.
        .bkgndColor = RGB(0, 209, 255)
        gPlayerFlashColor(2) = RGB(0, 128, 255)
        End With
        
        'The Yellow Army.
        With gPlayerID(4)
        .lngColor = RGB(255, 255, 0)
        .MaskIndex = 1                  'Use mask 1
        .txtColor = 0
        .bkgndColor = RGB(250, 242, 0)
        gPlayerFlashColor(3) = RGB(255, 255, 192)
        End With
        
        'The Purple Army.
        With gPlayerID(5)
        .lngColor = RGB(255, 0, 255)
        .MaskIndex = 1                  'Use mask 1
        .txtColor = 0
        .bkgndColor = RGB(255, 24, 255)
        gPlayerFlashColor(4) = RGB(255, 128, 255)
        End With
        
        'The Gray Army.
        With gPlayerID(6)
        .lngColor = RGB(255, 255, 255)  'Uses its own mask so set the fill colour white.
        .MaskIndex = 2                  'Use mask 2
        .txtColor = RGB(223, 223, 223) '0
        .bkgndColor = RGB(160, 160, 160)
        gPlayerFlashColor(5) = RGB(128, 128, 128)
        End With
    
    '2D display. All use mask #4 unless the system pallet cannot handle it,
    'in which case, use mask #0, the black and white blit stencil.
    Else
        'The Red Army.
        With gPlayerID(1)
        .lngColor = vbRed
        .MaskIndex = vMaskChoice
        .txtColor = RGB(223, 223, 223) '0
        .bkgndColor = vbRed
        gPlayerFlashColor(0) = RGB(255, 128, 128)
        End With
        
        'The Green Army.
        With gPlayerID(2)
        .lngColor = vbGreen
        .MaskIndex = vMaskChoice
        .txtColor = 0
        .bkgndColor = vbGreen
        gPlayerFlashColor(1) = RGB(192, 255, 192)
        End With
        
        'The Blue Army.
        With gPlayerID(3)
        .lngColor = RGB(0, 255, 255)
        .MaskIndex = vMaskChoice
        .txtColor = 0
        .bkgndColor = RGB(0, 255, 255)
        gPlayerFlashColor(2) = RGB(0, 128, 255)
        End With
        
        'The Yellow Army.
        With gPlayerID(4)
        .lngColor = vbYellow
        .MaskIndex = vMaskChoice
        .txtColor = 0
        .bkgndColor = vbYellow
        gPlayerFlashColor(3) = RGB(255, 255, 192)
        End With
        
        'The Purple Army.
        With gPlayerID(5)
        .lngColor = RGB(255, 0, 255)
        .MaskIndex = vMaskChoice
        .txtColor = RGB(223, 223, 223) '0
        .bkgndColor = RGB(255, 0, 255)
        gPlayerFlashColor(4) = RGB(255, 128, 255)
        End With
        
        'The Gray Army.
        With gPlayerID(6)
        .lngColor = RGB(127, 127, 127)  '128 breaks the mask mix.
        .MaskIndex = vMaskChoice
        .txtColor = RGB(223, 223, 223)
        .bkgndColor = RGB(180, 180, 180)
        gPlayerFlashColor(5) = RGB(128, 128, 128)
        End With
    End If
    
    'Change setup screen coulour and put an "X" in any un used players.
    'Name players
    For vPlayer = 0 To 5
        gPlayerID(vPlayer + 1).strColor = Phrase(vPlayer + 1)
        pctClr(vPlayer).BackColor = gPlayerID(vPlayer + 1).bkgndColor
        pctClr(vPlayer).Cls
        If txtPlayerStartCountries(vPlayer).Text < 1 Then
            pctClr(vPlayer).Print " X"
        End If
    Next vPlayer
    
    'And don't forget to refresh the map so the changes will
    'be seen on the next foreground sync.
    Call refreshMap
End Sub

    'Encrypt a string into decimal notation (max 4 digits)
Private Function Encrypt(str As String) As String
    Dim cntr As Long
    Dim rslt As Long
    Dim ln As Long
    Dim pos As Long
    
    ln = Len(str)
    rslt = 0
    pos = 1
    
    For cntr = 1 To ln
        rslt = rslt + pos * (Asc(Mid(str, cntr, 1)) - 30)
        pos = pos + 1
        If pos > 5 Then pos = 1
    Next
    Encrypt = Trim(CStr(rslt Mod 10000))
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim cntr
    Dim tmp As Byte
    
    On Error Resume Next
    
    gCancelUnload = 0
    
    Unload netMain
    
    'If player selected to not close netMain, cancel shutdown.
    If gCancelUnload = -1 Then
        Cancel = -1
        Exit Sub
    End If
    
    If (gCurrentMode <> 13) _
    And (gCurrentMode <> 18) _
    And (gCurrentMode <> 3) _
    And (Not SetupScreen.Visible) _
    And (gCancelUnload <> -2) _
    And Not gServerMode _
    And Not gHeadlessMode Then
        If MsgBox(Phrase(191), vbYesNo, gcApplicationName) = vbNo Then  'Do you really want to quit this war?
            Cancel = -1
            Exit Sub
        End If
    End If
    
    Call SetBit(chkMsnArmyWipeout.Value = vbChecked, 0, tmp)
    Call SetBit(chkMsnConquerHold.Value = vbChecked, 1, tmp)
    Call SetBit(chkMsnMustComplete.Value = vbChecked, 2, tmp)
    Call SetBit(chkMsnWinImmediate.Value = vbChecked, 3, tmp)
    Call SetBit(chkMsnAreUnique.Value = vbChecked, 4, tmp)
    Call SetBit(mnuViewLFont.Checked, 7, tmp)

    SaveSetting gcApplicationName, "settings", "supAdvanced", str(tmp)
        'Save the windows state - Maximised or minimised
        'Save reshuffle screen status
        'Save tool bar state
    If netWorkSituation <> cNetNone Then
        Call IxServerCloseSession(eIxCommand.CLEANUP_SESSION, True)
        'Unload netMain
    End If
    SaveSetting gcApplicationName, "settings", "state", Me.WindowState
    SaveSetting gcApplicationName, "settings", "BorderWidth", gBorderWidth
    SaveSetting gcApplicationName, "settings", "seeReshuf", mnuOptReport.Checked
    SaveSetting gcApplicationName, "settings", "ToolBox", mnuOptToolbox.Checked
    SaveSetting gcApplicationName, "settings", "Path", (App.Path)
    SaveSetting gcApplicationName, "settings", "AIblock", str(evalChk.TimeNow)
    SaveSetting gcApplicationName, "settings", "MRVersion", GetVersionInfo(".")
    SaveSetting gcApplicationName, "settings", "SmoothDisplay", mnuViewQualityDisplay.Checked
    SaveSetting gcApplicationName, "settings", "Border", mnuViewBorder.Checked
    SaveSetting gcApplicationName, "settings", "FontBold", Picture1.FontBold
    SaveSetting gcApplicationName, "settings", "FontItalic", Picture1.FontItalic
    SaveSetting gcApplicationName, "settings", "FontName", Picture1.FontName
    SaveSetting gcApplicationName, "settings", "3Ddisplay", mnu3Ddisplay.Checked
    SaveSetting gcApplicationName, "settings", "AutoRestart", mnuAutoRestart.Checked
    SaveSetting gcApplicationName, "settings", "FlashInfoBox", mnuFlashInfoBox.Checked
    SaveSetting gcApplicationName, "settings", "CheckForUpdates", hlpCheckForUpdates.Checked
    
    tmp = 0
    Call SetBit(optnFastWar.Checked, 0, tmp)
    Call SetBit(Not mnuOptStats.Checked, 4, tmp)
    Call SetBit(Not mnuMisSeeReminder.Checked, 6, tmp)
    SaveSetting gcApplicationName, "settings", "supPersonal", str(tmp)
    
    'Release the random number generator.
    Call ReleaseRandomNumberGenerator
    
    LogInfo "TheManiForm.Form_Unload", "Application end."
    
    'Forcefully shut down. "mrisk" is the app name when running
    'in the IDE during development.
    If App.EXEName = "mrisk" Then
        End
    Else
        ExitProcess 0
    End If
End Sub

'Open the Global Siege home page.
Private Sub hlpMRhome_Click()
    On Error Resume Next
    Call OpenWebPage(gHomeWebPage)
End Sub

'Open the passed web page. Not sure why it is done this way but
'it works reliably. TODO: this may be inproved in future versions.
Public Sub OpenWebPage(pUrl As String)
    Dim vTempFile As String, dummy As String
    Dim vBrowserExec As String * 255
    Dim vReturnVal As Long
    Dim vFileNumber As Integer
    
    On Error Resume Next
    vBrowserExec = Space(255)
    vTempFile = GetTmpDataDir & "\temphtm.HTM"

    vFileNumber = FreeFile()
    
    'Create a temp HTML file.
    Open vTempFile For Output As #vFileNumber
    Write #vFileNumber, " <\HTML>"
    Close #vFileNumber

    ' Then find the application associated with it.
    vReturnVal = FindExecutable(vTempFile, dummy, vBrowserExec)
    vBrowserExec = Trim$(vBrowserExec)
    
    
    'Could not find a browser.
    If vReturnVal <= 32 Or IsEmpty(vBrowserExec) Then ' Error
        MsgBox Phrase(215) + vbCrLf + vbCrLf _
                & gcDefaultHomePageClearURL
    
    'An application is found, launch it.
    Else
        vReturnVal = ShellExecute(TheMainForm.hWnd, "open", vBrowserExec, _
            pUrl, dummy, SW_SHOWNORMAL)
        
        'Error, Web Page '<url>' not Opened
        If vReturnVal <= 32 Then
            MsgBox Phrase(216)
        End If
    End If

    Kill vTempFile
End Sub

'Stop the info box from flashing.
Public Sub Picture1_Click()
    On Error Resume Next
    
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    If gPauseActive Then
        Call ActivatePauseMode(False)
    End If
    If gPlayerID(gPlayerTurn).playerWho = 0 Then
        Call ClickMapNow
        Call SyncForgroundMap("Picture1_Click")
    End If
End Sub

    'Trigger main loop
Public Sub ClickMapNow()
    If gPauseActive Then
        Call ActivatePauseMode(False)
    End If
    If clickLock Then
        Exit Sub
    End If
    clickLock = True
    If frmMissions.Visible Then
        frmMissions.Hide
    End If
    Call map1Clicked
    clickLock = False
End Sub

'Entry point for computer players.
Private Sub map1Clicked()
    Dim CountryNumber As Integer
    Dim newScore As Integer
    Dim lstPlayer As Integer
    
    If gMapSetupLock Or gWarRestartLock Then
        Exit Sub
    End If
    If gPlayerTurn = 0 Then gPlayerTurn = 1
    If (netWorkSituation <> cNetNone) _
    And (net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber) Then
        Exit Sub
    End If
    
    If gCurrentMode = 13 Or gCurrentMode = 18 Then
        Exit Sub
    End If
    
    Call GetCardHitPosition
    
    If gCurrentMode = 5 And gLastCardClicked = 0 Then         'Stop looking at cards
        Call Mode5
        Call handleUpdate
        If GetPlayerController(gPlayerTurn) = 0 Then
            net.madeUpdate = False
        End If
        Call SyncForgroundMap("map1Clicked 1")
        Exit Sub
    End If
    If GetPlayerController(gPlayerTurn) = 0 Then      'If human
        If (gPickedUpUnits = 0) And (gPlayerValue = 0) Then  'Stop possible lock-up
            ToggleKeys (True)
        End If
    End If
    lstPlayer = gPlayerTurn
    If GetPlayerController(gPlayerTurn) = 0 Then
        CountryNumber = GetMouseHitPosition(gCurrentMousePosX, gCurrentMousePosY)
    Else
        CountryNumber = AutoCountry     'Computer's turn
    End If
    
    If CountryNumber = 0 Then               'Not a country.
        If gLastCardClicked > 0 Then
            If gPlayerID(gPlayerTurn).card(0) > 0 Then
                Call LookAtCards                'Hit cards
            End If
        Else
            Call CheckWinDuringTurn(lstPlayer)
            Call HitLittleCards
            Call handleUpdate
            If GetPlayerController(gPlayerTurn) = 0 Then
                net.madeUpdate = False
            End If
        End If
    Else
        Call processMode(CountryNumber)
        Call CheckWinDuringTurn(lstPlayer)               'Has this player won?
        Call handleUpdate
        If GetPlayerController(gPlayerTurn) = 0 Then
            net.madeUpdate = False
        End If
    End If
End Sub

    'Return true if countryNmbr is on a border with playerNumber (country)
Private Function onAborder(CountryNumber As Integer, PlayerNumber As Integer) As Boolean
    Dim cntr1 As Integer
    Dim nbrNmbr As Integer
    
    For cntr1 = 1 To 7
        nbrNmbr = CountryID(CountryNumber).neighbour(cntr1)
        If nbrNmbr = 0 Then
            onAborder = False
            Exit Function
        ElseIf gCountryOwner(nbrNmbr) = PlayerNumber Then
            onAborder = True
            Exit Function
        End If
    Next cntr1
    onAborder = False
End Function

'Return true if pTargetCountry is on a border with a country owned by pPlayerNumber
'And pPlayerNumber has points in range to attack with.
Private Function CanIAttackThisCountry(pTargetCountry As Integer, pPlayerNumber As Integer) As Boolean
    Dim vNeighborIndex As Integer
    Dim vNeighbor As Integer
    
    CanIAttackThisCountry = False
    
    For vNeighborIndex = 1 To 7
        vNeighbor = CountryID(pTargetCountry).neighbour(vNeighborIndex)
        If vNeighbor = 0 Then
            Exit Function
        ElseIf gCountryOwner(vNeighbor) = pPlayerNumber _
        And gCtryScore(vNeighbor) > 1 Then
            CanIAttackThisCountry = True
            Exit Function
        End If
    Next
End Function

'Load rolled dice into the networking global for dispatch to client terminals.
Private Sub PackDiceForNet(pAttackDice() As Integer, pDefenceDice() As Integer)
    Dim vIndex As Long
    
    For vIndex = 0 To cMaxNumberOfDice - 1
        net.RolledAttackDice(vIndex) = CByte(pAttackDice(vIndex))
        net.RolledDefenceDice(vIndex) = CByte(pDefenceDice(vIndex))
    Next
End Sub

'Attack by rolling the dice, update the scores and audits, check for wins
'and print updates to the info box. This sub should really be split.
Private Sub AttackCountry()
    Dim vDefendingPlayer As Integer
    Dim vAttackDice(cMaxNumberOfDice - 1) As Integer
    Dim vDefenceDice(cMaxNumberOfDice - 1) As Integer
    Dim vAttackUnits As Integer
    Dim vDefendUnits As Integer
    Dim vUnitDiff As Integer
    
    
    vAttackUnits = gPickedUpUnits
    vDefendUnits = gCtryScore(gTargetCtry)
    
    'Roll dice.
    Call AttackRollDice(vAttackUnits, vDefendUnits, vAttackDice, vDefenceDice)
    
    'Pack the dice for remote terminals.
    Call PackDiceForNet(vAttackDice, vDefenceDice)
    
    'Work out the results of the thrown dice.
    Call DiceBattleDamage(vAttackUnits, vDefendUnits, vAttackDice, vDefenceDice)
    Call addChangeToList(gTargetCtry + 130, 0, 0)
    
    'Display the dice.
    Call DisplayDiceOnBoard(vAttackDice, vDefenceDice)
    
    'Update audits and stats for the attacker.
    vUnitDiff = gPickedUpUnits - vAttackUnits
    If vUnitDiff > 0 Then
        gPickedUpUnits = vAttackUnits
        gPlayerStats(gPlayerTurn).UnitsLost _
            = gPlayerStats(gPlayerTurn).UnitsLost + vUnitDiff
        gPlayerStats(gCountryOwner(gTargetCtry)).UnitsBeaten _
            = gPlayerStats(gCountryOwner(gTargetCtry)).UnitsBeaten + vUnitDiff
        Call AuditUpdateScore(gPlayerTurn, vUnitDiff)
    End If
    
    'Update audits and stats for the defender.
    vUnitDiff = gCtryScore(gTargetCtry) - vDefendUnits
    If vUnitDiff > 0 Then
        Call ChangeScoreUnderAttack(gTargetCtry, vDefendUnits)
        gPlayerStats(gCountryOwner(gTargetCtry)).UnitsLost _
            = gPlayerStats(gCountryOwner(gTargetCtry)).UnitsLost + vUnitDiff
        gPlayerStats(gPlayerTurn).UnitsBeaten _
            = gPlayerStats(gPlayerTurn).UnitsBeaten + vUnitDiff
        Call AuditUpdateScore(gCountryOwner(gTargetCtry), vUnitDiff)
    End If
    
    vDefendingPlayer = gCountryOwner(gTargetCtry)
    If gCtryScore(gTargetCtry) <= 0 Then       'Won battle
        
        'Update country stats.
        gPlayerStats(gCountryOwner(gTargetCtry)).CountriesLost _
            = gPlayerStats(gCountryOwner(gTargetCtry)).CountriesLost + 1
        gPlayerStats(gPlayerTurn).CountriesDefeated _
            = gPlayerStats(gPlayerTurn).CountriesDefeated + 1
        
        gCountryOwner(gTargetCtry) = gPlayerTurn
        Call printNewScore(gTargetCtry, gPickedUpUnits)
        'TP111 - first call to colorcountry.
        Call ColorCountry(gTargetCtry, gPlayerID(gPlayerTurn).lngColor)
        gPickedUpUnits = 0
        boolIssueCard = True
        Call ToggleKeys(True)               'Just incase
        Call CheckWinDuringTurn(gPlayerTurn)
        Call addChangeToList(gTargetCtry, 0, 2)
        If gCurrentMode = 13 Or gCurrentMode = 18 Then
            Exit Sub
        End If
        Call CheckVultureCard(vDefendingPlayer)
        If gPlayerID(gPlayerTurn).card(4) = 0 Then
            Call AttackClicked
        Else
            ToggleKeys (False)
        End If
        Exit Sub
    End If

    If gPickedUpUnits <= 0 Then                      'End of battle
        'TP111 - second call
        Call ColorCountry(gTargetCtry, gPlayerID(gCountryOwner(gTargetCtry)).lngColor)
        Call addChangeToList(gTargetCtry, 0, 1)
        If GetPlayerController(gPlayerTurn) > 0 Then
            gCurrentMode = 1
            Call ToggleKeys(True)
            
            'Stop following A2 path if I lost a battle.
            With A2opportunity
            If .IsActive Then
                .IsActive = False
                .pathPointer = 0
                .Path(0) = 0
            End If
            End With
            
            Exit Sub
        End If
        Call AttackClicked
        'Call handleUpdate
        Exit Sub
    End If
    
    InfoBoxPrint 0                   'cls
    InfoBoxPrint 9, gPlayerTurn
    InfoBoxPrint 3, 1
    InfoBoxPrnCR 1, 64               'attacks
    InfoBoxPrnCR 8, gTargetCtry
    InfoBoxPrint 5                   'bold
    InfoBoxPrint 1, 101              'with
    InfoBoxPrint 2, gPickedUpUnits
    If gPickedUpUnits > 1 Then
        InfoBoxPrnCR 1, 102          'units
    Else
        InfoBoxPrnCR 1, 103          'unit
    End If
    InfoBoxPrnCR 6                       'normal
    If gPlayerID(gPlayerTurn).playerWho = 0 Then
        InfoBoxPrnCR 1, 65                   '<Click defending country>
        InfoBoxPrnCR 1, 104                  '<Click defending country>
    End If
End Sub

'Limit moves between countries
Private Function limitMoves(trNmbr As Integer) As Integer
    Dim i As Long
    
    limitMoves = trNmbr
    If gMoveTimes(gSourceCtry) = 0 Then
        If gMoveTimes(gTargetCtry) = 0 Then      '0 to 0 , to-> 1
            gMoveTimes(gTargetCtry) = 1
            limitMoves = trNmbr
            If gMoveLimit = 2 Then               'Fill everyting with 2 if no moves
                For i = 1 To 42
                    If (i <> gSourceCtry) Then
                        gMoveTimes(i) = 2
                    End If
                Next
            End If
            Exit Function
        Else                                    '0 to 1
            limitMoves = trNmbr                 '0 to 2
            Exit Function
        End If
    ElseIf gMoveTimes(gSourceCtry) = 1 Then
        If gMoveLimit = 2 Then                   '1 move only
            limitMoves = 0
            Exit Function
        End If
        If gMoveTimes(gTargetCtry) < 2 Then        '1 to 0, to->2
            gMoveTimes(gTargetCtry) = 2            '1 to 1, to->2
            If gMovedIn(gSourceCtry) + trNmbr >= 5 Then
                trNmbr = 5 - gMovedIn(gSourceCtry)
                gMovedIn(gSourceCtry) = 5
                gMoveTimes(gSourceCtry) = 2
                limitMoves = trNmbr
                Exit Function
            Else
                gMovedIn(gSourceCtry) = gMovedIn(gSourceCtry) + trNmbr
                limitMoves = trNmbr
                Exit Function
            End If
        ElseIf gMoveTimes(gTargetCtry) = 2 Then    '1 to 2, from->2
            If gMovedIn(gSourceCtry) + trNmbr >= 5 Then
                trNmbr = 5 - gMovedIn(gSourceCtry)
                gMovedIn(gSourceCtry) = 5
                gMoveTimes(gSourceCtry) = 2
                limitMoves = trNmbr
                Exit Function
            Else
                gMovedIn(gSourceCtry) = gMovedIn(gSourceCtry) + trNmbr
                limitMoves = trNmbr
                Exit Function
            End If
        End If
    ElseIf gMoveTimes(gSourceCtry) = 2 Then
        If gMoveTimes(gTargetCtry) < 2 Then        '2 to 0, to->2
            gMoveTimes(gTargetCtry) = 2            '2 to 1, to->2
            If gMoveLimit = 2 Then
                limitMoves = 0
                Exit Function
            End If
            trNmbr = 1
            limitMoves = trNmbr
            Exit Function
        Else                                    '2 to 2
            trNmbr = 0
            limitMoves = trNmbr
            Exit Function
        End If
    End If
End Function

'Return True if move is possible from country (moveFrom).
Private Function isMoveOK(moveFrom As Integer) As Boolean
    If gMoveLimit = 0 Then               'Unrestricted
        isMoveOK = gMoveTimes(moveFrom) <> 2 'Move all
    ElseIf gMoveLimit = 1 Then           'Restricted
        If gMoveTimes(moveFrom) = 0 Then
            isMoveOK = True             'All
        ElseIf gMoveTimes(moveFrom) = 1 Then
            isMoveOK = True             'Only 5
        Else
            isMoveOK = False            'None
        End If
    Else    ' gMoveLimit=2               'Only 1 move
        If gMoveTimes(moveFrom) = 0 Then
            isMoveOK = True             'All
        Else
            isMoveOK = False            'None
        End If
    End If
End Function

'Change scores, do housekeeping.
Private Sub mooveEm(trNmbr As Integer)
    Dim newScoreTo As Integer, newScoreFrom As Integer
    
    newScoreFrom = gCtryScore(gSourceCtry) - trNmbr
    newScoreTo = gCtryScore(gTargetCtry) + trNmbr
    Call printNewScore(gSourceCtry, newScoreFrom)
    Call printNewScore(gTargetCtry, newScoreTo)
End Sub

    'Takes action depending on current mode
Private Sub processMode(CountryNumber As Integer)
    
    Select Case gCurrentMode
    
    Case 4                  'Must change cards
        Exit Sub
    Case 3                  'Printing map using timer
        Exit Sub
    Case 2                  'Place new units
        Call mode2(CountryNumber)
        Call ToggleKeys(gPlayerValue = 0)
    Case 10                 'Moove to
        Call mode10(CountryNumber)
    Case 11                 'Moove from
        Call mode11(CountryNumber)
    Case 12                'Mooving
        Call mode12(CountryNumber)
    Case 20                'Attack who?
        Call mode20(CountryNumber)
    Case 21                'Attack from?
        Call mode21(CountryNumber)
        Call ToggleKeys(gPickedUpUnits = 0)
    Case 22             'Picking up armies
        Call mode22(CountryNumber)
        Call ToggleKeys((gPickedUpUnits = 0) And (gPlayerID(gPlayerTurn).card(4) = 0))
    Case 23                'Attacking
        Call mode23(CountryNumber)
        Call ToggleKeys((gPickedUpUnits = 0) And (gPlayerID(gPlayerTurn).card(4) = 0))
    Case 24                'Retreat (pussy)
        Call mode24(CountryNumber)
        Call ToggleKeys((gPickedUpUnits = 0) And (gPlayerID(gPlayerTurn).card(4) = 0))
    Case 1                   'Idle, do nothing
        Exit Sub
    Case 25
        Call mode25(CountryNumber)
        Call ToggleKeys((gPickedUpUnits = 0) And (gPlayerID(gPlayerTurn).card(4) = 0))
    Case 0                    'Test mode
        'Call ColorCountry(CountryNumber, longColrNow)
        cmdAttack.Enabled = True
        mnuAttack.Enabled = cmdAttack.Enabled
    End Select
End Sub

    'Place new units on map
Private Sub mode2(CountryNumber As Integer)
    Dim trnsfr As Integer
    If gCountryOwner(CountryNumber) <> gPlayerTurn Then
        Exit Sub
    ElseIf gPlayerValue - transferNmbr < 1 Then
        trnsfr = gPlayerValue
    Else
        trnsfr = transferNmbr
    End If
    gPlayerValue = gPlayerValue - trnsfr
    gPlayerStats(gPlayerTurn).UnitsIssued = gPlayerStats(gPlayerTurn).UnitsIssued + trnsfr
    Call printNewScore(CountryNumber, (gCtryScore(CountryNumber) + trnsfr))

    If gPlayerValue = 0 Then
        Call ToggleKeys(True)
        
        InfoBoxPrint 0
        InfoBoxPrnCR 9, gPlayerTurn
        If GetPlayerController(gPlayerTurn) = 0 Then
            Call AttackClicked
            Exit Sub
        End If
        gCurrentMode = 1
        Exit Sub
    End If
    Call printPlaceUnits        ' You have xx units to place
End Sub

    ' You have xx units to place
Public Sub printPlaceUnits()
    Call ColorCountryUnderAttack(0)
    InfoBoxPrint 0
    InfoBoxPrnCR 9, gPlayerTurn
    InfoBoxPrint 1, 70               'you have
    InfoBoxPrint 2, gPlayerValue
    If gPlayerValue > 1 Then
        InfoBoxPrnCR 1, 71                   'units to place
        If gLanguage <> eLanguage.German _
        And gLanguage <> eLanguage.Spanish _
        And gLanguage <> eLanguage.French Then
            InfoBoxPrint 7                   'Extra space if not German or Spanish or French
        End If
    Else
        InfoBoxPrnCR 1, 105                  'unit to place
        If gLanguage <> eLanguage.French Then
            InfoBoxPrint 7                   'Extra space if not French
        End If
    End If
    InfoBoxPrint 7
    If gPlayerID(gPlayerTurn).playerWho = 0 Then
        InfoBoxPrint 1, 72               '<Click destination country>
    End If
End Sub

    'Put cards down
Public Sub Mode5()
    Dim tmp As Long
    Dim cntr As Integer
    
    If gMapSetupLock Or gWarRestartLock Then
        Exit Sub
    End If
    gCurrentMode = 2
    Call ColorCountryUnderAttack(0)
    Call CardOutOfHand(gPlayerTurn)
    
    Call showPlayerInfo
    Call CheckCards
    
    If gCurrentMode = 2 Then
        Call ToggleKeys(gPlayerValue = 0)
        Call ToglleCardKeys(False)
        If gPlayerValue = 0 Then
            If notHitMove Then
                Call AttackClicked
            Else
                Call MoveClicked
            End If
        End If
        Exit Sub
    End If

    If chkCardsVulture.Value > 0 Then
        Call DrawAllCards
    End If
End Sub

    'Get "moove to" country
Private Sub mode10(CountryNumber As Integer)
    If Not onAborder(CountryNumber, gPlayerTurn) Then
        Exit Sub
    ElseIf gCountryOwner(CountryNumber) = gPlayerTurn Then
        gTargetCtry = CountryNumber
        Call ColorCountryUnderAttack(0)
        
        InfoBoxPrint 0                   'cls
        InfoBoxPrint 9, gPlayerTurn
        InfoBoxPrint 3, 1                'space
        InfoBoxPrint 1, 73               'moves
        InfoBoxPrnCR 1, 74               'to
        InfoBoxPrint 8, CountryNumber
        InfoBoxPrint 5                   'bold
        InfoBoxPrint 1, 107              'from
        InfoBoxPrnCR 3, 1                'space
        InfoBoxPrnCR 6                   'normal
        InfoBoxPrint 7
        InfoBoxPrnCR 1, 108              '<source>
        
        gCurrentMode = 11
    End If
End Sub

    'Get "Attack who" country
Private Sub mode20(CountryNumber As Integer)
    If gCountryOwner(CountryNumber) <> gPlayerTurn _
    And CanIAttackThisCountry(CountryNumber, gPlayerTurn) Then
        gTargetCtry = CountryNumber
        Call ColorCountryUnderAttack(0)
        Call ColorCountryUnderAttack(gTargetCtry)
        
        InfoBoxPrint 0
        InfoBoxPrint 9, gPlayerTurn
        InfoBoxPrint 3, 1
        InfoBoxPrnCR 1, 64               'attacks
        InfoBoxPrnCR 8, CountryNumber
        InfoBoxPrint 5
        InfoBoxPrnCR 1, 101
        InfoBoxPrnCR 6
        InfoBoxPrint 7
        If gPlayerID(gPlayerTurn).playerWho = 0 Then
            InfoBoxPrnCR 1, 109         '<Click attacking country(s)>
        End If
        
        gPickedUpUnits = 0
        gCurrentMode = 21
        Call addChangeToList(gTargetCtry + 130, 0, 1)
    End If
End Sub

    'Get "moove from" country
Private Sub mode11(CountryNumber As Integer)
        If gTargetCtry = CountryNumber Then   'Click same country
            Call MoveClicked
            Exit Sub
        ElseIf gCountryOwner(CountryNumber) <> gPlayerTurn Then
            Exit Sub
        ElseIf Not isNeighbour(gTargetCtry, CountryNumber) Then
            Exit Sub                        'not a neighbour
        ElseIf gCtryScore(CountryNumber) = 1 Then
            Exit Sub                        '1 unit on source country
        End If
        
        gSourceCtry = CountryNumber
        gCurrentMode = 12
        Call mode12(CountryNumber)
        Call addChangeToList(CountryNumber, 0, 1)
        Exit Sub
End Sub

    'Get "attack from" country
Private Sub mode21(CountryNumber As Integer)
    If gCountryOwner(CountryNumber) <> gPlayerTurn Then
        Exit Sub
    ElseIf Not isNeighbour(CountryNumber, gTargetCtry) Then
        Exit Sub
    ElseIf gCtryScore(CountryNumber) <= 1 Then
        Call AttackClicked
        Exit Sub
    End If
    gSourceCtry = CountryNumber
    gCurrentMode = 22
    Call mode22(CountryNumber)
    Call addChangeToList(CountryNumber, 0, 0)
End Sub

    'Pickup armies to attack with
Private Sub mode22(CountryNumber As Integer)
    Dim trnsfr As Integer
    If CountryNumber = gTargetCtry Then
        gCurrentMode = 23
        Call mode23(CountryNumber)
        Exit Sub
    ElseIf gCountryOwner(CountryNumber) <> gPlayerTurn Then
        Exit Sub
    ElseIf Not isNeighbour(CountryNumber, gTargetCtry) Then
        Exit Sub
    ElseIf gCtryScore(CountryNumber) <= 1 Then
        Exit Sub
    ElseIf gCtryScore(CountryNumber) - transferNmbr < 1 Then
        trnsfr = gCtryScore(CountryNumber) - 1
    Else
        trnsfr = transferNmbr
    End If
    gPickedUpUnits = gPickedUpUnits + trnsfr
    If gPickedUpUnits = 0 Then
        Call ToggleKeys(True)       'To be sure!!!
    End If
    Call printNewScore(CountryNumber, (gCtryScore(CountryNumber) - trnsfr))
    
    InfoBoxPrint 0
    InfoBoxPrint 9, gPlayerTurn
    InfoBoxPrint 3, 1
    InfoBoxPrnCR 1, 64               'attacks
    InfoBoxPrnCR 8, gTargetCtry
    InfoBoxPrint 5                   'bold
    If gPickedUpUnits > 1 Then
        InfoBoxPrint 1, 101          'with
        InfoBoxPrint 2, gPickedUpUnits
        InfoBoxPrnCR 1, 102          'units
    Else
        InfoBoxPrint 1, 101          'with
        InfoBoxPrint 2, gPickedUpUnits
        InfoBoxPrnCR 1, 103          'unit
    End If
    InfoBoxPrnCR 6                   'normal
    If gPlayerID(gPlayerTurn).playerWho = 0 Then
        If gPickedUpUnits > 0 Then
            InfoBoxPrnCR 1, 65           '<Click defending>
        Else
            InfoBoxPrint 7
        End If
        If gCtryScore(gSourceCtry) > 1 And gPlayerID(gPlayerTurn).playerWho = 0 Then
            InfoBoxPrnCR 1, 109          '<attacking ctry>
        End If
    End If
    
    Call addChangeToList(gTargetCtry + 130, 0, 0)
    Call addChangeToList(gSourceCtry, 0, 0)
End Sub

    'Attacking
Private Sub mode23(CountryNumber As Integer)
    Dim dummy As Integer
    
    If CountryNumber = gTargetCtry Then
        Call AttackCountry
    ElseIf gCountryOwner(CountryNumber) = gPlayerTurn Then
        If isNeighbour(CountryNumber, gTargetCtry) Then
            'Call colorCountry(gTargetCtry, playerID(gCountryOwner(gTargetCtry)).lngColor)
            Call ColorCountryUnderAttack(0)
            
            InfoBoxPrint 0
            InfoBoxPrnCR 9, gPlayerTurn
            InfoBoxPrnCR 1, 110              'retreats 0 units from
            InfoBoxPrnCR 8, gTargetCtry
            InfoBoxPrint 7
            If gPlayerID(gPlayerTurn).playerWho = 0 Then
                InfoBoxPrnCR 1, 104              '<attking ctry>
                InfoBoxPrnCR 1, 65               '<dfnding ctry>
            End If
            
            retreatArmies = 0
            Call ToggleKeys(False)
            gCurrentMode = 24            'Retreat
            retreatArmies = 0
            Call mode24(CountryNumber)
            Call addChangeToList(gTargetCtry, 0, 1)
            Call addChangeToList(gSourceCtry, 0, 1)
            Exit Sub
        End If
    End If
    If gPlayerID(gPlayerTurn).card(4) = 0 Then
        ToggleKeys (gPickedUpUnits = 0)
    Else
        ToggleKeys (False)
    End If
End Sub

    'Retreating
Private Sub mode24(CountryNumber As Integer)
    Dim trnsfr As Integer
    
    If CountryNumber = gTargetCtry Then   'Attack again
        gCurrentMode = 23
        Call ToggleKeys(True)
        
        InfoBoxPrint 0
        InfoBoxPrint 9, gPlayerTurn
        InfoBoxPrint 3, 1
        InfoBoxPrnCR 1, 64           'attacks
        InfoBoxPrnCR 8, gTargetCtry
        
        Call AttackCountry
        Exit Sub
    ElseIf gCountryOwner(CountryNumber) <> gPlayerTurn Then
        Exit Sub
    ElseIf Not isNeighbour(CountryNumber, gTargetCtry) Then
        Exit Sub
    ElseIf gPickedUpUnits - transferNmbr <= 0 Then
        trnsfr = gPickedUpUnits
        Call ToggleKeys(True)
        Call AttackClicked
    Else
        trnsfr = transferNmbr
    End If
    gPickedUpUnits = gPickedUpUnits - trnsfr
    retreatArmies = retreatArmies + trnsfr
    Call printNewScore(CountryNumber, (gCtryScore(CountryNumber) + trnsfr))
    Call addChangeToList(CountryNumber, 0, 1)
    If gCurrentMode = 23 Then
        Exit Sub
    End If
    
    InfoBoxPrint 0
    InfoBoxPrnCR 9, gPlayerTurn
    InfoBoxPrint 1, 111          'retreates
    InfoBoxPrint 3, 1
    InfoBoxPrint 2, retreatArmies
    If retreatArmies > 1 Then
        InfoBoxPrnCR 1, 112                          'units
    Else
        InfoBoxPrnCR 1, 113                          'unit
    End If
    InfoBoxPrnCR 8, gTargetCtry
    If gPickedUpUnits > 0 Then
        InfoBoxPrint 5                               'bold
        InfoBoxPrint 1, 101                          'with
        InfoBoxPrint 2, gPickedUpUnits
        If gPickedUpUnits > 1 Then
            InfoBoxPrnCR 1, 114                      'units left
        Else
            InfoBoxPrnCR 1, 115                      'unit left
        End If
        InfoBoxPrint 6                               'normal
        If gPlayerID(gPlayerTurn).playerWho = 0 Then
            InfoBoxPrnCR 1, 104                          '<attking ctry>
            InfoBoxPrnCR 1, 65                           '<dfnding ctry>
        End If
    End If
End Sub

    'Retreat
Private Sub mode25(CountryNumber As Integer)
    gCurrentMode = 24
    retreatArmies = 0
End Sub

        'Click destination country to complete transfer
Private Sub mode12(CountryNumber As Integer)
    If CountryNumber = gTargetCtry Then
        Call MoveClicked
        Exit Sub
        'Otherwise keep going
    ElseIf isNeighbour(gTargetCtry, CountryNumber) _
        And gCountryOwner(CountryNumber) = gPlayerTurn Then
        gSourceCtry = CountryNumber
        Call mode12Change
    End If
End Sub

    'Move scores between countries
Private Sub mode12Change()
    Dim trNmbr As Integer
    
    trNmbr = transferNmbr
    If gMoveLimit <> 0 Then
        trNmbr = limitMoves(trNmbr)
    End If
    If gCtryScore(gSourceCtry) - trNmbr <= 1 Then
        trNmbr = gCtryScore(gSourceCtry) - 1
    End If
    Call mooveEm(trNmbr)
    
    InfoBoxPrint 0
    InfoBoxPrint 9, gPlayerTurn
    InfoBoxPrint 3, 1
    InfoBoxPrint 1, 73                   'moves
    InfoBoxPrnCR 1, 74                   'to
    InfoBoxPrint 8, gTargetCtry
    InfoBoxPrint 3, 1
    InfoBoxPrnCR 1, 107                  'from
    InfoBoxPrint 8, gSourceCtry
    InfoBoxPrnCR 4                       'Period
    InfoBoxPrint 7
    If gPlayerID(gPlayerTurn).playerWho = 0 Then
        If gCtryScore(gSourceCtry) > 1 Then
            InfoBoxPrnCR 1, 108              '<click source ctry>
        Else
            InfoBoxPrint 7
        End If
        InfoBoxPrnCR 1, 72                   '<click dest ctry>
    End If
End Sub

    'returns TRUE if 2 countries are neighbours
Private Function isNeighbour(ctryFrom As Integer, ctryTo As Integer) As Boolean
    Dim counter As Integer
    
    For counter = 1 To 7
        If CountryID(ctryFrom).neighbour(counter) = 0 Then
            isNeighbour = False
            Exit Function
        ElseIf CountryID(ctryFrom).neighbour(counter) = ctryTo Then
            isNeighbour = True
            Exit Function
        End If
    Next counter
    isNeighbour = False
End Function

'Returns the country number of mouseX and mouseY hit position, 0 if none.
Private Function GetMouseHitPosition(mouseX As Integer, mouseY As Integer) As Integer
    Dim vCountryNumber As Integer
    Dim vSrcHiddenX As Integer
    Dim vSrcHiddenY As Integer
    
    'Work out the conversion factors.
    vSrcHiddenX = CInt(mouseX / gPictureMaskRatioX)
    vSrcHiddenY = CInt(mouseY / gPictureMaskRatioY)
    
    'Debug.Print vSrcHiddenX, vSrcHiddenY
    
    'Test each individual country to determine if it is under the passed mouse position.
    For vCountryNumber = 1 To 42
        If IsCountryUnderPosition(vSrcHiddenX, vSrcHiddenY, vCountryNumber) Then
            GetMouseHitPosition = vCountryNumber
            Exit Function
        End If
    Next vCountryNumber
    
    'Check for click in the ocean to the left of Alaska and return Kamchatta.
    If vSrcHiddenX < 36 And vSrcHiddenY < 240 And vSrcHiddenY > 80 Then
        GetMouseHitPosition = 38
    
    'Check for click in the ocean to the right of Kamchatta and return Alaska.
    ElseIf vSrcHiddenX > 900 And vSrcHiddenY < 240 And vSrcHiddenY > 80 Then
        GetMouseHitPosition = 1
    
    'Click in the water. Would be cool to animate a splash.
    Else
        GetMouseHitPosition = 0
    End If
End Function

'Return true if the passed country is under the passed mouse position.
Private Function IsCountryUnderPosition(mouseX As Integer, mouseY As Integer, CountryNumber As Integer) As Boolean
    Dim vMaskPositionX As Integer
    Dim vMaskPositionY As Integer
    
    IsCountryUnderPosition = False
    
    'Translate the position of the point under the mouse to the
    'black and white stencil blit mask Mask4.pctMaskArray(0).
    vMaskPositionX = mouseX - CountryID(CountryNumber).destX + CountryID(CountryNumber).srcX
    vMaskPositionY = mouseY - CountryID(CountryNumber).destY + CountryID(CountryNumber).srcY

    'Test if the passed position is within the confines (ie blitting rectangle) of the country.
    If mouseX > CountryID(CountryNumber).destX _
    And mouseX < CountryID(CountryNumber).destX + CountryID(CountryNumber).Width _
    And mouseY > CountryID(CountryNumber).destY _
    And mouseY < CountryID(CountryNumber).destY + CountryID(CountryNumber).Height Then
        
        'Test certain countries for nearness of hit. Some countries are hard to hit because
        'they are small like Japan or have water between boundries like Indonesia.
        If IsCtryAreaUnderPosition(mouseX, mouseY, CountryNumber) Then
            IsCountryUnderPosition = True
        
        'If the pixel on the stencil is white, then we have a hit. Position worked out above.
        Else
            IsCountryUnderPosition = (GetPixel(Mask4.pctMaskArray(0).hdc, vMaskPositionX, vMaskPositionY) > RGB(80, 80, 80))
        End If
    End If
End Function

'Test certain countries for nearness of mouse hit. To get to this stage, the
'position under test must be within the confines of the country. By that I mean
'that it is within its blitting rectangle.
Private Function IsCtryAreaUnderPosition(mouseX As Integer, mouseY As Integer, CountryNumber As Integer) As Boolean
    Select Case CountryNumber
    Case 14
        'Iceland. Anywhere within its boundry is considered a hit.
        IsCtryAreaUnderPosition = True
    Case 26
        'Madagascar
        IsCtryAreaUnderPosition = True
    Case 41
        'New Ginea
        IsCtryAreaUnderPosition = True
    Case 42
        'Eastern Australia and NZ.
        IsCtryAreaUnderPosition = True
    Case 3
        'Greenland has ocean between its islands.
        IsCtryAreaUnderPosition = (mouseX > 217) And (mouseX < 310) _
                                And (mouseY > 96) And (mouseY < 182)
    Case 9
        'Mexico is really thin.
        IsCtryAreaUnderPosition = ((mouseX > 155) And (mouseX < 236) _
                                And (mouseY > 332) And (mouseY < 393)) Or ((mouseX > 135) And (mouseX < 172) And (mouseY > 308) And (mouseY < 333))
    Case 15
        'Great Britain.
        IsCtryAreaUnderPosition = (mouseX > 362) And (mouseX < 409) _
                                And (mouseY > 229) And (mouseY < 277)
    Case 18
        'Western Europe.
        IsCtryAreaUnderPosition = (mouseX > 380) And (mouseX < 410) And (mouseY > 287) And (mouseY < 325)
    Case 19
        'Southern Europe
        IsCtryAreaUnderPosition = (mouseX > 431) And (mouseX < 487) _
                                And (mouseY > 295) And (mouseY < 328)
    Case 27
        'Middle East
        IsCtryAreaUnderPosition = (mouseX > 493) And (mouseX < 553) And (mouseY > 315) And (mouseY < 359)
        
    Case 30
        'Siam
        IsCtryAreaUnderPosition = (mouseX > 654) And (mouseX < 781) _
                                And (mouseY > 375) And (mouseY < 472)
    Case 37
        'Japan
        IsCtryAreaUnderPosition = (mouseX > 782) And (mouseX < 839) _
                                And (mouseY > 275) And (mouseY < 327)
    Case 39
        'Indonesia
        IsCtryAreaUnderPosition = (mouseX > 693) And (mouseX < 775) _
                                And (mouseY > 427) And (mouseY < 477)
    Case Else
        'Nup.
        IsCtryAreaUnderPosition = False
    End Select
End Function
                    
    'Print score on fresh map
Private Sub printValue(intCtryNmbr As Integer)
    Mask4.Map1.ForeColor = 0
    If intCtryNmbr = 15 Then
        Mask4.Map1.ForeColor = &HFFFFFF
    ElseIf intCtryNmbr = 37 Then
        Mask4.Map1.ForeColor = &HFFFFFF
    ElseIf intCtryNmbr = 26 Then
        Mask4.Map1.ForeColor = &HFFFFFF
    End If
    Mask4.Map1.CurrentX = CountryID(intCtryNmbr).printX
    Mask4.Map1.CurrentY = CountryID(intCtryNmbr).printY
    Mask4.Map1.Print gCtryScore(intCtryNmbr)
End Sub

Private Sub Picture1_DblClick()
    
    'If clickLock Then Exit Sub
    'clickLock = True
    'Call updateTestViewer("Start map dbl click")
    Call map1DblClicked
    'Call updateTestViewer("End map dbl click")
    'clickLock = False
End Sub

Private Sub map1DblClicked()
    Dim tst1 As Boolean
    Dim rtrn2 As Integer
    
    If (netWorkSituation <> cNetNone) _
    And (net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber) Then Exit Sub
    
    tst1 = (gCurrentMode = 12) Or (gCurrentMode = 22) _
            Or (gCurrentMode = 24) Or (gCurrentMode = 2)
    If tst1 Then
        Call Picture1_Click
        Exit Sub
    ElseIf gCurrentMode = 23 And gPlayerID(gPlayerTurn).playerWho = 0 Then
        gCurrentMode = 25
        Timer1.Enabled = True
        Timer1.Interval = diceSpeed
    End If
    'If gCurrentMode = 25 Then
    '    Call attackCountry
    '    Exit Sub
    'End If
End Sub

    'Show info when right mouse button has been pressed
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Call updateTestViewer("Start mouse down")
    Call map1MouseDown(Button, Shift, x, y)
    'Call updateTestViewer("End mouse down")
End Sub

Private Sub map1MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cntry_hit As Integer
    Dim cont_hit As Integer
    Dim dvCode As Long
    
    On Error Resume Next
    
    If Button <> 2 Then
        Exit Sub
    End If
    
    cntry_hit = GetMouseHitPosition(gCurrentMousePosX, gCurrentMousePosY)
    If cntry_hit = 0 Then
        Exit Sub
    End If
    If gCheatMode.createMap Then       'Global Siege editor
        frmEdit.setCountry cntry_hit, CountryID(cntry_hit).ctryName, gCtryScore(cntry_hit), gCountryOwner(cntry_hit)
        frmEdit.Show vbModal
        Exit Sub
    End If
    cont_hit = ContinentOfCtry(cntry_hit) - 1
    frmIntelligence.country = CountryID(cntry_hit).ctryName
    frmIntelligence.Occupier = gPlayerID(gCountryOwner(cntry_hit)).strColor
    frmIntelligence.Continent = Phrase(116) + Trim(Continents(cont_hit).ContNameText) + ", " _
        + Phrase(117) + str(udContVal(cont_hit).Value) + Phrase(102) + " " _
        + Phrase(118) + Trim(gPlayerID(gCountryOwner(cntry_hit)).strColor) + "."
    frmIntelligence.Show vbModal
End Sub

    ' Allow external modules to edit country owners and scores
Public Sub editMap(editCntry As Integer, newScore As Integer, newOwner As Integer)
    gCtryScore(editCntry) = newScore
    gCountryOwner(editCntry) = newOwner
    Call refreshMap
    gPlayerTurn = gPlayerTurn - 1
    Call SaveCheckpoint
    gPlayerTurn = gPlayerTurn + 1
End Sub

Private Sub maxCard_GotFocus()
    On Error Resume Next
    SetupScreen.SetFocus
End Sub

'Fast/slow dice selected. Keep controls in sinc, prevent call backs
Private Sub changeDiceSpeed(forceFast As Boolean)
    Static Locked As Boolean
    
    If Locked Then Exit Sub
    Locked = True
    'mnuDiceFast.Checked = forceFast
    'chkFastDice.Value = Abs(CInt(forceFast))
    'ToolCheck(1).Value = Abs(CInt(forceFast))
    If forceFast Then
        diceSpeed = diceFast
    Else
        diceSpeed = diceSlow
    End If
    Locked = False
    net.setupControlChange = True
End Sub

Private Sub mnuFileExit_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub mnuFileLoadWar_Click()
    Dim vWarFileName As String
    
    On Error GoTo errHandle
    If gMapSetupLock Or gWarRestartLock Then
        Exit Sub
    End If
    
    'If half way through the draw win sequence or waiting to restart the war.
    tmrDrawWin.Tag = ""
    
    If frmStats.Visible Then
        frmStats.Hide
    End If
    
    'frmAdvanced.Hide
    frmOpen.currentWar = gCurrentWarPath
    
    frmOpen.Show vbModal
    vWarFileName = frmOpen.SelectedWar
    
    If vWarFileName = "" Then
        Exit Sub
    End If
    
    'Invalidate stats.
    Call ValidateAllStats(False)
    
    If OpenWar(vWarFileName) Then
        gCurrentWarPath = vWarFileName
        Call RevertToCheckpoint
        If gPlayerTurn = 0 Then gPlayerTurn = 1
        gPlayerTurn = gPlayerTurn - 1
        Call SaveCheckpoint
        gPlayerTurn = gPlayerTurn + 1
    End If
    If SetupScreen.Visible Then
        TheMainForm.BackColor = &H8000000F
        Timer1.Enabled = False
        Timer2.Enabled = False
    End If
    
    If (netWorkSituation = cNetHost) Then
        net.setupControlChange = True
        netMain.sendSetupScreen
        net.setupControlChange = False
    End If
    mnuOptUndo.Enabled = gCheatMode.undoEnabled
    Toolbar1.Buttons(8).Enabled = mnuOptUndo.Enabled
    Call SyncForgroundMap("mnuFileLoadWar_Click")
    Exit Sub
errHandle:
    If Not gHeadlessMode Then
        MsgBox LimitTextWidth(Phrase(218) & Trim(vWarFileName) & Phrase(219), 30), vbCritical, Phrase(220)  'A file error has occured. Perhapse
    End If
    LogError "mnuFileLoadWar_Click", "Error: " & Err.Description
    boolDrawnWin = False
    Call putBack
    Call RevertToCheckpoint
    TheMainForm.BackColor = &H8000000F
    Exit Sub
End Sub

    'Open back up war to recover from error (warBak -> warSit)
Private Sub putBack()
    Call copyWar(warBak, warSit)
End Sub

    'Save back up war (warSit -> warBak)
Private Sub backUp()
    Call copyWar(warSit, warBak)
End Sub

Private Sub copyWar(warFrom As WarControlType, warTo As WarControlType)
    warTo = warFrom
End Sub

Private Sub mnuFileNew_Click()
    On Error Resume Next
    
    If gMapSetupLock Then
        Exit Sub
    End If
    
    'Hide the stats form if visible.
    If frmStats.Visible Then
        frmStats.Hide
    End If
    
    'If half way through the draw win sequence or waiting to restart the war.
    tmrDrawWin.Tag = ""
    
    tmpTimer1 = Timer1.Enabled
    tmpTimer2 = Timer2.Enabled
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    'If networked, stop lients from being able to go to the setup screen
    'and back to the game and beint awarded more units.
    If GetPlayerController(gPlayerTurn) = 0 _
    And netWorkSituation <> cNetNone _
    And net.playerOwner(gPlayerTurn - 1) = myTerminalNumber Then
        Call ForfeitTurn
    End If
    
    Call setupNewGame
    If netWorkSituation = cNetHost Then
        netMain.sendSetupScreen
        net.setupControlChange = False
    ElseIf netWorkSituation = cNetClient Then
        cmdSetupOk.Enabled = False
        gCurrentMode = gPreviousSettings.PrevMode
        cmdSUPcncl.Enabled = Not (gCurrentMode = 13 Or gCurrentMode = 18)
        mnuDeclareWar.Enabled = cmdSetupOk.Enabled
        mnuCancelSetup.Enabled = cmdSUPcncl.Enabled
    End If
End Sub

Public Sub mnuFileReset_Click()
    On Error Resume Next
    
    If gMapSetupLock Then
        Exit Sub
    End If
    
    'Make sure the transfer rate is set to 1. This stops human players
    'being lwft with a high transfer rate but onlt 1 is transfered at
    'a time if they have the first turn after a reset war.
    tfRate1.Value = True
    transferNmbr = 1
    
    'If half way through the draw win sequence or waiting to
    'restart the war, do it now and exit.
    If tmrDrawWin.Tag <> "" Then
        Call tmrDrawWin_Timer
        Exit Sub
    End If
    
    'Hide the stats form if visible.
    If frmStats.Visible Then
        frmStats.Hide
    End If
    
    If (gCurrentMode <> 13) _
    And (gCurrentMode <> 18) _
    And (gCurrentMode <> 3) _
    And (Not SetupScreen.Visible) _
    And Not gServerMode _
    And Not gHeadlessMode Then
        'Do you really want to quit this war?
        If MsgBox(Phrase(191), vbYesNo, gcApplicationName) <> vbYes Then
            Exit Sub
        End If
    End If
    mnuFileNew_Click
    cmdSetupOk_Click
End Sub

'Save the war as the current name if it is not locked and there is a valid vile path.
'Open the Save As... dialog box if it is.
Private Sub mnuFileSaveWar_Click()
    On Error Resume Next
    
    If warSit.Locked _
    Or warSit.sCtryOwner(1) = 0 _
    Or Trim(gCurrentWarPath) = "" _
    Or Dir(gCurrentWarPath) = "" Then
        mnuFileWarAs_Click
    Else
        Call SaveWar(gCurrentWarPath, warSit.filename, warSit.fileDescription, False)
    End If
End Sub

Private Sub mnuFileWarAs_Click()
    Dim vWarTitle As String
    Dim vWarDescription As String
    Dim vWarLocked As Boolean
    'Dim vWarFilePath As String
    
    On Error Resume Next
    frmSaveAs.Show vbModal
    
    vWarTitle = frmSaveAs.WarTitle
    vWarDescription = frmSaveAs.WarDescription
    vWarLocked = frmSaveAs.WarLocked
    gCurrentWarPath = frmSaveAs.WarFilePath
    
    If vWarTitle = "" Then
        Exit Sub
    End If
    
    If SetupScreen.Visible Then
        Call ChangeTitlebarText(Phrase(34) + Trim(vWarTitle))   'Global Siege Set Up...
    Else
        Call ChangeTitlebarText(Phrase(35) + Trim(vWarTitle))   'Global Siege -
    End If
    Call ShowMenuBar(SetupScreen.Visible)
    
    Call SaveWar(gCurrentWarPath, vWarTitle, vWarDescription, vWarLocked)
    
    If (netMain.CountTerminals <> 0) Then
        netMain.sendSetupScreen
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    On Error Resume Next
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelpContents_Click()
    On Error Resume Next
    Call OpenWebPage(gHelpWebPage)
End Sub

'Show the player's mission if it is their thurn.
'The global variable gAskedToSeeMission is reset at the end of the turn.
Public Sub mnuMissionSee_Click()
    Dim tmp As Long
    
    On Error Resume Next
    If SetupScreen.Visible Then
        Exit Sub
    End If
    
    'Human on this terminal does not control the current player.
    If (chkMsnMissionsOn.Value = vbChecked And GetPlayerController(gPlayerTurn) <> 0) _
    And Not gCheatMode.seeMissions Then
        
        'Hide remote player's mission.
        If netWorkSituation <> cNetNone _
        And net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber Then
            gAskedToSeeMission = True
        End If
        
        'Display random message for computer player.
        If Not gAskedToSeeMission Then
            tmpTimer1 = Timer1.Enabled
            tmpTimer2 = Timer2.Enabled
            Timer1.Enabled = False
            Timer2.Enabled = False
            
            'First time, "AI interupt request - The XX Army wants to keep the mission a secret!"
            tmp = MsgBox(Trim(gPlayerID(gPlayerTurn).strColor) + _
                Phrase(190), vbMsgBoxRight, Phrase(196))
            gAskedToSeeMission = True
            Timer1.Enabled = tmpTimer1
            Timer2.Enabled = tmpTimer2
            
        '50% chance
        '"SPRUNG!       Ah Hah! Caught spying!
        'You must wait until it's your turn to see what your mission is!"
        ElseIf Int(GenRandom4 * 2) = 1 Then
            MRbox Phrase(119), Phrase(120)
        
        '25% chance
        '"Stop peeking! NO! My mission!
        'You must wait until it's your turn to see what your mission is!"
        ElseIf Int(GenRandom4 * 2) = 1 Then
            MRbox Phrase(121), Phrase(122)
        
        '12.5% chance
        '"Conspiracy    This is classified information. If I tell you, I will have to kill you.
        'Or you could wait for your turn to see what your own mission is."
        ElseIf Int(GenRandom4 * 2) = 1 Then
            MRbox Phrase(123), Phrase(124)
        
        '6.25% chance.
        '"NO!           My mission! I finish it! You just wait until your turn and finish your own mission!"
        ElseIf Int(GenRandom4 * 2) = 1 Then
            MRbox Phrase(125), Phrase(126)
        
        '6.25% chance
        '"Spying is NOT tolerated in war!
        'Your spies were all captured and executed. They died slowly in one of our many torture chambers.
        'Let that be a warning to you human scum! Your human armies will soon be exterminated, and our
        'Global Siege world will be a better place for computer players to live!     You intolerable species"
        Else
            tmpTimer1 = Timer1.Enabled
            tmpTimer2 = Timer2.Enabled
            Timer1.Enabled = False
            Timer2.Enabled = False
            tmp = MsgBox(LimitTextWidth(Phrase(127), 45), vbApplicationModal, Phrase(128))
            Timer1.Enabled = tmpTimer1
            Timer2.Enabled = tmpTimer2
        End If
    Else
        'Show the human player sitting at this terminal their mission.
        MRbox gPlayerID(gPlayerTurn).strColor, gMissions(gPlayerID(gPlayerTurn).mission).DescriptionText, True
        gSeenMission(gPlayerTurn) = True
    End If
End Sub

    'A better looking message box
Public Function MRbox(Title As String, Content As String, Optional showModal As Boolean = False) As Long
    
    On Error GoTo openError
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
        frmIntelligence.country = Title
        frmIntelligence.Continent = Content
        If showModal Then
            frmIntelligence.Show , TheMainForm
        Else
            frmIntelligence.Show vbModal, TheMainForm
        End If
    End If
    MRbox = 0
    Exit Function
    
openError:
    Resume Next
    
End Function

Private Sub mnuOptLanguage_Click()
    On Error Resume Next

    frmLanguage.Show vbModal
    Call SetNewWords
    Call refreshMap
    Call SyncForgroundMap("mnuOptLanguage_Click")
End Sub

Private Sub mnuOptReport_Click()
    If mnuOptReport.Checked Then
        mnuOptReport.Checked = False
    Else
        mnuOptReport.Checked = True
    End If
End Sub

Private Sub mnuOptToolbox_Click()
    Dim vState As Boolean
    On Error Resume Next
    
    vState = Not mnuOptToolbox.Checked
    mnuOptToolbox.Checked = vState
    Call ShowToolBar(vState)
End Sub

'Show or hide the toolbar. If pRefresh is TRUE then resize the form.
Private Sub ShowToolBar(pShow As Boolean, Optional pResize As Boolean = True)
    Toolbar1.Visible = pShow
    If pResize Then
        Call Form_Resize
    End If
End Sub

'Return the height of the toolbar or 0 if it is hidden.
Private Function GetToolbarHeight() As Long
    If mnuOptToolbox.Checked Then
        GetToolbarHeight = Toolbar1.Height
    Else
        GetToolbarHeight = 0
    End If
End Function

Private Sub mnuOptUndo_Click()
    Dim tmp As Long
    On Error Resume Next
    
    If clickLock Or gCurrentMode = 25 Or gMapSetupLock Or gWarRestartLock Then
        Exit Sub
    End If
    
    If (warSit.sNmbrOfPlayers = 0) Or (warSit.sCtryOwner(1) = 0) Then
        Exit Sub
    End If
    
    'If half way through the draw win sequence or waiting to restart the war.
    tmrDrawWin.Tag = ""
    
    'Ask player if they really want to do this.
    If Not gCheatMode.cheatActive _
    And gPlayerStats(gPlayerTurn).IsValid _
    And netWorkSituation = cNetNone Then
        If MsgBox("This action is considered cheating and will invalidate your stats." & vbCrLf _
        & "Do you really want to do this?", vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    
    'Cheat.
    'gCheatMode.cheatActive = True
    gPlayerStats(gPlayerTurn).IsValid = False
    gPlayerStats(gPlayerTurn).InvalidatedReason = "Cheated by using the" & vbCrLf & "undo option."
    
    boolDrawnWin = False
    Call RevertToCheckpoint
    SetupScreen.Visible = warSit.SetupScreen
    Call ShowMenuBar(SetupScreen.Visible)
    
    'mnuMsnMissionOn.Enabled = (SetupScreen.Visible And netWorkSituation <> cNetClient)
    Call EnableMissionOptions
    
    If flashingBorder _
    And (GetPlayerController(gPlayerTurn) = 0) _
    And (Not SetupScreen.Visible) Then
        If netWorkSituation <> cNetNone Then       'Flash color if human on this terminal
            If (net.playerOwner(gPlayerTurn - 1) = myTerminalNumber) Then
                TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
                tmrFlashInfoBox.Enabled = True
            Else
                TheMainForm.BackColor = &H8000000F
            End If
        Else
            TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
            tmrFlashInfoBox.Enabled = True
        End If
    Else
        TheMainForm.BackColor = &H8000000F
    End If
    If (netWorkSituation <> cNetNone) _
    And (myTerminalNumber = net.playerOwner(gPlayerTurn - 1)) Then
        netMain.SendRefresh
    End If
    
    'No info box flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    Call SyncForgroundMap("mnuOptUndo_Click")
End Sub

Private Sub optCardMode_Click(Index As Integer)
    Dim vIsEnabled As Boolean
    
    vIsEnabled = frmSetupCards.Enabled
    
    udMaximumCardValue.Enabled = (Index = 2 And vIsEnabled)
    lblSetupMaxCardValue.Enabled = (Index = 2 And vIsEnabled)
    txtMaximumCardValue.Enabled = (Index = 2 And vIsEnabled)
    chkCardsHidden.Enabled = (Index <> 0 And vIsEnabled)
    chkCardsVulture.Enabled = (Index <> 0 And vIsEnabled)
    Call EnableFixedValuesCardControls(vIsEnabled)
    Call EnableDeckCountsCardControls(vIsEnabled)
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

'Enable / disable the Fixed Card Values setup options.
Private Sub EnableFixedValuesCardControls(pEnabled As Boolean)
    Dim vIndex As Long
    
    frmFixedValues.Enabled = pEnabled
    lbCardValues.Enabled = pEnabled
    
    For vIndex = 0 To 3
        txtFixedValues(vIndex).Enabled = pEnabled
        udFixedValues(vIndex).Enabled = pEnabled
        lblFixedValues(vIndex).Enabled = pEnabled
    Next
End Sub

'Enable / disable the Card Deck setup options.
Private Sub EnableDeckCountsCardControls(pEnabled As Boolean)
    Dim vIndex As Long
    
    frmTheDeck.Enabled = pEnabled
    lblTheDeck.Enabled = pEnabled
    For vIndex = 0 To 3
        txtCardDeck(vIndex).Enabled = pEnabled
        udCardDeck(vIndex).Enabled = pEnabled
        lblCardDeck(vIndex).Enabled = pEnabled
    Next
End Sub

Private Sub optLimitSupply_Click()
    gMoveLimit = 1
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

    'Fast war clicked
Private Sub optnFastWar_Click()
    Call changeWarSpeed(Not optnFastWar.Checked)
End Sub

    'Fast/slow war selected. Keep controls in sinc, prevent call backs
Private Sub changeWarSpeed(forceFast As Boolean)
    Static Locked As Boolean
    
    If Locked Then
        Exit Sub
    End If
    Locked = True
    optnFastWar.Checked = forceFast
    'chkFast.Value = -CInt(forceFast)
    If forceFast Then
        playSpeed = playFast
        Toolbar1.Buttons(7).Value = tbrPressed
    Else
        playSpeed = playSlow
        Toolbar1.Buttons(7).Value = tbrUnpressed
    End If
    Call changeDiceSpeed(forceFast)
    Locked = False
    net.setupControlChange = True
    
    Exit Sub
    If Timer2.Enabled Then
        Timer2.Enabled = False
        Timer2.Interval = playSpeed * 5
        Timer2.Enabled = True
    End If
    If Timer1.Enabled Then
        Timer1.Enabled = False
        Timer1.Interval = diceSpeed
        Timer1.Enabled = True
    End If
End Sub

Private Sub optNoSupply_Click()
    gMoveLimit = 2
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub optSupplyLines_Click()
    gMoveLimit = 0
    net.setupControlChange = True
    Call CheckSetupForChange
End Sub

Private Sub tabSetup_Click()
    Call PluralReinforcementTabText
    frameSetup(tabSetup.SelectedItem.Index - 1).ZOrder 0
End Sub

Private Sub txtPlayerStartCountries_GotFocus(Index As Integer)
    On Error Resume Next
    SetupScreen.SetFocus
End Sub

Private Sub tfRate1_Click()
    Dim tstPress As Integer
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    tstPress = (GetPlayerController(gPlayerTurn) <> 0) _
    And (Not gComputerPressed)
    If tstPress Then
        Exit Sub
    End If
    transferNmbr = 1
    If gCurrentMode = 5 Then
        Call Mode5
    End If
End Sub

Private Sub tfRate10_Click()
    Dim tstPress As Integer
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    tstPress = (GetPlayerController(gPlayerTurn) <> 0) _
    And (Not gComputerPressed)
    If tstPress Then
        Exit Sub
    End If
    transferNmbr = 10
    If gCurrentMode = 5 Then
        Call Mode5
    End If

End Sub

Private Sub tfRate2_Click()
    Dim tstPress As Integer
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    tstPress = (GetPlayerController(gPlayerTurn) <> 0) _
    And (Not gComputerPressed)
    If tstPress Then
        Exit Sub
    End If
    transferNmbr = 2
    If gCurrentMode = 5 Then
        Call Mode5
    End If

End Sub

Private Sub tfRate20_Click()
    Dim tstPress As Integer
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    tstPress = (GetPlayerController(gPlayerTurn) <> 0) _
    And (Not gComputerPressed)
    If tstPress Then
        Exit Sub
    End If
    transferNmbr = 20
    If gCurrentMode = 5 Then
        Call Mode5
    End If

End Sub

Private Sub tfRate5_Click()
    Dim tstPress As Integer
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    tstPress = (GetPlayerController(gPlayerTurn) <> 0) _
    And (Not gComputerPressed)
    If tstPress Then
        Exit Sub
    End If
    transferNmbr = 5
    If gCurrentMode = 5 Then
        Call Mode5
    End If

End Sub

Private Sub tfRate50_Click()
    Dim tstPress As Integer
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    tstPress = (GetPlayerController(gPlayerTurn) <> 0) _
    And (Not gComputerPressed)
    If tstPress Then
        Exit Sub
    End If
    transferNmbr = 50
    If gCurrentMode = 5 Then
        Call Mode5
    End If

End Sub

Private Sub tfRateAll_Click()
    Dim tstPress As Integer
    
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    tstPress = (GetPlayerController(gPlayerTurn) <> 0) _
    And (Not gComputerPressed)
    If tstPress Then
        Exit Sub
    End If
    transferNmbr = 9999
    If gCurrentMode = 5 Then
        Call Mode5
    End If

End Sub

Private Sub Timer1_Timer()
    If clickLock Then Exit Sub
    Call handleTimer1
    If SetupScreen.Visible Then   'Safety trap
        gCurrentMode = 3
        Timer1.Enabled = False
        Timer2.Enabled = False
        Exit Sub
    End If
End Sub

Private Sub handleTimer1()
    Dim vNull As Long
    
    'Mode 13, draw win and switch to mode 18.
    If gCurrentMode = 13 And Not gMapSetupLock Then
        Timer1.Enabled = False
        Call DrawWin
        Exit Sub
    
    'If mode 18, we should not have made it here.
     ElseIf gCurrentMode = 18 Then
        Timer1.Enabled = False
        Exit Sub
     
    'Mode 3, setting up the board by filling in countries in a random sequence.
    ElseIf gCurrentMode = 3 Then
        Timer2.Enabled = False
        If gCtryOrder(timerCounter) > 0 Then
            Call ColorCountry(gCtryOrder(timerCounter), gPlayerID(gCountryOwner(gCtryOrder(timerCounter))).lngColor)
            Call SyncForgroundMap("handleTimer1 1")
        End If
        Timer2.Enabled = False
        timerCounter = timerCounter + 1
        If timerCounter < 43 Then
            Exit Sub
        End If
        
        'All countries have been filled in.
        Timer1.Enabled = False
        If gPlayerTurn = 0 Then
            gPlayerTurn = 1
        End If
        gPlayerTurn = gPlayerTurn - 1
        Call SaveCheckpoint
        gPlayerTurn = gPlayerTurn + 1
        mnuOptUndo.Enabled = gCheatMode.undoEnabled  '(playerID(gPlayerTurn).playerWho = 0)
        Toolbar1.Buttons(8).Enabled = gCheatMode.undoEnabled     '(playerID(gPlayerTurn).playerWho = 0)
        Call ValidateAllStats(True)
        
        Call resetChangeList
        gPlayerValue = GetPlayerValue(gPlayerTurn)
        TheMainForm.pctInfoBox.BackColor = gPlayerID(gPlayerTurn).bkgndColor
        Call DrawLittleCards
        If flashingBorder And (GetPlayerController(gPlayerTurn) = 0) Then
            If netWorkSituation <> cNetNone Then       'Flash color if human on this terminal
                If (net.playerOwner(gPlayerTurn - 1) = myTerminalNumber) Then
                    TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
                    tmrFlashInfoBox.Enabled = True
                Else
                    TheMainForm.BackColor = &H8000000F
                End If
            Else
                TheMainForm.BackColor = gPlayerID(gPlayerTurn).bkgndColor
                tmrFlashInfoBox.Enabled = True
            End If
        Else
            TheMainForm.BackColor = &H8000000F
        End If
        
        ' You have xx units to place
        Call printPlaceUnits
        
        Call CheckToSeeMission(gPlayerTurn)
        gCurrentMode = 2
        Call AutoPlayerSelect
        gMapSetupLock = False
        Call AuditPlayerRecord
        Call AuditAddPointsIssued(gPlayerTurn, gPlayerValue)
        
        'Set sync so that the scores get printed.
        gSyncViewportNeeded = True
        Call SyncForgroundMap("handleTimer1 2")
        Exit Sub
    
    ElseIf gMapSetupLock Then
        Exit Sub
    
    'Overrun attack. Attack again.
    ElseIf gCurrentMode = 25 Then
        Call overRunAttack
        Call handleUpdate
        If GetPlayerController(gPlayerTurn) = 0 Then
            net.madeUpdate = False
        End If
        Call SyncForgroundMap("handleTimer1 3")
        Exit Sub
    End If
    
    'Should never get here, but just incase it does.
    Timer1.Enabled = False
    If gPlayerID(gPlayerTurn).card(4) = 0 Then
        ToggleKeys (gPickedUpUnits = 0)
    Else
        ToggleKeys (False)
    End If
End Sub

    'Handle attacking in overrun mode
Private Sub overRunAttack()
    Call AttackCountry
    If gCurrentMode <> 25 Then
        If gCurrentMode = 13 Then
            'Timer1.Interval = 1700
            'Timer1.Enabled = True
            Exit Sub
        End If
        Timer1.Enabled = False
        If gPlayerID(gPlayerTurn).card(4) = 0 Then
            ToggleKeys (gPickedUpUnits = 0)
        Else
            ToggleKeys (False)
        End If
        Exit Sub
    End If
    If gPickedUpUnits = 0 Then
        Call ToggleKeys(True)
        Timer1.Enabled = False
        Call AttackClicked
    ElseIf gTargetCtry = 0 Then
        Call ToggleKeys(True)
        Timer1.Enabled = False
    End If
    Call CheckWinDuringTurn(gPlayerTurn)
    Exit Sub
End Sub

'Return the number of countries pPlayer holds.
Public Function CountCountriesHeld(pPlayer As Integer) As Long
    Dim vIndex As Long
    
    CountCountriesHeld = 0
    For vIndex = 1 To 42
        If gCountryOwner(vIndex) = pPlayer Then
            CountCountriesHeld = CountCountriesHeld + 1
        End If
    Next vIndex
End Function

'Human clicked card exchange button.
Private Sub cmdExchange_Click()
    'Stop the info box from flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    Call CardExchangeClicked
    Call SyncForgroundMap("cmdExchange_Click")
End Sub

    '-------------------------
    '   Setup processes
    '-------------------------

'Check players have been setup correctly then continue.
Private Sub cmdSetupOk_Click()
    Dim tmpNmbr As Integer
    'Call updateTestViewer("cmdSetupOk_Click")
    
    If gMapSetupLock Or gWarRestartLock Then
        'Call updateTestViewer("1 cmdSetupOk_Click")
        Exit Sub
    End If
    If frmStats.Visible And netWorkSituation <> cNetHost Then
        frmStats.Hide
    End If
    
    'frmAdvanced.Hide
    If gCurrentMode = 100 Or netWorkSituation = cNetClient Then
        Call joinWar
        Exit Sub
    End If
    
    mnuOptUndo.Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    
    tmpNmbr = nmbrOfPlayers
    nmbrOfPlayers = CountActiveStartingPlayers
    If nmbrOfPlayers < 2 Then
        MsgBox Phrase(137), vbOKOnly, Phrase(221)   'Setup
        nmbrOfPlayers = tmpNmbr
        SetupScreen.Visible = True
        Call ShowMenuBar(SetupScreen.Visible)
        'mnuMsnMissionOn.Enabled = (SetupScreen.Visible And netWorkSituation <> cNetClient)
        Call EnableMissionOptions
        Exit Sub
    ElseIf CountAllocatedCountries < 42 Then
        MsgBox Phrase(138), vbOKOnly, Phrase(221)   'Setup
        nmbrOfPlayers = tmpNmbr
        SetupScreen.Visible = True
        Call ShowMenuBar(SetupScreen.Visible)
        'mnuMsnMissionOn.Enabled = (SetupScreen.Visible And netWorkSituation <> cNetClient)
        Call EnableMissionOptions
        Exit Sub
    End If
    
    '>>>>>> Commited to start new war
    
    Randomize
    gPauseActive = False
    clickLock = False
    gMapSetupLock = True
    tmrFindCPUspeed.Interval = 1700
    tmrFindCPUspeed.Enabled = True
    With A2opportunity
        .IsActive = False
        .pathPointer = 0
        .Path(0) = 0
    End With
    
    sWinMessage = ""
    
    'cardmode = Abs(CInt(optCardsFixed.Value) + CInt(optCardsIncrease.Value * 2))
    'gateDefence = gateDefenseTable(cardmode)
    If GetCardMode = 2 Then
        gMaxCardValue = CInt(txtMaximumCardValue.Text)
    End If
    'If chkCardsHidden.Value = 1 Then
    '    CardsUp = False
    'Else
    '    CardsUp = True            'CardsUp = Not (chkCardsHidden.Value)
    'End If
    gCurrentCardValue = gcCardStartValue          'Set card value
    Call AssignPlayerIDs
    
    If optPlr1FirstPlayer.Value Then           '1st player starts
        gPlayerTurn = 1
        Do
            If gPlayerID(gPlayerTurn).startWith <> 0 Then
                Exit Do
            End If
            gPlayerTurn = gPlayerTurn + 1
        Loop
    Else                             'Random player starts
        Do
            gPlayerTurn = Int(GenRandom4 * 6) + 1
            If gPlayerID(gPlayerTurn).startWith <> 0 Then
                Exit Do
            End If
        Loop
    End If
    
    'optimizeDice = (chkOptimizeDefenceDice.Value = 1)
    
    Call initializePlayers
    
    'If missions are on.
    If chkMsnMissionsOn.Value = vbChecked Then
        Call DealNewMissions
    Else
        Call ClearMissions
    End If
    
    TheMainForm.pctInfoBox.BackColor = &HFFFFFF
    TheMainForm.BackColor = &H8000000F
    Call ClearMainCardsArea
    Call ClearDiceFromBoard
    InfoBoxPrint 0
    InfoBoxPrnCR 1, 139                          'setting up the board
    
    gAskedToSeeMission = False
    boolIssueCard = False               'Hasn't got a card yet!
    SetupScreen.Visible = False
    Call ShowMenuBar(SetupScreen.Visible)
    'mnuMsnMissionOn.Enabled = (SetupScreen.Visible And netWorkSituation <> cNetClient)
    Call EnableMissionOptions
    Call ChangeTitlebarText(Phrase(35) & Trim(warSit.filename))                   'Global Siege -
    gComputerPressed = False
    boolDrawnWin = False
    notHitMove = True
    gComputerAquiredCards = False
    
    Call ToglleCardKeys(False)
    Call resetMoveTimes                'Limited moves clear
    Call ResetCardsNewWar
    Call ShowNewMap                    'Normal
    If netWorkSituation = cNetHost Then
        Call sendOwnerScoreOrder
    End If
    Call SyncForgroundMap("cmdSetupOk_Click")
End Sub

'Show a message box to remind player to check the mission.
Private Sub CheckToSeeMission(Optional playerWho As Integer = 0)
    On Error Resume Next
    If frmMissions.Visible Then
        frmMissions.Hide
    End If
    If (gCurrentMode <> 2 And gCurrentMode <> 3) _
    Or SetupScreen.Visible _
    Or Not TheMainForm.Visible _
    Or Not mnuMisSeeReminder.Checked Then
        Exit Sub
    End If
    If playerWho = 0 Then
        playerWho = gPlayerTurn
    End If
    If gPlayerID(playerWho).playerWho = 0 _
    And Not gSeenMission(playerWho) _
    And gPlayerID(playerWho).mission > 0 _
    And Not gServerMode _
    And Not gHeadlessMode Then
        frmMissions.Show , TheMainForm
    End If
End Sub

    'Set up players.
Private Sub AssignPlayerIDs()
    Dim vIndex As Integer
    
    Call resetStats
    
    For vIndex = 0 To 5
        gPlayerID(vIndex + 1).startWith = txtPlayerStartCountries(vIndex).Text
        gPlayerID(vIndex + 1).playerWho = playerSelect_getIndex(vIndex)
    Next vIndex
    
    With A2opportunity
        .IsActive = False
        .pathPointer = 0
        .Path(0) = 0
    End With
End Sub

    'Clear war statisticts for each player
Private Sub resetStats()
    Dim cntr As Integer
    
    For cntr = 0 To 5
        gPlayerStats(cntr + 1).CardsTraded = 0
        gPlayerStats(cntr + 1).CountriesDefeated = 0
        gPlayerStats(cntr + 1).CountriesLost = 0
        gPlayerStats(cntr + 1).PlayersWipedOut = 0
        gPlayerStats(cntr + 1).StartingMission = Phrase(75)
        gPlayerStats(cntr + 1).UnitsBeaten = 0
        gPlayerStats(cntr + 1).UnitsFromCards = 0
        gPlayerStats(cntr + 1).UnitsIssued = 0
        gPlayerStats(cntr + 1).UnitsLost = 0
    Next cntr
End Sub

Private Function autoWho(Hmn As Boolean, Auto1 As Boolean, Auto2 As Boolean) As Integer
    If Hmn Then
        autoWho = 0
    ElseIf Auto1 Then
        autoWho = 1
    Else
        autoWho = 2
    End If
End Function

'Counts active starting players. Active starting players are players
'with 1 or more starting countries on the setup screen.
Private Function CountActiveStartingPlayers() As Integer
    Dim vCntr As Long
    
    CountActiveStartingPlayers = 0
    
    For vCntr = 0 To 5
        If Trim(udPlayerStartCountries(vCntr).Value) <> 0 Then
            CountActiveStartingPlayers = CountActiveStartingPlayers + 1
        End If
    Next
End Function

    'Reset move limiter
Private Sub resetMoveTimes()
    Dim cntr As Integer
    
    For cntr = 1 To 42
        gMoveTimes(cntr) = 0
        gMovedIn(cntr) = 0
    Next cntr
End Sub

'Draw "Victory!" message after a short delay. This part reprints the
'map, clears the info box, cleans up a few things, sets gCurrentMode
'to 18 and fires off a timer to draw the win after a pause.
Public Sub DrawWin(Optional pRemoteTerminal As Boolean = False)
    On Error GoTo ErrHand
    Timer2.Enabled = False
    Timer1.Enabled = False
    
    If frmMissions.Visible Then
        frmMissions.Hide
    End If
    
    If boolDrawnWin _
    Or gCurrentMode = 18 _
    Or gWarRestartLock _
    Or SetupScreen.Visible _
    Or tmrDrawWin.Tag <> "" Then
        Exit Sub
    End If
    
    gWarRestartLock = True
    Call refreshMap
    
    'Sync map to viewport with scores.
    Call DrawAllCards
    Call SyncForgroundMap("drawWin")
    With A2opportunity
        .IsActive = False
        .pathPointer = 0
        .Path(0) = 0
    End With
    
    boolDrawnWin = True
    
    If gPlayerTurn = 0 Then
        gPlayerTurn = 1
    End If
    
    'Disable undo buttons unless cheat mode active.
    mnuOptUndo.Enabled = gCheatMode.undoEnabled
    Toolbar1.Buttons(8).Enabled = gCheatMode.undoEnabled
    
    'Short puuse for dramatic affect before going on with the rest.
    Timer1.Enabled = False
    Timer2.Enabled = False
    tmrDrawWin.Tag = "PrintWinMessageOnViewport"
    tmrDrawWin.Interval = 2000
    tmrDrawWin.Enabled = True
    Call frmStats.CalculateStats
    Exit Sub
ErrHand:
    Resume Next
End Sub

'Choose a font size for the win message by scaling up until it fits in the middle of the screen.
Private Sub ScaleWinMessageFont()
    Dim vCntr As Long
    Dim vFontPoints As Long
    Dim vTestStrY As String
    Dim vTestStrX As String
    Dim vMaxHeight As Long
    Dim vMaxWidth As Long
    
    'Start really really small.
    vFontPoints = 10
    
    'Load test strings with maximum expected strings.
    vTestStrY = Phrase(371)
    vTestStrX = Phrase(371) & "X"
    
    'Set up the max sizes relative to the mask size.
    'Mask4 ScaleWidth and ScaleHeight are 944 and 650
    vMaxHeight = Mask4.Map1.ScaleHeight '* 0.55
    vMaxWidth = Mask4.Map1.ScaleWidth '* 0.75
    
    'Scale up the font until it no longer fits the info box either height or width.
    For vFontPoints = 6 To 900
        Mask4.Map1.Font.Size = vFontPoints / 4
        
        If Mask4.Map1.TextHeight(vTestStrY) > vMaxHeight _
        Or Mask4.Map1.TextWidth(vTestStrX) > vMaxWidth Then
            Mask4.Map1.Font.Size = (vFontPoints - 1) / 4
            Exit For
        End If
    Next
End Sub

'Try various fonts from a list in global variable gcDrawWinFonts.
'gcDrawWinFonts is a comma delimited string of font names with the
'first fonts to try being the first in the list.
'This needs to be done because not all Windows has the same font set.
'If we set to a font that doesn't exist, it will go to the default font.
Private Sub SetWinMessageFont()
    Dim vFontName() As String
    Dim vIndex As Long
    
    On Error Resume Next
    
    vFontName = Split(gcDrawWinFonts, ",")
    
    For vIndex = 0 To UBound(vFontName)
        Mask4.Map1.Font.name = vFontName(vIndex)
        If Mask4.Map1.Font.name = vFontName(vIndex) Then
            Exit For
        End If
    Next
End Sub

'Print the Victory message with shadow effects on the viewport.
'Called from tmrDrawWin timer which was triggered by sub DrawWin.
Private Sub PrintWinMessageOnViewport()
    Dim vPrintPosX As Long
    Dim vPrintPosY As Long
    Dim vTextWidth As Long
    Dim vTextHeight As Long
    
    On Error GoTo ErrHand
    
    'Clear the cards from the map.
    Call ClearMainCardsArea
    Call CleaLittleCards
    
    'Clear the dice from the map.
    Call ClearDiceFromBoard
    
    'Set up the font.
    Mask4.Map1.Font.Bold = False
    Call SetWinMessageFont
    Call ScaleWinMessageFont
    
    'Set up the destination of the Victory message on the viewport.
    vTextWidth = Mask4.Map1.TextWidth(Phrase(371))
    vTextHeight = Mask4.Map1.TextHeight(Phrase(371))
    vPrintPosX = (Mask4.Map1.ScaleWidth - vTextWidth) * 0.5
    vPrintPosY = (Mask4.Map1.ScaleHeight - vTextHeight) * 0.2
    
    'Print a black win message to the right and below to give a shadow affect.
    Mask4.Map1.CurrentX = vPrintPosX + 10
    Mask4.Map1.CurrentY = vPrintPosY + 10
    Mask4.Map1.ForeColor = vbBlack
    Mask4.Map1.Print Phrase(371)        '"Victory!"
    
    'Decide which colour to use.
    If mnu3Ddisplay.Checked Then
        Mask4.Map1.ForeColor = RGB(241, 254, 1)
    Else
        Mask4.Map1.ForeColor = vbYellow
    End If
    
    'Print the main win message.
    Mask4.Map1.CurrentX = vPrintPosX
    Mask4.Map1.CurrentY = vPrintPosY
    Mask4.Map1.Print Phrase(371)        '"Victory!"
    
    'Make sure the viewport gets refreshed.
    gSyncViewportNeeded = True
    
    'Set current mode to 18, win message has been drawn.
    gCurrentMode = 18
    
    'Clear and print the win message to the info box.
    InfoBoxPrint 0
    TheMainForm.pctInfoBox.Print sWinMessage
    TheMainForm.pctInfoBox.Tag = sWinMessage
    
    'Game won, no info box flashing.
    tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    'Sync the background image to the viewport without printing the scores.
    Call SyncForgroundMap("PrintWinMessageOnMap")
    
    'Short puuse for dramatic affect before going on with the rest.
    Timer1.Enabled = False
    Timer2.Enabled = False
    tmrDrawWin.Tag = "ShowStatsAfterWin"
    tmrDrawWin.Interval = 2000
    tmrDrawWin.Enabled = True
    Exit Sub
ErrHand:
    Resume Next
End Sub

'Show the stats after printing the win message then trigger off another delay
'before restarting, if auto restart is on.
Private Sub ShowStatsAfterWin()
    Dim drX As Integer
    Dim drY As Integer
    Dim cntr1 As Long
    Dim cntr2 As Integer
    Dim Pclr As Long
    Dim countSpeed As Long
    Const numberOfDots = 510
    Dim dummy As Long
    Dim destX As Long
    Dim destY As Long
    
    On Error GoTo ErrHand
    
    'Show stats if player wants to see them.
    If Not SetupScreen.Visible _
    And gCurrentMode = 18 Then
        Call frmStats.ShowStats
    End If
    
    gWarRestartLock = False
    
    'This will be the second printing of Clem's counter
    'if the winner is a remote terminal.
    Call AuditPrintClemsCounter
    Call netMain.ClearForfeitVotes
    
    'Fire restart timer if auto restart is selected.
    If TheMainForm.mnuAutoRestart.Checked _
    And netWorkSituation <> cNetClient Then
        ' Survey of player wins, auto restart
        Plr(gPlayerTurn) = Plr(gPlayerTurn) + 1
        cntr2 = 0
        Debug.Print "---------------------------------"
        For cntr1 = 1 To 6
            Debug.Print gPlayerID(cntr1).strColor; "   ", "Defence:"; playerDefence(cntr1), "Won:"; Plr(cntr1)
            cntr2 = cntr2 + CInt(Plr(cntr1))
        Next
        Debug.Print
        Debug.Print "Total games played:", cntr2
        Debug.Print "---------------------------------"
        
        'Set restart timer if auto restart is switched on.
        Timer1.Enabled = False
        Timer2.Enabled = False
        tmrDrawWin.Tag = "AutoRestartAfterWin"
        tmrDrawWin.Interval = GetSetting(gcApplicationName, "settings", "AutoRestartDelay", 9000)
        tmrDrawWin.Enabled = True
    Else
        tmrDrawWin.Tag = ""
    End If
    Call SyncForgroundMap("ShowStatsAfterWin")
    Exit Sub
ErrHand:
    Resume Next
End Sub

'Restart the war by pressing the restart button.
Private Sub AutoRestartAfterWin()
    tmrDrawWin.Enabled = False
    tmrDrawWin.Tag = ""
    If gCurrentMode = 18 Then
        cmdSetupOk_Click
    End If
End Sub

'Timer used for the draw win delay and also to show stats and restart the war.
Private Sub tmrDrawWin_Timer()
    On Error Resume Next
    
    tmrDrawWin.Enabled = False
    
    Select Case tmrDrawWin.Tag
    
    Case "PrintWinMessageOnViewport"
        Call PrintWinMessageOnViewport
        
    Case "ShowStatsAfterWin"
        Call ShowStatsAfterWin
        
    Case "AutoRestartAfterWin"
        Call AutoRestartAfterWin
        
    Case Else
        tmrDrawWin.Tag = ""
        gWarRestartLock = False
    End Select
End Sub


    '--------------------------
    '   Auto player functions
    '--------------------------
    
    
    'Select player, if auto player then turn on timer2
Private Sub AutoPlayerSelect()
    If gCurrentMode = 13 Then
        Timer2.Enabled = False
        Exit Sub
    ElseIf gCurrentMode = 18 Then
        Timer2.Enabled = False
        Exit Sub
    End If
    
    'If player is not on this terminal.
    If netWorkSituation <> cNetNone Then
        If net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber Then
            Timer2.Enabled = False
            Exit Sub
        End If
    End If
    
    If IsComputerPlayer(gPlayerTurn) Then
        Call AutoCheckCards
        Timer2.Interval = playSpeed * 16
        Timer2.Enabled = True
        Exit Sub
    End If
End Sub

    'Computer player's timer
Private Sub Timer2_Timer()
    If clickLock Then Exit Sub
    'Call updateTestViewer("Start T2")
    Call handleTimer2
    If SetupScreen.Visible Then   'Safety trap
        gCurrentMode = 3
        Timer1.Enabled = False
        Timer2.Enabled = False
        Exit Sub
    End If
    ''Call updateTestViewer("End T2")
End Sub

Private Sub handleTimer2()
    If gMapSetupLock Or gWarRestartLock Then
        Exit Sub
    ElseIf gCurrentMode = 13 Then
        Timer2.Enabled = False
        Exit Sub
    ElseIf gCurrentMode = 18 Then
        Timer2.Enabled = True
        Exit Sub
    End If

    If GetPlayerController(gPlayerTurn) = 0 Then      ' Stop computer players controlling human
        Timer2.Enabled = False
        'Call updateTestViewer("Who = 0 ")
        Exit Sub
    End If
    
    If netWorkSituation <> cNetNone Then                   ' Stop computer players controlling remote player
        If net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber Then
            Timer2.Enabled = False
            'Call updateTestViewer("Remote player ")
            Exit Sub
        End If
    End If
    
    Call processComputerMode
    
    Call handleUpdate
    net.madeUpdate = False
    If IsComputerPlayer(gPlayerTurn) Then
        Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
        Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
    End If
    
    Call SyncForgroundMap("handleTimer2")
End Sub

    'Handle mode processing got computer players 1 and 2
Private Sub processComputerMode()
    Dim tst As Boolean
    Dim tmp As Integer
    'Debug.Print "Timer2_timer: Currentmode="; gCurrentMode; ", gComputerAquiredCards="; gComputerAquiredCards
    
    If gPlayerID(gPlayerTurn).playerWho = A3Index Then
        Timer2.Enabled = False
        Timer1.Enabled = False
        Call A3ProcessIntelligentPlayer
        'Call updateTestViewer("T2 process A3 ")
        Exit Sub
    End If
    
    If gCurrentMode = 6 Then
        gCurrentMode = 2
    End If
    
    If gComputerAquiredCards Then
        Timer2.Enabled = False
        cmdAttack.Enabled = cmdMove.Enabled
        mnuAttack.Enabled = cmdAttack.Enabled
        notHitMove = True
        tmp = gCurrentMode
        Call AutoCheckCards
        If gCurrentMode <> 8 Then
            gCurrentMode = tmp
        End If
        gComputerAquiredCards = False  'change later
        Timer2.Interval = playSpeed * 22
        'Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
        'Call updateTestViewer("T2 Cards aquired ")
        Exit Sub
    End If
    
    If gCurrentMode = 8 Then      'Tell humans turning in cards
        Timer2.Enabled = False
        Call AutoTurn3In(gDiceArray(11))
        Call SyncForgroundMap("processComputerMode")
        Exit Sub
        
    ElseIf gCurrentMode = 7 Then       'Turn in cards
        Timer2.Enabled = False
        gComputerPressed = True
        Call CardExchangeClicked
        gComputerPressed = False
        gCurrentMode = 2
        net.madeUpdate = False
        resetChangeList
        net.changeList(0) = 0
        If gPlayerID(gPlayerTurn).card(4) > 0 Then
            gComputerAquiredCards = True
        End If
        Timer2.Interval = playSpeed * 22
        'Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
        'Call updateTestViewer("T2 mode 7 ")
        Exit Sub
        
    ElseIf gCurrentMode = 2 Then        'Place units
        Timer2.Interval = playSpeed * 18
        Timer2.Enabled = False
        If gPlayerID(gPlayerTurn).playerWho = 1 Then
            Call AutoPlayer1
        Else
            Call AutoPlayer2
            If AutoCountry < 1 Then     'Just in case, should never get here (bodgy!)
                AutoCountry = 1
            End If
        End If
        'Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
        'Call updateTestViewer("T2 Mode2 ")
        Exit Sub
        
    ElseIf gCurrentMode = 1 Then
        Timer2.Interval = playSpeed * 13
        Timer2.Enabled = False
        If gPlayerID(gPlayerTurn).playerWho = 1 Then
            Call Auto1Attack        'Attack
        Else
            Call Auto2Attack
        End If
        'Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
        'Call updateTestViewer("T2 mode 1 ")
        Exit Sub
        
    ElseIf gCurrentMode = 20 Then    'Attack Who
        Timer2.Enabled = False
        If gPlayerID(gPlayerTurn).playerWho = 1 Then
            Call AutoMode20
        Else
            Call Auto200Mode20
        End If
        'Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
        If gCurrentMode = 20 And Not A2opportunity.IsActive Then    'Attempt to stop that stupid Hang state
            'Look for an In-your-face easy attack.
            If Not EasyAttack Then
                gCurrentMode = 16
            Else
                prefercont = 0
            End If
        End If
        'Call updateTestViewer("T2 mode 20")
        Exit Sub
        
    ElseIf gCurrentMode = 21 Then    'Attack from
        Timer2.Enabled = False
        If gPlayerID(gPlayerTurn).playerWho = 1 Then
            Call AutoMode21
        Else
            Call Auto2Mode21
        End If
        'Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
        If gCurrentMode = 21 And Not IsEasyAttackOn Then    'Attempt to stop that stupid Hang state
            gCurrentMode = 16
        End If
        'Call updateTestViewer("T2 mode 21")
        Exit Sub
        
    ElseIf gCurrentMode = 22 Then
        Timer2.Enabled = False
        Call ResetEasyAttackConts
        Call AutoMode22
        'If gCurrentMode <> 13 And gCurrentMode <> 18 Then
        '    Timer2.Enabled = True
        'Else
            'Call updateTestViewer("T2 mode 22a")
            Exit Sub
        'End If
        If gPickedUpUnits > 0 Then
            'Call updateTestViewer("T2 mode 22b")
            Exit Sub
        End If
        Exit Sub
        'Call updateTestViewer("T2 mode 22c")
        
    ElseIf gCurrentMode = 23 Then
        Timer2.Enabled = False
        Call AutoMode23
        If gCurrentMode <> 13 And gCurrentMode <> 18 Then
            If gCurrentMode = 1 Then
                Timer2.Interval = playSpeed * 18
            End If
            'Timer2.Enabled = True
        Else
            'Call updateTestViewer("T2 mode 23a")
            Exit Sub
        End If
        If GetPlayerController(gPlayerTurn) = 2 Then
            If gCurrentMode = 1 And AtkForCard.On Then
                Call Auto1Move
                'Call updateTestViewer("T2 mode 23b")
                Exit Sub
            End If
            'Call updateTestViewer("T2 mode 23d")
            Exit Sub
        End If
        If gPickedUpUnits > 0 Then
            'Call updateTestViewer("T2 mode 23c")
            Exit Sub
        End If
        
    ElseIf gCurrentMode = 10 Then
        Timer2.Enabled = False
        If GetPlayerController(gPlayerTurn) = 2 Then
            Call AutoMode10        'Change to end of turn
        Else
            Call AutoMode10
        End If
        'Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
        If gCurrentMode = 10 And Not A2FindStranded Then
            gCurrentMode = 14
        End If
        'Call updateTestViewer("T2 mode 10")
        Exit Sub
        
    ElseIf gCurrentMode = 11 Then
        Timer2.Enabled = False
        Call AutoMode11
        If A2MoveList.Pointer <> 0 And gCurrentMode <> 16 Then        'Move to another country (same origin)
            gComputerPressed = True
            Call MoveClicked
            gComputerPressed = False
            AutoCountry = A2MoveList.List(A2MoveList.Pointer)   'CtryTo
            Call ClickMapNow
            'Timer2.Enabled = True
            'Call updateTestViewer("T2 mode 11a")
            Exit Sub
        End If
        If AutoSecondMove Then
            AutoSecondMove = False
            AutoSecondMove = False
            gCurrentMode = 10
            Timer2.Interval = playSpeed * 14
            'Timer2.Enabled = True
        ElseIf A2FindStranded Then  '1
            gCurrentMode = 16
            Timer2.Interval = playSpeed * 14
            'Timer2.Enabled = True
        Else
            gCurrentMode = 14    'End after timer
        End If
        'Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
        'Call updateTestViewer("T2 mode 11b")
        Exit Sub
    ElseIf gCurrentMode = 16 Then
        With AtkForCard
            .from = 0
            .To = 0
            .On = False
        End With
        If gCtryScore(gTargetCtry) < 2 Or gCountryOwner(gTargetCtry) <> gPlayerTurn Then
            If Not A2FindStranded Then
                gCurrentMode = 14
                Timer2.Interval = playSpeed * 18
                'Timer2.Enabled = True
                'Call updateTestViewer("T2 mode 16a")
                Exit Sub
            End If
        End If
        gComputerPressed = True
        Call MoveClicked
        gComputerPressed = False
        Timer2.Interval = playSpeed * 14
            If gCurrentMode <> 10 Then
                gCurrentMode = 14
                Timer2.Interval = playSpeed * 18
            End If
        'Timer2.Enabled = True
        'Call updateTestViewer("T2 mode 16b")
        Exit Sub
    ElseIf gCurrentMode = 14 Then
        With AtkForCard
            .from = 0
            .To = 0
            .On = False
        End With
        gCurrentMode = 1
        Timer2.Enabled = False
        gComputerPressed = True
        Call EndClicked
        gComputerPressed = False
        'Call updateTestViewer("T2 mode 14a")
        Exit Sub
    End If
    
    If AutoGoAgain Then
        Timer2.Interval = playSpeed * 12
        gCurrentMode = 1
        'Timer2.Enabled = True
        'Call updateTestViewer("T2 AutoGoAgain ")
        Exit Sub
    ElseIf gCountryOwner(gTargetCtry) = gPlayerTurn Then
        If gCtryScore(gTargetCtry) > 1 Then
            Timer2.Enabled = False
            Call Auto1Move
            'Timer2.Enabled = gCurrentMode <> 13 And gCurrentMode <> 18
            If gCurrentMode = 10 Then
                'Call updateTestViewer("T2 mode 10x")
                Exit Sub
            End If
        End If
    End If
    Timer2.Interval = playSpeed * 15
    'Timer2.Enabled = True
    gCurrentMode = 14
    'Call updateTestViewer("T2 fell through ")
    Exit Sub
End Sub

    'Start move thingy
Private Sub Auto1Move()
    Dim tst1 As Integer
    Dim cntr As Integer
    
    If AtkForCard.On Then
        Call A200MoveBack
        Exit Sub
    End If
    
    If gTargetCtry = 0 Then
        gCurrentMode = 14
        Exit Sub
    End If
    
    tst1 = (gCountryOwner(gTargetCtry) = gPlayerTurn) And (gCtryScore(gTargetCtry) > 1)
    If tst1 Then
        If enemyAbout(gTargetCtry, gPlayerTurn, 0) Then
            Exit Sub
        End If
        Call AutoGetMove
    End If
End Sub

Private Sub A200MoveBack()
    Dim cntr As Integer
    
    If A200outGunned(AtkForCard.To) _
    Or enemyAbout(AtkForCard.from, gPlayerTurn, 0) Then
        AutoCountry = AtkForCard.from
        Call AutoGetMove
    Else
        If A2FindStranded Then  '4
            gCurrentMode = 16
            Timer2.Interval = playSpeed * 14
            Timer2.Enabled = True
        Else
            gCurrentMode = 14
        End If
    End If
End Sub

    'Return T if ctryWhere score > all enemy score on borders + getAvg
Private Function A200outGunned(ctryWhere As Integer) As Boolean
    Dim cntr As Integer
    Dim howMany As Integer
    Dim eTotal As Long
    
    eTotal = 2 '(getAvg / gateDefence * 2)
    For cntr = 1 To 7
        If CountryID(ctryWhere).neighbour(cntr) = 0 Then Exit For
        If gCountryOwner(CountryID(ctryWhere).neighbour(cntr)) _
        <> gCountryOwner(ctryWhere) Then
            eTotal = eTotal + gCtryScore(CountryID(ctryWhere).neighbour(cntr)) - 1
        End If
    Next cntr
    
    If (eTotal / (playerDefence(gPlayerTurn))) > (gCtryScore(ctryWhere)) Then
        A200outGunned = True
    Else
        A200outGunned = False
    End If
End Function

Private Sub auto2move()
    Dim tst1 As Integer
    Dim cntr As Integer
    
    tst1 = (gCountryOwner(gTargetCtry) = gPlayerTurn) And (gCtryScore(gTargetCtry) > 1)
    If tst1 Then
        Call Auto2GetMove
    End If
End Sub

    'Make sure there is a country to move into nearby
Private Sub AutoGetMove()
    Dim cntr As Integer
    
    Timer2.Interval = playSpeed * 13
    Timer2.Enabled = True
    gComputerPressed = True
    Call MoveClicked
    gComputerPressed = False
    Exit Sub        'found one
End Sub

Private Sub Auto2GetMove()
    Dim cntr As Integer, ctry As Integer
    Dim tst As Boolean
    
    For cntr = 1 To 7
        ctry = CountryID(gTargetCtry).neighbour(cntr)
        If ctry = 0 Then
            Exit For
        End If
        
        If gCountryOwner(ctry) = gPlayerTurn Then
            tst = (autoIsAGate(ctry, gPlayerTurn)) _
                And (enemyInCont(ContinentOfCtry(ctry), ctry, gPlayerTurn))
    
            If tst Or (enemyAbout(ctry, gPlayerTurn, 0)) Then
                Timer2.Interval = playSpeed * 14
                Timer2.Enabled = True
                gComputerPressed = True
                Call MoveClicked
                gComputerPressed = False
                Exit Sub        'found one
            End If
        End If
    Next cntr
End Sub

    'Move to
Private Sub AutoMode10()
    Dim cntr As Integer, cntr2 As Integer, rslt As Integer
    Dim prospects(42) As Integer
    Dim Pointer As Integer
    Dim tst As Boolean
    Dim TmpNbr As Integer
    Dim TmpNbr1 As Integer
    Dim firstChoice As Integer
    Dim Neigbors(7) As Integer
    
    Pointer = 1
    For cntr = 1 To 42
        prospects(cntr) = 0
    Next cntr
    
        'Is next to ctry touching enemy
    Timer2.Interval = playSpeed * 16
    
    If AtkForCard.On Then
        nextCountryMove = 0
        If gCountryOwner(AtkForCard.from) = gPlayerTurn Then
            AutoCountry = AtkForCard.from
            Call ClickMapNow
            Exit Sub
        Else
            AtkForCard.On = False
            AtkForCard.from = 0
            AtkForCard.To = 0
        End If
    End If
    
    nextCountryMove = 0
    gSourceCtry = gTargetCtry
    If gCtryScore(gTargetCtry) < 2 Then
        gComputerPressed = True
        Call EndClicked
        gComputerPressed = False
        Exit Sub
    End If
    
    Call A3.fillRandNeigborList(gTargetCtry, Neigbors)
    firstChoice = 0
    With A2MoveList
    .Pointer = 0
    
        'Find all countries to move into
    If Not enemyAbout(gTargetCtry, gPlayerTurn, 0) Then
        'Any neigbours facing enemy directly?
        For cntr = 1 To 7
            TmpNbr = CountryID(gTargetCtry).neighbour(cntr)
            If TmpNbr = 0 Then
                Exit For
            End If
                'Found one facing an enemy
            If gCountryOwner(TmpNbr) = gPlayerTurn _
            And enemyAbout(TmpNbr, gPlayerTurn, gTargetCtry) Then
                Call AddToMoveList(TmpNbr)
            End If
        Next cntr
    End If
    
    If Not enemyAbout(gTargetCtry, gPlayerTurn, 0) Then
        'Look for friendly neighbour in hostile continent. A2 only.
        If .Pointer = 0 Then
            For cntr = 1 To 7
                TmpNbr = CountryID(gTargetCtry).neighbour(cntr)
                If TmpNbr = 0 Then
                    Exit For
                End If
                If gCountryOwner(TmpNbr) = gPlayerTurn And autoIsAGate(TmpNbr, gPlayerTurn) Then
                    If Not OwnContinent(ContinentOfCtry(TmpNbr), gPlayerTurn) Then
                        'Found one in an unfriendly continent
                        Call AddToMoveList(TmpNbr)
                    End If
                End If
            Next
        End If
        
        If .Pointer = 0 Then
            'If next to ctry touching gate with an unfriendly continent?
            For cntr = 1 To 7
                TmpNbr = CountryID(gTargetCtry).neighbour(cntr)
                If TmpNbr = 0 Then Exit For
                If autoIsAGate(TmpNbr, gPlayerTurn) _
                And Not OwnContinent(ContinentOfCtry(TmpNbr), gPlayerTurn) Then
                    Call AddToMoveList(TmpNbr)
                End If
            Next cntr
        End If
        
        'Any neigbours with neigbours facing enemy directly?
        If .Pointer = 0 Then
            For cntr = 1 To 7
                TmpNbr = CountryID(gTargetCtry).neighbour(cntr)
                If TmpNbr = 0 Then
                    Exit For
                End If
                    'Found one facing an enemy
                If gCountryOwner(TmpNbr) = gPlayerTurn _
                And Not enemyAbout(TmpNbr, gPlayerTurn, gTargetCtry) Then
                    For cntr2 = 1 To 7
                        TmpNbr1 = CountryID(TmpNbr).neighbour(cntr2)
                        If TmpNbr1 = 0 Then
                            Exit For
                        End If
                        If enemyAbout(TmpNbr1, gPlayerTurn, gTargetCtry) Then
                            'Found one facing an unfriendly continent
                            Call AddToMoveList(TmpNbr)
                            Exit For
                        End If
                    Next
                End If
            Next cntr
        End If
        
        'Look for friendly neighbour facing hostile continent.
        If .Pointer = 0 Then
            For cntr = 1 To 7
                TmpNbr = CountryID(gTargetCtry).neighbour(cntr)
                If TmpNbr = 0 Then
                    Exit For
                End If
                If gCountryOwner(TmpNbr) = gPlayerTurn And autoIsAGate(TmpNbr, gPlayerTurn) Then
                    For cntr2 = 1 To 7
                        TmpNbr1 = CountryID(TmpNbr).neighbour(cntr2)
                        If TmpNbr1 = 0 Then
                            Exit For
                        End If
                        If Not OwnContinent(ContinentOfCtry(TmpNbr1), gPlayerTurn) Then
                            'Found one facing an unfriendly continent
                            Call AddToMoveList(TmpNbr)
                            Exit For
                        End If
                    Next
                End If
            Next
        End If
        
        If .Pointer = 0 Then
            'If next to ctry touching gate with a friendly continent?
            For cntr = 1 To 7
                TmpNbr = CountryID(gTargetCtry).neighbour(cntr)
                If TmpNbr = 0 Then Exit For
                If autoIsAGate(TmpNbr, gPlayerTurn) And OwnContinent(ContinentOfCtry(gTargetCtry), gPlayerTurn) Then
                    Call AddToMoveList(TmpNbr)
                End If
            Next cntr
        End If
        
        If .Pointer = 0 Then
            'If not a gate, next to neighbour with a neighbour on a gate?
            For cntr = 1 To 7
                TmpNbr = CountryID(gTargetCtry).neighbour(cntr)
                If TmpNbr = 0 Then Exit For
                For cntr2 = 1 To 7
                    TmpNbr1 = CountryID(TmpNbr).neighbour(cntr2)
                    If TmpNbr1 = 0 Then
                        Exit For
                    End If
                    If autoIsAGate(TmpNbr1, gPlayerTurn) Then
                        'Found one facing an unfriendly continent
                        Call AddToMoveList(TmpNbr)
                        Exit For
                    End If
                Next
            Next cntr
        End If

        'Is ctry 42 (only remote ctry)? Allow dumb player.
        If .Pointer = 0 And gTargetCtry = 42 Then
            Call AddToMoveList(40)
            Call AddToMoveList(41)
        End If
    End If
    
    If .Pointer > 0 Then
        Call MoveTransferAverage
        AutoCountry = .List(1)
        Call ClickMapNow
    Else
        gMoveTimes(gTargetCtry) = 2
    End If
    End With
End Sub

    ' Add move to list and increment pointer
Private Sub AddToMoveList(CountryNumber As Integer)
    With A2MoveList
    If GetPlayerController(gPlayerTurn) = 1 And A2MoveList.Pointer > 1 And CountryNumber < 39 Then
       'Limit dumb players to 2 moves only unless in Australia.
       Exit Sub
    End If
    If .Pointer < 7 And Not InMoveList(CountryNumber) Then
        .Pointer = .Pointer + 1
        .List(.Pointer) = CountryNumber
    End If
    End With
End Sub

    ' Return True if CountryNumber already in move list
Private Function InMoveList(CountryNumber As Integer) As Boolean
    Dim cntr As Integer
    With A2MoveList
    InMoveList = False
    For cntr = 1 To .Pointer
        If CountryNumber = .List(.Pointer) Then
            InMoveList = True
        End If
    Next
    End With
End Function

    'Work out the number of units to move into the each country.
Private Sub MoveTransferAverage()
    Dim avg As Integer
    Dim cntr As Integer
    Dim cntry As Integer
    
    With A2MoveList
    cntry = gCtryScore(gTargetCtry)
    For cntr = 1 To 7
        .Transfer(cntr) = 0
    Next
    If .Pointer > 1 Then
        Do While cntry > 1
            cntr = GetMinFromMoveList
            .Transfer(cntr) = .Transfer(cntr) + 1
            cntry = cntry - 1
        Loop
    End If
    .Transfer(1) = 9999
    End With
End Sub

    'Get country in MoveList with lowest score.
Private Function GetMinFromMoveList() As Integer
    Dim cntr As Integer
    Dim LowestScore As Integer
    LowestScore = 9999
    
    With A2MoveList
    For cntr = 1 To .Pointer
        If gCtryScore(.List(cntr)) + .Transfer(cntr) < LowestScore Then
            GetMinFromMoveList = cntr
            LowestScore = gCtryScore(.List(cntr)) + .Transfer(cntr)
        End If
    Next
    End With
End Function

    'Work out the number of units to move into the first country
Private Function firstMoveTF(cFrom As Integer, fChoice As Integer, cNext As Integer) As Integer
    Dim avg As Integer
    
    avg = (gCtryScore(cFrom) + gCtryScore(fChoice) + gCtryScore(cNext)) / 2
    firstMoveTF = avg - gCtryScore(fChoice)
    If firstMoveTF < 0 Then
        firstMoveTF = 0
    End If
End Function

    'Returns true if country is a gate.
    'Used by A3 to load private array.
    'Use "A3.IsAgate" instead of this function for speed.
    'If given a player, gates are modified depending on what is held.
Public Function autoIsAGate(country As Integer, Optional playerWho As Integer = 0) As Boolean
    Dim cntr As Integer, cntr2 As Integer, Cont As Integer
    Dim TempGate(6) As Integer
    
    autoIsAGate = False
    Cont = ContinentOfCtry(country, playerWho)
    If Cont = 0 Then        'Trap
        Exit Function
    End If
    
    Call SetTempGatesForCont(Cont, TempGate, playerWho)
    For cntr = 1 To 5
        If TempGate(cntr) = 0 Then
            Exit For
        ElseIf TempGate(cntr) = country Then
            autoIsAGate = True
            Exit For
        End If
    Next cntr
End Function

    'Set gates depending on what is currently held
Private Sub SetTempGatesForCont(Cont As Integer, TempGates() As Integer, Optional playerWho As Integer = 0)
    Dim cntr As Integer
    
    For cntr = 1 To 5
        TempGates(cntr) = Continents(Cont - 1).GateCountries(cntr)
    Next
    
    If playerWho > 0 Then
        If Cont = 1 Then
            'Hold North America
            If OwnContinent(1, playerWho) Then
                If gCountryOwner(38) = playerWho Then TempGates(1) = 38     '1 -> 38
                If gCountryOwner(14) = playerWho Then TempGates(2) = 14     '3 -> 14
                If gCountryOwner(10) = playerWho Then TempGates(3) = 10     '9 -> 10
            End If
        ElseIf Cont = 2 Then
            'South America
            If OwnContinent(2, playerWho) Then
                If gCountryOwner(9) = playerWho Then TempGates(1) = 9       '10 -> 9
                If gCountryOwner(21) = playerWho Then TempGates(2) = 21     '12 -> 21
            End If
        ElseIf Cont = 3 Or Cont = 4 Then
            'Hold Europe, Africa and Mid East?
            If OwnContinent(3, playerWho) _
            And OwnContinent(4, playerWho) _
            And gCountryOwner(27) = playerWho Then
                TempGates(1) = 14
                TempGates(2) = 20
                TempGates(3) = 21
                TempGates(4) = 27
                TempGates(5) = 0
                If OwnContinent(1, playerWho) And gCountryOwner(3) = playerWho Then
                    TempGates(1) = 3       '14 -> 3
                End If
                If OwnContinent(2, playerWho) And gCountryOwner(12) = playerWho Then
                    TempGates(3) = 12       '21 -> 12
                End If
            'Europe
            ElseIf Cont = 3 And OwnContinent(3, playerWho) Then
                If gCountryOwner(3) = playerWho Then TempGates(1) = 3       '14 -> 3
                If gCountryOwner(21) = playerWho Then TempGates(2) = 21     '18 -> 21
            'Africa
            ElseIf Cont = 4 And OwnContinent(4, playerWho) Then
                If gCountryOwner(12) = playerWho Then TempGates(1) = 12     '21 -> 12
                If gCountryOwner(27) = playerWho Then TempGates(3) = 27     '23 -> 27
            End If
        ElseIf Cont = 5 Then
            'Hold Asia and Ukraine?
            If OwnContinent(5, playerWho) _
            And gCountryOwner(20) = playerWho Then
                TempGates(1) = 20
                TempGates(2) = 27
                TempGates(3) = 30
                TempGates(4) = 38
                TempGates(5) = 0
                If gCountryOwner(1) = playerWho Then TempGates(4) = 1           '38 -> 1
            ElseIf gCountryOwner(1) = playerWho Then
                TempGates(5) = 1           '38 -> 1
            End If
            If gCountryOwner(39) = playerWho Then TempGates(3) = 39         '30 -> 39
        ElseIf Cont = 6 Then
            If gCountryOwner(29) = playerWho _
            And gCountryOwner(33) = playerWho Then
                TempGates(1) = 29       '30 -> 29
                TempGates(2) = 33       '30 -> 33
                TempGates(3) = 30       '30 -> 33
            ElseIf gCountryOwner(30) = playerWho Then
                TempGates(1) = 30         '39 -> 30
            End If
        End If
    End If
End Sub

    'Move from
Private Sub AutoMode11()
    Dim tmp As Integer

    If AtkForCard.On Then
        AutoCountry = AtkForCard.To
        
        If OwnContinent(ContinentOfCtry(AtkForCard.from), gPlayerTurn) Then
            transferNmbr = getAvg / (playerDefence(gPlayerTurn) * 1.5)
        Else
            transferNmbr = getAvg / (playerDefence(gPlayerTurn) * 2)
        End If
        
        Call ClickMapNow
        gCurrentMode = 16
        Exit Sub
    Else
        AutoCountry = gSourceCtry
    End If
    
    With A2MoveList
    If .Pointer > 0 Then
        tmp = gSourceCtry
        gTargetCtry = .List(.Pointer)
        transferNmbr = .Transfer(.Pointer)
        Call ClickMapNow
        gSourceCtry = tmp
        .Pointer = .Pointer - 1
    End If
    End With
    Exit Sub
    
    '---------------------------
    If nextCountryMove = 0 Then
        transferNmbr = 9999
        Call ClickMapNow
    Else
        tmp = gSourceCtry
        Call ClickMapNow
        gSourceCtry = tmp
    End If
    
End Sub

    'Return true if worth attacking again
Private Function AutoGoAgain() As Boolean
    Dim rtrn As Integer
    Dim tst As Boolean
    
    rtrn = gCtryScore(gTargetCtry)
    
    If gPlayerID(gPlayerTurn).playerWho = 1 Then      'Dumb puter
        If gCountryOwner(gTargetCtry) <> gPlayerTurn Then
            AutoGoAgain = False
            Exit Function
        ElseIf rtrn < (getAvg) Then
            AutoGoAgain = False
        Else
            AutoGoAgain = True
        End If
        Exit Function
    End If
    
    AutoGoAgain = False
    Exit Function
    '**************** forget it !!! ***********************
    
    If gCountryOwner(gTargetCtry) = gPlayerTurn Then
        'AutoGoAgain = False
        'Exit Function
    End If
    
    With A2opportunity                  'Has a path been defined?
    If .IsActive Then                     'Follow pre set path
        If .Path(A2opportunity.pathPointer) > 0 Then         'Still going
            If gCountryOwner(.Path(A2opportunity.pathPointer)) = gPlayerTurn Then
                .IsActive = False
                .pathPointer = 0
                .Path(0) = 0
            Else
                AutoGoAgain = True
                Exit Function
            End If
        Else                            'End of path
            .IsActive = False
            .pathPointer = 0
            .Path(0) = 0
        End If
    End If
    End With
    
    If A3.findOpportunity(0) Then       'Try to find another path
        AutoGoAgain = True
        Exit Function
    End If
    
    If AtkForCard.On Then
        AutoGoAgain = False
        Exit Function
    End If
    
    tst = (OwnContinent(ContinentOfCtry(gTargetCtry), gPlayerTurn)) _
            Or (rtrn < 3)
    tst = tst And (rtrn < getAvg)
    If tst Then
        AutoGoAgain = False
        Exit Function
    End If

    AutoGoAgain = True
End Function

    ' Return the average score of the other players
Private Function averageScore() As Integer
    Dim cntr As Integer, rtrn As Integer
    
    rtrn = 0
    For cntr = 1 To 42
        'rtrn = rtrn + gCtryScore(cntr)
        If gCountryOwner(cntr) <> gPlayerTurn Then   ' Don't count my own points
            rtrn = rtrn + gCtryScore(cntr)
        Else
            rtrn = rtrn + 1
        End If
    Next cntr
    rtrn = Int(rtrn / 42)
    averageScore = rtrn
End Function

Private Sub AutoMode23()
    AutoCountry = gTargetCtry
    Call ClickMapNow
End Sub

    'Pickup armies
Private Sub AutoMode22()
    transferNmbr = 9999
    AutoCountry = gTargetCtry
    Call ClickMapNow
    
    Timer2.Interval = diceSpeed
End Sub

    'Attack with everything not on a border with another enemy
Private Sub AutoFindAttacker()
    Dim cntr As Integer, rtrn As Integer
    Dim tst1 As Boolean, tst2 As Boolean
    
    For cntr = 1 To 7
        If CountryID(gTargetCtry).neighbour(cntr) = 0 Then
            Exit For
        ElseIf gCountryOwner(CountryID(gTargetCtry).neighbour(cntr)) _
        = gPlayerTurn And gCtryScore(CountryID(gTargetCtry).neighbour(cntr)) > 1 Then 'If Not enemyAbout(countryID(gTargetCtry).neighbour(cntr), gPlayerTurn, gTargetCtry) Then
            AutoCountry = CountryID(gTargetCtry).neighbour(cntr)
            tst1 = (OwnContinent(ContinentOfCtry(AutoCountry, gPlayerTurn), gPlayerTurn))
            tst2 = (ContinentOfCtry(AutoCountry, gPlayerTurn) <> ContinentOfCtry(gTargetCtry, gPlayerTurn))
            If enemyAbout(CountryID(gTargetCtry).neighbour(cntr), gPlayerTurn, gTargetCtry) Then
                transferNmbr = gCtryScore(AutoCountry) - (getAvg / (playerDefence(gPlayerTurn)))
                If transferNmbr < 1 Then
                    transferNmbr = 0
                End If
            ElseIf tst1 Then
                transferNmbr = gCtryScore(AutoCountry) - (getAvg / (playerDefence(gPlayerTurn) * 2))
                If transferNmbr < 1 Then
                    transferNmbr = 0
                End If
            ElseIf tst2 Then
                transferNmbr = gCtryScore(AutoCountry) - (getAvg / (playerDefence(gPlayerTurn) * 2))
                If transferNmbr < 1 Then
                    transferNmbr = 0
                End If
            Else
                transferNmbr = 9999
            End If
            Call ClickMapNow
        End If
    Next cntr
    If gPickedUpUnits = 0 Then
        With A2opportunity
        .IsActive = False
        .pathPointer = 0
        .Path(0) = 0
        End With
        If AutoGoAgain Then
            gCurrentMode = 1
        Else
            Call Auto1Move
            If gCurrentMode = 10 Then
                Exit Sub
            End If
            If A2FindStranded Then  '2
                gCurrentMode = 16
                Timer2.Interval = playSpeed * 14
                Timer2.Enabled = True
            Else
                gCurrentMode = 14
            End If
        End If
    End If
End Sub

    ' Count all available ally units in range of victim.
    ' Don't include agressor (country launching attack)
Private Function countUnitsInrange(agressor As Integer, Victim As Integer, whoAlly As Integer) As Integer
    Dim cntr As Integer, rtrn As Integer, tmp As Integer
    Dim tst1 As Boolean, tst2 As Boolean
    Dim tmpCtry As Integer
    
    countUnitsInrange = 0
    For cntr = 1 To 7
        If CountryID(Victim).neighbour(cntr) = 0 Then
            Exit For
        ElseIf Not enemyAbout(CountryID(Victim).neighbour(cntr), whoAlly, Victim) _
            And CountryID(Victim).neighbour(cntr) <> agressor Then
            tmpCtry = CountryID(Victim).neighbour(cntr)
            tst1 = (OwnContinent(ContinentOfCtry(tmpCtry, gPlayerTurn), whoAlly))
            tst2 = (ContinentOfCtry(tmpCtry, gPlayerTurn) <> ContinentOfCtry(Victim, gPlayerTurn))
            If tst1 Then        ' I own continent
                tmp = gCtryScore(tmpCtry) - (getAvg / (playerDefence(gPlayerTurn) * 2))
                If tmp > 0 Then
                    countUnitsInrange = countUnitsInrange + tmp
                End If
            ElseIf tst2 Then    ' From different continent
                tmp = gCtryScore(tmpCtry) - (getAvg / (playerDefence(gPlayerTurn) * 2))
                If tmp > 0 Then
                    countUnitsInrange = countUnitsInrange + tmp
                End If
            End If
        End If
    Next cntr
End Function

    'Counts neighbours without enemies
Private Function countAllies(Victim As Integer, Who As Integer, ctry As Integer) As Integer
    Dim cntr As Integer, rtrn As Integer
    
    countAllies = 0
    For cntr = 1 To 7
        If CountryID(Victim).neighbour(cntr) = 0 Then
            Exit For
        ElseIf Not enemyAbout(CountryID(Victim).neighbour(cntr), gPlayerTurn, Victim) Then
            countAllies = countAllies + 1
        End If
    Next cntr
End Function

    'find a victim
Private Sub AutoMode20()
    Dim cntr1 As Integer
    Dim ctry As Integer
    Dim HoldCtry As Integer
    
    For cntr1 = 1 To 49
        ctry = Int(GenRandom4 * 7 + 1)
        If CountryID(AutoCountry).neighbour(ctry) > 0 Then
            If gCountryOwner(CountryID(AutoCountry).neighbour(ctry)) <> gPlayerTurn Then
                HoldCtry = AutoCountry
                AutoCountry = CountryID(HoldCtry).neighbour(ctry)
                Call ClickMapNow
                AutoCountry = HoldCtry
                Exit Sub
            End If
        End If
    Next cntr1
    Call Auto1Move
    If gCurrentMode = 10 Then
        Exit Sub
    End If
    Timer2.Enabled = False
    gComputerPressed = True
    Call EndClicked
    gComputerPressed = False

End Sub

    'Returns true if picked a country to attack
Private Function A200inCont() As Boolean
    Dim cntr1 As Integer, nbrs As Integer
    Dim ctry As Integer, avg As Integer
    Dim MyPoints As Integer, HisPoints As Integer
    Dim HoldCtry As Integer
    Dim tst As Boolean, tst2 As Boolean, tst3 As Boolean
    Dim Highest As Integer, lowest As Integer, LowestInCont As Integer
    Dim hiCtry As Integer, nextEnemy As Integer, tmp As Integer
    
        'Can't attack from a country not owned
    If gCountryOwner(AutoCountry) <> gPlayerTurn Then
        A200inCont = False
        Exit Function
    End If
    
    avg = getAvg
    
    For cntr1 = 1 To 7              'Find number of neighbours
        If CountryID(AutoCountry).neighbour(cntr1) = 0 Then
            nbrs = cntr1
            Exit For
        End If
    Next cntr1
    
    nbrs = nbrs - 1
    lowest = 9999
    nextEnemy = 20
    For cntr1 = 1 To 40             'Find victim in cont
        ctry = Int(GenRandom4 * nbrs) + 1      'If 2 the same, random first ctry
        MyPoints = countUnitsInrange(AutoCountry, CountryID(AutoCountry).neighbour(ctry), gPlayerTurn) + gCtryScore(AutoCountry)
        HisPoints = gCtryScore(CountryID(AutoCountry).neighbour(ctry))
        
        tst = (CountryID(AutoCountry).neighbour(ctry) > 0) _
            And (gCountryOwner(CountryID(AutoCountry).neighbour(ctry)) <> gPlayerTurn) _
            And (ContinentOfCtry(AutoCountry) = ContinentOfCtry(CountryID(AutoCountry).neighbour(ctry)))
            
        tst2 = (MyPoints > HisPoints + 1) Or (MyPoints > (avg))
        
        If tst And tst2 Then
            If lowest >= gCtryScore(CountryID(AutoCountry).neighbour(ctry)) - gCtryScore(AutoCountry) / 2 + 1 Then
                tmp = CountEnemyNbrs(ContinentOfCtry(AutoCountry), CountryID(AutoCountry).neighbour(ctry))
                tmp = tmp + modifyDirection(AutoCountry, CountryID(AutoCountry).neighbour(ctry))
                If tmp <= nextEnemy Then
                    nextEnemy = tmp
                    lowest = gCtryScore(CountryID(AutoCountry).neighbour(ctry))
                    hiCtry = CountryID(AutoCountry).neighbour(ctry)
                End If
            End If
        End If
    Next cntr1
    
    If lowest < 9999 Then
        HoldCtry = AutoCountry
        AutoCountry = hiCtry
        Call ClickMapNow
        If gCurrentMode <> 20 Then
            AutoCountry = HoldCtry
            A200inCont = True
            Exit Function
        End If
    End If
    A200inCont = False
End Function

    'Returns (number of enemy inside Cont neigbouring ctry) * 2
    '+ 1 if any enemy outside Cont neigbouring ctry
Private Function EasyCountEnemyNbrs(Cont As Integer, ctry As Integer, Who As Integer) As Integer
    Dim cntr1 As Integer
    Dim tmp As Integer
    Dim tst As Boolean, tst1 As Boolean
    Dim chkCEN As Integer
    
    EasyCountEnemyNbrs = 0
    If ctry = 42 Then
        EasyCountEnemyNbrs = -2     'High priority for East Aust.
    End If
    
    If killHim(ctry) Then
        EasyCountEnemyNbrs = EasyCountEnemyNbrs - 10
    End If
    
    
    'Any neigbouring enemy outside of Cont, add 1.
    'If I own that cont, -2.
    If autoIsAGate(ctry) Then
        For cntr1 = 1 To 7
            If CountryID(ctry).neighbour(cntr1) = 0 Then
                Exit For
            End If
            tmp = CountryID(ctry).neighbour(cntr1)
            tst = (ContinentOfCtry(tmp) <> ContinentOfCtry(ctry)) And _
                Not OwnContinent(ContinentOfCtry(tmp), Who) ' (enemyInCont(ContinentOfCtry(tmp), tmp, Who))
            tst1 = OwnContinent(ContinentOfCtry(tmp), Who)
            'tst1 = (ContinentOfCtry(tmp) <> ContinentOfCtry(ctry)) And Not (enemyInCont(ContinentOfCtry(tmp), tmp, Who))
            If tst And tst1 Then
                EasyCountEnemyNbrs = EasyCountEnemyNbrs - 1
                Exit For
            Else
                EasyCountEnemyNbrs = EasyCountEnemyNbrs + 2
                Exit For
            End If
        Next cntr1
    End If
        'Number inside Cont *4
    For cntr1 = 1 To 7
        If CountryID(ctry).neighbour(cntr1) = 0 Then
            Exit Function
        End If
        tst = (ContinentOfCtry(CountryID(ctry).neighbour(cntr1)) = Cont) _
        And (gCountryOwner(CountryID(ctry).neighbour(cntr1)) <> Who)
        If tst Then
            EasyCountEnemyNbrs = EasyCountEnemyNbrs + 4
        End If
    Next cntr1
End Function

'Find a country in this continent to launch an attack from.
'Return 0 if none found.
'TODO: use A3 random countries and neigbours.
Private Function EasyAttackLaunchFrom(Cont As Integer, Who As Integer) As Integer
    Dim ctry As Integer
    Dim cntr1 As Integer
    Dim nbr As Integer
    Dim EnemyCount As Integer
    Dim EnemyCountLeast As Integer
    
    EasyAttackLaunchFrom = 0
    EnemyCountLeast = 9999
    
    For ctry = Continents(Cont - 1).FirstCountry To Continents(Cont - 1).LastCountry
        If gCountryOwner(ctry) <> Who Then
            EnemyCount = EasyCountEnemyNbrs(Cont, ctry, Who)
            For cntr1 = 1 To 7
                nbr = CountryID(ctry).neighbour(cntr1)
                If nbr = 0 Then
                    Exit For
                End If
                If gCountryOwner(nbr) = Who Then
                    If EnemyCount < EnemyCountLeast _
                    And gCtryScore(nbr) > gCtryScore(ctry) + 4 Then
                        EnemyCountLeast = EnemyCount
                        EasyAttackLaunchFrom = nbr
                    End If
                End If
            Next
        End If
    Next
End Function

'Reset Easy Attack markers.
Private Sub ResetEasyAttackConts()
    Dim i As Long
    
    For i = 1 To 6
        EasyAttackConts(i) = False
    Next
End Sub


'Look for an In-your-face easy attack.
Private Function EasyAttack() As Boolean
    Dim cntr1 As Long
    Dim Cont As Integer
    Dim ContScore As Integer
    Dim LaunchCountry As Integer
    
    EasyAttack = False
    
    'Find easy continent to take.
    For cntr1 = 0 To 5          'Look at all conts
        Cont = ContPriority(cntr1)
        
        'Make sure this Cont hasn't already
        'been proccessed. Stops hang.
        If Not EasyAttackConts(Cont) Then
            If Not OwnContinent(Cont, gPlayerTurn) Then
            
                'Do I have enough points to take this continent?
                ContScore = PointsAvailableInCont(Cont, gPlayerTurn, gPickedUpUnits)
                
                If ContScore > 0 Then
                    LaunchCountry = EasyAttackLaunchFrom(Cont, gPlayerTurn)
                    If LaunchCountry > 0 Then
                        AutoCountry = LaunchCountry
                        EasyAttackConts(Cont) = True
                        If A200inCont Then
                            EasyAttack = True
                            Exit For
                        ElseIf A200outOfCont Then
                            EasyAttack = True
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next cntr1
    
    IsEasyAttackOn = EasyAttack
End Function

'Return my points proportioned to enemy points within this cont.
'+ve is my advantage, -ve is enemy advantage, 0 is neutral.
'Cont is continent, Who is my player number, offset is extra points I can use.
Private Function PointsAvailableInCont(Cont As Integer, Who As Integer, Optional Offset As Integer = 0) As Integer
    Dim cntr1 As Integer, cntr2 As Integer
    Dim Highest As Integer, MyPoints As Integer, theirPoints As Integer
    Dim tst As Integer, tmp As Long
    
    On Error GoTo ErrHand
    
    PointsAvailableInCont = Offset
    
    For cntr2 = Continents(Cont - 1).FirstCountry To Continents(Cont - 1).LastCountry
        If gCountryOwner(cntr2) = Who Then
            PointsAvailableInCont = PointsAvailableInCont + gCtryScore(cntr2) - 1
        Else
            PointsAvailableInCont = PointsAvailableInCont - gCtryScore(cntr2)
        End If
    Next cntr2
    Exit Function
ErrHand:
    Resume Next
End Function

    ' Stop attacking front from ending up where it is not needed by
    ' lowering enemy count if path is being looked at.
Private Function modifyDirection(cFrom As Integer, cTo As Integer) As Integer

        ' >> In Asia
    If cFrom = 32 And cTo = 28 Then
            ' Own Europe and Africa
        If OwnContinent(3, gPlayerTurn) _
        And OwnContinent(4, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 32 And cTo = 33 Then
            ' Own Australia
        If OwnContinent(6, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 27 And cTo = 29 Then
            ' Own Australia
        If OwnContinent(6, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 27 And cTo = 28 Then
            ' Own NA
        If OwnContinent(1, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 32 And cTo = 31 Then
            ' Own NA
        If OwnContinent(1, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 28 And cTo = 32 Then
            ' Own NA
        If OwnContinent(1, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 30 And cTo = 33 Then
            ' Own NA
        If OwnContinent(1, gPlayerTurn) Then
            modifyDirection = -9
        End If
    ElseIf cFrom = 33 And cTo = 34 Then
            ' Own NA
        If OwnContinent(1, gPlayerTurn) Then
            modifyDirection = -9
        End If
        
        ' >> In N America
    ElseIf cFrom = 1 And cTo = 2 Then
            ' Own Europe
        If OwnContinent(3, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 1 And cTo = 4 Then
            ' Own SA
        If OwnContinent(2, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 4 And cTo = 7 Then
            ' Own SA
        If OwnContinent(2, gPlayerTurn) Then
            modifyDirection = -5
        End If
    ElseIf cFrom = 4 And cTo = 5 Then
            ' Own Europe
        If OwnContinent(3, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 6 And cTo = 8 Then
            ' Own SA
        If OwnContinent(2, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 6 And cTo = 5 Then
            ' Own Asia
        If OwnContinent(5, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 8 And cTo = 6 Then
            ' Own Europe
        If OwnContinent(3, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 9 And cTo = 8 Then
            ' Own Asia or Europe
        If OwnContinent(5, gPlayerTurn) Or OwnContinent(3, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 8 And cTo = 7 Then
            ' Own Asia
        If OwnContinent(5, gPlayerTurn) Then
            modifyDirection = -1
        End If
    ElseIf cFrom = 3 And cTo = 2 Then
            ' Own Asia
        If OwnContinent(5, gPlayerTurn) Then
            modifyDirection = -5
        End If
    ElseIf cFrom = 2 And cTo = 3 Then
            ' Own Europe
        If OwnContinent(3, gPlayerTurn) Then
            modifyDirection = -1
        End If
    
        ' >> In Africa
    ElseIf cFrom = 21 And cTo = 24 Then
            ' Always go this way
        modifyDirection = -3

    ElseIf cFrom = 22 And cTo = 23 Then
            ' DO NOT own Europe or SA
        If Not OwnContinent(2, gPlayerTurn) Or Not OwnContinent(3, gPlayerTurn) Then
            modifyDirection = -9
        End If
    
    ElseIf cFrom = 23 And cTo = 26 Then
            ' Always go this way
        modifyDirection = -1
    
        ' >> In Europe
    ElseIf cFrom = 20 And cTo = 16 Then
            ' Own NA
        If OwnContinent(1, gPlayerTurn) Then
            modifyDirection = -3
        End If
    ElseIf cFrom = 18 And cTo = 15 Then
            ' Own NA
        If OwnContinent(1, gPlayerTurn) Then
            modifyDirection = -3
        End If

    End If
End Function

    'Return true if advantage in killing ctry
Private Function killHim(ctry As Integer) As Boolean
    Dim ctryLeft As Integer, cntr As Integer
    
    killHim = False
    ctryLeft = CountCountriesHeld(gCountryOwner(ctry))
    
    If gCountryOwner(ctry) = gPlayerTurn Then Exit Function
    
        'Is his cards worth it?
    If chkCardsVulture.Value = vbChecked Then
        If ((gPlayerID(gCountryOwner(ctry)).card(1) > 0) _
        Or (gPlayerID(gCountryOwner(ctry)).card(0) = 4)) _
        And (ctryLeft < 4) _
        And ((chkMsnMissionsOn.Value = vbUnchecked) _
        Or (chkMsnArmyWipeout.Value = vbUnchecked) _
        Or (chkMsnMustComplete.Value = vbChecked)) Then
            killHim = True
            Exit Function
        End If
    End If
    
    If chkMsnMissionsOn.Value = vbChecked _
    And (ctryLeft < 6) Then
        If gMissions(gPlayerID(gPlayerTurn).mission).TargetArmy = gCountryOwner(ctry) Then
            killHim = True
            Exit Function
        End If
    End If
End Function

    'Return True if OK to kill other players
Public Function okToKill() As Boolean
    If (chkMsnMissionsOn.Value = vbChecked And chkMsnMustComplete.Value = vbUnchecked) _
    Or chkCardsVulture.Value = vbUnchecked Then
        okToKill = False
    Else
        okToKill = True
    End If
End Function

    'Return True if enemy is mission target for gPlayerTurn
Public Function isTarget(pEnemy As Integer) As Boolean
    If chkMsnMissionsOn.Value = vbChecked _
    And gMissions(gPlayerID(gPlayerTurn).mission).TargetArmy = pEnemy Then
        isTarget = True
    Else
        isTarget = False
    End If
End Function

    'Returns (number of enemy inside Cont neigbouring ctry) * 2
    '+ 1 if any enemy outside Cont neigbouring ctry
Private Function CountEnemyNbrs(Cont As Integer, ctry As Integer) As Integer
    Dim cntr1 As Integer
    Dim tmp As Integer
    Dim tst As Boolean, tst1 As Boolean
    Dim chkCEN As Integer
    
    CountEnemyNbrs = 0
    If ctry = 42 Then
        CountEnemyNbrs = -2     'High priority for East Aust.
    End If
    
    If killHim(ctry) Then
        CountEnemyNbrs = CountEnemyNbrs - 10
    End If
    
    
        'Any neigbouring enemy outside of Cont, add 1.
    If autoIsAGate(ctry) Then
        For cntr1 = 1 To 7
            If CountryID(ctry).neighbour(cntr1) = 0 Then
                Exit For
            End If
            tmp = CountryID(ctry).neighbour(cntr1)
            tst = (ContinentOfCtry(tmp) <> ContinentOfCtry(AutoCountry)) And _
                Not OwnContinent(ContinentOfCtry(tmp), gPlayerTurn) ' (enemyInCont(ContinentOfCtry(tmp), tmp, gPlayerTurn))
            'tst1 = ownContinent(ContinentOfCtry(tmp), gPlayerTurn)
            'tst1 = (ContinentOfCtry(tmp) <> ContinentOfCtry(AutoCountry)) And Not (enemyInCont(ContinentOfCtry(tmp), tmp, gPlayerTurn))
            If tst Then
                CountEnemyNbrs = CountEnemyNbrs + 2
                Exit For
            End If
        Next cntr1
    End If
        'Number inside Cont *4
    For cntr1 = 1 To 7
        If CountryID(ctry).neighbour(cntr1) = 0 Then
            Exit Function
        End If
        tst = (ContinentOfCtry(CountryID(ctry).neighbour(cntr1)) = _
            ContinentOfCtry(AutoCountry)) And _
            (gCountryOwner(CountryID(ctry).neighbour(cntr1)) <> gPlayerTurn)
        If tst Then
            CountEnemyNbrs = CountEnemyNbrs + 4
        End If
    Next cntr1
End Function

    'Returns true if picked a victim out side continent
Private Function A200outOfCont() As Boolean
    Dim cntr1 As Integer, nbrs As Integer
    Dim ctry As Integer, avg As Integer
    Dim MyPoints As Integer, HisPoints As Integer
    Dim HoldCtry As Integer
    Dim tst As Boolean, tst2 As Boolean, tst3 As Boolean
    Dim Highest As Integer, lowest As Integer, LowestInCont As Integer
    Dim hiCtry As Integer
    
    avg = getAvg
    
        'Can't launch an attack from another player's territory
    If gCountryOwner(AutoCountry) <> gPlayerTurn Then
        A200outOfCont = False
        Exit Function
    End If
    
    For cntr1 = 1 To 7              'Find number of neighbours
        If CountryID(AutoCountry).neighbour(cntr1) = 0 Then
            nbrs = cntr1
            Exit For
        End If
    Next cntr1
    nbrs = nbrs - 1
    Highest = 0
    lowest = 9999
    
    For cntr1 = 1 To 30             'Find a victim out of continent
        ctry = Int(GenRandom4 * nbrs) + 1
        
        MyPoints = countUnitsInrange(AutoCountry, CountryID(AutoCountry).neighbour(ctry), gPlayerTurn) + gCtryScore(AutoCountry)
        HisPoints = gCtryScore(CountryID(AutoCountry).neighbour(ctry))
        
        tst = (CountryID(AutoCountry).neighbour(ctry) > 0) _
            And (gCountryOwner(CountryID(AutoCountry).neighbour(ctry)) <> gPlayerTurn) _
            And (MyPoints > HisPoints) And (MyPoints > (avg / playerDefence(gPlayerTurn)))
            
        If prefercont <> 0 Then
            tst2 = (prefercont = ContinentOfCtry(CountryID(AutoCountry).neighbour(ctry), gPlayerTurn))
        Else
            tst2 = True
        End If
        
        tst3 = (killHim(CountryID(AutoCountry).neighbour(ctry))) _
            And (CountryID(AutoCountry).neighbour(ctry) > 0) _
            And (MyPoints > HisPoints)
        
        
        If tst And tst2 Then                     'and tst2?
            If Highest <= BigInCont(ContinentOfCtry(CountryID(AutoCountry).neighbour(ctry), gPlayerTurn), gPlayerTurn, gCtryScore(AutoCountry)) Then
                Highest = BigInCont(ContinentOfCtry(CountryID(AutoCountry).neighbour(ctry), gPlayerTurn), gPlayerTurn, gCtryScore(AutoCountry))
                If lowest > gCtryScore(CountryID(AutoCountry).neighbour(ctry)) Then
                    lowest = gCtryScore(CountryID(AutoCountry).neighbour(ctry))
                    hiCtry = CountryID(AutoCountry).neighbour(ctry)
                End If
            End If
        ElseIf tst3 Then
            Highest = 9999
            lowest = gCtryScore(CountryID(AutoCountry).neighbour(ctry))
            hiCtry = CountryID(AutoCountry).neighbour(ctry)
            Exit For
        End If
    Next cntr1
    
    If Highest > 0 Then
        HoldCtry = AutoCountry
        AutoCountry = hiCtry 'countryID(holdCtry).neighbour(ctry)
        Call ClickMapNow
        AutoCountry = HoldCtry
        If gCurrentMode = 21 Then
            A200outOfCont = True
            Exit Function
        End If
    End If
    
    A200outOfCont = False
    
End Function

    'Smart player's attack who a better approach
Private Sub Auto200Mode20()
    With A2opportunity
    IsEasyAttackOn = False
    If .IsActive Then                     'Follow pre set path
        If .Path(.pathPointer) > 0 Then         'Still going
            If gCountryOwner(.Path(.pathPointer)) = gPlayerTurn Then
                .IsActive = False
                .pathPointer = 0
                .Path(0) = 0
            Else
                AutoCountry = .Path(.pathPointer)
                Call ClickMapNow
                If gCurrentMode = 20 Then
                    'Shit... it failed!
                    .IsActive = False
                    .pathPointer = 0
                    .Path(0) = 0
                    'If EasyAttack Then
                        'prefercont = 0
                        'Exit Sub
                    'End If
                End If
            End If
            Exit Sub
        Else                            'End of path
            .IsActive = False
            .pathPointer = 0
            .Path(0) = 0
            If .StopWhenFinished Then
                .StopWhenFinished = False
                'Look for an In-your-face easy attack.
                If EasyAttack Then
                    prefercont = 0
                    Exit Sub
                Else
                    Call CheckForMove
                End If
                Exit Sub
            End If
        End If
    End If
    End With
    
    If AtkForCard.On Then
        AutoCountry = AtkForCard.To
        Call ClickMapNow
        If gCurrentMode = 20 Then
            'Look for an In-your-face easy attack.
            If EasyAttack Then
                prefercont = 0
                Exit Sub
            Else
                gCurrentMode = 16
            End If
        End If
        Exit Sub
    End If

    Call A3.findOpportunity(gPickedUpUnits)     'Try to find a path
    If A2opportunity.IsActive Then        'Follow pre set path if one found
        AutoCountry = A2opportunity.Path(A2opportunity.pathPointer)
        Call ClickMapNow
        Exit Sub
    ElseIf prefercont = 0 Then
        If A200inCont Then
            Exit Sub
        ElseIf A200outOfCont Then
            Exit Sub
        End If
    Else
        If A200outOfCont Then
            prefercont = 0
            Exit Sub
        ElseIf A200inCont Then
            prefercont = 0
            Exit Sub
        End If
        prefercont = 0
    End If
    
    'Look for an In-your-face easy attack.
    If EasyAttack Then
        prefercont = 0
        Exit Sub
    End If
    
    'Look for a thorn-in-side advantage.
    If A3.A2FindThornInSide(gPlayerTurn, CSng(getAvg / playerDefence(gPlayerTurn)) + 1) Then
        Exit Sub
    ElseIf A200AttackAny Then
        prefercont = 0
        Exit Sub
    End If
    
    'Look for ANOTHER In-your-face easy attack.
    If EasyAttack Then
        prefercont = 0
        Exit Sub
    End If
    
    'Call A2move                'Later
    'Exit Sub
    '-------------------------------------
    Call CheckForMove
End Sub

Private Sub CheckForMove()
    Call Auto1Move
    If gCurrentMode = 10 Then
        Exit Sub
    End If
    If A2FindStranded Then  '3
        gCurrentMode = 16
        Timer2.Interval = playSpeed * 14
        Timer2.Enabled = True
    Else
        gCurrentMode = 14
    End If
End Sub

    'Find stranded units
Private Function A2FindStranded() As Boolean
    Dim i As Integer
    Dim ctry As Integer
    Dim randList(42) As Integer
    
    Call A3.fillRandList(randList)
    If gPlayerID(gPlayerTurn).playerWho < 2 Then Exit Function
    
    For i = 1 To 42
        ctry = randList(i)
        If gCountryOwner(ctry) = gPlayerTurn _
        And Not autoIsAGate(ctry, gPlayerTurn) _
        And Not enemyAbout(ctry, gPlayerTurn, 0) _
        And gCtryScore(ctry) > 1 Then
            If isMoveOK(ctry) Then
                gTargetCtry = ctry
                A2FindStranded = True
                Exit Function
            End If
        End If
    Next
    A2FindStranded = False
    Exit Function
End Function

    'If not attacked, try to attack anyone to get a card.
Private Function A200AttackAny() As Boolean
    Dim cntr As Integer, ctry As Integer, cntr2 As Integer
    Dim lowest As Integer, attackWho As Integer, nbrNmbr As Integer
    Dim bigstDif As Integer
    Dim atkFrom As Integer
    Dim tmp As Long
    
    With AtkForCard
    If .On Then
        A200AttackAny = True
        AutoCountry = attackWho
        gComputerPressed = True
        transferNmbr = 9999
        Call ClickMapNow
        gComputerPressed = False
        Exit Function
    End If
        .from = 0
        .To = 0
        .On = False
    End With
    
    A200AttackAny = False
    If (GetCardMode = 0) Or (boolIssueCard) Then
        Exit Function
    End If
    
    lowest = 9999
    bigstDif = 0
    For cntr = 1 To 500
        ctry = Int(GenRandom4 * 42 + 1)
        If (gCountryOwner(ctry) = gPlayerTurn) And (gCtryScore(ctry) > 2) Then
            For cntr2 = 1 To 7
                nbrNmbr = CountryID(ctry).neighbour(cntr2)
                If nbrNmbr = 0 Then Exit For
                
                If gCountryOwner(nbrNmbr) <> gPlayerTurn Then
                    If (gCtryScore(nbrNmbr) < lowest) _
                    And (gCtryScore(ctry) - gCtryScore(nbrNmbr) >= bigstDif) Then
                        lowest = gCtryScore(nbrNmbr)
                        attackWho = nbrNmbr
                        bigstDif = gCtryScore(ctry) - gCtryScore(nbrNmbr)
                        atkFrom = ctry
                    ElseIf (gCtryScore(nbrNmbr) = lowest) _
                    And ((gCtryScore(ctry) - gCtryScore(nbrNmbr)) >= bigstDif) Then
                        If (GenRandom4 * 2) > 1 Then
                            attackWho = nbrNmbr
                            bigstDif = gCtryScore(ctry) - gCtryScore(nbrNmbr)
                            atkFrom = ctry
                        End If
                    End If
                End If
            Next cntr2
        End If
    Next cntr
    
    If atkFrom = 0 Then Exit Function
    
        'Attack with 10% tollerance
    If (gCtryScore(atkFrom)) > (lowest + (lowest / 10) + 2) Then
        AutoCountry = attackWho
        gComputerPressed = True
        transferNmbr = 9999
        Call ClickMapNow
        gComputerPressed = False
        A200AttackAny = True
        
        With AtkForCard
            .from = atkFrom
            .To = attackWho
            .On = True
        End With
        
        Exit Function
    End If
End Function

    'computer presses ATTACK if enemy on border     mode 1 -> 20
Private Sub Auto1Attack()
    Dim cntr1 As Integer
    Dim ctry As Integer
    
    gComputerPressed = True
    Call AttackClicked
    gComputerPressed = False
    If gCurrentMode <> 20 Then
        gCurrentMode = 14
    End If
End Sub

    'Click currently selected country (AutoCountry)
Private Sub AutoMode21()
    Call ClickMapNow
    Call AutoFindAttacker
End Sub

    'Start AutoPlayer1
Private Sub AutoPlayer1()
    Dim ctry As Integer
    Call Auto1PutCtry
End Sub

    'Find a suitable country to put new units
Private Sub Auto1PutCtry()
    Dim cntr As Integer, ctry As Integer, cntr2 As Integer
    
    For cntr = 1 To 500
        ctry = Int(GenRandom4 * 42) + 1
        If gCountryOwner(ctry) = gPlayerTurn Then
            If enemyAbout(ctry, gPlayerTurn, 0) Then
                AutoCountry = ctry
                gComputerPressed = True
                transferNmbr = 9999
                Call ClickMapNow
                gComputerPressed = False
                Exit Sub
            End If
        End If
    Next cntr
End Sub

    'Return true if countryNmbr is on a border with player <> playerNumber (country)
    'and <> targetNmbr (0 if not required)
Private Function enemyAbout(CountryNumber As Integer, PlayerNumber As Integer, targetNmbr As Integer) As Boolean
    Dim cntr1 As Integer
    Dim nbrNmbr As Integer
    
    If CountryNumber = 0 Then
        enemyAbout = False
        Exit Function
    End If
    
    For cntr1 = 1 To 7
        nbrNmbr = CountryID(CountryNumber).neighbour(cntr1)
        If nbrNmbr = 0 Then
            enemyAbout = False
            Exit Function
        ElseIf gCountryOwner(nbrNmbr) <> PlayerNumber Then
            If targetNmbr <> CountryID(CountryNumber).neighbour(cntr1) Then
                enemyAbout = True
                Exit Function
            End If
        End If
    Next cntr1
    enemyAbout = False
End Function

    'Check cards and turn in if a set exists
    'TODO: Refactor, comment, document. Similar job to FindHumanCards().
Private Sub AutoCheckCards()
    Dim nbrOfCrds As Integer, cntr As Integer, tmpry As Integer
    Dim tst As Boolean
    Dim tmp2 As Long
    Dim Danger As Boolean
    Dim CardValue As Integer
    Dim dvCode As Long
    
    Call CardOutOfHand(gPlayerTurn)
    Danger = inDanger
    CardValue = gCurrentCardValue
    If (GetCardMode = 2) _
    And (gPlayerID(gPlayerTurn).card(4) = 0) _
    And (Not Danger) _
    And (CardValue < 20) _
    And (CardValue < gMaxCardValue) _
    And (gPlayerID(gPlayerTurn).playerWho <> 1) Then
        gCurrentMode = 2
        Exit Sub
    End If
    
    'Use dice to do the sorting.
    For cntr = 0 To 9
        gDiceArray(cntr) = 0
    Next cntr
    
    
    For cntr = 0 To 7
        If gPlayerID(gPlayerTurn).card(cntr) = 0 Then
            Exit For
        ElseIf gPlayerID(gPlayerTurn).card(cntr) = 3 Then
            gDiceArray(0) = 3
        ElseIf gPlayerID(gPlayerTurn).card(cntr) = 2 Then
            gDiceArray(1) = 2
        ElseIf gPlayerID(gPlayerTurn).card(cntr) = 1 Then
            gDiceArray(2) = 1
        ElseIf gPlayerID(gPlayerTurn).card(cntr) = 4 Then
            gDiceArray(3) = 4
        End If
    Next cntr
    
    nbrOfCrds = 0
    tst = False
    For cntr = 0 To 2
        If gDiceArray(cntr) <> 0 Then
            nbrOfCrds = nbrOfCrds + 1
        End If
    Next cntr
    If nbrOfCrds = 3 Then
        tst = True
    ElseIf ((gDiceArray(3) = 4) And (GetCardMode = 1)) Then 'Or ((gDiceArray(3) = 4) And (Danger)) Then
        If nbrOfCrds = 2 Then
            tst = True
            Call SortDice(1, 4)
        End If
    End If
    
    If tst Then                 'All different
        gDiceArray(3) = 0
        gCurrentMode = 8
        gDiceArray(11) = 1
        Call AutoCrdInfo
        Exit Sub
    End If

    For cntr = 0 To 9
        gDiceArray(cntr) = 0
    Next cntr
    nbrOfCrds = 0
    
    For cntr = 0 To 9
        If gPlayerID(gPlayerTurn).card(cntr) > 0 Then
            nbrOfCrds = nbrOfCrds + 1
        End If
        gDiceArray(cntr) = gPlayerID(gPlayerTurn).card(cntr)
    Next cntr
    
    If (nbrOfCrds < 3) Then
        gCurrentMode = 2
        Exit Sub
    End If
    
    Call SortDice(1, nbrOfCrds)
    
        'Check all same, no jokers
    For cntr = 1 To nbrOfCrds - 2
        tst = ((gDiceArray(cntr - 1) = gDiceArray(cntr)) And (gDiceArray(cntr) = gDiceArray(cntr + 1)))
            
        If tst Then                 'All the same (no joker)
            gCurrentMode = 8
            gDiceArray(11) = cntr
            Call AutoCrdInfo
            Exit Sub
        End If
    Next cntr
    
    If ((gDiceArray(0) = 4) And (GetCardMode = 2)) _
    Or ((gDiceArray(0) = 4) And (Danger)) Then
        For cntr = 2 To nbrOfCrds - 1
            If gDiceArray(cntr - 1) = gDiceArray(cntr) Then
                gDiceArray(1) = gDiceArray(cntr - 1)
                gDiceArray(2) = gDiceArray(1)
                gCurrentMode = 8
                gDiceArray(11) = 1
                Call AutoCrdInfo
                Exit Sub
            End If
        Next cntr
    End If
    
    tmpry = nbrOfCrds
    For cntr = 1 To nbrOfCrds - 1
        If gDiceArray(cntr - 1) = gDiceArray(cntr) Then
            gDiceArray(cntr - 1) = 0
            tmpry = tmpry - 1
        End If
    Next cntr
    If tmpry < 3 Then       'Bodgy check for 2 jokers
        For cntr = 0 To 9
            gDiceArray(cntr) = 0
        Next cntr
        nbrOfCrds = 0
        
        For cntr = 0 To 9
            If gPlayerID(gPlayerTurn).card(cntr) > 0 Then
                nbrOfCrds = nbrOfCrds + 1
            End If
            gDiceArray(cntr) = gPlayerID(gPlayerTurn).card(cntr)
        Next cntr
        Call SortDice(1, nbrOfCrds)
        If ((gDiceArray(0) = 4) And (gDiceArray(1) = 4) And (GetCardMode = 2)) _
        Or ((gDiceArray(0) = 4) And (gDiceArray(1) = 4) And (Danger)) Then
            gCurrentMode = 8
            gDiceArray(11) = 1
            Call AutoCrdInfo
            Exit Sub
        End If
        Exit Sub
    ElseIf tmpry > 3 Then
    End If

    Call SortDice(1, nbrOfCrds)     'All different (including joker)
        tst = (gDiceArray(0) <> gDiceArray(1)) _
                And (gDiceArray(1) <> gDiceArray(2)) _
                And (gDiceArray(2) <> gDiceArray(0))
        If tst Then
            gCurrentMode = 8
            gDiceArray(11) = 1
            Call AutoCrdInfo
            Exit Sub
        End If
    If gDiceArray(0) = 4 Then             'Joker
        gCurrentMode = 8
        gDiceArray(11) = 1
        Call AutoCrdInfo
        Exit Sub
    End If
End Sub

    'True if card node <> increasing and joker with 2 same
    'or 2 jokers if in danger: NOT currently used or tested
Private Function isJoker(Danger As Boolean) As Boolean
    Dim cntr1 As Long
    Dim nbrOfCrds As Integer
    Dim cntr As Long
    Dim tst As Boolean
    
    For cntr = 0 To 9
        If gPlayerID(gPlayerTurn).card(cntr) > 0 Then
            nbrOfCrds = nbrOfCrds + 1
        End If
        gDiceArray(cntr) = gPlayerID(gPlayerTurn).card(cntr)
    Next cntr
    Call SortDice(1, nbrOfCrds)
    
    For cntr = 1 To nbrOfCrds - 2
        tst = ((gDiceArray(cntr - 1) = gDiceArray(cntr)) And (gDiceArray(cntr) = gDiceArray(cntr + 1))) _
            Or ((gDiceArray(cntr - 1) = 4) And (GetCardMode = 2)) _
            Or ((gDiceArray(cntr - 1) = 4) And (Danger))
        If tst Then                 'All the same (no joker)
            gCurrentMode = 8
            gDiceArray(11) = cntr
            Call AutoCrdInfo
            Exit Function
        End If
    Next cntr
    
    gCurrentMode = 2
End Function


    'Return true if in danger of being wiped out in the next go
Private Function inDanger() As Boolean
    Dim avg As Integer
    Dim cntr1 As Integer
    Dim countValue As Integer

    countValue = 0
    For cntr1 = 1 To 42
        If gCountryOwner(cntr1) = gPlayerTurn Then
            countValue = countValue + gCtryScore(cntr1)
        End If
    Next cntr1
    
    If GetCardMode = 2 Then            'Get average value
        avg = (Int(totalScore / 8.4 + (gCurrentCardValue / 4)))
    Else
        avg = (Int(totalScore / 8.4))
    End If
    If countValue <= avg Then
        inDanger = True
    Else
        inDanger = False
    End If

End Function

    'Return total of all scores
Private Function totalScore() As Integer
    Dim cntr As Integer, rtrn As Integer
    
    rtrn = 0
    For cntr = 1 To 42
        rtrn = rtrn + gCtryScore(cntr)
    Next cntr
    totalScore = rtrn
End Function

    'Tell humans what is happening
Private Sub AutoCrdInfo()

    InfoBoxPrint 0
    InfoBoxPrnCR 9, gPlayerTurn
    InfoBoxPrint 7
    InfoBoxPrint 5                           'bold
    InfoBoxPrnCR 1, 130                      'checking cards
    InfoBoxPrint 6                           'normal
    
    With AtkForCard
        .from = 0
        .To = 0
        .On = False
    End With
    
End Sub

    'Turn in 3 cards with value of gDiceArray(atPosition)
Private Sub AutoTurn3In(atPosition As Integer)
    Dim cntr1 As Integer, cntr2 As Integer
    Dim tst As Integer, tmpry As Integer
    
    For cntr1 = atPosition To atPosition + 2
        For cntr2 = 1 To 10
            tst = (gDiceArray(cntr1 - 1) = gPlayerID(gPlayerTurn).card(cntr2 - 1)) _
                    And (gPlayerID(gPlayerTurn).pickedCards(cntr2 - 1)) = True
            If tst Then
                gPlayerID(gPlayerTurn).pickedCards(cntr2 - 1) = False
                tmpry = gCurrentMode
                gCurrentMode = 5
                Call DrawBigCard(gDiceArray(cntr1 - 1), cntr2, False)
                gCurrentMode = tmpry
                Exit For
            End If
        Next cntr2
    Next cntr1
    gCurrentMode = 7
    Timer2.Interval = playSpeed * 20
End Sub

    'Return current defence value
Private Function getAvg() As Integer
    If IsEasyAttackOn Then
        getAvg = 18
    ElseIf GetCardMode = 2 Then            'Get average value
        getAvg = (Int(averageScore + (gCurrentCardValue / 2)) + 1) + 6
    Else
        getAvg = (Int(averageScore) + 18)
    End If
End Function

    'Put defence units then find country for rest.
Private Sub AutoPlayer2()           'Put new units (mode2)
    Dim cntr1 As Integer, cntr2 As Integer, enemyPoints As Integer, Cont As Integer
    Dim MyPoints As Integer
    Dim tst As Boolean
    Dim prospects(42) As Integer            'Enemy points on gates
    Dim Pointer As Integer                  'pointer for prospects()
    Dim highestGate As Integer              'Highest enemy on this ctry
    Dim highestPoints As Integer            'Highest points so far
    Dim avg As Integer                      'average points
    Dim nbr As Integer
    Dim TempGate(5) As Integer
    
    With A2opportunity
    ' Search harder if I have the points.
    If gPickedUpUnits + gPlayerValue > 20 + CInt(GenRandom4 * 5) Then
        cntr1 = CInt(gA3MaxSearchDepth * 1.3)
    ElseIf gPickedUpUnits + gPlayerValue > 15 + CInt(GenRandom4 * 5) Then
        cntr1 = CInt(gA3MaxSearchDepth * 1.2)
    Else
        cntr1 = CInt(gA3MaxSearchDepth * 1.1)    '13
    End If
    If A3.findOpportunity(gPickedUpUnits + gPlayerValue, CLng(cntr1)) Then  'Try to find a path
        cntr1 = 0
        Do
            cntr2 = Int(GenRandom4 * 5) + 1           'From 1 to 6
            nbr = CountryID(.Path(.pathPointer)).neighbour(cntr2)
            If nbr > 0 Then
                If gCountryOwner(nbr) = gPlayerTurn Then
                    AutoCountry = nbr
                    gComputerPressed = True
                    transferNmbr = (9999)       '10% of value to highest threat
                    Call ClickMapNow
                    gComputerPressed = False
                    Exit Sub
                End If
            End If
            cntr1 = cntr1 + 1
        Loop While cntr1 < 32
        
        .IsActive = False                     'Failed to find a country
        .pathPointer = 0                    'for units next to start of path
        .Path(0) = 0
        .unitsRequired = 0
    End If
    End With
    
    
    
    Pointer = 1                             'Clear prospects
    highestGate = 0
    highestPoints = 0
    For cntr1 = 1 To 42
        prospects(cntr1) = 0
    Next cntr1

    avg = getAvg
    For cntr1 = 1 To 6
    Cont = ContPriority(cntr1 - 1)
   
        If OwnContinent(Cont, gPlayerTurn) Then
            For cntr2 = 1 To 5
                Call SetTempGatesForCont(Cont, TempGate, gPlayerTurn)
                enemyPoints = AutoPointsOnGate(Cont, cntr2, TempGate)    'Get highest points on gate
                If TempGate(cntr2) = 0 Then
                    Exit For
                End If
                
                MyPoints = gCtryScore(TempGate(cntr2))
                enemyPoints = enemyPoints - MyPoints
                
                'Dont worry about gates between held conts
                If MyPoints < (getAvg + 1) Then
                    tst = (ContOfCtryNbr(TempGate(cntr2)))
                    If Not tst Then
                        'Don't own neibouring continent
                        enemyPoints = Int(averageScore + (gCurrentCardValue / playerDefence(gPlayerTurn))) + 1
                    Else
                        enemyPoints = 0
                    End If
                ElseIf enemyPoints < 1 Then
                    enemyPoints = 0
                End If
                            
                If enemyPoints <> 0 Then
                    prospects(Pointer) = TempGate(cntr2)
                    Pointer = Pointer + 1
                    If highestGate = 0 Then
                        highestGate = TempGate(cntr2)
                        highestPoints = enemyPoints
                    ElseIf highestPoints < enemyPoints Then
                        highestGate = TempGate(cntr2)
                        highestPoints = enemyPoints
                    End If
                End If
            Next cntr2
        End If
    Next cntr1
    
    If highestGate <> 0 Then
        MyPoints = Int(gPlayerValue / 2)
        AutoCountry = highestGate
                
        gComputerPressed = True
        transferNmbr = (MyPoints / 5)       '10% of value to highest threat
        MyPoints = MyPoints - transferNmbr
        Call ClickMapNow
        gComputerPressed = False
        
        Pointer = Pointer - 1
        If MyPoints < Pointer Then
            Pointer = MyPoints
        End If
        
                'Bodge bodge bodge!!!
        If (Pointer = 0) Then
            Pointer = 1
            transferNmbr = 1
        Else
            transferNmbr = Int(MyPoints / Pointer)
        End If
        
        For cntr1 = 1 To Pointer
            AutoCountry = prospects(cntr1)
            gComputerPressed = True
            Call ClickMapNow
            gComputerPressed = False
        Next cntr1
        Call AutoPlace2
        Exit Sub
    End If
    Call AutoPlace2
End Sub

    'select ctry to place new units depending on player
Private Sub AutoPlace2()
    If gPlayerID(gPlayerTurn).playerWho = 2 Then
        If AutoPut2 Then
            Exit Sub
        End If
    End If
    Call AutoPlayer1
End Sub

    'Find best spot to put new attack units
Private Function AutoPut2() As Boolean
    Dim prospects(6) As Integer
    Dim Pointer As Integer
    Dim tmp As Integer
    Dim cntr1 As Integer
    Dim cntr2 As Integer
    Dim Cont As Integer
    Dim HstCont As Integer
    Dim Highest As Integer
    Dim tst As Boolean
    Dim tst1 As Boolean
    Dim tst2 As Boolean
    
    Highest = 0
    Pointer = 1
    For cntr1 = 0 To 5              'Compare all conts
        Cont = ContPriority(cntr1)
        prospects(Cont) = BigInCont(Cont, gPlayerTurn, (gPlayerValue))
    Next cntr1
    
    For cntr2 = 0 To 5              'Try to find cont 6 times
        Highest = 0
        For cntr1 = 0 To 5          'Look at all conts
            Cont = ContPriority(cntr1)
            If prospects(Cont) > Highest Then
                Highest = prospects(Cont)
                HstCont = Cont
            End If
        Next cntr1
        If Highest < 1 Then
            AutoPut2 = False
            Exit Function
        ElseIf putInThisCont(HstCont) Then
            AutoPut2 = True
            Exit Function
        End If
        
        prospects(HstCont) = 0
        Highest = 0
    Next cntr2
    AutoPut2 = False
End Function

    'Returns true if any neighbours are in specified continent
Private Function AgateNextToCont(ctry As Integer, Cont As Integer, Who As Integer) As Boolean
    Dim cntr1 As Integer
    
    For cntr1 = 1 To 7
        If CountryID(ctry).neighbour(cntr1) = 0 Then
            AgateNextToCont = False
            Exit Function
        ElseIf ContinentOfCtry(CountryID(ctry).neighbour(cntr1), gPlayerTurn) = Cont Then
            AgateNextToCont = True
            Exit Function
        End If
    Next cntr1
    AgateNextToCont = False
End Function

    'Returns number of any neighbours player(who)
Private Function AllyNextToCtry(ctry As Integer, Who As Integer) As Integer
    Dim cntr1 As Integer
    
    For cntr1 = 1 To 7
        If CountryID(ctry).neighbour(cntr1) = 0 Then
            AllyNextToCtry = 0
            Exit Function
        ElseIf gCountryOwner(CountryID(ctry).neighbour(cntr1)) = Who Then
            AllyNextToCtry = CountryID(ctry).neighbour(cntr1)
            Exit Function
        End If
    Next cntr1
    AllyNextToCtry = 0
End Function

    'Return smallest difference on ctry border inside cont
    'ie. my points - smallest enemy, higer the better for attacking
Private Function getMinDiff(Cont As Integer, ctry As Integer) As Integer
    Dim cntr1 As Integer, hiDiff As Integer, tmp As Integer
    Dim tst As Boolean
    
    getMinDiff = -10000
    For cntr1 = 1 To 7
        tmp = CountryID(ctry).neighbour(cntr1)
        If tmp = 0 Then
            Exit Function
        End If
        tst = (ContinentOfCtry(ctry, gPlayerTurn) = ContinentOfCtry(tmp, gPlayerTurn)) _
            And (gCountryOwner(tmp) <> gPlayerTurn) _
            And (getMinDiff < gCtryScore(ctry) + gPlayerValue - gCtryScore(tmp))
        If tst Then
            getMinDiff = gCtryScore(ctry) + gPlayerValue - gCtryScore(tmp)
        End If
    Next cntr1
End Function

    'Pick a country in this continent
Private Function putInThisCont(Cont As Integer) As Boolean
    Dim cntr1 As Integer
    Dim tmp As Integer, tmp1 As Integer, tmp2 As Integer
    Dim Highest As Integer, scoreDiff As Integer
    Dim ctry As Integer, prospect As Integer
    Dim tst As Boolean
    Dim TempGate(5) As Integer
    
        'Find a spot inside continent first
    Highest = -10000
    For cntr1 = Continents(Cont - 1).FirstCountry To Continents(Cont - 1).LastCountry
        
        tst = (gCountryOwner(cntr1) = gPlayerTurn) _
                And enemyInCont(Cont, cntr1, gPlayerTurn)
        If tst Then
            scoreDiff = getMinDiff(Cont, cntr1)
            If scoreDiff > Highest Then
                Highest = scoreDiff
                ctry = cntr1
            End If
            If scoreDiff = Highest Then
                tmp1 = CountEnemyNbrs(Cont, ctry)
                tmp2 = CountEnemyNbrs(Cont, cntr1)
                tst = ((tmp1 = tmp2) And GenRandom4() > 0.66) Or (tmp1 < tmp2)
                If tst Then        'Find best spot, or random if choices are same
                    Highest = scoreDiff
                    ctry = cntr1
                End If
            End If
        End If
    Next cntr1

    If ctry > 0 Then
        Highest = gCtryScore(ctry) + gPlayerValue
    ElseIf Highest = -10000 Then
        Highest = 0
    Else
        Highest = 1
    End If
    
        'Now find a spot outside cont (but touching)
    Call SetTempGatesForCont(Cont, TempGate, gPlayerTurn)
    For cntr1 = 1 To 5
        If TempGate(cntr1) = 0 Then
            Exit For
        End If
        
        prospect = AllyNextToCtry(TempGate(cntr1), gPlayerTurn)
        If prospect > 0 Then
            tmp = (gPlayerValue + gCtryScore(prospect) - 1) - (getAvg / playerDefence(gPlayerTurn))
            If tmp < 1 Then
                tmp = 1
            End If
            tst = (gCountryOwner(prospect) = gPlayerTurn) _
                    And enemyInCont(Cont, prospect, gPlayerTurn) _
                    And (Highest < tmp)
        Else
            tst = False
        End If
                
        If tst Then
            Highest = gCtryScore(cntr1)
            ctry = prospect
            prefercont = Cont
        End If
    Next cntr1
    
    If Highest = 0 Then
        putInThisCont = False
        prefercont = 0
        Exit Function
    End If
    putInThisCont = True
    AutoCountry = ctry
    
    If gCurrentMode = 2 Then
        
        'Click country.
        gComputerPressed = True
        If gCurrentCardValue < 10 Then
            transferNmbr = 9999
        Else
            transferNmbr = 9999 'gCurrentCardValue / 2
        End If
        Call ClickMapNow
        gComputerPressed = False
    End If
End Function

    'Returns true if enemy about in continent
Private Function enemyInCont(Cont As Integer, ctry As Integer, Who As Integer) As Boolean
    Dim cntr1 As Integer
    enemyInCont = False
    If ContinentOfCtry(ctry, gPlayerTurn) = Cont Then
        For cntr1 = 1 To 7
            If CountryID(ctry).neighbour(cntr1) = 0 Then
                Exit For
            ElseIf gCountryOwner(CountryID(ctry).neighbour(cntr1)) <> Who Then
                If ContinentOfCtry(CountryID(ctry).neighbour(cntr1), gPlayerTurn) = Cont Then
                    enemyInCont = True
                    Exit For
                End If
            End If
        Next cntr1
    End If
End Function

    'Returns value in cont as a percentage.
Private Function BigInCont(Cont As Integer, Who As Integer, Offset As Integer) As Integer
    Dim cntr1 As Integer, cntr2 As Integer
    Dim Highest As Integer, MyPoints As Integer, theirPoints As Integer
    Dim tst As Integer, tmp As Long
    
    On Error GoTo ErrHand
    tst = (IsContPartOfMission(Cont, Who)) ' And (GetOccupiedContinentValues > 1)
    If tst Then
        MyPoints = ((getAvg) + Offset) * ((7 - CountActiveStartingPlayers) * 4 - 3)
    Else
        MyPoints = Offset
    End If
    
    theirPoints = 0
    For cntr1 = 1 To 6
        For cntr2 = Continents(Cont - 1).FirstCountry To Continents(Cont - 1).LastCountry
            If cntr1 = Who Then
                If gCountryOwner(cntr2) = Who Then
                    MyPoints = MyPoints + gCtryScore(cntr2)
                End If
            End If
                '** What's going on here?? Something to do with missions?
                theirPoints = theirPoints + gCtryScore(cntr2)
            
        Next cntr2
        Highest = theirPoints
    Next cntr1
    
    If Highest <> 0 Then
        tst = (MyPoints / Highest) * 100
        If tst > 15000 Then
            tst = 12000
        End If
        BigInCont = Int(tst) + 1
    Else
        BigInCont = 0
    End If
    If getAvg > 20 Then         'A bit of blonde logic
        BigInCont = BigInCont * (GenRandom4 * 0.5 + 0.75) '(GenRandom4 * 2 + 0.2)
    Else
        'BigInCont = BigInCont * (GenRandom4 * 0.5 + 0.75)
    End If
    
    For cntr2 = Continents(Cont - 1).FirstCountry To Continents(Cont - 1).LastCountry
        If killHim(cntr2) Then
            BigInCont = BigInCont * 2
            Exit For
        End If
    Next cntr2
    Exit Function
ErrHand:
    Resume Next
End Function

    'Returns true if who is in cont
Private Function isInCont(Cont As Integer, Who As Integer) As Boolean
    Dim cntr1 As Integer
    
    For cntr1 = Continents(Cont - 1).FirstCountry To Continents(Cont - 1).LastCountry
        If gCountryOwner(cntr1) = Who Then
            isInCont = True
            Exit Function
        End If
    Next cntr1
    isInCont = False
End Function

    'returns true if own all conts on gate
Private Function ContOfCtryNbr(cntry As Integer) As Boolean
    Dim cntr As Integer, Cont As Integer, temp As Integer
        
    Cont = ContinentOfCtry(cntry, gPlayerTurn)
    For cntr = 1 To 7
        temp = CountryID(cntry).neighbour(cntr)
        If temp = 0 Then
            ContOfCtryNbr = True
            Exit Function
        ElseIf Not OwnContinent(ContinentOfCtry(temp, gPlayerTurn), gPlayerTurn) Then
            ContOfCtryNbr = False
            Exit Function
        End If
    Next cntr
    ContOfCtryNbr = True
End Function

    'Returns which continent a country is in.
    'If playerWho <> 0 then extend borders into other continents.
Private Function ContinentOfCtry(cntry As Integer, Optional playerWho As Integer = 0) As Integer
    Dim cntr As Integer
    Dim rslt As Integer
    
    For cntr = 0 To 5
        If (cntry >= Continents(cntr).FirstCountry) _
            And (cntry <= Continents(cntr).LastCountry) Then
            rslt = cntr + 1
            Exit For
        End If
    Next cntr
    If playerWho > 0 Then
            'Make Mid East part of Africa if held.
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 27, 4)
            'Ukraine to Asia.
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 20, 5)
            'North Africa to South America
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 21, 2)
            'North Africa to Europe
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 21, 3)
            'Greenland to Europe
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 3, 3)
            'Central America to South America
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 9, 2)
            'Venusuala to North America
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 10, 1)
            'Kamchata to North America
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 38, 1)
            'Iceland to North America
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 14, 1)
            'NW Territory to Asia
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 1, 5)
            'Indonesia to Asia
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 39, 5)
            'Siam to Australia
        rslt = HoldCountryAndCont(rslt, playerWho, cntry, 30, 6)
    End If
    ContinentOfCtry = rslt
End Function

    'Return True if own country and continents.
Private Function HoldCountryAndCont(ContSoFar As Integer, playerWho As Integer, _
CntryLookAt As Integer, CntryCompare As Integer, Cont1 As Integer, _
Optional Cont2 As Integer = 0) As Integer
    If gCountryOwner(CntryCompare) = playerWho And CntryLookAt = CntryCompare Then
        If Cont2 = 0 Then
            If OwnContinent(Cont1, playerWho) Then
                ContSoFar = Cont1
            End If
        Else
            If OwnContinent(Cont1, playerWho) And OwnContinent(Cont2, playerWho) Then
                ContSoFar = Cont1
            End If
        End If
    End If
    HoldCountryAndCont = ContSoFar
    
End Function

    'Returns how many points are sitting on the gate
Private Function AutoPointsOnGate(Continent As Integer, gate As Integer, TempGates() As Integer) As Integer
    Dim cntr As Integer, rslt As Integer, cntry
    
    rslt = 0
    cntry = TempGates(gate)
    If cntry = 0 Then
        Exit Function
    End If
    For cntr = 1 To 7
        If CountryID(cntry).neighbour(cntr) = 0 Then
            Exit For
        End If
        If gCountryOwner(CountryID(cntry).neighbour(cntr)) <> gPlayerTurn Then
            If gCtryScore(CountryID(cntry).neighbour(cntr)) > rslt Then
                rslt = gCtryScore(CountryID(cntry).neighbour(cntr))
            End If
        End If
    Next cntr
    AutoPointsOnGate = rslt
End Function

Private Sub Auto2Attack()           'Attack
    Call Auto1Attack
End Sub

    'Pick up attacking points depending on situation
Private Sub Auto2Mode21()           'Attack from
    Dim tst0 As Boolean, tst1 As Boolean, tst2 As Boolean
    Dim tst3 As Boolean, Tst4 As Boolean, tst5 As Boolean
    Dim tstCommitAll As Boolean
    Dim avg As Integer
    
    If AtkForCard.On Then
        Call A200AttackWithAll
        Exit Sub
    ElseIf A2opportunity.IsActive Then        'Get some strange numbers sometimes
        With A2opportunity
        If .Path(.pathPointer) <> 0 And .Path(.pathPointer) > 42 Then
            .IsActive = False
            .pathPointer = 0
            .Path(0) = 0
            Exit Sub
        End If
        End With
        Call A2opportunityFrom
        Exit Sub
    End If
        
    avg = getAvg
    
        ' Keep borders covered
    tst0 = ((AutoCountry = 30) And (Not OwnContinent(6, gPlayerTurn)) _
        And (gTargetCtry <> 39)) _
        Or ((AutoCountry = 21) And (Not OwnContinent(2, gPlayerTurn)) _
        And (gTargetCtry <> 12)) _
        Or ((AutoCountry = 9) And (Not OwnContinent(2, gPlayerTurn)) _
        And (gTargetCtry <> 10)) _
        Or ((AutoCountry = 1) And (Not OwnContinent(5, gPlayerTurn)) _
        And (gTargetCtry <> 38)) _
        Or ((AutoCountry = 3) And (Not OwnContinent(3, gPlayerTurn)) _
        And (gTargetCtry <> 14)) _
        Or ((AutoCountry = 38) And (Not OwnContinent(1, gPlayerTurn)) _
        And (gTargetCtry <> 1)) _
        Or ((AutoCountry = 14) And (Not OwnContinent(1, gPlayerTurn)) _
        And (gTargetCtry <> 1))
    
    
    tst1 = (OwnContinent(ContinentOfCtry(AutoCountry, gPlayerTurn), gPlayerTurn))
    tst2 = (gCtryScore(AutoCountry) > gCtryScore(gTargetCtry) + (avg / playerDefence(gPlayerTurn)) + 1)
    tst3 = (ContinentOfCtry(AutoCountry, gPlayerTurn) <> ContinentOfCtry(gTargetCtry, gPlayerTurn))
    Tst4 = (ContinentOfCtry(AutoCountry, gPlayerTurn) <> ContinentOfCtry(gTargetCtry, gPlayerTurn)) _
        And (tst1)
    tst5 = (ContinentOfCtry(AutoCountry, gPlayerTurn) = ContinentOfCtry(gTargetCtry, gPlayerTurn)) _
        And A3.IsAgate(CLng(AutoCountry))
    
    
    If IsEasyAttackOn Then  'Commit all if easy attack.
        transferNmbr = 9999
    ElseIf tst0 Then    'Block out countries when going past (not thru) gates (eg Siam/Aust)
        transferNmbr = (gCtryScore(AutoCountry) * 2 / 3)
    ElseIf tst1 Then    'Own Cont in at moment (ie leaving). If I own all countries in cont, then I must be leaving.
        transferNmbr = gCtryScore(AutoCountry) - (avg / playerDefence(gPlayerTurn))
    ElseIf tst3 Then    'Leaving cont not owned
        transferNmbr = gCtryScore(AutoCountry) - (avg / (playerDefence(gPlayerTurn) * 2))
    ElseIf Tst4 Then    'Never get here - Tst4 says same thing as tst1.
        transferNmbr = gCtryScore(AutoCountry) - (avg / playerDefence(gPlayerTurn))
    ElseIf tst5 Then    'Past a gate but not leaving cont.
        transferNmbr = CInt(CLng(gCtryScore(AutoCountry)) * 4 / 5)
        If (boolIssueCard) And (transferNmbr > 4) Then
            transferNmbr = transferNmbr - 1
        End If
        
    Else
        transferNmbr = 9999
    End If
    If transferNmbr < 2 Then
        If AutoGoAgain Then
            gCurrentMode = 1
        Else
            If EasyAttack Then
                prefercont = 0
                Exit Sub
            End If
            Call Auto1Move
            If gCurrentMode = 10 Then
                Exit Sub
            End If
            If A2FindStranded Then  '4
                gCurrentMode = 16
                Timer2.Interval = playSpeed * 14
                Timer2.Enabled = True
            Else
                gCurrentMode = 14
            End If
        End If
        Exit Sub
    End If
    Call ClickMapNow
    Call AutoFindAttacker
End Sub

Private Sub A200AttackWithAll()
    Dim cntr As Integer, nbr As Integer
    Dim tmpAtk As Attack
    
    If AtkForCard.To <> gTargetCtry Then
        tmpAtk = AtkForCard
        gComputerPressed = True
        Call AttackClicked
        gComputerPressed = False
        If gCurrentMode <> 20 Then
            gCurrentMode = 16
        Else
            AtkForCard = tmpAtk
        End If
        Exit Sub
    End If
    
    transferNmbr = 9999
    AutoCountry = AtkForCard.from
    Call ClickMapNow
    
    For cntr = 1 To 7
        nbr = CountryID(AtkForCard.To).neighbour(cntr)
        If nbr = 0 Then Exit For
        If gCountryOwner(nbr) = gPlayerTurn And nbr <> AtkForCard.from Then
            If Not enemyAbout(nbr, gPlayerTurn, AtkForCard.To) Then
                AutoCountry = nbr
                transferNmbr = gCtryScore(nbr) / 2
                Call ClickMapNow
            End If
        End If
    Next
End Sub

    'Following predefined path.
    'Any country < 0 is from country.
Private Sub A2opportunityFrom()
    Dim cntr As Integer, nbr As Integer, ctry As Integer
    
    With A2opportunity
    transferNmbr = 9999
    ctry = .Path(.pathPointer)
    If .Path(.pathPointer + 1) < -1 Then
        Do While .Path(.pathPointer + 1) < -1
            .pathPointer = .pathPointer + 1
            nbr = -.Path(.pathPointer) - 10
            If gCountryOwner(nbr) = gPlayerTurn And gCtryScore(nbr) > 1 Then
                AutoCountry = nbr
                Call ClickMapNow
            End If
        Loop
    Else
        For cntr = 1 To 7
            nbr = CountryID(ctry).neighbour(cntr)
            If nbr = 0 Then Exit For
            If gCountryOwner(nbr) = gPlayerTurn And gCtryScore(nbr) > 1 Then
                AutoCountry = nbr
                Call ClickMapNow
            End If
        Next
    End If
    If gPickedUpUnits = 0 Then
        With A2opportunity
        .IsActive = False
        .pathPointer = 0
        .Path(0) = 0
        End With
    End If
    .pathPointer = .pathPointer + 1
    End With
End Sub

    'Kick the dog when network seems stalled - host only
Private Sub TimerWatch_Timer()
    On Error GoTo erhnd
    If SetupScreen.Visible Then
        Exit Sub
    End If
    
    TimerWatch.Enabled = False
    
    If netWorkSituation = cNetNone Then
        TimerWatch.Enabled = False
        TimerWatch.Interval = 5000
        Exit Sub
    End If
    
    If net.playerOwner(gPlayerTurn - 1) = myTerminalNumber Then
        TimerWatch.Enabled = False
        TimerWatch.Interval = 5000
        Exit Sub            'My player so ignore
    End If
    
    Call netMain.KickNextTerminal(CLng(net.playerOwner(gPlayerTurn - 1)))
    TimerWatch.Interval = 10000
    TimerWatch.Enabled = True
    Exit Sub
erhnd:
    TimerWatch.Interval = 10000
    Resume Next
End Sub

Private Sub tmrFindCPUspeed_Timer()
    tmrFindCPUspeed.Enabled = False
    CPUspeedTimer = True
    If gMapSetupLock Then
        gCurrentMode = 3
        Timer1.Enabled = True
        tmrFindCPUspeed.Enabled = True
    End If
End Sub

'Flash the info box to attract human attention.
Private Sub tmrFlashInfoBox_Timer()
    Static vFlashCount As Long
    Dim vFontPoints As Single
    
    On Error Resume Next
    
    'Do not flash again if already flashed for this player during this turn.
    'This is cleared by RefreshMap() if currentplayer <> tag.
    If tmrFlashInfoBox.Tag = CStr(gPlayerTurn) _
    Or gCurrentMode = 13 _
    Or gCurrentMode = 18 _
    Or Not mnuFlashInfoBox.Checked Then
        'Bail out here if info box is already the correct colour.
        If vFlashCount = 0 Then
            vFlashCount = 0
            tmrFlashInfoBox.Enabled = False
            tmrFlashInfoBox.Interval = 1000
            Exit Sub
        Else
            vFlashCount = 5
        End If
    End If
    
    'Change info box colour on odd and even counts.
    If vFlashCount Mod 2 = 1 Then
        pctInfoBox.BackColor = gPlayerID(gPlayerTurn).bkgndColor
    Else
        pctInfoBox.BackColor = gPlayerFlashColor(gPlayerTurn - 1)
    End If
    
    'Check if this is the last time.
    If vFlashCount >= 5 Then
        tmrFlashInfoBox.Enabled = False
        tmrFlashInfoBox.Interval = 1000
        pctInfoBox.BackColor = gPlayerID(gPlayerTurn).bkgndColor
        vFlashCount = 0
        tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    Else
        tmrFlashInfoBox.Interval = 200
        vFlashCount = vFlashCount + 1
    End If
    
    'Print the tag onto the info box.
    vFontPoints = pctInfoBox.Font.Size
    pctInfoBox.Cls
    pctInfoBox.Print "";
    If InStr(1, pctInfoBox.Tag, "<FONT15>") Then
        pctInfoBox.Font.Bold = True
        pctInfoBox.Font.Size = vFontPoints * 1.5
        Call CenterPrintText(pctInfoBox, Replace(pctInfoBox.Tag, vbCrLf & "<FONT15>", ""))
        pctInfoBox.Font.Bold = False
        pctInfoBox.Font.Size = vFontPoints
    Else
        pctInfoBox.Print pctInfoBox.Tag
    End If
End Sub

'Various houskeeping tasks.
Private Sub tmrWatchDog_Timer()
    Dim vTagTime As Long
    
    On Error Resume Next
    
    tmrWatchDog.Enabled = False
    
    'Hide Declare War and Cancel Setup buttons and show in the menu bar
    'if the size is reduced passed a certain level.
    cmdSetupOk.Visible = Picture1.Height > 320 And Picture1.Width > 570
    cmdSUPcncl.Visible = cmdSetupOk.Visible
    mnuDeclareWar.Visible = Not cmdSetupOk.Visible And SetupScreen.Visible
    mnuCancelSetup.Visible = Not cmdSUPcncl.Visible And SetupScreen.Visible
    
    'Hide attack, move and pass buttons and show in the menu bar
    'if the size is reduced passed a certain level.
    cmdAttack.Visible = Picture1.Height > 260 And Picture1.Width > 450
    cmdMove.Visible = cmdAttack.Visible
    cmdEnd.Visible = cmdAttack.Visible
    mnuAttack.Visible = Not cmdAttack.Visible And Not SetupScreen.Visible
    mnuMove.Visible = Not cmdMove.Visible And Not SetupScreen.Visible
    mnuPass.Visible = Not cmdEnd.Visible And Not SetupScreen.Visible
    
    'Check that the menu hides after a certain time when its flag is set.
    If Trim(mnuFile.Tag) <> "" Then
        vTagTime = CLng(mnuFile.Tag)
        If CLng(Time * 100000) > vTagTime + 5 _
        Or CLng(Time * 100000) < vTagTime Then
            mnuFile.Tag = ""
        End If
    End If
    
    'Check that the timer is running for the computer player.
    'This is needed when a human player is changed to a computer
    'player and the war is resumed right after Global Siege is
    'first started..
    If gPlayerID(gPlayerTurn).playerWho = 1 _
    Or gPlayerID(gPlayerTurn).playerWho = 2 Then
        If Timer2.Interval = 0 Then
            Timer2.Interval = 100
        End If
    End If
    
    'Show English button if any language other than English is selected.
    cmdEnglish.Visible = (gLanguage <> (GetSystemDefaultLangID And &HFF))
    
    'Disable controls when networked.
    If netWorkSituation = cNetClient Then
        mnuOptUndo.Enabled = False
        Toolbar1.Buttons(8).Enabled = False     'Undo
        Toolbar1.Buttons(4).Enabled = False     'Open
        Toolbar1.Buttons(3).Enabled = False     'Reset
    Else
        mnuOptUndo.Enabled = (GetPlayerController(gPlayerTurn) = 0 _
                            And netWorkSituation <> cNetClient _
                            And gCurrentMode <> 2 _
                            And gCurrentMode <> 5) Or gCheatMode.undoEnabled
        Toolbar1.Buttons(8).Enabled = mnuOptUndo.Enabled
        Toolbar1.Buttons(4).Enabled = True
        Toolbar1.Buttons(3).Enabled = True
    End If
    
    'Enable the counter menu bar if I am the host.
    mnuNetCntr.Enabled = (netWorkSituation <> cNetNone)
    
    'Ensure that the fast war on the toolbar is in sync with everything else.
    If optnFastWar.Checked Then
        Toolbar1.Buttons(7).Value = tbrPressed
    Else
        Toolbar1.Buttons(7).Value = tbrUnpressed
    End If
    
    'Make sure the DrawWin timer has an empty tag if the gCurrentMode is no longer 18.
    If gCurrentMode <> 18 And gCurrentMode <> 13 And tmrDrawWin.Tag <> "" Then
        tmrDrawWin.Tag = ""
        tmrDrawWin.Enabled = False
        gWarRestartLock = False
    End If
    
    'Make sure DrawWin is called to prevent hang after win.
    If gCurrentMode = 13 Then
        If tmrDrawWin.Tag = "" Then
            gWarRestartLock = False
            Call DrawWin
        End If
    End If
    
    'Disable the "See Missions" menu item if you are not meant to see it.
    mnuMissionSee.Enabled = Not SetupScreen.Visible _
                            And gPlayerID(gPlayerTurn).playerWho = 0
    Toolbar1.Buttons(14).Enabled = mnuMissionSee.Enabled
    'Disable/enable Missions menu item depending on network situation.
    'mnuMsnMissionOn.Enabled = SetupScreen.Visible And netWorkSituation <> cNetClient
    
    'Turn off and disable the auto restart option if I am a networked client.
    If netWorkSituation = cNetClient Then
        mnuAutoRestart.Checked = False
        mnuAutoRestart.Enabled = False
    
    Else
        mnuAutoRestart.Enabled = True
    End If
    
    'Ensure locks have not been left on causing restart buttons to become unusable.
    If gMapSetupLock And gCurrentMode <> 3 Then
        gMapSetupLock = False
    End If
    
    tmrWatchDog.Enabled = True
End Sub

'The tool menu across the top.
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
    Case "toolNew" '"Setup"
        Call mnuFileNew_Click
    Case "toolReset"     '"Reset"
        Call mnuFileReset_Click
    Case "toolOpen"     '"Open"
        Call mnuFileLoadWar_Click
    Case "toolSave"     '"Save"
        Call mnuFileSaveWar_Click
    Case "toolNet"     '"Net"
        mnuNetAdvanced_Click
    Case "toolFast"     '"Fast"
        Call changeWarSpeed(Not optnFastWar.Checked)
    Case "toolChat"     '"Chat"
        Call mnuNetChat_Click
    Case "toolUndo"     '"Undo"
        Call mnuOptUndo_Click
    Case "toolFind"     '"Find"
        Call mnuNetClientInternet_Click
    'Case "Help"
    '    Call mnuHelpContents_Click
    Case "toolMission"     '"Mission"
        Call mnuMissionSee_Click
    Case "toolConts"     '"Conts"
        Call hlpContMap_Click
    End Select
End Sub

'Choose a font size for the Info Box by scaling up until it fits the box.
Private Sub ScaleInfoBoxFont()
    Dim vCntr As Long
    Dim vFontPoints As Long
    Dim vLongWords(7) As String
    Dim vTestStrY As String
    Dim vTestStrX As String
    
    'Load up some sample long sentences in an effort
    'to find the longest one for the selected language
    'because there are vast differences between lingos.
    vLongWords(0) = Phrase(109)                             '"<Click attacking country(s)>"
    vLongWords(1) = Phrase(72)                              '"<Click destination country>"
    vLongWords(2) = Phrase(104)                             '"<Click attacking country>"
    vLongWords(3) = Phrase(130)                             '"Checking cards"
    vLongWords(4) = Phrase(70) & " 0 " & Phrase(105)
    vLongWords(5) = Phrase(5) & Phrase(73)
    vLongWords(6) = ""
    vLongWords(7) = ""
    
    'Load test strings with maximum expected strings.
    vTestStrY = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    
    vTestStrX = ""
    For vCntr = 0 To UBound(vLongWords)
        If pctInfoBox.TextWidth(vLongWords(vCntr)) > pctInfoBox.TextWidth(vTestStrX) Then
            vTestStrX = vLongWords(vCntr)
        End If
    Next
    
    'Start really really small.
    vFontPoints = 1
    
    'Scale up the font until it no longer fits the info box either height or width.
    For vFontPoints = 6 To 400
        pctInfoBox.Font.Size = vFontPoints / 4
        
        If pctInfoBox.TextHeight(vTestStrY) > pctInfoBox.ScaleHeight _
        Or pctInfoBox.TextWidth(vTestStrX) > pctInfoBox.ScaleWidth Then
            pctInfoBox.Font.Size = (vFontPoints - 1) / 4
            Exit For
        End If
    Next
    
    'Clear and reprint the text in the info box.
    'Problem is that you loose any bold characters
    'but that could be fixed up in a future version
    'if anyone complains.
    pctInfoBox.Cls
    pctInfoBox.Print "";
    If InStr(1, pctInfoBox.Tag, "<FONT15>") Then
        pctInfoBox.Font.Bold = True
        pctInfoBox.Font.Size = ((vFontPoints - 1) / 4) * 1.5
        Call CenterPrintText(pctInfoBox, Replace(pctInfoBox.Tag, vbCrLf & "<FONT15>", ""))
        pctInfoBox.Font.Bold = False
        pctInfoBox.Font.Size = (vFontPoints - 1) / 4
    Else
        pctInfoBox.Print pctInfoBox.Tag
    End If
End Sub

'Choose a font size for the front map by scaling up until it fits the smallest box.
Private Sub ScaleViewportMapFont()
    Dim vCntr As Long
    Dim vFontPoints As Long
    Dim vTestStrY As String
    Dim vTestStrX As String
    
    'Start really really small.
    vFontPoints = 1
    
    'Load test strings with maximum expected strings.
    vTestStrY = "0"
    vTestStrX = "000"
    
    'Scale up the font until it no longer fits the info box either height or width.
    For vFontPoints = 6 To 400
        Picture1.Font.Size = vFontPoints / 4
        
        If Picture1.TextHeight(vTestStrY) > 26 * gPictureMaskRatioY _
        Or Picture1.TextWidth(vTestStrX) > 38 * gPictureMaskRatioX Then
            Picture1.Font.Size = (vFontPoints - 1) / 4
            Exit For
        End If
    Next
    
    'Reprint.
    Call DrawLittleCardText
End Sub

'Choose a font size for the little cards and set
'the global variable "gLittleCardFontSize".
Private Sub ScaleLittleCardFont()
    Dim vCntr As Long
    Dim vOrigFontSize As Single
    Dim vFontPoints As Long
    Dim vTestStrY As String
    Dim vTestStrX As String
    Dim vMaxHeight As Long
    Dim vMaxWidth As Long
    
    vOrigFontSize = Picture1.Font.Size
    
    vMaxHeight = gMsk.LittleCardTop * gPictureMaskRatioY
    vMaxWidth = (gMsk.LittleCrdSngWidth * 4) * gPictureMaskRatioX
    
    'Start really really small.
    vFontPoints = 1
    
    'Load test strings with maximum expected strings.
    vTestStrY = Trim(Phrase(5))             '"The Purple Army"
    vTestStrX = vTestStrY

    'Scale up the font until it no longer fits the text area either height or width.
    For vFontPoints = 6 To 400
        Picture1.Font.Size = vFontPoints / 4
        
        If Picture1.TextHeight(vTestStrY) > vMaxHeight _
        Or Picture1.TextWidth(vTestStrX) > vMaxWidth Then
            gLittleCardFontSize = (vFontPoints - 1) / 4
            Exit For
        End If
    Next
    
    Picture1.Font.Size = vOrigFontSize
End Sub

'Resize form and all elements within as required.
Public Sub ResizeForm()
    Dim vFormHeight As Long
    Dim vFormWidth As Long
    Dim vSetupTop As Long
    Dim vSetupLeft As Long
    Dim vBorder As Long

    vFormHeight = TheMainForm.ScaleHeight
    vFormWidth = TheMainForm.ScaleWidth
    
    If mnuViewBorder.Checked Then
        vBorder = CLng(vFormWidth / gBorderWidth)      '8
    Else
        vBorder = 0
    End If
    
    If vFormHeight < GetToolbarHeight + vBorder Then
        Exit Sub
    End If
    
    If vFormWidth < 2 * vBorder Then
        Exit Sub
    End If
    
    'If tag is non zero length then do not resize the picture.
    If Picture1.Tag = "" Then
        Picture1.Move vBorder, GetToolbarHeight, vFormWidth - (2 * vBorder), _
                        vFormHeight - GetToolbarHeight - vBorder
    Else
        TheMainForm.Picture1.Width = 14220
        TheMainForm.Picture1.Height = 9810
        Picture1.Move (vFormWidth - Picture1.Width) / 2, GetToolbarHeight
    End If
    
    gPictureMaskRatioX = GetPictureMaskRatioX
    gPictureMaskRatioY = GetPictureMaskRatioY
    
    Call ScaleViewportMapFont
    Call ScaleLittleCardFont
    
    'Hide Declare War and Cancel Setup buttons and show in the menu bar
    'if the size is reduced passed a certain level.
    cmdSetupOk.Visible = Picture1.Height > 320 And Picture1.Width > 570
    cmdSUPcncl.Visible = cmdSetupOk.Visible
    mnuDeclareWar.Visible = Not cmdSetupOk.Visible And SetupScreen.Visible
    mnuCancelSetup.Visible = Not cmdSUPcncl.Visible And SetupScreen.Visible
    
    'Hide attack, move and pass buttons and show in the menu bar
    'if the size is reduced passed a certain level.
    cmdAttack.Visible = Picture1.Height > 260 And Picture1.Width > 450
    cmdMove.Visible = cmdAttack.Visible
    cmdEnd.Visible = cmdAttack.Visible
    mnuAttack.Visible = Not cmdAttack.Visible And Not SetupScreen.Visible
    mnuMove.Visible = Not cmdMove.Visible And Not SetupScreen.Visible
    mnuPass.Visible = Not cmdEnd.Visible And Not SetupScreen.Visible
    
    'Hide the card cancel button.
    cmdCardCncl.Visible = Picture1.Height > 250 And cmdExchange.Visible
    
    'Resize the attack, move and pass buttons.
    If Picture1.Height < 370 Or Picture1.Width < 500 Then
        cmdAttack.Height = 17
        cmdAttack.Font.Size = 8
        cmdExchange.Height = 17
        cmdExchange.Font.Size = 8
    ElseIf Picture1.Height < 420 Or Picture1.Width < 630 Then
        cmdAttack.Height = 25
        cmdAttack.Font.Size = 10
        cmdExchange.Height = 25
        cmdExchange.Font.Size = 10
    Else
        cmdAttack.Height = 33
        cmdAttack.Font.Size = 10
        cmdExchange.Height = 25
        cmdExchange.Font.Size = 10
    End If
    cmdMove.Height = cmdAttack.Height
    cmdMove.Font.Size = cmdAttack.Font.Size
    cmdEnd.Height = cmdAttack.Height
    cmdEnd.Font.Size = cmdAttack.Font.Size
    cmdCardCncl.Height = cmdExchange.Height
    cmdCardCncl.Font.Size = cmdExchange.Font.Size
    
    'Move things in the bottom left corner.
    'Info Box (pctInfoBox)
    pctInfoBox.Left = 8 * gPictureMaskRatioX
    pctInfoBox.Width = 226 * gPictureMaskRatioX
    pctInfoBox.Height = 136 * gPictureMaskRatioY
    pctInfoBox.Top = Picture1.Height - pctInfoBox.Height - (8 * gPictureMaskRatioY)
    Call ScaleInfoBoxFont
    
    'Transfer selector options box.
    pctTransfer.Top = pctInfoBox.Top - pctTransfer.Height - 5
    'If cmdEnd.Visible Then
    '    pctTransfer.Left = cmdEnd.Left + cmdEnd.Width + 10
    'Else
    '    pctTransfer.Left = pctInfoBox.Left
    'End If
    pctTransfer.Left = pctInfoBox.Left
    
    'Attack, Move and Pass buttons.
    cmdEnd.Top = pctInfoBox.Top - cmdEnd.Height - 5
    cmdEnd.Left = pctInfoBox.Left + pctTransfer.Width '+ 5
    cmdMove.Top = cmdEnd.Top - cmdMove.Height
    cmdMove.Left = cmdEnd.Left
    cmdAttack.Top = cmdMove.Top - cmdAttack.Height
    cmdAttack.Left = cmdMove.Left
    
    cmdEnglish.Left = Picture1.Width - cmdEnglish.Width
    cmdEnglish.Top = Picture1.Height - cmdEnglish.Height
    
    SetupScreen.Top = 0
    SetupScreen.Left = 0
    SetupScreen.Height = Picture1.Height
    SetupScreen.Width = Picture1.Width
    
    'Frame frameSetupControls is a lightweight control which has
    'no scale mode meaning it uses twips instead of pixels.
    'Hard coded the move values for better looks.
    vSetupLeft = (SetupScreen.ScaleWidth - frameSetupControls.Width) / 2
    If vSetupLeft < -260 Then
        vSetupLeft = -260
    End If
    frameSetupControls.Left = vSetupLeft
    
    vSetupTop = (SetupScreen.ScaleHeight - frameSetupControls.Height) / 2
    If vSetupTop > 232 Then
        vSetupTop = 232
    ElseIf vSetupTop < -330 Then
        vSetupTop = -330
    End If
    frameSetupControls.Top = vSetupTop
    
    'Make sure the viewport gets refreshed.
    gSyncViewportNeeded = True
    
    Call SyncForgroundMap("Form_Resize")
    cmdExchange.Top = gMsk.CrdMainTop * GetPictureMaskRatioY
    cmdExchange.Left = gMsk.CrdMainLeft * GetPictureMaskRatioX - cmdExchange.Width
    cmdCardCncl.Top = cmdExchange.Top + cmdExchange.Height
    cmdCardCncl.Left = cmdExchange.Left
    
    Call CheckAppFontAndSize
End Sub

Private Sub Form_Resize()
    Call ResizeForm
End Sub

'Handle large fonts and small windows size (originally written 27.11.98).
'Move setup controls if required.
Private Sub CheckAppFontAndSize()
    Dim cntr As Long
    
    If pctTransfer.Height >= 150 Then           'Large font set
        Call CompactSetupControls
        Call LargeFontSetupControls
    ElseIf GetDeviceCaps(Picture1.hdc, 8) < 700 _
    Or cmdSUPcncl.Left > TheMainForm.Width Then
        Call CompactSetupControls                      '640x480
    End If
End Sub

'Move setup controls into top left corner if they do not fit the window.
Private Sub CompactSetupControls()
    Dim cntr As Long
    'With PlayerNumber
    '    .Height = 975
    '    .Left = 4800
    '    .Top = 480
    '    .Width = 1455
    'End With
    'With Frame8
    '    .Height = 1215
    '    .Left = 4800
    '    .Top = 1680
    '    .Width = 1455
    'End With
    'With Frame7
    '    .Height = 3135
    '    .Left = 4800
    '    .Top = 3120
    '    .Width = 1455
    'End With
    
    'optSupplyLines.Top = 840
    'optLimitSupply.Top = 1080
    'optNoSupply.Top = 1320
    'chkFast.Top = 1800
    'chkFastDice.Top = 2160
    'chkOptimizeDefenceDice.Top = 2520
    'chkBorder.Top = 3000
    'chkExtraStartingUnits.Top = 3600
    'Label3.Visible = False
    'plrStrt(8).Top = 4080
    'plrStrt(8).Left = 600
    'pctUpDown(8).Top = 4080
    'pctUpDown(8).Left = 1095
    
    'With Frame9
    '    .Height = 4575
    '    .Left = 6360
    '    .Top = 480
    '    '.Width = 2175
    'End With
    'With cmdSetupOk
    '    .Height = 495
    '    .Left = 6840
    '    .Top = 5160
    '    .Width = 1215
    'End With
    'With cmdSUPcncl
    '    .Height = 495
    '    .Left = 6840
    '    .Top = 5760
    '    .Width = 1215
    'End With
    
    'For cntr = 0 To 5
    '    plrOpt(cntr).Height = 975
    '    plrOpt(cntr).Left = 120
    '    plrOpt(cntr).Width = 4575
    '    plrOpt(cntr).Top = 480 + 960 * cntr
    'Next cntr
End Sub

'Large font set, change size and pos of things.
Private Sub LargeFontSetupControls()
    Dim i As Long
    
    On Error GoTo erhand

    With cmdAttack
        .Height = 33
        .Left = 80
        .Top = 280
        If L = 3 Then
            .Width = 65
        Else
            .Width = 55
        End If
    End With
    With cmdCardCncl
        .Height = 25
        .Left = 344
        .Top = 472
        .Width = 81
    End With
    With cmdEnd
        .Height = 33
        .Left = 80
        .Top = 360
        If L = 3 Then
            .Width = 73
        Else
            .Width = 55
        End If
    End With
    With cmdExchange
        .Height = 25
        .Left = 344
        .Top = 448
        .Width = 81
        .Font.Size = 5
    End With
    With cmdMove
        .Height = 33
        .Left = 80
        .Top = 320
        If L = 3 Then
            .Width = 73
        Else
            .Width = 55
        End If
    End With
    With tfRate1
        .Height = 195
        .Left = 120
        .Top = 240
        .Width = 495
    End With
    With tfRate2
        .Height = 195
        .Left = 120
        .Top = 480
        .Width = 495
    End With
    With tfRate5
        .Height = 195
        .Left = 120
        .Top = 720
        .Width = 495
    End With
    With tfRate10
        .Height = 195
        .Left = 120
        .Top = 960
        .Width = 495
    End With
    With tfRate20
        .Height = 195
        .Left = 120
        .Top = 1200
        .Width = 495
    End With
    tfRate50.Visible = False
    With tfRateAll
        .Height = 195
        .Left = 120
        .Top = 1440
        .Width = 615
    End With
    
    Exit Sub
erhand:
    Resume Next
End Sub

'----------------------------------------------------------------------------
'  Actions for the up/down controls on the setup screen.
'----------------------------------------------------------------------------

Private Sub udCardDeck_Change(Index As Integer)
    On Error Resume Next
    If udCardDeck(Index).Tag = "" Then
        Call CheckCardDeckCounts(Index)
        txtCardDeck(Index).Text = CStr(udCardDeck(Index).Value)
        net.setupControlChange = True
    End If
End Sub

'Change card deck counts in the setup screen without all the checks. This is to allow
'changes to the setup screen from the host in a network war or from file without
'changing the value back if it breaks the minimum card requirements.
Private Sub ChangCardDeckCounts(pSuiteIX As Integer, pNewCount As Integer)
    udCardDeck(pSuiteIX).Tag = "No Checking"
    udCardDeck(pSuiteIX).Value = pNewCount
    txtCardDeck(pSuiteIX).Text = CStr(pNewCount)
    udCardDeck(pSuiteIX).Tag = ""
End Sub

'Check there are enough cards in the deck. Minimum is the worst case
'scenario which is 5 cards for 5 palyers, so there must be at least 25 cards.
Private Sub CheckCardDeckCounts(Index As Integer)
    Dim vCardCount As Long
    
    vCardCount = CountStartingCardsInDeck
    If vCardCount < 25 Then
        udCardDeck(Index).Tag = "No Callback"
        udCardDeck(Index).Value = udCardDeck(Index).Value + 25 - vCardCount
        udCardDeck(Index).Tag = ""
    End If
End Sub

'Return the number of starting cards in the deck as selected from the Setup form.
Private Function CountStartingCardsInDeck() As Long
    Dim vIndex As Long
    
    On Error Resume Next
    For vIndex = 0 To udCardDeck.Count - 1
        CountStartingCardsInDeck = CountStartingCardsInDeck + udCardDeck(vIndex).Value
    Next
End Function

Private Sub udDiceThrown_Change(Index As Integer)
    On Error Resume Next
    txtDiceThrown(Index).Text = CStr(udDiceThrown(Index).Value)
    Call PrintDiceOdds
    net.setupControlChange = True
End Sub

Private Sub udFixedValues_Change(Index As Integer)
    On Error Resume Next
    txtFixedValues(Index).Text = CStr(udFixedValues(Index).Value)
    net.setupControlChange = True
End Sub

Private Sub udContVal_Change(Index As Integer)
    On Error Resume Next
    txtContVal(Index).Text = CStr(udContVal(Index).Value)
    net.setupControlChange = True
    Call frmContinents.UpdateContDetails
End Sub

Private Sub udNewUnitClac_Change(Index As Integer)
    On Error Resume Next
    txtNewUnitClac(Index).Text = CStr(udNewUnitClac(Index).Value)
    net.setupControlChange = True
    Call PluralReinforcementTabText
End Sub

'Player controller u/d control.
Private Sub vscrollPlayerSelect_Change(Index As Integer)
    On Error Resume Next
    If vscrollPlayerSelect(Index).Tag = "" Then
        Call ChangePlayerSelect(Index)
    End If
End Sub

'Change plurals in the Reinforcement text and
'move recruit text and number selectors as required. as required.
Private Sub PluralReinforcementTabText()
    If udNewUnitClac(0).Value = 1 Then
        lblNewUnitClac(1).Caption = "country occupied"
    Else
        lblNewUnitClac(1).Caption = "countries occupied"
    End If
    If udNewUnitClac(1).Value = 1 Then
        lblNewUnitClac(2).Caption = "battalion will be drafted with"
    Else
        lblNewUnitClac(2).Caption = "battalions will be drafted with"
    End If
    If udNewUnitClac(2).Value = 1 Then
        lblNewUnitClac(4).Caption = "battalion"
    Else
        lblNewUnitClac(4).Caption = "battalions"
    End If
    
    'Move recruit text and number selectors as required.
    txtNewUnitClac(0).Left = lblNewUnitClac(0).Left + lblNewUnitClac(0).Width
    udNewUnitClac(0).Left = txtNewUnitClac(0).Left + txtNewUnitClac(0).Width - 15
    lblNewUnitClac(1).Left = udNewUnitClac(0).Left + udNewUnitClac(0).Width + 45
    
    udNewUnitClac(1).Left = txtNewUnitClac(1).Left + txtNewUnitClac(1).Width - 15
    lblNewUnitClac(2).Left = udNewUnitClac(1).Left + udNewUnitClac(1).Width + 45
    
    txtNewUnitClac(2).Left = lblNewUnitClac(3).Left + lblNewUnitClac(3).Width
    udNewUnitClac(2).Left = txtNewUnitClac(2).Left + txtNewUnitClac(2).Width - 15
    lblNewUnitClac(4).Left = udNewUnitClac(2).Left + udNewUnitClac(2).Width + 45
End Sub

Private Sub ChangePlayerSelect(pIndex As Integer)
    Dim pHoldPlrIx As Integer
    
    pHoldPlrIx = playerSelect_getIndex(pIndex)
    
    'Change the player.
    PlayerSelect(pIndex).ListIndex = vscrollPlayerSelect(pIndex).Value
    
    'Invalidate stats.
    If gPlayerStats(pIndex + 1).IsValid Then
        gPlayerStats(pIndex + 1).IsValid = False
        gPlayerStats(pIndex + 1).InvalidatedReason = "Changed controller" _
                                                & vbCrLf & "during the war."
    End If
    
    'Flag the setup controls as having changed.
    If (pHoldPlrIx = remoteIndex) Xor (PlayerSelect(pIndex).ListIndex = CVi(remoteIndex)) Then
        net.setupControlChange = True
    End If
    
    'Network stuff for player change.
    If netWorkSituation <> cNetNone Then
        If net.playerOwner(pIndex) <> 0 _
        And net.playerOwner(pIndex) <> myTerminalNumber Then
            Call playerSelect_showIndex(pIndex, remoteIndex)
            net.setupControlChange = False
        Else
            Call CheckSetupForChange
        End If
        
        'Make sure remote player listed as "Remote Player".
        If netWorkSituation = cNetClient _
        And net.playerOwner(pIndex) = myTerminalNumber _
        And vscrollPlayerSelect(pIndex).Value = CVi(remoteIndex) Then
            RenameRemote CByte(pIndex), Phrase(28)                 'Remote Player.
        End If
    End If
End Sub

'Return the color of enabled or disabled text.
Public Function SetEnabledTextColor(pEnabled As Boolean) As Long
    If pEnabled Then
        SetEnabledTextColor = vbBlack
    Else
        SetEnabledTextColor = &H80000011
    End If
End Function

'Enable or disable player selector controls.
Private Sub PlayerSelect_Enable(pPlayerIX As Integer, pEnabled As Boolean)
    PlayerSelect(pPlayerIX).ForeColor = SetEnabledTextColor(pEnabled)
    'PlayerSelect(pPlayerIX).Enabled = pEnabled
    vscrollPlayerSelect(pPlayerIX).Enabled = pEnabled
End Sub

    'playerSelect(player).listIndex = newIndex
Private Sub playerSelect_showIndex(whichPlr As Integer, newIndex As Integer)
    On Error Resume Next
    PlayerSelect(whichPlr).ListIndex = CVi(newIndex)
    vscrollPlayerSelect(whichPlr).Tag = "No action."
    vscrollPlayerSelect(whichPlr).Value = CVi(newIndex)
    vscrollPlayerSelect(whichPlr).Tag = ""
End Sub


Public Function playerSelect_getIndex(whichPlr As Integer) As Integer
    On Error Resume Next
    playerSelect_getIndex = CVo(PlayerSelect(whichPlr).ListIndex)
End Function


'----------------------------------------------------------------------------

'Click action for maximum card value.
Private Sub udMaximumCardValue_Change()
    Dim vModifier As Integer
    
    On Error Resume Next
    
    'If increasing, change the modifier. We ned to use a modifier because
    'the step size changes do not take affect until the next click.
    If udMaximumCardValue.Value > CInt(txtMaximumCardValue.Text) Then
        vModifier = 0
    Else
        vModifier = 1
    End If
    
    txtMaximumCardValue.Text = CStr(udMaximumCardValue.Value)
    
    'Increase units per click for high values.
    If udMaximumCardValue.Value < 20 + (5 * vModifier) Then
        udMaximumCardValue.SmallChange = 1
    ElseIf udMaximumCardValue.Value < 80 + (10 * vModifier) Then
        udMaximumCardValue.SmallChange = 5
    Else
        udMaximumCardValue.SmallChange = 10
    End If
    
    net.setupControlChange = True
End Sub

'Click action for number of starting countries for player(Index)
Private Sub udPlayerStartCountries_Change(Index As Integer)
    On Error Resume Next
    
    If udPlayerStartCountries(Index).Tag = "" Then
        Call CheckPlayerStartingCountries(Index)
        txtPlayerStartCountries(Index).Text = udPlayerStartCountries(Index).Value
    End If
End Sub

'Change a single player's number of starting countries.
Private Sub CheckPlayerStartingCountries(Index As Integer)
    Dim vCountriesLeft As Integer
    
    'Ensure there are enough countries.
    
    
    vCountriesLeft = udPlayerStartCountries(Index).Value + 42 - CountAllocatedCountries
    If udPlayerStartCountries(Index).Value > vCountriesLeft Then
        udPlayerStartCountries(Index).Tag = "No Callback"
        udPlayerStartCountries(Index).Value = vCountriesLeft
        udPlayerStartCountries(Index).Tag = ""
    End If
    
    'Place an "X" in the color box.
    If udPlayerStartCountries(Index).Value < 1 Then
        pctClr(Index).Print " X"
    Else
        pctClr(Index).Cls
    End If
    
    'Change the number of starting armies in the Starting Armies setup box.
    Call SetStartingArmyCount(CountActiveStartingPlayers)
    net.setupControlChange = True
End Sub

'Set the number of starting armies without re-ordering. This is used when army counts are
'changed in code or by changing a players number of starting countries from 0 to 1 or back.
Private Sub SetStartingArmyCount(pNewArmyCount As Integer)
    If pNewArmyCount > 1 Then
        'Stop Starting Army functions from re-ordering knocked out players.
        udStartingArmies.Tag = "No Player Reorg"
        udStartingArmies.Value = pNewArmyCount
        txtStartingArmies.Text = CStr(pNewArmyCount)
        udStartingArmies.Tag = ""
    Else
        txtStartingArmies.Text = CStr(pNewArmyCount)
    End If
End Sub

'Total number of countries accounted for by players
Private Function CountAllocatedCountries() As Integer
    Dim cntr As Integer
    
    CountAllocatedCountries = 0
    For cntr = 0 To 5
        CountAllocatedCountries = CountAllocatedCountries + udPlayerStartCountries(cntr).Value
    Next cntr
End Function

'Sync the u/d control value with the player starting countries text boxes.
'This has to go.
Private Sub SyncPlrStartingCountries()
    Dim vCntr As Long
    
    For vCntr = 0 To 5
        udPlayerStartCountries(vCntr).Value = CInt(txtPlayerStartCountries(vCntr).Text)
    Next
End Sub

'Click action for number of starting armies u/d click.
Private Sub udStartingArmies_Change()
    On Error Resume Next
    
    If udStartingArmies.Tag = "" Then
        txtStartingArmies.Text = CStr(udStartingArmies.Value)
    
        'Don't re-order players if this was changed in code.
        If udStartingArmies.Tag = "" Then
            Call ChangeStartingArmyCount
        End If
    End If
End Sub

'Change starting countries for pPlayer in the setup screen without all the checks.
Private Sub ChangPlayerStartingCtryCount(pPlayerIX As Integer, pNewCount As Integer)
    udPlayerStartCountries(pPlayerIX).Tag = "No Checking"
    udPlayerStartCountries(pPlayerIX).Value = pNewCount
    txtPlayerStartCountries(pPlayerIX).Text = CStr(pNewCount)
    udPlayerStartCountries(pPlayerIX).Tag = ""
    
    'Place an "X" in the color box.
    If udPlayerStartCountries(pPlayerIX).Value < 1 Then
        pctClr(pPlayerIX).Print " X"
    Else
        pctClr(pPlayerIX).Cls
    End If
End Sub

'Change the number of starting players. Triggered from number of
'starting armies scroll bar click on the setup screen.
Private Sub ChangeStartingArmyCount()
    Dim vCountryIndex As Long, tmp As Long
    Dim ctryTotal As Long
    Dim vMaxUnitsForPlayer As Long
    Dim vPlayerIndex As Long
    
    'Start all at 0.
    For vCountryIndex = 0 To 5
        txtPlayerStartCountries(vCountryIndex).Text = 0
    Next
    
    'Distribute countries between players as evenly as possible.
    vMaxUnitsForPlayer = CLng(txtStartingArmies.Text) - 1
    vPlayerIndex = vMaxUnitsForPlayer - 1
    
    '1 for each country.
    For vCountryIndex = 1 To 42
        txtPlayerStartCountries(vPlayerIndex).Text = CStr(CInt(txtPlayerStartCountries(vPlayerIndex).Text) + 1)
        vPlayerIndex = vPlayerIndex - 1
        If vPlayerIndex < 0 Then
            vPlayerIndex = CLng(txtStartingArmies.Text) - 1
        End If
    Next
    
    'Sync the u/d control value with the player starting countries text boxes.
    'Call SyncPlrStartingCountries
    For vPlayerIndex = 0 To 5
        udPlayerStartCountries(vPlayerIndex).Tag = "No Checking"
        udPlayerStartCountries(vPlayerIndex).Value = CInt(txtPlayerStartCountries(vPlayerIndex).Text)
        udPlayerStartCountries(vPlayerIndex).Tag = ""
        
        'Place an "X" in the color box.
        If udPlayerStartCountries(vPlayerIndex).Value < 1 Then
            pctClr(vPlayerIndex).Print " X"
        Else
            pctClr(vPlayerIndex).Cls
        End If
    Next
    
    'How many extra points to distibute.
    Select Case CInt(txtStartingArmies.Text)
    Case 2
        udExtraStartingUnits.Value = 40
    Case 3
        udExtraStartingUnits.Value = 35
    Case 4
        udExtraStartingUnits.Value = 30
    Case 5
        udExtraStartingUnits.Value = 25
    Case Else
        udExtraStartingUnits.Value = 20
    End Select
    
    txtExtraStartingUnits.Text = CStr(udExtraStartingUnits.Value)
    net.setupControlChange = True
End Sub

'Click action for Distribute Extra Units u/d buttons.
Private Sub udExtraStartingUnits_Change()
    Dim vModifier As Integer
    On Error Resume Next
    
    'If increasing, change the modifier. We ned to use a modifier because
    'the step size changes do not take affect until the next click.
    If udExtraStartingUnits.Value > CInt(txtExtraStartingUnits.Text) Then
        vModifier = 0
    Else
        vModifier = 1
    End If
    
    'Increase units per click for high values.
    If udExtraStartingUnits.Value < 15 + (5 * vModifier) Then
        udExtraStartingUnits.SmallChange = 1
    ElseIf udExtraStartingUnits.Value < 100 + (10 * vModifier) Then
        udExtraStartingUnits.SmallChange = 5
        'udExtraStartingUnits.Value = udExtraStartingUnits.Value _
                                        - udExtraStartingUnits.Value Mod 5
    Else
        udExtraStartingUnits.SmallChange = 10
        'udExtraStartingUnits.Value = udExtraStartingUnits.Value _
                                        - udExtraStartingUnits.Value Mod 10
    End If
    
    'Commit the change.
    txtExtraStartingUnits.Text = CStr(udExtraStartingUnits.Value)
    net.setupControlChange = True
End Sub

'----------------------------------------------------------------------------


    'Changes the order of player listings
Private Function CVi(Index As Integer) As Integer
    Select Case Index
    Case 0
        CVi = 0
    Case 1
        CVi = 2
    Case 2
        CVi = 3
    Case 3
        CVi = 1
    Case 4
        CVi = 4
    End Select
End Function

Private Function CVo(Index As Integer) As Integer
    Select Case Index
    Case 0
        CVo = 0
    Case 1
        CVo = 3
    Case 2
        CVo = 1
    Case 3
        CVo = 2
    Case 4
        CVo = 4
    End Select
End Function



'******************************************************************************************
' A3

    ' Initialize A3 at start up
Private Sub setA3()
    Dim cntr As Long
    Dim cntr2 As Long
    
    For cntr = 1 To 42
        A3ContOfCntry(cntr) = CLng(ContinentOfCtry(CInt(cntr)))
        For cntr2 = 1 To 7
            A3CntryNeigbors(cntr, cntr2) = CLng(CountryID(cntr).neighbour(cntr2))
        Next
    Next
    
    For cntr = 0 To 5
        A3ContValue(cntr + 1) = udContVal(cntr).Value
    Next
    
    Call A3.A3Initialize
End Sub

Private Sub rcrs(ctry As Integer, Depth As Long, effort As Long)
    Dim cntr As Long, tmp2 As Long
    
        ' Test for what I am looking for
    If ContinentOfCtry(ctry) = 6 Then
        effort = effort + gCtryScore(ctry)
        If effort < A3UnitsLeft Then
            A3UnitsLeft = effort
            copyA3Path
        End If
        'refreshMap
        'MsgBox " "
        'Debug.Print "x";
    End If
    
        ' Set limit
    If Depth > 12 Then
        Exit Sub
    End If
    
    For cntr = 1 To 7
            ' No more itterations left at this level
        If CountryID(ctry).neighbour(cntr) = 0 Then
            Exit Sub
            
            ' Test target and take actions id pass
        ElseIf gCountryOwner(CountryID(ctry).neighbour(cntr)) <> 1 Then
            A3Itterations = A3Itterations + 1
            tmp2 = gCountryOwner(CountryID(ctry).neighbour(cntr))
            gCountryOwner(CountryID(ctry).neighbour(cntr)) = 1
            A3RecordTo(Depth) = CountryID(ctry).neighbour(cntr)
            
            rcrs CountryID(ctry).neighbour(cntr), Depth + 1, effort + gCtryScore(CountryID(ctry).neighbour(cntr)) + 1
            
            A3RecordTo(Depth) = 0
            gCountryOwner(CountryID(ctry).neighbour(cntr)) = tmp2
        End If
    Next
End Sub

Private Sub copyA3Path()
    Dim cntr As Long
    For cntr = 0 To 49
        A3PathTo(cntr) = A3RecordTo(cntr)
    Next
End Sub

    ' A3's turn
Private Sub A3ProcessIntelligentPlayer()

End Sub

    ' Give A3 info about the current map status
Public Sub A3UpdateMap(whosTurn As Long)
    Dim cntr As Long
    
    For cntr = 1 To 6
        A3NoCntrysOwn(cntr) = 0
    Next
    
    For cntr = 1 To 42
        A3CntryOwner(cntr) = CLng(gCountryOwner(cntr))
        A3CntryScore(cntr) = CLng(gCtryScore(cntr))
        A3NoCntrysOwn(gCountryOwner(cntr)) = A3NoCntrysOwn(gCountryOwner(cntr)) + 1
        Call A3SetNeigbors(cntr, whosTurn)
    Next
End Sub

    ' Count the number if neigbors that have points and put in list
    ' A3CntryAllys(cntr, 0) = number of nbrs
    ' A3CntryAllys(cntr, >0) = actual neighbors
Private Sub A3SetNeigbors(cntr As Long, whosTurn As Long)
    Dim cntr2 As Long
    Dim nbr As Integer
    
    A3CntryAllys(cntr, 0) = 0
    If gCountryOwner(cntr) <> CInt(whosTurn) Then
        For cntr2 = 1 To 6
            nbr = CountryID(cntr).neighbour(cntr2)
            If nbr = 0 Then
                Exit For
            End If
            If gCountryOwner(nbr) = CInt(whosTurn) And gCtryScore(nbr) > 1 Then
                A3CntryAllys(cntr, 0) = A3CntryAllys(cntr, 0) + 1
                A3CntryAllys(cntr, A3CntryAllys(cntr, 0)) = CLng(nbr)
            End If
        Next
    End If
End Sub

    ' Return the country owner
Public Function whoOwnesThisCountry(ctry As Integer) As Long
    whoOwnesThisCountry = gCountryOwner(ctry)
End Function

    ' Return the country score
Public Function getCountryScore(ctry As Byte) As Integer
    getCountryScore = gCtryScore(CInt(ctry))
End Function

    ' Load cards for whichPlayer into array in A3
Public Sub LoadCards(whichPlayer As Integer)
    Dim cntr As Integer
    
    For cntr = 1 To 10
        otherPlayerCards(whichPlayer, cntr) = CByte(gPlayerID(whichPlayer).card(cntr - 1))
    Next
End Sub

    ' Load cards for whichPlayer into passed list
Public Sub loadCardList(whichPlayer As Integer, cardList() As Integer)
    Dim cntr As Integer
    
    For cntr = 1 To 10
        cardList(cntr) = gPlayerID(whichPlayer).card(cntr - 1)
    Next
End Sub

'Testing: mr#onandon
'Load player defence with scores for testing.
Private Sub pd()
    Dim i As Long
    Dim x As Single
    
    x = 0.5
    For i = 1 To 6
        playerDefence(i) = x
        Debug.Print "Player "; i, playerDefence(i)
        x = x + (4.5 / 5)
        Plr(i) = 0
    Next
End Sub

'Tesing:Print country details.
Private Sub PrintMapOrder()
    Dim i As Long
    For i = 1 To 42
        Debug.Print "Ctry "; i; " = "; CountryID(i).ctryName
    Next
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Hide or show the menu bar. A bit long winded but the idea is to
    'only call the function as few times as required.
    If mnuFullScreen.Checked And Trim(mnuFile.Tag) = "" Then
        If y < 5 And Not mnuFile.Visible Then
            Call ShowMenuBar(True)
            'Debug.Print "Show"
        ElseIf y > 5 And mnuFile.Visible Then
            Call ShowMenuBar(False)
            'Debug.Print "Hide"
        End If
    End If
End Sub

