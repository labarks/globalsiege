VERSION 5.00
Begin VB.Form frmMissions 
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2940
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMissions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmMissions.frx":000C
   PaletteMode     =   2  'Custom
   ScaleHeight     =   196
   ScaleMode       =   0  'User
   ScaleWidth      =   239.37
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00000000&
      Caption         =   "Ignore"
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
      Left            =   2760
      MaskColor       =   &H00000000&
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Open"
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
      Left            =   1920
      MaskColor       =   &H00000000&
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox chkShowAgain 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblShowAgain 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Don't show any more."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   2595
      Width           =   1530
   End
   Begin VB.Label lblContinent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your Mission Briefing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   1995
      Left            =   0
      Picture         =   "frmMissions.frx":2EAE
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmMissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Left = TheMainForm.Left + ((TheMainForm.Width - Width) \ 2)
    Top = TheMainForm.Top + ((TheMainForm.Height - Height) \ 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'No info box flashing.
    TheMainForm.tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
End Sub

Private Sub chkShowAgain_Click()
    On Error Resume Next
    'No info box flashing.
    TheMainForm.tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    TheMainForm.mnuMisSeeReminder.Checked = Not CBool(Abs(chkShowAgain.Value))
    OKButton.SetFocus
End Sub

Private Sub Command1_Click()
    'No info box flashing.
    TheMainForm.tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    Me.Hide
    Call TheMainForm.mnuMissionSee_Click
End Sub

Private Sub Form_Paint()
    'lblCountry.Caption = Phrase(355)        'TOP SECRET
    On Error Resume Next 'GoTo errhand
    lblContinent.Caption = Phrase(356)      'Your mission briefing is enclosed in this document.
    Command1.Caption = Phrase(358)          'Open
    OKButton.Caption = Phrase(357)          'Ignore
    lblShowAgain.Caption = Phrase(136)      'Don't show any more.
    If TheMainForm.Visible Then
        TheMainForm.SetFocus
    End If
    Exit Sub
ErrHand:
    Unload Me
End Sub

Private Sub OKButton_Click()
    'No info box flashing.
    TheMainForm.tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    Me.Hide
End Sub

