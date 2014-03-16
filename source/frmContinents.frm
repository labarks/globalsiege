VERSION 5.00
Begin VB.Form frmContinents 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "<Var.ExeName> Continents"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4560
   Icon            =   "frmContinents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmContinents.frx":030A
   PaletteMode     =   2  'Custom
   ScaleHeight     =   3150
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3165
      Left            =   0
      Picture         =   "frmContinents.frx":15BB
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00000000&
         Cancel          =   -1  'True
         Caption         =   "Close"
         Default         =   -1  'True
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
         Left            =   3525
         TabIndex        =   1
         Top             =   2820
         Width           =   975
      End
      Begin VB.Label lblContName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Australia 2"
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
         Index           =   5
         Left            =   3585
         TabIndex        =   7
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label lblContName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asia 7"
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
         Index           =   4
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblContName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Africa 3"
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
         Index           =   3
         Left            =   1875
         TabIndex        =   5
         Top             =   2640
         Width           =   585
      End
      Begin VB.Label lblContName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Europe 5"
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
         Index           =   2
         Left            =   1845
         TabIndex        =   4
         Top             =   480
         Width           =   645
      End
      Begin VB.Label lblContName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "South America 2"
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
         Height          =   690
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label lblContName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "North America 5"
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
         Index           =   0
         Left            =   225
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmContinents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Me.Caption = SubstituteStringTokens(Me.Caption)
    cmdClose.Caption = Phrase(336)  'Close
    Call UpdateContDetails
End Sub

'Update the continent names and values.
Public Sub UpdateContDetails()
    Dim vIndex As Long
    
    On Error Resume Next
    
    For vIndex = 0 To 5
        lblContName(vIndex).Caption = Phrase(414 + vIndex) & ": " _
                                    & CStr(TheMainForm.udContVal(vIndex).Value)
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Pause button.
    If KeyCode = 19 Then
        Call TheMainForm.ActivatePauseMode
    End If
End Sub
