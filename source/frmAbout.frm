VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About <Var.ExeName>"
   ClientHeight    =   4470
   ClientLeft      =   2340
   ClientTop       =   1815
   ClientWidth     =   6780
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmAbout.frx":030A
   ScaleHeight     =   3085.273
   ScaleMode       =   0  'User
   ScaleWidth      =   6366.771
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLicense 
      Height          =   3135
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmAbout.frx":A26D
      Top             =   1200
      Width           =   5055
   End
   Begin VB.CommandButton cmdLicense 
      Caption         =   "License"
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
      Left            =   5400
      TabIndex        =   5
      Top             =   840
      Width           =   1260
   End
   Begin VB.CommandButton cmdCredits 
      Caption         =   "Credits"
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
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   780
      Left            =   240
      Picture         =   "frmAbout.frx":A50B
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   2
      Top             =   240
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
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
      Height          =   345
      Left            =   5400
      TabIndex        =   0
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      Caption         =   "<Var.ExeName>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   4245
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   4245
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLicense_Click()
    On Error Resume Next
    Call frmCredits.DisplayText("License")
    frmCredits.Show vbModal
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdCredits_Click()
    On Error Resume Next
    Call frmCredits.DisplayText("Credits")
    frmCredits.Show vbModal
End Sub

Private Sub Form_Load()
    Me.Caption = SubstituteStringTokens(Me.Caption)
    lblTitle.Caption = SubstituteStringTokens(lblTitle.Caption)
    lblVersion.Caption = Phrase(197) & SubstituteStringTokens("<Var.Maj>.<Var.Min>.<Var.Rev>")
    cmdCredits.Caption = Phrase(326)  'Credits
End Sub

