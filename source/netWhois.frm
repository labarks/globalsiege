VERSION 5.00
Begin VB.Form netWhois 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Whois"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3060
   Icon            =   "netWhois.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrWhois 
      Interval        =   500
      Left            =   120
      Top             =   1440
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox pctWhois 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "netWhois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WhoisFColor(1 To 6) As Long     'Text colour.
Dim WhoisBColor(1 To 6) As Long     'Background color, not used but required.
Dim WhoisSClor(1 To 6)  As Long     'Map score text color.

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmdClose.Top = Me.ScaleHeight - cmdClose.Height
    cmdClose.Left = Me.ScaleWidth - cmdClose.Width
    pctWhois.Left = 0
    pctWhois.Top = 0
    pctWhois.Width = Me.ScaleWidth
    pctWhois.Height = Me.ScaleHeight - cmdClose.Height
    Me.Top = netChatterBox.Top + netChatterBox.Height
    Me.Left = RiskForm1.Left + RiskForm1.Width - Me.Width
    
    'Set text colour.
    Call RiskForm1.set2Dmode(WhoisFColor, WhoisBColor, WhoisSClor)
    Call Whois
End Sub

Private Sub Whois()
    Dim ArmyNo As Integer
    
    pctWhois.Cls
    For ArmyNo = 1 To 6
        pctWhois.ForeColor = WhoisFColor(ArmyNo)
        
        'Mark the current player
        If ArmyNo = RiskForm1.playerTurn Then
            pctWhois.Font.Underline = True
        Else
            pctWhois.Font.Underline = False
        End If
        
        pctWhois.Print RiskForm1.GetArmyOrControllerName(CByte(ArmyNo))
        
        'If (RiskForm1.playerSelect_getIndex(CInt(ArmyNo - 1)) = 0) Then
        '    pctWhois.Print netMain.txtTerminalName
        'Else
        '    pctWhois.Print RiskForm1.PlayerSelect(ArmyNo - 1).Text
        'End If
    Next
End Sub

Private Sub tmrWhois_Timer()
    Call Whois
End Sub
