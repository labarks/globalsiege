VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "<Var.ExeName> Credits & License"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5085
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtCredits 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Display credits or license text.
'pTextType can be "Credits" or "License".
Public Sub DisplayText(Optional pTextType As String = "Credits")
    On Error Resume Next
    Select Case LCase(pTextType)
    Case "credits"
        Call ShowCredits
    Case "license"
        Call ShowLicense
    End Select
End Sub

' Below is only needed when VB6's SP6 is not installed.
'Private Sub txtCredits_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 67 And Shift = vbCtrlMask Then
'       Clipboard.Clear
'       Clipboard.SetText txtCredits.SelText
'   End If
'End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

' Print a line of text.
Private Sub PutLine(Optional lineOfText As String)
    txtCredits.Text = txtCredits.Text & lineOfText & vbCrLf
End Sub

'Display credits.
Private Sub ShowCredits()
    Me.Caption = SubstituteStringTokens("<Var.ExeName> Credits")
    txtCredits.Text = ""
    
    PutLine "DEVELOPMENT:"
    PutLine "     Doug Burner"
    PutLine
    PutLine "PRE RELEASE TESTING AND ADVICE:"
    PutLine "     Gary Wilson"
    PutLine "     Gary Ryan"
    PutLine "     Gary Whitehead"
    PutLine "     Lars Monsees"
    PutLine "     Adrian Carter"
    PutLine "     Robin Deal"
    PutLine "     Steve Colbourne"
    PutLine
    PutLine "LANGUAGE TRANSLATIONS:"
    PutLine "     Mauro Castelnuovo"
    PutLine "     Lars Monsees"
    PutLine "     Paolo Pelloni"
    PutLine "     Gabriel A. Marturano"
    PutLine "     Carl-Fredrik Bergdahl"
    PutLine "     Gunnar Baardsen"
    PutLine "     Walfroy Ract-madoux"
    PutLine "     Simon Kofod"
    PutLine
    'PutLine "BUG REPORTS"
    'PutLine "     Timothy Steele"
    'PutLine
    PutLine "ARTWORK:"
    PutLine "     Doug Burner"
    PutLine "     John O'Brien (http://members.upnaway.com/~obees/soldiers/)"
End Sub

'Read license file in the program's home directory and display.
Private Sub ShowLicense()
    Me.Caption = SubstituteStringTokens("<Var.ExeName> License")
    txtCredits.Text = LoadProgramFile("License.txt")
End Sub

