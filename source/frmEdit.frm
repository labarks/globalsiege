VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "<Var.EXEName> Editor"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3225
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmEdit.frx":030A
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4680
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Graphics Options"
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   2895
      Begin VB.CommandButton cmdPaint 
         Caption         =   "Paint"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtClr 
         Height          =   285
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "255"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtClr 
         Height          =   285
         Index           =   1
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "255"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtClr 
         Height          =   285
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "255"
         Top             =   1080
         Width           =   375
      End
      Begin VB.PictureBox pctClr 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   2475
         TabIndex        =   12
         Top             =   1440
         Width           =   2535
      End
      Begin MSComctlLib.Slider sldColor 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider sldColor 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider sldColor 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickStyle       =   3
      End
      Begin VB.Label Label3 
         Caption         =   "R"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "G"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "B"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "1"
      Top             =   720
      Width           =   735
   End
   Begin VB.PictureBox pctColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox pctColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox pctColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox pctColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox pctColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox pctColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblCountry 
      Caption         =   "Country name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Units"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Occupying army"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim editCntry As Integer
Dim newOwner As Integer

Public Sub setCountry(whichCntry As Integer, ctryName As String, ctryScore As Integer, ctryOwner As Integer)
    editCntry = whichCntry
    lblCountry.Caption = ctryName
    txtScore.Text = ctryScore
    newOwner = ctryOwner
    updateColors
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Pause button.
    If KeyCode = 19 Then
        Call TheMainForm.ActivatePauseMode
    End If
End Sub

Private Sub updateColors()
    Dim cntr As Integer
    
    For cntr = 0 To 5
        pctColor(cntr).Cls
    Next
    pctColor(newOwner - 1).Print " *"
    
    txtClr(0).Text = (gPlayerID(newOwner).lngColor \ &H1) And &HFF
    txtClr(1).Text = (gPlayerID(newOwner).lngColor \ &H100) And &HFF
    txtClr(2).Text = (gPlayerID(newOwner).lngColor \ &H10000) And &HFF
    sldColor(0).Value = txtClr(0).Text
    sldColor(1).Value = txtClr(1).Text
    sldColor(2).Value = txtClr(2).Text
    UpdatePctClr
    DoEvents
End Sub

Private Sub cmdOK_Click()
    Dim ctryScore As Integer
    
    If Not IsNumeric(txtScore.Text) Then
        txtScore.Text = 1
    ElseIf CLng(txtScore.Text) > 999 Or CLng(txtScore.Text) < 1 Then
        txtScore.Text = 1
    End If
    ctryScore = CInt(txtScore.Text)
    
    Call TheMainForm.editMap(editCntry, ctryScore, newOwner)
    
    'Make sure the viewport gets refreshed.
    TheMainForm.gSyncViewportNeeded = True
    
    Call TheMainForm.SyncForgroundMap("frmEdit.cmdOK_Click")
    Unload Me
End Sub

Private Sub Form_Load()
    Dim x As Single
    Dim y As Single
    Me.Caption = SubstituteStringTokens(Me.Caption)
    'Call TheMainForm.ActivatePauseMode(True)
    Call TheMainForm.getMousePos(x, y)
    Me.Left = Abs(x - Me.Width) + TheMainForm.Left
    Me.Top = Abs(y - Me.Height) + TheMainForm.Top
    Me.Width = 2385
    Me.Height = 2025
    Me.Caption = Phrase(337)        'Global Siege Editor
    Label1.Caption = Phrase(338)    'Occupying army
    Label2.Caption = Phrase(339)    'Units
End Sub

Private Sub pctColor_Click(Index As Integer)
    On Error Resume Next
    newOwner = Index + 1
    updateColors
    txtScore.SetFocus
End Sub

Private Sub txtScore_GotFocus()
    txtScore.SelStart = 0
    txtScore.SelLength = Len(txtScore.Text)
End Sub

' *** Colour picker ***

Private Sub sldColor_Scroll(Index As Integer)
    txtClr(Index).Text = CStr(sldColor(Index).Value)
    UpdatePctClr
End Sub

' Put new colour in pctColor.
Private Sub UpdatePctClr()
    pctClr.BackColor = RGB(CInt(txtClr(0).Text), _
                            CInt(txtClr(1).Text), _
                            CInt(txtClr(2).Text))
End Sub

'Paint map new color.
Private Sub PaintMap()
    Dim cntr As Integer
    Dim seeClr As Long
    
    seeClr = RGB(CInt(txtClr(0).Text), _
                CInt(txtClr(1).Text), _
                CInt(txtClr(2).Text))
    For cntr = 1 To 42
        Call TheMainForm.ColorCountry(cntr, seeClr)
    Next cntr
    Call TheMainForm.SyncForgroundMap("frmEdit.PaintMap")
End Sub

Private Sub cmdPaint_Click()
    PaintMap
End Sub

