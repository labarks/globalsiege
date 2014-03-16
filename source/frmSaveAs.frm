VERSION 5.00
Begin VB.Form frmSaveAs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save As..."
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSaveAs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmSaveAs.frx":000C
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4395
   ScaleMode       =   0  'User
   ScaleWidth      =   3080.096
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLocked 
      Caption         =   "Locked"
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
      ToolTipText     =   "Lock war to prevent accidental deletion"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtDescription 
      Height          =   2175
      Left            =   0
      MaxLength       =   499
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCncl 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   0
      MaxLength       =   100
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.CheckBox chkDefault 
      Caption         =   "&Make this the default war"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Open this war the next time MissionRisk starts"
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Title"
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
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
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
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "frmSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gWarLocked As Boolean
Dim gWarDescription As String
Dim gWarFilePath As String
Dim gWarTitle As String

Property Get WarLocked() As Boolean
    WarLocked = gWarLocked
End Property

Property Get WarDescription() As String
    WarDescription = gWarDescription
End Property

Property Get WarTitle() As String
    WarTitle = gWarTitle
End Property

Property Get WarFilePath() As String
    WarFilePath = gWarFilePath
End Property

Private Sub cmdCncl_Click()
    Unload Me
End Sub

'Check and prepair the current war ready for saving.
Private Sub cmdSave_Click()
    Dim vWarFileName As String
    Dim wFiles As WarControlType
    Dim cntr As Long
    
    On Error Resume Next
    
    txtTitle.Text = Trim(txtTitle.Text)
    
    If txtTitle.Text = "" Then
        Exit Sub
    End If
    
    'Strip characters from the title that cannot be used as a file name.
    'The funny chars listed below can be used in file names and will not be encoded.
    vWarFileName = GetWarDataDir & "\" _
    & EncodeNonAscii(txtTitle.Text, , " $@#%^&! ~-_+='`;.") _
    & gcWarFileExtension
    
    'File with the same name already exists.
    If Dir(vWarFileName) <> "" Then
        If MsgBox(txtTitle.Text & Phrase(263) & vbCrLf & vbCrLf _
        & Phrase(264), vbYesNo, Phrase(265)) <> vbYes Then ' already exists.; Overwrite any way?
            txtTitle.SelStart = 0
            txtTitle.SelLength = Len(txtTitle.Text)
            txtTitle.SetFocus
            Exit Sub
        End If
        If IsWarFileLocked(vWarFileName) Then
            MsgBox """" & vWarFileName & """" & vbCrLf & Phrase(266)   ' cannot be modified or deleted.
            txtTitle.SelStart = 0
            txtTitle.SelLength = Len(txtTitle.Text)
            txtTitle.SetFocus
            Exit Sub
        End If
    End If
    
    'Actions if marked as the startup war.
    If chkDefault.Value Then
        SaveSetting gcApplicationName, "settings", "StartingWar", vWarFileName
    End If
    
    gWarFilePath = vWarFileName
    gWarTitle = txtTitle.Text
    gWarDescription = txtDescription.Text
    gWarLocked = chkLocked.Value
    Unload Me
End Sub

Private Sub Form_Load()
    'Call TheMainForm.ActivatePauseMode(True)
    chkLocked.Visible = gCheatMode.createMap
    frmSaveAs.Caption = Phrase(147)
    Label2.Caption = Phrase(148)
    Label1.Caption = Phrase(146)
    chkDefault.Caption = Phrase(145)
    cmdCncl.Caption = "&" + Phrase(63)
    cmdSave.Caption = Phrase(149)
    chkDefault.ToolTipText = Phrase(324) 'Open this war the next time the app starts
    chkLocked.ToolTipText = Phrase(325) 'Lock war to prevent accidental deletion
    
End Sub

Private Sub txtDescription_GotFocus()
    'Call SelectAndHighlightText(txtDescription)
End Sub

Private Sub txtTitle_GotFocus()
    Call SelectAndHighlightText(txtTitle)
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSave_Click
    End If
End Sub
