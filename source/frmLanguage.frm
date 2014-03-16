VERSION 5.00
Begin VB.Form frmLanguage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Language"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   2985
   Icon            =   "frmLanguage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLang 
      Caption         =   "Phrase Index"
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Load phrases from ""phrases.txt"""
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdLang 
      Caption         =   "Load Phrase File"
      Height          =   735
      Index           =   8
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Load phrases from ""phrases.txt"""
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdLang 
      Height          =   735
      Index           =   7
      Left            =   1080
      Picture         =   "frmLanguage.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Danish"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdLang 
      Height          =   735
      Index           =   6
      Left            =   240
      Picture         =   "frmLanguage.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Norsk"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdLang 
      Height          =   735
      Index           =   5
      Left            =   1920
      Picture         =   "frmLanguage.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Svenska"
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdLang 
      Height          =   735
      Index           =   2
      Left            =   1080
      Picture         =   "frmLanguage.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Francais"
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdLang 
      Height          =   735
      Index           =   4
      Left            =   240
      Picture         =   "frmLanguage.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Espanol"
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdLang 
      Height          =   735
      Index           =   3
      Left            =   1920
      Picture         =   "frmLanguage.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Deutsch"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdLang 
      Height          =   735
      Index           =   1
      Left            =   1080
      Picture         =   "frmLanguage.frx":1C96
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Italiano"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdLang 
      Default         =   -1  'True
      Height          =   735
      Index           =   0
      Left            =   240
      Picture         =   "frmLanguage.frx":20D8
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "English"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Please select a language."
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2895
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancle_Click()
    Unload Me
End Sub

'Change language and update all the phrases.
Public Sub Translate(pNewLang As Long)

    SaveSetting gcApplicationName, "settings", "Lang", pNewLang
    gLanguage = pNewLang
    Call LoadPhrases
    On Error Resume Next
    Unload Me
End Sub

'If Cheatmode.Testing, Language 4 prints phrase numbers to the screen.
Private Sub cmdLang_Click(Index As Integer)
    
    Select Case Index
    
    'English
    Case 0
        Me.Hide
        TheMainForm.Refresh
        Call Translate(eLanguage.English)
    
    'Italian
    Case 1
        Me.Hide
        TheMainForm.Refresh
        Call Translate(eLanguage.Italian)
    
    'French
    Case 2
        Me.Hide
        TheMainForm.Refresh
        Call Translate(eLanguage.French)
    
    'German
    Case 3
        Me.Hide
        TheMainForm.Refresh
        Call Translate(eLanguage.German)
    
    'Spanish
    Case 4
        Me.Hide
        TheMainForm.Refresh
        Call Translate(eLanguage.Spanish)
    
    'Swedish
    Case 5
        Me.Hide
        TheMainForm.Refresh
        Call Translate(eLanguage.Swedish)
    
    'Norwegian
    Case 6
        Me.Hide
        TheMainForm.Refresh
        Call Translate(eLanguage.Norwegian)
    
    'Danish
    Case 7
        Me.Hide
        TheMainForm.Refresh
        Call Translate(eLanguage.Danish)
    
    'Read phrase file
    Case 8
        Me.Hide
        TheMainForm.Refresh
        Call Translate(eLanguage.PhraseFile)
    
    'Show phrase numbers instead of phrases.
    Case 9
        Me.Hide
        TheMainForm.Refresh
        Call Translate(eLanguage.PhraseNumbers)
    
    'Default to English.
    Case Else
        Call Translate(eLanguage.English)
    End Select
    
End Sub

Private Sub Form_Load()
    If Dir(App.Path & "\" & gcPhraseFileName) <> "" Then
        'Rearange buttons and show file button.
        'cmdLang(6).Left = cmdLang(0).Left
        'cmdLang(7).Left = cmdLang(1).Left
        cmdLang(8).Visible = True
    Else
        'Hide file button.
        cmdLang(8).Visible = False
    End If
    
    'Show phrase index button if in testing mode.
    cmdLang(9).Visible = gCheatMode.testing
    
    Label1.Caption = "- Please select a language." + vbCrLf _
        + "- Selezionare una lingua." + vbCrLf _
        + "- Wählen Sie bitte eine Sprache aus." + vbCrLf _
        + "- Por favor seleccione un idioma." + vbCrLf _
        + "- S'il vous plaît sélectionnez un langage." + vbCrLf _
        + "- Var god välj språk." + vbCrLf _
        + "- Vennligst velg språk."
End Sub

'Save selected language to registry. Default to English.
Private Sub Form_Unload(Cancel As Integer)
    Dim vLanguage As String
    
    vLanguage = GetSetting(gcApplicationName, "settings", "Lang", "No lang selected")
    If IsNumeric(vLanguage) Then
        gLanguage = CLng(vLanguage) And &HFF
    Else
        gLanguage = eLanguage.English
    End If
    SaveSetting gcApplicationName, "settings", "Lang", CStr(gLanguage)
End Sub
