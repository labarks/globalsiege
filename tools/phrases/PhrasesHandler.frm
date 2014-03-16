VERSION 5.00
Begin VB.Form frmPhraseHandler 
   Caption         =   "Phrase Editor"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8430
   Icon            =   "PhrasesHandler.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   360
      TabIndex        =   6
      Top             =   0
      Width           =   1935
      Begin VB.CheckBox chkPaste 
         Caption         =   "Auto Paste English to clipboard"
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox varList 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Variables"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   7455
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   1575
         TabIndex        =   24
         Top             =   120
         Width           =   1575
         Begin VB.OptionButton optEnglish 
            Caption         =   "English"
            Height          =   195
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optRemark 
            Caption         =   "Remark"
            Height          =   195
            Left            =   0
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox text1 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Text            =   "PhrasesHandler.frx":0442
         Top             =   720
         Width           =   6975
      End
      Begin VB.TextBox Text2 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "PhrasesHandler.frx":0448
         Top             =   2640
         Width           =   6975
      End
      Begin VB.Label Label2 
         Caption         =   "English"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Italian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   2280
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   5175
         TabIndex        =   10
         Top             =   240
         Width           =   5175
         Begin VB.CommandButton cmdEnter 
            Caption         =   "Enter"
            Height          =   735
            Left            =   1920
            TabIndex        =   22
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "-"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   21
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "+"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   20
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtCount 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   0
            TabIndex        =   19
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton optLanguage 
            Caption         =   "Spanish"
            Height          =   195
            Index           =   3
            Left            =   2400
            TabIndex        =   18
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optLanguage 
            Caption         =   "German"
            Height          =   195
            Index           =   2
            Left            =   1560
            TabIndex        =   17
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optLanguage 
            Caption         =   "French"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   16
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optLanguage 
            Caption         =   "Italian"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.CommandButton cmdLast 
            Caption         =   "Last entry"
            Height          =   735
            Left            =   1320
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
         Begin VB.OptionButton optLanguage 
            Caption         =   "Swedish"
            Height          =   195
            Index           =   4
            Left            =   3480
            TabIndex        =   13
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optLanguage 
            Caption         =   "Norwegian"
            Height          =   195
            Index           =   5
            Left            =   4080
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optLanguage 
            Caption         =   "Dainish"
            Height          =   195
            Index           =   6
            Left            =   4080
            TabIndex        =   11
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Phrase number"
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu brk1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInc 
         Caption         =   "&Increment when entered"
         Checked         =   -1  'True
         Shortcut        =   ^I
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOut 
         Caption         =   "&Out put to phrases.txt"
      End
      Begin VB.Menu mnuIn 
         Caption         =   "&Read phrases.txt"
      End
   End
End
Attribute VB_Name = "frmPhraseHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************
'Phrase handler/editor for MissionRisk
'Doug      15-10-1999
'
'Updated to handle
'more languages     25-10-1999
'***************************************

    'Token for phrase part of file
Const phraseToken As String = "Private Sub InitialisePhrases()"
    'Token to start new phrase
Const SelPhrsToken As String = "sPhrase"

Private Type language
    Remark As String            '"'REM:" stetements
    lang(7) As String           'English=0, Italian=1, etc...
End Type

Dim sourcePath As String        'Path to source file
Dim phraseFileName As String    'Source file name
Dim FirstPart As String         'Code part of source file
Dim phrasePart As String        'Phrase part of source file
Dim phrase() As language        'Individual phrases
Dim textChanged As Boolean      'True if phrase text has changed
Dim fileChanged As Boolean      'True if phrases have been changed
Dim textLock As Boolean         'Stop callback for phrase number text box
Dim lastActiveLanguage As Long  'Keep track of language changes


    'Increment/decrement phrase pointer and put new phrase in text boxes
    'If Index is -1 then goto last phrase
Private Sub cmdChange_Click(Index As Integer)
    Dim pointer As Long
    
    If Not IsNumeric(txtCount.Text) Then Exit Sub
    textLock = True
    pointer = CLng(txtCount.Text)
    If textChanged Then
        If MsgBox("Enter change?", vbYesNo) = vbYes Then
            Call cmdEnter_Click
        End If
    End If
    
    If Index = 0 Then                       '+ key
        pointer = pointer + 1
    ElseIf Index = 1 Then                   '- key
        pointer = pointer - 1
    ElseIf Index = -1 Then                  'Last entry
        pointer = UBound(phrase) - 1
    End If
    If pointer < 0 Or pointer >= UBound(phrase) Then
        Exit Sub                            'Out of bounds
    End If
    
    txtCount.Text = CStr(pointer)
    If optEnglish.Value Then
        text1.Text = phrase(pointer).lang(0)
    Else
        text1.Text = phrase(pointer).Remark
    End If
    Text2.Text = phrase(pointer).lang(activeLanguage)
    Call pastePhrase
    textChanged = False
    Text2.SetFocus
    textLock = False
End Sub

    'Goto last phrase
Private Sub cmdLast_Click()
    Call cmdChange_Click(-1)
End Sub

Private Sub Form_Load()
    Dim fileContents As String
    
    'varList.AddItem "<Var.EXEName>"
    'varList.AddItem "<Var.Min>"
    'varList.AddItem "<Var.Maj>"
    'varList.AddItem "<Var.Eval>"
    
    'sourcePath = "C:\WINDOWS\Desktop\Risk development\MR source code\"
    sourcePath = App.Path & "\"
    phraseFileName = "Phrases.bas"
    'phraseFileName = "tst.bas"
    fileContents = seeFile(sourcePath & phraseFileName)
    
    Me.Show
    Call breakString(fileContents)
    If optEnglish.Value Then
        text1.Text = phrase(CLng(txtCount.Text)).lang(0)
    Else
        text1.Text = phrase(CLng(txtCount.Text)).Remark
    End If
    Text2.Text = phrase(CLng(txtCount.Text)).lang(activeLanguage)
    Call pastePhrase
    textChanged = False
    textLock = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Answer As Long
    
    If fileChanged Or textChanged Then
        Answer = MsgBox("Save changes?", vbYesNoCancel)
        If Answer = vbYes Then
            Call mnuSave_Click
        ElseIf Answer = vbCancel Then
            Cancel = -1
        End If
    End If
End Sub

    'Read German from "c:/windows/desktop/MissionRisk phrases.txt"
Private Sub mnuIn_Click()
    Dim i As Long
    Dim strInPut As String
    Dim tmp As String
    
    Open App.Path & "/phrases.txt" For Binary As #1
    strInPut = Space(LOF(1))
    Get #1, , strInPut
    Close #1
    'Debug.Print strInPut
    
    Call BreakTranslatedString(strInPut)
End Sub

    'Auto incrment when enter pressed
Private Sub mnuInc_Click()
    mnuInc.Checked = Not mnuInc.Checked
End Sub

    'Put english to "c:/windows/desktop/MissionRisk phrases.txt"
Private Sub mnuOut_Click()
    Dim i As Long
    Dim outPut As String
    
    outPut = ""
    For i = 0 To UBound(phrase)
        outPut = outPut & "~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
        
        If Len(Trim(phrase(i).Remark)) > 0 Then
            outPut = outPut & "Remark " & CStr(i) & ": " & Chr(34) & phrase(i).Remark & Chr(34) & vbCrLf & vbCrLf
        End If
        outPut = outPut & "[Phrase <" & Trim(CStr(i)) & "> - English]" & vbCrLf
        'outPut = outPut & "[" & Trim(CStr(i)) & ".]" & vbCrLf
        outPut = outPut & Chr(34) & Replace(phrase(i).lang(0), "&", "") & Chr(34) & vbCrLf & vbCrLf
        
        'outPut = outPut & "[Phrase <" & Trim(CStr(i)) & "> - German]" & vbCrLf & Chr(34) & " " & Chr(34) & vbCrLf & vbCrLf
        'outPut = outPut & "[Phrase <" & Trim(CStr(i)) & "> - Norwegian]" & vbCrLf & Chr(34) & "" & Chr(34) & vbCrLf & vbCrLf
        'outPut = outPut & "[Phrase <" & Trim(CStr(i)) & "> - Italian]" & vbCrLf
        'outPut = outPut & "[Phrase <" & Trim(CStr(i)) & "> - French]" & vbCrLf & Chr(34) & Replace(phrase(i).lang(2), "&", "") & Chr(34) & vbCrLf & vbCrLf
        'outPut = outPut & "[Phrase <" & Trim(CStr(i)) & "> - Translation]" & vbCrLf & Chr(34) & Replace(phrase(i).lang(0), "&", "") & Chr(34) & vbCrLf & vbCrLf
        outPut = outPut & "[Phrase <" & Trim(CStr(i)) & "> - Translation]" & vbCrLf & Chr(34) & Chr(34) & vbCrLf & vbCrLf

        'outPut = outPut & "[Phrase <" & Trim(CStr(i)) & "> - Spanish]" & vbCrLf
        'outPut = outPut & Chr(34) & phrase(i).lang(4) & Chr(34) & vbCrLf & vbCrLf
        
        'outPut = outPut & "[Phrase <" & Trim(CStr(i)) & "> - German]" & vbCrLf
        'outPut = outPut & Chr(34) & phrase(i).lang(3) & Chr(34) & vbCrLf & vbCrLf
    Next
    
    Clipboard.SetText outPut
    'Open "c:/windows/desktop/MissionRisk phrases.txt" For Binary As #1
    Open sourcePath & "MissionRisk phrases.txt" For Binary As #1
    Put #1, , outPut
    Close #1
End Sub

    'Save phrases to source file
Private Sub mnuSave_Click()
    Dim cntr1 As Long
    Dim outPut As String
    
    If textChanged Then
        If MsgBox("Enter change to current phrase?", vbYesNo) = vbYes Then
            Call cmdEnter_Click
        End If
    End If
    
    outPut = FirstPart
    For cntr1 = 0 To UBound(phrase) - 2
        outPut = outPut & AssembleCode(cntr1)
    Next
    outPut = outPut & vbCrLf & "End Sub     'Do not put code beyond this point!!!" & vbCrLf
    Call saveFile(sourcePath & phraseFileName, outPut)
End Sub

    'Save string "Contents" to file "destFile"
Private Sub saveFile(destFile As String, Contents As String)
    Dim Spaces As String
    
    Open destFile For Binary As #1
    If LOF(1) > Len(Contents) Then
        Spaces = Space(LOF(1) - Len(Contents))
    End If
    Put #1, , Contents & Spaces
    Close #1
    fileChanged = False
End Sub

    'Return a VB readable string made up of phrases(cntr1):
    'SelPhrs nnn, "English",
    '           & "Italian"
Private Function AssembleCode(nnn As Long) As String
    Dim Str As String
    
    Str = ""
    If nnn > 1 And nnn Mod 100 = 0 Then
        Str = Str & "    call initialisePhrases" & CStr(nnn) & vbCrLf & "End Sub"
        Str = Str & vbCrLf & vbCrLf & "Private Sub initialisePhrases" & CStr(nnn) & "()"
    End If
    
    Str = Str & vbCrLf & "    'REM:" & AssembleRem(phrase(nnn).Remark)
    Str = Str & vbCrLf & "    sPhraseEng " & CStr(nnn) & ", " _
    & AssembleString(phrase(nnn).lang(0))
    
    Str = Str & vbCrLf & "    sPhraseIta " & CStr(nnn) & ", " _
    & AssembleString(phrase(nnn).lang(1))
    
    Str = Str & vbCrLf & "    sPhraseFra " & CStr(nnn) & ", " _
    & AssembleString(phrase(nnn).lang(2))
    
    Str = Str & vbCrLf & "    sPhraseGer " & CStr(nnn) & ", " _
    & AssembleString(phrase(nnn).lang(3))
    
    Str = Str & vbCrLf & "    sPhraseSpa " & CStr(nnn) & ", " _
    & AssembleString(phrase(nnn).lang(4))
    
    Str = Str & vbCrLf & "    sPhraseSwe " & CStr(nnn) & ", " _
    & AssembleString(phrase(nnn).lang(5))
    
    Str = Str & vbCrLf & "    sPhraseNor " & CStr(nnn) & ", " _
    & AssembleString(phrase(nnn).lang(6))
    
    Str = Str & vbCrLf & "    sPhraseDan " & CStr(nnn) & ", " _
    & AssembleString(phrase(nnn).lang(7))
    
    AssembleCode = Str & vbCrLf
End Function

    'Replace variables with VB code and c/returns with " + VbCrLf _"
Private Function AssembleString(ByVal Str As String) As String
    Str = Chr(34) & Str & Chr(34)
    
    Str = Replace(Str, vbCrLf, Chr(34) & " + vbCrLf _" & vbCrLf & "                 + " & Chr(34))
    Str = Replace(Str, "<Var.EXEName>", Chr(34) & " & Trim(App.EXEName) & " & Chr(34))
    Str = Replace(Str, "<Var.Min>", Chr(34) & " + str(App.Minor) + " & Chr(34))
    Str = Replace(Str, "<Var.Maj>", Chr(34) & " + str(App.Major) + " & Chr(34))
    Str = Replace(Str, "<Var.Eval>", Chr(34) & " + Trim(CStr(evaluationPeriod)) + " & Chr(34))
    
    AssembleString = Str
End Function

    'Look after line continuations in remarks
Private Function AssembleRem(ByVal Str As String) As String
    AssembleRem = Replace(Str, vbCrLf, " _" & vbCrLf & "     ")
End Function

    'Return the contents of file (fpath)
Private Function seeFile(fPath As String) As String
    Dim strB As String

    Open fPath For Binary As #1
    strB = Space(LOF(1))
    Get #1, , strB
    Close #1
    seeFile = strB
    fileChanged = False
End Function

Private Sub BreakTranslatedString(fileContents As String)
    Dim where As Long
    Dim startPhrase As Long
    Dim endPhrase As Long
    Dim startComment As Long
    Dim totalPhrase As Long
    Dim cntr As Long
    
    where = 1
    phrasePart = Mid(fileContents, where)
    
    'Count total phrases by counting how many English phrases there are
    where = 1
    totalPhrase = UBound(phrase)
    
    'Extract Italian
    where = 1
    For cntr = 0 To totalPhrase - 1
        startPhrase = where
        'where = InStr(where + 2, phrasePart, "> - Italian")
        'where = InStr(where + 2, phrasePart, "> - Spanish")
        'where = InStr(where + 2, phrasePart, "> - Swedish")
        'where = InStr(where + 2, phrasePart, "> - Norwegian")
        'where = InStr(where + 2, phrasePart, "> - French")
        where = InStr(where + 2, phrasePart, "> - Translation")
        
        If where = 0 Then
            Exit For
        End If
        startPhrase = InStr(where, phrasePart, Chr(34)) + 1            'Point to first quote marks (")
        endPhrase = InStr(startPhrase, phrasePart, Chr(34))   'Start of next language's phrase
        If endPhrase = 0 Then
            Exit For
        End If
        phrase(cntr).lang(7) = Mid(phrasePart, startPhrase, endPhrase - startPhrase)
    Next
End Sub

    'Seperate file contents into firstPart, phrasePart and
    'read phrases from phrasePart
Private Sub breakString(fileContents As String)
    Dim where As Long
    Dim startPhrase As Long
    Dim endPhrase As Long
    Dim startComment As Long
    Dim totalPhrase As Long
    Dim cntr As Long
    
    where = InStr(fileContents, phraseToken) + Len(phraseToken)
    FirstPart = Mid(fileContents, 1, where - 1)
    phrasePart = Mid(fileContents, where)
    
    'Count total phrases by counting how many English phrases there are
    where = 1
    totalPhrase = 0
    Do
        totalPhrase = totalPhrase + 1
        where = InStr(where + 2, phrasePart, "sPhraseEng")
    Loop While where <> 0
    
    totalPhrase = totalPhrase - 1
    ReDim phrase(totalPhrase + 1)
    
    phrasePart = Replace(phrasePart, "'Rem:", "'REM:")              'Make all upper case
    phrasePart = Replace(phrasePart, "'rem:", "'REM:")
    
    'Extract English and comments
    where = 1
    For cntr = 0 To totalPhrase - 1
        startPhrase = where
        where = InStr(where + 2, phrasePart, "sPhraseEng")
        startComment = InStr(startPhrase, phrasePart, "'REM:") + 5
        If startComment < where And startComment > 5 Then           'Extract comment
            phrase(cntr).Remark = getRem(Mid(phrasePart, startComment, where - startComment - 6))
        End If
        
        'Extract English
        startPhrase = InStr(where, phrasePart, Chr(34))             'Point to first quote marks (")
        endPhrase = InStr(startPhrase, phrasePart, SelPhrsToken)    'Start of next language's phrase
        If endPhrase = 0 Then
            endPhrase = InStr(startPhrase, phrasePart, "End Sub")
        End If
        phrase(cntr).lang(0) = getPhrase(Mid(phrasePart, startPhrase, endPhrase - startPhrase))
    Next
    
    'Extract Italian
    where = 1
    For cntr = 0 To totalPhrase - 1
        startPhrase = where
        where = InStr(where + 2, phrasePart, "sPhraseIta")
        If where = 0 Then
            Exit For
        End If
        startPhrase = InStr(where, phrasePart, Chr(34))             'Point to first quote marks (")
        endPhrase = InStr(startPhrase, phrasePart, SelPhrsToken)    'Start of next language's phrase
        If endPhrase = 0 Then
            endPhrase = InStr(startPhrase, phrasePart, "End Sub")
        End If
        phrase(cntr).lang(1) = getPhrase(Mid(phrasePart, startPhrase, endPhrase - startPhrase))
    Next
    
    'Extract French
    where = 1
    For cntr = 0 To totalPhrase - 1
        startPhrase = where
        where = InStr(where + 2, phrasePart, "sPhraseFra")
        If where = 0 Then
            Exit For
        End If
        startPhrase = InStr(where, phrasePart, Chr(34))             'Point to first quote marks (")
        endPhrase = InStr(startPhrase, phrasePart, SelPhrsToken)    'Start of next language's phrase
        If endPhrase = 0 Then
            endPhrase = InStr(startPhrase, phrasePart, "End Sub")
        End If
        phrase(cntr).lang(2) = getPhrase(Mid(phrasePart, startPhrase, endPhrase - startPhrase))
    Next
    
    'Extract German
    where = 1
    For cntr = 0 To totalPhrase - 1
        startPhrase = where
        where = InStr(where + 2, phrasePart, "sPhraseGer")
        If where = 0 Then
            Exit For
        End If
        startPhrase = InStr(where, phrasePart, Chr(34))             'Point to first quote marks (")
        endPhrase = InStr(startPhrase, phrasePart, SelPhrsToken)    'Start of next language's phrase
        If endPhrase = 0 Then
            endPhrase = InStr(startPhrase, phrasePart, "End Sub")
        End If
        phrase(cntr).lang(3) = getPhrase(Mid(phrasePart, startPhrase, endPhrase - startPhrase))
    Next
    
    'Extract Spanish
    where = 1
    For cntr = 0 To totalPhrase - 1
        startPhrase = where
        where = InStr(where + 2, phrasePart, "sPhraseSpa")
        If where = 0 Then
            Exit For
        End If
        startPhrase = InStr(where, phrasePart, Chr(34))             'Point to first quote marks (")
        endPhrase = InStr(startPhrase, phrasePart, SelPhrsToken)    'Start of next language's phrase
        If endPhrase = 0 Then
            endPhrase = InStr(startPhrase, phrasePart, "End Sub")
        End If
        phrase(cntr).lang(4) = getPhrase(Mid(phrasePart, startPhrase, endPhrase - startPhrase))
    Next
    
    'Extract Swedish
    where = 1
    For cntr = 0 To totalPhrase - 1
        startPhrase = where
        where = InStr(where + 2, phrasePart, "sPhraseSwe")
        If where = 0 Then
            Exit For
        End If
        startPhrase = InStr(where, phrasePart, Chr(34))             'Point to first quote marks (")
        endPhrase = InStr(startPhrase, phrasePart, SelPhrsToken)    'Start of next language's phrase
        If endPhrase = 0 Then
            endPhrase = InStr(startPhrase, phrasePart, "End Sub")
        End If
        phrase(cntr).lang(5) = getPhrase(Mid(phrasePart, startPhrase, endPhrase - startPhrase))
    Next
    
    'Extract Norwegian
    where = 1
    For cntr = 0 To totalPhrase - 1
        startPhrase = where
        where = InStr(where + 2, phrasePart, "sPhraseNor")
        If where = 0 Then
            Exit For
        End If
        startPhrase = InStr(where, phrasePart, Chr(34))             'Point to first quote marks (")
        endPhrase = InStr(startPhrase, phrasePart, SelPhrsToken)    'Start of next language's phrase
        If endPhrase = 0 Then
            endPhrase = InStr(startPhrase, phrasePart, "End Sub")
        End If
        phrase(cntr).lang(6) = getPhrase(Mid(phrasePart, startPhrase, endPhrase - startPhrase))
    Next
    
    'Extract Dainish
    where = 1
    For cntr = 0 To totalPhrase - 1
        startPhrase = where
        where = InStr(where + 2, phrasePart, "sPhraseDan")
        If where = 0 Then
            Exit For
        End If
        startPhrase = InStr(where, phrasePart, Chr(34))             'Point to first quote marks (")
        endPhrase = InStr(startPhrase, phrasePart, SelPhrsToken)    'Start of next language's phrase
        If endPhrase = 0 Then
            endPhrase = InStr(startPhrase, phrasePart, "End Sub")
        End If
        phrase(cntr).lang(7) = getPhrase(Mid(phrasePart, startPhrase, endPhrase - startPhrase))
    Next
End Sub

    'Replace "vbCrLf" and VB code then return remaing string between quotes
Private Function getPhrase(Str As String) As String
    Dim Qstart As Long
    Dim Qend As Long
    Dim Qnext As Long
    
    Str = Replace(Str, "vbCrLf", Chr(34) & vbCrLf & Chr(34))
    'Str = Replace(Str, "Trim(App.EXEName)", Chr(34) & "<Var.EXEName>" & Chr(34))
    'Str = Replace(Str, "str(App.Minor)", Chr(34) & "<Var.Min>" & Chr(34))
    'Str = Replace(Str, "str(App.Major)", Chr(34) & "<Var.Maj>" & Chr(34))
    'Str = Replace(Str, "Trim(CStr(evaluationPeriod))", Chr(34) & "<Var.Eval>" & Chr(34))
    
    getPhrase = ""
    Qend = 1
    Do
        Qstart = InStr(Qend, Str, Chr(34))                  'Open quote
        Qend = InStr(Qstart + 1, Str, Chr(34)) + 1          'Close qoute
        getPhrase = getPhrase & Mid(Str, Qstart + 1, Qend - Qstart - 2) 'Read between quotes
    Loop While InStr(Qend, Str, Chr(34)) <> 0
End Function

    'Look after line continuations in remarks
Private Function getRem(Str As String) As String
    Str = Replace(Str, " _" & vbCrLf, vbCrLf)
    Str = Replace(Str, "  ", "")
    getRem = Replace(Str, vbCrLf & " ", vbCrLf)
End Function

    'Enter phrase from text box into phrase array.
    'Increment to next phrase if "Increment" checked
    'Create room for a new phrase if last element
Private Sub cmdEnter_Click()
    Dim pointer As Long
    Dim tmp As Long
    
    If Not IsNumeric(txtCount.Text) Then
        Exit Sub
    End If
    pointer = CLng(txtCount.Text)
    tmp = repeatedPhrase(pointer, text1.Text)           'Check to see if phrase repeated
    
    If tmp >= 0 Then
        If MsgBox("Phrase " & CStr(tmp) & " already says that." & vbCrLf _
        & "Do you want to continue?", vbYesNo, "Repeated Phrase") = vbNo Then
            text1.Text = phrase(pointer).lang(0)
            Text2.Text = phrase(pointer).lang(activeLanguage)
            textChanged = False
            textLock = False
            Call pastePhrase(tmp)
            Exit Sub
        End If
    End If
    
    If optEnglish.Value Then
        phrase(CLng(txtCount.Text)).lang(0) = text1.Text
    Else
        phrase(pointer).Remark = text1.Text
    End If
    
    If text1.Text <> "" And Text2.Text = "" Then        'Use English phrase if blank
        'phrase(CLng(txtCount.Text)).lang(1) = "**" & text1.Text
        phrase(CLng(txtCount.Text)).lang(activeLanguage) = Text2.Text
    Else
        phrase(CLng(txtCount.Text)).lang(activeLanguage) = Text2.Text
    End If

    textChanged = False
    fileChanged = True
    If CLng(txtCount.Text) >= UBound(phrase) - 1 Then   'Make room for next phrase if required
        ReDim Preserve phrase(UBound(phrase) + 1)
    End If
    If mnuInc.Checked Then                              'Point to next phrase if Increment checked
        txtCount.Text = CStr(pointer + 1)
    End If
    Call pastePhrase
    Text2.SetFocus
End Sub

    'Paste 'Phrase(phraseNumber)' to clipboard if checked
Private Sub pastePhrase(Optional phraseNumber As Long = 0)
    If chkPaste.Value = 0 Then
        Exit Sub
    End If
    Clipboard.Clear
    Clipboard.SetText text1.Text
    'Clipboard.SetText "Phrase(" & Trim(CStr(phraseNumber)) & ")"
End Sub


    'Make sure phrase is not repeated
Private Function repeatedPhrase(pointer As Long, Str As String) As Long
    Dim cntr1 As Long
    
    If Not optEnglish.Value Then
        repeatedPhrase = -1
        Exit Function
    End If
    For cntr1 = 0 To UBound(phrase) - 1
        If phrase(cntr1).lang(0) = Str And cntr1 <> pointer Then
            repeatedPhrase = cntr1
            Exit Function
        End If
    Next
    repeatedPhrase = -1
End Function

Private Sub optEnglish_Click()
    text1.Text = phrase(CLng(txtCount.Text)).lang(0)
    textChanged = False
End Sub

Private Sub optRemark_Click()
    text1.Text = phrase(CLng(txtCount.Text)).Remark
    textChanged = False
End Sub

Private Sub optLanguage_Click(Index As Integer)
    If textChanged Then
        If MsgBox("Enter change?", vbYesNo) = vbYes Then
            phrase(CLng(txtCount.Text)).lang(lastActiveLanguage) = Text2.Text
        End If
    End If
    Label3.Caption = optLanguage(Index).Caption
    Text2.Text = phrase(CLng(txtCount.Text)).lang(Index + 1)
    textChanged = False
    textLock = False
    Text2.SetFocus
End Sub

Private Sub optLanguage_GotFocus(Index As Integer)
    lastActiveLanguage = activeLanguage
End Sub

    'Keep track of changes
Private Sub text1_Change()
    textChanged = True
End Sub
Private Sub Text2_Change()
    textChanged = True
End Sub

    'Update phrase if not locked out
Private Sub txtCount_Change()
    Dim pointer As Long
    If textLock Then Exit Sub
    
    On Error GoTo errhand
    If Not IsNumeric(txtCount.Text) Then Exit Sub
    pointer = CLng(txtCount.Text)
    textLock = True
    If textChanged Then
        If MsgBox("Enter change?", vbYesNo) = vbYes Then
            Call cmdEnter_Click
        End If
    End If
    
    If pointer < 0 Or pointer >= UBound(phrase) Then
        textLock = False
        Exit Sub
    End If
    
    If optEnglish.Value Then
        text1.Text = phrase(pointer).lang(0)
    Else
        text1.Text = phrase(pointer).Remark
    End If
    Text2.Text = phrase(pointer).lang(activeLanguage)
    textChanged = False
    textLock = False
    Exit Sub
errhand:
    textLock = False
    Exit Sub
End Sub

    'Return which language is currently active
Private Function activeLanguage() As Long
    activeLanguage = optLanguage(0).Value _
                    + optLanguage(1).Value * 2 _
                    + optLanguage(2).Value * 3 _
                    + optLanguage(3).Value * 4 _
                    + optLanguage(4).Value * 5 _
                    + optLanguage(5).Value * 6 _
                    + optLanguage(6).Value * 7
    activeLanguage = Abs(activeLanguage)
End Function

