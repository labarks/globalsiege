VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form netChatterBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Chat"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   3165
   Icon            =   "frmChatterBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmChatterBox.frx":030A
   PaletteMode     =   2  'Custom
   ScaleHeight     =   3615
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pctSendHide 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   700
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   3375
      TabIndex        =   14
      Top             =   1095
      Width           =   3375
      Begin VB.TextBox txtWrite 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000016&
         Height          =   390
         Left            =   0
         MaxLength       =   200
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         ToolTipText     =   "Type your message and press enter to send"
         Top             =   0
         Width           =   3135
      End
      Begin VB.CommandButton cmdHide 
         BackColor       =   &H80000007&
         Caption         =   "&Hide"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         MaskColor       =   &H8000000F&
         TabIndex        =   17
         Top             =   400
         Width           =   1095
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H80000007&
         Caption         =   "&Send"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         MaskColor       =   &H8000000F&
         TabIndex        =   16
         Top             =   400
         Width           =   975
      End
      Begin VB.CheckBox chkOptions 
         Alignment       =   1  'Right Justify
         Caption         =   "&Options < <"
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
         TabIndex        =   15
         Top             =   400
         UseMaskColor    =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Timer tmrWhois 
      Interval        =   500
      Left            =   2040
      Top             =   2400
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1931
      _Version        =   393217
      BackColor       =   -2147483641
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChatterBox.frx":31AC
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
   Begin VB.Frame frmOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Select recipient(s) of private messages"
      Top             =   1795
      Width           =   3135
      Begin VB.PictureBox pctVoteNudge 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   3615
         TabIndex        =   9
         Top             =   1500
         Width           =   3615
         Begin VB.CheckBox chkChime 
            BackColor       =   &H80000007&
            Caption         =   "Squelch"
            ForeColor       =   &H80000016&
            Height          =   255
            Left            =   2160
            TabIndex        =   12
            ToolTipText     =   "Beep when message received from remote terminal."
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton cmdVote 
            BackColor       =   &H80000007&
            Caption         =   "&Vote"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            MaskColor       =   &H8000000F&
            TabIndex        =   11
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdNudge 
            BackColor       =   &H80000007&
            Caption         =   "&Beep"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            MaskColor       =   &H8000000F&
            TabIndex        =   10
            ToolTipText     =   "Send alert tone to recipient terminals."
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblSquelch 
            BackColor       =   &H80000007&
            Caption         =   "Squelch"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   2400
            TabIndex        =   13
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.PictureBox pctWhois 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   225
         ScaleHeight     =   1440
         ScaleWidth      =   2850
         TabIndex        =   8
         Top             =   0
         Width           =   2910
      End
      Begin VB.CheckBox chkRecipient 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   7
         ToolTipText     =   "Send message to terminal controlling the red army"
         Top             =   0
         Width           =   225
      End
      Begin VB.CheckBox chkRecipient 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "Send message to terminal controlling the green army"
         Top             =   240
         Width           =   225
      End
      Begin VB.CheckBox chkRecipient 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "Send message to terminal controlling the blue army"
         Top             =   480
         Width           =   225
      End
      Begin VB.CheckBox chkRecipient 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   4
         ToolTipText     =   "Send message to terminal controlling the yellow army"
         Top             =   720
         Width           =   225
      End
      Begin VB.CheckBox chkRecipient 
         BackColor       =   &H00FF00FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Send message to terminal controlling the purple army"
         Top             =   960
         Width           =   225
      End
      Begin VB.CheckBox chkRecipient 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Send message to terminal controlling the gray army"
         Top             =   1200
         Width           =   225
      End
   End
End
Attribute VB_Name = "netChatterBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gNudgeTimer              As Long     'Seconds.
Dim gNudgeCount              As Long
Dim gSendTimer               As Long
Dim gSendCount               As Long
Dim gSoundTimer              As Long
Dim gSoundCount              As Long
Dim gSoundFile               As String
Dim gNoSquelch              As Boolean

Private Sub chkChime_Click()
    On Error Resume Next
    If txtWrite.Visible Then
        txtWrite.SetFocus
    End If
End Sub

Private Sub chkOptions_Click()
    On Error Resume Next
    Call ShowOptions(chkOptions.Value = vbChecked)
    If txtWrite.Visible Then
        txtWrite.SetFocus
    End If
End Sub

'Show or hide the options in the Chat box.
Private Sub ShowOptions(pShow As Boolean)
    On Error Resume Next
    
    'Show Options.
    If pShow Then
        Me.Height = (frmOptions.Top + frmOptions.Height) - (Me.ScaleHeight - Me.Height)
        chkOptions.Caption = "Less < <"
    
    'Hide Options.
    Else
        Me.Height = (frmOptions.Top) - (Me.ScaleHeight - Me.Height)
        chkOptions.Caption = "More > >"
    End If
End Sub

Private Sub Form_Resize()
    Static sAlreadyInHere As Boolean
    
    On Error Resume Next
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    If sAlreadyInHere Then
        Exit Sub
    End If
    sAlreadyInHere = True
    
    'Width adjustments.
    txtChat.Width = Me.ScaleWidth
    pctSendHide.Width = Me.ScaleWidth
    txtWrite.Width = Me.ScaleWidth
    frmOptions.Width = Me.ScaleWidth
    pctWhois.Width = Me.ScaleWidth - chkRecipient(0).Width
    
    
    'Height adjustments.
    'Show options.
    If chkOptions = vbChecked Then
        txtChat.Height = Me.ScaleHeight - pctSendHide.ScaleHeight - frmOptions.Height
    
    'Hide options.
    Else
        txtChat.Height = Me.ScaleHeight - pctSendHide.ScaleHeight
    End If
    pctSendHide.Top = txtChat.Height
    frmOptions.Top = pctSendHide.Top + pctSendHide.ScaleHeight
    
    txtChat.SelStart = Len(txtChat.Text)
    sAlreadyInHere = False
    Exit Sub
ErrHand:
    sAlreadyInHere = False
    Exit Sub
End Sub

Private Sub cmdNudge_Click()
    Dim vNudgText As String
    Dim vIndex As Long
    
    On Error Resume Next
    
    If txtWrite.Visible Then
        txtWrite.SetFocus
    End If
    
    gNudgeCount = gNudgeCount + 1
    If gNudgeCount > 2 Then
        If CLng(Time * 100000) > gNudgeTimer + 5 _
        Or CLng(Time * 100000) < gNudgeTimer Then
            gNudgeTimer = CLng(Time * 100000)
            gNudgeCount = 0
        Else
            Exit Sub
        End If
    End If
    
    For vIndex = 0 To 2
        vNudgText = vNudgText & "\clr=" & CLng(GenRandom4 * &HFFFFFE) & " " & "*"
    Next
    
    XmitMessage vNudgText & "\beep=1"
End Sub

Private Sub cmdSend_Click()
    On Error Resume Next
    If txtWrite.Visible Then
        txtWrite.SetFocus
        SendKeys vbCr
    End If
End Sub

Private Sub cmdVote_Click()
    Dim vPlayerIndex As Long
    Dim vTerminalIndex As Long
    Dim vParts() As String
    
    On Error Resume Next
    
    If txtWrite.Visible Then
        txtWrite.SetFocus
    End If
    
    'Show the Network Admin Panel with the Connections tab showing.
    netMain.Show , TheMainForm
    netMain.tabInfo.Tabs.Item(3).Selected = True
    
    'If a player is selected, try to find their controlling terminal
    'and highlight it in the connections listbox.
    For vPlayerIndex = 0 To 5
        If chkRecipient(vPlayerIndex).Value = vbChecked Then
            For vTerminalIndex = 1 To netMain.lsvConnections.ListItems.Count
                vParts = Split(netMain.lsvConnections.ListItems(vTerminalIndex).Tag, ",")
                If CInt(vParts(2)) = net.playerOwner(vPlayerIndex) Then
                    netMain.lsvConnections.ListItems(vTerminalIndex).Selected = True
                    Call netMain.lsvConnectionsItemSelected(vTerminalIndex)
                    netMain.lsvConnections.SetFocus
                    Exit Sub
                End If
            Next
        End If
    Next
    
    netMain.lsvConnections.ListItems(1).Selected = True
    Call netMain.lsvConnectionsItemSelected(1)
    netMain.lsvConnections.SetFocus
End Sub

Private Sub Form_Activate()
    Static sDoNotMove As Boolean
    
    On Error Resume Next
    txtChat.SelStart = Len(txtChat.Text)
    frmOptions.ToolTipText = Phrase(329)        'Select recipient(s) of private messages
    cmdHide.Caption = Phrase(334)             '&Hide
    txtChat.ToolTipText = ""
    txtWrite.ToolTipText = Phrase(335)       'Type your message and press enter to send
    cmdSend.Enabled = False
    
    'Move the chatter box only if it has not already been seen.
    If Not sDoNotMove Then
        sDoNotMove = True
        Me.Top = TheMainForm.Top + 435
        Me.Left = TheMainForm.Left + TheMainForm.Width - Me.Width
    End If
End Sub

Private Sub Form_Load()
    Me.Top = TheMainForm.Top + 435
    Me.Left = TheMainForm.Left + TheMainForm.Width - Me.Width
    With txtChat
    .Top = 0
    .Left = 0
    .Text = ""
    
    End With
    chkOptions.Left = 0
    txtWrite.Left = 0
    txtWrite.Text = ""
    txtChat.SelStart = Len(txtChat.Text)
    cmdSend.Enabled = False
    chkChime.Value = GetSetting(gcApplicationName, "settings", "MsgChimeActive", vbUnchecked)
    gSoundFile = GetSetting(gcApplicationName, "settings", "MsgChimeActiveFile", App.Path & "\Ding.wav")
    
    gNudgeTimer = CLng(Time * 100000)
    gNudgeCount = 0
    gSendTimer = CLng(Time * 100000)
    gSendCount = 0
    
    Call ShowOptions(chkOptions.Value = vbChecked)
    Call Whois
    Call ChooseChatBoxFont
End Sub

'Hide if cloded by the control menu otherwise save and unload on app exit.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If UnloadMode = vbFormControlMenu Then
        Cancel = -1
        txtWrite.Text = ""
        cmdSend.Enabled = False
        Me.Hide
        If TheMainForm.Visible Then
            TheMainForm.SetFocus
        End If
    Else
        SaveSetting gcApplicationName, "settings", "MsgChimeActiveFile", gSoundFile
        SaveSetting gcApplicationName, "settings", "MsgChimeActive", chkChime.Value
    End If
End Sub

Private Sub cmdHide_Click()
    On Error Resume Next
    txtWrite.Text = ""
    cmdSend.Enabled = False
    Me.Hide
    If TheMainForm.Visible Then
        TheMainForm.SetFocus
    End If
End Sub

'Choose a font size for the chat box by scaling up until it fits the smallest box just nicely.
Public Sub ChooseChatBoxFont()
    Dim vCntr As Long
    Dim vFontPoints As Long
    Dim vTestStrY As String
    
    netChatterBox.txtChat.SelStart = Len(netChatterBox.txtChat.Text)
    netChatterBox.txtWrite.Font.name = TheMainForm.Picture1.FontName
    netChatterBox.txtChat.SelFontName = TheMainForm.Picture1.FontName
    netChatterBox.txtWrite.Font.Italic = TheMainForm.Picture1.FontItalic
    netChatterBox.txtChat.SelItalic = TheMainForm.Picture1.FontItalic
    netChatterBox.txtWrite.Font.Bold = TheMainForm.Picture1.FontBold
    netChatterBox.txtChat.SelBold = TheMainForm.Picture1.FontBold
    
    'Load test strings with maximum expected string size.
    vTestStrY = "XWQ@jI(|{["
    
    'Scale up the font until it no longer fits the info box either height or width.
    For vFontPoints = 6 To 400
        netChatterBox.Font.Size = vFontPoints / 4
        
        If netChatterBox.TextHeight(vTestStrY) > 230 Then
            txtWrite.Font.Size = (vFontPoints - 1) / 4
            txtChat.SelFontSize = (vFontPoints - 1) / 4
            Exit For
        End If
    Next
End Sub

'From remote terminals only.
Public Sub printMessageByte(BytBuf() As Byte, Optional PlayerByte As Byte = 255)
    Dim strText As String
    strText = StrConv(BytBuf(), vbUnicode)
    strText = Mid(strText, 3)               'Cut header
    Call WriteMessage(strText, PlayerByte)
    
    'Squelch only if switched on by the user and if this is
    'destined for the chat box. This stops the squelch sound
    'every time the counter is updated.
    If chkChime.Value = vbChecked _
    And gNoSquelch = False _
    And InStr(1, strText, "\DestWindow=Counter") <= 0 Then
        PlaySoundFromFile RandomSquelchSound
    End If
End Sub

'Pick random squelch sound file (squelch0.wav - squelch3.wav)
Private Function RandomSquelchSound() As String
    Dim i As Long
    
    i = CLng(GenRandom4 * 3)
    RandomSquelchSound = App.Path & "\squelch" & CStr(i) & ".wav"
End Function

'Print text to the chatter box. Add carriage return only if needed.
Public Sub printMessageString(pText As String, Optional PlayerByte As Byte = 255)
    
    'Quick and dirty way to garantee a vbcrlf at the end of the text.
    pText = Replace(pText & vbCrLf, vbCrLf & vbCrLf, vbCrLf)
    Call WriteMessage(pText, PlayerByte)
End Sub

Private Sub WriteMessage(strText As String, Optional PlayerByte As Byte = 255)
    Dim vDestinationBox As RichTextBox
    
    On Error GoTo ErrHand
    
    'Select which rich text box this message is destined for.
    If InStr(1, LCase(strText), "\destwindow=counter ") Then
        
        'To the audit box on TheMainForm.
        'Set vDestinationBox = TheMainForm.txtAudit
        Set vDestinationBox = frmCounter.txtAudit
        'If Not vDestinationBox.Visible Then
        '    Exit Sub
        'End If
    
    'Future expansion to allow the host to send messages to a
    'rich text message box that is yet to be created.
    'ElseIf InStr(1, LCase(strText), "\DestWindow=MessageBox ") Then
    '
    '    'Hide the audit box.
    '    Set vDestinationBox = fmrMessageBox
    Else
        
        'To the normal chat window.
        Set vDestinationBox = txtChat
        If Not Me.Visible _
        And Not gHeadlessMode Then
            Me.Show , TheMainForm
        End If
    End If
    
    With vDestinationBox
    .SelStart = Len(.Text)
    .SelColor = vbWhite
        
    'Knock off CRLF and place at front if not the first line in box.
    If Len(.Text) = 0 Then
        Call WriteText(Mid(strText, 1, Len(strText) - 2), vDestinationBox)
    Else
        Call WriteText(vbCrLf & Mid(strText, 1, Len(strText) - 2), vDestinationBox)
    End If
    
    .SelStart = Len(.Text)
    End With
    
    Set vDestinationBox = Nothing
    Exit Sub
ErrHand:
    Set vDestinationBox = Nothing
    Exit Sub
End Sub

'Decode directives and write message.
'Directive syntax: "\command=value"
Private Sub WriteText(strText As String, pDetinationBox As RichTextBox)
    Dim vSelText As String
    Dim vDirective As String
    Dim vPos As Long
    Dim vStart As Long
    
    Dim vSelColor As Long
    Dim vSelFontName As String
    Dim vSelFontSize As Long
    Dim vSelCharOffset As Long
    Dim vSelItalic As Long
    Dim vSelUnderline As Long
    Dim vSelStrikeThru As Long
    Dim vSelBold As Long

    On Error GoTo ErrHand
    
    'Save text font properties.
    With pDetinationBox
    vSelBold = .SelBold
    vSelColor = .SelColor
    vSelFontName = .SelFontName
    vSelFontSize = .SelFontSize
    vSelCharOffset = .SelCharOffset
    vSelItalic = .SelItalic
    vSelUnderline = .SelUnderline
    vSelStrikeThru = .SelStrikeThru
    
    
    
    vStart = 1
    vPos = InStr(vStart, strText, "\")
    gNoSquelch = False
    Do
        If vPos = 0 Then
            vSelText = Mid(strText, vStart)
            .SelText = vSelText
            Exit Do
        Else
            vSelText = Mid(strText, vStart, vPos - vStart)
            .SelText = vSelText
            vStart = vPos + DecodeDirective(Mid(strText, vPos), pDetinationBox)
            vPos = InStr(vStart, strText, "\")
        End If
    Loop
    
    'Restore properties.
    .SelBold = vSelBold
    .SelColor = vSelColor
    .SelFontName = vSelFontName
    .SelFontSize = vSelFontSize
    .SelCharOffset = vSelCharOffset
    .SelItalic = vSelItalic
    .SelUnderline = vSelUnderline
    .SelStrikeThru = vSelStrikeThru
    
    End With
    
    Exit Sub
ErrHand:
    Exit Sub
End Sub

'Decode and implement directive.
Private Function DecodeDirective(strText As String, pDetinationBox As RichTextBox) As Long
    Dim vDirLen As Long
    Dim vEqlPos As Long
    Dim vCntr As Long
    Dim i As Long
    Dim vDirective As String
    Dim vValue As String
    
    On Error Resume Next
    
    With pDetinationBox
    
    vDirLen = InStr(1, strText, " ")
    If vDirLen = 0 Then
        vDirLen = Len(strText)
    End If
    DecodeDirective = vDirLen
    
    'Decode directive and value.
    vEqlPos = InStr(1, Mid(strText, 1, vDirLen), "=")
    If vEqlPos > 0 Then
        vDirective = Mid(strText, 2, vEqlPos - 2)
        vValue = Mid(strText, vEqlPos + 1, vDirLen - vEqlPos)
    Else
        vDirective = Trim(Mid(strText, 2, vDirLen - 1))
        vValue = ""
    End If
    
    'Colour. "\clr=nnnnn "
    If vDirective = "clr" Then
        If IsNumeric(vValue) Then
            .SelStart = Len(.Text)
            .SelColor = CLng(vValue)
        Else
            .SelStart = Len(.Text)
            .SelColor = CLng(vbWhite)
        End If
    
    'Font. "\fname=Font_Name_Text "
    'Underscore replaces while space.
    ElseIf vDirective = "fname" Then
        .SelStart = Len(.Text)
        .SelFontName = Trim(Replace(vValue, "_", " "))
        'Default
        '.SelStart = Len(.Text)
        '.SelFontName = "Comic Sans MS"
    
    'Font size. "\fsize=[nn] "
    ElseIf vDirective = "fsize" Then
        If IsNumeric(vValue) Then
            .SelStart = Len(.Text)
            .SelFontSize = Abs(CLng(vValue))
        End If
    
    'Font char offset. "\fofst=[nn] " where positive integer is superscript,
    'negative integer is subscript, 0 is normal.
    ElseIf vDirective = "fofst" Then
        If IsNumeric(vValue) Then
            .SelStart = Len(.Text)
            .SelCharOffset = CLng(vValue)
        End If
    
    'Font bold. "\fbold=[0|1] "
    ElseIf vDirective = "fbold" Then
        If IsNumeric(vValue) Then
            .SelStart = Len(.Text)
            .SelBold = CBool(vValue)
        End If
    
    'Font italic. "\fital=[0|1] "
    ElseIf vDirective = "fital" Then
        If IsNumeric(vValue) Then
            .SelStart = Len(.Text)
            .SelItalic = CBool(vValue)
        End If
    
    'Font underline. "\fulin=[0|1] "
    ElseIf vDirective = "fulin" Then
        If IsNumeric(vValue) Then
            .SelStart = Len(.Text)
            .SelUnderline = CBool(vValue)
        End If
    
    'Font strike-through. "\fstru=[0|1] "
    ElseIf vDirective = "fstru" Then
        If IsNumeric(vValue) Then
            .SelStart = Len(.Text)
            .SelStrikeThru = CBool(vValue)
        End If
    
    
    'Play sound. "\beep "
    ElseIf vDirective = "beep" Then
        gNoSquelch = True
        If IsNumeric(vValue) Then
            vCntr = CLng(vValue)
            If vCntr > 10 Then
                vCntr = 10
            End If
        Else
            vCntr = 1
        End If
        
        'Stop those annoying idiots.
        gSoundCount = gSoundCount + 1
        If gSoundCount > 1 Then
            If CLng(Time * 100000) > gSoundTimer + 10 _
            Or CLng(Time * 100000) < gSoundTimer Then
                gSoundTimer = CLng(Time * 100000)
                gSoundCount = 0
            Else
                Exit Function
            End If
        End If
        
        'Play sound.
        For i = 1 To vCntr
            PlaySoundFromFile gSoundFile
            'pause 200, True
            Call Sleep(200)
        Next
    
    'Squelch noise (hiss). "\squelch "
    ElseIf vDirective = "squelch" Then
        gNoSquelch = True
        If IsNumeric(vValue) Then
            vCntr = CLng(vValue)
            If vCntr > 10 Then
                vCntr = 10
            End If
        Else
            vCntr = 1
        End If
        
        'Stop those annoying idiots.
        gSoundCount = gSoundCount + 1
        If gSoundCount > 1 Then
            If CLng(Time * 100000) > gSoundTimer + 10 _
            Or CLng(Time * 100000) < gSoundTimer Then
                gSoundTimer = CLng(Time * 100000)
                gSoundCount = 0
            Else
                Exit Function
            End If
        End If
        
        'Play squelch sound.
        For i = 1 To vCntr
            PlaySoundFromFile RandomSquelchSound
            'pause 200, True
            Call Sleep(200)
        Next
    
    'Escape the escape character is "\".
    ElseIf Mid(vDirective, 1, 1) = "\" Then
        .SelText = "\"
        DecodeDirective = 2
    
    'Don't uderstand directive. Skip it.
    'Or to print it, uncomment the next three lines.
    'Else
        '.SelText = "\"
        'DecodeDirective = 1
    End If
    End With
End Function

'Return text color for this player.
Private Function GetPlayerColor(PlayerByte As Byte, strText As String) As Long
    Dim ArmyNo As Long
    
    GetPlayerColor = vbWhite    'Default
    
    If netWorkSituation = cNetNone Then
        If PlayerByte < 6 Then
            For ArmyNo = 0 To 5
                If (TheMainForm.playerSelect_getIndex(CInt(ArmyNo)) = 0) _
                And ((Mid(strText, 1, InStr(1, strText, ">>") + 1) = netMain.txtTerminalName & ">>") _
                    Or (Mid(strText, 1, InStr(1, strText, ">>") + 1) = netMain.txtTerminalName & "(prv)>>")) _
                Then
                    GetPlayerColor = gPlayerID(ArmyNo + 1).lngColor
                    Exit For
                End If
            Next
        End If
    Else
        For ArmyNo = 0 To 5
            If net.playerOwner(ArmyNo) = PlayerByte Then
                If (net.playerOwner(ArmyNo) <> myTerminalNumber _
                And ((Mid(strText, 1, InStr(1, strText, ">>") + 1)) = TheMainForm.PlayerSelect(ArmyNo).Text & ">>") _
                    Or (Mid(strText, 1, InStr(1, strText, ">>") + 1) = TheMainForm.PlayerSelect(ArmyNo).Text & "(prv)>>")) _
                Or (net.playerOwner(ArmyNo) = myTerminalNumber _
                And TheMainForm.playerSelect_getIndex(CInt(ArmyNo)) <> remoteIndex) _
                Then
                    GetPlayerColor = gPlayerID(ArmyNo + 1).lngColor
                    Exit For
                End If
            End If
        Next
    End If
End Function

    'Private message. Check if I can read it then print
Public Sub privateMessage(BytBuf() As Byte, PlayerByte As Byte)
    Dim strText As String
    
    On Error Resume Next
    strText = StrConv(BytBuf(), vbUnicode)
    
    If Not controlArmy(Mid(strText, 3, 2)) Then
        Exit Sub
    End If
    strText = Mid(strText, 5)       'Cut header and destination
    WriteMessage gGsLeUtils.LE1d(strText), PlayerByte 'Host
    If chkChime.Value = vbChecked And gNoSquelch = False Then
        PlaySoundFromFile RandomSquelchSound
    End If
End Sub

    'Return true if this terminal controls this army (bit coded decimal string)
Private Function controlArmy(strNmbrCode As String) As Boolean
    Dim nbrCode As Long
    Dim termCode As Long
    Dim cntr As Long
    
    nbrCode = CLng(strNmbrCode) - 10
    
    If nbrCode = 63 Then            'Every one can see if all checked
        controlArmy = True
        Exit Function
    End If
    
    For cntr = 0 To 5
        If net.playerOwner(cntr) = myTerminalNumber _
        And (CLng(2 ^ cntr) And nbrCode) <> 0 Then
            controlArmy = True
            Exit Function
        End If
    Next
    controlArmy = False
End Function

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If txtWrite.Visible Then
        txtWrite.SetFocus
        SendKeys Chr(KeyAscii)
        'KeyAscii = 0
    End If
End Sub

'Get color of my terminal. Find first human, else first computer, else white.
'Return command for that color.
Public Function CreateColorCodeForTerminal() As String
    Dim i As Long
    
    'Find first human player from top to bottom.
    For i = 0 To 5
        If (TheMainForm.playerSelect_getIndex(CInt(i)) = 0) Then
            CreateColorCodeForTerminal = "\clr=" & CStr(gPlayerID(i + 1).lngColor) & " "
            Exit Function
        End If
    Next
    
    'Else find first computer player from top to bottom.
    For i = 0 To 5
        If (TheMainForm.playerSelect_getIndex(CInt(i)) <> remoteIndex) Then
            CreateColorCodeForTerminal = "\clr=" & CStr(gPlayerID(i + 1).lngColor) & " "
            Exit Function
        End If
    Next
    
    CreateColorCodeForTerminal = "\clr=" & CStr(vbWhite) & " "
End Function

'Keep track of key presses, if CR then xmit string
Private Sub txtWrite_KeyPress(KeyAscii As Integer)
    Dim x As Long
    Dim strSend As String

    On Error Resume Next
    
    If KeyAscii = Asc(vbCr) Then
        
        'Limit amount of times that text can be sent over time.
        gSendCount = gSendCount + 1
        If gSendCount > 3 Then
            If CLng(Time * 100000) > gSendTimer + 5 _
            Or CLng(Time * 100000) < gSendTimer Then
                gSendTimer = CLng(Time * 100000)
                gSendCount = 0
            Else
                txtWrite.Text = ""
                KeyAscii = 0
                cmdSend.Enabled = False
                Exit Sub
            End If
        End If
        
        If Not CheckForCheatCode Then
            Call XmitMessage(CreateColorCodeForTerminal & _
                Mid(txtWrite.Text, 1, 200) & " \clr=" & CStr(vbWhite) & " ")
        End If
        txtWrite.Text = ""
        KeyAscii = 0
        cmdSend.Enabled = False
    Else
        cmdSend.Enabled = True
    End If
End Sub

'Send message
Private Sub XmitMessage(pMessageText As String)
    If netMain.CountTerminals = 0 Then
        If AnyPlayerSelected Then
            Call netMain.XmitString(0, 17, myTerminalNumber, _
                    getDest & gGsLeUtils.LE1(netMain.txtTerminalName & "(prv)>>" & pMessageText) & vbCrLf)
            Call WriteMessage(netMain.txtTerminalName & "(prv)>>" & pMessageText & vbCrLf, myTerminalNumber)
        Else
            Call netMain.XmitString(0, 1, myTerminalNumber, _
                    netMain.txtTerminalName & ">>" & pMessageText & vbCrLf)
            WriteMessage netMain.txtTerminalName & ">>" & pMessageText & vbCrLf, myTerminalNumber
        End If
    Else
        If AnyPlayerSelected Then
            Call netMain.XmitStringAll(0, 17, myTerminalNumber, _
                    getDest & gGsLeUtils.LE1(netMain.txtTerminalName & "(prv)>>" & pMessageText) & vbCrLf)
            Call WriteMessage(netMain.txtTerminalName & "(prv)>>" & pMessageText & vbCrLf, myTerminalNumber)
        Else
            Call netMain.XmitStringAll(0, 1, myTerminalNumber, _
                    netMain.txtTerminalName & ">>" & pMessageText & vbCrLf)
            Call WriteMessage(netMain.txtTerminalName & ">>" & pMessageText & vbCrLf, myTerminalNumber)
        End If
    End If
End Sub

    'Return TRUE if cheat code typed
Private Function CheckForCheatCode() As Boolean
    Dim mr As String
    Dim CheatResponse As String
    
    mr = Trim(LCase(txtWrite.Text))
    CheckForCheatCode = False
    
    If Len(mr) < 6 Then
        Exit Function
    ElseIf Left(mr, 3) <> "mr#" Then
        Exit Function
    End If
    
    mr = Trim(Mid(mr, 4))
        
    'Add 50 units
    If mr = gCheatMode.inCodes(0) _
    And gCurrentMode = 2 Then '
        TheMainForm.gPlayerValue = TheMainForm.gPlayerValue + 50
        Call TheMainForm.printPlaceUnits
        CheckForCheatCode = True
        'txtChat = txtChat & vbCrLf & gCheatMode.responses(0)
        CheatResponse = gCheatMode.responses(0)
    
    'See other cards
    ElseIf mr = gCheatMode.inCodes(1) _
    And Not gCheatMode.seeCards Then
        gCheatMode.seeCards = True
        Call TheMainForm.refreshMap
        CheckForCheatCode = True
        CheatResponse = gCheatMode.responses(1)
    
    'Change map
    ElseIf mr = gCheatMode.inCodes(2) Then
        gCheatMode.createMap = True
        gCheatMode.undoEnabled = True
        CheckForCheatCode = True
        CheatResponse = gCheatMode.responses(2)
    
    'See missions
    ElseIf mr = gCheatMode.inCodes(3) _
    And Not gCheatMode.seeMissions Then
        gCheatMode.seeMissions = True
        CheckForCheatCode = True
        CheatResponse = gCheatMode.responses(3)
        
    'Auto restart
    ElseIf mr = gCheatMode.inCodes(4) Then
        If Not TheMainForm.mnuAutoRestart.Checked Then
            'gCheatMode.autoRestart = True
            TheMainForm.mnuAutoRestart.Checked = True
            CheckForCheatCode = False
            CheatResponse = gCheatMode.responses(4)
        Else
            'gCheatMode.autoRestart = False
            TheMainForm.mnuAutoRestart.Checked = False
            CheckForCheatCode = False
            CheatResponse = "Auto restart deactivated"
        End If
        Call WriteMessage(CheatResponse & vbCrLf, 0)
            
    'Show gCurrentMode and timer status
    ElseIf mr = gCheatMode.inCodes(5) Then
        gCheatMode.testing = Not gCheatMode.testing
        CheckForCheatCode = False
        Call WriteMessage(gCheatMode.responses(5) & vbCrLf, 0)
        Call TheMainForm.updateTestViewer(" ")
    
    'Modeling mode. Change form to a standard size for screenshots.
    ElseIf mr = gCheatMode.inCodes(6) Then
        CheatResponse = gCheatMode.responses(6) _
                        & vbCrLf & "Original size " & CStr(TheMainForm.Height) _
                        & "x" & CStr(TheMainForm.Width)
        TheMainForm.Height = 7875
        TheMainForm.Width = 9885
        Call WriteMessage(CheatResponse & vbCrLf, 0)
    
    
    ElseIf mr = gCheatMode.inCodes(7) Then
        gAppLogLevel = 9
        CheatResponse = gCheatMode.responses(7)
        Call WriteMessage(CheatResponse & vbCrLf, 0)
        
    Else
        Exit Function
    End If
    gCheatMode.cheatActive = gCheatMode.cheatActive Or CheckForCheatCode
    
    'Block cheat mode if networked.
    If gCheatMode.cheatActive And netWorkSituation > cNetNone Then
        Call TheMainForm.TurnOffCheatCodes
        CheckForCheatCode = False
        Call WriteMessage("No cheats while networked." & vbCrLf, 0)
    End If
    
    If CheckForCheatCode Then               'Dob if affects chances in war.
        If netMain.CountTerminals = 0 Then
            Call netMain.XmitString(0, 1, myTerminalNumber, _
                netMain.txtTerminalName & ">>" & txtWrite.Text & vbCrLf)
        Else
            Call netMain.XmitStringAll(0, 1, myTerminalNumber, _
                netMain.txtTerminalName & ">>" & txtWrite.Text & vbCrLf)
        End If
        Call WriteMessage(netMain.txtTerminalName & ">>" & txtWrite.Text _
                        & vbCrLf, myTerminalNumber)
                        
        CheatResponse = Phrase(260) & vbCrLf & CheatResponse
        If netMain.CountTerminals = 0 Then
            Call netMain.XmitString(0, 1, 0, CheatResponse & vbCrLf)
        Else
            Call netMain.XmitStringAll(0, 1, 0, CheatResponse & vbCrLf)
        End If
        Call WriteMessage(CheatResponse & vbCrLf)
    End If
End Function

    'Get the destination players (bit code) and convert to a decimal string
Private Function getDest() As String
    Dim cntr As Long
    Dim rslt As Long
    
    For cntr = 0 To 5
        If chkRecipient(cntr).Value Then
            rslt = rslt + CLng(2 ^ cntr)
        End If
    Next
    getDest = Trim(CStr(rslt + 10))     ' Add 10 to ensure 2 digits
End Function

Private Sub Whois()
    Dim ArmyNo As Integer
    
    pctWhois.Cls
    For ArmyNo = 1 To 6
        pctWhois.ForeColor = gPlayerID(ArmyNo).bkgndColor
        
        'Mark the current player
        If ArmyNo = gPlayerTurn Then
            pctWhois.Font.Underline = True
        Else
            pctWhois.Font.Underline = False
        End If
        
        pctWhois.Print TheMainForm.GetArmyOrControllerName(CByte(ArmyNo))
    Next
End Sub

'Return TRUE if any player has been selected.
Private Function AnyPlayerSelected() As Boolean
    Dim vPlayer As Long
    
    AnyPlayerSelected = False
    
    For vPlayer = 0 To 5
        If chkRecipient(vPlayer).Value = vbChecked Then
            AnyPlayerSelected = True
            Exit For
        End If
    Next
End Function

'Enable buttons if any player is selected and then call Whois.
Private Sub tmrWhois_Timer()
    Call Whois
End Sub

