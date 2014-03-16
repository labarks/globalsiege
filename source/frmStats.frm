VERSION 5.00
Begin VB.Form frmStats 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "War Statistics"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7725
   ControlBox      =   0   'False
   Icon            =   "frmStats.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmStats.frx":030A
   PaletteMode     =   2  'Custom
   ScaleHeight     =   5850
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkShowAgain 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   8
      Top             =   5400
      Width           =   195
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00000000&
      Cancel          =   -1  'True
      Caption         =   "OK"
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
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox pctStats 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   5
      Left            =   5160
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.PictureBox pctStats 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   4
      Left            =   2640
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   5
      Top             =   2760
      Width           =   2415
   End
   Begin VB.PictureBox pctStats 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   3
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.PictureBox pctStats 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   2
      Left            =   5160
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox pctStats 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   1
      Left            =   2640
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox pctStats 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   0
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblShowAgain 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Don't show anymore"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   210
      Left            =   360
      TabIndex        =   7
      Top             =   5445
      Width           =   1440
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkShowAgain_Click()
    TheMainForm.mnuOptStats.Checked = Not CBool(Abs(chkShowAgain.Value))
End Sub

Private Sub cmdOK_Click()
    Me.Hide
    TheMainForm.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 18 Then
        Call TheMainForm.mnuFileReset_Click
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Pause button.
    If KeyCode = 19 Then
        Call TheMainForm.ActivatePauseMode
    End If
End Sub

Private Sub Form_Load()
    Me.Top = TheMainForm.Top + (TheMainForm.Height - Me.Height) / 2
    Me.Left = TheMainForm.Left + (TheMainForm.Width - Me.Width) / 2
End Sub

'Modify the score depending on number of units involved in the war.
Public Function ScoreModifier(pUnitsInvloved As Long) As Single
    
    If pUnitsInvloved < 1 Then
        ScoreModifier = 0
    ElseIf pUnitsInvloved < 30 Then
        ScoreModifier = (pUnitsInvloved / 30) ^ 2
    ElseIf pUnitsInvloved <= 100 Then
        ScoreModifier = 1
    ElseIf pUnitsInvloved <= 900 Then
        ScoreModifier = 1 - ((pUnitsInvloved - 100) / 900)
    Else
        ScoreModifier = 0.2
    End If
End Function

'Calculate and return the score from the passed stats.
Private Function CalculateScore( _
pUnitsIBeat As Long, _
pUnitsILost As Long, _
pCountriesIBeat As Long, _
pCountriesILost As Long, _
pHumanStartedCount As Long, _
pMyStartingUnits As Long, _
pSmartComputerCount As Long, _
pTotalComputerCount As Long) As Single
    
    Dim vScore As Single
    
    On Error Resume Next
    
    'If less than 10 units were beaten or lost in the war.
    If (pUnitsIBeat + pUnitsILost) <= 10 Then
        
        'Too few battles to qualify for a score.
        vScore = 0
    Else
        
        'Work out the score by proportion of beaten/lost.
        vScore = 20 * ((pCountriesIBeat / pCountriesILost) * (pUnitsIBeat / pUnitsILost))
        
        'Modify the score by rewarding quick and efficient battles. The fewer
        'units killed on either side the better.
        vScore = vScore * ScoreModifier((pUnitsILost + pUnitsIBeat - 2) / 2)
        
        'Consider starting countries without extra units distributed.
        If TheMainForm.chkExtraStartingUnits.Value = vbUnchecked Then
            
            vScore = vScore * ((63 / pHumanStartedCount - pMyStartingUnits) / 4.5)
        End If
        
        'Consider the proportion of smart and dumb computer players.
        vScore = vScore * ((pSmartComputerCount / 2 + 2) / pTotalComputerCount)
        
        'Don't let the score be negative.
        If vScore < 0 Then
            vScore = 0
        End If
        
    End If
    
    'Return the score.
    CalculateScore = vScore
End Function

'Display available stats for last war
Public Sub ShowStats()
    On Error Resume Next
    
    If TheMainForm.mnuOptStats.Checked _
    And Not gHeadlessMode Then
        Me.Show , TheMainForm
    End If
    
    Exit Sub
errorHand:
    'MsgBox "error"
    Resume Next
End Sub

Public Function CalculateStats()
    Dim vIndex As Long
    Dim vScore As Single
    Dim vCountriesIBeat As Long
    Dim vCountriesILost As Long
    Dim vUnitsIBeat As Long
    Dim vUnitsILost As Long
    Dim vHumanStartedCount As Long
    Dim vMyStartingUnits As Long
    Dim vSmartComputerCount As Long
    Dim vTotalComputerCount As Long
    Dim vPrintXPosStart As Long
    Dim vIxStatsReport As String
    
    On Error Resume Next
    
    cmdOK.Caption = Phrase(340)
    frmStats.Caption = Phrase(345)
    lblShowAgain.Caption = Phrase(136)
    
    vSmartComputerCount = 1
    vTotalComputerCount = 1
    
    'Make sure the stats global has the correct terminals associated with the players.
    Call TheMainForm.ListPlayerControllers
    
    'Count the computer players and human players. gPlayerStats().PlrController
    'knows if the remote player is controlled by a human or computer player at
    'the other end.
    For vIndex = 1 To 6
        With gPlayerStats(vIndex)
        
        If .PlrController = 1 _
        Or .PlrController = 2 _
        Or .PlrController = 3 Then
            vTotalComputerCount = vTotalComputerCount + 1
        End If
        
        If .PlrController = 2 Then
            vSmartComputerCount = vSmartComputerCount + 1
        End If
        
        If gPlayerID(vIndex).startWith > 0 Then
            vHumanStartedCount = vHumanStartedCount + 1
        End If
        End With
        
    Next
    On Error Resume Next
    
    'For each player.
    For vIndex = 0 To 5
        
        'Set the stats window to the player's background colour and clear.
        pctStats(vIndex).BackColor = gPlayerID(vIndex + 1).bkgndColor
        pctStats(vIndex).Cls
        
        'Print the player's terminal name in bold.
        pctStats(vIndex).Font.Bold = True
        pctStats(vIndex).Print TheMainForm.GetArmyOrControllerName(CByte(vIndex) + 1)
        pctStats(vIndex).Font.Bold = False
        
        'Check if the player was involved in the war.
        If gPlayerID(vIndex + 1).startWith = 0 Then
            
            'Not involved so print "Not involved".
            pctStats(vIndex).Print LimitWidth(Phrase(346))
        Else
            
            'Player was involved. Work out and print stats.
            vCountriesIBeat = gPlayerStats(vIndex + 1).CountriesDefeated
            vCountriesILost = gPlayerStats(vIndex + 1).CountriesLost
            vUnitsIBeat = gPlayerStats(vIndex + 1).UnitsBeaten
            vUnitsILost = gPlayerStats(vIndex + 1).UnitsLost
            vMyStartingUnits = CLng(gPlayerID(vIndex + 1).startWith)
            
            vScore = CalculateScore(vUnitsIBeat + 1, _
                                    vUnitsILost + 1, _
                                    vCountriesIBeat + 1, _
                                    vCountriesILost + 1, _
                                    vHumanStartedCount, _
                                    vMyStartingUnits, _
                                    vSmartComputerCount, _
                                    vTotalComputerCount)
            
            'Print the stats and score to the player's stats window.
            
            'Print the mission.
            pctStats(vIndex).Print LimitWidth(gPlayerStats(vIndex + 1).StartingMission) & vbCrLf
            pctStats(vIndex).CurrentY = 1170
            vPrintXPosStart = 1800
            
            'Countries conquered.
            pctStats(vIndex).Print Phrase(348);         'Countries conquered:
            pctStats(vIndex).CurrentX = vPrintXPosStart
            pctStats(vIndex).Print CStr(vCountriesIBeat)
            
            'Countries lost.
            pctStats(vIndex).Print Phrase(349);         'Countries surrendered:
            pctStats(vIndex).CurrentX = vPrintXPosStart
            pctStats(vIndex).Print CStr(vCountriesILost)
            
            'Units beaten.
            pctStats(vIndex).Print Phrase(350);         'Enemy casualties:
            pctStats(vIndex).CurrentX = vPrintXPosStart
            pctStats(vIndex).Print CStr(vUnitsIBeat)
            
            'Units defeated.
            pctStats(vIndex).Print Phrase(351);         'Your casualties:
            pctStats(vIndex).CurrentX = vPrintXPosStart
            pctStats(vIndex).Print CStr(vUnitsILost)
            
            pctStats(vIndex).Font.Bold = True
            
            'If cheat mode is actice.
            If gCheatMode.cheatActive And gPlayerID(vIndex + 1).playerWho = 0 Then
                
                'Print "CHEAT!" and set the score to 0.
                pctStats(vIndex).Print
                pctStats(vIndex).Print Phrase(353);
                vScore = 0
                
            'If stats were invalidated.
            ElseIf Not gPlayerStats(vIndex + 1).IsValid Then
                
                'Print the reason for being invalidated and set the score to 0.
                pctStats(vIndex).Print gPlayerStats(vIndex + 1).InvalidatedReason
                vScore = 0
              
            'If stats are valid.
            Else
                
                'Print the score.
                pctStats(vIndex).Print
                pctStats(vIndex).Print Phrase(354);         'Score:
                pctStats(vIndex).CurrentX = vPrintXPosStart
                pctStats(vIndex).Print CStr(CLng(vScore))
                
                'Must be a human player on this terminal or a human player
                'on the remote terminal. No stats are submitted for computer
                'controlled players.
                If netWorkSituation = cNetHost _
                And Not TheMainForm.IsComputerPlayer(vIndex + 1) _
                And ((net.playerOwner(vIndex) <> myTerminalNumber) _
                And (net.Controller(vIndex) = 0) _
                Or net.playerOwner(vIndex) = myTerminalNumber) Then
                    
                    'Format the report to send to the Indexing Server.
                    'gGsLeUtils.LE6(vControllingTerminal),vCountriesIBeat,
                    'vCountriesILost,vUnitsIBeat,vUnitsILost,vScore.
                    vIxStatsReport = vIxStatsReport _
                                    & gGsLeUtils.LE6(net.ClientName(net.playerOwner(vIndex))) & "," _
                                    & CStr(vCountriesIBeat) & "," _
                                    & CStr(vCountriesILost) & "," _
                                    & CStr(vUnitsIBeat) & "," _
                                    & CStr(vUnitsILost) & "," _
                                    & CStr(CLng(vScore)) & vbCrLf
                End If
            End If
            pctStats(vIndex).Font.Bold = False

        End If
        pctStats(vIndex).Refresh
    Next
    
    'Lop off the last CRLF and encrypt.
    vIxStatsReport = gGsLeUtils.LE6(CleanList(vIxStatsReport, vbCrLf))
    
    'Send stats to the Indexing Server.
    Call IxServerUploadStats(vIxStatsReport)
    
End Function

    'Wrap text as required to fit in picture box
Private Function LimitWidth(str As String) As String
    Dim cntr As Long
    Dim MaxWidth As Long
    Dim StartLine As Long
    Dim LastSpace As Long
    Dim FormatString As String
    
    On Error Resume Next
    
    MaxWidth = pctStats(0).Width - 200
    StartLine = 1
    LastSpace = 1
    FormatString = ""
    
    For cntr = 1 To Len(str)
        If pctStats(0).TextWidth(Mid(str, StartLine, cntr - StartLine)) >= MaxWidth Then
            If LastSpace <= StartLine Then
                FormatString = FormatString & Mid(str, StartLine, cntr - StartLine) & vbCrLf
                LastSpace = cntr
                StartLine = cntr
            Else
                FormatString = FormatString & Mid(str, StartLine, LastSpace - StartLine) & vbCrLf
                StartLine = LastSpace + 1
                LastSpace = LastSpace + 1
            End If
            
        ElseIf Mid(str, cntr, 1) = " " Then
            LastSpace = cntr
        End If
    Next
    LimitWidth = FormatString & Mid(str, StartLine, Len(str))
End Function

'Don't actually close the form if it was closed by the user. Hide it instead
'so that they can reopen it at any time.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Me.Hide
        Cancel = -1
        Exit Sub
    End If
End Sub
