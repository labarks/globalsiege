Attribute VB_Name = "modCheckWin"
Option Explicit

'Remeber what player is to win at the start of their next turn.
Public gWinMemoryBits As Byte

'Check for world domination and player wipeout during the player's turn.
Public Sub CheckWinDuringTurn(pPlayer As Integer, Optional pRemoteTerminal As Boolean = False)
    If gCurrentMode = 13 Or gCurrentMode = 18 Then
        Exit Sub
    End If
    Call ChkWinDominate(pPlayer)
    Call ChkWinWipeout(pPlayer)
    If gCurrentMode = 13 Or gCurrentMode = 18 Then
        If netWorkSituation <> cNetNone And Not pRemoteTerminal Then
            Call TheMainForm.AuditShadowAppend
            Call TheMainForm.AuditPlayerCompare
            Call netMain.IHaveWon(1)
            net.madeUpdate = True
        End If
        Call TheMainForm.DrawWin
    End If
End Sub

'Check for win by world domination.
Private Sub ChkWinDominate(pPlayer As Integer)
    Dim vCountry As Integer
    
    If gCurrentMode = 13 Or gCurrentMode = 18 Then
        TheMainForm.Timer2.Enabled = False
        Exit Sub
    End If
    
    'Check if passed player owns all countries.
    For vCountry = 1 To 42
        If gCountryOwner(vCountry) <> pPlayer Then
            Exit Sub
        End If
    Next vCountry
    
    'I win.
    Call WeHaveAWinner(pPlayer, gPlayerID(pPlayer).strColor & vbCrLf & vbCrLf _
                    & Phrase(66) & vbCrLf & vbCrLf & Phrase(100))
End Sub

'Signify that we have a winner and put win message in a global variable
'for printing after a certain time.
Private Sub WeHaveAWinner(pWinningPlayer As Integer, pWinInfoText As String)
    TheMainForm.Timer2.Enabled = False
    Call TheMainForm.ToggleKeys(False)
    Call TheMainForm.refreshMap
    TheMainForm.pctInfoBox.BackColor = gPlayerID(pWinningPlayer).bkgndColor
    If Not TheMainForm.SetupScreen.Visible Then
        TheMainForm.BackColor = gPlayerID(pWinningPlayer).bkgndColor
    End If
    gCurrentMode = 13
    Call DrawLittleCards
    
    'Clear info box and set the win message.
    TheMainForm.InfoBoxPrint 0
    TheMainForm.sWinMessage = pWinInfoText
End Sub

'Check for win by player wipeout.
Private Sub ChkWinWipeout(pPlayer As Integer)
    Dim vPlayer As Integer
    Dim vCountry As Integer
    
    'Already have a winner.
    If gCurrentMode = 13 Or gCurrentMode = 18 Then
        Exit Sub
    End If
    
    'Check all players for win by wipeout.
    For vPlayer = 1 To 6
        
        'Is this player's mission to wipe out another player?
        If (gPlayerID(vPlayer).mission < 7) _
        And (gPlayerID(vPlayer).mission <> 0) _
        And (TheMainForm.CountCountriesOwned(pPlayer) > 0) Then
            
            'Check if the player's target army has been exterminated.
            If TheMainForm.CountCountriesOwned(gPlayerID(vPlayer).mission) = 0 Then
                
                'Change to 24 countries if player did not complete own mission.
                If TheMainForm.chkMsnMustComplete.Value = vbChecked _
                And Not GetBit(CLng(vPlayer), gWinMemoryBits) Then
                    If vPlayer <> pPlayer Then
                        gPlayerID(vPlayer).mission = 14
                        Exit Sub
                    End If
                End If
                
                'Player must wait until their next turn to win.
                If TheMainForm.chkMsnWinImmediate.Value = vbUnchecked Then
                    Call SetBit(True, CLng(vPlayer), gWinMemoryBits)
                
                'Player wins immediately.
                Else
                    
                    '"<player> has won the war by wiping out <army>."
                    Call WeHaveAWinner(vPlayer, gPlayerID(vPlayer).strColor & vbCrLf & vbCrLf _
                            & Phrase(66) & vbCrLf & vbCrLf _
                            & Phrase(69) & vbCrLf _
                            & gPlayerID(gPlayerID(vPlayer).mission).strColor & ".")
                    Exit Sub
                End If
            End If
        End If
    Next vPlayer
End Sub

'Check for player wipeout and continent hold victory at the start of the player's turn.
Public Function CheckWinStartOfTurn(pPlayerTurn As Integer, _
Optional pRemoteTerminal As Boolean = False) As Boolean
    
    'Make sure not already in win mode.
    If gCurrentMode = 13 Or gCurrentMode = 18 Then
        TheMainForm.Timer2.Enabled = False
        Exit Function
    End If
    
    'If dominate mission, bail out here because there is
    'no way that a player can win by world domintaion at
    'the start of their turn.
    If gPlayerID(pPlayerTurn).mission = 0 Then
        CheckWinStartOfTurn = False
    
    'Check if player wipeout mission.
    ElseIf gPlayerID(pPlayerTurn).mission < 7 Then
        
        'Has the player been marked as winning at the start of my turn?
        If GetBit(CLng(pPlayerTurn), gWinMemoryBits) _
        And TheMainForm.chkMsnWinImmediate.Value = vbUnchecked Then
            
            '<player> has won by wiping out <target army>.
            Call WeHaveAWinner(pPlayerTurn, gPlayerID(pPlayerTurn).strColor & vbCrLf & vbCrLf _
                        & Phrase(66) & vbCrLf & vbCrLf _
                        & Phrase(69) & vbCrLf _
                        & gPlayerID(gPlayerID(pPlayerTurn).mission).strColor & ".")
            If netWorkSituation <> cNetNone And Not pRemoteTerminal Then
                Call netMain.IHaveWon(2)
                net.madeUpdate = True
            End If
            Call TheMainForm.DrawWin
            CheckWinStartOfTurn = True
        End If
    
    'Check for 24 countriy missions.
    ElseIf gPlayerID(pPlayerTurn).mission = 14 Then
        If TheMainForm.CountCountriesOwned(pPlayerTurn) >= gcMission14 Then
            
            '<player> has won the war By occupying And holding 24 countries.
            Call WeHaveAWinner(pPlayerTurn, gPlayerID(pPlayerTurn).strColor & vbCrLf _
                    & Phrase(66) & vbCrLf & vbCrLf _
                    & Phrase(67) & vbCrLf _
                    & gMissions(gPlayerID(pPlayerTurn).mission).WinMessageText)
            If netWorkSituation <> cNetNone And Not pRemoteTerminal Then
                Call netMain.IHaveWon(2)
                net.madeUpdate = True
            End If
            Call TheMainForm.DrawWin
            CheckWinStartOfTurn = True
        End If
    
    'Check for continent hold missions.
    ElseIf HoldAllConts(pPlayerTurn, gPlayerID(pPlayerTurn).mission) Then
        
        '<player> has won by occupying and holding <continents>.
        Call WeHaveAWinner(pPlayerTurn, gPlayerID(pPlayerTurn).strColor & vbCrLf & vbCrLf _
                    & Phrase(66) & vbCrLf & vbCrLf _
                    & Phrase(67) & vbCrLf _
                    & gMissions(gPlayerID(pPlayerTurn).mission).WinMessageText)
        If netWorkSituation <> cNetNone And Not pRemoteTerminal Then
            Call netMain.IHaveWon(2)
            net.madeUpdate = True
        End If
        Call TheMainForm.DrawWin
        CheckWinStartOfTurn = True
    
    Else
        CheckWinStartOfTurn = False
    End If
End Function

'Return true the passed player holds all the continents of the passed mission.
Private Function HoldAllConts(pPlayer As Integer, pMissionIndex As Integer) As Boolean
    Dim vIndex As Long
    Dim vMissionConts() As String
    
    vMissionConts = Split(gMissions(pMissionIndex).TargetConts, ",")
    
    For vIndex = 0 To UBound(vMissionConts)
        HoldAllConts = True
        If Not HoldContinent(CInt(vMissionConts(vIndex)), pPlayer) Then
            HoldAllConts = False
            Exit For
        End If
    Next
End Function

'Return true if passed player holds the passed continent. Used to check for win.
Private Function HoldContinent(pContinent As Integer, pPlayer As Integer) As Boolean
    Dim vIndex As Integer
    
    HoldContinent = True
    For vIndex = Continents(pContinent - 1).FirstCountry To Continents(pContinent - 1).LastCountry
        If gCountryOwner(vIndex) <> pPlayer Then
            HoldContinent = False
            Exit For
        End If
    Next
End Function

'Return True if passed player owns a third continent.
'Not currently used.
Private Function HoldAnyThirdCont(pCont1 As Integer, pCont2 As Integer, pPlayer As Integer) As Boolean
    Dim vContinent As Integer
    
    For vContinent = 1 To 6
        If (vContinent <> pCont1) And (vContinent <> pCont2) _
        And HoldContinent(vContinent, pPlayer) Then
            HoldAnyThirdCont = True
            Exit Function
        End If
    Next vContinent
    HoldAnyThirdCont = False
End Function



