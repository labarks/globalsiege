Attribute VB_Name = "modDice"
Option Explicit

Public Const cMaxNumberOfDice As Long = 5

Public gDiceArray(19) As Integer            'Dice results. 1-3 attack, 4-5 defend

Dim gDiceOdds(29) As String

'Initialise the gDiceOdds array by populating it with previously worked out dice probabilities.
Public Sub InitialiseDiceOddsArray()
    gDiceOdds(0) = "50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50"
    gDiceOdds(1) = "58,63,67,70,70,53,58,62,66,66,49,55,60,63,63,47,53,58,61,61,46,52,56,60,60"
    gDiceOdds(2) = "42,47,51,53,53,37,42,45,47,47,33,38,40,42,42,30,34,37,39,39,28,32,34,36,36"
    gDiceOdds(3) = "50,57,61,63,63,43,50,54,57,57,39,46,50,53,53,37,43,47,50,50,35,41,45,48,48"
    gDiceOdds(4) = "50,55,58,59,59,45,50,53,55,55,42,47,50,52,52,41,45,48,50,50,39,44,46,48,48"
    gDiceOdds(5) = "50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50"
    gDiceOdds(6) = "58,42,34,29,26,75,61,46,37,32,83,75,63,49,40,87,82,75,65,52,91,87,82,75,66"
    gDiceOdds(7) = "42,25,17,13,9,58,39,25,18,13,66,54,37,25,18,71,62,51,35,25,74,68,60,48,34"
    gDiceOdds(8) = "50,31,21,15,11,69,50,32,22,16,79,68,50,33,23,85,78,67,50,34,89,84,77,66,50"
    gDiceOdds(9) = "50,36,29,25,22,64,50,38,31,27,71,62,50,40,33,75,69,60,50,41,78,73,67,59,50"
    gDiceOdds(10) = "50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50"
    gDiceOdds(11) = "58,54,66,70,70,61,56,69,73,74,51,45,59,64,64,48,41,56,61,61,46,39,54,60,60"
    gDiceOdds(12) = "42,39,49,52,52,46,44,55,59,61,34,31,41,44,44,30,27,36,39,39,28,24,33,36,36"
    gDiceOdds(13) = "50,47,59,63,63,53,50,63,68,70,41,36,50,55,55,37,32,45,50,50,35,29,43,48,48"
    gDiceOdds(14) = "50,48,56,59,59,52,50,60,64,65,44,40,50,53,54,41,36,47,50,50,40,34,45,48,48"
    gDiceOdds(15) = "50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50"
    gDiceOdds(16) = "58,32,32,29,26,82,59,57,55,54,84,59,63,51,43,88,60,73,65,52,91,59,79,75,66"
    gDiceOdds(17) = "42,19,16,12,9,68,41,41,40,41,68,43,37,27,21,71,45,49,35,25,74,46,57,48,34"
    gDiceOdds(18) = "50,22,19,15,11,79,50,49,47,47,81,51,50,36,27,85,53,64,50,34,89,53,73,66,50"
    gDiceOdds(19) = "50,28,28,25,22,72,50,49,48,48,72,51,50,41,35,75,52,59,50,41,78,52,65,59,50"
    gDiceOdds(20) = "50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50"
    gDiceOdds(21) = "58,69,68,70,70,43,56,58,61,61,48,58,59,62,62,47,57,58,61,61,46,55,57,60,60"
    gDiceOdds(22) = "42,57,52,53,53,31,43,42,43,44,32,42,41,42,42,30,39,38,39,39,28,36,35,36,36"
    gDiceOdds(23) = "50,66,63,63,64,34,50,49,51,51,38,51,50,52,52,37,49,48,50,50,35,47,46,48,48"
    gDiceOdds(24) = "50,62,59,60,60,38,50,49,51,51,41,51,50,52,52,40,49,48,50,50,39,48,47,48,48"
    gDiceOdds(25) = "50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50"
    gDiceOdds(26) = "58,56,39,30,30,56,59,40,32,28,77,77,63,48,40,86,85,75,65,52,86,88,82,75,66"
    gDiceOdds(27) = "42,44,23,14,14,44,41,23,15,12,61,60,37,25,18,70,68,52,35,26,70,72,60,48,34"
    gDiceOdds(28) = "50,50,27,17,16,50,50,28,18,14,73,72,50,32,24,83,82,68,50,35,84,86,76,65,50"
    gDiceOdds(29) = "50,50,33,26,26,50,50,34,27,24,67,66,50,39,33,74,73,61,50,41,74,76,67,59,50"
End Sub

'Return the dice odds for the chosen combination of dice settings.
'The return is a comma delimited string of 25 dice odds, one for
'each dice combination.
'Format: "A1D1,A1D2,A1D3,...,A5D5"
Public Function GetDiceOdds() As String
    Dim vIndex As Long
    Dim vSetting As Long
    
    On Error Resume Next
    
    'Dice rules.
    For vIndex = 0 To TheMainForm.optDiceRules.Count - 1
        If TheMainForm.optDiceRules(vIndex).Value Then
            vSetting = vIndex
            Exit For
        End If
    Next
    
    'Sort dice.
    If TheMainForm.chkSortDice.Value = vbChecked Then
        vSetting = vSetting + 5
    End If
    
    'All same.
    For vIndex = 0 To TheMainForm.optDiceSame.Count - 1
        If TheMainForm.optDiceSame(vIndex) Then
             vSetting = vSetting + (vIndex * 10)
             Exit For
        End If
    Next
    
    GetDiceOdds = gDiceOdds(vSetting)
End Function

'Work out the dice odds for all setup screen dice combinations and save
'to "DiceOdds.txt" in the config directory. Used only during development.
Public Function CrunchDiceOdds() As String
    Dim vSetting As Long
    Dim vAttackDice As Integer
    Dim vDefenceDice As Integer
    
    TheMainForm.Show
    CrunchDiceOdds = ""
    
    For vSetting = 0 To 29
        
        'Rules.
        TheMainForm.optDiceRules(vSetting Mod 5).Value = vbChecked
        
        'Sort dice.
        TheMainForm.chkSortDice.Value = vSetting \ 5 Mod 2
        
        'All same.
        TheMainForm.optDiceSame(vSetting \ 10).Value = vbChecked
        
        CrunchDiceOdds = CrunchDiceOdds & "gDiceOdds(" & vSetting & ") = """
        For vAttackDice = 1 To 5
            TheMainForm.udDiceThrown(0) = vAttackDice
            For vDefenceDice = 1 To 5
                TheMainForm.udDiceThrown(1) = vDefenceDice
                DoEvents
                CrunchDiceOdds = CrunchDiceOdds & WorkOutDiceOdds(6000000) & ","
            Next
        Next
        CrunchDiceOdds = Mid(CrunchDiceOdds, 1, Len(CrunchDiceOdds) - 1) & """" & vbCrLf
    Next
    Call SaveConfigFile("DiceOdds.txt", CrunchDiceOdds)
    Debug.Print CrunchDiceOdds
End Function

'Work out dice odds by rolling the dice behind the sceens.
'Return the attack winning odds out of 100.
Public Function WorkOutDiceOdds(Optional pDepth As Long = 10000) As Integer
    Dim vAttackDice(cMaxNumberOfDice - 1) As Integer
    Dim vDefenceDice(cMaxNumberOfDice - 1) As Integer
    Dim vTestNumber As Long
    Dim vAttackUnits As Integer
    Dim vDefendUnits As Integer
    Dim vAttackLoss As Long
    Dim vDefendLoss As Long
    Dim vTotalLoss As Long
    Dim vStartTime As Long
    
    'No dice option.
    If TheMainForm.optDiceRules(0).Value Then
        'TheMainForm.lblOddsCalculated.Caption = "50 : 50"
        WorkOutDiceOdds = 50
    Else
        vStartTime = Round(CDbl(Time) * 100000)
        'Call GenRepeat4(True)
        For vTestNumber = 0 To pDepth
            vAttackUnits = 10
            vDefendUnits = 10
            Call RollDice(vAttackUnits, vDefendUnits, vAttackDice, vDefenceDice, False)
            'Sort dice if selected in the setup screen.
            If TheMainForm.chkSortDice.Value = vbChecked Then
                Call BubbleSort(vAttackDice, False)
                Call BubbleSort(vDefenceDice, False)
            End If
            Call DiceBattleDamage(vAttackUnits, vDefendUnits, vAttackDice, vDefenceDice)
            If vAttackUnits <= 0 Or vDefendUnits <= 0 Then
                Exit For
            End If
            vAttackLoss = vAttackLoss + (10 - vAttackUnits)
            vDefendLoss = vDefendLoss + (10 - vDefendUnits)
            
            'Limit time to 5 seconds.
            'If vTestNumber Mod 1000 = 5 Then
            '    If Round(CDbl(Time) * 100000) >= vStartTime + 5 Then
            '        Exit For
            '    End If
            'End If
        Next
         
        vTotalLoss = vAttackLoss + vDefendLoss
        'TheMainForm.lblOddsCalculated.Caption = CStr(Round(vDefendLoss / vTotalLoss * 100)) _
                    & " : " & CStr(Round(vAttackLoss / vTotalLoss * 100))
        
        WorkOutDiceOdds = Round(vDefendLoss / vTotalLoss * 100)
    End If
End Function

'Roll dice as dictated by the user's selected rules.
Public Sub AttackRollDice(pAttackUnits As Integer, pDefendUnits As Integer, _
pAttackDice() As Integer, pDefenceDice() As Integer)
    Dim vAttackUnits As Integer
    Dim vDefendUnits As Integer
    
    If Not TheMainForm.optDiceRules(0).Value Then
    
        'Roll dice.
        Call RollDice(pAttackUnits, pDefendUnits, pAttackDice, pDefenceDice)
        
        'Sort dice if selected in the detup screen.
        If TheMainForm.chkSortDice.Value = vbChecked Then
            Call BubbleSort(pAttackDice, False)
            Call BubbleSort(pDefenceDice, False)
        End If
    End If
End Sub

'Roll dice following rules chosen on the setup screen and change passed attack
'and defence points resulting from the rolls. The resulting dice are in the passed
'pDice array.
Private Sub RollDice(pAttackPoints As Integer, pDefencePoints As Integer, _
pAttackDice() As Integer, pDefenceDice() As Integer, Optional pRepeatRandom As Boolean = False)
    Dim vAttackDiceCount As Integer
    Dim vDefenceDiceCount As Integer
    Dim vIndex As Integer
    
    'Check attacker and defender have units left.
    If pAttackPoints = 0 Then
        Exit Sub
    ElseIf pDefencePoints = 0 Then
        Exit Sub
    End If
    
    'How many attack dice to roll.
    If pAttackPoints <= TheMainForm.udDiceThrown(0).Value Then
        vAttackDiceCount = pAttackPoints
    Else
        vAttackDiceCount = TheMainForm.udDiceThrown(0).Value
    End If
    
    'How many defence dice to roll. This may change if Optimize Defence Dice is selected.
    If pDefencePoints <= TheMainForm.udDiceThrown(1).Value Then
        vDefenceDiceCount = pDefencePoints
    Else
        vDefenceDiceCount = TheMainForm.udDiceThrown(1).Value
    End If
    
    'Roll attackers dice.
    For vIndex = 0 To vAttackDiceCount - 1
        If pRepeatRandom Then
            pAttackDice(vIndex) = Int(GenRepeat4 * 6 + 1)
        Else
            pAttackDice(vIndex) = Int(GenRandom4 * 6 + 1)
        End If
    Next
    
    'Roll defence dice.
    For vIndex = 0 To vDefenceDiceCount - 1
        If pRepeatRandom Then
            pDefenceDice(vIndex) = Int(GenRepeat4 * 6 + 1)
        Else
            pDefenceDice(vIndex) = Int(GenRandom4 * 6 + 1)
        End If
    Next
    
End Sub

'Compare dice and increase/decrease scores as required.
Private Function BattleCompareDice(pAttackUnits As Integer, pDefendUnits As Integer, _
pAttackDie As Integer, pDefenceDie As Integer) As Integer
    Dim vAttackLoss As Integer
    Dim vDefenceLoss As Integer
    
    'Attacker won.
    If pAttackDie > pDefenceDie Then
        pDefendUnits = pDefendUnits - 1
    
    'Defender won.
    ElseIf pAttackDie < pDefenceDie Then
        pAttackUnits = pAttackUnits - 1
    
    'Draw, both attack and defend dice are the dame.
    ElseIf pAttackDie = pDefenceDie Then
        
        'Attacked wins draw.
        If TheMainForm.optDiceRules(1) Then
            pDefendUnits = pDefendUnits - 1
            
        'Defender wins draw.
        ElseIf TheMainForm.optDiceRules(2) Then
            pAttackUnits = pAttackUnits - 1
            
        'Both retreat.
        ElseIf TheMainForm.optDiceRules(3) Then
            'Do nothing.
            
        'Both loose.
        ElseIf TheMainForm.optDiceRules(4) Then
            
            'Ensure units don't annihilate each other.
            If pAttackUnits = 1 And pDefendUnits = 1 Then
                pAttackUnits = pAttackUnits - 1
            Else
                pAttackUnits = pAttackUnits - 1
                pDefendUnits = pDefendUnits - 1
            End If
        
        End If
    End If
End Function

'No dice. Instead of throwing dice, results are determined on a 1 for 1 basis.
'That means that whoever has the most units will win.
'Return TRUE if this is the case.
Private Function CheckNoDiceDice(pAttackUnits As Integer, pDefendUnits As Integer, _
pAttackDice() As Integer, pDefenceDice() As Integer) As Boolean
    Dim vAttackUnits As Integer
    Dim vDefendUnits As Integer
    
    'No dice. Instead of throwing dice, results are determined on a 1 for 1 basis.
    'That means that whoever has the most units will win.
    If TheMainForm.optDiceRules(0).Value Then
        vAttackUnits = pAttackUnits
        vDefendUnits = pDefendUnits
        
        'Attacker wins.
        If vAttackUnits > vDefendUnits Then
            pAttackUnits = vAttackUnits - vDefendUnits
            pDefendUnits = 0
        
        'Defender wins.
        Else
        
            'The last unit is won by the defender. This is to ensure that at least
            'one unit is left to occupy the defending territory.
            If vAttackUnits = vDefendUnits Then
                vDefendUnits = vDefendUnits + 1
            End If
            
            pDefendUnits = vDefendUnits - vAttackUnits
            pAttackUnits = 0
            
        End If
        CheckNoDiceDice = True
    Else
        CheckNoDiceDice = False
    End If
End Function

'Take action if all the same dice were thrown. Return TRUE if something was done.
Private Function CheckAllSameDice(pAttackUnits As Integer, pDefendUnits As Integer, _
pAttackDice() As Integer, pDefenceDice() As Integer) As Boolean
    Dim vAttackAllSame As Boolean
    Dim vDefenceAllSame As Boolean
    
    CheckAllSameDice = False
    
    'Check selected rules for all same dice.
    If Not TheMainForm.optDiceSame(0).Value Then
        vAttackAllSame = IsArrayAllSame(pAttackDice)
        vDefenceAllSame = IsArrayAllSame(pDefenceDice)
        
        'Carry on the attack if both attacker and defender throw all same dice.
        If Not (vAttackAllSame And vDefenceAllSame) Then
            
            'Attacker's dice are all the same.
            If vAttackAllSame Then
                
                'Attacker instantly wins the attack.
                If TheMainForm.optDiceSame(1).Value Then
                    pDefendUnits = pDefendUnits - CountNonzeroArrayElements(pDefenceDice)
                
                'Attacker instantly looses the attack.
                ElseIf TheMainForm.optDiceSame(2).Value Then
                    pAttackUnits = pAttackUnits - CountNonzeroArrayElements(pAttackDice)
                    
                End If
                
                CheckAllSameDice = True
            
            'Defender's dice are all the same.
            ElseIf vDefenceAllSame Then
            
                'Defender instantly wins the attack.
                If TheMainForm.optDiceSame(1).Value Then
                    pAttackUnits = pAttackUnits - CountNonzeroArrayElements(pAttackDice)
                
                'Defender instantly looses the attack.
                ElseIf TheMainForm.optDiceSame(2).Value Then
                    pDefendUnits = pDefendUnits - CountNonzeroArrayElements(pDefenceDice)
                    
                End If
                
                CheckAllSameDice = True
                
            End If
        End If
    End If
End Function

'Actions for sorted dice. Return TRUE if something was done.
Private Function CheckSortedDice(pAttackUnits As Integer, pDefendUnits As Integer, _
pAttackDice() As Integer, pDefenceDice() As Integer) As Boolean
    Dim vAttackIndex As Long
    Dim vDefendIndex As Long
    
    If TheMainForm.chkSortDice.Value = vbChecked Then
        For vAttackIndex = 0 To CountNonzeroArrayElements(pAttackDice) - 1
            If pAttackDice(vAttackIndex) = 0 Or pDefenceDice(vAttackIndex) = 0 Then
                
                'Attacker or defender has run out of dice do compare. Attack ends.
                Exit For
                
            Else
                
                Call BattleCompareDice(pAttackUnits, pDefendUnits, _
                    pAttackDice(vAttackIndex), pDefenceDice(vAttackIndex))
                    
            End If
        Next
        CheckSortedDice = True
    Else
        CheckSortedDice = False
    End If
End Function

'Actions for unsorted dice. Return TRUE if something was done.
Private Function CheckUnsortedDice(pAttackUnits As Integer, pDefendUnits As Integer, _
pAttackDice() As Integer, pDefenceDice() As Integer) As Boolean
    Dim vAttackIndex As Long
    Dim vDefendIndex As Long
    Dim vRound As Long
    Dim vAttackUnits As Integer
    Dim vDefendUnits As Integer
    
    If TheMainForm.chkSortDice.Value <> vbChecked Then
        vAttackIndex = 0
        vDefendIndex = 0
        For vRound = 0 To CountNonzeroArrayElements(pAttackDice) + CountNonzeroArrayElements(pDefenceDice)
            
            'Ensure indicies do not go out of bounds.
            If vAttackIndex >= cMaxNumberOfDice Or vDefendIndex >= cMaxNumberOfDice - 1 Then
                Exit For
            Else
                If pAttackDice(vAttackIndex) = 0 Or pDefenceDice(vDefendIndex) = 0 Then
                    
                    'Attacker or defender has run out of dice do compare.
                    Exit For
                Else
                    
                    'Compare dice, modify scores, update the audits and eliminate the loosing die.
                    vAttackUnits = pAttackUnits
                    vDefendUnits = pDefendUnits
                    Call BattleCompareDice(pAttackUnits, pDefendUnits, _
                        pAttackDice(vAttackIndex), pDefenceDice(vDefendIndex))
                    
                    'Discover if the attacker lost the battle and eliminate their die.
                    'This method may seem like a bit of a hack but is saves passing
                    'heaps of variables through functions and less dependent on global
                    'variables making the code more loosley coupled. Note that the next
                    'three if statements are NOT mutually exclusive so we need to test
                    'each one. No ElseIf here matey.
                    If pAttackUnits < vAttackUnits Then
                        vAttackIndex = vAttackIndex + 1
                    End If
                    
                    'Discover if the defender lost the battle and eliminate their die.
                    If pDefendUnits < vDefendUnits Then
                        vDefendIndex = vDefendIndex + 1
                    End If
                    
                    'If both retreated, eliminate dice for both parties.
                    If pAttackUnits = vAttackUnits And pDefendUnits = vDefendUnits Then
                        vAttackIndex = vAttackIndex + 1
                        vDefendIndex = vDefendIndex + 1
                    End If
                    
                End If
            End If
        Next
        CheckUnsortedDice = True
    Else
        CheckUnsortedDice = False
    End If
End Function

'Dice have been rolled. This function changes scores of passed attacker and defender
'depending on the dice thrown and rules selected.
Public Sub DiceBattleDamage(pAttackUnits As Integer, pDefendUnits As Integer, _
pAttackDice() As Integer, pDefenceDice() As Integer)
    
    'Check and take action if no dice selected.
    If CheckNoDiceDice(pAttackUnits, pDefendUnits, pAttackDice(), pDefenceDice) Then
        Exit Sub
    
    'Check selected rules for all same dice.
    ElseIf CheckAllSameDice(pAttackUnits, pDefendUnits, pAttackDice(), pDefenceDice) Then
        Exit Sub
    
    'Compare sorted dice.
    ElseIf CheckSortedDice(pAttackUnits, pDefendUnits, pAttackDice(), pDefenceDice) Then
        Exit Sub
    
    'Compare un-sorted dice.
    ElseIf CheckUnsortedDice(pAttackUnits, pDefendUnits, pAttackDice(), pDefenceDice) Then
        Exit Sub
    End If
    
End Sub

'Save the empty dice area from the recently cleaned mask (Map1) to hidden picture boxe.
'This affects dice only. SnapEmptyDCardArea() does the same thing for cards.
Public Sub SnapEmptyDiceArea()
    Dim vDummy As Long
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
    
        'Dice.
        Mask4.pctClearDice.Cls
        Mask4.pctClearDice.Print ""
        vDummy = BitBlt(Mask4.pctClearDice.hdc, 0, 0, gMsk.DiceWidth, gMsk.DiceHeight, _
                Mask4.Map1.hdc, gMsk.DiceLeft, gMsk.DiceTop, vbSrcCopy)
    End If
End Sub

'Copy a clean image of the dice area to the map. This clean image was saved earlier
'by function SnapEmptyDiceArea().
Public Sub ClearDiceFromBoard()
    Dim vDummy As Long
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
    
        'Dice.
        vDummy = BitBlt(Mask4.Map1.hdc, gMsk.DiceLeft, gMsk.DiceTop, _
                gMsk.DiceWidth, gMsk.DiceHeight, Mask4.pctClearDice.hdc, 0, 0, vbSrcCopy)
    End If
End Sub

'Display the dice on the board in preset locations.
Public Sub DisplayDiceOnBoard(pAttackDice() As Integer, pDefenceDice() As Integer)
    Dim vCntr As Integer
    Dim vAtkDiceThrown As Integer
    Dim vDefDiceThrown As Integer
    
    'Count the number of dice thrown.
    vAtkDiceThrown = CountNonzeroArrayElements(pAttackDice)
    vDefDiceThrown = CountNonzeroArrayElements(pDefenceDice)
    
    If vAtkDiceThrown > 0 And vDefDiceThrown > 0 Then
        Call ClearDiceFromBoard
    End If
    
    'Print all attack dice.
    For vCntr = 0 To vAtkDiceThrown - 1
        If pAttackDice(vCntr) > 0 Then
            Call PrintDieHere(pAttackDice(vCntr), 0, _
                CInt(gMsk.DieAttackDestX(vAtkDiceThrown, vCntr + 1)), _
                CInt(gMsk.DieAttackDestY(vAtkDiceThrown, vCntr + 1)))
        Else
            Exit For
        End If
    Next
    
    'Print all defence dice.
    For vCntr = 0 To vDefDiceThrown - 1
        If pDefenceDice(vCntr) > 0 Then
            Call PrintDieHere(pDefenceDice(vCntr), 1, _
                CInt(gMsk.DieDefendDestX(vDefDiceThrown, vCntr + 1)), _
                CInt(gMsk.DieDefendDestY(vDefDiceThrown, vCntr + 1)))
        Else
            Exit For
        End If
    Next
End Sub

'Display chosen dice at specified position.
'pDiceType is 0 for attack die and 1 for defence die.
Private Sub PrintDieHere(pDiceNumber As Integer, pDiceType As Long, pPosX As Long, pPosY As Long)
    Dim pSourceX As Long
    Dim pSourceY As Long
    Dim dummy As Long
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
    
        If pDiceNumber = 0 Then
            Exit Sub
        End If
        
        With gMsk
        
        'Select die colour and work out the source image from the Mask4.
        If pDiceType = 0 Then
            'Attack die.
            pSourceY = (.DieHeight * 2) + (CInt(GenRandom4) * .DieHeight + 1)
        Else
            'Defence die.
            pSourceY = CInt(GenRandom4) * (.DieHeight + 1)
        End If
        
        pSourceX = (.DieWidth) * (pDiceNumber - 1)
        
        dummy = BitBlt(Mask4.Map1.hdc, pPosX + .DiceLeft, _
                pPosY + .DiceTop, .DieWidth, .DieHeight, _
                Mask4.pctDice.hdc, pSourceX, pSourceY, vbSrcCopy)
        End With
        
        'Make sure the viewport gets refreshed.
        TheMainForm.gSyncViewportNeeded = True
    End If
End Sub

