Attribute VB_Name = "modCards"
Option Explicit

'---------------------------
'   Card functions
'---------------------------

Public gLastCardClicked As Integer         'The last card clicked by the user.
Public gCurrentCardValue As Integer         'Value of cards
Public gMaxCardValue As Integer             'Maximum value for cards
Public gCardDeck(3) As Integer            'Cards left in pack

Dim mouseCardsX As Integer              'Mouse position over card box

'Check and take action if cards area has been clicked by a human player.
Public Sub CardsClicked()
    
    'Not actually used any more but left here incase of
    'reactivation one day.
    If gPauseActive Then
        Call TheMainForm.ActivatePauseMode(False)
    End If
    
    'Check if in win mode.
    If gCurrentMode = 13 Or gCurrentMode = 18 Then
        
        TheMainForm.Timer2.Enabled = False
    
    'Check if the human player can look at the card.
    ElseIf Not TheMainForm.gMapSetupLock _
    And Not TheMainForm.gWarRestartLock _
    And TheMainForm.GetPlayerController(gPlayerTurn) = 0 _
    And gPlayerID(gPlayerTurn).card(0) <> 0 Then
        
        Call LookAtCards
    
    End If
    
End Sub

'Return the X position of the card the mouse is over relative to pXOffset. The offset is needed
'because the upper row of cards, the Vulture Cards, are slightly offset to the right of the main cards.
Private Function GetCardXHitPos(pXOffSet As Integer) As Integer
    Dim vOffsetMouseX As Integer
    Dim vIndex As Integer
    
    vOffsetMouseX = (TheMainForm.gCurrentMousePosX / TheMainForm.GetPictureMaskRatioX) - pXOffSet
    GetCardXHitPos = 0
    
    With gMsk
    For vIndex = 1 To 5
        If vOffsetMouseX > .CrdDestX(vIndex) _
        And vOffsetMouseX < .CrdDestX(vIndex) + (.CrdSnglWidth - .CrdSnglSrcBuffer) Then
            GetCardXHitPos = vIndex
            Exit For
        End If
    Next
    End With
End Function

'Return TRUE if the mouse is over a card and update the global
'variable "gLastCardClicked" with the card number.
Public Function GetCardHitPosition() As Boolean
    
    'Check first row of cards.
    With TheMainForm
    If .gCurrentMousePosY > gMsk.CrdMainTop * .gPictureMaskRatioY _
    And .gCurrentMousePosY < (gMsk.CrdMainTop + gMsk.CrdMainHeight) * .gPictureMaskRatioY Then
        gLastCardClicked = GetCardXHitPos(gMsk.CrdMainLeft)
    
    'Check second (Vulture) row of cards (above).
    ElseIf .gCurrentMousePosY > gMsk.CrdVultTop * .gPictureMaskRatioY _
    And .gCurrentMousePosY < (gMsk.CrdVultTop + gMsk.CrdVultHeight) * .gPictureMaskRatioY Then
        gLastCardClicked = GetCardXHitPos(gMsk.CrdVultLeft)
        If gLastCardClicked <> 0 Then
            'Add 5 to indicate that the last card clicked was a vulture card.
            gLastCardClicked = gLastCardClicked + 5
        End If
    Else
        gLastCardClicked = 0
    End If
    
    'Confirm player has a card in this position.
    If gLastCardClicked > 0 Then
        If gPlayerID(gPlayerTurn).card(gLastCardClicked - 1) = 0 Then
            gLastCardClicked = 0
        End If
    End If
    GetCardHitPosition = (gLastCardClicked > 0)
    End With
End Function

'Get the card mode selected from the setup screen. Counterpart to SetCardMode().
'0 = none
'1 = fixed
'2 = increasing
Public Function GetCardMode() As Integer
    GetCardMode = Abs(CInt(TheMainForm.optCardMode(1).Value + (TheMainForm.optCardMode(2).Value * 2)))
End Function

'Set the card mode to the passed value. Counterpart to GetCardMode().
'0 = none
'1 = fixed
'2 = increasing
Public Sub SetCardMode(pCardMode As Integer)
    On Error GoTo ErrHand:
    TheMainForm.optCardMode(pCardMode).Value = True
    Exit Sub
ErrHand:
    'Set to fixed value if a bad card mode was passed.
    TheMainForm.optCardMode(1).Value = True
    Exit Sub
End Sub

'Human player is checking cards. Print "Checking Cards" if required.
'Turn cards over if hidden. Preselect cards if not already in card
'mode otherwise select or unselect the clicked card as required.
Public Sub LookAtCards()
    Dim vLastMode As Integer
    Dim vNumberOfCardsPicked As Integer                 'how many cards picked up
    
    With TheMainForm
    
    'Safety catch.
    If gPickedUpUnits > 0 Then
        Exit Sub
    End If
    vLastMode = gCurrentMode
    
    'If player is not already looking at their cards.
    If Not (gCurrentMode = 4 Or gCurrentMode = 5 Or gCurrentMode = 6) Then
        Call .ColorCountryUnderAttack(0)
        gCurrentMode = 5
        vNumberOfCardsPicked = 1
        Call CardOutOfHand(gPlayerTurn)
        Call .ToggleKeys(False)
        Call .ToglleCardKeys(True)
        
        'Print "Checking cards" in the info box.
        .InfoBoxPrint 0
        .InfoBoxPrnCR 9, gPlayerTurn
        .InfoBoxPrint 7
        .InfoBoxPrint 5                       'bold
        .InfoBoxPrnCR 1, 130                  'checking cards
        .InfoBoxPrint 6                       'normal
        
        'If increasing card value option selected, print the current card value.
        If GetCardMode = 2 Then
            .InfoBoxPrnCR 7
            .InfoBoxPrint 1, 181                  'current value
            .InfoBoxPrint 2, gCurrentCardValue
            .InfoBoxPrint 1, 102                  'units
        End If
        
        'Turn cards over if the hidden option is selected on the setup screen.
        If .chkCardsHidden.Value = vbChecked Then
            Call DrawAllCards
        End If
        
        'Preselect cards for human players.
        If (vLastMode = 2 Or vLastMode = 20) _
        And (gLastCardClicked = 0 Or .chkCardsHidden.Value = vbChecked) Then
            Call PreSelectCardsForHuman
        End If
        
        'Jump out here if hidden cards have just been turned over.
        If .chkCardsHidden.Value = vbChecked Then
            .cmdExchange.Enabled = IsValidCardSetSelected
            Exit Sub
        End If
    End If
    
    'Pre select cards for humans.
    If gCurrentMode = 6 Then
        gCurrentMode = 5
        Call DrawAllCards
        Call PreSelectCardsForHuman
        .cmdExchange.Enabled = IsValidCardSetSelected
    Else
        Call DrawAllCards
        vNumberOfCardsPicked = SelectACard(vNumberOfCardsPicked)
        .cmdExchange.Enabled = IsValidCardSetSelected
    End If
    
    End With
End Sub

'Unselect all the passed player's cards.
Public Sub CardOutOfHand(pPlayer As Integer)
    Dim vIndex As Integer
    
    'For each card.
    For vIndex = 0 To 9
        gPlayerID(pPlayer).pickedCards(vIndex) = True
    Next
End Sub

'Toggle the selection status of the passed card. If already selected, make it
'unselected and if unselected, make the passed card selected.
'Called by LookAtCards() to select or unselect the clicked card.
'For some reason, selected cards are marked as FALSE.
Private Function SelectACard(pNumberOfCardsPicked As Integer) As Integer
    
    If gLastCardClicked > 0 Then
        If gPlayerID(gPlayerTurn).card(gLastCardClicked - 1) > 0 Then
            
            'If the card clicked is already selected, unselect it
            'and reduce the number of cards picked.
            If Not gPlayerID(gPlayerTurn).pickedCards(gLastCardClicked - 1) Then
                gPlayerID(gPlayerTurn).pickedCards(gLastCardClicked - 1) = True
                pNumberOfCardsPicked = pNumberOfCardsPicked - 1
                Call DrawAllCards
                Exit Function
            
            Else
                
                'The clicked card is not already selected.
                Call DrawBigCard(gPlayerID(gPlayerTurn).card(gLastCardClicked - 1), _
                        gLastCardClicked, False)
                pNumberOfCardsPicked = pNumberOfCardsPicked + 1
                gPlayerID(gPlayerTurn).pickedCards(gLastCardClicked - 1) = False
                'gCurrentMode = 2
                
            End If
        End If
    End If
End Function

'Draw all cards held by the current player. Clear main card area and
'vulture card area using clean snap shots taken earlier then draw
'all cards held by the current player including vulture cards.
Public Sub DrawAllCards()
    Dim vIndex As Integer
    Dim vDummy As Long
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
    
        'Clear main card screen using a clean snap shot taken earlier.
        Call ClearMainCardsArea
        
        If TheMainForm.chkCardsVulture.Value = vbChecked Then
            
            'Clear vulture card screen using a clean snap shot taken earlier.
            vDummy = BitBlt(Mask4.Map1.hdc, gMsk.CrdVultLeft, gMsk.CrdVultTop, _
                    gMsk.CrdVultWidth, gMsk.CrdVultHeight, _
                    Mask4.pctVultureCards.hdc, 0, 0, vbSrcCopy)
        End If
        
        'Print each card.
        For vIndex = 1 To 10
            
            'Print all cards until we reach a vacant card.
            If gPlayerID(gPlayerTurn).card(vIndex - 1) <> 0 Then
                
                'Draw the card.
                Call DrawBigCard(gPlayerID(gPlayerTurn).card(vIndex - 1), _
                        vIndex, gPlayerID(gPlayerTurn).pickedCards(vIndex - 1))
            Else
                
                'Vacant spot reached.
                Exit For
                
            End If
            
        Next
        
        'Make sure the viewport gets refreshed.
        TheMainForm.gSyncViewportNeeded = True
    End If
    
End Sub

'Save the empty card areas from the recently cleaned mask (Map1) to hidden picture boxes.
'This affects the Main cards, Vulture cards, Little cards across the top. Function
'SnapEmptyDiceArea() does the same thing for dice.
Public Sub SnapEmptyCardArea()
    Dim vDummy As Long
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
    
        'Main cards.
        Mask4.pctMainCards.Cls
        Mask4.pctMainCards.Print ""
        vDummy = BitBlt(Mask4.pctMainCards.hdc, 0, 0, gMsk.CrdMainWidth, gMsk.CrdMainHeight, _
                Mask4.Map1.hdc, gMsk.CrdMainLeft, gMsk.CrdMainTop, vbSrcCopy)
        
        'Vulture cards.
        Mask4.pctVultureCards.Cls
        Mask4.pctVultureCards.Print ""
        vDummy = BitBlt(Mask4.pctVultureCards.hdc, 0, 0, gMsk.CrdVultWidth, gMsk.CrdVultHeight, _
                Mask4.Map1.hdc, gMsk.CrdVultLeft, gMsk.CrdVultTop, vbSrcCopy)
        
        'Little cards across the top.
        Mask4.pctLittleCards.Cls
        Mask4.pctLittleCards.Print ""
        vDummy = BitBlt(Mask4.pctLittleCards.hdc, 0, 0, Mask4.Map1.ScaleWidth, gMsk.LittleCardHeight, _
                Mask4.Map1.hdc, 0, gMsk.LittleCardTop, vbSrcCopy)
        
        'Dice.
        Mask4.pctClearDice.Cls
        Mask4.pctClearDice.Print ""
        vDummy = BitBlt(Mask4.pctClearDice.hdc, 0, 0, gMsk.DiceWidth, gMsk.DiceHeight, _
                Mask4.Map1.hdc, gMsk.DiceLeft, gMsk.DiceTop, vbSrcCopy)
    End If
End Sub

'Clear cards from across the top of the map.
Public Sub CleaLittleCards()
    Dim vDummy As Integer
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
        vDummy = BitBlt(Mask4.Map1.hdc, 0, gMsk.LittleCardTop, _
                Mask4.Map1.ScaleWidth, gMsk.LittleCardHeight, _
                Mask4.pctLittleCards.hdc, 0, 0, vbSrcCopy)
    End If
End Sub

'Draw vulture cards aquired from killing a player.
Private Sub DrawVultureCards()
    Dim vDummy As Long
    Dim vIndex As Integer
    
    'Only if vulture cards are on and not running in headless mode.
    If TheMainForm.chkCardsVulture.Value = vbChecked _
    And Not gHeadlessMode Then
    
        'Clear the extra card section by copying a recent copy of that area.
        vDummy = BitBlt(Mask4.Map1, gMsk.CrdVultLeft, gMsk.CrdVultTop, _
                gMsk.CrdVultWidth, gMsk.CrdVultHeight, _
                Mask4.pctVultureCards.hdc, 0, 0, vbSrcCopy)
        
        'Make sure the viewport gets refreshed.
        TheMainForm.gSyncViewportNeeded = True
        
        'Draw all the extra cards until we reach a vacant card.
        For vIndex = 6 To 10
            
            If gPlayerID(gPlayerTurn).card(vIndex - 1) <> 0 Then
                
                'Draw the card.
                Call DrawBigCard(gPlayerID(gPlayerTurn).card(vIndex - 1), _
                        vIndex, gPlayerID(gPlayerTurn).pickedCards(vIndex - 1))
            Else
                
                'Vacant spot reached.
                Exit For
                
            End If
        Next
    End If
End Sub

'Draw one card to the correct position.
'pWhichCard = 0 is turned card.
'pNotSelected means that this card is selected when false.
Public Sub DrawBigCard(pWhichCard As Integer, vCardPosition As Integer, pNotSelected As Boolean)
    Dim vSourceX As Integer
    Dim vSourceY As Integer
    Dim vDummy As Integer
    Dim vPosFromLeft As Integer
    Dim vDestX As Long
    Dim vDestY As Long
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
    
        'Work out card position on Picture1.
        If vCardPosition > 5 Then
            
            'Vulture card.
            vPosFromLeft = vCardPosition - 5
            vDestX = gMsk.CrdVultLeft
            vDestY = gMsk.CrdVultTop
        
        Else
            
            'Normal card.
            vPosFromLeft = vCardPosition
            vDestX = gMsk.CrdMainLeft
            vDestY = gMsk.CrdMainTop
        
        End If
        
        'Select the card image to use and work out its source position of the card.
        'Note that cards are 1 pixel apart.
        vSourceX = (gMsk.CrdSnglWidth) * pWhichCard
        If gCurrentMode = 5 Or gCurrentMode = 4 _
        Or TheMainForm.chkCardsHidden.Value = vbUnchecked Then
            
            If pNotSelected Then
                
                'Choose from the unselected card images.
                vSourceY = 0
                
            Else
                
                'Choose from the selected card images.
                vSourceY = gMsk.CrdSnglHeight
                
            End If
        
        Else
            
            'Card face is down so select the card back image.
            vSourceY = 0
            vSourceX = 0
            
        End If
        
        'Blit it onto the background map.
        vDummy = BitBlt(Mask4.Map1.hdc, gMsk.CrdDestX(vPosFromLeft) + vDestX, _
                gMsk.CrdDesty(vPosFromLeft) + vDestY, _
                gMsk.CrdSnglWidth - gMsk.CrdSnglSrcBuffer, _
                gMsk.CrdSnglHeight - gMsk.CrdSnglSrcBuffer, _
                Mask4.pctCardSource.hdc, vSourceX, vSourceY, vbSrcCopy)
        
        'Make sure the viewport gets refreshed.
        TheMainForm.gSyncViewportNeeded = True
    End If
End Sub

'Check for little cards being clicked or if clicked hit near main cards. This
'action causes the main cards to be turned over and a valid set preselected
'if one exists.
Public Sub HitLittleCards()
    
    With TheMainForm
    
    'Only if current player is a human player.
    If .GetPlayerController(gPlayerTurn) = 0 Then
        
        'Check for hit on little cards
        If (.gCurrentMousePosY > 1) _
        And (.gCurrentMousePosY < (gMsk.LittleCardTop + gMsk.LittleCardHeight) _
                                                * .gPictureMaskRatioY) Then
            
            'Do as if the main cards were clicked.
            Call CardsClicked
        
        'Check for a hit near main cards.
        ElseIf .gCurrentMousePosY > (gMsk.CrdMainTop - (gMsk.CrdMainHeight * 1)) * .gPictureMaskRatioY _
        And .gCurrentMousePosY < (gMsk.CrdMainTop + (gMsk.CrdMainHeight * 2)) * .GetPictureMaskRatioY _
        And .gCurrentMousePosX > (gMsk.CrdMainLeft - (gMsk.CrdMainWidth * 0.3)) * .gPictureMaskRatioX _
        And .gCurrentMousePosX < (gMsk.CrdMainLeft + (gMsk.CrdMainWidth * 1.1)) * .GetPictureMaskRatioX Then
            
            'Do as if the main cards were clicked.
            Call CardsClicked
            
        End If
    End If
    End With
End Sub

'Draw all players' cards at the top of the map.
Public Sub DrawLittleCards()
    Dim vPlayer As Integer
    Dim vNull As Integer

    'Clear the little cards from across the top of the map.
    Call CleaLittleCards
    
    
    'Print little cards for all players.
    For vPlayer = 1 To 6
        
        'Only draw cards for armies that are in the war.
        If TheMainForm.CountCountriesOwned(vPlayer) > 0 Then
            Call PrintLittleCard(vPlayer, vPlayer - 1)
        End If
    Next
    
End Sub

'Draw all player deatils above the little cards at the top
'of the map if they are in the war.
Public Sub DrawLittleCardText()
    Dim vPlayer As Integer
    Dim vOrigFontBold As Boolean
    Dim vOrigFontSize As Single
    
    With TheMainForm
    
    'Save current font settings.
    vOrigFontBold = .Picture1.Font.Bold
    vOrigFontSize = .Picture1.Font.Size
    .Picture1.Font.Bold = True
    .Picture1.Font.Size = TheMainForm.gLittleCardFontSize
    
    'Print little cards for all players.
    For vPlayer = 1 To 6
        
        'Only print names for armies that are in the war.
        If TheMainForm.CountCountriesOwned(vPlayer) > 0 Then
            Call PrintLittleCardText(vPlayer, vPlayer - 1)
        End If
    Next
    
    'Restore global variables to original values.
    .Picture1.Font.Bold = vOrigFontBold
    .Picture1.Font.Size = vOrigFontSize
    End With
End Sub

'Print player's army name above the little cards on the viewport (Picture1).
'Underline and print using player's color if it is that player's turn. If
'global variable gPrintLittleCardColors is set to TRUE, print all in the
'associated player's colour. Param pLocationIndex indicates which slot position
'to use. This is the same as the pPlayer -1 but was left this way incase in the
'future it is decided to move surviving players to one side or spread them out.
Private Sub PrintLittleCardText(pPlayer As Integer, pLocationIndex As Integer)
    Dim vPrintLocation As Integer
    
    With TheMainForm
    vPrintLocation = (((gMsk.LittleCardWidth * pLocationIndex) / 6) _
                        + gMsk.LittleCardLeft + gMsk.LittleCardPadding) _
                        * .GetPictureMaskRatioX
    
    .Picture1.Font.Underline = (pPlayer = gPlayerTurn)
    If pPlayer = gPlayerTurn Or gPrintLittleCardColors Then
        .Picture1.ForeColor = gPlayerID(pPlayer).bkgndColor
    Else
        .Picture1.ForeColor = &HFFFFFF
    End If
        
    .Picture1.CurrentX = vPrintLocation
    .Picture1.CurrentY = 0
    .Picture1.Print RTrim(gPlayerID(pPlayer).strColor)
    .Picture1.Font.Underline = False
    End With
End Sub

'Put pPlayer's cards at top of screen. Param pLocationIndex indicates which slot position
'to use. This is the same as the pPlayer -1 but was left this way incase in the
'future it is decided to move surviving players to one side or spread them out.
Private Sub PrintLittleCard(pPlayer As Integer, pLocationIndex As Integer)
    Dim vDestX As Integer
    Dim vSourceX As Integer
    Dim vCardIX As Integer
    Dim vDummy As Long
    Dim vPrintLocation As Integer
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
    
        'Work out the position of the group of cards.
        vPrintLocation = (gMsk.LittleCardWidth * pLocationIndex) / 6
        
        'Add the offset from the left of the screen.
        vDestX = vPrintLocation + gMsk.LittleCardLeft
        
        'Print each card until we reach an empty card.
        For vCardIX = 1 To 6
            If gPlayerID(pPlayer).card(vCardIX - 1) <> 0 Then
                
                'Show the back of the card or the front?
                If TheMainForm.chkCardsHidden.Value = vbUnchecked Or gCheatMode.seeCards Then
                    
                    'Show the front. Select which card to print.
                    vSourceX = gPlayerID(pPlayer).card(vCardIX - 1) * gMsk.LittleCrdSngWidth
                
                Else
                    
                    'Cards are hidden, show the back of the card.
                    vSourceX = 0
                
                End If
                
                'Blit the card to the back port.
                vDummy = BitBlt(Mask4.Map1.hdc, vDestX, gMsk.LittleCardTop, gMsk.LittleCrdSngWidth, gMsk.LittleCrdSngHeight, _
                        Mask4.pctLittleCrdSource.hdc, vSourceX, 0, vbSrcCopy)
                
                'Set the X destination to the position of the next card.
                vDestX = vDestX + gMsk.LittleCrdSngWidth + gMsk.LittleCardPadding
                
                'Make sure the viewport gets refreshed.
                TheMainForm.gSyncViewportNeeded = True
            Else
                
                'Empty card, stop printing cards.
                Exit For
                
            End If
        Next
    End If
End Sub

'Clear the main cards area by copying the clean area saved to the mask
'form by function SnapEmptyCardArea().
Public Sub ClearMainCardsArea()
    Dim vDummy As Long
    
    'Only if not running in headless mode.
    If Not gHeadlessMode Then
        vDummy = BitBlt(Mask4.Map1.hdc, gMsk.CrdMainLeft, gMsk.CrdMainTop, _
                gMsk.CrdMainWidth, gMsk.CrdMainHeight, _
                Mask4.pctMainCards.hdc, 0, 0, vbSrcCopy)
        
        'Make sure the viewport gets refreshed.
        TheMainForm.gSyncViewportNeeded = True
    End If
End Sub

'Prints cards at the start of the turn or when vulture cards are acquired.
'Check if the current player MUST change cards.
Public Sub CheckCards()
    Dim vCardIX As Integer
    
    With TheMainForm
    
    'Clear the cards
    Call ClearMainCardsArea
    
    'For each card position.
    For vCardIX = 1 To 5
        
        'In not an empty card.
        If gPlayerID(gPlayerTurn).card(vCardIX - 1) <> 0 Then
            
            'Print the card as not selected.
            Call DrawBigCard(gPlayerID(gPlayerTurn).card(vCardIX - 1), vCardIX, True)
        
        Else
            
            'Empty card, no more cards to print.
            vCardIX = 0
            Exit For
        
        End If
        
    Next
    
    'If there are cards to select.
    If vCardIX <> 0 Then
        
        'Update the info box.
        .InfoBoxPrnCR 0
        .InfoBoxPrint 5                           'bold
        .InfoBoxPrint 10                          'font size * 1.5
        .InfoBoxPrnCR 1, 131                      'must trade cards
        .InfoBoxPrint 11                          'reset font size
        .InfoBoxPrint 6                           'normal
        
        'Show and enable card keys only.
        Call .ToggleKeys(False)
        Call .ToglleCardKeys(True)
        
        'Card selection mode.
        gCurrentMode = 6
        
        'Take different actions if the cards are hidden.
        If .chkCardsHidden.Value = vbUnchecked Then
            Call LookAtCards
        Else
            .cmdExchange.Enabled = IsValidCardSetSelected
        End If
    End If
    
    End With
End Sub

'Collect all cards and reset the deck at the start of a new war.
Public Sub ResetCardsNewWar()
    Dim vPlayer As Integer
    Dim vCardIX As Integer
    
    'For each player.
    For vPlayer = 1 To 6
        
        'For each card.
        For vCardIX = 1 To 10
            
            'Set the card to empty.
            gPlayerID(vPlayer).card(vCardIX - 1) = 0
            
            'Set the card to unpicked.
            gPlayerID(vPlayer).pickedCards(vCardIX - 1) = True
        Next
    Next
    
    'Reset the card deck.
    Call CollectUsedCards
End Sub

'Put used cards back on the deck. This is done by filling the card deck
'to full and then going through all player's hands and reducing the
'appropriate card count from the deck. Ugly but it works beautifully.
Private Sub CollectUsedCards()
    Dim vPlayer As Long
    Dim vCardIX As Long
    
    With TheMainForm
    
    'Replenish all cards.
    gCardDeck(0) = .udCardDeck(0).Value      'Artillery
    gCardDeck(1) = .udCardDeck(1).Value      'Infantry
    gCardDeck(2) = .udCardDeck(2).Value      'Cavalry
    gCardDeck(3) = .udCardDeck(3).Value      'Wild
    
    'For each player.
    For vPlayer = 1 To 6
        
        'For each card the player holds.
        For vCardIX = 0 To 4
            
            'If it is not an empty slot.
            If gPlayerID(vPlayer).card(vCardIX) <> 0 Then
                
                'Reduce the same suit as this card in the card deck by 1.
                gCardDeck(gPlayerID(vPlayer).card(vCardIX) - 1) _
                        = gCardDeck(gPlayerID(vPlayer).card(vCardIX) - 1) - 1
            
            End If
        Next
    Next
    End With
End Sub

'After a battle victory, check if the defender has been wiped out and if so,
'acquire their cards if playing Vulture Cards.
Public Sub CheckVultureCard(pDefender As Integer)
    Dim vCardIX As Long
    Dim vCardIxDef As Long
    
    With TheMainForm
    
    'Check if playing Vulture Cards.
    If .chkCardsVulture.Value = vbUnchecked Then
        .ToggleKeys (gPickedUpUnits = 0)
        Exit Sub
    End If
    
    'Check if the defender has been wiped out.
    If TheMainForm.CountCountriesHeld(pDefender) > 0 Then
        Exit Sub
    End If
    
    'Find the last card of the current player.
    'TODO: Refactor using a for loop.
    vCardIX = 1
    Do While gPlayerID(gPlayerTurn).card(vCardIX - 1) <> 0
        vCardIX = vCardIX + 1
    Loop
    
    'Append all of the defenders cards to the current players cards.
    'TODO: Refactor using a for loop.
    vCardIxDef = 1
    Do While gPlayerID(pDefender).card(vCardIxDef - 1) <> 0
        gPlayerID(gPlayerTurn).card(vCardIX - 1) = gPlayerID(pDefender).card(vCardIxDef - 1)
        Call TheMainForm.AuditTradeCard(gPlayerTurn, -gPlayerID(pDefender).card(vCardIxDef - 1))
        gPlayerID(pDefender).card(vCardIxDef - 1) = 0
        vCardIX = vCardIX + 1
        vCardIxDef = vCardIxDef + 1
    Loop
    
    'Draw cards and check if the there are 5 or more cards in hand
    'and must be exchanged.
    Call CheckCards
    
    'Draw extra cards aquired from killing a player above the other cards.
    If vCardIX > 5 Then
        Call DrawVultureCards
    End If
    
    'Make sure the viewport gets refreshed.
    Call TheMainForm.ColorCountryUnderAttack(0)
    TheMainForm.gSyncViewportNeeded = True
    
    'Different actions for computer players and human players.
    If TheMainForm.GetPlayerController(gPlayerTurn) <> 0 Then
        
        'Computer player.
        .gComputerAquiredCards = True
    Else
        
        'Human player.
        If gPlayerID(gPlayerTurn).card(4) = 0 Then
            Call .ToggleKeys(gPickedUpUnits = 0)
        Else
            Call .ToggleKeys(False)
        End If
    End If
    
    End With
End Sub

'Return true if a valid card set is selected.
Private Function IsValidCardSetSelected() As Boolean
    Dim vCardIX As Integer
    Dim vSelectedIX As Integer
    Dim vChosenCards(2) As Integer
    
    'Find the cards that have been selected.
    vSelectedIX = 0
    For vCardIX = 0 To 9
        If Not gPlayerID(gPlayerTurn).pickedCards(vCardIX) Then
            
            'Card chosen.
            If vSelectedIX <= 2 Then
                
                'Add to the list.
                vChosenCards(vSelectedIX) = gPlayerID(gPlayerTurn).card(vCardIX)
                
                'Increase the selected card count.
                vSelectedIX = vSelectedIX + 1
                
            ElseIf vSelectedIX > 3 Then
                
                'More than 3 cards selected, bail out.
                Exit For
                
            End If
        End If
    Next
    
    'If 3 cards were selected.
    If vSelectedIX = 3 Then
        
        'Check if the selected cards form a valid set.
        IsValidCardSetSelected = GetCardValue(vChosenCards(0), vChosenCards(1), vChosenCards(2)) > 0
        
    Else
        
        'Not 3 cards selected.
        IsValidCardSetSelected = False
        
    End If
    
End Function

'Echange the selected cards if possible. Increase card values if
'increasing cards selected. Update networked players.
Public Sub CardExchangeClicked()
    Dim vCardIX As Integer
    Dim vSelectedIX As Integer
    Dim vNewPoints As Integer
    Dim vLeftoverIX As Integer
    Dim vChosenCard(3) As Integer
    Dim vLeftOverCards(11) As Integer
    
    'Check pause status.
    If gPauseActive Then
        Call TheMainForm.ActivatePauseMode(False)
    End If
    
    'Exit if game was won.
    If gCurrentMode = 13 Or gCurrentMode = 18 Then
        Exit Sub
    End If
    
    'If the current player is not a human on the local maching and
    'the computer player on the local machine didn't call this then
    'bail out. It must have been a remote player.
    If TheMainForm.GetPlayerController(gPlayerTurn) <> 0 _
    And Not TheMainForm.gComputerPressed Then
        Exit Sub
    End If
    
    'No cards selected but must change cards.
    If gCurrentMode = 6 Then
        Call LookAtCards
        Exit Sub
    End If
    
    'Find the cards that have been selected.
    vSelectedIX = 1
    vLeftoverIX = 1
    For vCardIX = 1 To 10
        If gPlayerID(gPlayerTurn).pickedCards(vCardIX - 1) = False Then
            If vSelectedIX > 3 Then
                
                'Too many cards selected, bail out here.
                Exit Sub
            End If
            
            'Add chosen card to the list.
            vChosenCard(vSelectedIX) = gPlayerID(gPlayerTurn).card(vCardIX - 1)
            net.cardsPickedPos(vSelectedIX - 1) = vCardIX
            vSelectedIX = vSelectedIX + 1
        
        Else
            
            'Card not picked so add it to the leftover cards.
            vLeftOverCards(vLeftoverIX) = gPlayerID(gPlayerTurn).card(vCardIX - 1)
            vLeftoverIX = vLeftoverIX + 1
            
        End If
    Next
    
    'Check 3 cards were selected.
    If vSelectedIX <> 4 Then
        Exit Sub
    End If
    
    'Find the point value of the selected cards.
    vNewPoints = GetCardValue(vChosenCard(1), vChosenCard(2), vChosenCard(3))
    If vNewPoints = 0 Then
        Exit Sub
    End If
    
    '"Trade cards in for <vNewPoints> units?"
    If TheMainForm.GetPlayerController(gPlayerTurn) = 0 Then
        If MsgBox(Phrase(132) + str(vNewPoints) + Phrase(133), vbYesNo, "Trading Cards") <> vbYes Then
            Call TheMainForm.Mode5
            Exit Sub
        End If
    End If
    
    'Change card value for when "Increasing Card Value" option is selected.
    'It is increased anyway but that doesn't matter.
    If gCurrentCardValue < 25 Then
        gCurrentCardValue = gCurrentCardValue + 2
    ElseIf gCurrentCardValue = 25 Then
        gCurrentCardValue = 30
    Else
        gCurrentCardValue = gCurrentCardValue + 5
    End If
    If gCurrentCardValue >= gMaxCardValue Then
        gCurrentCardValue = gMaxCardValue
    End If
    
    'Add new points to the player's value.
    TheMainForm.gPlayerValue = TheMainForm.gPlayerValue + vNewPoints + gPickedUpUnits
    Call TheMainForm.AuditAddPointsIssued(gPlayerTurn, vNewPoints)
    Call TheMainForm.AuditTradeCard(gPlayerTurn, vChosenCard(1) + vChosenCard(2) + vChosenCard(3))
    
    'Give player back his leftover cards.
    vLeftoverIX = 0
    For vCardIX = 1 To 10
        gPlayerID(gPlayerTurn).card(vCardIX - 1) = vLeftOverCards(vCardIX)
        If vLeftOverCards(vCardIX) > 0 Then
            vLeftoverIX = vCardIX
        End If
    Next
    
    'Unselect all cards and force the player to continue trading
    'cards if the player still has more than 5 cards.
    Call CardOutOfHand(gPlayerTurn)
    
    Call DrawAllCards
    Call DrawLittleCards
    
    'Send to remote players if networked.
    If netWorkSituation <> cNetNone Then
        Call netMain.ChangeCards(net.cardsPickedPos, CByte(gPlayerTurn))
    End If
    Call TheMainForm.Mode5
End Sub

'Try to find and select a valid set of cards for human players.
Private Sub PreSelectCardsForHuman()
    If FindHumanCards Then
        Call ClickHumanCards(gDiceArray(11))
    End If
End Sub

Public Sub TestCards()
    gPlayerID(gPlayerTurn).card(0) = 2
    gPlayerID(gPlayerTurn).card(1) = 1
    gPlayerID(gPlayerTurn).card(2) = 4
End Sub

'Find cards to pre-select for human players. Return True if found.
'Works by trying to find the most valuable set first, then consider
'jokers, then find 3 of a kind, then consider jokers etc. The order
'that it looks gives priority to the best combinations. Increasing
'and fixed card modes are taken into account.
'TODO: Refactor. This is really ugly! Similar job to AutoCheckCards()
'Use the proper sort algorithm, modUtilities.BubbleSort().
Private Function FindHumanCards() As Boolean
    Dim vNumberbrOfCards As Integer
    Dim vIndex As Integer
    Dim vCountCards As Integer
    Dim vTest1 As Boolean
    
    'Bounce back out if not a human player or not controlled by this terminal.
    If gPlayerID(gPlayerTurn).playerWho <> 0 _
    Or net.playerOwner(gPlayerTurn - 1) <> myTerminalNumber Then
        FindHumanCards = False
        Exit Function
    End If
    
    'Puts any selected cards back on the board.
    Call CardOutOfHand(gPlayerTurn)
    
    'Use dice to do the sorting because dice are un used at this
    'stage of the war and all the needed functions are already written
    'and tested.
    'Clear selected cards left over from outside procedure.
    For vIndex = 0 To 9
        gDiceArray(vIndex) = 0
    Next
    
    'Stack cards to try to find all different.
    For vIndex = 1 To 8
        If gPlayerID(gPlayerTurn).card(vIndex - 1) = 0 Then
            Exit For
        ElseIf gPlayerID(gPlayerTurn).card(vIndex - 1) = 3 Then
            gDiceArray(0) = 3
        ElseIf gPlayerID(gPlayerTurn).card(vIndex - 1) = 2 Then
            gDiceArray(1) = 2
        ElseIf gPlayerID(gPlayerTurn).card(vIndex - 1) = 1 Then
            gDiceArray(2) = 1
        ElseIf gPlayerID(gPlayerTurn).card(vIndex - 1) = 4 Then
            gDiceArray(3) = 4
        End If
    Next
    
    vNumberbrOfCards = 0
    vTest1 = False
    
    'Test cards to find a set of different cards.
    For vIndex = 0 To 2
        If gDiceArray(vIndex) <> 0 Then
            vNumberbrOfCards = vNumberbrOfCards + 1
        End If
    Next
    
    'Do we have 3 different non-joker cards?
    If vNumberbrOfCards = 3 Then
        vTest1 = True
    'Do we have 2 different cards and a joker and cards are fixed?
    ElseIf ((gDiceArray(3) = 4) And (GetCardMode = 1)) Then
        If vNumberbrOfCards = 2 Then
            vTest1 = True
            Call TheMainForm.SortDice(1, 4)
        End If
    End If
    
    'All cards are different one of the above contitions was met.
    If vTest1 Then
        gDiceArray(3) = 0
        gCurrentMode = 8
        gDiceArray(11) = 1
        FindHumanCards = True
        Exit Function
    End If
    
    'Clear selected cards.
    For vIndex = 0 To 9
        gDiceArray(vIndex) = 0
    Next vIndex
    vNumberbrOfCards = 0
    
    'Load all cards into dice global array for sorting.
    For vIndex = 0 To 9
        If gPlayerID(gPlayerTurn).card(vIndex) > 0 Then
            vNumberbrOfCards = vNumberbrOfCards + 1
        End If
        gDiceArray(vIndex) = gPlayerID(gPlayerTurn).card(vIndex)
    Next
    
    'Not enough cards, bail out.
    If (vNumberbrOfCards < 3) Then
        FindHumanCards = False
        Exit Function
    End If
    
    Call TheMainForm.SortDice(1, vNumberbrOfCards)
    
    'Sorting puts the most valuable cards first giving them priority.
    'Are all same, no jokers?
    For vIndex = 1 To vNumberbrOfCards - 2
        vTest1 = ((gDiceArray(vIndex - 1) = gDiceArray(vIndex)) And (gDiceArray(vIndex) = gDiceArray(vIndex + 1)))
        If vTest1 Then
            gDiceArray(11) = vIndex
            FindHumanCards = True
            Exit Function
        End If
    Next vIndex
    
    'Do we have a joker and two other non joker cards?
    If gDiceArray(0) = 4 Then
        For vIndex = 2 To vNumberbrOfCards - 1
            If gDiceArray(vIndex - 1) <> 4 And gDiceArray(vIndex) <> 4 Then
                gDiceArray(1) = gDiceArray(vIndex - 1)
                gDiceArray(2) = gDiceArray(vIndex)
                gDiceArray(11) = 1
                FindHumanCards = True
                Exit Function
            End If
        Next vIndex
    End If
    
    'Check for 2 jokers and another card.
    vCountCards = vNumberbrOfCards
    For vIndex = 1 To vNumberbrOfCards - 1
        If gDiceArray(vIndex - 1) = gDiceArray(vIndex) Then
            gDiceArray(vIndex - 1) = 0
            vCountCards = vCountCards - 1
        End If
    Next
    
    'Check for 2 jokers.
    If vCountCards < 3 Then
        'Clear array.
        For vIndex = 1 To 10
            gDiceArray(vIndex - 1) = 0
        Next vIndex
        vNumberbrOfCards = 0
        'Load up array.
        For vIndex = 0 To 9
            If gPlayerID(gPlayerTurn).card(vIndex) > 0 Then
                vNumberbrOfCards = vNumberbrOfCards + 1
            End If
            gDiceArray(vIndex) = gPlayerID(gPlayerTurn).card(vIndex)
        Next vIndex
        'Sort array.
        Call TheMainForm.SortDice(1, vNumberbrOfCards)
        'Two dice and anything else.
        If gDiceArray(0) = 4 And gDiceArray(1) = 4 Then
            gDiceArray(11) = 1
            FindHumanCards = True
            Exit Function
        End If
        Exit Function
    End If
    
    FindHumanCards = False
End Function

'Turn in 3 cards with value of gDiceArray(atPosition).
'TODO: Refactor and comment. Use the proper sort algorithm modUtilities.BubbleSort().
Private Sub ClickHumanCards(atPosition As Integer)
    Dim vPosIndex As Integer
    Dim vCardIndex As Integer
    Dim vHoldCurMode As Integer
    
    For vPosIndex = atPosition To atPosition + 2
        For vCardIndex = 1 To 10
            If (gDiceArray(vPosIndex - 1) = gPlayerID(gPlayerTurn).card(vCardIndex - 1)) _
            And (gPlayerID(gPlayerTurn).pickedCards(vCardIndex - 1)) Then
                gPlayerID(gPlayerTurn).pickedCards(vCardIndex - 1) = False
                vHoldCurMode = gCurrentMode
                gCurrentMode = 5
                Call DrawBigCard(gDiceArray(vPosIndex - 1), vCardIndex, False)
                gCurrentMode = vHoldCurMode
                Exit For
            End If
        Next
    Next
End Sub

'Return the value of passed cards.
'TODO: Refactor. This is really ugly!
Public Function GetCardValue(pCard1 As Integer, pCard2 As Integer, pCard3 As Integer) As Integer
    Dim tst2 As Boolean
    
    With TheMainForm
    
    'Test for wild.
    If (pCard1 = 4) Or (pCard2 = 4) Or (pCard3 = 4) Then
        If GetCardMode = 2 Then
            GetCardValue = gCurrentCardValue
            Exit Function
        Else
            tst2 = (pCard1 = 4 And pCard2 = 4) Or (pCard2 = 4 And pCard3 = 4) Or (pCard1 = 4 And pCard3 = 4)
            tst2 = tst2 Or ((pCard1 <> pCard2) And (pCard2 <> pCard3) And (pCard3 <> pCard1))
            If tst2 Then            'All different
                GetCardValue = .udFixedValues(3).Value   '10
                Exit Function
            End If
            
            If (pCard1 = 1) Or (pCard2 = 1) Then            'All Artillary
                GetCardValue = .udFixedValues(0).Value   '4
                Exit Function
            End If
            
            If (pCard1 = 2) Or (pCard2 = 2) Then            'All Infantry
                GetCardValue = .udFixedValues(1).Value   '6
                Exit Function
            End If
            
            If (pCard1 = 3) Or (pCard2 = 3) Then            'All Cavalry
                GetCardValue = .udFixedValues(2).Value   '8
                Exit Function
            End If
            gDiceArray(0) = pCard1
            gDiceArray(1) = pCard2
            gDiceArray(2) = pCard3
            Call .SortDice(1, 3)
            
            If (gDiceArray(0) = 4) And (gDiceArray(1) = 4) Then            '2 jokers, maximum value
                GetCardValue = .udFixedValues(3).Value   '10
                Exit Function
            End If
        End If
    End If
    
    If (pCard1 = pCard2) And (pCard2 = pCard3) Then
        If GetCardMode = 2 Then
            GetCardValue = gCurrentCardValue
            Exit Function
        Else
            If pCard1 = 1 Then            'All Artillary
                GetCardValue = .udFixedValues(0).Value   '4
                Exit Function
            End If
            If pCard1 = 2 Then            'All Infantry
                GetCardValue = .udFixedValues(1).Value   '6
                Exit Function
            End If
            If pCard1 = 3 Then            'All Cavalry
                GetCardValue = .udFixedValues(2).Value   '8
                Exit Function
            End If
        End If
    End If

    Call ArangeCards(pCard1, pCard2, pCard3)
    
    If (pCard1 = 3) And (pCard2 = 2) And (pCard3 = 1) Then
        
        'All different
        If GetCardMode = 2 Then
            GetCardValue = gCurrentCardValue
            Exit Function
        Else
            GetCardValue = .udFixedValues(3).Value   '10
            Exit Function
        End If
    End If
    GetCardValue = 0
    End With
End Function

'Simple card bubble sort.
'TODO: Use the proper sort algorithm modUtilities.BubbleSort().
Private Sub ArangeCards(pCard1 As Integer, pCard2 As Integer, pCard3 As Integer)
    Dim vHoldCard As Integer
    
    If pCard1 < pCard2 Then
        vHoldCard = pCard2
        pCard2 = pCard1
        pCard1 = vHoldCard
    End If
    If pCard2 < pCard3 Then
        vHoldCard = pCard3
        pCard3 = pCard2
        pCard2 = vHoldCard
    End If
    If pCard1 < pCard2 Then
        vHoldCard = pCard2
        pCard2 = pCard1
        pCard1 = vHoldCard
    End If
End Sub

'Pack all remaining cards into the passed card deck.
'The passed deck is a flat integer array.
Private Sub PackCardDeck(pCardDeck() As Integer)
    Dim vIndex As Long
    Dim vCardType As Long
    Dim vCntr As Long
    
    vIndex = 0
    ReDim pCardDeck(HowManyStartingCards) As Integer
    
    For vCardType = 1 To 4
        For vCntr = 0 To gCardDeck(vCardType - 1) - 1
            pCardDeck(vIndex) = vCardType
            vIndex = vIndex + 1
        Next
    Next
    ReDim Preserve pCardDeck(vIndex) As Integer
End Sub

'Count the number of cards starting in the deck by adding up
'the selected counts in the startup screen.
Private Function HowManyStartingCards() As Long
    Dim vIndex As Long
    
    For vIndex = 0 To 3
        HowManyStartingCards = HowManyStartingCards + TheMainForm.udCardDeck(vIndex).Value
    Next
End Function

'Collect used cards and repack. Notify humans if required.
Private Sub CollectAndShuffleUsedCards()
    Dim vPlayer As Integer
    Dim vMsgBoxRslt As Long
    
    net.reshuffleCards = True
    
    'Collect used cards and put them in the card deck.
    Call CollectUsedCards
    
    If Not TheMainForm.SetupScreen.Visible _
    And TheMainForm.mnuOptReport.Checked _
    And Not gHeadlessMode Then
    
        'Look for human players because there is no point showing the
        'reshuffle notice if they are all computer controlled.
        For vPlayer = 1 To 6
            If TheMainForm.GetPlayerController(vPlayer) = 0 Then
                
                'Human player found.
                If TheMainForm.GetInitialPoints(vPlayer) > 0 Then
                    
                    'Show the reshuffle notice.
                    TheMainForm.InfoBoxPrint 0       'Cls
                    frmIntelligence.lblShowAgain(0).Visible = True
                    frmIntelligence.lblShowAgain(0).Caption = Phrase(136)
                    frmIntelligence.lblShowAgain(1).Visible = True
                    frmIntelligence.lblShowAgain(1).Caption = Phrase(136)
                    frmIntelligence.chkShowAgain.Visible = True
                    
                    'Show the reshuffle notice.
                    If netWorkSituation = cNetNone Then
                        vMsgBoxRslt = TheMainForm.MRbox(Phrase(134), Phrase(135))     '"Reshuffling cards"
                    Else
                        TheMainForm.MRbox Phrase(134), Phrase(135), True        '"Reshuffling cards"
                    End If
                    
                    Exit For
                    
                End If
            End If
        Next
    End If
End Sub

'Deal a card to the current player at the end of their turn.
'Reshuffle the deck if required.
Public Sub DealACard()
    Dim vCardDeck() As Integer
    Dim vShuffle As Long
    Dim vDealtCard As Integer
    
    'Only if cards are in the game.
    If GetCardMode <> 0 Then

        'Pack all remaining cards into a deck.
        Call PackCardDeck(vCardDeck)
        
        'Check if there are any cards left.
        If UBound(vCardDeck) = 0 Then
            
            'Deck empty. Collect used cards and
            'repack. Notify humans as required.
            Call CollectAndShuffleUsedCards
            Call PackCardDeck(vCardDeck)
            net.reshuffleCards = True
        
        Else
            
            'No reshuffle needed.
            net.reshuffleCards = False
            
        End If
            
        'Pick a random card from the deck.
        vDealtCard = vCardDeck(Int(GenRandom4 * UBound(vCardDeck)))
        
        'Put the card in the player's hand.
        Call PutCard(vDealtCard)
        
        'Reduce the number of cards left in the deck.
        gCardDeck(vDealtCard - 1) = gCardDeck(vDealtCard - 1) - 1
        
        'Apdate the audit log.
        Call TheMainForm.AuditAddCardsIssued(gPlayerTurn, vDealtCard)
    End If
End Sub

'Put cards back and show reshuffle notice from remote player only
'if there is a human player on this terminal,
Public Sub PutCardsBack()
    Dim vPlayer As Integer
    
    'For each player.
    For vPlayer = 1 To 6
        
        'Bail if setup is visible or we have opted to not see the notice.
        If TheMainForm.SetupScreen.Visible _
        Or Not TheMainForm.mnuOptReport.Checked Then
            Exit For
        End If
        
        'If this player is a human player.
        If TheMainForm.GetPlayerController(vPlayer) = 0 Then
            
            'If the player started in the game even if they are now wiped out.
            If TheMainForm.GetInitialPoints(vPlayer) > 0 Then
                
                'Print the reshuffle notice on the info box and a popup message box.
                TheMainForm.InfoBoxPrint 0       'Cls
                frmIntelligence.lblShowAgain(0).Caption = Phrase(136)   'Don't show any more.
                frmIntelligence.lblShowAgain(0).Visible = True
                frmIntelligence.lblShowAgain(1).Caption = Phrase(136)   'Don't show any more.
                frmIntelligence.lblShowAgain(1).Visible = True
                frmIntelligence.chkShowAgain.Visible = True
                If netWorkSituation = cNetNone Then
                    vPlayer = TheMainForm.MRbox(Phrase(134), Phrase(135))   'Re-shuffling cards. Used cards are being collected, re-shuffled and put back in the pack.
                Else
                    TheMainForm.MRbox Phrase(134), Phrase(135), True    'Re-shuffling cards. Used cards are being collected, re-shuffled and put back in the pack.
                End If
                Exit For
            End If
        End If
    Next
    
    'Put used cards back in the card deck.
    Call CollectUsedCards
End Sub

'Put the passed card in the players hand next to last card.
Public Sub PutCard(pCard As Integer)
    Dim vCardIndex As Integer
    
    'For all cards slots in the player's hand.
    For vCardIndex = 0 To 4
        
        'If the slot is empty.
        If gPlayerID(gPlayerTurn).card(vCardIndex) = 0 Then
            
            'Put the card in the empty slot and exit.
            gPlayerID(gPlayerTurn).card(vCardIndex) = pCard
            Exit For
        
        End If
    Next
End Sub
