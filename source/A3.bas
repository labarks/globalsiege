Attribute VB_Name = "A3"
Option Explicit
' The 3rd AI for Global Siege
' 10.7.99

Private Const endOfPath As Long = -1    ' Signal end of path

    ' Various commands embedded in path
Private Enum pathAction
    endTurn                             ' End of list
    reEvaluateAttack                    ' End of attack sequence, force re-evaluation
    placeUnits                          ' Put new units into country(too)
    AttackCountry                       ' Attack country(too) from country(from)
    moveToCountry                       ' Move to country(too) from country(from)
    pickUp                              ' Picking up units
    retreat                             ' Retreat into country(too) from country(from)
End Enum

    ' Final list of actions and countries for A3
Private Type warPathType
    too(100) As Byte                    ' Attack, move this ctry or place units here
    from(100) As Byte                   ' Source country
    transferNumber(100) As Long         ' Number to attack with, move or place
    expectedUnitsLeft(100) As Long      ' Should have this many units left after attack
    action(100) As pathAction           ' What to do
    myDamage As Long                    ' Number of units I have lost following path
    hisDamage As Long                   ' Number of units enemy has lost
    Pointer As Byte                     ' Where upto in list
End Type

Private Type A3PathType
    from As Long                        ' Attack from this country
    To As Long                          ' Defender
    UnitsNeeded As Long                 ' Number of units left
    'outOfUnits as Boolean               ' True if out of units
End Type
 
Private Type A3Paths
    pathToConts(6, 50) As A3PathType          ' Path to continents with pax points
    pathToContsCost(6) As Long          ' Units needed to get to continents with everything
    pathToContsDist(6, 50) As A3PathType      ' Path to continents shortest distance/cost
    pathToContsDistance(6) As Long      ' Units needed to get to continents shortest distance/cost
    pathKillPlayer(6, 50) As A3PathType       ' Path to kill player
    pathKillPlayerCost(6) As Long       ' Units needed to kill player
    pathToWin(50) As A3PathType               ' Path to win
    pathToWinCost As Long               ' Units required to win
End Type

Public gA3MaxSearchDepth As Long                 ' Maximum depth for A3 recursive search.
Public gA2MaxSearchDepth As Long            ' Maximum depth for A2rcSearch

Private OtherPlayerValues(6) As Long        ' New unit value of all players
Public otherPlayerCards(6, 10) As Byte      ' Cards other players hold

Public startOfNewTurn As Boolean              ' True if not started and initialized turn yet

Private warPath As warPathType              ' Final list of actions and countries for A3

Private A3RandNeighbor(42, 7) As Long        ' List of neibours (1 to 7) in random order
Private A3RandCtry(100, 42) As Long          ' List of countries in random search order

Public A3CntryNeigbors(42, 7) As Long          ' Neigbors of each country
Public A3ContOfCntry(42) As Long            ' Continent each country belongs to
Public A3ContValue(6) As Long               ' Value of each continent

Public A3CntryOwner(42) As Long                ' Current country owner
Public A3CntryScore(42) As Long                ' Current country score
Public A3CntryAllys(42, 7) As Long              ' Friendly neigbors with points

Public A3Itterations As Long               'For recursive testing

Public A3RecordTo(50) As Long                'Record of current path during search
Public A3RecordFrom(50) As Long
Public A3RecordUnits(50) As Long

Public A3PathTo(50) As Long                   ' Path passed to external modules
Public A3PathFrom(50) As Long
Public A3PathUnits(50) As Long
Public A3PathCost As Long

Public A3UnitsLeft As Long

Private A3PlayerTurn As Long                ' Current player turn
Private Path As A3Paths                     ' Paths found by search

Public A3NoCntrysOwn(6) As Long              ' Countries players own
Private A3IsAgate(42) As Boolean            ' True if a gate


    ' Initialize A3 at program start
Public Sub A3Initialize()
    'Set the maximum recursive search depth for the A3 soldiers.
    gA3MaxSearchDepth = GetSetting(gcApplicationName, "settings", "A3MaxSearchDepth", 8)
    gA2MaxSearchDepth = gA3MaxSearchDepth
    startOfNewTurn = True
    Call randomizeNeigbor
    Call randomizeCountries
    Call loadGates
End Sub

    ' Find all gates
Private Sub loadGates()
    Dim i As Long
    For i = 1 To 42
        A3IsAgate(i) = TheMainForm.autoIsAGate(CInt(i))
    Next
End Sub

    ' Return True if cCtry is a continent's gate
Public Function IsAgate(ctry As Long)
    IsAgate = A3IsAgate(ctry)
End Function

    ' Return a random list of countries
Public Sub fillRandList(ByRef randList() As Integer)
    Dim i As Long
    Dim chosen As Long
    
    chosen = CLng(GenRandom4 * 99)
    
    'Just in case.
    If A3RandCtry(chosen, 1) = 0 Then
        Call randomizeCountries
    End If
    
    For i = 1 To 42
        randList(i) = A3RandCtry(chosen, i)
    Next
End Sub

    ' Fill list with random neigbour order
Public Sub fillRandNeigborList(cntry As Integer, ByRef randNbrList() As Integer)
    Dim i As Long
    
    If GenRandom4 < 0.3 Then           '30% chance or re order
        Call randomizeNeigbor
    End If
    
    For i = 1 To 6
        randNbrList(i) = CInt(A3RandNeighbor(cntry, i))
    Next
End Sub


    ' Set random list of neigbours
Public Sub randomizeNeigbor()
    Dim cntr1 As Long, cntr2 As Long, ptr As Long
    Dim neiborChoises(7) As Long
    Dim chosen As Long
    Dim neigborsLeft As Long
    
    For cntr1 = 1 To 42
        neigborsLeft = countNeigbours(cntr1)
        ptr = 1
        For cntr2 = 1 To 7
            neiborChoises(cntr2) = A3CntryNeigbors(cntr1, cntr2)
        Next
        Do While neigborsLeft > 0
            chosen = CLng(GenRandom4 * 5) + 1
            If neiborChoises(chosen) > 0 Then
                A3RandNeighbor(cntr1, ptr) = neiborChoises(chosen)
                neiborChoises(chosen) = 0
                ptr = ptr + 1
                neigborsLeft = neigborsLeft - 1
            End If
        Loop
    Next
End Sub

    ' Return the number of neigbours country has
Private Function countNeigbours(country As Long) As Long
    Dim cntr As Long
    
    countNeigbours = 0
    For cntr = 1 To 7
        If A3CntryNeigbors(country, cntr) = 0 Then
            Exit For
        Else
            countNeigbours = countNeigbours + 1
        End If
    Next
End Function

    ' Count total for testing
Private Sub tg()
    Dim cntr As Long, rslt As Single, total As Long
    
    For cntr = 1 To 42
        If countNeigbours(cntr) > total Then
            total = countNeigbours(cntr)
        End If
    Next
    Debug.Print "Average = "; total
End Sub

    ' Set random search list of countries
Private Sub randomizeCountries()
    Dim cntr1 As Long, cntr2 As Long, ptr As Long
    Dim countriesChosen(42) As Long
    Dim chosen As Long
    Dim countriesLeft As Long
    Const lastFew As Long = 5
    
    For cntr1 = 0 To 99
            
            'Reset counter and de-select countries
        countriesLeft = 42
        For cntr2 = 1 To 42
            countriesChosen(cntr2) = cntr2
        Next
        
            'Randomly pick all but the last few
        Do While countriesLeft > lastFew
            chosen = CLng(GenRandom4 * 41) + 1
            If countriesChosen(chosen) > 0 Then
                A3RandCtry(cntr1, countriesLeft) = countriesChosen(chosen)
                countriesChosen(chosen) = 0
                countriesLeft = countriesLeft - 1
            End If
        Loop
        
            'Linearly pick last few for speed.
            'Randomly select search direction
        If GenRandom4 > 0.5 Then
            For chosen = 1 To 42
                If countriesChosen(chosen) > 0 Then
                    A3RandCtry(cntr1, countriesLeft) = countriesChosen(chosen)
                    countriesChosen(chosen) = 0
                    countriesLeft = countriesLeft - 1
                    If countriesLeft <= 0 Then
                        Exit For
                    End If
                End If
            Next
        Else
            For chosen = 42 To 1 Step -1
                If countriesChosen(chosen) > 0 Then
                    A3RandCtry(cntr1, countriesLeft) = countriesChosen(chosen)
                    countriesChosen(chosen) = 0
                    countriesLeft = countriesLeft - 1
                    If countriesLeft <= 0 Then Exit For
                End If
            Next
        End If
    Next
End Sub

'TheMainForm.ContHeldByEnemy
    ' Try to put a thorn in someones side.
Public Function A2FindThornInSide(Player As Integer, Defence As Single) As Boolean
    Dim MyPoints As Integer
    Dim HisPoints As Integer
    Dim modifier As Integer
    Dim Extra As Integer
    
    modifier = 20
    A2FindThornInSide = False
    A2opportunity.pathPointer = 0
    Call TheMainForm.A3UpdateMap(CLng(Player))
    
    With TheMainForm
    If .OwnContinent(5, Player) Then
        'If I own Asia, attack Ukraine?
        MyPoints = A3CntryScore(32) + A3CntryScore(28) - 2
        HisPoints = A3CntryScore(20) - 1
        If A3CntryOwner(20) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then  '10% safety
            ' Attack Ukraine.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(20, 32, 28)
            A2FindThornInSide = True
        End If
        'If I own Asia, attack Australia, Alaska?
        MyPoints = A3CntryScore(30) - 1
        HisPoints = A3CntryScore(39) - 1
        If A3CntryOwner(39) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Australia.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(39, 30, 0)
            A2FindThornInSide = True
        End If
        'Alaska?
        MyPoints = A3CntryScore(38) - 1
        HisPoints = A3CntryScore(1) - 1
        If A3CntryOwner(1) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Alaska.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(1, 38, 0)
            A2FindThornInSide = True
        End If
    End If
    
    If .OwnContinent(3, Player) _
    And .OwnContinent(4, Player) Then
        'If I own Europe and Africa attack Mid East?
        MyPoints = A3CntryScore(22) + A3CntryScore(23) - 2
        HisPoints = A3CntryScore(27) - 1
        If A3CntryOwner(27) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then  '10% safety
            ' Attack Mid East.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(27, 22, 23)
            A2FindThornInSide = True
        End If
        ' Attack Brazil?
        MyPoints = A3CntryScore(21) - 1
        HisPoints = A3CntryScore(12) - 1
        If A3CntryOwner(12) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Brazil.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(12, 21, 0)
            A2FindThornInSide = True
        End If
    End If
    
    If .OwnContinent(2, Player) _
    And .OwnContinent(3, Player) Then
        'If I own South America and Europe attack North Africa?
        MyPoints = A3CntryScore(12) + A3CntryScore(18) - 2
        HisPoints = A3CntryScore(21) - 1
        If A3CntryOwner(21) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then  '10% safety
            ' Attack North Africa.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(21, 12, 28)
            A2FindThornInSide = True
        End If
    End If
    
    If .OwnContinent(3, Player) Then
        ' If I own Europe, attack Greenland, North Africa?
        MyPoints = A3CntryScore(14) - 1
        HisPoints = A3CntryScore(3) - 1
        If A3CntryOwner(3) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Greenland.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(3, 14, 0)
            A2FindThornInSide = True
        End If
        ' North Africa?
        MyPoints = A3CntryScore(18) - 1
        HisPoints = A3CntryScore(21) - 1
        If A3CntryOwner(21) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Greenland.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(21, 18, 0)
            A2FindThornInSide = True
        End If
    End If
    
    If .OwnContinent(4, Player) Then
        'I own Africa, attack Mid East?
        MyPoints = A3CntryScore(23) - 1
        HisPoints = A3CntryScore(27) - 1
        If A3CntryOwner(27) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) _
        And ((A3CntryScore(22) >= A3CntryScore(19) And A3CntryOwner(19) <> Player) _
        Or A3CntryOwner(19) = Player) Then '10% safety, don't open path into Africa.
            ' Attack Mid East.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(27, 23, 0)
            A2FindThornInSide = True
        End If
    End If
    
    If .OwnContinent(5, Player) _
    And .OwnContinent(4, Player) Then
        'If I own Asia and Africa, attack Southern Europe?
        MyPoints = A3CntryScore(22) - 1
        HisPoints = A3CntryScore(19) - 1
        If A3CntryOwner(19) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Southern Europe.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(19, 22, 0)
            A2FindThornInSide = True
        End If
    End If
    
    If .OwnContinent(2, Player) Then
        'If I own South America, attack North Africa?
        MyPoints = A3CntryScore(12) - 1
        HisPoints = A3CntryScore(21) - 1
        If A3CntryOwner(21) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack North Africa.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(21, 12, 0)
            A2FindThornInSide = True
        End If
        'If I own South America, attack Central America?
        MyPoints = A3CntryScore(10) - 1
        HisPoints = A3CntryScore(9) - 1
        If A3CntryOwner(9) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Central America.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(9, 10, 0)
            A2FindThornInSide = True
        End If
    End If
    
    If .OwnContinent(1, Player) Then
        'If I own North America, attack Kamchata, Peru, Iceland?
        MyPoints = A3CntryScore(1) - 1
        HisPoints = A3CntryScore(38) - 1
        If A3CntryOwner(38) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Kamchata.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(38, 1, 0)
            A2FindThornInSide = True
        End If
        ' Peru?
        MyPoints = A3CntryScore(9) - 1
        HisPoints = A3CntryScore(10) - 1
        If A3CntryOwner(10) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Peru.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(10, 9, 0)
            A2FindThornInSide = True
        End If
        ' Iceland?
        MyPoints = A3CntryScore(3) - 1
        HisPoints = A3CntryScore(14) - 1
        If A3CntryOwner(14) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Iceland.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(14, 3, 0)
            A2FindThornInSide = True
        End If
    End If
    
    If .OwnContinent(6, Player) Then
        'If I own Australia, attack Siam?
        MyPoints = A3CntryScore(39) - 1
        HisPoints = A3CntryScore(30) - 1
        If A3CntryOwner(30) <> Player _
        And MyPoints > HisPoints + 3 + (MyPoints / modifier) Then '10% safety
            ' Attack Siam.
            A2opportunity.unitsRequired = HisPoints + 1
            Call LayPath(30, 39, 0)
            A2FindThornInSide = True
        End If
    End If
    End With
    
    With A2opportunity
    .Path(.pathPointer) = CInt(endOfPath)
    .pathPointer = 0
    .StopWhenFinished = A2FindThornInSide
    End With
End Function

Private Sub LayPath(Victim As Integer, Attacker1 As Integer, attacker2 As Integer)
    With A2opportunity
    .IsActive = True
    If Not VictimAlreadyInPath(Victim) Then
        .Path(.pathPointer) = Victim
        .Path(.pathPointer + 1) = -Attacker1 - 10
        If attacker2 > 0 Then
            .Path(.pathPointer + 2) = -attacker2 - 10
            .pathPointer = .pathPointer + 3
        Else
            .pathPointer = .pathPointer + 2
        End If
        .Path(.pathPointer) = 0
    End If
    End With
End Sub

Private Function VictimAlreadyInPath(Victim As Integer) As Boolean
    Dim cntr As Long
    With A2opportunity
    VictimAlreadyInPath = False
    For cntr = 0 To .pathPointer
        If .Path(cntr) = Victim Then
            VictimAlreadyInPath = True
            Exit For
        End If
    Next
    End With
End Function

    'A2 helper
    'Try to kill player for cards or try to win
Public Function findOpportunity(newPoints As Integer, Optional pMaxRcDepth As Long = -1) As Boolean
    Dim cntr As Long
    Dim ptr As Long
    Dim tmp As Long
    Dim cntry As Long
    Dim vA2MaxSearchDepth As Long
    
    findOpportunity = False
    If pMaxRcDepth > 0 Then
        gA2MaxSearchDepth = pMaxRcDepth
    Else
        gA2MaxSearchDepth = gA3MaxSearchDepth
    End If
    
    A3PlayerTurn = CLng(gPlayerTurn)
    Call TheMainForm.A3UpdateMap(A3PlayerTurn)
    
    ptr = CLng(GenRandom4 * 100)
    'If ptr < 3 Then randomizeCountries      '2% chance of re-organizing country list
    
    A3Itterations = 0
    
    A3UnitsLeft = 0
    Path.pathToWinCost = 0
    For cntr = 1 To 6
        Path.pathKillPlayerCost(cntr) = -99999
    Next
    
    'Start search from last country
    If gTargetCtry > 0 Then
        cntry = CLng(gTargetCtry)
        If A3CntryOwner(cntry) = A3PlayerTurn Then
            tmp = A3CntryScore(cntry)
            If tmp - 1 + newPoints > 2 Then
                A3CntryScore(cntry) = 1
                A2rcSearch cntry, 0, tmp - 1 + CLng(newPoints), 0
                A3CntryScore(cntry) = tmp
            End If
        End If
    End If
    'Search rest of countries
    For cntr = 1 To 42
        cntry = A3RandCtry(ptr, cntr)
        If A3CntryOwner(cntry) = A3PlayerTurn And cntry <> CLng(gTargetCtry) Then
            tmp = A3CntryScore(cntry)
            If tmp - 1 + newPoints > 2 Then
                A3CntryScore(cntry) = 1
                A2rcSearch cntry, 0, tmp - 1 + CLng(newPoints), 0
                A3CntryScore(cntry) = tmp
            End If
        End If
    Next
    
        'Found a path to win?
    If Path.pathToWinCost > 0 Then
        A2opportunity.IsActive = True
        A2opportunity.pathPointer = 0
        A2opportunity.unitsRequired = CInt(Path.pathToWinCost)
        For cntr = 0 To 42
            A2opportunity.Path(cntr) = CInt(Path.pathToWin(cntr).To)
            If Path.pathToWin(cntr).To = 0 Then Exit For
        Next
        findOpportunity = True
        
    Else        'Try to find something to kill
        findOpportunity = A2tryToKillSomething
    End If
End Function

    'Try to find something to kill from list
    'Then fill A2opportunity. If dangerous to kill then don't (missions)
Private Function A2tryToKillSomething() As Boolean
    Dim i As Integer, ptr As Integer
    Dim KillPlayerList(6) As Integer
    Dim KillPlayerReward(6) As Integer
    Dim bestReward As Integer
    Dim bestEnemy As Integer
    Dim tmp1 As Integer
    
    ptr = 1
    For i = 1 To 6
        If Path.pathKillPlayerCost(i) > 0 Then      'Found one to kill
            If TheMainForm.okToKill Then              'No danger in killing
                KillPlayerList(ptr) = i             'Point to victim
                If TheMainForm.isTarget(i) Then       'Is part of mission?
                    KillPlayerReward(ptr) = 9999
                Else                                'Reward = CardPoints + points left
                    tmp1 = A2exchangeValue(ptr)
                    If tmp1 = 0 Then
                        KillPlayerReward(ptr) = 0
                    Else
                        KillPlayerReward(ptr) = Path.pathKillPlayerCost(i) + tmp1
                    End If
                End If
                ptr = ptr + 1
            End If
        End If
    Next
    
    If ptr > 1 Then
        bestReward = -9999
        For i = 1 To 6                              'Select best victim
            If bestReward < KillPlayerReward(i) Then
                bestReward = KillPlayerReward(i)
                bestEnemy = i
            End If
            If i = ptr - 1 Then Exit For
        Next
        
        A2opportunity.IsActive = True                 'Load path
        A2opportunity.pathPointer = 0
        A2opportunity.unitsRequired = CInt(Path.pathKillPlayerCost(KillPlayerList(bestEnemy)))
        For i = 0 To 42
            A2opportunity.Path(i) = CInt(Path.pathKillPlayer(KillPlayerList(bestEnemy), i).To)
            If Path.pathKillPlayer(KillPlayerList(bestEnemy), i).To = 0 Then Exit For
        Next
        A2tryToKillSomething = True
    Else
        A2tryToKillSomething = False
    End If
End Function

    'Points I would get for my cards and other player's cards
    'Only by counting them, not by their value
Private Function A2exchangeValue(ownerOfCards As Integer) As Integer
    Dim MyCardList(10) As Integer       'My Cards
    Dim HisCardList(10) As Integer
    Dim i As Long, NoMyCards As Long
    Dim NoHisCards As Long
    Dim NoTotalCards As Long
    Dim Joker As Boolean
    
    Call TheMainForm.loadCardList(CInt(A3PlayerTurn), MyCardList)
    Call TheMainForm.loadCardList(ownerOfCards, HisCardList)
    
    NoMyCards = 0
    NoHisCards = 0
    Joker = False
    For i = 1 To 10         'Count cards
        If MyCardList(i) > 0 Then
            NoMyCards = NoMyCards + 1
        End If
        If HisCardList(i) > 0 Then
            NoHisCards = NoHisCards + 1
        End If
                            'Any Jokers in his list?
        Joker = Joker Or HisCardList(i) = 4
    Next
    
    NoTotalCards = NoMyCards + NoHisCards
        'Some bodgy calculations
    If NoHisCards = 0 Then
        A2exchangeValue = 0
    ElseIf NoTotalCards > 2 And Joker Then
        A2exchangeValue = GetCardValue(1, 2, 3)
    ElseIf NoTotalCards < 4 And NoHisCards > 1 Then
        A2exchangeValue = GetCardValue(1, 1, 1)
    ElseIf NoTotalCards = 4 Then
        A2exchangeValue = GetCardValue(2, 2, 2)
    ElseIf NoTotalCards = 5 Then
        A2exchangeValue = GetCardValue(3, 3, 3)
    ElseIf NoTotalCards > 5 Then
        A2exchangeValue = GetCardValue(1, 2, 3)
    End If
    
        ' Consider total number of cards acquired
    A2exchangeValue = A2exchangeValue + NoHisCards
End Function

    ' Recursive serch for A2 - find quick win or kill army (only)
Private Sub A2rcSearch(ByVal ctry As Long, ByVal Depth As Long, _
ByVal UnitsLeft As Long, ByVal pathPointer As Long)
    Dim cntr As Long, cntr2 As Long, tmp As Long
    Dim previousOwner As Long, ctryNeigbor As Long
    Dim ctryValue As Long, tmpCtryNbr As Long
    Dim allyNeigborUnits(7) As Long
    Dim A2incentive As Long                 'Incentive to finish path on a gate or enemy
    
    A3RecordFrom(pathPointer) = ctry
        ' Search neigboring countries
    For cntr = 1 To 7
        ctryNeigbor = A3RandNeighbor(ctry, cntr)
        ctryValue = A3CntryScore(ctryNeigbor)
        
            ' No more iterations (neigbours) left on this path
        If ctryNeigbor = 0 Then
            Exit For
            
            ' Test target and take actions if pass
        ElseIf A3CntryOwner(ctryNeigbor) <> A3PlayerTurn Then
                'Set situation parameters
            A3Itterations = A3Itterations + 1
            previousOwner = A3CntryOwner(ctryNeigbor)
            A3NoCntrysOwn(previousOwner) = A3NoCntrysOwn(previousOwner) - 1
            A3CntryOwner(ctryNeigbor) = A3PlayerTurn
            A3NoCntrysOwn(A3PlayerTurn) = A3NoCntrysOwn(A3PlayerTurn) + 1
            A3RecordTo(pathPointer) = ctryNeigbor
            UnitsLeft = UnitsLeft - ctryValue
            A3CntryScore(ctryNeigbor) = 1
            
                ' Check for ally units (unrolled loop for speed)
            allyNeigborUnits(1) = 0
            allyNeigborUnits(2) = 0
            allyNeigborUnits(3) = 0
            allyNeigborUnits(4) = 0
            allyNeigborUnits(5) = 0
            allyNeigborUnits(6) = 0
                ' Look at all ally neibours for points
            If UnitsLeft > 1 And A3CntryAllys(ctryNeigbor, 0) > 0 Then
                For cntr2 = 1 To A3CntryAllys(ctryNeigbor, 0)
                    tmp = A3CntryAllys(ctryNeigbor, cntr2)
                    If A3CntryScore(tmp) > 1 Then
                        UnitsLeft = UnitsLeft + A3CntryScore(tmp) - 1
                        allyNeigborUnits(cntr2) = A3CntryScore(tmp)
                        A3CntryScore(tmp) = 1
                    End If
                Next
            End If
            
            
            A3RecordTo(pathPointer + 1) = endOfPath
                    
                ' Have I won?
            If A3NoCntrysOwn(A3PlayerTurn) = 42 And UnitsLeft > Path.pathToWinCost Then
                Path.pathToWinCost = UnitsLeft
                For cntr2 = 0 To 49
                    Path.pathToWin(cntr2).To = A3RecordTo(cntr2)
                    Path.pathToWin(cntr2).from = A3RecordFrom(cntr2)
                    If A3RecordTo(cntr2) = 0 Then Exit For
                Next
            End If
            
                ' Have I killed some one?
            If A3NoCntrysOwn(previousOwner) = 0 Then
                If UnitsLeft + 2 >= Path.pathKillPlayerCost(previousOwner) Then
                        'Do as few times as possible
                    A2incentive = A2getincentive(previousOwner, pathPointer)
                    If UnitsLeft + A2incentive > Path.pathKillPlayerCost(previousOwner) Then
                        Path.pathKillPlayerCost(previousOwner) = UnitsLeft + A2incentive
                        For cntr2 = 0 To 49
                            Path.pathKillPlayer(previousOwner, cntr2).To = A3RecordTo(cntr2)
                            Path.pathKillPlayer(previousOwner, cntr2).from = A3RecordFrom(cntr2)
                            If A3RecordTo(cntr2) = 0 Then Exit For
                        Next
                    End If
                End If
            End If
            
                ' Recurse if I can
            If Depth < gA2MaxSearchDepth Then
                A2rcSearch ctryNeigbor, Depth + 1, UnitsLeft - 1, pathPointer + 1
            End If
            
                'Undo situation parameters
            If A3CntryAllys(ctryNeigbor, 0) > 0 Then
                For cntr2 = 1 To A3CntryAllys(ctryNeigbor, 0)
                    tmp = A3CntryAllys(ctryNeigbor, cntr2)
                    If allyNeigborUnits(cntr2) > 0 Then
                        A3CntryScore(tmp) = allyNeigborUnits(cntr2)
                        UnitsLeft = UnitsLeft - A3CntryScore(tmp) + 1
                    End If
                Next
            End If
            
            A3CntryScore(ctryNeigbor) = ctryValue
            UnitsLeft = UnitsLeft + ctryValue
            A3NoCntrysOwn(A3PlayerTurn) = A3NoCntrysOwn(A3PlayerTurn) - 1
            A3CntryOwner(ctryNeigbor) = previousOwner
            A3NoCntrysOwn(previousOwner) = A3NoCntrysOwn(previousOwner) + 1
        End If
    Next
    A3RecordTo(pathPointer) = endOfPath     'Put "End of Path" at top of list
End Sub

    'Add incentive for path to end on a gate or next to an enemy
Private Function A2getincentive(previousOwner As Long, pathPointer As Long) As Long
    Dim i As Long
    Dim ctry As Long
    Dim nbr As Long
    
    ctry = A3RecordTo(pathPointer)
    
        'Is last country a gate? Probably a more efficient way than this
    If A3IsAgate(ctry) Then
        A2getincentive = 1
    ElseIf ctry = 41 Or ctry = 40 Then      'Keep away from East Aust
        A2getincentive = 1
    End If
    
        'Is last country next to an enemy?
    For i = 1 To 7
        nbr = A3RandNeighbor(ctry, i)
        If nbr = 0 Then
            Exit Function
        ElseIf A3CntryOwner(nbr) <> A3PlayerTurn Then
            A2getincentive = A2getincentive + 1
            Exit Function
        End If
    Next
End Function
