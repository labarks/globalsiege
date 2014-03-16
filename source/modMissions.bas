Attribute VB_Name = "modMissions"
Option Explicit

'Each mission is marked as "not active" if it is unavailable. Missions are picked
'at random and if it is not active or not feasible (for example, players shouldn't
'be dealt missions to kill them selves), another mission is picked at random until
'an active and feasible mission is picked. The mission is then marked as "not active"
'and the process is done again for the next player.
'
'The reasons for missions being marked as “not active” are:
'   - Mission type is un-checked in the Missions tool menu
'   - The target army is not in the war
'   - The mission has already been dealt to another player
'
'The two types of missions are army wipeout missions and conquer and hold missions.
'
'Mission Numbers:
'0. You must wipe out all other players and conquer the world.
'1. Your mission is to wipe out the Red Army.
'2. Your mission is to wipe out the Green Army.
'3. Your mission is to wipe out the Blue Army.
'4. Your mission is to wipe out the Yellow Army.
'5. Your mission is to wipe out the Purple Army.
'6. Your mission is to wipe out the gray Army.
'7. You must conquer North America and South America and hold them until your next turn.
'8. You must conquer North America And Australia and hold them until your next turn.
'9. You must conquer South America And Europe and hold them until your next turn.
'10. You must conquer Europe and Australia and hold them until your next turn.
'11. You must conquer South America And Africa and hold them until your next turn.
'12. You must conquer Africa and Australia and hold them until your next turn.
'13. You must conquer South America And Australia and hold them until your next turn.
'14. You must occupy any 18 countries and hold them until your next turn.

'The number of countries for mission 14, "You must conquer and hold 18 countries"
Public Const gcMission14 As Integer = 18

Public Type MissionType
    DescriptionText As String   'Mission description shown to the player.
    WinMessageText As String    'Printed to the info box after a win.
    TargetArmy As Integer       'Mission to wipe out this army, 0 if none.
    TargetConts As String       'Comma delimited list of target continents, empty if none.
    IsActive As Boolean
End Type

Public gSeenMission(6) As Boolean
Public gMissions(14) As MissionType         'All available missions.
Public gAskedToSeeMission As Boolean        'False if not asked to see mission this turn.

'Returns true if the passed countinent is part of pPlayer's mission.
Public Function IsContPartOfMission(pContinent As Integer, pPlayer As Integer) As Boolean
    IsContPartOfMission = InStr(1, gMissions(gPlayerID(pPlayer).mission).TargetConts, CStr(pContinent)) > 0
End Function

'Populate the Missions Listbox with currently valid missions.
Public Sub PopulateMissionList()
    Dim vIndex As Integer
    
    TheMainForm.lstMissionList.Clear
    
    'Missions are on.
    If TheMainForm.chkMsnMissionsOn.Value = vbChecked Then
        
        'Kill player missions.
        If TheMainForm.chkMsnArmyWipeout.Value = vbChecked Then
            For vIndex = 1 To 6
                TheMainForm.lstMissionList.AddItem CStr(vIndex) & ". " & gMissions(vIndex).DescriptionText
            Next
        End If
        
        'Conquer and hold missions.
        If TheMainForm.chkMsnConquerHold.Value = vbChecked Then
            For vIndex = 7 To 13
                TheMainForm.lstMissionList.AddItem CStr(vIndex) & ". " & gMissions(vIndex).DescriptionText
            Next
        End If
        
        'Occupy 18 countries is not available when army wipeout mission
        'selected, continent hold mission is unselected and must complete
        'own mission is unselected.
        If TheMainForm.chkMsnArmyWipeout.Value = vbUnchecked _
        Or TheMainForm.chkMsnConquerHold.Value = vbChecked _
        Or TheMainForm.chkMsnMustComplete.Value = vbChecked Then
            TheMainForm.lstMissionList.AddItem CStr(14) & ". " & gMissions(14).DescriptionText
        End If
    
    'Missions are off, dominate world.
    Else
        TheMainForm.lstMissionList.AddItem gMissions(0).DescriptionText
    End If
    
    'Enable/disable mission option checkboxes that are applicable.
    Call EnableMissionOptions
End Sub

'Enable or disable mission options. Boolean parameter pForceOff
'will override if set to false.
Public Sub EnableMissionOptions(Optional pForceOff As Boolean = True)
    Dim vEnabled As Boolean
    
    With TheMainForm
    vEnabled = pForceOff _
    And netWorkSituation <> cNetClient _
    And .chkMsnMissionsOn.Enabled _
    And .chkMsnMissionsOn.Value = vbChecked
    
    .chkMsnArmyWipeout.Enabled = vEnabled
    .chkMsnConquerHold.Enabled = vEnabled
    
    'Enable/disable mission option checkboxes that are applicable.
    .chkMsnMustComplete.Enabled = vEnabled And .chkMsnArmyWipeout.Value = vbChecked
    .chkMsnWinImmediate.Enabled = vEnabled And .chkMsnArmyWipeout.Value = vbChecked
    .chkMsnAreUnique.Enabled = vEnabled And .chkMsnConquerHold.Value = vbChecked
    End With
End Sub

'Return the player's mission as text.
Public Function GetMissionDescriptionText(pPlayer) As String
    GetMissionDescriptionText = gMissions(gPlayerID(pPlayer).mission).DescriptionText
End Function

'Print all players' missions for testing.
Public Sub PrintPlayerMissions()
    Dim vIndex
    For vIndex = 1 To 6
        Debug.Print "player "; vIndex, GetMissionDescriptionText(vIndex)
    Next
End Sub

'Fill in mission information during application startup.
Public Sub InitializeMissions()
    Dim vIndex As Integer
    
    'Dominate world.
    gMissions(0).IsActive = True
    gMissions(0).DescriptionText = Phrase(75) '"You must wipe out all other players and conquer the world."
    gMissions(0).WinMessageText = Phrase(100) '"by dominating the world."
    gMissions(0).TargetArmy = 0
    gMissions(0).TargetConts = ""
    
    'Kill army missions.
    For vIndex = 1 To 6
        With gMissions(vIndex)
        .IsActive = True
        .DescriptionText = Phrase(76) & Phrase(vIndex + 361) + "."  '"Your mission is to wipe out the <Colour> Army."
        .WinMessageText = gPlayerID(vIndex).strColor & "."
        .TargetArmy = vIndex
        .TargetConts = ""
        End With
    Next
    
    'Conquer and hold missions.
    For vIndex = 7 To 14
        With gMissions(vIndex)
        .IsActive = True
        .DescriptionText = Phrase(70 + vIndex)
        .WinMessageText = Phrase(216 + vIndex)
        .TargetArmy = 0
        .TargetConts = ""
        End With
    Next
    
    'Set target continents.
    gMissions(7).TargetConts = "1,2"    'North America and South America
    gMissions(8).TargetConts = "1,6"    'North America And Australia
    gMissions(9).TargetConts = "2,3"    'South America And Europe
    gMissions(10).TargetConts = "3,6"   'Europe and Australia
    gMissions(11).TargetConts = "2,4"   'South America And Africa
    gMissions(12).TargetConts = "4,6"   'Africa and Australia
    gMissions(13).TargetConts = "2,6"   'South America And Australia
End Sub

'Clear missions from all playerIDs and set all missions to active.
Public Sub ClearMissions()
    Dim vIndex As Integer
    
    'Reset missions for each player.
    gWinMemoryBits = 0
    For vIndex = 1 To 6
        gSeenMission(vIndex) = False
        gPlayerID(vIndex).mission = 0
        gPlayerStats(vIndex).StartingMission = Phrase(75) 'You must wipe out all other players and conquer the world.
    Next
    
    'Set all missions to active.
    For vIndex = 0 To UBound(gMissions)
        gMissions(vIndex).IsActive = True
    Next
End Sub

'Deal missions to all players if missions are on.
Public Sub DealNewMissions()
    Const cMaxMissionAttempts As Long = 200         'How many times we try to deal missions.
    Const cMaxRedealAttempts As Long = 50           'How many times we re-deal missions due to deadlock.
    Dim vPlayer As Long
    Dim vDealAttempts As Long
    Dim vMission As Integer
    Dim vAttempts As Long
    Dim vDeadlock As Boolean
    
    'It is possible to have a mission deadlock. For example, 3 players - Red, Green, Blue.
    'Red must kill Green, Green must kill Red, Blue has nothing left to kill. This loop will
    'deal the missions again up to cMaxRedealAttempts (50) times if a deadlock is created.
    For vAttempts = 0 To cMaxRedealAttempts
        
        'Clear missions from playerIDs and set all missions to active.
        Call ClearMissions
        
        'Reset wipeout missions. Mark as active if there are 3 or more
        'players in the war and the target army is in the war.
        For vMission = 1 To 6
            gMissions(vMission).IsActive = TheMainForm.chkMsnArmyWipeout.Value = vbChecked _
                                        And TheMainForm.nmbrOfPlayers >= 3 _
                                        And gPlayerID(vMission).startWith > 0
        Next
        
        'Reset hold missions by marking all active.
        For vMission = 7 To UBound(gMissions)
            gMissions(vMission).IsActive = TheMainForm.chkMsnConquerHold.Value = vbChecked
        Next
        
        'Pick a mission for each player. Quick and dirty method used.
        For vPlayer = 1 To 6
             If gPlayerID(vPlayer).startWith > 0 Then
                
                'If missions are on and neither kill player nor continent missions
                'are selected, then everyone gets mission 14, 18 countries. Sweet.
                If TheMainForm.chkMsnArmyWipeout.Value = vbUnchecked _
                And TheMainForm.chkMsnConquerHold.Value = vbUnchecked Then
                    vDealAttempts = 0
                    gPlayerID(vPlayer).mission = 14
                
                'vDealAttempts (200) times to be sure. Defaults to 0 (dominate) if it
                'falls through which indicates a deadlock. Mission must be active and
                'the player should not get a mission to kill them selves.
                Else
                    For vDealAttempts = 0 To cMaxMissionAttempts
                        vMission = Int(GenRandom4 * UBound(gMissions)) + 1
                        If gMissions(vMission).IsActive And vMission <> vPlayer Then
                            
                            'Mission has been chosen for this player.
                            gPlayerID(vPlayer).mission = vMission
                            
                            'Conquer and hold missions selected as unique.
                            gMissions(vMission).IsActive = TheMainForm.chkMsnAreUnique.Value = vbUnchecked _
                                                        And gMissions(vMission).TargetConts <> ""
                            Exit For
                        End If
                    Next
                End If
                
                'If it fell through to here then there must have been
                'a deadlock. Flag the deadlock so we can deal missions again.
                vDeadlock = vDealAttempts >= cMaxMissionAttempts
            End If
        Next
        
        'Only try to deal the missions again if there is a deadlock flagged.
        If Not vDeadlock Then
            Exit For
        End If
    Next
    
    'Re activate "18 countries" mission incase other mission becomes invalid, like if
    'someone wiped out your target army and you must win your own mission to win.
    gMissions(14).IsActive = True
    
    'Remember starting missions for the stats report.
    Call RememberStartingMissions
End Sub

'Remember what missions were dealt to each player at the start of the war.
'Missions can change during war. frmStats uses these.
Public Sub RememberStartingMissions()
    Dim vIndex As Long
    
    For vIndex = 1 To 6
        gPlayerStats(vIndex).StartingMission = gMissions(gPlayerID(vIndex).mission).DescriptionText
        gSeenMission(vIndex) = False
    Next
End Sub
