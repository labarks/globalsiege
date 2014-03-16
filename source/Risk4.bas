Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------------------------------------------------
'Global Siege
'
'------------------------------------------------------------------------------
'                   Version History, public releases.
'
'MissionRisk
'26/11/1998 V1.00.06
'28/11/1998 v1.01.00
'27/01/1999 v1.02.00
'16/02/1999 v1.02.02
'25/02/1999 v2.00.00     started ~
'07/07/1999 v2.00.03
'17/09/1999 v2.00.04
'24/09/1999 v2.01.04     released
'20/10/1999 v2.01.05
'20/11/1999 v2.02.00
'17/03/2002 v2.06.00
'10/06/2002 v2.09.00     v3 beta 1
'07/07/2002 v2.09.00     v3 beta 2
'
'GlobalSiege
'16/01/2011 v0.9.0080   GlobalSiege Beta 1
'07/05/2011 v0.9.0127   GlobalSiege Beta 2
'14/12/2011 v0.9.0273   GlobalSiege Beta 3
'15/12/2011 v0.9.0279   GlobalSiege Beta 3.1
'29/12/2011 v0.9.0291   GlobalSiege Beta 3.2

'------------------------------------------------------------------------------
'                               Build History
'
'04/02/2010 V3.01.08    Added auditing system to prevent network cheating.
'15/02/2010 V3.01.18    Fixed mission dealing bug resolving hang caused by
'                       convoluted code in sub DealNewMissions().
'16/02/2010 V3.01.19    Improved card dealing functions. Cards are now laid
'                       out into an array making random selection much more
'                       robust.
'18/02/2010 V3.01.24    Screen stretching basically working.
'19/02/2010 V3.01.32    Mouse hit detection working.
'21/02/2010 V3.01.36    Full screen enabled hiding the title bar and menu.
'23/02/2010 V3.01.40    Added Clem's counter (mr#audit) and added more complex
'                       font commands to the chatter box.
'27/02/2010 V3.01.53    Updated the toolbar, scores are drawn directly on the
'                       front map (viewport) for clarity. Picture box moved back
'                       on the viewport and using a font resizing system to
'                       adjust the font size of the scores to fit correctly.
'                       End user can now select the font for the viewport.
'28/02/2010 V3.01.55    Preselect cards for humans, moved little card titles
'                       to the viewport for clarity.
'03/03/2010 V3.01.58    Unclaim on disconnect bug fixed in function lostPlayerOwner()
'06/03/2010 V3.01.74    Added deadlock detection to the mission dealing process
'                       DealNewMissions().
'                       Cheat code "mr#currentmode" now "mr#testing122"
'07/03/2010 V3.01.80    Voting system added.
'15/03/2010 V3.02.00    Map and masks updated.
'21/03/2010 V3.02.13    Cards and dice updated.
'23/03/2010 V3.02.20    Win XP look and feel using a manifest file.
'10/04/2010 V3.02.31    New random number generator using CryptGenRandom.
'11/10/2010 V3.02.50    Manifest files moved to a resource file for Windows 7.
'
'17/10/2010 V3.02.51    New name, "FarbeSieg"
'17/10/2010 V0.00.00    New name, "GlobalSiege"
'31/10/2010 V0.00.24    Added tabbed option box to the setup screen and started
'                       adding more available war options to choose from.
'15/01/2011 V0.9.80     Released  Beta 1. Legal notice prepared, website ready.
'19/01/2011 V0.9.81     Added cheat code mr#model to change the screen to a
'                       standard size for taking screenshots.
'07/05/2011 V0.9.127    Released Beta 2.
'14/12/2011 V0.9.273    Released Beta 3.
'15/12/2011 V0.9.279    Released Beta 3.1.
'29/12/2011 V0.9.291    Released Beta 3.2.
'  --- Discontinued. Reffer to the GlobalSiege Technical Manual for the change log. ---
'
'------------------------------------------------------------------------------
'
'       ** Checklist: Make sure the following is done before any public release **
'
'   - gcAppDevelopMode is set to FALSE
'   - gcAppTestingMode is set to FALSE
'   - gcDefaultHomePageClearURL is set to "http://www.globalsiege.net"
'   - gcOnlineServerBaseURL is set to "http://www.globalsiege.net/gs/v090200"
'
'   - Also check that c_DEBUG_MODE is set to FALSE in gs-indexing-server.php
'     on line 15.
'------------------------------------------------------------------------------

'** Set to FALSE before public release **
'Set to TRUE for comprehensive unencrypted logging of application data.
Public Const gcAppDevelopMode As Boolean = False

'** Set to FALSE before public release **
'Set to TRUE for application testing. This will show testpoint text boxes on the main form.
Public Const gcAppTestingMode As Boolean = False

'Application name.
Public Const gcApplicationName As String = "GlobalSiege"

'Default Home web page.
Public Const gcDefaultHomePageClearURL As String = "http://globalsiege.net"
Public Const gcDefaultHomePageSecureURL As String = "https://globalsiege.net"

'** Set to http://www.globalsiege.net/gs/v090200 before public release **
'Application web home.
Public Const gcOnlineServerBaseURL As String = "http://globalsiege.net/gs/v00090200"
'Public Const gcOnlineServerBaseURL As String = "http://10.1.1.102/gs/v00090200"

'Default Help web page.
Public Const gcDefaultHelpPageURL As String = gcDefaultHomePageClearURL & "/docs/"

'Default Download web page.
Public Const gcDefaultDownloadPageURL As String = gcDefaultHomePageClearURL & "/download/"

'Default create login account web page.
Public Const gRegisterAccountWebPage As String = gcDefaultHomePageSecureURL & "/wp-login.php?action=register"

'Default Lost Password web page.
Public Const gLostPasswordWebPage  As String = gcDefaultHomePageSecureURL & "/wp-login.php?action=lostpassword"

'Main index page containing version updates and other details such as
'home pages and other directives.
Public Const gcNewsServerURL As String = gcOnlineServerBaseURL & "/news/"

'Default Session Indexing Server web pages.
Public Const gcIndexServerURL As String = gcOnlineServerBaseURL & "/indexserver/"

'Minimum starting units distributed at the start of a player's turn.
Public Const gcStartUnitsMin As Integer = 3

'Increasing card values at the start of the war.
Public Const gcCardStartValue As Integer = 5

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As _
    String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32.dll" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

'Used to get the default language.
Public Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer

'Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'Constants
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
Public Const KEY_READ = &H20019

Public Const remoteIndex As Long = 3           '"Remote player" list index
Public Const A3Index As Long = 4
Public Const playFast As Integer = 12          'Fast clock speed
Public Const playSlow As Integer = 100         'Slow clock speed
Public Const diceFast As Integer = 10           'Fast dice
Public Const diceSlow As Integer = 300          'Slow dice speed

Public Const fileBuffer = 18                    'Size of spare space in save game file

Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type

Public Type CountryIDType
    ctryName As String * 30
    destX As Integer
    destY As Integer
    Width As Integer
    Height As Integer
    srcX As Integer
    srcY As Integer
    printX As Integer
    printY As Integer
    PrintRules As String        'Rules for printing: L=left, R=right, W=white (default to mid and black)
    neighbour(7) As Integer
End Type

'This cannot be changed without corrupting saved games.
Public Type playType
    strColor As String
    lngColor As Long
    srfsLite As Boolean     'Lighten color using darker mask if T
    MaskIndex As Long       'Hdc of source
    txtColor As Long
    bkgndColor As Long
    card(10) As Integer
    pickedCards(10) As Boolean
    startWith As Integer
    mission As Integer
    playerWho As Integer
    'remoteComputer As Byte  'Client no for each player for network
End Type

Public Type ContinentType
    ContNameText As String
    FirstCountry As Integer
    LastCountry As Integer
    'ContUnitValue As Integer
    GateCountries(5) As Integer
    ContPriority As Integer
End Type

Public Type coord
    Xpos As Integer
    Ypos As Integer
End Type

Public Type Attack
    from As Integer
    To As Integer
    On As Boolean
End Type

Public Type regData                    'Rego checking format
    iCode(0 To 50) As Integer
    sName As String * 50
End Type

Public Type WarControlType          'Save to file as
    filename As String * 50              'File name
    fileDescription As String * 500      'Description
    playerStart(5) As Integer       'Start ctrys
    playerCtrlr(5) As Integer   '0,1,2=Human,Av,Smart
    crdHidden As Integer            'hidden
    capture As Integer              'Vulture cards
    cardmode As Integer             '0,1,2=none,fixed,incr
    cardMax As Integer
    firstPlayer As Integer          'T=red first
    warMissions As Integer           '1 for on
    warSupply As Integer            'supply lines
    warSupplyLimit As Integer
    warSupplyNo As Integer
    warOptDice As Integer           'Deprecated
    warFastWar As Integer           '1 for fast war
    warFastDice As Integer          '1 for fast dice
    warBorder As Integer            'Flashing border
    nmbrPlrs As Integer             'Number of players
    Locked As Boolean               'Locked if True (cant modify)
    SetupScreen As Boolean
    tmpTimer1 As Boolean                'Temp storage for timer values
    tmpTimer2 As Boolean
    cardstmp As Boolean                 'VultureCards memory
    PrevMode As Integer                 'Previous mode (New game cancel)
    
    sCtryScore(49) As Integer       'Individual country scores
    sCtryOwner(49) As Integer       'player on country
    sNmbrOfPlayers As Integer            'Number of players
    sPlayerID(6) As playType
    sPlayerTurn As Integer               'Whos go
    sCurrentCardValue As Integer         'Value of cards
    sGateDefence As Single               'the lower, more defensive smart players (2.5)
    sBoolIssueCard As Boolean
    sCards(3) As Integer           'Cards left in pack
    
    kCardsUp As Boolean
    kMaxCardValue As Integer
    kMissionsOn As Boolean
    kMoveLimit As Integer
    kPlaySpeed As Integer
    kFlashingBorder As Boolean
    kDiceSpeed As Integer
    kOptimizeDice As Boolean            'Depricated
    kCardMode As Integer
    
    chkExtraStartingUnits As Boolean
    distUnits As Byte
    
    'New stuff.
    GSVersion As String                         'Version info to help with loading new stuff
    'Dice setup info.
    optDiceRules(10) As Boolean                   '0-4 in use, 5-10 as spare.
    chkSortDice As Byte
    optDiceSame(10) As Boolean                      '0-2 in use, 3-10 spare.
    udDiceThrown(5) As Byte                        '0-1 taken 2-5 spare.
    
    'Card setup info.
    udCardDeck(10) As Byte                          '0-3 taken, 4-10 spare.
    udFixedValues(10) As Byte                       '0-3 taken, 4-10 spare.
    
    'Reinforcments tab.
    udNewUnitClac(2)  As Long
    udContVal(5)      As Long
    
End Type

Public Type AuditPlayerType
    Player As Long                      'Keep track of country count per player to prevent cheating.
    PlayerMax As Long                   'Ensure extra points are not "sneaked" in somehow.
    PlayerCard As Long                  'Keep track of total player card value.
End Type

Public Type PreviousSettingsType
    PrevMode As Integer                 'Previous mode (New game cancel)
    PrevBorder As Long                  'Previou border (New game cancel)
    Advanced As Byte                    'Advanced setup settings
End Type

Public Type registryData
    playerAI110 As String       'GetSetting(gcApplicationName, "settings", "playerAI110")
    playerAI As String          'GetSetting(gcApplicationName, "settings", "playerAI")
    Language As String          'CInt(GetSetting(gcApplicationName, "settings", "Lang", 0))
    State As String             'GetSetting(gcApplicationName, "settings", "state")
    toolbox As Boolean          'GetSetting(gcApplicationName, "settings", "toolBox", "True")
    seeReshuf As Boolean        'GetSetting(gcApplicationName, "settings", "seeReshuf")
    Start As String             'Trim (GetSetting(gcApplicationName, "settings", "start"))
    supPersonal As Integer      'CByte(GetSetting(gcApplicationName, "settings", "supPersonal", 12))
    InstallDate As Date         'GetSetting(gcApplicationName, "settings", "InstallDate")
    lastPortV26 As Integer         'CInt(GetSetting(gcApplicationName, "settings", "lastPortV26", Str(gcDefaultPortNumber)))
    lastHost As String          'GetSetting(gcApplicationName, "settings", "lastHost", "")
    rfrsRateV29 As Integer         'CInt(GetSetting(gcApplicationName, "settings", "rfrsRateV29", "1"))
    
    regName As String
    RegCode As String
End Type

Public Type evalData
    timeLast As Long
    TimeNow As Long
    dateInIniFile As Long
    regName As String
    dateInRegistry As Long
    RegCode As String
    dateNow As Long
    runsToday As Integer
    dateDiff As Integer
    prevVersion As Boolean
    failReg As Boolean
    fileCS As String
End Type

Public Type cheat                      'Different cheat modes
    inCodes(10) As String
    responses(10) As String
    testing As Boolean
    createMap As Boolean
    seeMissions As Boolean
    seeCards As Boolean
    undoEnabled As Boolean
    autoRestart As Boolean
    cheatActive As Boolean
End Type

Public Type A2opportunityList
    IsActive As Boolean                   'True if path found
    pathPointer As Integer              'Current position in path
    unitsRequired As Integer            'Points needed
    Path(42) As Integer                 'Attack path
    StopWhenFinished As Boolean         'Prevent opening week spots
End Type

Public Type PlayerWarStatistics
    StartingMission As String
    UnitsBeaten As Long
    UnitsLost As Long
    CountriesDefeated As Long
    CountriesLost As Long
    PlayersWipedOut As Long
    UnitsIssued As Long
    CardsTraded As Long
    UnitsFromCards As Long
    PlrController As Long
    IsValid As Boolean                  'Stats are valid only if player has played the game from the start.
    InvalidatedReason As String         'Reason for invalidation above.
End Type

Public Type MoveList
    List(7) As Integer              'List of moves found (during AutoMode10)
    Transfer(7) As Integer          'Tranfer rate of each move
    Pointer As Integer              'Pointer to MoveList elements
End Type

Public Type CoreType
    Message As String               'Reason for core dump
End Type

Public CountryID(49) As CountryIDType    'Individual country data
Public Continents(5) As ContinentType
Public ContPriority(5) As Integer         'Order of priority
Public gCtryScore(49) As Integer       'Individual country scores
Public warSit As WarControlType                 'Save as
Public A2MoveList As MoveList
Public gPlayerID(6) As playType
Public gPlayerStats(6) As PlayerWarStatistics
Public A2opportunity As A2opportunityList
Public gCheatMode As cheat
Public evalChk As evalData
Public CoreDump As CoreType
Public gPlayerTurn As Integer               'Whos go
Public gSourceCtry As Integer              'Attacking or move from country
Public gTargetCtry As Integer                'Defending or recieving country
Public gCurrentMode As Integer              'Attack=1, move=10 etc.. Test=0
Public gPauseActive As Boolean                  'True if paused
Public gPickedUpUnits As Integer            'Units player has picked up fo transfer
Public gCountryOwner(49) As Integer       'player on country
Public gServerMode As Boolean               'Is this running as a headless server?
Public gHeadlessMode As Boolean             'No display if set to TRUE
Public gGsLeUtils As CGsLeUtils           'Light encryption functions

'Program starts here.
Sub Main()
    Dim iccex As InitCommonControlsExStruct
    Dim hMod As Long
    
    ' constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx
    Const ICC_ANIMATE_CLASS As Long = &H80&
    Const ICC_BAR_CLASSES As Long = &H4&
    Const ICC_COOL_CLASSES As Long = &H400&
    Const ICC_DATE_CLASSES As Long = &H100&
    Const ICC_HOTKEY_CLASS As Long = &H40&
    Const ICC_INTERNET_CLASSES As Long = &H800&
    Const ICC_LINK_CLASS As Long = &H8000&
    Const ICC_LISTVIEW_CLASSES As Long = &H1&
    Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&
    Const ICC_PAGESCROLLER_CLASS As Long = &H1000&
    Const ICC_PROGRESS_CLASS As Long = &H20&
    Const ICC_TAB_CLASSES As Long = &H8&
    Const ICC_TREEVIEW_CLASSES As Long = &H2&
    Const ICC_UPDOWN_CLASS As Long = &H10&
    Const ICC_USEREX_CLASSES As Long = &H200&
    Const ICC_STANDARD_CLASSES As Long = &H4000&
    Const ICC_WIN95_CLASSES As Long = &HFF&
    Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all values above

    With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_STANDARD_CLASSES _
                Or ICC_BAR_CLASSES _
                Or ICC_COOL_CLASSES _
                Or ICC_HOTKEY_CLASS _
                Or ICC_INTERNET_CLASSES _
                Or ICC_LISTVIEW_CLASSES _
                Or ICC_NATIVEFNTCTL_CLASS _
                Or ICC_PROGRESS_CLASS _
                Or ICC_TAB_CLASSES _
                Or ICC_TREEVIEW_CLASSES _
                Or ICC_UPDOWN_CLASS _
                Or ICC_USEREX_CLASSES _
                Or ICC_STANDARD_CLASSES
       '.lngICC = ICC_STANDARD_CLASSES ' vb intrinsic controls (buttons, textbox, etc)
       ' if using Common Controls; add appropriate ICC_ constants for type of control you are using
       ' example if using CommonControls v5.0 Progress bar:
       ' .lngICC = ICC_STANDARD_CLASSES Or ICC_PROGRESS_CLASS
    End With
    On Error Resume Next ' error? InitCommonControlsEx requires IEv3 or above
    hMod = LoadLibraryA("shell32.dll") ' patch to prevent XP crashes when VB usercontrols present
    InitCommonControlsEx iccex
    If Err Then
        InitCommonControls ' try Win9x version
        Err.Clear
    End If
    
    On Error GoTo 0
    Set gGsLeUtils = New CGsLeUtils  'Initialise encryption functions.
    'On Error Resume Next
    
    'Dim vResp As VbMsgBoxResult
    'vResp = MsgBox("here", vbOKOnly)
    'Load TheMainForm
    'If GetDeviceCaps(TheMainForm.Picture1.hdc, 12) < 24 Then
    '    vResp = MsgBox(LimitTextWidth(Phrase(89), 50), vbOKOnly)
    'End If
    'TheMainForm.Show
    Load TheMainForm
    If hMod Then FreeLibrary hMod

'** Tip 1: Avoid using VB Frames when applying XP/Vista themes
'          In place of VB Frames, use pictureboxes instead.
'** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
'          Doing so will prevent them from being themed.
End Sub

'Return the app version info formatted nicely.
Public Function GetVersionInfo(Optional pSeperator As String = "") As String
    GetVersionInfo = Format(CStr(App.Major), "00") & pSeperator _
                    & Format(App.Minor, "00") & pSeperator _
                    & Format(App.Revision, "0000")
End Function

'------------------------------------------------------------------------------------
'Get and set if required a unique ID.
Public Function GetUniqueId() As String
    Static vUid As String
    
    If vUid = "" Then
        vUid = GetSetting(gcApplicationName, "settings", "dllptr02", "NA")
    End If
    
    If vUid = "NA" Then
        vUid = netMain.BytToHex(gGsLeUtils.LE1(CreateUniqueId()))
        SaveSetting gcApplicationName, "settings", "dllptr02", vUid
    End If
    GetUniqueId = gGsLeUtils.LE1d(netMain.HexToByte(vUid))
End Function

'Greate a unique ID for this user.
'Format: encrypted <Timestamp>:<Host name>:<User name>
Public Function CreateUniqueId() As String
    Dim vDetails As String
    Dim i As Long
    Dim vLen As Long
    Dim vTimeStamp As String
    
    On Error Resume Next
    
    'vDetails = Replace(GetMACAddress(), " ", "")
    vDetails = Format(Now, "ssnnhhddmmyy") & ":" _
                & GetLocalHostName & ":" _
                & GetLoggedInUserName
    
    CreateUniqueId = gGsLeUtils.LE6(vDetails)
    
End Function

' Get the the user name for the currently looged in user from the system.
Public Function GetLoggedInUserName() As String
    Dim vHoldStr
    Dim vUserName As String * 100
    Dim vBuffLen As Long
    
    vBuffLen = 99
    vHoldStr = GetUserName(vUserName, vBuffLen)
    vBuffLen = 1
    Do While Asc(Mid(vUserName, vBuffLen, 1)) > 0
        vBuffLen = vBuffLen + 1
    Loop
    GetLoggedInUserName = Mid(vUserName, 1, vBuffLen - 1)
End Function

'RegNow code for affiliate tracking.
Private Function regGetBuyURL(ByVal publisher As String, ByVal appName As String, ByVal appVer As String) As String
    Dim hKey As Long    ' receives a handle opened registry key
    Dim stringbuffer As String  ' receives data read from the registry
    Dim datatype As Long  ' receives data type of read value
    Dim slength As Long  ' receives length of returned data
    Dim retval As Long  ' return value
    
    ' form the registry key path
    Dim keyPath
    keyPath = "SOFTWARE\Digital River\SoftwarePassport\" & publisher & "\" & appName & "\" & appVer
            
    ' open the registry key
    ' try to get from HKEY_LOCAL_MACHINE first
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, hKey)
    ' if fail to get from HKEY_LOCAL_MACHINE branch, try HKEY_CURRENT_USER
    If retval <> 0 Then
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, keyPath, 0, KEY_READ, hKey)
    End If
    
    If retval = 0 Then
        ' Make room in the buffer to receive the incoming data.
        stringbuffer = Space(1024)
        slength = 1024
        
        ' Read the "BuyURL" value from the registry key.
        retval = RegQueryValueEx(hKey, "BuyURL", 0, datatype, ByVal stringbuffer, slength)
        If retval = 0 Then
            stringbuffer = Left(stringbuffer, slength - 1)
        Else
            ' "BuyURL" does not exists
            stringbuffer = ""
        End If
        
        ' Close the registry key.
        retval = RegCloseKey(hKey)
    End If
    
    regGetBuyURL = stringbuffer
End Function

'Print ascii char set.
Public Sub test5()
    Dim i As Long
    For i = 0 To 255
        Debug.Print i; " = "; Chr(i),
        If i Mod 10 = 0 Then
            Debug.Print
        End If
    Next
End Sub

'----------------------------------------------------------------------------

#Const CompileBlock = False
#If CompileBlock Then


#End If
