VERSION 5.00
Begin VB.Form Mask4 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000011&
   Caption         =   "Mask4"
   ClientHeight    =   12120
   ClientLeft      =   -90
   ClientTop       =   1935
   ClientWidth     =   14145
   Icon            =   "Mask4.frx":0000
   LinkTopic       =   "Form1"
   Palette         =   "Mask4.frx":000C
   PaletteMode     =   2  'Custom
   ScaleHeight     =   808
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   943
   Begin VB.Frame frameMasks 
      Caption         =   "Masks"
      Height          =   6015
      Left            =   0
      TabIndex        =   10
      Top             =   5280
      Width           =   6495
      Begin VB.PictureBox pctMaskArray 
         AutoRedraw      =   -1  'True
         Height          =   2535
         Index           =   4
         Left            =   2400
         Picture         =   "Mask4.frx":12BD
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   16
         Top             =   3000
         Width           =   3975
      End
      Begin VB.PictureBox pctMaskArray 
         AutoRedraw      =   -1  'True
         Height          =   2535
         Index           =   3
         Left            =   1800
         Picture         =   "Mask4.frx":1CEEA
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   14
         Top             =   2400
         Width           =   3975
      End
      Begin VB.PictureBox pctMaskArray 
         AutoRedraw      =   -1  'True
         Height          =   2535
         Index           =   2
         Left            =   1080
         Picture         =   "Mask4.frx":2F65D
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   13
         Top             =   1680
         Width           =   3975
      End
      Begin VB.PictureBox pctMaskArray 
         AutoRedraw      =   -1  'True
         Height          =   2535
         Index           =   1
         Left            =   600
         Picture         =   "Mask4.frx":3BB9A
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   12
         Top             =   960
         Width           =   3975
      End
      Begin VB.PictureBox pctMaskArray 
         AutoRedraw      =   -1  'True
         Height          =   2535
         Index           =   0
         Left            =   120
         Picture         =   "Mask4.frx":4D9BA
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   11
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame frmMap1 
      Caption         =   "Map1"
      Height          =   10455
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   13335
      Begin VB.PictureBox Map1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9810
         Left            =   120
         Picture         =   "Mask4.frx":5E901
         ScaleHeight     =   650
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   944
         TabIndex        =   5
         Top             =   360
         Width           =   14220
      End
   End
   Begin VB.Frame pctCards 
      Caption         =   "Cards - Position info in Mask4's code"
      Height          =   4335
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   5655
      Begin VB.PictureBox pctCardSource 
         AutoRedraw      =   -1  'True
         Height          =   2415
         Left            =   120
         Picture         =   "Mask4.frx":8C5A1
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   18
         Top             =   240
         Width           =   3975
      End
      Begin VB.PictureBox pctDice 
         AutoRedraw      =   -1  'True
         Height          =   2415
         Left            =   4320
         Picture         =   "Mask4.frx":B1B8B
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   17
         Top             =   600
         Width           =   3615
      End
      Begin VB.PictureBox pctLittleCrdSource 
         AutoRedraw      =   -1  'True
         Height          =   900
         Left            =   120
         Picture         =   "Mask4.frx":CE09D
         ScaleHeight     =   56
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   164
         TabIndex        =   3
         Top             =   3360
         Width           =   2520
      End
   End
   Begin VB.Frame frmMap1Blanks 
      Caption         =   "Map1 Blanks"
      Height          =   3255
      Left            =   480
      TabIndex        =   0
      Top             =   9360
      Width           =   12615
      Begin VB.PictureBox pctLittleCards 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   600
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   781
         TabIndex        =   15
         Top             =   120
         Width           =   11775
      End
      Begin VB.PictureBox pctClearDice 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000E&
         Height          =   1740
         Left            =   120
         Picture         =   "Mask4.frx":D76F3
         ScaleHeight     =   116
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   217
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.PictureBox pctMainCards 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   7560
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   269
         TabIndex        =   7
         Top             =   1320
         Width           =   4095
      End
      Begin VB.PictureBox pctVultureCards 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   1680
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   269
         TabIndex        =   1
         Top             =   1320
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Mask4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Me.Hide
End Sub

'Load contsants saved in this mask form.
Public Sub LoadMaskConstants()
    'Dim vCtry As Long
    'Dim vNbr As Long
    
    Call LoadPictureConstants
    Call LoadAllCountryData
    Call LoadContinents
    
    ''Use to assist load the country data way back in 15/03/2010
    ''Call LoadMaskData
    'For vCtry = 1 To 42
    '    With CountryID(vCtry)
    '    Debug.Print "LoadCountryData"; vCtry; ","; .destX; ","; .destY; ","; _
    '                .Width; ","; .Height; ","; .srcX; ","; .srcY; ","; """"; Trim(.ctryName); """";
    '
    '    For vNbr = 1 To 7
    '        Debug.Print ","; CountryID(vCtry).neighbour(vNbr);
    '    Next
    '    Debug.Print ","; .printX; ","; .printY; ","; """"; .PrintRules; """"
    '    End With
    'Next
End Sub

'Load all constant country data into the global variable CountryID(1-42).
'Format: <Country Number>, <DestX>, <DextY>, <Width>, <Height>, <SourceX>, <SourceY>, <Country Name>,
'<Neighbor1>, ..., <Nieghbor7>, <Score Position X>, <Score Position Y>, <Score Attributes>
'Where <DestX> and <DestY> are the X and Y coordinates on the destination map Mask4.Map1
'<SourceX> and <SourceY> are the X and Y coordinates on the blitting masks Map4.pctMaskArray(0-4)
'<Score Position X>, <Score Position Y> are the X and Y coordinates of the country unit value
'<Score Attributes> is the rules governing how the score is printed.
'"M" will keep the score centered on the X and Y score position.
'"L" will treat the X and Y position as the left most position to print the score.
'The score will be printed from this position.
'"R" will trteat the X and Y position as the right most position the score can be printed.
'The score will extend leftwards of this position.
'"W" will print the score using a white font. This is for when the score is in the water off the country.
Private Sub LoadAllCountryData()
    LoadCountryData 1, 27, 112, 84, 105, 7, 1, "Alaska", 2, 4, 38, 0, 0, 0, 0, 71, 155, "M"
    LoadCountryData 2, 103, 73, 102, 107, 92, 1, "Northwest Territories", 1, 3, 4, 5, 0, 0, 0, 142, 150, "M"
    LoadCountryData 3, 191, 70, 148, 119, 195, 1, "Greenland", 2, 5, 6, 14, 0, 0, 0, 297, 120, "M"
    LoadCountryData 4, 106, 165, 71, 77, 1, 107, "British Columbia", 1, 2, 5, 7, 0, 0, 0, 143, 207, "M"
    LoadCountryData 5, 170, 173, 85, 93, 318, 269, "Ontario", 2, 3, 4, 6, 7, 8, 0, 203, 226, "M"
    LoadCountryData 6, 239, 188, 82, 76, 149, 121, "Quebec", 3, 5, 8, 0, 0, 0, 0, 264, 226, "M"
    LoadCountryData 7, 120, 233, 67, 78, 232, 121, "Western United States", 4, 5, 8, 9, 0, 0, 0, 152, 269, "M"
    LoadCountryData 8, 181, 235, 106, 98, 300, 121, "Eastern United States", 5, 6, 7, 9, 0, 0, 0, 208, 278, "M"
    LoadCountryData 9, 136, 299, 105, 98, 1, 201, "Mexico", 7, 8, 10, 0, 0, 0, 0, 178, 363, "RW"
    LoadCountryData 10, 206, 378, 86, 44, 107, 201, "Colombia", 9, 11, 12, 0, 0, 0, 0, 238, 394, "L"
    LoadCountryData 11, 203, 418, 84, 92, 107, 246, "Peru", 10, 12, 13, 0, 0, 0, 0, 253, 472, "M"
    LoadCountryData 12, 239, 399, 100, 126, 1, 300, "Brazil", 10, 11, 13, 21, 0, 0, 0, 278, 437, "M"
    LoadCountryData 13, 244, 497, 43, 110, 344, 1, "Argentina", 11, 12, 0, 0, 0, 0, 0, 260, 533, "M"
    LoadCountryData 14, 338, 183, 44, 36, 1, 427, "Iceland", 3, 15, 16, 0, 0, 0, 0, 355, 176, "W"
    LoadCountryData 15, 371, 229, 41, 52, 1, 464, "Great Britain", 14, 16, 17, 18, 0, 0, 0, 377, 244, "RW"
    LoadCountryData 16, 413, 169, 51, 78, 46, 427, "Scandinavia", 14, 15, 17, 20, 0, 0, 0, 433, 175, "RW"
    LoadCountryData 17, 413, 230, 70, 55, 1, 517, "Germania", 15, 16, 18, 19, 20, 0, 0, 454, 261, "M"
    LoadCountryData 18, 377, 272, 55, 60, 102, 339, "Spain", 15, 17, 19, 21, 0, 0, 0, 405, 296, "M"
    LoadCountryData 19, 424, 271, 74, 56, 194, 200, "The Mediterranean", 17, 18, 20, 21, 22, 27, 0, 459, 285, "M"
    LoadCountryData 20, 448, 169, 120, 142, 407, 1, "Prussia", 16, 17, 19, 27, 28, 32, 0, 515, 246, "M"
    LoadCountryData 21, 349, 330, 95, 106, 407, 144, "Algeria", 12, 18, 19, 22, 23, 24, 0, 389, 384, "M"
    LoadCountryData 22, 422, 333, 79, 54, 255, 581, "Egypt", 19, 21, 23, 27, 27, 0, 0, 454, 360, "M"
    LoadCountryData 23, 436, 377, 94, 106, 102, 400, "Ethiopia", 21, 22, 24, 25, 26, 27, 0, 476, 406, "M"
    LoadCountryData 24, 414, 411, 58, 60, 158, 339, "Congo", 21, 23, 25, 0, 0, 0, 0, 439, 444, "M"
    LoadCountryData 25, 418, 459, 77, 90, 197, 400, "South Africa", 23, 24, 26, 0, 0, 0, 0, 449, 497, "M"
    LoadCountryData 26, 507, 458, 29, 62, 192, 257, "Madagascar", 23, 25, 0, 0, 0, 0, 0, 535, 471, "LW"
    LoadCountryData 27, 473, 276, 95, 119, 222, 269, "Saudi Arabia", 19, 20, 22, 23, 28, 29, 0, 515, 353, "M"
    LoadCountryData 28, 548, 253, 73, 96, 275, 389, "Afghanistan", 20, 27, 29, 32, 33, 0, 0, 579, 297, "M"
    LoadCountryData 29, 562, 306, 81, 108, 595, 261, "India", 27, 28, 30, 33, 0, 0, 0, 600, 350, "M"
    LoadCountryData 30, 634, 345, 71, 81, 73, 109, "Cambodia", 29, 33, 39, 0, 0, 0, 0, 675, 374, "M"
    LoadCountryData 31, 638, 124, 113, 155, 528, 1, "Siberia", 32, 33, 34, 35, 36, 0, 0, 682, 201, "M"
    LoadCountryData 32, 558, 83, 131, 200, 463, 251, "Krasnoyarsk", 20, 28, 31, 33, 0, 0, 0, 603, 199, "M"
    LoadCountryData 33, 606, 266, 134, 108, 349, 452, "China", 28, 29, 30, 31, 32, 34, 0, 669, 315, "M"
    LoadCountryData 34, 679, 257, 92, 67, 349, 378, "Korea", 31, 33, 35, 37, 38, 0, 0, 731, 279, "M"
    LoadCountryData 35, 706, 184, 87, 83, 503, 157, "Magadan", 31, 34, 36, 38, 0, 0, 0, 741, 227, "M"
    LoadCountryData 36, 727, 126, 137, 89, 197, 491, "Chukotka", 31, 35, 38, 0, 0, 0, 0, 796, 162, "M"
    LoadCountryData 37, 773, 252, 56, 69, 72, 507, "Japan", 34, 38, 0, 0, 0, 0, 0, 805, 312, "LW"
    LoadCountryData 38, 762, 131, 164, 150, 484, 452, "Kamchatka", 1, 34, 35, 36, 37, 0, 0, 865, 179, "M"
    LoadCountryData 39, 679, 404, 103, 67, 1, 577, "Indonesia", 30, 40, 41, 0, 0, 0, 0, 733, 426, "M"
    LoadCountryData 40, 754, 473, 78, 86, 335, 561, "Western Australia", 39, 41, 42, 0, 0, 0, 0, 789, 515, "M"
    LoadCountryData 41, 790, 418, 86, 34, 484, 603, "New Guinea", 39, 40, 42, 0, 0, 0, 0, 825, 412, "MW"
    LoadCountryData 42, 821, 473, 90, 103, 595, 157, "Eastern Australia NZ", 40, 41, 0, 0, 0, 0, 0, 836, 518, "M"
End Sub

'Load individual country data.
'Rules for printing: L=left, M=middle, R=right, W=white (default to mid and black)
Private Sub LoadCountryData(pID As Long, pMapX As Long, pMapY As Long, _
pWidth As Long, pHeight As Long, pMaskX As Long, pMaskY As Long, pName As String, _
 Neighbor1 As Integer, Neighbor2 As Integer, Neighbor3 As Integer, _
 Neighbor4 As Integer, Neighbor5 As Integer, Neighbor6 As Integer, Neighbor7 As Integer, _
 pPrintScoreX As Long, pPrintScoreY As Long, Optional pPrintScoreRules As String = "")
    With CountryID(pID)
        .destX = pMapX
        .destY = pMapY
        .Width = pWidth
        .Height = pHeight
        .srcX = pMaskX
        .srcY = pMaskY
        .ctryName = pName
        .neighbour(1) = Neighbor1
        .neighbour(2) = Neighbor2
        .neighbour(3) = Neighbor3
        .neighbour(4) = Neighbor4
        .neighbour(5) = Neighbor5
        .neighbour(6) = Neighbor6
        .neighbour(7) = Neighbor7
        .printX = pPrintScoreX
        .printY = pPrintScoreY
        .PrintRules = pPrintScoreRules
    End With
End Sub

'Load constans for each continent.
'Old old code but still used.
Private Sub LoadContinents()
    Dim vIndex As Integer
    Dim tmp
        
    Continents(0).ContNameText = " North America "
    'Continents(0).ContUnitValue = 6
    Continents(0).FirstCountry = 1
    Continents(0).LastCountry = 9
    Continents(0).ContPriority = 4
    tmp = Array(1, 3, 9, 0, 0)
    For vIndex = 1 To 5
        Continents(0).GateCountries(vIndex) = tmp(vIndex - 1)
    Next vIndex
    ContPriority(2) = 1
    
    Continents(1).ContNameText = " South America "
    'Continents(1).ContUnitValue = 3
    Continents(1).FirstCountry = 10
    Continents(1).LastCountry = 13
    Continents(1).ContPriority = 5
    tmp = Array(10, 12, 0, 0, 0)
    For vIndex = 1 To 5
        Continents(1).GateCountries(vIndex) = tmp(vIndex - 1)
    Next vIndex
    ContPriority(1) = 2
    
    Continents(2).ContNameText = " Europe "
    'Continents(2).ContUnitValue = 6
    Continents(2).FirstCountry = 14
    Continents(2).LastCountry = 20
    Continents(2).ContPriority = 2
    tmp = Array(14, 18, 19, 20, 0)
    For vIndex = 1 To 5
        Continents(2).GateCountries(vIndex) = tmp(vIndex - 1)
    Next vIndex
    ContPriority(4) = 3
    
    Continents(3).ContNameText = " Africa "
    'Continents(3).ContUnitValue = 4
    Continents(3).FirstCountry = 21
    Continents(3).LastCountry = 26
    Continents(3).ContPriority = 3
    tmp = Array(21, 22, 23, 0, 0)
    For vIndex = 1 To 5
        Continents(3).GateCountries(vIndex) = tmp(vIndex - 1)
    Next vIndex
    ContPriority(3) = 4
    
    Continents(4).ContNameText = " Asia "
    'Continents(4).ContUnitValue = 8
    Continents(4).FirstCountry = 27
    Continents(4).LastCountry = 38
    Continents(4).ContPriority = 1
    tmp = Array(27, 28, 30, 32, 38)
    For vIndex = 1 To 5
        Continents(4).GateCountries(vIndex) = tmp(vIndex - 1)
    Next vIndex
    ContPriority(5) = 5
    
    Continents(5).ContNameText = " Australia "
    'Continents(5).ContUnitValue = 3
    Continents(5).FirstCountry = 39
    Continents(5).LastCountry = 42
    Continents(5).ContPriority = 6
    tmp = Array(39, 0, 0, 0, 0)
    For vIndex = 1 To 5
        Continents(5).GateCountries(vIndex) = tmp(vIndex - 1)
    Next vIndex
    ContPriority(0) = 6
End Sub

'Load various constants from picture boxes on Mask4
Private Sub LoadPictureConstants()
    Dim vIX As Long
    
    'Main Cards container on the map.
    gMsk.CrdMainLeft = 520
    gMsk.CrdMainTop = 560
    gMsk.CrdMainWidth = 264
    gMsk.CrdMainHeight = 77
    
    'Vulture Cards container on the map.
    gMsk.CrdVultLeft = 545 '440
    gMsk.CrdVultTop = 482
    gMsk.CrdVultWidth = 212
    gMsk.CrdVultHeight = 77
    
    'Single card size and buffer size.
    gMsk.CrdSnglWidth = 60 '51
    gMsk.CrdSnglHeight = 85 '75
    gMsk.CrdSnglSrcBuffer = 7
    
    'Card destinations within card containers.
    gMsk.CrdDestX(1) = 0
    gMsk.CrdDesty(1) = 0
    gMsk.CrdDestX(2) = 53
    gMsk.CrdDesty(2) = 0
    gMsk.CrdDestX(3) = 106
    gMsk.CrdDesty(3) = 0
    gMsk.CrdDestX(4) = 159
    gMsk.CrdDesty(4) = 0
    gMsk.CrdDestX(5) = 212
    gMsk.CrdDesty(5) = 0
    
    'Little card container. The width of the back map is 944 pixels.
    gMsk.LittleCardLeft = 5 '24
    gMsk.LittleCardTop = 20
    gMsk.LittleCardWidth = 930 '939
    gMsk.LittleCardHeight = 41
    gMsk.LittleCrdSngWidth = 30
    gMsk.LittleCrdSngHeight = 41
    gMsk.LittleCardPadding = 1

    'Dice container.
    gMsk.DiceLeft = 340
    gMsk.DiceTop = 62
    gMsk.DiceWidth = 221
    gMsk.DiceHeight = 109
    
    'Single die size.
    gMsk.DieWidth = 39
    gMsk.DieHeight = 39
    
    'Dice destinations within dice containers.
    'The first array element is the number thrown. The second
    'array element is the position of each dice.
    '1 die thrown.
    gMsk.DieAttackDestX(1, 1) = 90
    gMsk.DieAttackDestY(1, 1) = 0
    
    gMsk.DieDefendDestX(1, 1) = 90
    gMsk.DieDefendDestY(1, 1) = 66
    
    '2 dice thrown.
    gMsk.DieAttackDestX(2, 1) = 66
    gMsk.DieAttackDestY(2, 1) = 0
    gMsk.DieAttackDestX(2, 2) = 110
    gMsk.DieAttackDestY(2, 2) = 0
    
    gMsk.DieDefendDestX(2, 1) = 66
    gMsk.DieDefendDestY(2, 1) = 66
    gMsk.DieDefendDestX(2, 2) = 110
    gMsk.DieDefendDestY(2, 2) = 66
    
    '3 dice thrown.
    gMsk.DieAttackDestX(3, 1) = 44
    gMsk.DieAttackDestY(3, 1) = 12
    gMsk.DieAttackDestX(3, 2) = 88
    gMsk.DieAttackDestY(3, 2) = 0
    gMsk.DieAttackDestX(3, 3) = 132
    gMsk.DieAttackDestY(3, 3) = 12
    
    gMsk.DieDefendDestX(3, 1) = 44
    gMsk.DieDefendDestY(3, 1) = 54
    gMsk.DieDefendDestX(3, 2) = 88
    gMsk.DieDefendDestY(3, 2) = 66
    gMsk.DieDefendDestX(3, 3) = 132
    gMsk.DieDefendDestY(3, 3) = 54
    
    '4 dice thrown.
    gMsk.DieAttackDestX(4, 1) = 22
    gMsk.DieAttackDestY(4, 1) = 12
    gMsk.DieAttackDestX(4, 2) = 66
    gMsk.DieAttackDestY(4, 2) = 0
    gMsk.DieAttackDestX(4, 3) = 110
    gMsk.DieAttackDestY(4, 3) = 0
    gMsk.DieAttackDestX(4, 4) = 154
    gMsk.DieAttackDestY(4, 4) = 12
    
    gMsk.DieDefendDestX(4, 1) = 22
    gMsk.DieDefendDestY(4, 1) = 54
    gMsk.DieDefendDestX(4, 2) = 66
    gMsk.DieDefendDestY(4, 2) = 66
    gMsk.DieDefendDestX(4, 3) = 110
    gMsk.DieDefendDestY(4, 3) = 66
    gMsk.DieDefendDestX(4, 4) = 154
    gMsk.DieDefendDestY(4, 4) = 54
    
    '5 dice thrown
    gMsk.DieAttackDestX(5, 1) = 0
    gMsk.DieAttackDestY(5, 1) = 12
    gMsk.DieAttackDestX(5, 2) = 44
    gMsk.DieAttackDestY(5, 2) = 6
    gMsk.DieAttackDestX(5, 3) = 88
    gMsk.DieAttackDestY(5, 3) = 0
    gMsk.DieAttackDestX(5, 4) = 132
    gMsk.DieAttackDestY(5, 4) = 6
    gMsk.DieAttackDestX(5, 5) = 176
    gMsk.DieAttackDestY(5, 5) = 12
    
    gMsk.DieDefendDestX(5, 1) = 0
    gMsk.DieDefendDestY(5, 1) = 54
    gMsk.DieDefendDestX(5, 2) = 44
    gMsk.DieDefendDestY(5, 2) = 60
    gMsk.DieDefendDestX(5, 3) = 88
    gMsk.DieDefendDestY(5, 3) = 66
    gMsk.DieDefendDestX(5, 4) = 132
    gMsk.DieDefendDestY(5, 4) = 60
    gMsk.DieDefendDestX(5, 5) = 176
    gMsk.DieDefendDestY(5, 5) = 54
End Sub

Private Sub Command2_Click()
    Map1.Refresh
End Sub
