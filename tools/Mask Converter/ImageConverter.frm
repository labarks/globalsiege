VERSION 5.00
Begin VB.Form ImageConverter 
   Caption         =   "Image Converter"
   ClientHeight    =   11325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   ScaleHeight     =   755
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   922
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Processed Image"
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdMap 
      Caption         =   "Convert to map"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdMask 
      Caption         =   "Convert to mask"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Copy Image From Clibpoard"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox pctDest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   5655
      Left            =   6600
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.PictureBox pctSource 
      AutoRedraw      =   -1  'True
      Height          =   5340
      Left            =   0
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   344
      TabIndex        =   0
      Top             =   480
      Width           =   5220
   End
End
Attribute VB_Name = "ImageConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hend As Long, ByVal lpHelpFile As String, ByVal WCommand As Long, ByVal dwData As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As _
    String) As Long
    
Private Type mapCoord
    ctryName As String * 30
    destX As Integer
    destY As Integer
    Width As Integer
    Height As Integer
    srcX As Integer
    srcY As Integer
    printX As Integer
    printY As Integer
    neighbour(1 To 7) As Integer
End Type

Private Type MapAndMaskType
    MapX As Long
    MapY As Long
    Width As Long
    Height As Long
    MaskX As Long
    MaskY As Long
End Type

Dim MapAndMask(1 To 42) As MapAndMaskType

Dim countryID(1 To 50) As mapCoord      'Individual country data
Dim strCurrentFile As String            'Filename for data files


'Load all data into the global variable.
Private Sub LoadMaskData()
    LoadCountryData 1, 27, 112, 84, 105, 7, 1
    LoadCountryData 2, 103, 73, 102, 107, 92, 1
    LoadCountryData 3, 191, 70, 148, 119, 195, 1
    LoadCountryData 4, 106, 165, 71, 77, 1, 107
    LoadCountryData 5, 170, 173, 85, 93, 318, 269
    LoadCountryData 6, 239, 188, 82, 76, 149, 121
    LoadCountryData 7, 120, 233, 67, 78, 232, 121
    LoadCountryData 8, 181, 235, 106, 98, 300, 121
    LoadCountryData 9, 136, 299, 105, 98, 1, 201
    LoadCountryData 10, 206, 378, 86, 44, 107, 201
    LoadCountryData 11, 203, 418, 84, 92, 107, 246
    LoadCountryData 12, 239, 399, 100, 126, 1, 300
    LoadCountryData 13, 244, 497, 43, 110, 344, 1
    LoadCountryData 14, 338, 183, 44, 36, 1, 427
    LoadCountryData 15, 371, 229, 41, 52, 1, 464
    LoadCountryData 16, 413, 169, 51, 78, 46, 427
    LoadCountryData 17, 413, 230, 70, 55, 1, 517
    LoadCountryData 18, 377, 272, 55, 60, 102, 339
    LoadCountryData 19, 424, 271, 74, 56, 194, 200
    LoadCountryData 20, 448, 169, 120, 142, 407, 1
    LoadCountryData 21, 349, 330, 95, 106, 407, 144
    LoadCountryData 22, 422, 333, 79, 54, 255, 581
    LoadCountryData 23, 436, 377, 94, 106, 102, 400
    LoadCountryData 24, 414, 411, 58, 60, 158, 339
    LoadCountryData 25, 418, 459, 77, 90, 197, 400
    LoadCountryData 26, 507, 458, 29, 62, 192, 257
    LoadCountryData 27, 473, 276, 95, 119, 222, 269
    LoadCountryData 28, 548, 253, 73, 96, 275, 389
    LoadCountryData 29, 562, 306, 81, 108, 595, 261
    LoadCountryData 30, 634, 345, 71, 81, 73, 109
    LoadCountryData 31, 638, 124, 113, 155, 528, 1
    LoadCountryData 32, 558, 83, 131, 200, 463, 251
    LoadCountryData 33, 606, 266, 134, 108, 349, 452
    LoadCountryData 34, 679, 257, 92, 67, 349, 378
    LoadCountryData 35, 706, 184, 87, 83, 503, 157
    LoadCountryData 36, 727, 126, 137, 89, 197, 491
    LoadCountryData 37, 773, 252, 56, 69, 72, 507
    LoadCountryData 38, 762, 131, 164, 150, 484, 452
    LoadCountryData 39, 679, 404, 103, 67, 1, 577
    LoadCountryData 40, 754, 473, 78, 86, 335, 561
    LoadCountryData 41, 790, 418, 86, 34, 484, 603
    LoadCountryData 42, 821, 473, 90, 103, 595, 157
End Sub

'Load individual country data.
Private Sub LoadCountryData(pID As Long, pMapX As Long, pMapY As Long, _
pWidth As Long, pHeight As Long, pMaskX As Long, pMaskY As Long)
    With MapAndMask(pID)
        .MapX = pMapX
        .MapY = pMapY
        .Width = pWidth
        .Height = pHeight
        .MaskX = pMaskX
        .MaskY = pMaskY
    End With
End Sub

    'Read Mask info from file
Private Sub ReadcountryIDInfo()
    Dim intChan As Integer
    Dim coLength As Integer
    Dim countryNumber As Integer
    Dim d1 As Integer, d2 As Integer, d3 As Long, d4 As Long
    Dim x1, x2, x3, x4
    Dim tst As Boolean
    Dim tmp69 As Long
    
    coLength = Len(countryID(1))
    intChan = FreeFile
    Open (strCurrentFile + "\Risk44INI.dat") For Random As intChan Len = coLength
    For countryNumber = 1 To 50
        Get #intChan, countryNumber, countryID(countryNumber)
        'MapColor(countryNumber) = &HFFFFFF      'Make all countriess white
    Next countryNumber
    Close intChan
End Sub



Private Sub cmdOpen_Click()
    pctSource.Visible = True
    pctDest.Visible = False
    pctSource.Picture = Clipboard.GetData()
End Sub

Private Sub cmdMask_Click()
    Dim cntr As Long
    
    pctDest.Left = 0
    pctDest.Width = 696 'pctSource.Width
    pctDest.Height = 650 'pctSource.Height
    
    For cntr = 1 To 42
        Call ConvertMapToMask(cntr)
    Next cntr
    pctSource.Visible = False
    pctDest.Visible = True
    pctDest.Refresh
End Sub

    'Convert map format to mask format
Private Sub ConvertMapToMask(pCountry As Long)
    Dim Dummy As Long
    
    'Mask OR Map1 (DSo)
    Dummy = BitBlt(pctDest.hdc, _
        MapAndMask(pCountry).MaskX, MapAndMask(pCountry).MaskY, _
        MapAndMask(pCountry).Width, MapAndMask(pCountry).Height, _
        pctSource.hdc, MapAndMask(pCountry).MapX, MapAndMask(pCountry).MapY, _
        &HEE0086)
End Sub

Private Sub cmdSave_Click()
    SavePicture pctDest.Image, App.Path + "\Output.bmp"
    Clipboard.Clear
    Clipboard.SetData pctDest.Image
End Sub

Private Sub cmdMap_Click()
    Dim cntr As Integer
    
    activate
    For cntr = 1 To 42
        colorMaskMap (cntr)
    Next cntr
    pctDest.Refresh
End Sub

    'Get info and resize
Private Sub activate()
    'strCurrentFile = App.Path
    'ReadcountryIDInfo
    
    'pctSource.AutoSize = True
    pctSource.Refresh
    'pctSource.Visible = False

    With pctDest
        .AutoSize = True
        .Top = 0
        .Left = 0
        .Height = pctSource.Height
        .Width = pctSource.Width
        .Refresh
    End With
End Sub

    'Convert mask format to map format
Private Sub colorMaskMap(ctryNumber As Integer)
    Dim Dummy As Long
    
    'Mask OR Map1 (DSo)
    Dummy = BitBlt(pctDest.hdc, _
        countryID(ctryNumber).destX, countryID(ctryNumber).destY, _
        countryID(ctryNumber).Width, countryID(ctryNumber).Height, _
        pctSource.hdc, countryID(ctryNumber).srcX, countryID(ctryNumber).srcY, _
        &HEE0086)
        
End Sub



Private Sub Form_Load()
    pctSource.AutoSize = True
    pctDest.AutoSize = True
    Call LoadMaskData
End Sub
