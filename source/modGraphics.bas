Attribute VB_Name = "modGraphics"
Option Explicit

'The main map is a bitmap image on a picture box named “Map1” contained in a hidden form
'named “Mask4”. All drawing is done directly on Map1 which is then copied to the foreground
'image named “Picture1” on the main form “Riskform1”. The image is resized as required
'during the copy to fit the current size of Picture1.
'
'There are a number of mask images also contained on Mask4 which are used to define the
'shape of individual countries. Using a Ternary Raster Operation, the mask acts as a stencil
'allowing countries to be of intricate shapes.

'When printing the army names across the top above the little cards, if
'gPrintLittleCardColors is set to false, all names are white except for
'the current player's name. If set to true, all names are printed in the
'associated army's color.
Public Const gPrintLittleCardColors As Boolean = False

'Various fonts to try during SetWinMessageFont().
Public Const gcDrawWinFonts As String = "Brush Script MT,Segoe Print,Comic Sans MS,Old English Text MT,Times New Roman,Comic Sans MS,Gabriola"

'Restores Window if Minimized
Public Const SW_SHOWNORMAL = 1

'StretchBlt API Constants.
Public Const STRETCH_ANDSCANS = &H1
Public Const STRETCH_ORSCANS = &H2
Public Const STRETCH_DELETESCANS = &H3
Public Const STRETCH_HALFTONE = &H4

'Windows style bits used for full screen mode.
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SYSMENU = &H80000
Private Const WS_CAPTION = &HC00000

'Used to get window style bits.
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

'Force total Redraw that shows new styles.
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
    
'StretchBlt - http://msdn.microsoft.com/en-us/library/dd145120(VS.85).aspx
'Raster Operations - http://msdn.microsoft.com/en-us/library/aa932106.aspx
'SetStretchBltMode - http://msdn.microsoft.com/en-us/library/dd145089(VS.85).aspx
'SetBrushOrgEx - http://msdn.microsoft.com/en-us/library/dd162967(VS.85).aspx
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, ByRef lpPt As Any) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Constants sourced from Mask4.
'Used to define positions and sizes of dice,
'cards, etc. Use integers because many of the functions
'use integers as parameters and no conversion is
'needed to convert to a long.
Public Type MaskPositionType
    CrdMainLeft         As Integer
    CrdMainTop          As Integer
    CrdMainWidth        As Integer
    CrdMainHeight       As Integer
    
    CrdVultLeft         As Integer
    CrdVultTop          As Integer
    CrdVultWidth        As Integer
    CrdVultHeight       As Integer
    
    CrdSnglHeight       As Integer
    CrdSnglWidth        As Integer
    CrdSnglSrcBuffer    As Integer
    
    CrdDestX(6)         As Integer
    CrdDesty(6)         As Integer
    
    LittleCardLeft      As Integer
    LittleCardTop       As Integer
    LittleCardWidth     As Integer
    LittleCardHeight    As Integer
    LittleCrdSngWidth   As Integer
    LittleCrdSngHeight  As Integer
    LittleCardPadding   As Integer
    
    DiceLeft            As Long
    DiceTop             As Long
    DiceWidth           As Long
    DiceHeight          As Long
    DieWidth            As Long
    DieHeight           As Long
    
    DieAttackDestX(5, 5) As Integer
    DieDefendDestX(5, 5) As Integer
    DieAttackDestY(5, 5) As Integer
    DieDefendDestY(5, 5) As Integer
End Type

Public gMsk As MaskPositionType

'DoBlt() is a wrapper for the block copy/transformation APIs. It selects and uses the
'appropriate block copy (blt) method (BitBlt or StretchBlt) depending on the capability
'of the operating system. Anything less than Windows 2000 cannot handle StretchBlt
'with full functionality. Return the results from StretchBlt() or BitBlt() API calls.
'
'StretchBlt - http://msdn.microsoft.com/en-us/library/dd145120(VS.85).aspx
'Raster Operations - http://msdn.microsoft.com/en-us/library/aa932106.aspx
'SetStretchBltMode - http://msdn.microsoft.com/en-us/library/dd145089(VS.85).aspx
'SetBrushOrgEx - http://msdn.microsoft.com/en-us/library/dd162967(VS.85).aspx
'vbSrcCopy = &HCC0020
Public Function DoBlt(pDestHDC As Long, pDestX As Long, pDestY As Long, pDestWidth As Long, pDestHeight As Long, _
pSrcHDC As Long, pSourceX As Long, pSourceY As Long, Optional pSourceWidth As Long, Optional pSourceHeight As Long, _
Optional pRasterOp As Long = vbSrcCopy) As Long
    Dim vReturn As Long
    
    On Error Resume Next
    
    'Only display graphics if not in headles mode.
    If Not gHeadlessMode Then
    
        'Set optional source width and height if required.
        If pSourceWidth = 0 Then
            pSourceWidth = pDestWidth
        End If
        If pSourceHeight = 0 Then
            pSourceHeight = pDestHeight
        End If
        
        'Set the SetStretchBltMode to user preference.
        If TheMainForm.mnuViewQualityDisplay.Checked Then
            vReturn = SetStretchBltMode(pDestHDC, STRETCH_HALFTONE)
        Else
            vReturn = SetStretchBltMode(pDestHDC, STRETCH_DELETESCANS)  'STRETCH_ANDSCANS 'STRETCH_DELETESCANS
        End If
        
        'Realign the brush as advised by Microsofr.
        vReturn = SetBrushOrgEx(pDestHDC, 0, 0, ByVal 0)
        
        'StretchBlt the back image from Mask4 to the fore image on TheMainForm.
        vReturn = StretchBlt(pDestHDC, pDestX, pDestY, pDestWidth, pDestHeight, _
                pSrcHDC, pSourceX, pSourceY, pSourceWidth, pSourceHeight, _
                pRasterOp)
        
        'If StretchBlt failed, BitBlt is and restrict the size of the fore picture.
        'by marking its tag with a non zero length comment.
        If vReturn = 0 Then
            If TheMainForm.Picture1.Tag = "" Then
                'TODO: fix the nex two lines.
                TheMainForm.Picture1.Tag = "No resize"
                Call TheMainForm.ResizeForm
            End If
            vReturn = BitBlt(pDestHDC, pDestX, pDestY, pSourceWidth, pSourceHeight, _
                    pSrcHDC, pSourceX, pSourceY, _
                    pRasterOp)
        End If
        
        DoBlt = vReturn
    End If
End Function

'Turn the window's title bar on or off.
Public Function FlipWindowsTitleBar(WindowHDC As Long, ByVal Value As Boolean) As Boolean
   Dim lStyle As Long

   lStyle = GetWindowLong(WindowHDC, GWL_STYLE)
   If Value Then
      lStyle = lStyle Or WS_CAPTION
   Else
      lStyle = lStyle And Not WS_CAPTION
   End If
   Call SetWindowLong(WindowHDC, GWL_STYLE, lStyle)
   Call RedrawWindow(WindowHDC)
End Function

' Redraw window with new style.
Private Sub RedrawWindow(WindowHDC As Long)
   Const swpFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE _
                            Or SWP_NOZORDER Or SWP_NOSIZE
   Call SetWindowPos(WindowHDC, 0, 0, 0, 0, 0, swpFlags)
End Sub

