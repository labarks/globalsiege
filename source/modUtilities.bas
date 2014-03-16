Attribute VB_Name = "modUtilities"
Option Explicit

'-------------------------------------------------------------------------------
'Miscellaneous but useful functions go here. These functions cannot really
'be categorised into any of the other modules.
'-------------------------------------------------------------------------------

'Select all text in the passed text box. Best used from the GotFocus event.
Public Sub SelectAndHighlightText(pTextBox As TextBox)
    pTextBox.SelStart = 0
    pTextBox.SelLength = Len(pTextBox.Text)
End Sub


'Shuffle the elements in the passed pDelimiter delimited string.
'Blank elements will be placed at the end.
'Test: ?ShuffleString("0,1,2,3,4,5,6,7,8,9")
Public Function ShuffleString(pInString As String, Optional pDelimiter As String = ",") As String
    Dim vArray() As String
    
    'Split the sting into an array.
    vArray = Split(pInString, pDelimiter)
    
    'Shuffle the array elements.
    vArray = ShuffleStringArray(vArray)
    
    'Rejoin the shuffled string and return.
    ShuffleString = Join(vArray, pDelimiter)
End Function

'Shuffle the passed sting array.
Public Function ShuffleStringArray(pSourceArray() As String) As String()
    Dim vSourceIndex As Long
    Dim vSourceUbound As Long
    Dim vDestIndex As Long
    Dim vDestUbound As Long
    Dim vDestArray() As String
    
    'Find the size of the source array.
    vSourceUbound = UBound(pSourceArray)
    
    'Bail out here if the array is empty.
    If vSourceUbound <= 0 Then
        ShuffleStringArray = pSourceArray
        Exit Function
    End If
    
    'Create a destination array of the same size as the source array.
    vDestUbound = vSourceUbound
    ReDim vDestArray(vDestUbound) As String
    
    'Set the first output destination index at position 0.
    vDestIndex = 0
    
    'For each element in the source array.
    For vDestIndex = 0 To vDestUbound
        
        'Pick a random source element.
        vSourceIndex = CLng(GenRandom4 * vSourceUbound)
        
        'Copy it over to the current destination element.
        vDestArray(vDestIndex) = pSourceArray(vSourceIndex)
        
        'Move the last source array element down to fill the picked element's place.
        pSourceArray(vSourceIndex) = pSourceArray(vSourceUbound)
        
        'The source array is now one element smaller.
        vSourceUbound = vSourceUbound - 1
    Next
    
    ShuffleStringArray = vDestArray
End Function

'Return the number of elements in the passed array that are non zero.
Public Function CountNonzeroArrayElements(pArray() As Integer) As Integer
    Dim vIndex As Integer
    
    For vIndex = 0 To UBound(pArray)
        If pArray(vIndex) = 0 Then
            Exit For
        End If
    Next
    CountNonzeroArrayElements = vIndex
End Function

'Return true if all non zero array elements are the same.
Public Function IsArrayAllSame(pArray() As Integer) As Boolean
    Dim vIndex As Long
    
    'Check that there are at least two elements in the array.
    If CountNonzeroArrayElements(pArray) >= 2 Then
    
        IsArrayAllSame = True
        
        For vIndex = 0 To UBound(pArray) - 2
            If pArray(vIndex + 1) = 0 Then
                Exit For
            ElseIf pArray(vIndex) <> pArray(vIndex + 1) Then
                IsArrayAllSame = False
                Exit For
            End If
        Next
        
    End If
End Function

'Sort the passed array in ascending descending order determined by pOrderAscending.
Public Sub BubbleSort(pSortArray() As Integer, Optional pOrderAscending As Boolean = True)
    Dim vIterations As Long
    Dim vIndex As Long
    Dim vLastIndex As Long
    Dim vSwapDone As Boolean
    Dim vHold As Integer
    Dim vSwapNeeded As Integer
    
    'The number of itterations are needed is one less than
    'the number of elements in the sort array.
    vLastIndex = UBound(pSortArray) - 1
    
    'Keep itterating until sorted or the 2nd last element reached.
    For vIterations = 0 To vLastIndex
        vSwapDone = False
        For vIndex = 0 To vLastIndex - vIterations
            
            'Determin if elements need to be swapped.
            If pOrderAscending Then
                vSwapNeeded = pSortArray(vIndex) > pSortArray(vIndex + 1)
            Else
                vSwapNeeded = pSortArray(vIndex) < pSortArray(vIndex + 1)
            End If
            
            'Swap elements if needed.
            If vSwapNeeded Then
                vHold = pSortArray(vIndex)
                pSortArray(vIndex) = pSortArray(vIndex + 1)
                pSortArray(vIndex + 1) = vHold
                vSwapDone = True
            End If
        Next
        
        'If no elements were swapped in the last itteration, the array is sorted.
        If Not vSwapDone Then
            Exit For
        End If
    Next
End Sub

'Test the bubble sort function above.
Private Sub TestBubbleSort()
    Dim vTest(7) As Integer
    Dim vIndex As Long
    vTest(0) = 5
    vTest(1) = 8
    vTest(2) = 3
    vTest(3) = 8
    vTest(4) = 12
    vTest(5) = 4
    vTest(6) = 9
    vTest(7) = 1
    Call BubbleSort(vTest, False)
    For vIndex = 0 To UBound(vTest)
        Debug.Print vIndex, vTest(vIndex)
    Next
End Sub

'Decode escaped characters from hex to the character code.
Public Function DecodeNonAscii(pText As String, Optional pSymbol As String = "$") As String
    Dim i As Long
    Dim vParts() As String
    Dim vCode As String
    
    vParts = Split(pText, pSymbol)
    If UBound(vParts) >= 0 Then
        DecodeNonAscii = vParts(0)
        For i = 1 To UBound(vParts)
            vCode = "&H" & Mid(vParts(i), 1, 2)
            If IsNumeric(vCode) Then
                DecodeNonAscii = DecodeNonAscii & Chr(vCode) & Mid(vParts(i), 3)
            Else
                DecodeNonAscii = DecodeNonAscii & pSymbol & vParts(i)
            End If
        Next
    End If
End Function

'Encode non-ascii characters using the passed escape character and a hex ascii code.
'pExcludeString contains a list of characters that should be ignored.
Public Function EncodeNonAscii(pText As String, _
Optional pSymbol As String = "$", Optional pExcludeString As String) As String
    Dim i As Long
    Dim vParts() As String
    Dim vCode As String
    
    For i = 1 To Len(pText)
        vCode = Mid(pText, i, 1)
        
        If pExcludeString = "" Or (pExcludeString <> "" And InStr(1, pExcludeString, vCode) = 0) Then
            If vCode < "0" _
            Or (vCode > "9" And vCode < "A") _
            Or (vCode > "Z" And vCode < "a") _
            Or vCode > "z" _
            Or vCode = pSymbol Then
                vCode = Hex(Asc(vCode))
                If Len(vCode) = 1 Then
                    vCode = "0" & vCode
                End If
                vCode = pSymbol & vCode
            End If
        End If
        
        EncodeNonAscii = EncodeNonAscii & vCode
    Next
End Function

'Test the conversion functions below.
Public Function TestC()
    Dim vByte() As Byte
    Dim vHStr As String
    
    vHStr = "f0ee0001"
    Call HexStringToByteArray(vByte, vHStr)
    Debug.Print ByteArrayToHexString(vByte)
End Function

'Convert passed byte array into a hex string.
Public Function ByteArrayToHexString(Pkt() As Byte) As String
    Dim vIndex As Long
    Dim vByt As String
    
    For vIndex = 0 To UBound(Pkt)
        vByt = Hex(Pkt(vIndex))
        If Len(vByt) < 2 Then
            vByt = "0" & vByt
        End If
        ByteArrayToHexString = ByteArrayToHexString & vByt
    Next
End Function

'Convert passed hex string to a byte array.
Public Sub HexStringToByteArray(pOutByte() As Byte, pInString As String)
    Dim vIndex As Long
    Dim vByt As String
    
    On Error Resume Next
    
    If Len(pInString) >= 2 Then
        ReDim pOutByte((Len(pInString) \ 2) - 1) As Byte
    
        For vIndex = 0 To Len(pInString) \ 2 - 1
            pOutByte(vIndex) = Format("&h" & Mid(pInString, vIndex * 2 + 1, 2))
        Next
    Else
         ReDim pOutByte(0) As Byte
    End If
End Sub

'Set bit of byteData at bitPos to onOrOff value
Public Sub SetBit(OnOrOff As Boolean, bitPos As Long, byteData As Byte)
    byteData = (CByte(OnOrOff) And (2 ^ bitPos)) Or byteData
End Sub

'True if bit bitPos is set in byteData
Public Function GetBit(bitPos As Long, byteData As Byte) As Boolean
    GetBit = ((byteData) And (2 ^ bitPos)) <> 0
End Function

'Convert passed number into two byte bytecode - bigendian.
Public Sub IntToByte(InNum As Long, OutByte() As Byte)
    ReDim OutByte(1) As Byte
    OutByte(0) = CByte((InNum \ &H100) And &HFF)
    OutByte(1) = CByte(InNum And &HFF)
End Sub

'Pack int in byte() at index.
Public Sub PackIntToByte(InNum As Integer, OutByte() As Byte, Index As Integer)
    OutByte(Index) = CByte((InNum \ &H100) And &HFF)
    OutByte(Index + 1) = CByte(InNum And &HFF)
End Sub

'Convert passed byte array at Index to a number.
Public Function ByteToInt(InByte() As Byte, Index As Long) As Long
    ByteToInt = CLng(InByte(Index)) * &H100 + CLng(InByte(Index + 1))
End Function

'Convert passed number into two byte bytecode - bigendian.
Public Sub LongToByte(InNum As Long, OutByte() As Byte)
    ReDim OutByte(3) As Byte
    OutByte(0) = CByte((InNum \ &H1000000) And &HFF)
    OutByte(1) = CByte((InNum \ &H10000) And &HFF)
    OutByte(2) = CByte((InNum \ &H100) And &HFF)
    OutByte(3) = CByte(InNum And &HFF)
End Sub

'Convert passed byte array at Index to a number.
Public Function ByteToLong(InByte() As Byte, Index As Long) As Long
    ByteToLong = CLng(InByte(Index)) * &H1000000 _
               + CLng(InByte(Index + 1)) * &H10000 _
               + CLng(InByte(Index + 2)) * &H100 _
               + CLng(InByte(Index + 3))
End Function

'Return position in Array of bFind, -1 if not found.
Public Function WhereInArray(bArray() As Byte, bFind() As Byte, _
Optional StartPos As Long = 0) As Long
    Dim i As Long
    Dim j As Long
    Dim IsFound As Boolean
    
    WhereInArray = -1
    For i = StartPos To UBound(bArray) - UBound(bFind)
        IsFound = True
        For j = 0 To UBound(bFind)
            If bArray(i + j) <> bFind(j) Then
                IsFound = False
                Exit For
            End If
        Next
        If IsFound Then
            WhereInArray = i
            Exit Function
        End If
    Next
End Function

'Return TRUE if bFind is in bArray.
Public Function IsInArray(bArray() As Byte, bFind() As Byte) As Boolean
    Dim i As Long
    Dim j As Long
    Dim IsFound As Boolean
    
    IsInArray = False
    For i = 0 To UBound(bArray) - UBound(bFind)
        IsFound = True
        For j = 0 To UBound(bFind)
            If bArray(i + j) <> bFind(j) Then
                IsFound = False
                Exit For
            End If
        Next
        If IsFound Then
            IsInArray = True
            Exit Function
        End If
    Next
End Function

'Copy BytesFrom to BytesTo starting at Index.
Public Sub CopySubBytes(BytesTo() As Byte, BytesFrom() As Byte, _
Start As Long, subLength As Long)
    Dim cntr As Long
    Dim UpperBound As Long
    
    'Make sure no array overrun happens.
    UpperBound = UBound(BytesFrom)
    If Start + subLength > UpperBound Then
        subLength = UpperBound - Start
    End If
    
    ReDim BytesTo(subLength) As Byte
    
    'Start copying.
    For cntr = 0 To subLength - 1
        BytesTo(cntr) = BytesFrom(cntr + Start)
    Next
End Sub

'Copy BytesFrom to BytesTo starting at Index.
Public Sub CopyBytes(BytesTo() As Byte, BytesFrom() As Byte, Index As Long)
    Dim cntr As Long
    Dim UpperBound As Long
    
    UpperBound = UBound(BytesFrom)
    
    'Make sure destination is large enough.
    If UBound(BytesTo) < UpperBound + Index Then
        ReDim Preserve BytesTo(UpperBound + Index) As Byte
    End If
    
    'Start copying.
    For cntr = 0 To UpperBound
        BytesTo(cntr + Index) = BytesFrom(cntr)
    Next
End Sub

'Append byte array 2 to byte araray 1.
Public Sub appendByteArray(bArry1() As Byte, bArry2() As Byte)
    Dim i As Long
    Dim sz1 As Long, sz2 As Long
    Dim returnArray() As Byte
    
    sz1 = UBound(bArry1)
    sz2 = UBound(bArry2)
    
    returnArray = bArry1
    ReDim Preserve returnArray(sz1 + sz2 + 1) As Byte
    
    For i = 0 To sz2
        returnArray(i + sz1 + 1) = bArry2(i)
    Next
    ReDim bArry1(UBound(returnArray)) As Byte
    bArry1 = returnArray
End Sub

'Return a sub array starting from the passed index to the end of the passed array.
'Used to chop the first part of the array off.
'**Check, may be able to use midb().
Public Sub GetRestOfByte(pBytArray() As Byte, pStartIndex As Long)
    Dim vIndex As Long
    Dim vNewArrayLen As Long
    Dim vNewArray() As Byte
    
    vNewArrayLen = UBound(pBytArray) - pStartIndex
    ReDim vNewArray(vNewArrayLen) As Byte
    
    For vIndex = 0 To vNewArrayLen
        vNewArray(vIndex) = pBytArray(vIndex + pStartIndex)
    Next
    
    ReDim pBytArray(vNewArrayLen) As Byte
    pBytArray = vNewArray
End Sub

'Wrap text at a width of "charsPerLine"
Public Function LimitTextWidth(strText As String, charsPerLine As Long) As String
    Dim cntr As Long
    Dim StartLine As Long
    Dim LastSpace As Long
    Dim FormatstrTexting As String
    
    StartLine = 1
    LastSpace = 1
    FormatstrTexting = ""
    
    For cntr = 1 To Len(strText)
        If Len(Mid(strText, StartLine, cntr - StartLine)) >= charsPerLine Then
            If LastSpace <= StartLine Then
                FormatstrTexting = FormatstrTexting & Mid(strText, StartLine, cntr - StartLine) & vbCrLf
                If Mid(strText, cntr, 1) = " " Then
                    cntr = cntr + 1
                End If
                LastSpace = cntr
                StartLine = cntr
            Else
                FormatstrTexting = FormatstrTexting & Mid(strText, StartLine, LastSpace - StartLine) & vbCrLf
                StartLine = LastSpace + 1
                LastSpace = LastSpace + 1
            End If
            
        Else
            If Mid(strText, cntr, 1) = " " Then
                LastSpace = cntr
            End If
            If Asc(Mid(strText, cntr, 1)) = 10 Then
                FormatstrTexting = FormatstrTexting & Mid(strText, StartLine, cntr - StartLine)
                LastSpace = cntr
                StartLine = cntr
            End If
        End If
    Next
    LimitTextWidth = FormatstrTexting & Mid(strText, StartLine, Len(strText))
End Function

'Remove repeating items from the passed list. Use pDelimiter as the delimiter.
Public Function CleanList(pList As String, Optional pDelimiter As String = ",") As String
    Dim vIndexOuter As Long
    Dim vIndexInner As Long
    Dim vItems() As String
    
    If Len(pList) > 0 Then
        vItems = Split(pList, pDelimiter)
        For vIndexOuter = 0 To UBound(vItems)
            If Len(vItems(vIndexOuter)) > 0 Then
                For vIndexInner = vIndexOuter + 1 To UBound(vItems)
                    If vItems(vIndexOuter) = vItems(vIndexInner) Then
                        vItems(vIndexInner) = ""
                    End If
                Next
                CleanList = CleanList & vItems(vIndexOuter) & pDelimiter
            End If
        Next
        If Len(CleanList) > Len(pDelimiter) Then
            CleanList = Mid(CleanList, 1, Len(CleanList) - Len(pDelimiter))
        End If
    End If
End Function

'Get the list element from the passed string at the passed index.
'Use pDelimiter as the delimiter. Remember that the first element
'is 0 and the second element is 1 etc.
Public Function GetListElement(pList As String, _
pIndex As Long, _
Optional pDelimiter As String = ",") As String
    Dim vItems() As String
    
    vItems = Split(pList, pDelimiter)
    
    If UBound(vItems) >= pIndex And pIndex >= 0 Then
        GetListElement = vItems(pIndex)
    Else
        GetListElement = ""
    End If
End Function

