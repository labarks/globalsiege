Attribute VB_Name = "modManifest"
Option Explicit

' Kernel32/User32 APIs for Unicode Filename Support
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function MoveFileA Lib "kernel32.dll" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function MoveFileW Lib "kernel32.dll" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Private Declare Function SetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Const FILE_ATTRIBUTE_NORMAL = &H80&
Private Const FILE_BEGIN As Long = 0
Private Const FILE_CURRENT As Long = 1
Private Const RT_MANIFEST As Long = 24&

Public Const INVALID_HANDLE_VALUE = -1&

Private Type RESOURCEHEADER ' 32 bytes
  ' http://msdn.microsoft.com/en-us/library/ms648027%28VS.85%29.aspx
  DataSize As Long
  HeaderSize As Long
  Type As Long  ' when resource section is numeric vs null-terminated string
  Name As Long  ' when resource ID is numeric vs null-terminated string
  DataVersion As Long
  MemoryFlags As Integer
  LanguageId As Integer
  Version As Long
  Characteristics As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

Public Sub Main()
    
    Dim iccex As InitCommonControlsExStruct, hMod As Long
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
       .lngICC = ICC_STANDARD_CLASSES ' vb intrinsic controls (buttons, textbox, etc)
       ' if using Common Controls; add appropriate ICC_ constants for type of control you are using
       ' example if using CommonControls v5.0 Progress bar:
       ' .lngICC = ICC_STANDARD_CLASSES Or ICC_PROGRESS_CLASS
    End With
    On Error Resume Next ' error? Requires IEv3 or above
    hMod = LoadLibrary("shell32.dll")
    InitCommonControlsEx iccex
    If Err Then
        InitCommonControls ' try Win9x version
        Err.Clear
    End If
    On Error GoTo 0
    '... show your main form next (i.e., Form1.Show)
    frmMain.Show
    If hMod Then FreeLibrary hMod


'** Tip 1: Avoid using VB Frames when applying XP/Vista themes
'          In place of VB Frames, use pictureboxes instead.
'** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
'          Doing so will prevent them from being themed.

End Sub

Public Function InsertManifestToResource(Filename As String, bManifest() As Byte, LangID As Long, ImportManifest As Boolean) As String

    ' function creates/replaces/deletes/extracts manifest files from VB resource files
    ' Granted this routine may be a bit awkward, but I extracted portions of several routines
    '   from another project of mine & honestly didn't want to rewrite it from scratch.
    
    ' If Creating/Replacing
    '       bManifest() is the manifest file
    '       LangID should be the language ID of the user's PC
    '       ImportManifest is False
    '       Return value is null if no errors, else an error message
    ' If Deleting
    '       bManifest() is an array of LBound=0:UBound=-1 {created by  StrConv(vbNullString,",")}
    '       LangID is ignored
    '       ImportManifest is False
    '       Return value is null if no errors, else an error message
    ' If Extracting
    '       bManifest() is an array of LBound=0:UBound=-1 {created by  StrConv(vbNullString,",")}
    '       ImportManifest is True
    '       Return value is null if no errors, else an error message
    '       If successful...
    '           LangID will be the resource language ID, if resource exists
    '           If manifest resource exists, bManifest() will be the manifest file
    '           If manifest resource does not exists, bManifest() is not changed
    '           So, checking the UBound of bManifest() will indicate if resource existed or not
    '       Else
    '           LangID should be ignored
    '           bManifest() is not changed
    
    Dim hFileOld As Long, hFileNew As Long
    Dim rezHeader As RESOURCEHEADER, vbRezHeader As RESOURCEHEADER
    Dim bData() As Byte, lData As Long
    Dim bOk As Boolean, bUnicode As Boolean
    Dim sTempFile As String
    
    ' RESOURCEHEADER.MemoryFlags constants
    Const PURE As Long = &H20
    Const MOVEABLE As Long = &H10

    If Not ImportManifest Then
        With rezHeader
            .DataSize = UBound(bManifest) + 1&
            .HeaderSize = Len(rezHeader)
            .Type = RT_MANIFEST * &H10000 Or &HFFFF&
            .Name = &H1FFFF
            .MemoryFlags = (PURE Or MOVEABLE)
            .LanguageId = LangID
        End With
    End If
    bUnicode = IsUnicodeSystem()

    
    ' see if the res file exists.
    If DoesFileExists(Filename, bUnicode) = False Then
        If rezHeader.DataSize Then ' else removing from res file, but no resfile
            ' create a new res file & bug out if error occurs
            hFileNew = CreateTheFile(Filename, False, bUnicode)
            If hFileNew = 0& Or hFileNew = INVALID_HANDLE_VALUE Then
                InsertManifestToResource = "Failed to create temporary RES file"
                Exit Function
            End If
            ' The VB resource file header is simply a RESOURCEHEADER
            With vbRezHeader
                .HeaderSize = 32
                .Name = 65535       ' (Numeric resource name of zero)
                .Type = .Name       ' (Numeric resource name of zero)
            End With
            WriteFile hFileNew, vbRezHeader, Len(vbRezHeader), lData, ByVal 0&
            If lData = Len(vbRezHeader) Then
                ' write the manifest resource header (32 bytes)
                WriteFile hFileNew, rezHeader, rezHeader.HeaderSize, lData, ByVal 0&
                If lData = rezHeader.HeaderSize Then
                    ' write the manifest file itself
                    WriteFile hFileNew, bManifest(0), UBound(bManifest) + 1&, lData, ByVal 0&
                    bOk = (lData > UBound(bManifest))
                End If
            End If
            CloseHandle hFileNew
            If Not bOk Then
                ' if faile to write (i.e., no disk space, permissions, whatever)....
                DeleteTheFile Filename, bUnicode
                InsertManifestToResource = "Failed to completely write the RES file." & vbCrLf & "Ensure disk not protected and space is available"
            End If
        ElseIf ImportManifest Then ' error, should not happen
            InsertManifestToResource = "File is not in the expected format"
        End If
        Exit Function
    End If
    
    If MsgBox("The selected resource file should not be included with any open VB6 projects." & vbCrLf & _
        "If needed, first close the VB6 project referencing that resource file.", vbOKCancel + vbInformation) = vbCancel Then
        InsertManifestToResource = "Aborted by user"
        Exit Function
    End If
        
    ' resource file exists, we will create a new one and copy data from target to our new one
    hFileOld = CreateTheFile(Filename, True, bUnicode)
    If hFileOld = 0& Or hFileOld = INVALID_HANDLE_VALUE Then
        InsertManifestToResource = "Failed to open the existing RES file"
        Exit Function
    End If
    
    ' read the res file's header & ensure it meets expectations
    ReadFile hFileOld, vbRezHeader, Len(vbRezHeader), lData, ByVal 0&
    If lData = Len(vbRezHeader) Then
        With vbRezHeader
            If .HeaderSize = 32 And .Name = 65535 And .Type = .Name Then
                If .DataSize = 0& And .Characteristics = .DataSize And .DataVersion = .DataSize Then
                    bOk = (.LanguageId = .DataSize And .MemoryFlags = .DataSize And .Version = .DataSize)
                End If
            End If
        End With
    End If
    If Not bOk Then
        CloseHandle hFileOld
        InsertManifestToResource = "The .res file you selected is not in the expected format."
        Exit Function
    End If
    
    If Not ImportManifest Then
        ' create the temp res file
        lData = 1&
        Do
            sTempFile = Filename & ".bak(" & CStr(lData) & ")"
            If DoesFileExists(sTempFile, bUnicode) = False Then Exit Do
            lData = lData + 1&
        Loop
        hFileNew = CreateTheFile(sTempFile, False, bUnicode)
        If hFileNew = 0& Or hFileNew = INVALID_HANDLE_VALUE Then
            InsertManifestToResource = "Failed to create the temporary RES file"
            CloseHandle hFileOld
            Exit Function
        End If
        
        ' write the RES file header we just read from the source file
        WriteFile hFileNew, vbRezHeader, Len(vbRezHeader), lData, ByVal 0&
        bOk = (lData = Len(vbRezHeader))
    End If
    
    If bOk Then
        If rezHeader.DataSize > 0& Then ' else deleting/importing, not inserting
            ' write the manifest resource header (32 bytes)
            WriteFile hFileNew, rezHeader, rezHeader.HeaderSize, lData, ByVal 0&
            bOk = (lData = rezHeader.HeaderSize)
            If bOk Then
                ' write the manifest file itself
                WriteFile hFileNew, bManifest(0), UBound(bManifest) + 1&, lData, ByVal 0&
                bOk = (lData > UBound(bManifest))
                Erase bManifest()
            End If
        End If
        If bOk Then
            If GetFileSize(hFileOld, ByVal 0&) > 32& Then ' else empty resource file
                ' now transfer the source file's resource data to the new temp file,
                ' but skipping any existing manifest resource; or extract the manifest file
                ' -- If the manifest was not written above, then the net result is manifest deleted from target resource file
                ' -- If the manifest was written above, then the net result is manifest inserted/changed in resource file
                ' -- Otherwise, we are extracting a manifest from a resource file and following applies:
                '       bManifest() will be the manifest file if successful
                '       LangID will be the resource's Lanaguge ID
                Do Until ReadRezData(hFileOld, hFileNew, bManifest(), LangID, ImportManifest, bOk) = False
                    ' Loop continues until ReadRezData returns false
                    ' If returns false and bOk is also false then resource data not read properly
                    ' If returns false and bOk is true, then resource data completely read
                    ' If returns True, bOk is ignored, not finished reading the data
                    If bOk = False Then Exit Do ' failed to parse the resource file properly
                Loop
            End If
        End If
    End If
    CloseHandle hFileOld
    If hFileNew Then CloseHandle hFileNew ' not created when extracting manifest
    
    If bOk Then ' success, delete source file & rename the temp file
        If Not ImportManifest Then
            DeleteTheFile Filename, bUnicode
            If bUnicode Then
                MoveFileW StrPtr(sTempFile), StrPtr(Filename)
            Else
                MoveFileA sTempFile, Filename
            End If
        End If
    Else ' failure
        If ImportManifest Then
            InsertManifestToResource = "Failed to successfully parse the RES file"
        Else
            DeleteTheFile sTempFile, bUnicode
            InsertManifestToResource = "Failed to completely write the RES file" & vbCrLf & "Ensure disk not protected and space is available"
        End If
    End If
    
End Function

Private Function ReadRezData(hFileFrom As Long, hFileTo As Long, outData() As Byte, LCID As Long, _
                                rtnManifest As Boolean, bContinue As Boolean) As Boolean

    Dim dcbDataSize   As Long   ' size of resource data
    Dim dcbHeaderSize As Long   ' size of header record
    Dim FileOffset   As Long    ' offset of header record within file
    ' Long, Integer & String variables for reading the file
    Dim resDataL As Long, resDataI As Integer
    Dim resName As Integer
    Dim bData() As Byte

    bContinue = False
    FileOffset = SetFilePointer(hFileFrom, 0&, 0&, FILE_CURRENT) ' cache current pointer value
    ' get resource structure data size
    ReadFile hFileFrom, dcbDataSize, 4&, resDataL, ByVal 0&
    If resDataL <> 4& Then Exit Function
    ' get resource item's size
    ReadFile hFileFrom, dcbHeaderSize, 4&, resDataL, ByVal 0&
    If resDataL <> 4& Then Exit Function
    ' minimum header size is 32, but can exceed 32 when string names are used for resource names/ids
    If dcbHeaderSize < 32& Then Exit Function
    
    ' get the resource name
    ReadFile hFileFrom, resName, 2&, resDataL, ByVal 0&
    If resDataL <> 2& Then Exit Function
    If resName = &HFFFF Then   ' if -1 then numerical name
        ReadFile hFileFrom, resName, 2&, resDataL, ByVal 0& ' and next 2bytes is the name
        If resDataL <> 2& Then Exit Function
    Else
        ' we have a unicode, double null-terminated string... not what we are looking for
        Do Until resName = 0       ' count characters, unicode=2bytes per char
            ReadFile hFileFrom, resName, 2&, resDataL, ByVal 0& ' looking for double null terminator
            If resDataL <> 2& Then Exit Function
        Loop
    End If
    
    If rtnManifest = True And resName = RT_MANIFEST Then ' when importing, we want the language identifier too
        If dcbDataSize > 0& Then ' not importing null resources
            ReadFile hFileFrom, resDataI, 2&, resDataL, ByVal 0&
            If resDataL <> 2& Then Exit Function
            If resDataI = &HFFFF Then  ' numerical ID
                ReadFile hFileFrom, resDataI, 2&, resDataL, ByVal 0& ' resource name
                If resDataL <> 2& Then Exit Function
            Else
                ' we have a unicode, double null-terminated string... not what we are looking for
                Do Until resDataI = 0       ' count characters, unicode=2bytes per char
                    ReadFile hFileFrom, resDataI, 2&, resDataL, ByVal 0& ' looking for double null terminator
                    If resDataL <> 2& Then Exit Function
                Loop
            End If
            SetFilePointer hFileFrom, 6&, ByVal 0&, FILE_CURRENT
            LCID = 0&
            ReadFile hFileFrom, LCID, 2&, resDataL, ByVal 0& ' resource name
            If resDataL <> 2& Then Exit Function
        End If
    End If
    
    ' here we can stop parsing for this resource item
    bContinue = True
    ' the resource data starts on DWORD boundary. The resource structure may or may not be
    ' DWORD aligned due to string resource names/IDs.
    ' Step 1. Determine where the resource data would start, following the resource structure
    resDataL = (((FileOffset + dcbHeaderSize) + 3&) And Not 3&)
    ' Step 2. Determine where the next resource item would start; again DWORD aligned
    resDataL = (((resDataL + dcbDataSize) + 3&) And Not 3&)
    
    If resName = RT_MANIFEST Then  ' found a manifest file at FileOffset; skip transferring it to temp file
        If rtnManifest = True Then  ' extraction
            If dcbDataSize > 0& Then    ' only extract if it has data
                resDataL = (((FileOffset + dcbHeaderSize) + 3&) And Not 3&)
                ReDim outData(0 To dcbDataSize - 1)
                SetFilePointer hFileFrom, resDataL, 0&, FILE_BEGIN 'set file pointer to where resource begins
                ReadFile hFileFrom, outData(0), dcbDataSize, resDataL, ByVal 0&
                bContinue = (resDataL = dcbDataSize)
            ElseIf resDataL < GetFileSize(hFileFrom, ByVal 0&) Then ' skip and see if another exists
                SetFilePointer hFileFrom, resDataL, 0&, FILE_BEGIN
                ReadRezData = True
            End If
        ElseIf resDataL < GetFileSize(hFileFrom, ByVal 0&) Then ' skip transfering to temp file
            SetFilePointer hFileFrom, resDataL, 0&, FILE_BEGIN
            ReadRezData = True
        End If
    Else
        ' transfer the resource to the temp file, unless importing
        If rtnManifest Then
            If resDataL < GetFileSize(hFileFrom, ByVal 0&) Then
                SetFilePointer hFileFrom, resDataL, 0&, FILE_BEGIN 'set file pointer to next resource begins
                ReadRezData = True
            End If
        Else
            dcbDataSize = resDataL - FileOffset ' calc size of resource including any padding
            ReDim bData(0 To dcbDataSize - 1&)  ' size array for reading it
            SetFilePointer hFileFrom, FileOffset, 0&, FILE_BEGIN 'set file pointer to where resource begins
            ReadFile hFileFrom, bData(0), dcbDataSize, dcbDataSize, ByVal 0& ' read data & copy it
            If dcbDataSize > UBound(bData) Then
                WriteFile hFileTo, bData(0), dcbDataSize, dcbDataSize, ByVal 0&
                If (dcbDataSize > UBound(bData)) Then
                    ReadRezData = (resDataL < GetFileSize(hFileFrom, ByVal 0&))
                Else
                    bContinue = False
                End If
            Else
                bContinue = False
            End If
        End If
    End If

End Function

Public Function IsUnicodeSystem() As Boolean

    IsUnicodeSystem = (Not (IsWindowUnicode(GetDesktopWindow) = 0&))
    
End Function


Public Function DoesFileExists(Filename As String, useUnicode As Boolean) As Boolean
    ' test to see if a file exists
    If useUnicode Then
        DoesFileExists = Not (GetFileAttributesW(StrPtr(Filename)) = INVALID_HANDLE_VALUE)
    Else
        DoesFileExists = Not (GetFileAttributes(Filename) = INVALID_HANDLE_VALUE)
    End If
End Function


Public Function DeleteTheFile(Filename As String, useUnicode As Boolean) As Boolean

    ' Function uses APIs to delete files :: unicode supported

    If useUnicode Then
        If Not (SetFileAttributesW(StrPtr(Filename), FILE_ATTRIBUTE_NORMAL) = 0&) Then
            DeleteTheFile = Not (DeleteFileW(StrPtr(Filename)) = 0&)
        End If
    Else
        If Not (SetFileAttributes(Filename, FILE_ATTRIBUTE_NORMAL) = 0&) Then
            DeleteTheFile = Not (DeleteFile(Filename) = 0&)
        End If
    End If

End Function

Public Function CreateTheFile(ByVal Filename As String, bOpen As Boolean, useUnicode As Boolean) As Long

    ' Function uses APIs to read/create files with unicode support

    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const GENERIC_WRITE As Long = &H40000000
    Const FILE_SHARE_WRITE As Long = &H2
    Const CREATE_ALWAYS As Long = 2
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    
    Dim Flags As Long, Access As Long
    Dim Disposition As Long, Share As Long
    
    If bOpen Then
        Access = GENERIC_READ
        Share = FILE_SHARE_READ
        Disposition = OPEN_EXISTING
        Flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
                Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    Else
        Access = GENERIC_READ Or GENERIC_WRITE
        Share = 0&
        If useUnicode Then
            Flags = GetFileAttributesW(StrPtr(Filename))
        Else
            Flags = GetFileAttributes(Filename)
        End If
        If Flags < 0& Then Flags = FILE_ATTRIBUTE_NORMAL
        ' CREATE_ALWAYS will delete previous file if necessary
        Disposition = CREATE_ALWAYS
    End If
    
    If useUnicode Then
        CreateTheFile = CreateFileW(StrPtr(Filename), Access, Share, ByVal 0&, Disposition, Flags, 0&)
    Else
        CreateTheFile = CreateFile(Filename, Access, Share, ByVal 0&, Disposition, Flags, 0&)
    End If

End Function



