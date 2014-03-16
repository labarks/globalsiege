Attribute VB_Name = "modFileIO"
Option Explicit

'19/10/2010 modFileIO.
'File IO related functions.
'
'All data files created by Global Siege are located in the user's application data
'directory pointed to by the %APPDATA% environment variable. This defaults to:
'"C:\Documents and Settings\<NAME>\Application Data\Global Siege" on Windows XP
'"C:\Users\<NAME>\AppData\Roaming\Global Siege" on Windows 7
'"<App.Path>\AppData" if the environment variable does not exist.


Public Const gcWarFileExtension As String = ".mrk"
Private Const cErrorLogFileName As String = "error.<DATE>.log"
Private Const cInfoLogFileName As String = "info.<DATE>.log"
Private Const cTestLogFileName As String = "test.<DATE>.log"

'Logging level used in function LogInfo().
Public gAppLogLevel As Long

'Instance ID to differentiate the logs of GlobalSiege instances when multiple
'instances are running. This is set by the config file.
Public gInstanceID As String

'Set the application log level and compress old log files. The application log level
'sets how detailed the application logs should be. Level 0 means only log inportant
'stuff and errors. Level 9 means log everything. Users can set level 9 by typing
'"mr#LogLevel9" in the Chat Box. The default log level can be set by adding the
'registry key "AppLogLevel".
Public Sub InitialiseFileIO()
    Dim vFileName As String
    Dim vDateStamp As String
    Dim vFileList() As String
    Dim vFileParts() As String
    Dim vIndex As Long
    
    On Error Resume Next
    
    'Set the log level.
    gAppLogLevel = CLng(GetSetting(gcApplicationName, "settings", "AppLogLevel", "0"))
    
    '**TODO: Compress old log files
    vDateStamp = Format(Date, "yymmdd")
    vFileList = Split(ListFiles(GetLogDataDir, "*"), vbCrLf)
    
    'For each valid file found in the log directory.
    For vIndex = 0 To UBound(vFileList)
        vFileParts = Split(vFileList(vIndex), ",")
        If UBound(vFileParts) = 1 Then
            
            'Find unzipped log files that are not stamped with today's date.
            If LCase(Right(vFileParts(1), Len(".log"))) = ".log" _
            And LCase(Right(vFileParts(1), Len(vDateStamp & ".log"))) <> vDateStamp & ".log" Then
                Debug.Print "Compress " & vFileParts(1)
            End If
        End If
    Next
    
End Sub

'Copy a file or a list of files from pFromDir to pToDir. Return true if
'at least one file was copied and there were no errors. Use wildcards to
'copy multiple files.
'Example: Call CopyFile(App.Path, GetWarDataDir, "*.mrk")
Public Function CopyFile(pFromDir As String, pToDir As String, pFileName As String) As Boolean
    Dim vFoundFile As String
    
    On Error GoTo ErrHand
    
    vFoundFile = Dir(pFromDir & "\" & pFileName, vbNormal)
    Do While vFoundFile <> ""
        FileCopy pFromDir & "\" & vFoundFile, pToDir & "\" & vFoundFile
        CopyFile = True
        vFoundFile = Dir
    Loop
    Exit Function
    
ErrHand:
    CopyFile = False
    Exit Function
End Function

'Return the system AppData directory from the APPDATA environment variable, app.path if not found.
'If the directory does not exist, it will be created.
Public Function GetAppDataDir() As String
    On Error GoTo ErrHand
    GetAppDataDir = Environ("APPDATA") & "\" & gcApplicationName
    If GetAppDataDir = "" Then
        GetAppDataDir = App.Path & "\AppData"
    End If
    
    If Dir(GetAppDataDir, vbDirectory) = "" Then
        MkDir GetAppDataDir
    End If
    Exit Function
ErrHand:
    If Not gHeadlessMode Then
        MsgBox Err.Description, vbCritical, Err.Description
    End If
    Exit Function
End Function

'Return the path to the passed directory name within the app data dir.
'Create the directory if required.
Private Function GetThisDataDir(pDirectoryName As String) As String
    On Error GoTo ErrHand
    
    GetThisDataDir = GetAppDataDir & "\" & pDirectoryName
    
    If Dir(GetThisDataDir, vbDirectory) = "" Then
        MkDir GetThisDataDir
    End If
    Exit Function
ErrHand:
    If Not gHeadlessMode Then
        MsgBox Err.Description, vbCritical, Err.Description
    End If
    Exit Function
End Function

'Return the saved wars directory.
Public Function GetWarDataDir() As String
    GetWarDataDir = GetThisDataDir("wars")
End Function

'Return the logs directory.
Public Function GetLogDataDir() As String
    GetLogDataDir = GetThisDataDir("logs")
End Function

'Return the config directory.
Public Function GetConfigDataDir() As String
    GetConfigDataDir = GetThisDataDir("config")
End Function

'Return the app's temp directory.
Public Function GetTmpDataDir() As String
    GetTmpDataDir = GetThisDataDir("tmp")
End Function

'Return a line delimited list of files found in the passed directory matching the
'file name passed in pFileName. Use wildcards for multiple files. Strip the file
'extension from the file name if passed.
'Output format: <file path & name>,<file name><CRLF>
Public Function ListFiles(pFileDir As String, pFileName As String, Optional pFileExt As String) As String
    Dim vFoundFile As String
    Dim vBareFileName As String
    
    On Error GoTo fileError
    vFoundFile = Dir(pFileDir & "\" & pFileName, vbNormal)
    Do While vFoundFile <> ""
        vBareFileName = Trim(Left(vFoundFile, Len(vFoundFile) - Len(pFileExt)))
        If vBareFileName <> "" Then
            ListFiles = ListFiles & pFileDir & "\" & vFoundFile & "," & vBareFileName & vbCrLf
        End If
        vFoundFile = Dir
    Loop
    Exit Function
    
fileError:
    Resume Next
End Function

'Return the contents of the file in the application directory with the passed file name.
Public Function LoadProgramFile(pProgFileName As String) As String
    Dim vFileRef As Integer
    
    On Error GoTo fileErr
    
    vFileRef = FreeFile
    Open App.Path & "\" & pProgFileName For Binary As vFileRef
    LoadProgramFile = Space(LOF(vFileRef))
    Get #vFileRef, , LoadProgramFile
    Close vFileRef
    
    Exit Function
fileErr:
    LoadProgramFile = ""
    Close vFileRef
    LogError "LoadProgramFile", Err.Description
    Exit Function
End Function

'Return the contents of the file in the config directory with the passed file name.
Public Function LoadConfigFile(pConfigFileName As String) As String
    Dim vFileRef As Integer
    
    On Error GoTo fileErr
    
    vFileRef = FreeFile
    Open GetConfigDataDir & "\" & pConfigFileName For Binary As vFileRef
    LoadConfigFile = Space(LOF(vFileRef))
    Get #vFileRef, , LoadConfigFile
    Close vFileRef
    
    Exit Function
fileErr:
    LoadConfigFile = ""
    Close vFileRef
    LogError "LoadConfigFile", Err.Description
    Exit Function
End Function

'Return the contents of the file passed in pTextFileName.
Public Function ReadTextFile(pTextFileName As String) As String
    Dim vFileRef As Integer
    
    On Error GoTo fileErr
    
    vFileRef = FreeFile
    Open pTextFileName For Binary As vFileRef
    ReadTextFile = Space(LOF(vFileRef))
    Get #vFileRef, , ReadTextFile
    Close vFileRef
    
    Exit Function
fileErr:
    ReadTextFile = ""
    Close vFileRef
    LogError "ReadTextFile", Err.Description
    Exit Function
End Function

'Save the contents of pConfigData to a file named pConfigFileName in the config data directory.
Public Sub SaveConfigFile(pConfigFileName As String, pConfigData As String)
    Dim vFileRef As Integer
    
    On Error Resume Next
    
    vFileRef = FreeFile
    Open GetConfigDataDir & "\" & pConfigFileName For Output As vFileRef
    Print #vFileRef, pConfigData;
    Close vFileRef
End Sub

'Save the contents of pFileContents to a file named pConfigFileName in the config data directory.
Public Sub SaveDataFile(pDataFileName As String, pFileContents As String)
    Dim vFileRef As Integer
    
    On Error Resume Next
    
    vFileRef = FreeFile
    Open GetLogDataDir & "\" & pDataFileName For Output As vFileRef
    Print #vFileRef, pFileContents;
    Close vFileRef
End Sub

'Append the passed log entry to the passed log file. The log entry
'is le6 encrypted if pLe6 is TRUE but will be overridden if gcAppDevelopMode is false.
Private Function WriteLogEntry(pLogFile As String, _
ByVal pFunctionName As String, _
pLogLine As String, _
Optional pLe6 As Boolean = False)
    Dim vFileRef As Integer
    Dim vLogLine As String
    Dim vLogFileName As String
    
    On Error Resume Next
    
    'Format the function name with brackets if present.
    If Len(pFunctionName) > 0 Then
        pFunctionName = pFunctionName & "(): "
    End If
    
    'Time stamp the record.
    vLogLine = Format(Now, "hh:mm:ss") & " - " & pFunctionName & pLogLine
    
    'Encrypt if the app is in production.
    If pLe6 And Not gcAppDevelopMode Then
        vLogLine = gGsLeUtils.LE6(vLogLine)
    End If
    
    'Format the file name with a date stamp.
    vLogFileName = Replace(pLogFile, "<DATE>", Format(Date, "yymmdd") & gInstanceID)
    
    'Append the formatted log entry to the log file.
    vFileRef = FreeFile
    Open GetLogDataDir & "\" & vLogFileName For Append As vFileRef
    Print #vFileRef, vLogLine
    Close vFileRef
    
End Function

'Write the passed error text to the error log. Output is gernerally in
'the following format:
'"12:54:04 - FunctionName(): <Err.Number> <Err.Description>"
'pFunctionName:     The name of the function that this was called from.
'pRecord:           The description text or record to be logged. Error
'                   messages passed by pRecord should be "Error: <Message>"
Public Sub LogError(ByVal pFunctionName As String, ByVal pRecord As String, Optional pLe6 As Boolean = False)
    On Error Resume Next
    
    'Write the record to the error log file.
    Call WriteLogEntry(cErrorLogFileName, pFunctionName, pRecord, pLe6)
    
    'Also write to the info log.
    Call LogInfo(pFunctionName, pRecord, 0, pLe6)
End Sub

'Write the passed info text to the info log if the passed log level is lower than or equal to
'the gAppLogLevel global variable. The record is le6 encrypted if gcAppDevelopMode is false.
'Output format is generally "12:54:04 - FunctionName(): <Record>"
'pFunctionName:     The name of the function that this was called from.
'pRecord:           The description text or record to be logged.
'pLogLevel:         The priority of the message. 0 is highest and is always logged and 9 is
'                   the lowest. The log level depends on the global variable gAppLogLevel
'                   which defaults to 0.
Public Sub LogInfo(ByVal pFunctionName As String, _
ByVal pRecord As String, _
Optional pLogLevel As Long = 0, _
Optional pLe6 As Boolean = False)
    On Error Resume Next
    
    'If higher priority that the log level (lower number) then log the record.
    'Log anyway if in debug mode.
    If pLogLevel <= gAppLogLevel Or gcAppDevelopMode Then
        
        'Write the record to the info log file.
        Call WriteLogEntry(cInfoLogFileName, pFunctionName, pRecord, pLe6)
    End If
End Sub

'Write the passed test info text to the test log with a time stamp.
'Only wirks if gcAppDevelopMode is set to TRUE.
'** No longer used. Everything now gets logged to LogInfo()
Public Sub LogTest(ByVal pRecord As String)
    Dim vLogFileName As String
    Dim vFileRef As Integer
    
    On Error Resume Next
    
    If gcAppDevelopMode Then
        vLogFileName = Replace(cTestLogFileName, "<DATE>", Format(Date, "yymmdd"))
    
        vFileRef = FreeFile
        Open GetLogDataDir & "\" & vLogFileName For Append As vFileRef
        Print #vFileRef, Format(Now, "hh:mm:ss") & " - " & pRecord
        Close vFileRef
    End If
End Sub

'Load the war file into the passed war struct. Return true if successful.
Public Function LoadWarFile(pWarFileName As String, ByRef pWarFile As WarControlType) As Boolean
    Dim vFileRef As Integer
    Dim vFileLength As Integer
    
    On Error GoTo fileErr
    
    If Dir(pWarFileName) = "" Then
        Exit Function
    End If
    
    vFileLength = Len(pWarFile) + Len(pWarFile.sPlayerID(1)) + fileBuffer
    vFileRef = FreeFile
    Open pWarFileName For Random As vFileRef Len = vFileLength
    Get #vFileRef, 1, pWarFile
    Close vFileRef
    Call SubstituteStringTokens(pWarFile.fileDescription)
    LoadWarFile = True
    Exit Function
fileErr:
    LoadWarFile = False
    Close vFileRef
    Exit Function
End Function

'Save the passed war file. Error handling must be done by the calling function.
Public Sub SaveWarFile(pWarFileName As String, pWarFile As WarControlType)
    Dim vFileRef As Integer
    Dim vFileLength As Integer
        
    vFileLength = Len(pWarFile) + Len(pWarFile.sPlayerID(1)) + fileBuffer
    vFileRef = FreeFile
    Open pWarFileName For Random As vFileRef Len = vFileLength
    Put #vFileRef, 1, pWarFile
    Close vFileRef
End Sub

'Return true if the passed war file (path and name) has been locked.
Public Function IsWarFileLocked(pWarFileName As String) As Boolean
    Dim vWarContainer As WarControlType
    Call LoadWarFile(pWarFileName, vWarContainer)
    IsWarFileLocked = vWarContainer.Locked
End Function

'Return simple but fast checksum of passed file.
Public Function GetFileCS(pFileName As String) As Double
    Dim vContentByte() As Byte
    Dim vFlen As Long
    Dim vTot As Double
    Dim i As Long
    
    On Error Resume Next
    
    If Dir(pFileName) = "" Then
        pFileName = App.Path & "\" & gcApplicationName & ".exe"
        If Dir(pFileName) = "" Then
            GetFileCS = 0
            Exit Function
        End If
    End If

    vFlen = FileLen(pFileName)
    ReDim vContentByte(vFlen - 1)
    Open pFileName For Binary As #1
    Get #1, 1, vContentByte
    
    For i = 0 To UBound(vContentByte)
        vTot = vTot + vContentByte(i) * i
    Next
    GetFileCS = vTot
    Close #1
End Function
