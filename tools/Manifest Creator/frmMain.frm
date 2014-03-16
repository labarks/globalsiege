VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manifest Creator"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1995
      TabIndex        =   2
      Text            =   "1.0.0.0"
      Top             =   1350
      Width           =   1605
   End
   Begin VB.ComboBox cboAction 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   105
      List            =   "frmMain.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5130
      Width           =   2655
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "..."
      Height          =   330
      Left            =   3135
      TabIndex        =   5
      ToolTipText     =   "MSDN description of settings"
      Top             =   2040
      Width           =   465
   End
   Begin VB.CommandButton cmdCode 
      Caption         =   "Sub Main To Clipboard"
      Height          =   435
      Left            =   105
      TabIndex        =   9
      ToolTipText     =   "Copy typical Sub Main to initialize common controls"
      Top             =   4605
      Width           =   3525
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3780
      Width           =   3480
   End
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   960
      Width           =   3465
   End
   Begin VB.ComboBox cboUIAccess 
      Height          =   315
      ItemData        =   "frmMain.frx":007D
      Left            =   1185
      List            =   "frmMain.frx":0087
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2895
      Width           =   2415
   End
   Begin VB.ComboBox cboLevel 
      Height          =   315
      ItemData        =   "frmMain.frx":0098
      Left            =   1185
      List            =   "frmMain.frx":00A5
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2490
      Width           =   2415
   End
   Begin VB.CheckBox chkVistaSecurity 
      Caption         =   "Include Vista security in manifest"
      Height          =   300
      Left            =   135
      TabIndex        =   4
      Top             =   2055
      Width           =   2955
   End
   Begin VB.CheckBox chkCmnCtrl 
      Caption         =   "Include common controls v6.0 in manifest"
      Height          =   300
      Left            =   135
      TabIndex        =   3
      Top             =   1755
      Width           =   4005
   End
   Begin VB.CommandButton cmdResFile 
      Caption         =   "Go..."
      Height          =   435
      Left            =   2805
      TabIndex        =   11
      ToolTipText     =   "<< Perform requested action"
      Top             =   5070
      Width           =   825
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   315
      Width           =   3465
   End
   Begin VB.Label Label1 
      Caption         =   "Executable's Version"
      Height          =   300
      Index           =   6
      Left            =   165
      TabIndex        =   18
      Top             =   1410
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "Required if saving to a VB resource file"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   4140
      Width           =   3390
   End
   Begin VB.Label Label1 
      Caption         =   "Resouce File Language Identifier"
      Height          =   300
      Index           =   4
      Left            =   150
      TabIndex        =   16
      Top             =   3525
      Width           =   3315
   End
   Begin VB.Label Label1 
      Caption         =   "Executable's Description"
      Height          =   300
      Index           =   3
      Left            =   180
      TabIndex        =   15
      Top             =   720
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Executable's Name"
      Height          =   300
      Index           =   2
      Left            =   210
      TabIndex        =   14
      Top             =   90
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "UI Access"
      Height          =   300
      Index           =   1
      Left            =   150
      TabIndex        =   13
      Top             =   2955
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Requested Level"
      Height          =   435
      Index           =   0
      Left            =   135
      TabIndex        =   12
      Top             =   2460
      Width           =   990
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdCode_Click()

    Dim sCode As String
    
    Clipboard.Clear
    sCode = "Private Declare Function LoadLibraryA Lib ""kernel32.dll"" (ByVal lpLibFileName As String) As Long" & vbCrLf
    sCode = sCode & "Private Declare Function FreeLibrary Lib ""kernel32.dll"" (ByVal hLibModule As Long) As Long" & vbCrLf
    sCode = sCode & "Private Declare Function InitCommonControlsEx Lib ""comctl32.dll"" (iccex As InitCommonControlsExStruct) As Boolean" & vbCrLf
    sCode = sCode & "Private Declare Sub InitCommonControls Lib ""comctl32.dll""()" & vbCrLf
    sCode = sCode & "Private Type InitCommonControlsExStruct" & vbCrLf
    sCode = sCode & "    lngSize As Long" & vbCrLf
    sCode = sCode & "    lngICC As Long" & vbCrLf
    sCode = sCode & "End Type" & vbCrLf & vbCrLf
    sCode = sCode & "Public Sub Main()" & vbCrLf
    sCode = sCode & "    Dim iccex As InitCommonControlsExStruct, hMod As Long" & vbCrLf
    sCode = sCode & "    ' constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx" & vbCrLf
    sCode = sCode & "    Const ICC_ANIMATE_CLASS As Long = &H80&" & vbCrLf
    sCode = sCode & "    Const ICC_BAR_CLASSES As Long = &H4&" & vbCrLf
    sCode = sCode & "    Const ICC_COOL_CLASSES As Long = &H400&" & vbCrLf
    sCode = sCode & "    Const ICC_DATE_CLASSES As Long = &H100&" & vbCrLf
    sCode = sCode & "    Const ICC_HOTKEY_CLASS As Long = &H40&" & vbCrLf
    sCode = sCode & "    Const ICC_INTERNET_CLASSES As Long = &H800&" & vbCrLf
    sCode = sCode & "    Const ICC_LINK_CLASS As Long = &H8000&" & vbCrLf
    sCode = sCode & "    Const ICC_LISTVIEW_CLASSES As Long = &H1&" & vbCrLf
    sCode = sCode & "    Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&" & vbCrLf
    sCode = sCode & "    Const ICC_PAGESCROLLER_CLASS As Long = &H1000&" & vbCrLf
    sCode = sCode & "    Const ICC_PROGRESS_CLASS As Long = &H20&" & vbCrLf
    sCode = sCode & "    Const ICC_TAB_CLASSES As Long = &H8&" & vbCrLf
    sCode = sCode & "    Const ICC_TREEVIEW_CLASSES As Long = &H2&" & vbCrLf
    sCode = sCode & "    Const ICC_UPDOWN_CLASS As Long = &H10&" & vbCrLf
    sCode = sCode & "    Const ICC_USEREX_CLASSES As Long = &H200&" & vbCrLf
    sCode = sCode & "    Const ICC_STANDARD_CLASSES As Long = &H4000&" & vbCrLf
    sCode = sCode & "    Const ICC_WIN95_CLASSES As Long = &HFF&" & vbCrLf
    sCode = sCode & "    Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all values above" & vbCrLf & vbCrLf
    sCode = sCode & "    With iccex" & vbCrLf
    sCode = sCode & "       .lngSize = LenB(iccex)" & vbCrLf
    sCode = sCode & "       .lngICC = ICC_STANDARD_CLASSES ' vb intrinsic controls (buttons, textbox, etc)" & vbCrLf
    sCode = sCode & "       ' if using Common Controls; add appropriate ICC_ constants for type of control you are using" & vbCrLf
    sCode = sCode & "       ' example if using CommonControls v5.0 Progress bar:" & vbCrLf
    sCode = sCode & "       ' .lngICC = ICC_STANDARD_CLASSES Or ICC_PROGRESS_CLASS" & vbCrLf
    sCode = sCode & "    End With" & vbCrLf
    sCode = sCode & "    On Error Resume Next ' error? InitCommonControlsEx requires IEv3 or above" & vbCrLf
    sCode = sCode & "    hMod = LoadLibraryA(""shell32.dll"") ' patch to prevent XP crashes when VB usercontrols present" & vbCrLf
    sCode = sCode & "    InitCommonControlsEx iccex" & vbCrLf
    sCode = sCode & "    If Err Then " & vbCrLf
    sCode = sCode & "        InitCommonControls ' try Win9x version" & vbCrLf
    sCode = sCode & "        Err.Clear" & vbCrLf
    sCode = sCode & "    End If" & vbCrLf
    sCode = sCode & "    On Error GoTo 0" & vbCrLf
    sCode = sCode & "    '... show your main form next (i.e., Form1.Show)" & vbCrLf
    sCode = sCode & "    If hMod Then FreeLibrary hMod" & vbCrLf & vbCrLf & vbCrLf
    sCode = sCode & "'** Tip 1: Avoid using VB Frames when applying XP/Vista themes" & vbCrLf
    sCode = sCode & "'          In place of VB Frames, use pictureboxes instead." & vbCrLf
    sCode = sCode & "'** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons" & vbCrLf
    sCode = sCode & "'          Doing so will prevent them from being themed." & vbCrLf & vbCrLf
    sCode = sCode & "End Sub"
    
    Clipboard.SetText sCode
    MsgBox "Code is in the clipboard. Paste it into a module.", vbInformation + vbOKOnly
    
End Sub

Private Sub cmdHelp_Click()
    
    ShellExecute Me.hWnd, "Open", "http://msdn.microsoft.com/en-us/library/bb756929.aspx", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub cmdResFile_Click()

    ' the Go... button

    Dim cDlg As cOpenSaveDialog, bOk As Boolean
    Dim sFile As String
    
    If cboAction.ListIndex = 0 Then ' writing
        If txtExeName = vbNullString Then
            MsgBox "The Executable's Name cannot be blank", vbInformation + vbOKOnly, "No Action"
            Exit Sub
        End If
    End If
    
    If cboAction.ListIndex = 3 Then ' viewing manifest
        frmManifest.txtManifest.Text = BuildManifest
        If frmManifest.Visible = False Then
            frmManifest.Caption = frmManifest.Caption & " - Read Only"
            frmManifest.Show , Me
        End If
        Exit Sub
    End If
    
    Set cDlg = New cOpenSaveDialog
    With cDlg
        .Filter = "Manifest File|*.exe*|Resource File|*.res"
        Select Case cboAction.ListIndex
        Case 0 ' writing manifest file
            .DialogTitle = "Save Manifest As..."
            .Flags = OFN_PATHMUSTEXIST
        Case 1 ' build from vbp file
            .DialogTitle = "Select VB Project File"
            .Filter = "Project Files|*.vbp|All Files|*.*"
            .Flags = OFN_FILEMUSTEXIST
        Case 2 ' deleting manifest file
            .DialogTitle = "Select Manifest/Resource File"
            .Flags = OFN_FILEMUSTEXIST
        Case 4 ' importing manifest file
            .DialogTitle = "Select Manifest/Resource File"
            .Flags = OFN_FILEMUSTEXIST
        End Select
        .Flags = .Flags Or OFN_DONTADDTORECENT Or OFN_EXPLORER
    End With
    If cboAction.ListIndex = 0 Then
        bOk = cDlg.ShowSave(Me.hWnd)
    Else
        bOk = cDlg.ShowOpen(Me.hWnd)
    End If
    If bOk Then
        sFile = cDlg.Filename
        If cboAction.ListIndex = 1 Then ' build from vbp file
            Call UploadVBPFile(sFile)
        Else
            If cDlg.FilterIndex = 1 Then ' manifest file vs resource file
                If cboAction.ListIndex = 0 Then
                    If StrComp(Right$(sFile, 9), ".manifest", vbTextCompare) Then sFile = sFile & ".manifest"
                End If
            Else
                If StrComp(Right$(sFile, 4), ".res", vbTextCompare) Then sFile = sFile & ".res"
            End If
            Set cDlg = Nothing
            
            If cboAction.ListIndex = 0 Then     ' writing
                Call WriteManifest(sFile)
            ElseIf cboAction.ListIndex = 4 Then ' viewing external
                Call ImportManifest(sFile)
            Else                                ' deleting
                Call DeleteManifest(sFile)
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    cboLevel.ListIndex = 0
    cboUIAccess.ListIndex = 0
    cboAction.ListIndex = 1
    chkVistaSecurity.Value = vbChecked
    FillLanguageIDs
    GetLanguageID 0&
End Sub

Private Sub DeleteManifest(Filename As String)

    ' routine deletes a *.manifest file or removes embedded manifests from VB resource files
    
    Dim bData() As Byte, sReturn As String
    If StrComp(Right$(Filename, 4), ".res", vbTextCompare) Then ' manifest file vs resource file
        If MsgBox("Are you sure you want to delete the selected file?", vbYesNo + vbDefaultButton2 + vbQuestion, "Confirmation") = vbYes Then
            If DeleteTheFile(Filename, IsUnicodeSystem()) = False Then
                MsgBox "Failed to delete the file.", vbExclamation + vbOKOnly, "Error"
            Else
                MsgBox "Manifest File Deleted", vbInformation + vbOKOnly
            End If
        End If
    Else
        bData = StrConv(vbNullString, vbFromUnicode) ' create an invalid non-null array
        sReturn = InsertManifestToResource(Filename, bData, 0&, False)
        If Len(sReturn) Then
            MsgBox sReturn, vbExclamation + vbOKOnly, "Error"
        Else
            MsgBox "Manifest File Removed from Resource File", vbInformation + vbOKOnly
        End If
    End If

End Sub

Private Sub ImportManifest(Filename As String)

    ' routine reads *.manifest file or extracts 1st embedded manifest from within a VB resource file

    Dim bData() As Byte, lRtn As Long, lValue As Long, sReturn As String, bUnicode As Boolean
    If StrComp(Right$(Filename, 4), ".res", vbTextCompare) Then ' manifest file vs resource file
        bUnicode = IsUnicodeSystem()
        lRtn = CreateTheFile(Filename, True, bUnicode)
        If lRtn = INVALID_HANDLE_VALUE Or lRtn = 0& Then
            MsgBox "Failed to create the manifest file. Ensure proper permissions exist", vbExclamation + vbOKOnly, "Error"
        Else
            lValue = GetFileSize(lRtn, ByVal 0&)
            If lValue = 0& Then
                MsgBox "Invalid file", vbInformation + vbOKOnly, "Error"
            Else
                ReDim bData(0 To lValue - 1)
                ReadFile lRtn, bData(0), lValue, lValue, ByVal 0&
                If lValue > UBound(bData) Then
                    sReturn = StrConv(bData, vbUnicode)
                Else
                    MsgBox "Failed to access that file. Ensure proper permissions and access exist", vbInformation + vbOKOnly, "Error"
                End If
            End If
            CloseHandle lRtn
        End If
    Else
        bData = StrConv(vbNullString, vbFromUnicode)
        sReturn = InsertManifestToResource(Filename, bData, lValue, True)
        If Len(sReturn) Then
            MsgBox sReturn, vbExclamation + vbOKOnly, "Error"
            sReturn = vbNullString
        ElseIf UBound(bData) = -1 Then ' didn't locate a manifest file
            MsgBox "No manifest was located in the selected resource file", vbInformation + vbOKOnly
        Else
            lRtn = cboLanguage.ListIndex
            GetLanguageID lValue
            If lRtn <> cboLanguage.ListIndex Then
                sReturn = cboLanguage.Text
                cboLanguage.ListIndex = lRtn
                MsgBox "The language identifier of the resource is:" & vbCrLf & sReturn, vbInformation + vbOKOnly, "FYI"
            End If
            sReturn = StrConv(bData, vbUnicode, lValue)
        End If
    End If
    If Len(sReturn) Then
        Dim f As Form
        Set f = New frmManifest
        f.Caption = "External Manifest - Read Only"
        f.txtManifest.Text = sReturn
        f.Show , Me
    End If

End Sub

Private Sub WriteManifest(Filename As String)

    ' routine writes a manifest to a file or inserts into an existing VB resource file
    ' If inserting into resource file and file does not exist, it will be created.

    Dim sManifest As String
    Dim sBuffer As String, lBuffer As Long
    Dim blnUnicode As Boolean
    Dim hFile As Long, bData() As Byte
    
    sManifest = BuildManifest
    lBuffer = ((Len(sManifest) + 3&) And Not 3&)
    If lBuffer > Len(sManifest) Then
        sBuffer = Space$(lBuffer - Len(sManifest))
        sManifest = sManifest & sBuffer
    End If
    
    bData() = StrConv(sManifest, vbFromUnicode)
    sManifest = vbNullString
    If StrComp(Right$(Filename, 4), ".res", vbTextCompare) Then ' writing to manifest file
        blnUnicode = IsUnicodeSystem()
        If DoesFileExists(Filename, blnUnicode) = True Then
            If MsgBox("Overwrite the existing file?", vbYesNo + vbDefaultButton2 + vbQuestion, "Confirmation") = vbNo Then Exit Sub
            If DeleteTheFile(Filename, blnUnicode) = False Then
                MsgBox "Cannot overwrite the existing file. Ensure proper permission exist.", vbExclamation + vbOKOnly, "Error"
                Exit Sub
            End If
        End If
        hFile = CreateTheFile(Filename, False, blnUnicode)
        If hFile = INVALID_HANDLE_VALUE Or hFile = 0& Then
            MsgBox "Failed to create the manifest file. Ensure proper permissions exist", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
        WriteFile hFile, bData(0), UBound(bData) + 1&, lBuffer, ByVal 0&
        CloseHandle hFile
        If lBuffer > UBound(bData) Then
            MsgBox "Manifest file written to disk", vbInformation + vbOKOnly
        Else
            MsgBox "Failed to fully write the manifest file.", vbExclamation + vbOKOnly, "Error"
            DeleteTheFile Filename, blnUnicode
        End If
    Else
        sBuffer = InsertManifestToResource(Filename, bData(), cboLanguage.ItemData(cboLanguage.ListIndex), False)
        If Len(sBuffer) Then
            MsgBox sBuffer, vbExclamation + vbOKOnly, "Error"
        Else
            MsgBox "Manifest inserted into the resource file", vbInformation + vbOKOnly
        End If
    End If
    
End Sub

Private Sub UploadVBPFile(Filename As String)

    ' routine fills in the app title, description and version from a .vbp file

    Dim hFile As Long, sLines() As String, bData() As Byte
    Dim lLine As Long, iPos As Long
    Dim sVersion As String, bSubMain As Boolean, bResFile As Boolean
    
    hFile = CreateTheFile(Filename, True, IsUnicodeSystem())
    If hFile = INVALID_HANDLE_VALUE Or hFile = 0& Then
        MsgBox "Failed to open the manifest file. Ensure proper permissions exist", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    lLine = GetFileSize(hFile, ByVal 0&)
    If lLine < 1& Then
        CloseHandle hFile
        MsgBox "File is not in expected format", vbExclamation + vbOKOnly, "Error"
    Else
        ReDim bData(0 To lLine - 1)
        ReadFile hFile, bData(0), lLine, lLine, ByVal 0&
        CloseHandle hFile
        If lLine > UBound(bData) Then
            sLines() = Split(StrConv(bData, vbUnicode), vbCrLf)
            Erase bData()
            txtDescription.Text = vbNullString
            txtExeName.Text = vbNullString
            txtVersion.Text = vbNullString
            sVersion = "M.m.0.R"
            For lLine = 0 To UBound(sLines)
                iPos = InStr(sLines(lLine), "=")
                If iPos Then
                    Select Case LCase$(Left$(sLines(lLine), iPos - 1))
                    Case "name"
                        If txtExeName.Text = vbNullString Then txtExeName.Text = Replace$(Mid$(sLines(lLine), iPos + 1), Chr$(34), vbNullString)
                    Case "title": txtExeName.Text = Replace$(Mid$(sLines(lLine), iPos + 1), Chr$(34), vbNullString)
                    Case "majorver": sVersion = Replace$(sVersion, "M", Mid$(sLines(lLine), iPos + 1))
                    Case "minorver": sVersion = Replace$(sVersion, "m", Mid$(sLines(lLine), iPos + 1))
                    Case "revisionver": sVersion = Replace$(sVersion, "R", Mid$(sLines(lLine), iPos + 1))
                    Case "description": txtDescription.Text = Replace$(Mid$(sLines(lLine), iPos + 1), Chr$(34), vbNullString)
                    Case "resfile32": bResFile = True
                    Case "startup"
                        If StrComp(Mid$(sLines(lLine), iPos + 1), """Sub Main""", vbTextCompare) = 0 Then bSubMain = True
                    End Select
                End If
            Next
            Erase sLines()
            If InStr(sVersion, "M") Then sVersion = Replace$(sVersion, "M", "1")
            If InStr(sVersion, "m") Then sVersion = Replace$(sVersion, "m", "0")
            If InStr(sVersion, "R") Then sVersion = Replace$(sVersion, "R", "0")
            txtVersion.Text = sVersion
            sVersion = "The fields have been populated with the data from your VBP file." & vbCrLf & vbCrLf & "The following is also provided..." & vbCrLf
            If bResFile Then
                sVersion = sVersion & "A resource file is referenced in that project" & vbCrLf
            Else
                sVersion = sVersion & "No resource file is referenced in that project" & vbCrLf
            End If
            If bSubMain = False Then
                sVersion = sVersion & "If you will be using a manifest file, you should create a Sub Main and start your application from that"
            End If
            MsgBox sVersion, vbInformation + vbOKOnly
        Else
            MsgBox "Failed to read the manifest file. Ensure proper permissions exist", vbExclamation + vbOKOnly, "Error"
        End If
    End If
    
End Sub

Private Function BuildManifest() As String

    ' routine creates a manifest file using values from the form
    
    Dim sTemplate As String
    
        sTemplate = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
        sTemplate = sTemplate & "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf
        sTemplate = sTemplate & "  <assemblyIdentity" & vbCrLf
        sTemplate = sTemplate & "    version=""#APPVERSION#""" & vbCrLf
        sTemplate = sTemplate & "    processorArchitecture=""X86""" & vbCrLf
        sTemplate = sTemplate & "    name=""#APPTITLE#""" & vbCrLf
        sTemplate = sTemplate & "    type=""win32""" & vbCrLf
        sTemplate = sTemplate & "    />" & vbCrLf
        sTemplate = sTemplate & "  <description>#APPDESCRIPTION#</description>" & vbCrLf
    
    If chkCmnCtrl.Value = vbChecked Then
        sTemplate = sTemplate & "    <dependency>" & vbCrLf
        sTemplate = sTemplate & "        <dependentAssembly>" & vbCrLf
        sTemplate = sTemplate & "            <assemblyIdentity" & vbCrLf
        sTemplate = sTemplate & "                type=""win32""" & vbCrLf
        sTemplate = sTemplate & "                name=""Microsoft.Windows.Common-Controls""" & vbCrLf
        sTemplate = sTemplate & "                version=""6.0.0.0""" & vbCrLf
        sTemplate = sTemplate & "                processorArchitecture=""X86""" & vbCrLf
        sTemplate = sTemplate & "                publicKeyToken=""6595b64144ccf1df""" & vbCrLf
        sTemplate = sTemplate & "                language=""*""" & vbCrLf
        sTemplate = sTemplate & "             />" & vbCrLf
        sTemplate = sTemplate & "        </dependentAssembly>" & vbCrLf
        sTemplate = sTemplate & "    </dependency>" & vbCrLf
    End If
    
    If chkVistaSecurity.Value = vbChecked Then
        sTemplate = sTemplate & "<!-- Identify the application security requirements: Vista and above -->" & vbCrLf
        sTemplate = sTemplate & "  <trustInfo xmlns=""urn:schemas-microsoft-com:asm.v2"">" & vbCrLf
        sTemplate = sTemplate & "      <security>" & vbCrLf
        sTemplate = sTemplate & "        <requestedPrivileges>" & vbCrLf
        sTemplate = sTemplate & "          <requestedExecutionLevel" & vbCrLf
        sTemplate = sTemplate & "            level=""#SECURITYLEVEL#""" & vbCrLf
        sTemplate = sTemplate & "            uiAccess=""#SECURITYACCESS#""" & vbCrLf
        sTemplate = sTemplate & "            />" & vbCrLf
        sTemplate = sTemplate & "        </requestedPrivileges>" & vbCrLf
        sTemplate = sTemplate & "      </security>" & vbCrLf
        sTemplate = sTemplate & "  </trustInfo>" & vbCrLf
        sTemplate = Replace$(sTemplate, "#SECURITYLEVEL#", cboLevel.Text)
        sTemplate = Replace$(sTemplate, "#SECURITYACCESS#", LCase$(cboUIAccess.Text))
    End If

    sTemplate = Replace$(sTemplate, "#APPTITLE#", Replace$(txtExeName.Text, Chr$(34), vbNullString))
    sTemplate = Replace$(sTemplate, "#APPDESCRIPTION#", Replace$(txtDescription.Text, Chr$(34), vbNullString))
    sTemplate = Replace$(sTemplate, "#APPVERSION#", Replace$(txtVersion.Text, Chr$(34), vbNullString))
    
    BuildManifest = sTemplate & "</assembly>"

End Function

Private Sub GetLanguageID(LCID As Long)
    
    ' routine retrieves user's language ID
    
    Dim Buffer As String, X As Long
    Const LOCALE_USER_DEFAULT = &H400
    Const LOCALE_ILANGUAGE = &H1
    
    If LCID = 0& Then
        Buffer = String$(256, 0)
        LCID = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ILANGUAGE, Buffer, Len(Buffer))
        If LCID > 0& Then LCID = Val("&H" & (Left$(Buffer, LCID - 1)))
        If LCID = 0& Then LCID = 1033 ' default to English
    End If
    For X = 0 To cboLanguage.ListCount - 1
        If cboLanguage.ItemData(X) = LCID Then Exit For
    Next
    If X = cboLanguage.ListCount Then
        cboLanguage.AddItem "Unknown (" & CStr(LCID) & ")"
        X = cboLanguage.NewIndex
        cboLanguage.ItemData(X) = LCID
    End If
    cboLanguage.ListIndex = X

End Sub

Private Sub FillLanguageIDs()
    ' http://msdn.microsoft.com/en-us/goglobal/bb964664.aspx
    With cboLanguage
        .AddItem "Afrikaans (South Africa)": .ItemData(.NewIndex) = 1078
        .AddItem "Albanian (Albania)": .ItemData(.NewIndex) = 1052
        .AddItem "Amharic (Ethiopia)": .ItemData(.NewIndex) = 1118
        .AddItem "Arabic (Saudi Arabia)": .ItemData(.NewIndex) = 1025
        .AddItem "Arabic (Algeria)": .ItemData(.NewIndex) = 5121
        .AddItem "Arabic (Bahrain)": .ItemData(.NewIndex) = 15361
        .AddItem "Arabic (Egypt)": .ItemData(.NewIndex) = 3073
        .AddItem "Arabic (Iraq)": .ItemData(.NewIndex) = 2049
        .AddItem "Arabic (Jordan)": .ItemData(.NewIndex) = 11265
        .AddItem "Arabic (Kuwait)": .ItemData(.NewIndex) = 13313
        .AddItem "Arabic (Lebanon)": .ItemData(.NewIndex) = 12289
        .AddItem "Arabic (Libya)": .ItemData(.NewIndex) = 4097
        .AddItem "Arabic (Morocco)": .ItemData(.NewIndex) = 6145
        .AddItem "Arabic (Oman)": .ItemData(.NewIndex) = 8193
        .AddItem "Arabic (Qatar)": .ItemData(.NewIndex) = 16385
        .AddItem "Arabic (Syria)": .ItemData(.NewIndex) = 10241
        .AddItem "Arabic (Tunisia)": .ItemData(.NewIndex) = 7169
        .AddItem "Arabic (U.A.E.)": .ItemData(.NewIndex) = 14337
        .AddItem "Arabic (Yemen)": .ItemData(.NewIndex) = 9217
        .AddItem "Armenian (Armenia)": .ItemData(.NewIndex) = 1067
        .AddItem "Assamese": .ItemData(.NewIndex) = 1101
        .AddItem "Azeri (Cyrillic)": .ItemData(.NewIndex) = 2092
        .AddItem "Azeri (Latin)": .ItemData(.NewIndex) = 1068
        .AddItem "Basque": .ItemData(.NewIndex) = 1069
        .AddItem "Belarusian": .ItemData(.NewIndex) = 1059
        .AddItem "Bengali (India)": .ItemData(.NewIndex) = 1093
        .AddItem "Bengali (Bangladesh)": .ItemData(.NewIndex) = 2117
        .AddItem "Bosnian (Bosnia/Herzegovina)": .ItemData(.NewIndex) = 5146
        .AddItem "Bulgarian": .ItemData(.NewIndex) = 1026
        .AddItem "Burmese": .ItemData(.NewIndex) = 1109
        .AddItem "Catalan": .ItemData(.NewIndex) = 1027
        .AddItem "Cherokee (United States)": .ItemData(.NewIndex) = 1116
        .AddItem "Chinese (People's Republic of China)": .ItemData(.NewIndex) = 2052
        .AddItem "Chinese (Singapore)": .ItemData(.NewIndex) = 4100
        .AddItem "Chinese (Taiwan)": .ItemData(.NewIndex) = 1028
        .AddItem "Chinese (Hong Kong SAR)": .ItemData(.NewIndex) = 3076
        .AddItem "Chinese (Macao SAR)": .ItemData(.NewIndex) = 5124
        .AddItem "Croatian": .ItemData(.NewIndex) = 1050
        .AddItem "Croatian (Bosnia/Herzegovina)": .ItemData(.NewIndex) = 4122
        .AddItem "Czech": .ItemData(.NewIndex) = 1029
        .AddItem "Danish": .ItemData(.NewIndex) = 1030
        .AddItem "Divehi": .ItemData(.NewIndex) = 1125
        .AddItem "Dutch (Netherlands)": .ItemData(.NewIndex) = 1043
        .AddItem "Dutch (Belgium)": .ItemData(.NewIndex) = 2067
        .AddItem "Edo": .ItemData(.NewIndex) = 1126
        .AddItem "English (United States)": .ItemData(.NewIndex) = 1033
        .AddItem "English (United Kingdom)": .ItemData(.NewIndex) = 2057
        .AddItem "English (Australia)": .ItemData(.NewIndex) = 3081
        .AddItem "English (Belize)": .ItemData(.NewIndex) = 10249
        .AddItem "English (Canada)": .ItemData(.NewIndex) = 4105
        .AddItem "English (Caribbean)": .ItemData(.NewIndex) = 9225
        .AddItem "English (Hong Kong SAR)": .ItemData(.NewIndex) = 15369
        .AddItem "English (India)": .ItemData(.NewIndex) = 16393
        .AddItem "English (Indonesia)": .ItemData(.NewIndex) = 14345
        .AddItem "English (Ireland)": .ItemData(.NewIndex) = 6153
        .AddItem "English (Jamaica)": .ItemData(.NewIndex) = 8201
        .AddItem "English (Malaysia)": .ItemData(.NewIndex) = 17417
        .AddItem "English (New Zealand)": .ItemData(.NewIndex) = 5129
        .AddItem "English (Philippines)": .ItemData(.NewIndex) = 13321
        .AddItem "English (Singapore)": .ItemData(.NewIndex) = 18441
        .AddItem "English (South Africa)": .ItemData(.NewIndex) = 7177
        .AddItem "English (Trinidad)": .ItemData(.NewIndex) = 11273
        .AddItem "English (Zimbabwe)": .ItemData(.NewIndex) = 12297
        .AddItem "Estonian": .ItemData(.NewIndex) = 1061
        .AddItem "Faroese": .ItemData(.NewIndex) = 1080
        .AddItem "Farsi": .ItemData(.NewIndex) = 1065
        .AddItem "Filipino": .ItemData(.NewIndex) = 1124
        .AddItem "Finnish": .ItemData(.NewIndex) = 1035
        .AddItem "French (France)": .ItemData(.NewIndex) = 1036
        .AddItem "French (Belgium)": .ItemData(.NewIndex) = 2060
        .AddItem "French (Cameroon)": .ItemData(.NewIndex) = 11276
        .AddItem "French (Canada)": .ItemData(.NewIndex) = 3084
        .AddItem "French (Democratic Rep. of Congo)": .ItemData(.NewIndex) = 9228
        .AddItem "French (Cote d'Ivoire)": .ItemData(.NewIndex) = 12300
        .AddItem "French (Haiti)": .ItemData(.NewIndex) = 15372
        .AddItem "French (Luxembourg)": .ItemData(.NewIndex) = 5132
        .AddItem "French (Mali)": .ItemData(.NewIndex) = 13324
        .AddItem "French (Monaco)": .ItemData(.NewIndex) = 6156
        .AddItem "French (Morocco)": .ItemData(.NewIndex) = 14348
        .AddItem "French (North Africa)": .ItemData(.NewIndex) = 58380
        .AddItem "French (Reunion)": .ItemData(.NewIndex) = 8204
        .AddItem "French (Senegal)": .ItemData(.NewIndex) = 10252
        .AddItem "French (Switzerland)": .ItemData(.NewIndex) = 4108
        .AddItem "French (West Indies)": .ItemData(.NewIndex) = 7180
        .AddItem "Frisian (Netherlands)": .ItemData(.NewIndex) = 1122
        .AddItem "Fulfulde (Nigeria)": .ItemData(.NewIndex) = 1127
        .AddItem "FYRO Macedonian": .ItemData(.NewIndex) = 1071
        .AddItem "Gaelic (Ireland)": .ItemData(.NewIndex) = 2108
        .AddItem "Gaelic (Scotland)": .ItemData(.NewIndex) = 1084
        .AddItem "Galician": .ItemData(.NewIndex) = 1110
        .AddItem "Georgian": .ItemData(.NewIndex) = 1079
        .AddItem "German (Germany)": .ItemData(.NewIndex) = 1031
        .AddItem "German (Austria)": .ItemData(.NewIndex) = 3079
        .AddItem "German (Liechtenstein)": .ItemData(.NewIndex) = 5127
        .AddItem "German (Luxembourg)": .ItemData(.NewIndex) = 4103
        .AddItem "German (Switzerland)": .ItemData(.NewIndex) = 2055
        .AddItem "Greek": .ItemData(.NewIndex) = 1032
        .AddItem "Guarani (Paraguay)": .ItemData(.NewIndex) = 1140
        .AddItem "Gujarati": .ItemData(.NewIndex) = 1095
        .AddItem "Hausa (Nigeria)": .ItemData(.NewIndex) = 1128
        .AddItem "Hawaiian (United States)": .ItemData(.NewIndex) = 1141
        .AddItem "Hebrew": .ItemData(.NewIndex) = 1037
        .AddItem "HID (Human Interface Device)": .ItemData(.NewIndex) = 1279
        .AddItem "Hindi": .ItemData(.NewIndex) = 1081
        .AddItem "Hungarian": .ItemData(.NewIndex) = 1038
        .AddItem "Ibibio (Nigeria)": .ItemData(.NewIndex) = 1129
        .AddItem "Icelandic": .ItemData(.NewIndex) = 1039
        .AddItem "Igbo (Nigeria)": .ItemData(.NewIndex) = 1136
        .AddItem "Indonesian": .ItemData(.NewIndex) = 1057
        .AddItem "Inuktitut": .ItemData(.NewIndex) = 1117
        .AddItem "Italian (Italy)": .ItemData(.NewIndex) = 1040
        .AddItem "Italian (Switzerland)": .ItemData(.NewIndex) = 2064
        .AddItem "Japanese": .ItemData(.NewIndex) = 1041
        .AddItem "Kannada": .ItemData(.NewIndex) = 1099
        .AddItem "Kanuri (Nigeria)": .ItemData(.NewIndex) = 1137
        .AddItem "Kashmiri": .ItemData(.NewIndex) = 2144
        .AddItem "Kashmiri (Arabic)": .ItemData(.NewIndex) = 1120
        .AddItem "Kazakh": .ItemData(.NewIndex) = 1087
        .AddItem "Khmer": .ItemData(.NewIndex) = 1107
        .AddItem "Konkani": .ItemData(.NewIndex) = 1111
        .AddItem "Korean": .ItemData(.NewIndex) = 1042
        .AddItem "Kyrgyz (Cyrillic)": .ItemData(.NewIndex) = 1088
        .AddItem "Lao": .ItemData(.NewIndex) = 1108
        .AddItem "Latin": .ItemData(.NewIndex) = 1142
        .AddItem "Latvian": .ItemData(.NewIndex) = 1062
        .AddItem "Lithuanian": .ItemData(.NewIndex) = 1063
        .AddItem "Malay (Malaysia)": .ItemData(.NewIndex) = 1086
        .AddItem "Malay (Brunei Darussalam)": .ItemData(.NewIndex) = 2110
        .AddItem "Malayalam": .ItemData(.NewIndex) = 1100
        .AddItem "Maltese": .ItemData(.NewIndex) = 1082
        .AddItem "Manipuri": .ItemData(.NewIndex) = 1112
        .AddItem "Maori (New Zealand)": .ItemData(.NewIndex) = 1153
        .AddItem "Marathi": .ItemData(.NewIndex) = 1102
        .AddItem "Mongolian (Cyrillic)": .ItemData(.NewIndex) = 1104
        .AddItem "Mongolian (Mongolian)": .ItemData(.NewIndex) = 2128
        .AddItem "Nepali": .ItemData(.NewIndex) = 1121
        .AddItem "Nepali (India)": .ItemData(.NewIndex) = 2145
        .AddItem "Norwegian (Bokmål)": .ItemData(.NewIndex) = 1044
        .AddItem "Norwegian (Nynorsk)": .ItemData(.NewIndex) = 2068
        .AddItem "Oriya": .ItemData(.NewIndex) = 1096
        .AddItem "Oromo": .ItemData(.NewIndex) = 1138
        .AddItem "Papiamentu": .ItemData(.NewIndex) = 1145
        .AddItem "Pashto": .ItemData(.NewIndex) = 1123
        .AddItem "Polish": .ItemData(.NewIndex) = 1045
        .AddItem "Portuguese (Brazil)": .ItemData(.NewIndex) = 1046
        .AddItem "Portuguese (Portugal)": .ItemData(.NewIndex) = 2070
        .AddItem "Punjabi": .ItemData(.NewIndex) = 1094
        .AddItem "Punjabi (Pakistan)": .ItemData(.NewIndex) = 2118
        .AddItem "Quecha (Bolivia)": .ItemData(.NewIndex) = 1131
        .AddItem "Quecha (Ecuador)": .ItemData(.NewIndex) = 2155
        .AddItem "Quecha (Peru)": .ItemData(.NewIndex) = 3179
        .AddItem "Rhaeto-Romanic": .ItemData(.NewIndex) = 1047
        .AddItem "Romanian": .ItemData(.NewIndex) = 1048
        .AddItem "Romanian (Moldava)": .ItemData(.NewIndex) = 2072
        .AddItem "Russian": .ItemData(.NewIndex) = 1049
        .AddItem "Russian (Moldava)": .ItemData(.NewIndex) = 2073
        .AddItem "Sami (Lappish)": .ItemData(.NewIndex) = 1083
        .AddItem "Sanskrit": .ItemData(.NewIndex) = 1103
        .AddItem "Sepedi": .ItemData(.NewIndex) = 1132
        .AddItem "Serbian (Cyrillic)": .ItemData(.NewIndex) = 3098
        .AddItem "Serbian (Latin)": .ItemData(.NewIndex) = 2074
        .AddItem "Sindhi (India)": .ItemData(.NewIndex) = 1113
        .AddItem "Sindhi (Pakistan)": .ItemData(.NewIndex) = 2137
        .AddItem "Sinhalese (Sri Lanka)": .ItemData(.NewIndex) = 1115
        .AddItem "Slovak": .ItemData(.NewIndex) = 1051
        .AddItem "Slovenian": .ItemData(.NewIndex) = 1060
        .AddItem "Somali": .ItemData(.NewIndex) = 1143
        .AddItem "Sorbian": .ItemData(.NewIndex) = 1070
        .AddItem "Spanish (Spain (Modern Sort))": .ItemData(.NewIndex) = 3082
        .AddItem "Spanish (Spain (Traditional Sort))": .ItemData(.NewIndex) = 1034
        .AddItem "Spanish (Argentina)": .ItemData(.NewIndex) = 11274
        .AddItem "Spanish (Bolivia)": .ItemData(.NewIndex) = 16394
        .AddItem "Spanish (Chile)": .ItemData(.NewIndex) = 13322
        .AddItem "Spanish (Colombia)": .ItemData(.NewIndex) = 9226
        .AddItem "Spanish (Costa Rica)": .ItemData(.NewIndex) = 5130
        .AddItem "Spanish (Dominican Republic)": .ItemData(.NewIndex) = 7178
        .AddItem "Spanish (Ecuador)": .ItemData(.NewIndex) = 12298
        .AddItem "Spanish (El Salvador)": .ItemData(.NewIndex) = 17418
        .AddItem "Spanish (Guatemala)": .ItemData(.NewIndex) = 4106
        .AddItem "Spanish (Honduras)": .ItemData(.NewIndex) = 18442
        .AddItem "Spanish (Latin America)": .ItemData(.NewIndex) = 58378
        .AddItem "Spanish (Mexico)": .ItemData(.NewIndex) = 2058
        .AddItem "Spanish (Nicaragua)": .ItemData(.NewIndex) = 19466
        .AddItem "Spanish (Panama)": .ItemData(.NewIndex) = 6154
        .AddItem "Spanish (Paraguay)": .ItemData(.NewIndex) = 15370
        .AddItem "Spanish (Peru)": .ItemData(.NewIndex) = 10250
        .AddItem "Spanish (Puerto Rico)": .ItemData(.NewIndex) = 20490
        .AddItem "Spanish (United States)": .ItemData(.NewIndex) = 21514
        .AddItem "Spanish (Uruguay)": .ItemData(.NewIndex) = 14346
        .AddItem "Spanish (Venezuela)": .ItemData(.NewIndex) = 8202
        .AddItem "Sutu": .ItemData(.NewIndex) = 1072
        .AddItem "Swahili": .ItemData(.NewIndex) = 1089
        .AddItem "Swedish": .ItemData(.NewIndex) = 1053
        .AddItem "Swedish (Finland)": .ItemData(.NewIndex) = 2077
        .AddItem "Syriac": .ItemData(.NewIndex) = 1114
        .AddItem "Tajik": .ItemData(.NewIndex) = 1064
        .AddItem "Tamazight (Arabic)": .ItemData(.NewIndex) = 1119
        .AddItem "Tamazight (Latin)": .ItemData(.NewIndex) = 2143
        .AddItem "Tamil": .ItemData(.NewIndex) = 1097
        .AddItem "Tatar": .ItemData(.NewIndex) = 1092
        .AddItem "Telugu": .ItemData(.NewIndex) = 1098
        .AddItem "Thai": .ItemData(.NewIndex) = 1054
        .AddItem "Tibetan (Bhutan)": .ItemData(.NewIndex) = 2129
        .AddItem "Tibetan (People's Republic of China)": .ItemData(.NewIndex) = 1105
        .AddItem "Tigrigna (Eritrea)": .ItemData(.NewIndex) = 2163
        .AddItem "Tigrigna (Ethiopia)": .ItemData(.NewIndex) = 1139
        .AddItem "Tsonga": .ItemData(.NewIndex) = 1073
        .AddItem "Tswana": .ItemData(.NewIndex) = 1074
        .AddItem "Turkish": .ItemData(.NewIndex) = 1055
        .AddItem "Turkmen": .ItemData(.NewIndex) = 1090
        .AddItem "Uighur (China)": .ItemData(.NewIndex) = 1152
        .AddItem "Ukrainian": .ItemData(.NewIndex) = 1058
        .AddItem "Urdu": .ItemData(.NewIndex) = 1056
        .AddItem "Urdu (India)": .ItemData(.NewIndex) = 2080
        .AddItem "Uzbek (Cyrillic)": .ItemData(.NewIndex) = 2115
        .AddItem "Uzbek (Latin)": .ItemData(.NewIndex) = 1091
        .AddItem "Venda": .ItemData(.NewIndex) = 1075
        .AddItem "Vietnamese": .ItemData(.NewIndex) = 1066
        .AddItem "Welsh": .ItemData(.NewIndex) = 1106
        .AddItem "Xhosa": .ItemData(.NewIndex) = 1076
        .AddItem "Yi": .ItemData(.NewIndex) = 1144
        .AddItem "Yiddish": .ItemData(.NewIndex) = 1085
        .AddItem "Yoruba": .ItemData(.NewIndex) = 1130
        .AddItem "Zulu": .ItemData(.NewIndex) = 1077
    End With
End Sub

Private Sub chkVistaSecurity_Click()
    cboLevel.Enabled = (chkVistaSecurity.Value = vbChecked)
    cboUIAccess.Enabled = cboLevel.Enabled
End Sub

Private Sub txtDescription_GotFocus()
    With txtDescription
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtExeName_GotFocus()
    With txtExeName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtVersion_GotFocus()
    With txtVersion
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
