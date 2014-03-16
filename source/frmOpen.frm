VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open war"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmOpen.frx":030A
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4980
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      HideSelection   =   0   'False
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open war"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4560
      Width           =   1140
   End
   Begin VB.CommandButton cmdCncl 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   4560
      Width           =   1140
   End
   Begin VB.CheckBox chkDefault 
      Caption         =   "&Make this the default war"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Open with this war next time MissionRisk starts"
      Top             =   4200
      Width           =   3480
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Delete the selected war"
      Top             =   4560
      Width           =   735
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Available Wars"
         Object.Width           =   17637
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   930
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gWarContainer As WarControlType
Dim gCurrentWarPath As String
Dim gFileOK As Boolean
Dim gSelectedWarPath As String

Property Let currentWar(pFilePathAndName As String)
    gCurrentWarPath = Trim(pFilePathAndName)
End Property

Property Get SelectedWar() As String
    SelectedWar = gSelectedWarPath
End Property

Private Sub cmdCncl_Click()
    gSelectedWarPath = ""
    Unload Me
End Sub

'Open War button clicked.
Private Sub cmdOpen_Click()
    If gFileOK Then
        gSelectedWarPath = ListView1.SelectedItem.Key
    Else
        gSelectedWarPath = ""
    End If
    
    If chkDefault.Value Then
        SaveSetting gcApplicationName, "settings", "StartingWar", ListView1.SelectedItem.Key
    End If
    Unload Me
End Sub

'Attempt to delete the war.
Private Sub cmdDelete_Click()
    On Error GoTo fileError
    If MsgBox(Phrase(192) & ListView1.SelectedItem.Text & "?", vbYesNo, Phrase(143)) Then
        Kill ListView1.SelectedItem.Key
        Call ListWarFiles
        Call HighlightLVKey(gCurrentWarPath)
    End If

    Exit Sub
fileError:
    MsgBox Phrase(261) & ListView1.SelectedItem.Text & "." & vbCrLf _
    & Phrase(262), vbCritical, Phrase(56)   'An error occured while deleting ;It might be already opened.
    Resume Next
End Sub

'Display form as modal.
Public Sub display()
    On Error Resume Next
    Show vbModal
End Sub

Private Sub Form_Load()
    ListView1.ColumnHeaders(1).Text = Phrase(142)
    ListView1.ColumnHeaders(1).Width = ListView1.Width
    cmdDelete.Caption = Phrase(143)
    cmdOpen.Caption = Phrase(200)
    cmdCncl.Caption = "&" + Phrase(63)
    Label1.Caption = Phrase(146)
    chkDefault.Caption = Phrase(145)
    frmOpen.Caption = Phrase(194)
    cmdDelete.ToolTipText = Phrase(323)  'Delete the selected war
    
    txtDescription.Text = ""
    Call ListWarFiles
End Sub

'Find war files and list in the list box ListView1.
Private Sub ListWarFiles()
    Dim vIndex As Long
    Dim vWars() As String
    Dim vWarDetails() As String
    Dim L As ListItem
    
    On Error GoTo fileError
    
    ListView1.ListItems.Clear
    vWars = Split(ListFiles(App.Path, "*" & gcWarFileExtension, gcWarFileExtension) _
            & ListFiles(GetWarDataDir, "*" & gcWarFileExtension, gcWarFileExtension), vbCrLf)
    For vIndex = 0 To UBound(vWars) - 1
        vWarDetails = Split(vWars(vIndex), ",")
        'vWarDetails(0) is the full file path and name.
        'vWarDetails(1) is the saved war name.
        Set L = ListView1.ListItems.Add(, vWarDetails(0), DecodeNonAscii(vWarDetails(1)))
        'L.Tag = vWarDetails(0)
    Next

    Exit Sub
    
fileError:
    Resume Next
End Sub

'Get war details and fill in the description box & enable/disable the delete button.
Private Sub GetWarDetails(pWarFileName As String, pShortName As String)
    Dim vWarContainer As WarControlType
    
    If LoadWarFile(pWarFileName, vWarContainer) Then
        cmdDelete.Enabled = Not vWarContainer.Locked
        txtDescription.Text = Trim(vWarContainer.fileDescription)
        gFileOK = True
    Else
        'There was a problem reading the passed war file.
        MsgBox pShortName & Phrase(48), vbCritical, Phrase(56)    ' has been corrupted and cannot be opened.; File error
        cmdDelete.Enabled = True
        txtDescription.Text = " -<Corrupted file>-"
        gFileOK = False
    End If
End Sub

Private Sub ListView1_DblClick()
    Call cmdOpen_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call GetWarDetails(Item.Key, Item.Text)
End Sub

'Hilight the current war. This must be done here because you
'can't select listview items until the form is visible.
Private Sub Form_Activate()
    Call HighlightLVKey(gCurrentWarPath)
End Sub

Private Sub HighlightLVKey(pKey As String)
    Dim vIndex As Long
    
    For vIndex = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(vIndex).Key = gCurrentWarPath Then
            
            ListView1.ListItems(vIndex).Selected = True
            Call ListView1_ItemClick(ListView1.ListItems(vIndex))
            ListView1.ListItems(vIndex).EnsureVisible
            ListView1.SetFocus
            Exit For
        End If
    Next
End Sub

