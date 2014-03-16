VERSION 5.00
Begin VB.Form frmIntelligence 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2940
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmIntelligence.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmIntelligence.frx":000C
   PaletteMode     =   2  'Custom
   ScaleHeight     =   196
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3480
      Begin VB.PictureBox pctIntelligenceLabelsnButtons 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   50
         ScaleHeight     =   2535
         ScaleWidth      =   3375
         TabIndex        =   1
         Top             =   120
         Width           =   3375
         Begin VB.PictureBox pctIntellegence 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000012&
            BorderStyle     =   0  'None
            Height          =   2535
            Left            =   0
            Picture         =   "frmIntelligence.frx":12BD
            ScaleHeight     =   169
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   225
            TabIndex        =   2
            Top             =   0
            Width           =   3375
            Begin VB.CommandButton OKButton 
               BackColor       =   &H00000000&
               Caption         =   "OK"
               Default         =   -1  'True
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   4
               Top             =   2160
               Width           =   1095
            End
            Begin VB.CheckBox chkShowAgain 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000016&
               Height          =   255
               Left            =   0
               MaskColor       =   &H80000016&
               TabIndex        =   3
               Top             =   2280
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label lblShowAgain 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Don't show any more."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   1
               Left            =   240
               TabIndex        =   10
               Top             =   2295
               Width           =   1590
            End
            Begin VB.Label lblCountry 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Country"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   1
               Left            =   0
               TabIndex        =   9
               Top             =   0
               Width           =   975
            End
            Begin VB.Label lblContinent 
               BackStyle       =   0  'Transparent
               Caption         =   "Belongs to North America"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   1455
               Index           =   1
               Left            =   105
               TabIndex        =   8
               Top             =   585
               Width           =   3015
            End
            Begin VB.Label lblShowAgain 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Don't show any more."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   0
               Left            =   255
               TabIndex        =   7
               Top             =   2310
               Width           =   1590
            End
            Begin VB.Label lblCountry 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Country"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   0
               Left            =   15
               TabIndex        =   6
               Top             =   15
               Width           =   975
            End
            Begin VB.Label lblContinent 
               BackStyle       =   0  'Transparent
               Caption         =   "Belongs to North America"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1455
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   600
               Width           =   3015
            End
         End
      End
   End
End
Attribute VB_Name = "frmIntelligence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Show information about individual countries (5.12.98)
Option Explicit

Private Sub Form_Activate()
    Left = TheMainForm.Left + ((TheMainForm.Width - Width) \ 2)
    Top = TheMainForm.Top + ((TheMainForm.Height - Height) \ 2)
End Sub

    'Center form
Private Sub Form_Load()
    OKButton.Caption = Phrase(340)      'OK
    lblShowAgain(0).Caption = Phrase(136)  'Don't show any more.
    lblShowAgain(1).Caption = Phrase(136)
    lblShowAgain(0).Visible = chkShowAgain.Visible
    lblShowAgain(1).Visible = chkShowAgain.Visible
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'No info box flashing.
    TheMainForm.tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
End Sub

Private Sub OKButton_Click()
    
    'No info box flashing.
    TheMainForm.tmrFlashInfoBox.Tag = CStr(gPlayerTurn)
    
    If (chkShowAgain.Visible) Then
        If chkShowAgain.Value = vbChecked Then
            'SaveSetting gcApplicationName, "settings", "seeReshuf", "off"
            TheMainForm.mnuOptReport.Checked = False
        End If
        chkShowAgain.Visible = False
        lblShowAgain(0).Visible = False
        lblShowAgain(1).Visible = False
    End If
    Me.Hide
End Sub

    'Display as modal
Public Sub display()
    On Error Resume Next
    Show vbModal
End Sub

    'Define string property for country
Property Let country(country As String)
    Dim sLen As Long
    Dim pad As String
    Dim cntr As Long
    
    pad = ""
    sLen = 17 - Len(Trim(country))
    For cntr = 1 To sLen
        pad = pad + " "
    Next cntr
    
    'MsgBox Len(Trim(Country))
    lblCountry(0).Caption = pad + country
    lblCountry(1).Caption = pad + country
    If gLanguage = eLanguage.Spanish Then    ' Spanish
        lblCountry(0).Font.Size = 14
        lblCountry(1).Font.Size = 14
    Else
        lblCountry(0).Font.Size = 16
        lblCountry(1).Font.Size = 16
    End If
End Property

Property Let Occupier(Occupier As String)
    'lblOccupier.Caption = "Occupied by " + Trim(Occupier) + "."
End Property

Property Let Continent(Continent As String)
    lblContinent(0).Caption = Continent
    lblContinent(1).Caption = Continent
End Property

