VERSION 5.00
Begin VB.Form frmManifest 
   Caption         =   "Manifest File"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6780
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtManifest 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmManifest.frx":0000
      Top             =   135
      Width           =   5385
   End
End
Attribute VB_Name = "frmManifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        txtManifest.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End If
End Sub
