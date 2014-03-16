VERSION 5.00
Begin VB.Form frmAspFilter 
   Caption         =   "PHP Filter"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   Icon            =   "frmAspFilter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   6975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   9855
   End
End
Attribute VB_Name = "frmAspFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Used to remove supurflous formatting for PHP files.
'02/12/2011     Converted to a PHP filter.

Private Sub Command1_Click()
    Text1.Text = ConvertASP(Text1.Text)
End Sub


Private Sub Form_Load()
    Dim vCmd As String
    Dim vOutDir As String
    Dim vFP As Integer
    Dim i As Long
    Dim vAspText As String
    Dim vLine As String
    Dim vFileName As String
    
    On Error GoTo ErrHand
    
    vCmd = Replace(Command(), Chr(34), "")
    vFileName = Mid(vCmd, InStrRev(vCmd, "\") + 1)
    
    If vCmd = "" Then
        'Open control panel.
        
    Else
        'Convert and send to upload directory.

        'Get input directory.
        If Dir(vCmd) = "" Then
            MsgBox "File not found: " & vCmd, vbCritical
            Exit Sub
        End If
        
        'Get output directory.
        vOutDir = GetSetting("ASP_Filter", "Settings", "OutDir", "")
        vOutDir = "C:\Documents and Settings\craig\Desktop\Upload Next"

        If vOutDir = "" Then
            MsgBox "Output directory not set.", vbOKOnly
            Exit Sub
        End If
        
        If Dir(vOutDir, vbDirectory) = "" Then
            MsgBox "Output directory not found.", vbOKOnly
            Exit Sub
        End If

        vFP = FreeFile
        Open vCmd For Input As vFP

        Do While Not EOF(vFP)
            Line Input #vFP, vLine
            vAspText = vAspText & vLine & vbCrLf
        Loop
        
        Close vFP
        
        'Text1.Text = vAspText
        vAspText = ConvertASP(vAspText)
        
        Text1.Text = vAspText
        
        If Dir(vOutDir & "\" & vFileName) <> "" Then
            Kill vOutDir & "\" & vFileName
        End If
        
        vFP = FreeFile
        Open vOutDir & "\" & vFileName For Output As vFP
        Print #vFP, vAspText
        Close vFP
    End If
    
    Exit Sub
ErrHand:
    MsgBox "Error: " & Err.Description, vbCritical
    End
End Sub

Private Function ConvertASP(pInstring As String) As String
    Dim vText() As String
    Dim vTmp As String
    Dim vResult As String
    Dim i As Long
    
    vText = Split(pInstring, vbCrLf)
    Text1.Text = ""
    vResult = ""
    
    'For each line.
    For i = 0 To UBound(vText)
        vTmp = Trim(vText(i))
        
        'Remove tabs.
        vTmp = Replace(vTmp, vbTab, "")
        
        'Replace double spaces.
        Do While InStr(vTmp, "  ")
            vTmp = Replace(vTmp, "  ", " ")
        Loop
        
        'Trim spaces.
        vTmp = Trim(vTmp)
        
        'Remove comments, '#' comments only.
        If Mid(vTmp, 1, 1) = "#" Then
            vTmp = ""
        End If
        
        'Remove blank lines.
        If Len(vTmp) > 0 Then
            vResult = vResult & vTmp & vbCrLf
        End If
        
    Next
    
    'Remove carriage returns arounf braces.
    vResult = Replace(vResult, "}" & vbCrLf, "}")
    vResult = Replace(vResult, "{" & vbCrLf, "{")
    vResult = Replace(vResult, vbCrLf & "}", "}")
    vResult = Replace(vResult, vbCrLf & "{", "{")
    
    'Replace double spaces if any have appeared.
    Do While InStr(vResult, "  ")
        vResult = Replace(vResult, "  ", " ")
    Loop
        
    ConvertASP = vResult
End Function


