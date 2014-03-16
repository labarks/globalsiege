Attribute VB_Name = "modSound"
Option Explicit

'Sound functions.
'
'Useful resources:
'http://www.vb6.us/tutorials/playsound-api

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SND_APPLICATION = &H80 ' look for application specific association
Private Const SND_ALIAS = &H10000 ' name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000 ' name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1 ' play asynchronously
Private Const SND_FILENAME = &H20000 ' name is a file name
Private Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4 ' lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Private Const SND_NOSTOP = &H10 ' don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000 ' don't wait if the driver is busy
Private Const SND_PURGE = &H40 ' purge non-static events for task
Private Const SND_RESOURCE = &H40004 ' name is a resource name or atom
Private Const SND_SYNC = &H0 ' play synchronously (default)

'Play the passed sound file.
Public Sub PlaySoundFromFile(ByVal pFileName As String)
    On Error Resume Next
    If Not gServerMode Then
        PlaySound pFileName, 0, SND_FILENAME Or SND_ASYNC
    End If
End Sub

'Play the passed system sound.
'Available system sounds:
'".Default"
'"AppGPFault"
'"Close"
'"EmptyRecycleBin"
'"MailBeep"
'"Maximize"
'"MenuCommand"
'"MenuPopup"
'"Minimize"
'"Open"
'"RestoreDown"
'"RestoreUp"
'"SystemAsterisk"
'"SystemExclaimation"
'"SystemExit"
'"SystemHand"
'"SystemQuestion"
'"SystemStart"
'Interaction.Beep
Public Sub PlaySystemSound(ByVal pSystemSound As String)
    On Error Resume Next
    If Not gServerMode Then
        PlaySound pSystemSound, 0, SND_ALIAS Or SND_ASYNC
    End If
End Sub

