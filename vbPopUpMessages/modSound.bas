Attribute VB_Name = "modSound"
Option Explicit

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000

Public Sub PlayWavFile(ByVal pstrFilename As String)
    If Not (Len(Dir(pstrFilename)) = 0) Then
        PlaySound pstrFilename, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
End Sub

