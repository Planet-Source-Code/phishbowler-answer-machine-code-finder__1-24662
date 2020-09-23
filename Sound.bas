Attribute VB_Name = "Sound"
'Sound BAS Assembled by Phishbowler

'Wav
  Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

    Const SND_SYNC = &H0
   Const SND_ASYNC = &H1
   Const SND_NODEFAULT = &H2
   Const SND_LOOP = &H8
   Const SND_NOSTOP = &H10
   Public Const SND_MEMORY = &H4



Public Sub Playwav(FilePath As String, PickOption As Integer)

Select Case PickOption%
Case 1:
wFlags% = SND_NOSTOP
Case 2:
wFlags% = SND_LOOP
Case 3:
wFlags% = SND_NODEFAULT
Case 4:
wFlags% = SND_ASYNC
Case 5:
wFlags% = SND_SYNC
Case 6:
wFlags% = SND_SYNC And SND_NOSTOP
Case 7:
wFlags% = SND_SYNC And SND_LOOP
Case 8:
wFlags% = SND_SYNC And SND_NODEFAULT
Case 9:
wFlags% = SND_ASYNC And SND_NOSTOP And SND_LOOP
Case 10:
wFlags% = SND_ASYNC And SND_LOOP
Case 11:
wFlags% = SND_ASYNC And SND_NODEFAULT
Case 12:
wFlags% = SND_SYNC Or SND_NOSTOP
Case 13:
wFlags% = SND_SYNC Or SND_LOOP
Case 14:
wFlags% = SND_SYNC Or SND_NODEFAULT
Case 15:
wFlags% = SND_ASYNC Or SND_NOSTOP
Case 16:
wFlags% = SND_ASYNC Or SND_LOOP
Case 17:
wFlags% = SND_ASYNC Or SND_NODEFAULT
End Select
   x% = sndPlaySound(FilePath$, wFlags%)
End Sub

