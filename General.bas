Attribute VB_Name = "General"
'General BAS - Created by Phishbowler

Public Sub pause(interval As Double)
    Dim Current
    
    Current = Timer
    Do While Timer - Current < interval
        DoEvents
    Loop
End Sub
Public Sub Timeout(interval As Integer)
    Dim Current
    
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub




Sub Copy(TextToCopy As String)
If TextToCopy$ = "" Then: Exit Sub
A$ = TextToCopy$
Clipboard.Clear
Clipboard.SetText (A$)

End Sub
Function TDate()
X = Format(Date, "mmmm/dd/yyyy")
TDate = X
End Function
Function TDate2()
X = Format(Date, "mm/dd/yy")
TDate2 = X
End Function
Function TDate3()
X = Format(Date, "mm/dd/yyyy")
TDate3 = X
End Function
Function TDate4()
X = Format(Date, "dddd/mmmm/yyyy")
TDate4 = X
End Function
Sub Shell_Write(TheExe As String)
    Dim Shellz As Long, NoFreeze As Long

    Shellz& = Shell("c:\windows\write.exe", 1): NoFreeze& = DoEvents()
End Sub
Sub Shell_Paint(TheExe As String)
    Dim Shellz As Long, NoFreeze As Long

    Shellz& = Shell("c:\windows\Pbrush.exe", 1): NoFreeze& = DoEvents()
End Sub
Sub Shell_Cdplayer(TheExe As String)
    Dim Shellz As Long, NoFreeze As Long

    Shellz& = Shell("c:\windows\Cdplayer.exe", 1): NoFreeze& = DoEvents()
End Sub
Sub Shell_Explorer(TheExe As String)
    Dim Shellz As Long, NoFreeze As Long

    Shellz& = Shell("c:\windows\Explorer.exe", 1): NoFreeze& = DoEvents()
End Sub
Sub Shell_DiskClean(TheExe As String)
' For win98 people
    Dim Shellz As Long, NoFreeze As Long

    Shellz& = Shell("c:\windows\Cleanmgr.exe", 1): NoFreeze& = DoEvents()
End Sub
Sub Shell_RegEdit(TheExe As String)
    Dim Shellz As Long, NoFreeze As Long

    Shellz& = Shell("c:\windows\regedit.exe", 1): NoFreeze& = DoEvents()
End Sub
Sub Shell_NotePad(TheExe As String)
    Dim Shellz As Long, NoFreeze As Long

    Shellz& = Shell("c:\windows\notepad.exe", 1): NoFreeze& = DoEvents()
End Sub
