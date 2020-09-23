Attribute VB_Name = "API"
'This Code Was Written By: Phishbowler
'Sept. 14, 2000
'
'For Money Making Opportunities, Visit
'Http://www.dreamstruct.com/
'
'Napster Users: Tired of Incomplete Songs?
'Get the good ol' Nap v2.0 Only available at:
'Http://come.to/NapsterResume

'This BAS Constructed by: Phishbowler
'

' WinMM
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' Kernel
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' User
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function iswindowenabled Lib "user32" Alias "IsWindowEnabled" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
'Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'Shell32
Public Declare Function SHAddToRecentDocs Lib "shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long

Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
    ' ----Public Declares for this code
    Public Const RSP_SIMPLE_SERVICE = 1
    Public Const RSP_UNREGISTER_SERVICE = 0
    ' ----What makes it invisible/visible in
    '     Ctrl-alt-delete
    ' Note: That if you run this program fro
    '     m your development
    'enviorment(VB) you will not see your de
    '     velopment
    'enviorment(VB) or your programs name in
    '     the
    'Ctrl-Alt-Delete Dialog.
    'From AciD email Me at Buckwheat9@juno.c
    '     om
        


Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40



' Global & Public Const
Const EM_UNDO = &HC7

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10



Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E

Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

'Public Const flags = SWP_NOMOVE Or SWP_NOSIZE

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4

Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const ENTER = 13
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

Private Const EM_LINESCROLL = &HB6
Private Const SPI_SCREENSAVERRUNNING = 97

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   x As Long
   y As Long
End Type












'Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Declare Function ReleaseCapture Lib "user32" () As Long






Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long




'Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'   Initialization File Controls
'ReadINI

'WriteINI







Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long




Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


Public Const SB_PAGEDOWN = 3
Public Const SB_LINEDOWN = 1
Public Const VK_SCROLL = &H91

    
        
    
Public Const WM_SETTEXT = &HC
Public Const WM_LBUTTONDOWN = &H201





Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT





Public Const VK_DOWN = &H28
Public Const VK_MENU = &H12
Public Const VK_SHIFT = &H10
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102



Public Const WM_MOVE = &HF012

Public Const WM_SYSCOMMAND = &H112


Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13



'Form back color fade codes begin here
'Works best when used in the Form_Paint() sub



Public Function GetCaption(WindowHandle As Long) As String
    'From Dos
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function


Sub ClickIcon(Icon)

Call SendMessage(Icon, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(Icon, WM_LBUTTONUP, 0, 0&)
End Sub

Public Sub SetText(Window As Long, Text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, Text$)
End Sub
Sub ClickIcon2(TheButin As Long)
    Call PostMessage(TheButin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(TheButin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub SendKeysAPI(TheWin As Long, AsciiChar As Integer)
    Call sendmessagebynum(TheWin&, WM_CHAR, AsciiChar, 0&)
End Sub
Sub SendKeysAPI2(TheWin As Long)
    Call PostMessage(TheWin&, WM_KEYDOWN, ALT_MASK, 0&)
   ' Call PostMessage(TheWin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub EnableWin(Window&)
    Dim dis
    dis = EnableWindow(Window&, 1)
End Sub
Sub Win_Enable(Window&)
    Dim dis
    dis = EnableWindow(Window&, 1)
End Sub

Sub DisableWin(Window&)
    Dim dis
    dis = EnableWindow(Window&, 0)
End Sub
Sub Win_Disable(Window&)
    Dim dis
    dis = EnableWindow(Window&, 0)
End Sub

Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
room% = firs%
FindChildByClass = room%

End Function


Function FindChildByTitle(parentw, childhand)

If UCase(GetCaption(GetWindow(parentw, 5))) Like UCase(childhand) Then GoTo bone
firs = GetWindow(parentw, GW_CHILD)

While firs

If UCase(GetCaption(GetWindow(parentw, 5))) Like UCase(childhand) & "*" Then GoTo bone
firs = GetWindow(GetWindow(parentw, 5), 2)
If UCase(GetCaption(GetWindow(parentw, 5))) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
room% = firs
FindChildByTitle = room%
End Function
Function GetText(Child)
GetTrim = sendmessagebynum(Child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(Child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Function GetClass(Child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(Child, Buffer$, 250)

GetClass = Buffer$
End Function



Sub Win_OnTop(TheFrm As Form)
    
    SetOnTop = SetWindowPos(TheFrm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)


End Sub
Sub Win_Killwin(TheWind&)
    Call PostMessage(TheWind&, WM_CLOSE, 0&, 0&)
End Sub

Public Function GetWindowNextAmount(Child, CycleAmount As Integer)
For x = 1 To CycleAmount
Child = GetWindow(Child, GW_HWNDNEXT)
Next x
GetWindowNextAmount = Child
End Function

Public Function GetWindowNext(Child)
Child = GetWindow(Child, GW_HWNDNEXT)
GetWindowNext = Child
End Function

Public Sub WindowClassList(Child, List As listBox)
' Place a Listbox on the Form,
' This will Display Subsequent Children in List
' Note Item 1 is Item 0 on List
' Use this in conjunction with Function GetWindowNext,
' Place Item # in Cycle Amount

Do

Item = Item + 1
List.AddItem Item & " " & GetText(Child)
Child = GetWindow(Child, GW_HWNDNEXT)

Loop Until Child = 0
End Sub











Sub Win_Center(frmz As Form)

    frmz.Top = (Screen.Height * 0.85) / 2 - frmz.Height / 2
    frmz.Left = Screen.Width / 2 - frmz.Width / 2
End Sub


Sub Win_Hide(TheWin As Long)
    Call ShowWindow(TheWin&, SW_HIDE)
End Sub

Sub Win_Maximize(THeWindow As Long)
    Dim max As Long

    max& = ShowWindow(THeWindow&, SW_MAXIMIZE)
End Sub
Sub Win_Minimize(THeWindow As Long)
    Dim Mini As Long

    Mini& = ShowWindow(THeWindow&, SW_MINIMIZE)
End Sub


Sub Win_Restore(THeWindow As Long)
    Dim res As Long

    res& = ShowWindow(THeWindow&, SW_RESTORE)
End Sub
Sub Win_Show(TheWin As Long)
    Call ShowWindow(TheWin&, SW_SHOW)
End Sub

Sub Win_StartButtin()
    Dim WinShell As Long, StartButtin As Long, Klick As Long

    WinShell& = FindWindow("Shell_TrayWnd", "")
    StartButtin& = FindWindowEx(WinShell&, 0, "Button", vbNullString)
    Call SendMessage(StartButtin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(StartButtin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Win_Shell(TheExe As String)
    Dim Shellz As Long, NoFreeze As Long

    Shellz& = Shell(TheExe$, 1): NoFreeze& = DoEvents()
End Sub
Sub Win_Unload(TheFrm As Form)
    Unload TheFrm
    End
    End
    Unload TheFrm
End Sub


Public Function GetChildCount(ByVal hWnd As Long) As Long
Dim hChild As Long

Dim i As Integer
   
If hWnd = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hWnd, GW_CHILD)

While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
i = i + 1
Wend

GetChildCount = i
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function



Sub LISTCopy(Source, Destination)
counts = SendMessage(Source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(Source, LB_GETTEXT, Adding, Buffer$)
addstrings% = SendMessageByString(Destination, LB_ADDSTRING, 0, Buffer$)
Next Adding

End Sub
Sub WIN_NotOnTop(the As Form)
'If You Dont Want Your Text On Top Of Everything
'But Shitty Code So You Gotta Make This In A EXE To See
'How It Werkx
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub




Public Sub Hide_Program_In_Closebox()
    Dim pid As Long
    Dim reserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub


Public Sub Show_Program_In_Closebox()
    Dim pid As Long
    Dim reserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
End Sub



Function APISpy_MouseOver()
'Dim CurP As POINTAPI
'Dim NowP%
'Dim ThenP%
'    Call GetCursorPos(CurP)
'    NowP% = WindowFromPoint(CurP.X, CurP.Y)

    'If NowP% <> ThenP% Then
    '    ThenP% = NowP%
    '    APISpy_MouseOver = NowP%
    'End If
End Function

Function APISpy_Parent()
Dim Parnt%
Parnt% = GetParent(APISpy_MouseOver)
APISpy_Parent = Parnt%
End Function

Function APISpy_ParentName()
winhand% = APISpy_Parent
pspace$ = String$(250, 0)
pclassname% = GetClassName(winhand%, pspace$, 250)
APISpy_ParentName = pspace$
End Function


Function APISpy_WindowText()
Dim WinTLen%, WindowText%, Spce$
WinTLen% = GetWindowTextLength(APISpy_MouseOver)
Spce$ = String$(WinTLen%, 0)
WindowText% = GetWindowText(APISpy_MouseOver, Spce$, (WinTLen% + 1))
APISpy_WindowText = Spce$
End Function



Function Windows_GetUser()
'returns the name of the user in windows
     
    ' Dim Spcs As String
    ' Dim lent As Long
    ' Spcs = Space$(255)
    ' lent = Len(Spcs)
    ' Call GetUserName(Spcs, lent)

    'If lent > 0 Then
    '     Windows_GetUser = Left$(Spcs, lent)
    'Else
    '     Windows_GetUser = vbNullString
    'End If

    End Function



Public Sub ClearRecentDocs()
SHAddToRecentDocs 0, 0 ' Clear All Items Under The Documents Menu

End Sub


Sub RunMenuByString(Application, StringSearch)
' From Hix he gets full credit

    Dim ToSearch As Integer, MenuCount As Integer, FindString
    Dim ToSearchSub As Integer, menuItemCount As Integer, GetString
    Dim SubCount As Integer, MenuString As String, GetStringMenu As Integer
    Dim MenuItem As Integer, RunTheMenu As Integer
    
    ToSearch% = GetMenu(Application)
    MenuCount% = GetMenuItemCount(ToSearch%)
    
    For FindString = 0 To MenuCount% - 1
        ToSearchSub% = GetSubMenu(ToSearch%, FindString)
        menuItemCount% = GetMenuItemCount(ToSearchSub%)
        For GetString = 0 To menuItemCount% - 1
            SubCount% = GetMenuItemID(ToSearchSub%, GetString)
            MenuString$ = String$(100, " ")
            GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)
            If InStr(UCase(MenuString$), UCase(StringSearch)) Then
                MenuItem% = SubCount%
                GoTo MatchString
            End If
    Next GetString
    Next FindString
MatchString:
    RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub
Sub Paste(hWnd As Integer)
A$ = Clipboard.GetText()
If A$ = "" Then: Exit Sub
x = SendMessageByString(whnd%, WM_PASTE, 0, A$)

End Sub

Sub CntAltDel_Disable()
     Dim ret As Integer
     Dim pOld As Boolean

     ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Sub CntAltDel_Enable()
     Dim ret As Integer
     Dim pOld As Boolean

     ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub




