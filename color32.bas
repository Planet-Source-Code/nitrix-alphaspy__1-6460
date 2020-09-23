Attribute VB_Name = "color32"

'''''''''''''''''''''''
'color32.bas by nitrix'
''''''''''''''''''''''`

'api declarations: 'A'
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

'api declarations: 'B'
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

'api declarations: 'C'
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Public Declare Function CharNext Lib "user32" Alias "CharNextA" (ByVal lpsz As String) As String
Public Declare Function CharPrev Lib "user32" Alias "CharPrevA" (ByVal lpszStart As String, ByVal lpszCurrent As String) As String
Public Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CopyFile& Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long)
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function CreateSolidBrush& Lib "gdi32" (ByVal crColor As Long)

'api declarations: 'D'
Public Declare Function DeleteFile& Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String)
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'api declarations: 'E'
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

'api declarations: 'F'
Public Declare Function FillRect Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'api declarations: 'G'
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetComputerName& Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC& Lib "user32" (ByVal hwnd As Long)
Public Declare Function GetFileSize& Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long)
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetObjectAPIBynum Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByVal lpObject As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetPixelFormat Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetProfileString& Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long)
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer

'api declarations: 'I'
Public Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Public Declare Function IsChild Lib "user32" (ByVal hwndparent As Long, ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InsertMenuByNum Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long

'api declarations: 'M'
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciExecute Lib "winmm.dll" Alias "MciExecute" (ByVal lpstrCommand As String) As Long
Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function MoveWindow& Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long)

'api declarations: 'O'
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'api declarations: 'P'
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'api declarations: 'R'
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function Rectangle& Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ReleaseDC& Lib "user32" (ByVal hwnd As Long, ByVal hDc As Long)
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal nBytes As Long)

'api declarations: 'S'
Public Declare Function SelectObject& Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "SHELL32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'api declarations: 'T'
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

'api declarations: 'V'
Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Public Declare Function VkKeyScanEx Lib "user32" Alias "VkKeyScanExA" (ByVal Ch As Byte, ByVal dwhkl As Long) As Integer

'api declarations: 'W'
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const MF_ENABLED = &H0&
Public Const MF_STRING = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const CB_ADDSTRING& = &H143
Public Const CB_DELETESTRING& = &H144
Public Const CB_FINDSTRINGEXACT& = &H158
Public Const CB_GETCOUNT& = &H146
Public Const CB_GETITEMDATA = &H150
Public Const CB_GETLBTEXT& = &H148
Public Const CB_RESETCONTENT& = &H14B
Public Const CB_SETCURSEL& = &H14E
Public Const CB_FINDSTRING& = &H18F

Public Const COLOR_MENU = 4

Public Const DT_LEFT = &H0
Public Const DT_SINGLELINE = &H20
Public Const DT_TOP = &H0

Public Const EM_GETLINECOUNT& = &HBA
Public Const EM_SETREADONLY = &HCF
Public Const ENTER_KEY = 13

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4

Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40

Public Const GWL_WNDPROC = -4
Public Const GWL_EXSTYLE = (-20)

Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_CHILD = 5

Public Const HWND_NOTTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_ADDSTRING& = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Const LB_GETCOUNT& = &H18B
Public Const LB_GETCURSEL& = &H188
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN& = &H18A
Public Const LB_INSERTSTRING = &H181
Public Const LB_RESETCONTENT& = &H184
Public Const LB_SETCURSEL = &H186
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETSEL = &H185

Public Const LF_FACESIZE = 32

Public Const ODS_SELECTED = &H1

Public Const ODT_MENU = 1

Public Const SC_SCREENSAVE = &HF140

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_FLAGS = SND_ASYNC Or SND_NODEFAULT
Public Const SND_FLAGS2 = SND_ASYNC Or SND_LOOP

Public Const SPI_SCREENSAVERRUNNING = 97
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_SETDRAGFULLWINDOWS = 37
Public Const SPIF_SENDWININICHANGE = &H2

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE& = 3
Public Const SW_MINIMIZE& = 6
Public Const SW_RESTORE& = 9
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE

Public Const SYS_ADD = &H0
Public Const SYS_DELETE = &H2
Public Const SYS_MESSAGE = &H1
Public Const SYS_ICON = &H2
Public Const SYS_TIP = &H4

Public Const SYSTEM_FONT = 13

Public Const TRANSPARENT = 1

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SNAPSHOT = &H2C
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_DRAWITEM = &H2B
Public Const WM_GETFONT = &H31&
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MEASUREITEM = &H2C
Public Const WM_MENUSELECT = &H11F
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &HF012
Public Const WM_RBUTTONDBLCLK& = &H206
Public Const WM_RBUTTONDOWN& = &H204
Public Const WM_RBUTTONUP& = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const WM_USER& = &H400

Public Const WS_EX_TRANSPARENT = &H20&

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const OP_FLAGS = PROCESS_READ Or RIGHTS_REQUIRED

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public SysTray As NOTIFYICONDATA

Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Type COLORRGB
  red As Long
  green As Long
  blue As Long
End Type

Public Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Public Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As Long
End Type

Public Function spy_color() As Long
'gets color the mouse is over
  On Error Resume Next
  Dim cursorpos As POINTAPI
  Dim rDC As Long, rPixel As Long
  Call GetCursorPos(cursorpos)
  rDC& = GetDC(0&)
  rPixel& = GetPixel(rDC&, cursorpos.X, cursorpos.Y)
  Call ReleaseDC(0&, rDC&)
  spy_color& = rPixel&
End Function
Function rgb_get(ByVal CVal As Long) As COLORRGB
  frmColor.Label20.Caption = Int(CVal / 65536)
  frmColor.Label18.Caption = Int((CVal - (65536 * frmColor.Label20.Caption)) / 256)
  frmColor.Label19.Caption = CVal - (65536 * frmColor.Label20.Caption + 256 * frmColor.Label18.Caption)

End Function
Public Function spy_colorqb() As String
 'gets qbasic color
  Dim sColor As Long
  sColor& = spy_color&
  Select Case sColor&
  Case QBColor(0)
    spy_colorqb$ = "0"
  Case QBColor(1)
    spy_colorqb$ = "1"
  Case QBColor(2)
    spy_colorqb$ = "2"
  Case QBColor(3)
    spy_colorqb$ = "3"
  Case QBColor(4)
    spy_colorqb$ = "4"
  Case QBColor(5)
    spy_colorqb$ = "5"
  Case QBColor(6)
    spy_colorqb$ = "6"
  Case QBColor(7)
    spy_colorqb$ = "7"
  Case QBColor(8)
    spy_colorqb$ = "8"
  Case QBColor(9)
    spy_colorqb$ = "9"
  Case QBColor(10)
    spy_colorqb$ = "10"
  Case QBColor(11)
    spy_colorqb$ = "11"
  Case QBColor(12)
    spy_colorqb$ = "12"
  Case QBColor(13)
    spy_colorqb$ = "13"
  Case QBColor(14)
    spy_colorqb$ = "14"
  Case QBColor(15)
    spy_colorqb$ = "15"
  Case Else
    spy_colorqb$ = "n/a"
  End Select
End Function
Public Function spy_cursorx() As Long
  'gets the x position of the cursor
  Dim cursorpos As POINTAPI
  Call GetCursorPos(cursorpos)
  spy_cursorx& = cursorpos.X
End Function

Public Function spy_cursory() As Long
'get y position of cursor
  Dim cursorpos As POINTAPI
  Call GetCursorPos(cursorpos)
  spy_cursory& = cursorpos.Y
End Function

Public Sub form_drag(frm As Form)
  'drags the form
  Call ReleaseCapture
  Call PostMessage(frm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0&)
End Sub

Public Sub form_center(frm As Form)
  'centers form
  frm.Top = (Screen.Height * 0.85) / 2 - frm.Height / 2
  frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub

Public Function dec_2bin(ByVal nDec As Integer) As String
    Dim i As Integer
    Dim j As Integer
    Dim sHex As String
    Const HexChar As String = "0123456789ABCDEF"
    
    sHex = Hex(nDec) 'That the only part that is different


    For i = 1 To Len(sHex)
        nDec = InStr(1, HexChar, Mid(sHex, i, 1)) - 1


        For j = 3 To 0 Step -1
            dec_2bin = dec_2bin & nDec \ 2 ^ j
            nDec = nDec Mod 2 ^ j
        Next j
    Next i
    'Remove the first unused 0
    i = InStr(1, dec_2bin, "1")
    If i <> 0 Then dec_2bin = Mid(dec_2bin, i)
End Function

Public Function hex_fill(iHex As String) As String
  Dim iX As Long
  iX& = Len(iHex$)
  Select Case iX&
  Case "6"
    hex_fill$ = iHex$
  Case "5"
    hex_fill$ = "0" & iHex$
  Case "4"
    hex_fill$ = "00" & iHex$
  Case "3"
    hex_fill$ = "000" & iHex$
  Case "2"
    hex_fill$ = "0000" & iHex$
  Case "1"
    hex_fill$ = "00000" & iHex$
  Case Else
    hex_fill$ = iHex$
  End Select
End Function
Public Function spy_colorvb() As String
 'gets vb color
  Dim sColor As Long
  sColor& = spy_color&
  Select Case sColor&
  Case vbBlack
    spy_colorvb$ = "vbBlack"
  Case vbRed
    spy_colorvb$ = "vbRed"
  Case vbGreen
    spy_colorvb$ = "vbGreen"
  Case vbYellow
    spy_colorvb$ = "vbYellow"
  Case vbBlue
    spy_colorvb$ = "vbBlue"
  Case vbMagenta
    spy_colorvb$ = "vbMagenta"
  Case vbCyan
    spy_colorvb$ = "vbCyan"
  Case vbWhite
    spy_colorvb$ = "vbWhite"
  Case Else
    spy_colorvb$ = "n/a"
  End Select
End Function
Public Function hex_2rgb(iHex As String) As String
  Dim iR As String, iG As String, iB As String
  iR$ = Mid(iHex$, 1, 2)
  iG$ = Mid(iHex$, 3, 2)
  iB$ = Mid(iHex$, 5, 2)
  If iHex$ = "0" Then
    hex_2rgb$ = "000000"
  Else
    hex_2rgb$ = iB$ + iG$ + iR$
  End If
End Function

