Attribute VB_Name = "API"
'Window Functions
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'The Send Message functions
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLONG Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'Flash a windows title bar
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
'Bring a window to the top
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
'Does thw window exist?
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'Is the window Enabled?
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
'Is the window visible?
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'Load the cursor
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
'Load the cursor from a file
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
'Load an icon
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
'System parameters
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
'The beep command
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
'Empty all data on the clipboard
Public Declare Function EmptyClipboard Lib "user32" () As Long
'Delete a file
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
'Play a sound, open/close cd-drive
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'Get windows tick count
Public Declare Function GetTickCount Lib "kernel32" () As Long
'Set the cursor
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'Show Cursor
Public Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
'Swap Mouse Buttons
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
'Empty Recycle BIN
Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
'Add to the recent documents list
Public Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pv As String)
'Load a Bitmap Image
Public Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
'Get the current User Name
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Get the computer Name
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Display a message box
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
'Move a File
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
'Set the computer Name
Public Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
'Set the cursor position
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Get the cursor position
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
'Execute a program
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Dial the internet
Public Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
'Disconnect the internet
Public Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
'The close windows functions
Public Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
'Get the system directory path
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Get the windows directory path
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Get the temporary directory path
Public Declare Function GetTempDirectory Lib "kernel32" Alias "GetTempPathA" (ByVal lBufferLength As Long, ByVal strBuffer As String) As Long
'Get the windows version
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
'Get a windows menu
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'Get a windows submenu
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Get the windows menus item ID
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Set the menu's bitmap images
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
'Keyboard events. Like SendKeys
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Create a folder
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'Create an extended directory
Public Declare Function CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'Play a sound
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
'Get the time
Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
'The run dialog box
Public Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal UnknownP1 As Long, ByVal UnknownP2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
'get the file information
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
'Delete a folder
Public Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
'Get the clipboard data
Public Declare Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As Long
'Set a parent
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Set the backcolor
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Type SHFILEINFO
        hIcon As Long
        iIcon As Long
        dwAttributes As Long
        szDisplayName As String * MAX_PATH
        szTypeName As String * 80
End Type

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400

Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91
Public Const VK_CAPITAL = &H14

Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

Public Const MF_BITMAP = &H4&

Public Const EWX_SHUTDOWN As Long = 1
Public Const EWX_REBOOT As Long = 2
Public Const EWX_LOGOFF As Long = 0

Public Const WM_GETTEXT = &HD
Public Const WM_SETTEXT = &HC

Public Const SW_SHOW = 5
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
Public Const SW_MINIMIZE As Long = 6
Public Const SW_HIDE = 0

Public Const Internet_Autodial_Force_Unattended As Long = 2

Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const SPI_SETDESKWALLPAPER = 20

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Public Enum PlaySoundSettings
    sndWavFile = 0
    sndWindowsStart = 1
    sndWindowsExit = 2
    sndRestoreUp = 3
    sndRestoreDown = 4
    sndApplicationError = 5
    sndQuestion = 6
    sndAsterisk = 7
    sndExclamation = 8
    sndSystemHand = 9
End Enum

Public Const IDC_ARROW = 32512&
Public Const IDC_CROSS = 32515&
Public Const IDC_IBEAM = 32513&
Public Const IDC_ICON = 32641&
Public Const IDC_NO = 32648&
Public Const IDC_SIZE = 32640&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_UPARROW = 32516&
Public Const IDC_WAIT = 32514&

Type PointAPI
        X As Long
        Y As Long
End Type

Public Sub MakeRecycleBinEmpty(Optional ByVal Drive As String, Optional NoConfirmation As Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)
 Dim hwnd, Flags As Long
 On Error Resume Next
 hwnd = Screen.ActiveForm.hwnd
 If Len(Drive) > 0 Then _
  Drive = Left$(Drive, 1) & ":\"
 Flags = (NoConfirmation And &H1) Or (NoProgress And &H2) Or (NoSound And &H4)
 SHEmptyRecycleBin hwnd, Drive, Flags
End Sub

Public Function StringToByteArray(str As String) As Variant
Dim bray() As Byte
Dim cnt As Integer
Dim ln As Integer
ln = Len(str)
ReDim bray(ln)
For cnt = 0 To ln - 1
    bray(cnt) = Asc(Mid(str, cnt + 1, 1))
Next cnt
bray(ln) = 0
StringToByteArray = bray
End Function

Public Function ShutDown()
Dim lngResult
lngResult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Function

Public Function Restart()
Dim lngResult
lngResult = ExitWindowsEx(EWX_REBOOT, 0&)
End Function

Public Function LogOff()
Dim lngResult
lngResult = ExitWindowsEx(EWX_LOGOFF, 0&)
End Function
