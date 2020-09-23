Attribute VB_Name = "killz"
'====================================================
'Sup, this is killz' bas file. It took me a long time to
'create and I got almost all the api functions from
'www.allapi.net's API Guide and API Toolshed. Its there
'for download so check it out. I use both aol95 and aol6.0
'Their aren't many aol functions here, but their are some
'useful ones. Im not a big aol programmer that much.
'I hope you find some use to this .bas file and dont
'be a lamer and rename this and say you coded it.
'
'
'                   Later,
'                         killz.
'====================================================


Option Explicit
Public busted As Boolean
Dim hMenu As Long
Dim PrevProc
Public Const MAX_PATH = 260
Public stopbust As Boolean
Public roombusted As Boolean
Global info
Global g
Global allcharacters
Global molestate()
Public Const HIGHEST_VOLUME_SETTING = 100 '%


Private Type VolumeSetting
    LeftVol As Integer
    RightVol As Integer
End Type

Public Type AOLSHIT
   GetUser As String
   sendim As Long
End Type

Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
' used for enumerating registrykeys
Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
' used for import/export registry key
Public Type LUID
  lowpart As Long
  highpart As Long
End Type
Public Type LUID_AND_ATTRIBUTES
  pLuid As LUID
  Attributes As Long
End Type
Public Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges As LUID_AND_ATTRIBUTES
End Type



Public Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hwndParent As Long
    cy As Long
    cx As Long
    y As Long
    X As Long
    style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type

Public Enum dwRop

    WHITENESS = &HFF0062
    BLACKNESS = &H42
    SRCAND = &H8800C6
    SRCCOPY = &HCC0020
    SRCINVERT = &H660046
    SRCERASE = &H440328
    SRCPAINT = &HEE0086
    
End Enum


Type WIN32_FIND_DATA ' 318 Bytes
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved_ As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
    End Type
    
    Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Const FLAG_ICC_FORCE_CONNECTION = &H1
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function CreateStatusWindow Lib "comctl32.dll" (ByVal style As Long, ByVal lpszText As String, ByVal hwndParent As Long, ByVal wID As Long) As Long
Public Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function ChangeDisplaySettings Lib "user32.dll" Alias "ChangeDisplaySettingsA" (ByRef lpDevMode As DEVMODE, ByVal dwFlags As Long) As Long
Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
lpRect As RECT) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal _
hDC As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal _
crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, _
ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "user32" (ByVal hDC As Long, ByVal hObject _
As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetFreeResources Lib "RSRC32" Alias "_MyGetFreeSystemResources32@4" (ByVal lWhat As Long) As Long
Public Declare Function CryptGetProvParam Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Public Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Public Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Public Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Declare Function RegisterDLL Lib "Regist10.dll" Alias "REGISTERDLL" (ByVal DllPath As String, bRegister As Boolean) As Boolean
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function AddPrinter Lib "winspool.drv" Alias "AddPrinterA" (ByVal pName As String, ByVal Level As Long, pPrinter As Any) As Long
Public Declare Function AddPrinterConn Lib "winspool.drv" Alias "AddPrinterConnectionA" (ByVal pName As String) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
Public Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function FindFirstFile& Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA)
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Declare Function DirectoryPathExi Lib "imagehlp.dll" Alias "MakeSureDirectoryPathExists" (ByVal lpPath As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Public Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Public Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal Dest&, ByVal nCount&)
Public Declare Function dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal Dest As Long, ByVal nCount&)
Public Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Public Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Public Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Public Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function WriteFile Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function DiskSpaceFree Lib "STKIT432.DLL" () As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef source As Any, ByVal nBytes As Long)
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function GetAllUsersProfileDirectory Lib "userenv.dll" Alias "GetAllUsersProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Public Declare Function GetDefaultUserProfileDirectory Lib "userenv.dll" Alias "GetDefaultUserProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Public Declare Function GetProfilesDirectory Lib "userenv.dll" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Public Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.
Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long

Dim sbox(255)
Dim Key(255)

'FindWindow/FindWindowEx/Childbytitle/childbyclass
Public Const FINDAOL = "AOL Frame25"
Public Const IEXPLORE = "IEFrame"
Public Const AIMSIGN = "AIM_CSignOnWnd"
Public Const AIMONLINE = "_Oscar_BuddyListWin"
Public Const AmericCH = "AOL Child"
Public Const MDI_Client = "MDIClient"
Public Const AIMTABG = "_Oscar_TabGroup"
Public Const AIMTABC = "_Oscar_TabCtrl"
Public Const IconButton = "_Oscar_IconBtn"

Public Const CB_GETCOUNT = &H146
Public Const CB_SETCURSEL = &H14E

'Encryption Const
Public Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
Public Const PROV_RSA_FULL As Long = 1
Public Const PP_NAME As Long = 4
Public Const PP_CONTAINER As Long = 6
Public Const CRYPT_NEWKEYSET As Long = 8
Public Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Public Const ALG_CLASS_HASH As Long = 32768
Public Const ALG_TYPE_ANY As Long = 0
Public Const ALG_TYPE_STREAM As Long = 2048
Public Const ALG_SID_RC4 As Long = 1
Public Const ALG_SID_MD5 As Long = 3
Public Const CALG_MD5 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Public Const CALG_RC4 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Public Const ENCRYPT_ALGORITHM As Long = CALG_RC4
Public Const ENCRYPT_NUMBERKEY As String = "16006833"
Public lngCryptProvider As Long
Public avarSeedValues As Variant
Public lngSeedLevel As Long
Public lngDecryptPointer As Long
Public astrEncryptionKey(0 To 131) As String
Public Const lngALPKeyLength As Long = 8
Public strKeyContainer As String
'My Constants
Public Const WM_GETCHATTEXT = 14
Public Const PL_GETCERTAIN = 13
' Color constants
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
'System Resources
Const GFSR_SYSTEMRESOURCES = 0
Const GFSR_GDIRESOURCES = 1
Const GFSR_USERRESOURCES = 2
' ExWindowStyles
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&
' Window styles
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_DBLCLKS = &H8
Public Const CS_OWNDC = &H20
Public Const CS_CLASSDC = &H40
Public Const CS_PARENTDC = &H80
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOCLOSE = &H200
Public Const CS_SAVEBITS = &H800
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_PUBLICCLASS = &H4000
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const CB_DELETESTRING = &H144
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_RESETCONTENT = &H14B
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_SIZE = &H5
Public Const WM_PASTE = &H302
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const conMCIAppTitle = "MCI Control Application"
Public Const conMCIErrInvalidDeviceID = 30257
Public Const conMCIErrDeviceOpen = 30263
Public Const conMCIErrCannotLoadDriver = 30266
Public Const conMCIErrUnsupportedFunction = 30274
Public Const conMCIErrInvalidFile = 30304
Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
Public Const FADE_BLUE = &HFF0000
Public Const FADE_YELLOW = &HFFFF&
Public Const FADE_WHITE = &HFFFFFF
Public Const FADE_BLACK = &H0&
Public Const FADE_PURPLE = &HFF00FF
Public Const FADE_GREY = &HC0C0C0
Public Const FADE_PINK = &HFF80FF
Public Const FADE_TURQUOISE = &HC0C000
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &HF012
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
Public Const LB_MULTIPLEADDSTRING = &H1B1
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
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9
Public Const VK_UP = &H26
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNA = 8
Public Const SW_MAX = 10
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const WM_SYSCOMMAND = &H112
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
Public Const PROCESS_VM_READ = &H10
Public Const ENTER_KEY = 13
Const MB_DEFBUTTON1 = &H0&
Const MB_DEFBUTTON2 = &H100&
Const MB_DEFBUTTON3 = &H200&
Const MB_ICONASTERISK = &H40&
Const MB_ICONEXCLAMATION = &H30&
Const MB_ICONHAND = &H10&
Const MB_ICONINFORMATION = MB_ICONASTERISK
Const MB_ICONQUESTION = &H20&
Const MB_ICONSTOP = MB_ICONHAND
Const MB_OK = &H0&
Const MB_OKCANCEL = &H1&
Const MB_YESNO = &H4&
Const MB_YESNOCANCEL = &H3&
Const MB_ABORTRETRYIGNORE = &H2&
Const MB_RETRYCANCEL = &H5&
' Standard ID's of cursors
Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_SIZE = 32640&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_NO = 32648&
Public Const IDC_APPSTARTING = 32650&
Public Const GWL_WNDPROC = -4

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&

Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte

Const DisplayErrorMsg = False

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer

    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer

    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Dim DevM As DEVMODE

Type COLORRGB
  red As Long
  Green As Long
  blue As Long
End Type

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   y As Long
End Type

Public Const RGN_OR = 2
Public lngRegion As Long
Public Prg As String, Sect As String ' for savesettings
Public skindir As String
Dim PI As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim suma As Double
Dim sumb As Double
Dim diffa As Double
Dim diffb As Double
Dim AnswerA As Double
Dim AnswerB As Double
Dim FirstMultA As Double
Dim FirstMultB As Double
Dim FirsteA As Double
Dim FirsteB As Double
Dim SecondMultA As Double
Dim SecondMultB As Double
Dim SecondeA As Double
Dim SecondeB As Double
Dim SumeA As Double
Dim SumeB As Double
Dim DivisionA As Double
Dim DivisionB As Double
Dim multiplicationa As Double
Dim multiplicationb As Double
Dim Divisor As Double
Dim Magnitude As Double
Dim SineA As Double
Dim SineB As Double
Dim CosineA As Double
Dim CosineB As Double
Dim Answer1A As Double
Dim Answer1B As Double
Dim Answer2A As Double
Dim Answer2B As Double
Dim BaseA As Double
Dim BaseB As Double
Dim LogA As Double
Dim LogB As Double
Dim MultA As Double
Dim MultB As Double
Public DialogCaption As String



Function FileFound(strFileName As String) As Boolean
    'Code Created by Lucian
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim hFindFirst As Long
    hFindFirst = FindFirstFile(strFileName, lpFindFileData)


    If hFindFirst > 0 Then
        FindClose hFindFirst
        FileFound = True
    Else
        FileFound = False
    End If
End Function


Public Sub Form_Center(f As Form)
    f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
    f.Left = Screen.Width / 2 - f.Width / 2
End Sub


Public Function BlankString() As String
    BlankString$ = Chr(32) & Chr(160)
End Function

Function GetClassNameNow(Ret As String)
Dim winwnd As Long
Dim lpClassName As String
Dim retval As Long
    winwnd = FindWindow(vbNullString, UCase(Ret$))
    If winwnd = 0 Then MsgBox "Couldn't find the window ...": Exit Function
    lpClassName = Space(256)
    retval = GetClassName(winwnd, lpClassName, 256)
    GetClassNameNow = Left$(lpClassName, retval)
End Function

Public Function MakeIt3d(TheForm As Form, TheControl As Control)
Dim OldMode As Long
If TheForm.AutoRedraw = False Then
    OldMode = TheForm.ScaleMode
        TheForm.ScaleMode = 3
        TheForm.AutoRedraw = True
        TheForm.CurrentX = TheControl.Left - 1
        TheForm.CurrentY = TheControl.Top + TheControl.Height
        TheForm.Line -Step(0, -(TheControl.Height + 1)), RGB(90, 90, 90)
        TheForm.Line -Step(TheControl.Width + 1, 0), RGB(90, 90, 90)
        TheForm.Line -Step(0, TheControl.Height + 1), RGB(255, 255, 255)
        TheForm.Line -Step(-(TheControl.Width + 1), 0), RGB(255, 255, 255)
        TheForm.AutoRedraw = False
    TheForm.ScaleMode = OldMode
End If
If TheForm.AutoRedraw = True Then
    OldMode = TheForm.ScaleMode
        TheForm.ScaleMode = 3
        TheForm.CurrentX = TheControl.Left - 1
        TheForm.CurrentY = TheControl.Top + TheControl.Height
        TheForm.Line -Step(0, -(TheControl.Height + 1)), RGB(90, 90, 90)
        TheForm.Line -Step(TheControl.Width + 1, 0), RGB(90, 90, 90)
        TheForm.Line -Step(0, TheControl.Height + 1), RGB(255, 255, 255)
        TheForm.Line -Step(-(TheControl.Width + 1), 0), RGB(255, 255, 255)
    TheForm.ScaleMode = OldMode
End If
End Function


Public Sub Window_Enable(Window)
    Call EnableWindow(Window, 1)
End Sub





Public Sub RemoveItem_Combo(ComboWin As Long, thestring As String)
Dim FindIt As Long, DeleteIt As Long
FindIt = SendMessageByString(ComboWin, CB_FINDSTRINGEXACT, -1, thestring)
If FindIt <> -1 Then
    Call SendMessageByString(ComboWin, CB_DELETESTRING, FindIt, 0)
End If
End Sub
Public Sub RemoveItem_ListBox(ListWin, thestring)
Dim FindIt As Long, DeleteIt As Long
FindIt = SendMessageByString(ListWin, LB_FINDSTRINGEXACT, -1, thestring)
If FindIt <> -1 Then
    Call SendMessageByString(ListWin, LB_DELETESTRING, FindIt, 0)
End If
End Sub
Public Sub Draw3DBorder(c As Control, iLook As Integer)
'Makes A Control Look 3D
Dim iOldScaleMode As Integer, iFirstColor As Integer
Dim iSecondColor As Integer, RAISED As Variant, PIXELS As Variant
    If iLook = RAISED Then
        iFirstColor = 15
        iSecondColor = 8
    Else
        iFirstColor = 8
        iSecondColor = 15
    End If
iOldScaleMode = c.Parent.ScaleMode
c.Parent.ScaleMode = PIXELS
c.Parent.Line (c.Left, c.Top - 1)-(c.Left + c.Width, c.Top - 1), QBColor(iFirstColor)
c.Parent.Line (c.Left - 1, c.Top)-(c.Left - 1, c.Top + c.Height), QBColor(iFirstColor)
c.Parent.Line (c.Left + c.Width, c.Top)-(c.Left + c.Width, c.Top + c.Height), QBColor(iSecondColor)
c.Parent.Line (c.Left, c.Top + c.Height)-(c.Left + c.Width, c.Top + c.Height), QBColor(iSecondColor)
c.Parent.ScaleMode = iOldScaleMode
End Sub
Public Sub WriteToLog(what As String, LoGPath As String)
Dim X As Long, sSTR As String
If LoGPath = "" Then Exit Sub
If InStr(LoGPath, ".") = 0 Then Exit Sub
X& = FreeFile
Open LoGPath For Binary Access Write As X&
    sSTR$ = what & Chr(10)
    Put #1, LOF(1) + 1, sSTR$
Close X&
End Sub
Public Function WindowSPYLabels(WinHdl, WinClass, WinTxT, WinStyle, WinIDNum, WinPHandle, WinPText, WinPClass, WinModule)
'Call This In A Timer
Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
Dim sClassName As String * 100, hWndOver As Long, hwndParent As Long
Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
Dim hInstance As Long, sParentWindowText As String * 100
Dim sModuleFileName As String * 100, r As Long
Static hWndLast As Long
    Call GetCursorPos(pt32)
    ptx = pt32.X
    pty = pt32.y
    hWndOver = WindowFromPointXY(ptx, pty)
    If hWndOver <> hWndLast Then
        hWndLast = hWndOver
        WinHdl.Caption = "Window Handle: " & hWndOver
        sWindowText = Space(100)
        r = GetWindowText(hWndOver, sWindowText, 100)
        WinTxT.Caption = "Window Text: " & Left(sWindowText, r)
        sClassName = Space(100)
        r = GetClassName(hWndOver, sClassName, 100)
        WinClass.Caption = "Window Class Name: " & Left(sClassName, r)
        lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)
        WinStyle.Caption = "Window Style: " & lWindowStyle
        hwndParent = GetParent(hWndOver)
            If hwndParent <> 0 Then
                wID = GetWindowWord(hWndOver, GWW_ID)
                WinIDNum.Caption = "Window ID Number: " & wID
                WinPHandle.Caption = "Parent Window Handle: " & hwndParent
                sParentWindowText = Space(100)
                r = GetWindowText(hwndParent, sParentWindowText, 100)
                WinPText.Caption = "Parent Window Text: " & Left(sParentWindowText, r)
                sParentClassName = Space(100)
                r = GetClassName(hwndParent, sParentClassName, 100)
                WinPClass.Caption = "Parent Window Class Name: " & Left(sParentClassName, r)
            Else
                WinIDNum.Caption = "Window ID Number: Not Available"
                WinPHandle.Caption = "Parent Window Handle: Not Available"
                WinPText.Caption = "Parent Window Text : Not Available"
                WinPClass.Caption = "Parent Window Class Name: Not Available"
            End If
                hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
                sModuleFileName = Space(100)
                r = GetModuleFileName(hInstance, sModuleFileName, 100)
        WinModule.Caption = "Module: " & Left(sModuleFileName, r)
    End If
End Function

Public Function Click_List(Window, index)
    Call SendMessage(Window, LB_SETCURSEL, ByVal CLng(index), ByVal 0&)
End Function
Public Function TileBitmap(TheForm As Form, theBitmap As PictureBox)
Dim Across As Integer, Down As Integer
theBitmap.AutoSize = True
    For Down = 0 To (TheForm.Width \ theBitmap.Width) + 1
        For Across = 0 To (TheForm.Height \ theBitmap.Height) + 1
            TheForm.PaintPicture theBitmap.Picture, Down * theBitmap.Width, Across * theBitmap.Height, theBitmap.Width, theBitmap.Height
    Next Across, Down
End Function
Public Sub Window_Maximize(Window)
    Call ShowWindow(Window, SW_MAXIMIZE)
End Sub
Public Sub Window_Minimize(Window)
    Call ShowWindow(Window, SW_MINIMIZE)
End Sub
Public Function MakeASCIIChart(list As ListBox)
Dim X As Long
For X = 33 To 255
    list.AddItem Chr(X)
Next X
End Function
Public Function WindowSPYTextBoxs(WinHdl As TextBox, WinClass As TextBox, WinTxT As TextBox, WinStyle As TextBox, WinIDNum As TextBox, WinPHandle As TextBox, WinPText As TextBox, WinPClass As TextBox, WinModule As TextBox)
'Call This In A Timer
Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
Dim sClassName As String * 100, hWndOver As Long, hwndParent As Long
Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
Dim hInstance As Long, sParentWindowText As String * 100
Dim sModuleFileName As String * 100, r As Long
Static hWndLast As Long
    Call GetCursorPos(pt32)
    ptx = pt32.X
    pty = pt32.y
    hWndOver = WindowFromPointXY(ptx, pty)
    If hWndOver <> hWndLast Then
        hWndLast = hWndOver
        WinHdl.Text = "Window Handle: " & hWndOver
        r = GetWindowText(hWndOver, sWindowText, 100)
        WinTxT.Text = "Window Text: " & Left(sWindowText, r)
        r = GetClassName(hWndOver, sClassName, 100)
        WinClass.Text = "Window Class Name: " & Left(sClassName, r)
        lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)
        WinStyle.Text = "Window Style: " & lWindowStyle
        hwndParent = GetParent(hWndOver)
            If hwndParent <> 0 Then
                wID = GetWindowWord(hWndOver, GWW_ID)
                WinIDNum.Text = "Window ID Number: " & wID
                WinPHandle.Text = "Parent Window Handle: " & hwndParent
                r = GetWindowText(hwndParent, sParentWindowText, 100)
                WinPText.Text = "Parent Window Text: " & Left(sParentWindowText, r)
                r = GetClassName(hwndParent, sParentClassName, 100)
                WinPClass.Text = "Parent Window Class Name: " & Left(sParentClassName, r)
            Else
                WinIDNum.Text = "Window ID Number: N/A"
                WinPHandle.Text = "Parent Window Handle: N/A"
                WinPText.Text = "Parent Window Text : N/A"
                WinPClass.Text = "Parent Window Class Name: N/A"
            End If
                hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
                r = GetModuleFileName(hInstance, sModuleFileName, 100)
        WinModule.Text = "Module: " & Left(sModuleFileName, r)
    End If
End Function

Public Sub ExtractAnIcon(CmmDlg As Control)
Dim sSourcePgm As String, lIcon As Long

Dim a%
    On Error Resume Next
  With CmmDlg
    .filename = sSourcePgm
    .CancelError = True
    .DialogTitle = "Select a DLL or EXE which includes Icons"
    .Filter = "Icon Resources (*.ico;*.exe;*.dll)|*.ico;*.exe;*.dll|All files|*.*"
    .Action = 1
    If Err Then
      Err.Clear
      Exit Sub
    End If
    sSourcePgm = .filename
    DestroyIcon lIcon
    End With
    Do
      lIcon = ExtractIcon(App.hInstance, sSourcePgm, a)
      If lIcon = 0 Then Exit Do
      a = a + 1
      DestroyIcon lIcon
    Loop
    If a = 0 Then
      MsgBox "No Icons in this file!"
    End If
End Sub




Public Sub Click(Icon)
    Call SendMessage(Icon, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Icon, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MIDI_Play(Midi As String)
Dim FilE As String
FilE$ = Dir(Midi$)
If FilE$ <> "" Then
    Call mciSendString("play " & Midi$, 0&, 0, 0)
End If
End Sub
Public Sub MIDI_Stop(Midi As String)
Dim FilE As String
FilE$ = Dir(Midi$)
If FilE$ <> "" Then
    Call mciSendString("stop " & Midi$, 0&, 0, 0)
End If
End Sub

Sub Click_Double(Icon&)
    Call SendMessageByNum(Icon&, WM_LBUTTONDBLCLK, &HD, 0)
End Sub


Public Function FindChildByTitle(Parent As Long, child As String)
    FindChildByTitle = FindWindowEx(Parent, 0&, vbNullString, child)
End Function



Sub Click_StartButton()
Dim Windows As Long, StartButton As Long
Windows& = FindWindow("Shell_TrayWnd", vbNullString)
StartButton& = FindWindowEx(Windows&, 0&, "Button", vbNullString)
Click (StartButton&)
End Sub


Public Sub Window_Hide(Window As Long)
    Call ShowWindow(Window, 0)
End Sub

Public Sub Window_Show(Window As Long)
    Call ShowWindow(Window, 5)
End Sub
Public Sub StayOffTop(f As Form)
    Call SetWindowPos(f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Sub DecompileProtect(ExeLocation)
Dim ThaFile As String, Cat As String
On Error Resume Next
    If ExeLocation = "" Then MsgBox "Executable File Not Found", vbOKOnly
ThaFile = FreeFile
Open ExeLocation For Binary As #ThaFile
    Cat = "."
Seek #ThaFile, 25
Put #ThaFile, , Cat
Close #1
If Err Then MsgBox "Not A Visual Basic Made File!", vbOKOnly, "Error In File": Exit Sub
MsgBox "Youre File Has Been Protected", vbOKOnly
End Sub

Public Function ClearDocuments()
    Call SHAddToRecentDocs(0, 0)
End Function

Public Function FindChildByClass(Parent, child)
    FindChildByClass = FindWindowEx(Parent, 0&, child, vbNullString)
End Function




Public Sub File_Delete(FilE$)
Dim NoFreeze As Long
If Not File_Exists(FilE$) Then Exit Sub
Kill FilE$
NoFreeze& = DoEvents()
End Sub


Public Sub DeleteListItem(list As ListBox, item$)

    item$ = list.ListIndex
    list.RemoveItem (item$)
End Sub


Public Function DirExists(TheDir)
Dim Test As Integer
On Error Resume Next
    If Right(TheDir, 1) <> "/" Then TheDir = TheDir & "/"
Test = Len(Dir$(TheDir))
If Err Or Test = 0 Then DirExists = False: Exit Function
DirExists = True
End Function
Public Function File_Exists(ByVal filename As String) As Integer
Dim Test As Integer
On Error Resume Next
    Test = Len(Dir$(filename))
If Err Or Test = 0 Then File_Exists = False: Exit Function
File_Exists = True
End Function



Public Function File_GetAttributes(TheFile As String)
Dim FilE As String
    FilE = Dir(TheFile)
If FilE <> "" Then File_GetAttributes = GetAttr(TheFile)
End Function
Public Sub File_SetHidden(TheFile As String)
Dim FilE As String
    FilE = Dir(TheFile)
If FilE <> "" Then SetAttr TheFile, vbHidden
End Sub

Public Sub File_SetReadOnly(TheFile As String)
Dim FilE As String
    FilE = Dir(TheFile)
If FilE <> "" Then SetAttr TheFile, vbReadOnly
End Sub


Public Sub LoadFonts(list As Control)
Dim X As Long
list.Clear
For X = 1 To Screen.FontCount
    list.AddItem Screen.Fonts(X - 1)
Next
End Sub
Public Function GetClass(child&) As String
Dim sString As String, Plop As String
sString$ = String$(250, 0)
    GetClass = GetClassName(child, sString$, 250)
    GetClass = sString$
End Function
Public Function GetCaption(Window)
Dim windowtitle As String, WindowText As String, WindowLength As Long
WindowLength& = GetWindowTextLength(Window)
    windowtitle$ = String$(WindowLength&, 0)
    WindowText$ = GetWindowText(Window, windowtitle$, (WindowLength& + 1))
    GetCaption = windowtitle$
End Function

Public Function GetText(child)
Dim TheTrimmer As Long, TrmSpace As String, GetStr As Long
TheTrimmer& = SendMessageByNum(child, WM_GETCHATTEXT, 0&, 0&)
    TrmSpace$ = Space$(TheTrimmer)
GetStr = SendMessageByString(child, PL_GETCERTAIN, TheTrimmer + 1, TrmSpace$)
    GetText = TrmSpace$
End Function



Public Function TaskBar_Hide()
Dim Bar As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Call ShowWindow(Bar&, 0)
End Function
Public Function TaskBar_Show()
Dim Bar As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
    Call ShowWindow(Bar&, 5)
End Function
Public Function StartButton_Hide()
Dim Bar As Long, Button As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(Bar&, 0&, "Button", vbNullString)
Call ShowWindow(Button&, 0)
End Function
Public Function StartButton_Show()
Dim Bar As Long, Button As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(Bar&, 0&, "Button", vbNullString)
    Call ShowWindow(Button&, 5)
End Function




Public Sub Window_Close(Window)
    Call SendMessageByNum(Window, WM_CLOSE, 0, 0)
End Sub

Public Sub CenterForm(f As Form)
    f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
    f.Left = Screen.Width / 2 - f.Width / 2
End Sub

Private Sub ListBox2Clipboard(list As ListBox)
Dim sn As Long, TheList As String
For sn = 0 To list.ListCount - 1
If sn = 0 Then
    TheList = list.list(sn)
Else
    TheList = TheList & "," & list.list(sn)
End If
Next
Clipboard.Clear
TimeOut 0.1
Clipboard.SetText TheList
End Sub



Public Sub RunMenuByString(Window, StringSearch)
Dim FindWin As Long, CountMenu As Long, FindString As Long, MenuItem As Long
Dim FindWinSub As Long, MenuItemCount As Long, getstring As Long
Dim SubCount As Long, MenuString As String, GetStringMenu As Long
FindWin& = GetMenu(Window)
CountMenu& = GetMenuItemCount(FindWin&)

For FindString = 0 To CountMenu& - 1
    FindWinSub& = GetSubMenu(FindWin&, FindString)
    MenuItemCount& = GetMenuItemCount(FindWinSub&)
For getstring = 0 To MenuItemCount& - 1
    SubCount& = GetMenuItemID(FindWinSub&, getstring)
    MenuString$ = String$(100, " ")
    GetStringMenu& = GetMenuString(FindWinSub&, SubCount&, MenuString$, 100, 1)
If InStr(UCase(MenuString$), UCase(StringSearch)) Then
    MenuItem& = SubCount&
    GoTo MatchString
End If
Next getstring
Next FindString

MatchString:
    Call SendMessage(Window, WM_COMMAND, MenuItem&, 0)
End Sub


Public Sub MakeShortcut(ShortcutDir, ShortcutName, ShortcutPath)
Dim WinShortcutDir As String, WinShortcutName As String, WinShortcutExePath As String, retval As Long
    WinShortcutDir$ = ShortcutDir
    WinShortcutName$ = ShortcutName
    WinShortcutExePath$ = ShortcutPath
retval& = fCreateShellLink("", WinShortcutName$, WinShortcutExePath$, "")
    Name "C:\Windows\Start Menu\Programs\" & WinShortcutName$ & ".LNK" As WinShortcutDir$ & "\" & WinShortcutName$ & ".LNK"
End Sub


Public Sub ParentChange(frm As Form, Window&)
    Call SetParent(frm.hwnd, Window&)
End Sub


Public Function ReadINI(Header As String, Key As String, location As String) As String
Dim sString As String
    sString = String(750, Chr(0))
    Key$ = LCase$(Key$)
    ReadINI$ = Left(sString, GetPrivateProfileString(Header$, ByVal Key$, "", sString, Len(sString), location$))
End Function

Public Sub File_ReName(FilE$, NewName$)
Dim NoFreeze As Long
    Name FilE$ As NewName$
    NoFreeze& = DoEvents()
End Sub



Public Sub RunMenu(menu1 As Integer, menu2 As Integer)
Static Working As Integer
Dim Menus As Long, SubMenu As Long, ItemID As Long, Works As Long, MenuClick As Long
Menus& = GetMenu(FindWindow("AOL Frame25", vbNullString))
SubMenu& = GetSubMenu(Menus&, menu1)
ItemID = GetMenuItemID(SubMenu&, menu2)
Works = CLng(0) * &H10000 Or Working
MenuClick = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, ItemID, 0&)
End Sub

Public Sub Window_SetText(Window, Text)
    Call SendMessageByString(Window, WM_SETTEXT, 0, Text)
End Sub

Public Sub shutdownwindows()
Dim EWX_SHUTDOWN
    Dim MsgRes As Long
    MsgRes = MsgBox("Do you really want to Shut Down Windows 9x", vbYesNo Or vbQuestion)
    If MsgRes = vbNo Then Exit Sub
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End Sub


Public Sub StayOnTop(TheForm As Form)
    Call SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub



Public Function StringInList(TheList As ListBox, FindMe As String)
Dim a As Long
If TheList.ListCount = 0 Then GoTo ListEmpty
For a = 0 To TheList.ListCount - 1
TheList.ListIndex = a
    If UCase(TheList.Text) = UCase(FindMe) Then
        StringInList = a
    Exit Function
    End If
Next a
ListEmpty:
StringInList = -1
End Function






Public Sub TimeOut(Length)
    Dim begin As Long
    begin = Timer
Do While Timer - begin >= Length
    DoEvents
Loop
End Sub
Public Sub Pause(Length)
'Same As Timeout
    Dim begin As Long
    begin = Timer
Do While Timer - begin >= Length
    DoEvents
Loop
End Sub




Public Sub waitforok()
Dim waitforok As Long, OK As Long, OKButton As Long
Do
    DoEvents
    OK = FindWindow("#32770", "America Online")
    DoEvents
Loop Until OK <> 0
OKButton = FindWindowEx(OK, 0&, vbNullString, "OK")
    Call SendMessageByNum(OKButton, WM_LBUTTONDOWN, 0, 0&)
    Call SendMessageByNum(OKButton, WM_LBUTTONUP, 0, 0&)
End Sub

Public Sub WriteToINI(Header As String, Key As String, KeyValue As String, location As String)
    Call WritePrivateProfileString(Header$, UCase$(Key$), KeyValue$, location$)
End Sub
Public Function Form_Drag(Form As Form)
'This Goes In Mouse Down Events Of A Label/Button
    Call ReleaseCapture
    Call SendMessage(Form.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Function


Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Long
    ZeroPos = InStr(1, sInput, Chr$(0))
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

Function GetWinDir()
    Dim sSave As String, Ret As Long
    sSave = Space(255)
    Ret = GetWindowsDirectory(sSave, 255)
    sSave = Left$(sSave, Ret)
    GetWinDir = sSave
End Function

Function GetProfilesDir(who)
Dim dirst
Dim ttt
dirst = GetWinDir()
ttt = InStr(4, dirst, "\")
If ttt <> 0 Then
If FileFound(GetWinDir() & "Profiles\" & who) = False Then GetProfilesDir = False: MsgBox "That profiles member does not exist": Exit Function
GetProfilesDir = dirst & "profiles\" & who
ElseIf ttt = 0 Then
GetProfilesDir = dirst & "\profiles\" & who
End If
End Function

Function GetShortPath(strng As String)
Dim txt$
Dim ttt&
txt$ = String(165, 0)
ttt& = GetShortPathName(strng$, txt$, 165)
GetShortPath = txt$
End Function

Function RandomWinPos(win As Long, X As String, y As String, wx2 As String, wy2 As String)
Randomize
X = SetWindowPos(win&, HWND_TOPMOST, X * Rnd, y * Rnd, wx2 * Rnd, wy2 * Rnd, &H40)
End Function

Function RandomCursorPos(X As String, y As String)
Randomize
X = SetCursorPos(X * Rnd, y * Rnd)
End Function

Function RunAOLToolbar(MenuNumber As String, letter As String)
Dim aolframe&
aolframe& = FindWindow("AOL Frame25", vbNullString)
Dim aoltoolbar&
aoltoolbar& = FindWindowEx(aolframe&, 0&, "AOL Toolbar", vbNullString)
Dim aoltoolbar2
aoltoolbar2 = FindWindowEx(aoltoolbar&, 0&, "_AOL_Toolbar", vbNullString)
Dim AOLIcon
AOLIcon = FindWindowEx(aoltoolbar2, 0&, "_AOL_Icon", vbNullString)
Dim Count
For Count = 1 To MenuNumber
AOLIcon = FindWindowEx(aoltoolbar2, AOLIcon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
Dim menu
menu = FindWindow("#32768", vbNullString)
Dim found
found = IsWindowVisible(menu)
Loop Until found <> 0
letter = Asc(letter)
Call PostMessage(menu, WM_CHAR, letter, 0&)
End Function



Public Function FindChatRoom() As Long
Dim Counter As Long
Dim AOLStatic5 As Long
Dim AOLIcon3 As Long
Dim AOLStatic4 As Long
Dim aollistbox As Long
Dim AOLStatic3 As Long
Dim aolimage As Long
Dim AOLIcon2 As Long
Dim RICHCNTL2 As Long
Dim AOLStatic2 As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLCombobox As Long
Dim richcntl As Long
Dim AOLStatic As Long
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
richcntl& = FindWindowEx(aolchild&, 0&, "RICHCNTL", vbNullString)
AOLCombobox& = FindWindowEx(aolchild&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(aolchild&, AOLStatic&, "_AOL_Static", vbNullString)
RICHCNTL2& = FindWindowEx(aolchild&, richcntl&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon2& = FindWindowEx(aolchild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
aolimage& = FindWindowEx(aolchild&, 0&, "_AOL_Image", vbNullString)
AOLStatic3& = FindWindowEx(aolchild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLStatic3& = FindWindowEx(aolchild&, AOLStatic3&, "_AOL_Static", vbNullString)
aollistbox& = FindWindowEx(aolchild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic4& = FindWindowEx(aolchild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(aolchild&, AOLIcon2&, "_AOL_Icon", vbNullString)
For i& = 1& To 7&
    AOLIcon3& = FindWindowEx(aolchild&, AOLIcon3&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic5& = FindWindowEx(aolchild&, AOLStatic4&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or richcntl& = 0& Or AOLCombobox& = 0& Or AOLIcon& = 0& Or AOLStatic2& = 0& Or RICHCNTL2& = 0& Or AOLIcon2& = 0& Or aolimage& = 0& Or AOLStatic3& = 0& Or aollistbox& = 0& Or AOLStatic4& = 0& Or AOLIcon3& = 0& Or AOLStatic5& = 0&): DoEvents
    aolchild& = FindWindowEx(mdiclient&, aolchild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
    richcntl& = FindWindowEx(aolchild&, 0&, "RICHCNTL", vbNullString)
    AOLCombobox& = FindWindowEx(aolchild&, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        AOLIcon& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(aolchild&, AOLStatic&, "_AOL_Static", vbNullString)
    RICHCNTL2& = FindWindowEx(aolchild&, richcntl&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon2& = FindWindowEx(aolchild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    aolimage& = FindWindowEx(aolchild&, 0&, "_AOL_Image", vbNullString)
    AOLStatic3& = FindWindowEx(aolchild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLStatic3& = FindWindowEx(aolchild&, AOLStatic3&, "_AOL_Static", vbNullString)
    aollistbox& = FindWindowEx(aolchild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic4& = FindWindowEx(aolchild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(aolchild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    For i& = 1& To 7&
        AOLIcon3& = FindWindowEx(aolchild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic5& = FindWindowEx(aolchild&, AOLStatic4&, "_AOL_Static", vbNullString)
    If AOLStatic& And richcntl& And AOLCombobox& And AOLIcon& And AOLStatic2& And RICHCNTL2& And AOLIcon2& And aolimage& And AOLStatic3& And aollistbox& And AOLStatic4& And AOLIcon3& And AOLStatic5& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindChatRoom& = aolchild&
    Exit Function
End If
End Function
Function SecsToMins(Secs As Integer)
    If Secs < 60 Then SecsToMins = "00:" & Format(Secs, "00") Else SecsToMins = Format(Secs / 60, "00") & ":" & Format(Secs - Format(Secs / 60, "00") * 60, "00")
End Function
Function FindToolbar2() As Long
Dim AOL&, tool1&, tool2&
AOL& = FindWindow("AOL Frame25", vbNullString)
tool1& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
tool2& = FindWindowEx(tool1&, 0&, "_AOL_Toolbar", vbNullString)
FindToolbar2& = tool2&
End Function

Function FindAOLChild() As Long
Dim AOL&, MDI&, child&
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindChildByClass(AOL&, "MDIClient")
child& = FindChildByClass(MDI&, "AOL Child")
FindAOLChild& = child&
End Function

Function ClickToolbar(Icon As Long)
Call SendMessage(Icon, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(Icon, WM_KEYUP, VK_SPACE, 0)
End Function


Function ClickReadMail()
Dim Toolbar&, icon1&, icon2&
Toolbar& = FindToolbar2()
icon1& = FindChildByClass(Toolbar&, "_AOL_Icon")
icon2& = GetWindow(icon1&, 2)
Call ClickToolbar(icon2&)
End Function

Function GetSN() As String
Dim child&, txt$, sn$, scn$, X
child& = FindAOLChild()
Do
DoEvents
txt$ = GetText(child&)
If InStr(txt$, "Welcome, ") Then
X = InStr(txt$, " ")
sn$ = Mid(txt$, X + 1, Len(txt$))
scn$ = Mid(sn$, 1, Len(sn$) - 1)
Exit Do
End If
child& = GetWindow(child&, 2)
Loop
GetSN$ = scn$
End Function


Function Find30Chat() As Long
'_AOL_Static, _AOL_View, _AOL_Edit, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Image, _AOL_Static, _AOL_Static, _AOL_Listbox, _AOL_Static, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Static
Dim child&, staticc&, what, i, view&, edit&, Icon&, aolimage&, list&
child& = FindAOLChild()
what = GetChildrenNum()
For i = 1 To what
staticc& = FindChildByClass(child&, "_AOL_Static")
view& = FindChildByClass(child&, "_AOL_View")
edit& = FindChildByClass(child&, "_AOL_Edit")
Icon& = FindChildByClass(child&, "_AOL_Icon")
aolimage& = FindChildByClass(child&, "_AOL_Image")
list& = FindChildByClass(child&, "_AOL_Listbox")
If staticc& <> 0 And view& <> 0 And edit& <> 0 And Icon& <> 0 And aolimage& <> 0 And list& <> 0 Then
Find30Chat& = child&
Exit Function
Else
child& = GetWindow(child&, 2)
End If
Next i
End Function

Function GetChildrenNum()
Dim child&, num
child& = FindAOLChild()
If child& <> 0 Then num = num + 1
While child&
DoEvents
child& = GetWindow(child&, 2)
If child& <> 0 Then num = num + 1
Wend
GetChildrenNum = num
End Function

Function Add30Room(TheList As ListBox, AddUser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = Find30Chat&
    If Room& = 0& Then Exit Function
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
 
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If AddUser = True Then
                TheList.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If

End Function



Public Function MoveSprite(ByRef Sprite As PictureBox, ByRef mask As PictureBox, ByRef Background As PictureBox, ByVal Direction As String, ByVal Distance_Pixels As Long, ByVal startX As Single, startY As Single, ByVal Speed As Long, Optional ByVal NumberOfFrames As Long = 1) As String

Dim X As Single, y As Single

Select Case Direction

Case "Up"
    
    X = startX
    
    For y = startY To Distance_Pixels + startY
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, X, y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, mask, 0, 0, mask.ScaleWidth / NumberOfFrames, mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
    
    Next y

Case "Down"
    
    X = startX
    
    For y = Distance_Pixels + startY To startY Step -1
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, X, y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, mask, 0, 0, mask.ScaleWidth / NumberOfFrames, mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
        
    Next y

Case "Left"
    
    y = startY

    For X = Distance_Pixels + startX To startX Step -1
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, X, y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, mask, 0, 0, mask.ScaleWidth / NumberOfFrames, mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
        
    Next X

Case "Right"

    y = startY

    For X = startX To Distance_Pixels + startX
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, X, y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, mask, 0, 0, mask.ScaleWidth / NumberOfFrames, mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
    
    Next X

End Select

End Function

Public Function DoBitBlt(ByRef destination As PictureBox, ByVal DestinationX As Long, ByVal DestinationY As Long, ByVal DestinationWidth As Long, ByVal DestinationHeight As Long, ByRef Sprite As PictureBox, ByVal SpriteX As Long, ByVal SpriteY As Long, ByVal SpriteWidth As Long, ByVal SpriteHeight As Long, ByRef mask As PictureBox, ByVal MaskX As Long, ByVal MaskY As Long, ByVal MaskWidth As Long, ByVal MaskHeight As Long) As Long

If DestinationWidth = SpriteWidth And DestinationHeight = SpriteHeight Then
    
    DoBitBlt = BitBlt(destination.hDC, DestinationX, DestinationY, DestinationWidth, DestinationHeight, mask.hDC, MaskX, MaskY, dwRop.SRCAND)
    DoBitBlt = BitBlt(destination.hDC, DestinationX, DestinationY, DestinationWidth, DestinationHeight, Sprite.hDC, SpriteX, SpriteY, dwRop.SRCPAINT)

ElseIf DestinationWidth <> SpriteWidth Or DestinationHeight <> SpriteHeight Then
    
    DoBitBlt = StretchBlt(destination.hDC, DestinationX, DestinationY, DestinationWidth, DestinationHeight, mask.hDC, MaskX, MaskY, MaskWidth, MaskHeight, dwRop.SRCAND)
    DoBitBlt = StretchBlt(destination.hDC, DestinationX, DestinationY, DestinationWidth, DestinationHeight, Sprite.hDC, SpriteX, SpriteY, SpriteWidth, SpriteHeight, dwRop.SRCPAINT)

Else

    DoBitBlt = 0
    
End If

End Function


Function ReplaceOneString(FullString As String, ReplaceWhat As String, ReplaceWith As String)
'case sensitive
Dim SearchFor$, LeftString$, RightString$
SearchFor$ = InStr(FullString$, ReplaceWhat$)
If SearchFor$ = 0 Then MsgBox "String not found.": Exit Function
LeftString$ = Left(FullString$, SearchFor$ - 1)
RightString$ = Mid(FullString$, SearchFor$ + 1, Len(FullString$))
ReplaceOneString = LeftString$ + ReplaceWith$ + RightString$
End Function

Function ROSNCS(FullString As String, ReplaceWhat As String, ReplaceWith As String)
'not case sensitive
'ROSNCS = Replace One String Not Case Sensative
Dim SearchFor$, LeftString$, RightString$
SearchFor$ = InStr(UCase(FullString$), UCase(ReplaceWhat$))
If SearchFor$ = 0 Then MsgBox "String not found.": Exit Function
LeftString$ = Left(FullString$, SearchFor$ - 1)
RightString$ = Mid(FullString$, SearchFor$ + 1, Len(FullString$))
ROSNCS = LeftString$ + ReplaceWith$ + RightString$
End Function

Sub CreateNewStartButton()
Dim twnd, bwnd, ncwnd
    Dim r As RECT
    twnd = FindWindow("Shell_TrayWnd", vbNullString)
    bwnd = FindWindowEx(twnd, ByVal 0&, "BUTTON", vbNullString)
    GetWindowRect bwnd, r
    ncwnd = CreateWindowEx(ByVal 0&, "BUTTON", "Hello !", WS_CHILD, 0, 0, r.Right - r.Left, r.Bottom - r.Top, twnd, ByVal 0&, App.hInstance, ByVal 0&)
    ShowWindow ncwnd, SW_NORMAL
    ShowWindow bwnd, SW_HIDE
End Sub

Sub DestroyNewSB()
Dim bwnd, ncwnd
    ShowWindow bwnd, SW_NORMAL
    DestroyWindow ncwnd
End Sub


Function StripSpace(txt As String) As String
If InStr(txt$, " ") = 0 Then StripSpace$ = txt$: Exit Function
While InStr(txt$, " ")
txt$ = ReplaceOneString(txt$, " ", "")
DoEvents
Wend
StripSpace$ = txt$
End Function

Public Function ScreenWipe(Form As Form, CutSpeed As Integer) As Boolean
    Dim OldWidth As Integer
    Dim OldHeight As Integer
Form.WindowState = 0
If CutSpeed <= 0 Then
MsgBox "You cannot use 0 as a speed value"
Exit Function
End If
Do
OldWidth = Form.Width
Form.Width = Form.Width - CutSpeed
DoEvents
If Form.Width <> OldWidth Then
Form.Left = Form.Left + CutSpeed / 2
DoEvents
End If
OldHeight = Form.Height
Form.Height = Form.Height - CutSpeed
DoEvents
If Form.Height <> OldHeight Then
Form.Top = Form.Top + CutSpeed / 2
DoEvents
End If
Loop While Form.Width <> OldWidth Or Form.Height <> OldHeight
End Function

Public Function LineCount(thestring As String) As Long
Dim charcount$
charcount$ = InStr(thestring$, Chr(13))
If charcount$ <> 0 Then LineCount& = 1
Do
DoEvents
charcount$ = InStr(charcount$ + 1, thestring$, Chr(13))
If charcount$ <> 0 Then LineCount& = LineCount& + 1
DoEvents
Loop Until charcount$ = 0
LineCount& = LineCount& + 1
End Function

Public Function GetChatName() As String
GetChatName$ = GetCaption(FindChatRoom())
End Function

Public Function StripChatName()
StripChatName = StripSpace(GetChatName())
End Function

Public Function RoomBuster(Room As String, Optional Count As Label) As Long
stopbust = False
busted = False
Dim AOL&, MDI&, Keyword&, aoledit&, msgboxx&, chatroom&
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
redo:
If stopbust = True Then Exit Function
Call RunAOLToolbar("11", "G")
Do
DoEvents
Keyword& = FindWindowEx(MDI&, 0&, "AOL Child", "Keyword")
aoledit& = FindWindowEx(Keyword&, 0&, "_AOL_Edit", vbNullString)
Loop Until Keyword& <> 0 And aoledit& <> 0
Call SendMessageByString(aoledit&, WM_SETTEXT, 0, "aol://2719:2-2-" & Room$)
Call PostMessage(aoledit&, WM_CHAR, 13, 0)
Do
DoEvents
msgboxx& = FindWindow("#32770", "America Online")
chatroom& = FindChatRoom()
Loop Until msgboxx& <> 0 Or chatroom& <> 0 And UCase(StripChatName()) = UCase(Room$)
If msgboxx& <> 0 Then
Call SendMessage(msgboxx&, WM_CLOSE, 0, 0)
Count.Caption = Count.Caption + 1
GoTo redo
Exit Function
End If
If chatroom& <> 0 Then
busted = True
stopbust = True
Count.Caption = Count.Caption + 1
Exit Function
End If
End Function

Public Function LoadListboxRooms(list As ListBox, Directory As String) As Long
Dim a$
Open Directory$ For Input As #1
Do While Not EOF(1)
Input #1, a$
list.AddItem a$
DoEvents
Loop
Close #1
End Function

Public Function SaveListboxRooms(list As ListBox, Directory As String) As Long
Dim i
Open Directory$ For Output As #1
For i = 0 To list.ListCount - 1
Write #1, list.list(i)
Next i
Close #1
End Function
Sub BarFadeFrm(frm, style)
Dim cx, cy, f, F1, F2, i
frm.AutoRedraw = True
frm.Cls
frm.ScaleMode = 3
cx = frm.ScaleWidth / 2
cy = frm.ScaleHeight / 2
For i = 255 To 0 Step -2
f = i / 255
F1 = 1 - f: F2 = 1 + f
If style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
frm.Line (cx * F1, 0)-(cx * F2, cy * 2), , BF
Next i
End Sub
Sub CFadeFrm(frm, style)
frm.AutoRedraw = True
frm.Cls
Dim cx, cy, i
frm.ScaleMode = 3
cx = frm.ScaleWidth \ 2
cy = frm.ScaleHeight \ 2
frm.drawwidth = 2
For i = 0 To 255
If style = 1 Then frm.Circle (cx, cy), i, RGB(i, i, i)  'Black to white
If style = 2 Then frm.Circle (cx, cy), i, RGB(0, i, i)  'Black to Cyan
If style = 3 Then frm.Circle (cx, cy), i, RGB(i, 0, i)  'Black to Purple
If style = 4 Then frm.Circle (cx, cy), i, RGB(i, i, 0)  'Black to Yellow
If style = 5 Then frm.Circle (cx, cy), i, RGB(0, 0, i)  'Black to Blue
If style = 6 Then frm.Circle (cx, cy), i, RGB(i, 0, 0)  'Black to Red
If style = 7 Then frm.Circle (cx, cy), i, RGB(0, i, 0)  'Black to Green
If style = 8 Then frm.Circle (cx, cy), i, RGB(0, i, 255)  'Blue to Green
If style = 9 Then frm.Circle (cx, cy), i, RGB(i, i, 255)  'Blue to White
If style = 11 Then frm.Circle (cx, cy), i, RGB(i, 0, 255)  'Blue to Purple
If style = 12 Then frm.Circle (cx, cy), i, RGB(0, 0, 255 - i)  'Blue to Black
If style = 13 Then frm.Circle (cx, cy), i, RGB(255, 0, i)  'Red to Purple
If style = 14 Then frm.Circle (cx, cy), i, RGB(255, i, i)  'Red to White
If style = 15 Then frm.Circle (cx, cy), i, RGB(255, i, 0)  'Red to Yellow
If style = 16 Then frm.Circle (cx, cy), i, RGB(255 - i, 0, 0)  'Red to Black
If style = 17 Then frm.Circle (cx, cy), i, RGB(i, 255, i)  'Green to White
If style = 18 Then frm.Circle (cx, cy), i, RGB(0, 255, i)  'Green to Blue
If style = 19 Then frm.Circle (cx, cy), i, RGB(i, 255, 0)  'Green to Yellow
If style = 20 Then frm.Circle (cx, cy), i, RGB(0, 255 - i, 0)  'Green to Black
If style = 21 Then frm.Circle (cx, cy), i, RGB(255 - i, 255 - i, 255 - i)  'White to Black
If style = 22 Then frm.Circle (cx, cy), i, RGB(255, 255, 255 - i) 'White to Yellow
If style = 23 Then frm.Circle (cx, cy), i, RGB(255, 255 - i, 255) 'White to Purple
If style = 24 Then frm.Circle (cx, cy), i, RGB(255 - i, 255, 255) 'White to Cyan
If style = 25 Then frm.Circle (cx, cy), i, RGB(255, 255, i)  'Yellow to White
If style = 26 Then frm.Circle (cx, cy), i, RGB(255, i, 255)  'Purple to White
If style = 27 Then frm.Circle (cx, cy), i, RGB(i, 255, 255)  'Cyan to White
If style = 28 Then frm.Circle (cx, cy), i, RGB(255 - i, 255 - i, 0)  'Yellow to Black
If style = 29 Then frm.Circle (cx, cy), i, RGB(255 - i, 0, 255 - i)  'Purple to Black
If style = 30 Then frm.Circle (cx, cy), i, RGB(0, 255 - i, 255 - i)  'Cyan to Black
Dim s1, s2, s3
If style = 31 Then frm.Circle (cx, cy), i, RGB(s1 - i, s2 - i, s3 - i)  'Selected color to black
Next i
End Sub

Sub DoubleFade(frm, style)
frm.AutoRedraw = True
frm.Cls
Dim cx, cy, f, F1, F2, i
frm.AutoRedraw = True
frm.Cls
frm.ScaleMode = 3
cx = frm.ScaleWidth / 2
cy = frm.ScaleHeight / 2
Dim drawwidth
drawwidth = 2
For i = 255 To 0 Step -2
f = i / 255
F1 = 1 - f: F2 = 1 + f
If style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
frm.Line (cx * F1, cy * F1)-(cx * F2, cy * F2), , BF
Next i
frm.ScaleMode = 3   ' Set ScaleMode to pixels.
cx = frm.ScaleWidth / 2 ' Get horizontal center.
cy = frm.ScaleHeight / 2    ' Get vertical center.
frm.drawwidth = 2
For i = 0 To 255
If style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
f = i / 255  ' Perform interim
F1 = 1 - f: F2 = 1 + f  ' calculations.
frm.Line (cx * F1, cy)-(cx, cy * F1)   ' Draw upper-left.
frm.Line -(cx * F2, cy) ' Draw upper-right.
frm.Line -(cx, cy * F2) ' Draw lower-right.
frm.Line -(cx * F1, cy) ' Draw lower-left.
Next i
End Sub
Sub ExplosiveFade(frm, style)
frm.AutoRedraw = True
frm.Cls
Dim cx, cy, f, F1, F2, i
frm.ScaleMode = 3
cx = frm.ScaleWidth / 2
cy = frm.ScaleHeight / 2
frm.drawwidth = 2
For i = 0 To 255
If style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
f = i / 255  ' Perform interim
F1 = 1 - f: F2 = 1 + f  ' calculations.
frm.Line (cx * F1, cy)-(cx, cy * F1)   ' Draw upper-left.
frm.Line -(cx * F2, cy) ' Draw upper-right.
frm.Line -(cx, cy * F2) ' Draw lower-right.
frm.Line -(cx * F1, cy) ' Draw lower-left.
Next i
End Sub
Sub FadeFrm(frm, style)
frm.ScaleMode = vbPixels
frm.AutoRedraw = True
frm.DrawStyle = vbInsideSolid
frm.Cls
frm.drawwidth = 2
frm.DrawMode = 13
frm.ScaleHeight = 256
Dim i
For i = 0 To 255
If style = 1 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, i), BF ' Black to white
If style = 2 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, i, i), BF ' Black to Cyan
If style = 3 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 0, i), BF ' Black to Purple
If style = 4 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, 0), BF ' Black to Yellow
If style = 5 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 0, i), BF ' Black to Blue
If style = 6 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 0, 0), BF ' Black to Red
If style = 7 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, i, 0), BF ' Black to Green
If style = 8 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, i, 255), BF ' Blue to Green
If style = 9 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, 255), BF ' Blue to White
If style = 11 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 0, 255), BF ' Blue to Purple
If style = 12 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 0, 255 - i), BF ' Blue to Black
If style = 13 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 0, i), BF ' Red to Purple
If style = 14 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, i, i), BF ' Red to White
If style = 15 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, i, 0), BF ' Red to Yellow
If style = 16 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 0, 0), BF ' Red to Black
If style = 17 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 255, i), BF ' Green to White
If style = 18 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 255, i), BF ' Green to Blue
If style = 19 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 255, 0), BF ' Green to Yellow
If style = 20 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 255 - i, 0), BF ' Green to Black
If style = 21 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 255 - i, 255 - i), BF ' White to Black
If style = 22 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 255, 255 - i), BF 'White to Yellow
If style = 23 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 255 - i, 255), BF 'White to Purple
If style = 24 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 255, 255), BF 'White to Cyan
If style = 25 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 255, i), BF ' Yellow to White
If style = 26 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, i, 255), BF ' Purple to White
If style = 27 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 255, 255), BF ' Cyan to White
If style = 28 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 255 - i, 0), BF ' Yellow to Black
If style = 29 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 0, 255 - i), BF ' Purple to Black
If style = 30 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 255 - i, 255 - i), BF ' Cyan to Black
If style = 31 Then If i = 193 Then Exit Sub: frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, i), BF ' black to Gray
Next i
End Sub
Sub RFadeFrm(frm, style)
Dim cx, cy, f, F1, F2, i
frm.AutoRedraw = True
frm.Cls
frm.ScaleMode = 3
cx = frm.ScaleWidth / 2
cy = frm.ScaleHeight / 2
For i = 255 To 0 Step -2
f = i / 255
F1 = 1 - f: F2 = 1 + f
If style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
frm.Line (cx * F1, cy * F1)-(cx * F2, cy * F2), , BF
Next i
End Sub
Sub SideFade(frm, style)
Dim drawwidth
Dim cx, cy, f, F1, F2, i
frm.AutoRedraw = True
frm.Cls
frm.ScaleMode = 3
cx = frm.ScaleWidth
cy = frm.ScaleHeight
drawwidth = 2
For i = 255 To 0 Step -2
f = i / 255
F1 = 1 - f: F2 = 1 + f
If style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
frm.Line (cx * F1, 0)-(cx * F2, cy * 2), , BF
Next i
End Sub
Sub Text3D(Ctrl As Control, Text, bevel, style, Font)
Ctrl.AutoRedraw = True
Ctrl.FontSize = bevel * 1.4
Ctrl.Font = Font
Dim i
For i = 0 To bevel * 10
If style = 1 Then Ctrl.ForeColor = RGB(i, i, i) ' Black to white
If style = 2 Then Ctrl.ForeColor = RGB(0, i, i) ' Black to Cyan
If style = 3 Then Ctrl.ForeColor = RGB(i, 0, i) ' Black to Purple
If style = 4 Then Ctrl.ForeColor = RGB(i, i, 0) ' Black to Yellow
If style = 5 Then Ctrl.ForeColor = RGB(0, 0, i) ' Black to Blue
If style = 6 Then Ctrl.ForeColor = RGB(i, 0, 0) ' Black to Red
If style = 7 Then Ctrl.ForeColor = RGB(0, i, 0) ' Black to Green
If style = 8 Then Ctrl.ForeColor = RGB(0, i, 255) ' Blue to Green
If style = 9 Then Ctrl.ForeColor = RGB(i, i, 255) ' Blue to White
If style = 11 Then Ctrl.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If style = 12 Then Ctrl.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If style = 13 Then Ctrl.ForeColor = RGB(255, 0, i) ' Red to Purple
If style = 14 Then Ctrl.ForeColor = RGB(255, i, i) ' Red to White
If style = 15 Then Ctrl.ForeColor = RGB(255, i, 0) ' Red to Yellow
If style = 16 Then Ctrl.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If style = 17 Then Ctrl.ForeColor = RGB(i, 255, i) ' Green to White
If style = 18 Then Ctrl.ForeColor = RGB(0, 255, i) ' Green to Blue
If style = 19 Then Ctrl.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If style = 20 Then Ctrl.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If style = 21 Then Ctrl.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If style = 22 Then Ctrl.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If style = 23 Then Ctrl.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If style = 24 Then Ctrl.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If style = 25 Then Ctrl.ForeColor = RGB(255, 255, i) ' Yellow to White
If style = 26 Then Ctrl.ForeColor = RGB(255, i, 255) ' Purple to White
If style = 27 Then Ctrl.ForeColor = RGB(i, 255, 255) ' Cyan to White
If style = 28 Then Ctrl.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If style = 29 Then Ctrl.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If style = 30 Then Ctrl.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If style = 31 Then Ctrl.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
Ctrl.CurrentY = i \ 2
Ctrl.CurrentX = i \ 2
Ctrl.Print Text
Next i
End Sub

Function RGBtoHEX(RGB)
Dim a$, Length
    a$ = Hex(RGB)
    Length = Len(a$)
    If Length = 5 Then a$ = "0" & a$
    If Length = 4 Then a$ = "00" & a$
    If Length = 3 Then a$ = "000" & a$
    If Length = 2 Then a$ = "0000" & a$
    If Length = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function


Public Function IMessage(who As String, message As String)
Call RunAOLToolbar("3", "I")
Do
DoEvents
Dim AOL&, MDI&, IM&, aoledit&, richcntl&, AOLIcon&
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindChildByClass(AOL&, "MDIClient")
IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
aoledit& = FindWindowEx(IM&, 0&, "_AOL_Edit", vbNullString)
richcntl& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
Loop Until richcntl& <> 0
Call SendMessageByString(aoledit&, WM_SETTEXT, 0, who$)
Call SendMessageByString(richcntl&, WM_SETTEXT, 0, message$)
AOLIcon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(IM&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(IM&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(IM&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(IM&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(IM&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(IM&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(IM&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(IM&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(IM&, AOLIcon&, "_AOL_Icon", vbNullString)
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(AOLIcon&, WM_KEYUP, VK_SPACE, 0)
Dim mesbox&
Do
DoEvents
mesbox& = FindWindow("#32770", "America Online Error")
Loop Until mesbox& <> 0 Or IM& = 0
If mesbox& <> 0 Then
Dim staticc&, txt$
staticc& = FindChildByClass(mesbox&, "Static")
txt$ = GetCaption(staticc&)
Call SendMessage(mesbox&, WM_CLOSE, 0, 0)
Call SendMessage(IM&, WM_CLOSE, 0, 0)
Exit Function
End If
End Function
Public Function DecryptWithALP(strdata As String) As String
    Dim strALPKey As String
    Dim strALPKeyMask As String
    Dim lngIterator As Long
    Dim blnOscillator As Boolean
    Dim strOutput As String
    Dim lngHex As Long
    If Len(strdata) = 0 Then
        Exit Function
    End If
    strALPKeyMask = Right$(String$(lngALPKeyLength, "0") + DoubleToBinary(CLng("&H" + Left$(strdata, 2))), lngALPKeyLength)
    strdata = Right$(strdata, Len(strdata) - 2)
    For lngIterator = lngALPKeyLength To 1 Step -1
        If Mid$(strALPKeyMask, lngIterator, 1) = "1" Then
            strALPKey = Left$(strdata, 1) + strALPKey
            strdata = Right$(strdata, Len(strdata) - 1)
        Else
            strALPKey = Right$(strdata, 1) + strALPKey
            strdata = Left$(strdata, Len(strdata) - 1)
        End If
    Next lngIterator
    lngIterator = 0
    Do Until Len(strdata) = 0
        blnOscillator = Not blnOscillator
        lngIterator = lngIterator + 1
        If lngIterator > lngALPKeyLength Then
            lngIterator = 1
        End If
        lngHex = IIf(blnOscillator, CLng("&H" + Left$(strdata, 2) - Asc(Mid$(strALPKey, lngIterator, 1))), CLng("&H" + Left$(strdata, 2) + Asc(Mid$(strALPKey, lngIterator, 1))))
        If lngHex > 255 Then
            lngHex = lngHex - 255
        ElseIf lngHex < 0 Then
            lngHex = lngHex + 255
        End If
        strOutput = strOutput + Chr$(lngHex)
        strdata = Right$(strdata, Len(strdata) - 2)
    Loop
    DecryptWithALP = strOutput
End Function
Public Function DecryptWithClipper(ByVal strdata As String, ByVal strcryptkey As String) As String
'Call EncryptWithClipper("hi", "password")
'Call DecryptWithClipper("hi", "password")
    Dim strDecryptionChunk As String
    Dim strDecryptedText As String
    On Error Resume Next
    InitCrypt strcryptkey
    Do Until Len(strdata) < 16
        strDecryptionChunk = ""
        strDecryptionChunk = Left$(strdata, 16)
        strdata = Right$(strdata, Len(strdata) - 16)
        If Len(strDecryptionChunk) > 0 Then
            strDecryptedText = strDecryptedText + PerformClipperDecryption(strDecryptionChunk)
        End If
    Loop
    DecryptWithClipper = strDecryptedText
End Function
Public Function DecryptWithCSP(ByVal strdata As String, ByVal strcryptkey As String) As String
    Dim lngEncryptionCount As Long
    Dim strDecrypted As String
    Dim strCurrentCryptKey As String
    If EncryptionCSPConnect() Then
        lngEncryptionCount = DecryptNumber(Mid$(strdata, 1, 8))
        strCurrentCryptKey = strcryptkey & lngEncryptionCount
        strDecrypted = EncryptDecrypt(Mid$(strdata, 9), strCurrentCryptKey, False)
        DecryptWithCSP = strDecrypted
        EncryptionCSPDisconnect
    End If
End Function
Public Function EncryptWithALP(strdata As String) As String
    Dim strALPKey As String
    Dim strALPKeyMask As String
    Dim lngIterator As Long
    Dim blnOscillator As Boolean
    Dim strOutput As String
    Dim lngHex As Long
    If Len(strdata) = 0 Then
        Exit Function
    End If
    Randomize
    For lngIterator = 1 To lngALPKeyLength
        strALPKey = strALPKey + Trim$(Hex$(Int(16 * Rnd)))
        strALPKeyMask = strALPKeyMask + Trim$(Int(2 * Rnd))
    Next lngIterator
    lngIterator = 0
    Do Until Len(strdata) = 0
        blnOscillator = Not blnOscillator
        lngIterator = lngIterator + 1
        If lngIterator > lngALPKeyLength Then
            lngIterator = 1
        End If
        lngHex = IIf(blnOscillator, CLng(Asc(Left$(strdata, 1)) + Asc(Mid$(strALPKey, lngIterator, 1))), CLng(Asc(Left$(strdata, 1)) - Asc(Mid$(strALPKey, lngIterator, 1))))
        If lngHex > 255 Then
            lngHex = lngHex - 255
        ElseIf lngHex < 0 Then
            lngHex = lngHex + 255
        End If
        strOutput = strOutput + Right$(String$(2, "0") + Hex$(lngHex), 2)
        strdata = Right$(strdata, Len(strdata) - 1)
    Loop
    For lngIterator = 1 To lngALPKeyLength
        If Mid$(strALPKeyMask, lngIterator, 1) = "1" Then
            strOutput = Mid$(strALPKey, lngIterator, 1) + strOutput
        Else
            strOutput = strOutput + Mid$(strALPKey, lngIterator, 1)
        End If
    Next lngIterator
    EncryptWithALP = Right$(String$(2, "0") + Hex$(BinaryToDouble(strALPKeyMask)), 2) + strOutput
End Function
Public Function EncryptWithClipper(ByVal strdata As String, ByVal strcryptkey As String) As String
'Call EncryptWithClipper("hi", "password")
'Call DecryptWithClipper("hi", "password")
    Dim strEncryptionChunk As String
    Dim strEncryptedText As String
    If Len(strdata) > 0 Then
        InitCrypt strcryptkey
        Do Until Len(strdata) = 0
            strEncryptionChunk = ""
            If Len(strdata) > 6 Then
                strEncryptionChunk = Left$(strdata, 6)
                strdata = Right$(strdata, Len(strdata) - 6)
            Else
                strEncryptionChunk = Left$(strdata + Space(6), 6)
                strdata = ""
            End If
            If Len(strEncryptionChunk) > 0 Then
                strEncryptedText = strEncryptedText + PerformClipperEncryption(strEncryptionChunk)
            End If
        Loop
    End If
    EncryptWithClipper = strEncryptedText
End Function
Public Function EncryptWithCSP(ByVal strdata As String, ByVal strcryptkey As String) As String
    Dim strEncrypted As String
    Dim lngEncryptionCount As Long
    Dim strCurrentCryptKey As String
    If EncryptionCSPConnect() Then
        lngEncryptionCount = 0
        strCurrentCryptKey = strcryptkey & lngEncryptionCount
        strEncrypted = EncryptDecrypt(strdata, strCurrentCryptKey, True)
        Do While (InStr(1, strEncrypted, vbCr) > 0) Or (InStr(1, strEncrypted, vbLf) > 0) Or (InStr(1, strEncrypted, Chr$(0)) > 0) Or (InStr(1, strEncrypted, vbTab) > 0)
            lngEncryptionCount = lngEncryptionCount + 1
            strCurrentCryptKey = strcryptkey & lngEncryptionCount
            strEncrypted = EncryptDecrypt(strdata, strCurrentCryptKey, True)
            If lngEncryptionCount = 99999999 Then
                Err.Raise vbObjectError + 999, "EncryptWithCSP", "This Data cannot be successfully encrypted"
                EncryptWithCSP = ""
                Exit Function
            End If
        Loop
        EncryptWithCSP = EncryptNumber(lngEncryptionCount) & strEncrypted
        EncryptionCSPDisconnect
    End If
End Function
Public Function GetCSPDetails() As String
    Dim lngDataLength As Long
    Dim bytContainer() As Byte
    If EncryptionCSPConnect Then
        If lngCryptProvider = 0 Then
            GetCSPDetails = "Not connected to CSP"
            Exit Function
        End If
        lngDataLength = 1000
        ReDim bytContainer(lngDataLength)
        If CryptGetProvParam(lngCryptProvider, PP_NAME, bytContainer(0), lngDataLength, 0) <> 0 Then
            GetCSPDetails = "Cryptographic Service Provider name: " & ByteToString(bytContainer, lngDataLength)
        End If
        lngDataLength = 1000
        ReDim bytContainer(lngDataLength)
        If CryptGetProvParam(lngCryptProvider, PP_CONTAINER, bytContainer(0), lngDataLength, 0) <> 0 Then
            GetCSPDetails = GetCSPDetails & vbCrLf & "Key Container name: " & ByteToString(bytContainer, lngDataLength)
        End If
        EncryptionCSPDisconnect
    Else
        GetCSPDetails = "Not connected to CSP"
    End If
End Function
Public Function DecryptNumber(ByVal strdata As String) As Long
    Dim lngIterator As Long
    For lngIterator = 1 To 8
        DecryptNumber = (10 * DecryptNumber) + (Asc(Mid$(strdata, lngIterator, 1)) - Asc(Mid$(ENCRYPT_NUMBERKEY, lngIterator, 1)))
    Next lngIterator
End Function
Public Function EncryptDecrypt(ByVal strdata As String, ByVal strcryptkey As String, ByVal Encrypt As Boolean) As String
    Dim lngDataLength As Long
    Dim strTempData As String
    Dim lngHaslngCryptKey As Long
    Dim lngCryptKey As Long
    If lngCryptProvider = 0 Then
        'Err.Raise vbObjectError + 999, "EncryptDecrypt", "Not connected to CSP"
        Exit Function
    End If
    If CryptCreateHash(lngCryptProvider, CALG_MD5, 0, 0, lngHaslngCryptKey) = 0 Then
        Err.Raise vbObjectError + 999, "EncryptDecrypt", "Error during CryptCreateHash."
    End If
    If CryptHashData(lngHaslngCryptKey, strcryptkey, Len(strcryptkey), 0) = 0 Then
        Err.Raise vbObjectError + 999, "EncryptDecrypt", "Error during CryptHashData."
    End If
    If CryptDeriveKey(lngCryptProvider, ENCRYPT_ALGORITHM, lngHaslngCryptKey, 0, lngCryptKey) = 0 Then
        Err.Raise vbObjectError + 999, "EncryptDecrypt", "Error during CryptDeriveKey!"
    End If
    strTempData = strdata
    lngDataLength = Len(strdata)
    If Encrypt Then
        If CryptEncrypt(lngCryptKey, 0, 1, 0, strTempData, lngDataLength, lngDataLength) = 0 Then
            Err.Raise vbObjectError + 999, "EncryptDecrypt", "Error during CryptEncrypt."
        End If
    Else
        If CryptDecrypt(lngCryptKey, 0, 1, 0, strTempData, lngDataLength) = 0 Then
            Err.Raise vbObjectError + 999, "EncryptDecrypt", "Error during CryptDecrypt."
        End If
    End If
    EncryptDecrypt = Mid$(strTempData, 1, lngDataLength)
    If lngCryptKey <> 0 Then
        CryptDestroyKey lngCryptKey
    End If
    If lngHaslngCryptKey <> 0 Then
        CryptDestroyHash lngHaslngCryptKey
    End If
End Function
Public Function EncryptionCSPConnect() As Boolean
    If Len(strKeyContainer) = 0 Then
        strKeyContainer = "FastTrack"
    End If
    If CryptAcquireContext(lngCryptProvider, strKeyContainer, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0 Then
        If CryptAcquireContext(lngCryptProvider, strKeyContainer, SERVICE_PROVIDER, PROV_RSA_FULL, 0) = 0 Then
            Err.Raise vbObjectError + 999, "EncryptionCSPConnect", "Error during CryptAcquireContext for a new key container." & vbCrLf & "A container with this name probably already exists."
            EncryptionCSPConnect = False
            Exit Function
        End If
    End If
    EncryptionCSPConnect = True
End Function
Public Function EncryptNumber(ByVal lngData As Long) As String
    Dim lngIterator As Long
    Dim strdata As String
    strdata = Format$(lngData, "00000000")
    For lngIterator = 1 To 8
        EncryptNumber = EncryptNumber & Chr$(Asc(Mid$(ENCRYPT_NUMBERKEY, lngIterator, 1)) + Val(Mid$(strdata, lngIterator, 1)))
    Next lngIterator
End Function
Public Sub EncryptionCSPDisconnect()
    If lngCryptProvider <> 0 Then
        CryptReleaseContext lngCryptProvider, 0
    End If
End Sub
Public Sub InitCrypt(ByRef strEncryptionKey As String)
    avarSeedValues = Array("A3", "D7", "09", "83", "F8", "48", "F6", "F4", "B3", "21", "15", "78", "99", "B1", "AF", _
    "F9", "E7", "2D", "4D", "8A", "CE", "4C", "CA", "2E", "52", "95", "D9", "1E", "4E", "38", "44", "28", "0A", "DF", _
    "02", "A0", "17", "F1", "60", "68", "12", "B7", "7A", "C3", "E9", "FA", "3D", "53", "96", "84", "6B", "BA", "F2", _
    "63", "9A", "19", "7C", "AE", "E5", "F5", "F7", "16", "6A", "A2", "39", "B6", "7B", "0F", "C1", "93", "81", "1B", _
    "EE", "B4", "1A", "EA", "D0", "91", "2F", "B8", "55", "B9", "DA", "85", "3F", "41", "BF", "E0", "5A", "58", "80", _
    "5F", "66", "0B", "D8", "90", "35", "D5", "C0", "A7", "33", "06", "65", "69", "45", "00", "94", "56", "6D", "98", _
    "9B", "76", "97", "FC", "B2", "C2", "B0", "FE", "DB", "20", "E1", "EB", "D6", "E4", "DD", "47", "4A", "1D", "42", _
    "ED", "9E", "6E", "49", "3C", "CD", "43", "27", "D2", "07", "D4", "DE", "C7", "67", "18", "89", "CB", "30", "1F", _
    "8D", "C6", "8F", "AA", "C8", "74", "DC", "C9", "5D", "5C", "31", "A4", "70", "88", "61", "2C", "9F", "0D", "2B", _
    "87", "50", "82", "54", "64", "26", "7D", "03", "40", "34", "4B", "1C", "73", "D1", "C4", "FD", "3B", "CC", "FB", _
    "7F", "AB", "E6", "3E", "5B", "A5", "AD", "04", "23", "9C", "14", "51", "22", "F0", "29", "79", "71", "7E", "FF", _
    "8C", "0E", "E2", "0C", "EF", "BC", "72", "75", "6F", "37", "A1", "EC", "D3", "8E", "62", "8B", "86", "10", "E8", _
    "08", "77", "11", "BE", "92", "4F", "24", "C5", "32", "36", "9D", "CF", "F3", "A6", "BB", "AC", "5E", "6C", "A9", _
    "13", "57", "25", "B5", "E3", "BD", "A8", "3A", "01", "05", "59", "2A", "46")
    SetKey strEncryptionKey
End Sub
Public Function PerformClipperDecryption(ByVal strdata As String) As String
    Dim bytChunk(1 To 4, 0 To 32) As String
    Dim bytCounter(0 To 32) As Byte
    Dim lngIterator As Long
    Dim strDecryptedData As String
    On Error Resume Next
    bytChunk(1, 32) = Mid(strdata, 1, 4)
    bytChunk(2, 32) = Mid(strdata, 5, 4)
    bytChunk(3, 32) = Mid(strdata, 9, 4)
    bytChunk(4, 32) = Mid(strdata, 13, 4)
    lngSeedLevel = 32
    lngDecryptPointer = 31
    For lngIterator = 0 To 32
        bytCounter(lngIterator) = lngIterator + 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel - 1) = PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey())
        bytChunk(2, lngSeedLevel - 1) = PerformXOR(PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey()), PerformXOR(bytChunk(3, lngSeedLevel), Hex(bytCounter(lngSeedLevel - 1))))
        bytChunk(3, lngSeedLevel - 1) = bytChunk(4, lngSeedLevel)
        bytChunk(4, lngSeedLevel - 1) = bytChunk(1, lngSeedLevel)
        lngDecryptPointer = lngDecryptPointer - 1
        lngSeedLevel = lngSeedLevel - 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel - 1) = PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey())
        bytChunk(2, lngSeedLevel - 1) = bytChunk(3, lngSeedLevel)
        bytChunk(3, lngSeedLevel - 1) = bytChunk(4, lngSeedLevel)
        bytChunk(4, lngSeedLevel - 1) = PerformXOR(PerformXOR(bytChunk(1, lngSeedLevel), bytChunk(2, lngSeedLevel)), Hex(bytCounter(lngSeedLevel - 1)))
        lngDecryptPointer = lngDecryptPointer - 1
        lngSeedLevel = lngSeedLevel - 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel - 1) = PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey())
        bytChunk(2, lngSeedLevel - 1) = PerformXOR(PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey()), PerformXOR(bytChunk(3, lngSeedLevel), Hex(bytCounter(lngSeedLevel - 1))))
        bytChunk(3, lngSeedLevel - 1) = bytChunk(4, lngSeedLevel)
        bytChunk(4, lngSeedLevel - 1) = bytChunk(1, lngSeedLevel)
        lngDecryptPointer = lngDecryptPointer - 1
        lngSeedLevel = lngSeedLevel - 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel - 1) = PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey())
        bytChunk(2, lngSeedLevel - 1) = bytChunk(3, lngSeedLevel)
        bytChunk(3, lngSeedLevel - 1) = bytChunk(4, lngSeedLevel)
        bytChunk(4, lngSeedLevel - 1) = PerformXOR(PerformXOR(bytChunk(1, lngSeedLevel), bytChunk(2, lngSeedLevel)), Hex(bytCounter(lngSeedLevel - 1)))
        lngDecryptPointer = lngDecryptPointer - 1
        lngSeedLevel = lngSeedLevel - 1
    Next lngIterator
    strDecryptedData = HexToString(bytChunk(1, 0) & bytChunk(2, 0) & bytChunk(3, 0) & bytChunk(4, 0))
    If InStr(strDecryptedData, Chr$(0)) > 0 Then
        strDecryptedData = Left$(strDecryptedData, InStr(strDecryptedData, Chr$(0)) - 1)
    End If
    PerformClipperDecryption = strDecryptedData
End Function
Public Function PerformClipperDecryptionChunk(ByVal strdata As String, ByRef strEncryptionKey() As String) As String
    Dim astrDecryptionLevel(1 To 6) As String
    Dim strDecryptedString As String
    astrDecryptionLevel(5) = Mid(strdata, 1, 2)
    astrDecryptionLevel(6) = Mid(strdata, 3, 2)
    strDecryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrDecryptionLevel(5), strEncryptionKey((4 * lngDecryptPointer) + 3)))))
    astrDecryptionLevel(4) = PerformXOR(strDecryptedString, astrDecryptionLevel(6))
    strDecryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrDecryptionLevel(4), strEncryptionKey((4 * lngDecryptPointer) + 2)))))
    astrDecryptionLevel(3) = PerformXOR(strDecryptedString, astrDecryptionLevel(5))
    strDecryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrDecryptionLevel(3), strEncryptionKey((4 * lngDecryptPointer) + 1)))))
    astrDecryptionLevel(2) = PerformXOR(strDecryptedString, astrDecryptionLevel(4))
    strDecryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrDecryptionLevel(2), strEncryptionKey(4 * lngDecryptPointer)))))
    astrDecryptionLevel(1) = PerformXOR(strDecryptedString, astrDecryptionLevel(3))
    strDecryptedString = astrDecryptionLevel(1) & astrDecryptionLevel(2)
    PerformClipperDecryptionChunk = strDecryptedString
End Function
Public Function PerformClipperEncryption(ByVal strdata As String) As String
    Dim bytChunk(1 To 4, 0 To 32) As String
    Dim lngCounter As Long
    Dim lngIterator As Long
    On Error Resume Next
    strdata = StringToHex(strdata)
    bytChunk(1, 0) = Mid(strdata, 1, 4)
    bytChunk(2, 0) = Mid(strdata, 5, 4)
    bytChunk(3, 0) = Mid(strdata, 9, 4)
    bytChunk(4, 0) = Mid(strdata, 13, 4)
    lngSeedLevel = 0
    lngCounter = 1
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel + 1) = PerformXOR(PerformXOR(PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey()), bytChunk(4, lngSeedLevel)), Hex(lngCounter))
        bytChunk(2, lngSeedLevel + 1) = PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey())
        bytChunk(3, lngSeedLevel + 1) = bytChunk(2, lngSeedLevel)
        bytChunk(4, lngSeedLevel + 1) = bytChunk(3, lngSeedLevel)
        lngCounter = lngCounter + 1
        lngSeedLevel = lngSeedLevel + 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel + 1) = bytChunk(4, lngSeedLevel)
        bytChunk(2, lngSeedLevel + 1) = PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey())
        bytChunk(3, lngSeedLevel + 1) = PerformXOR(PerformXOR(bytChunk(1, lngSeedLevel), bytChunk(2, lngSeedLevel)), Hex(lngCounter))
        bytChunk(4, lngSeedLevel + 1) = bytChunk(3, lngSeedLevel)
        lngCounter = lngCounter + 1
        lngSeedLevel = lngSeedLevel + 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel + 1) = PerformXOR(PerformXOR(PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey()), bytChunk(4, lngSeedLevel)), Hex(lngCounter))
        bytChunk(2, lngSeedLevel + 1) = PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey())
        bytChunk(3, lngSeedLevel + 1) = bytChunk(2, lngSeedLevel)
        bytChunk(4, lngSeedLevel + 1) = bytChunk(3, lngSeedLevel)
        lngCounter = lngCounter + 1
        lngSeedLevel = lngSeedLevel + 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel + 1) = bytChunk(4, lngSeedLevel)
        bytChunk(2, lngSeedLevel + 1) = PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey())
        bytChunk(3, lngSeedLevel + 1) = PerformXOR(PerformXOR(bytChunk(1, lngSeedLevel), bytChunk(2, lngSeedLevel)), Hex(lngCounter))
        bytChunk(4, lngSeedLevel + 1) = bytChunk(3, lngSeedLevel)
        lngCounter = lngCounter + 1
        lngSeedLevel = lngSeedLevel + 1
    Next lngIterator
    PerformClipperEncryption = bytChunk(1, 32) & bytChunk(2, 32) & bytChunk(3, 32) & bytChunk(4, 32)
End Function
Public Function PerformClipperEncryptionChunk(ByVal strdata As String, ByRef strEncryptionKey() As String) As String
    Dim astrEncryptionLevel(1 To 6) As String
    Dim strEncryptedString As String
    astrEncryptionLevel(1) = Mid(strdata, 1, 2)
    astrEncryptionLevel(2) = Mid(strdata, 3, 2)
    strEncryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrEncryptionLevel(2), strEncryptionKey(4 * lngSeedLevel)))))
    astrEncryptionLevel(3) = PerformXOR(strEncryptedString, astrEncryptionLevel(1))
    strEncryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrEncryptionLevel(3), strEncryptionKey((4 * lngSeedLevel) + 1)))))
    astrEncryptionLevel(4) = PerformXOR(strEncryptedString, astrEncryptionLevel(2))
    strEncryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrEncryptionLevel(4), strEncryptionKey((4 * lngSeedLevel) + 2)))))
    astrEncryptionLevel(5) = PerformXOR(strEncryptedString, astrEncryptionLevel(3))
    strEncryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrEncryptionLevel(5), strEncryptionKey((4 * lngSeedLevel) + 3)))))
    astrEncryptionLevel(6) = PerformXOR(strEncryptedString, astrEncryptionLevel(4))
    strEncryptedString = astrEncryptionLevel(5) & astrEncryptionLevel(6)
    PerformClipperEncryptionChunk = strEncryptedString
End Function
Public Function PerformTranslation(ByVal strdata As String) As Double
    Dim strTranslationString As String
    Dim strTranslationChunk As String
    Dim lngTranslationIterator As Long
    Dim lngHexConversion As Long
    Dim lngHexConversionIterator As Long
    Dim dblTranslation As Double
    Dim lngTranslationMarker As Long
    Dim lngTranslationModifier As Long
    Dim lngTranslationLayerModifier As Long
    strTranslationString = strdata
    strTranslationString = Right$(strTranslationString, 8)
    strTranslationChunk = String$(8 - Len(strTranslationString), "0") + strTranslationString
    strTranslationString = ""
    For lngTranslationIterator = 1 To 8
        lngHexConversion = Val("&H" + Mid$(strTranslationChunk, lngTranslationIterator, 1))
        For lngHexConversionIterator = 3 To 0 Step -1
            If lngHexConversion And 2 ^ lngHexConversionIterator Then
                strTranslationString = strTranslationString + "1"
            Else
                strTranslationString = strTranslationString + "0"
            End If
        Next lngHexConversionIterator
    Next lngTranslationIterator
    dblTranslation = 0
    For lngTranslationIterator = Len(strTranslationString) To 1 Step -1
        If Mid(strTranslationString, lngTranslationIterator, 1) = "1" Then
            lngTranslationLayerModifier = 1
            lngTranslationMarker = (Len(strTranslationString) - lngTranslationIterator)
            lngTranslationModifier = 2
            Do While lngTranslationMarker > 0
                Do While (lngTranslationMarker / 2) = (lngTranslationMarker \ 2)
                    lngTranslationModifier = (lngTranslationModifier * lngTranslationModifier) Mod 255
                    lngTranslationMarker = lngTranslationMarker / 2
                Loop
                lngTranslationLayerModifier = (lngTranslationModifier * lngTranslationLayerModifier) Mod 255
                lngTranslationMarker = lngTranslationMarker - 1
            Loop
            dblTranslation = dblTranslation + lngTranslationLayerModifier
        End If
    Next lngTranslationIterator
    PerformTranslation = dblTranslation
End Function
Public Function PerformXOR(ByVal strdata As String, ByVal strMask As String) As String
    Dim strXOR As String
    Dim lngXORIterator As Long
    Dim lngXORMarker As Long
    lngXORMarker = Len(strdata) - Len(strMask)
    If lngXORMarker < 0 Then
        strXOR = Left$(strMask, Abs(lngXORMarker))
        strMask = Mid$(strMask, Abs(lngXORMarker) + 1)
    ElseIf lngXORMarker > 0 Then
        strXOR = Left$(strdata, Abs(lngXORMarker))
        strdata = Mid$(strdata, lngXORMarker + 1)
    End If
    For lngXORIterator = 1 To Len(strdata)
        strXOR = strXOR + Hex$(Val("&H" + Mid$(strdata, lngXORIterator, 1)) Xor Val("&H" + Mid$(strMask, lngXORIterator, 1)))
    Next lngXORIterator
    PerformXOR = Right(strXOR, 8)
End Function
Public Sub SetKey(ByVal strEncryptionKey As String)
    Dim intEncryptionKeyIterator As Integer
    For intEncryptionKeyIterator = 0 To 131 Step 10
        If intEncryptionKeyIterator = 130 Then
            astrEncryptionKey(intEncryptionKeyIterator + 0) = Mid(strEncryptionKey, 1, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 1) = Mid(strEncryptionKey, 3, 2)
        Else
            astrEncryptionKey(intEncryptionKeyIterator + 0) = Mid(strEncryptionKey, 1, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 1) = Mid(strEncryptionKey, 3, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 2) = Mid(strEncryptionKey, 5, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 3) = Mid(strEncryptionKey, 7, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 4) = Mid(strEncryptionKey, 9, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 5) = Mid(strEncryptionKey, 11, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 6) = Mid(strEncryptionKey, 13, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 7) = Mid(strEncryptionKey, 15, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 8) = Mid(strEncryptionKey, 17, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 9) = Mid(strEncryptionKey, 19, 2)
        End If
    Next
End Sub


Public Function BinaryToDouble(ByVal strdata As String) As Double
    Dim dblOutput As Double
    Dim lngIterator As Long
    Do Until Len(strdata) = 0
        dblOutput = dblOutput + IIf(Right$(strdata, 1) = "1", (2 ^ lngIterator), 0)
        strdata = Left$(strdata, Len(strdata) - 1)
        lngIterator = lngIterator + 1
    Loop
    BinaryToDouble = dblOutput
End Function

Public Function DoubleToBinary(ByVal dblData As Double) As String
    Dim strOutput As String
    Dim lngIterator As Long
    Do Until (2 ^ lngIterator) > dblData
        strOutput = IIf(((2 ^ lngIterator) And dblData) > 0, "1", "0") + strOutput
        lngIterator = lngIterator + 1
    Loop
    DoubleToBinary = strOutput
End Function
Public Function HexToString(ByVal strdata As String) As String
    Dim strOutput As String
    Do Until Len(strdata) < 2
        strOutput$ = strOutput$ + Chr$(CLng("&H" + Left$(strdata, 2)))
        strdata = Right$(strdata, Len(strdata) - 2)
    Loop
    HexToString = strOutput
End Function

Public Function StringToHex(ByVal strdata As String) As String
    Dim strOutput As String
    Do Until Len(strdata) = 0
        strOutput = strOutput + Right$(String$(2, "0") + Hex$(Asc(Left$(strdata, 1))), 2)
        strdata = Right$(strdata, Len(strdata) - 1)
    Loop
    StringToHex = strOutput
End Function
Public Function ByteToString(ByRef bytData() As Byte, ByVal lngDataLength As Long) As String
    Dim lngIterator As Long
    For lngIterator = LBound(bytData) To (LBound(bytData) + lngDataLength)
        ByteToString = ByteToString & Chr$(bytData(lngIterator))
    Next lngIterator
End Function


Public Function AdvReplaceString(strSearch As String, strOld As String, strNew As String) As String
   'This is new. The old string and the new string dont
   'have to be the same length. You can replace as many
   'characters as you want at once.
    
    Dim lngFoundPos As Long
    Dim strReturn As String
    Dim strReplace As String
    Dim strIn As String
    Dim strFind As String
    Dim lngStartPos As Long
    strIn = strSearch
    strFind = strOld
    strReplace = strNew
    lngFoundPos = 1
    lngStartPos = 1
    strReturn = ""
    Do While lngFoundPos <> 0
        lngFoundPos = InStr(lngStartPos, strIn, strFind)
        If lngFoundPos <> 0 Then
            strReturn = strReturn & Mid$(strIn, lngStartPos, lngFoundPos - lngStartPos) & strReplace
        Else
            strReturn = strReturn & Mid$(strIn, lngStartPos)
            End If
        lngStartPos = lngFoundPos + Len(strFind)
        Loop
    AdvReplaceString = strReturn
    End Function

Public Function UpDown(ByVal AnyStr As String) As String
'this code isn't useful, its just to show you how
'to manipulate strings.
'I am a string freak.
  Dim i As Integer, b As String
  For i = 1 To Len(AnyStr)
  Select Case i Mod 2
    Case 0
      AnyStr = Left(AnyStr, i - 1) + LCase(Mid(AnyStr, i, 1)) + Right(AnyStr, Len(AnyStr) - i)
    Case 1
      AnyStr = Left(AnyStr, i - 1) + UCase(Mid(AnyStr, i, 1)) + Right(AnyStr, Len(AnyStr) - i)
  End Select
  Next
  UpDown = AnyStr
  
End Function

Public Function ReverseString(ByRef YourString As String) As String
   Dim idx As Long
   Dim ByteArray() As Byte
   Dim tmpByte As Byte
   Dim MAX As Long
   
   ByteArray = StrConv(YourString, vbFromUnicode)
   MAX = Len(YourString) - 1
   
   For idx = 0 To MAX \ 2
      tmpByte = ByteArray(idx)
      ByteArray(idx) = ByteArray(MAX - idx)
      ByteArray(MAX - idx) = tmpByte
   Next idx
   ReverseString = StrConv(ByteArray, vbUnicode)
   
End Function

Public Function ScrambleText(Word As String) As String
Dim g, i, position As Integer
Dim letter, newword As String
g = Len(Word)

ReDim scram(1 To g)
 For i = 1 To g
  scram(i) = ""
Next i

For i = 1 To g
letter = Mid(Word, i, 1)

Randomize
Do
 position = Int(Rnd * Len(Word)) + 1
  Loop Until scram(position) = ""
  scram(position) = letter

Next i

For i = 1 To g
newword = newword & scram(i)
Next i
ScrambleText = newword
End Function


Function IsVowel(txt As String) As Boolean
If Len(txt$) <> 1 Then Exit Function
If UCase(txt$) = UCase("A") Or UCase(txt$) = UCase("E") Or UCase(txt$) = UCase("I") Or UCase(txt$) = UCase("o") Or UCase(txt$) = UCase("u") Then
IsVowel = True
Else
IsVowel = False
End If
End Function

Function CaseTalker(txt As String) As String
Dim i, letter$, whats$
For i = 1 To Len(txt$)
letter$ = Mid(txt$, i, 1)
If IsCharAlpha(Asc(letter$)) = 0 Then whats$ = whats$ + letter$: GoTo endone
If IsVowel(letter$) = True Then
whats$ = whats$ + LCase(letter$)
ElseIf IsVowel(letter$) = False Then
whats$ = whats$ + UCase(letter$)
End If
endone:
Next i
CaseTalker = whats$
End Function

Public Function IconFromBinary(filename As String, PictureControl As Object, frm As Form) As Boolean
On Error GoTo ErrorHandler:
Dim lRet As Long
Dim hIcon As Long
Dim lHdc As Long
Dim sFile As String
If Dir(filename) = "" Then Exit Function
lHdc = PictureControl.hDC
If lHdc = 0 Then Exit Function
frm.AutoRedraw = True
PictureControl.AutoRedraw = True
sFile = filename & Chr(0)
hIcon = ExtractIcon(frm.hwnd, sFile, 0)
lRet = DrawIcon(lHdc, 0, 0, hIcon)
If lRet <> 0 Then
    PictureControl.Refresh
    DestroyIcon hIcon
    IconFromBinary = Err.LastDllError = 0
End If
ErrorHandler:
End Function

Sub ExplodeForm(f As Form, Movement As Integer)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, y%, cx%, cy%
Dim TheScreen As Long
Dim Brush As Long
GetWindowRect f.hwnd, myRect
formWidth = (myRect.Right - myRect.Left)
formHeight = myRect.Bottom - myRect.Top
TheScreen = GetDC(0)
Brush = CreateSolidBrush(f.BackColor)
For i = 1 To Movement
cx = formWidth * (i / Movement)
cy = formHeight * (i / Movement)
X = myRect.Left + (formWidth - cx) / 2
y = myRect.Top + (formHeight - cy) / 2
Rectangle TheScreen, X, y, X + cx, y + cy
Next i
X = ReleaseDC(0, TheScreen)
DeleteObject (Brush)
End Sub

Public Sub ImplodeForm(f As Form, Movement As Integer)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, y%, cx%, cy%
Dim TheScreen As Long
Dim Brush As Long
GetWindowRect f.hwnd, myRect
formWidth = (myRect.Right - myRect.Left)
formHeight = myRect.Bottom - myRect.Top
TheScreen = GetDC(0)
Brush = CreateSolidBrush(f.BackColor)
For i = Movement To 1 Step -1
cx = formWidth * (i / Movement)
cy = formHeight * (i / Movement)
X = myRect.Left + (formWidth - cx) / 2
y = myRect.Top + (formHeight - cy) / 2
Rectangle TheScreen, X, y, X + cx, y + cy
Next i
X = ReleaseDC(0, TheScreen)
DeleteObject (Brush)
End Sub

Function KillDupes(list As ListBox)
Dim i, X
For i = 0 To list.ListCount - 1
    For X = 0 To list.ListCount - 1
    If i = X Then GoTo Nextx
        If StripSpace(LCase(list.list(X))) = StripSpace(LCase(list.list(i))) Then ' aha! if the items are equal
                             
        list.RemoveItem X
    End If
Nextx:
    Next X
Next i
End Function

Function IsNotVowel(letter As String) As Boolean
If Len(letter$) <> 1 Then Exit Function
If IsVowel(letter$) = False Then IsNotVowel = True
End Function

Function DeltreeDetector(filename As String) As Boolean
'this is not fool proof
'it will only find the word deltree
'in the executable or file you choose
Dim a$, sFile, txt$
sFile = FreeFile
Open filename$ For Input As #1
Do While Not EOF(1)
Line Input #sFile, a$
If InStr(UCase(a$), UCase("deltree")) Then
DeltreeDetector = True
GoTo quitit
End If
DoEvents
Loop
DeltreeDetector = False
quitit:
Close #sFile
End Function

Function CharMaker(TheText As String) As String
Dim i, first$, txt$, wow$, trimit$
For i = 1 To Len(TheText$)
first$ = Mid(TheText$, i, 1)
txt$ = Asc(first$)
wow$ = wow$ & "Chr(" & txt$ & ") & "
Next i
trimit$ = Left(wow$, Len(wow$) - 3)
CharMaker$ = trimit$
End Function
Function SetDWORDValue(SubKey As String, Entry As String, Value As Long)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then
      rtn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4)
      If Not rtn = ERROR_SUCCESS Then
         If DisplayErrorMsg = True Then
            MsgBox ErrorMsg(rtn)
         End If
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user want errors displayed
         MsgBox ErrorMsg(rtn) 'display the error
      End If
   End If
End If

End Function
Function GetDWORDValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         GetDWORDValue = lBuffer  'return the value
      Else                        'otherwise, if the value couldnt be retreived
         GetDWORDValue = "Error"  'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants errors displayed
            MsgBox ErrorMsg(rtn)        'tell the user what was wrong
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetDWORDValue = "Error"        'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MsgBox ErrorMsg(rtn)        'tell the user what was wrong
      End If
   End If
End If

End Function

Function SetBinaryValue(SubKey As String, Entry As String, Value As String)
Dim i
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      lDataSize = Len(Value)
      ReDim ByteArray(lDataSize)
      For i = 1 To lDataSize
      ByteArray(i) = Asc(Mid$(Value, i, 1))
      Next
      rtn = RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if the was an error writting the value
         If DisplayErrorMsg = True Then 'if the user want errors displayed
            MsgBox ErrorMsg(rtn)        'display the error
         End If
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MsgBox ErrorMsg(rtn) 'display the error
      End If
   End If
End If

End Function

Function GetBinaryValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened
      lBufferSize = 1
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize) 'get the value from the registry
      sBuffer = Space(lBufferSize)
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         GetBinaryValue = sBuffer 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetBinaryValue = "Error" 'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants to errors displayed
            MsgBox ErrorMsg(rtn)  'display the error to the user
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetBinaryValue = "Error" 'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants to errors displayed
         MsgBox ErrorMsg(rtn)  'display the error to the user
      End If
   End If
End If

End Function
Function DeleteKey(keyname As String)

Call ParseKey(keyname, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, keyname, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = RegDeleteKey(hKey, keyname) 'delete the key
      rtn = RegCloseKey(hKey)  'close the key
   End If
End If

End Function

Function GetMainKeyHandle(MainKeyName As String) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
   
Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function

Function ErrorMsg(lErrorCode As Long) As String
    Dim GetErrorMsg
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1009, 1015
            GetErrorMsg = "The Registry Database is corrupt!"
       Case 2, 1010
            GetErrorMsg = "Bad Key Name"
       Case 1011
            GetErrorMsg = "Can't Open Key"
       Case 4, 1012
            GetErrorMsg = "Can't Read Key"
       Case 5
            GetErrorMsg = "Access to this key is denied"
       Case 1013
            GetErrorMsg = "Can't Write Key"
       Case 8, 14
            GetErrorMsg = "Out of memory"
       Case 87
            GetErrorMsg = "Invalid Parameter"
       Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case Else
            GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
End Select

End Function

Function GetStringValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      sBuffer = Space(255)     'make a buffer
      lBufferSize = Len(sBuffer)
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         sBuffer = Trim(sBuffer)
         GetStringValue = Left(sBuffer, Len(sBuffer) - 1) 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetStringValue = "Error" 'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants errors displayed then
            MsgBox ErrorMsg(rtn)  'tell the user what was wrong
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetStringValue = "Error"       'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants errors displayed then
         MsgBox ErrorMsg(rtn)        'tell the user what was wrong
      End If
   End If
End If

End Function

Private Sub ParseKey(keyname As String, Keyhandle As Long)
    
rtn = InStr(keyname, "\") 'return if "\" is contained in the Keyname

If Left(keyname, 5) <> "HKEY_" Or Right(keyname, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + keyname 'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(keyname)
   keyname = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(Left(keyname, rtn - 1)) 'seperate the Keyname
   keyname = Right(keyname, Len(keyname) - rtn)
End If

End Sub
Function CreateKey(SubKey As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegCreateKey(MainKeyHandle, SubKey, hKey) 'create the key
   If rtn = ERROR_SUCCESS Then 'if the key was created then
      rtn = RegCloseKey(hKey)  'close the key
   End If
End If

End Function
Function SetStringValue(SubKey As String, Entry As String, Value As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value)) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         If DisplayErrorMsg = True Then 'if the user wants errors displayed
            MsgBox ErrorMsg(rtn)        'display the error
         End If
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MsgBox ErrorMsg(rtn)        'display the error
      End If
   End If
End If

End Function

Function IsAIMExe(filename As String) As Boolean
Dim free, filelength, Text$, X, regNeeded, namee, VersionTag
If Len(filename$) = 0 Then Exit Function
free = FreeFile
Open filename$ For Binary As #free
filelength = LOF(free)
For X = 1 To filelength Step 32000
Text$ = Space(32000)
Get #free, X, Text$
regNeeded = InStr(1, Text$, "regNeeded", 1)
namee = InStr(1, Text$, "name", 1)
VersionTag = InStr(1, Text$, "VersionTag", 1)
If regNeeded <> 0 And namee <> 0 And VersionTag <> 0 Then
IsAIMExe = True
GoTo quit
End If
Next X
quit:
Close #free
End Function

Function LoadToList(catchname As String, list As ListBox, filename As String)
'the catchname is what will refuse loading
'if the file was not created by your program
'so example: Call LoadToList("killz is cool", list1, "C:\whatever.txt")
Dim free, a$
free = FreeFile
Open filename$ For Binary As #free
Input #free, a$
If Not UCase(a$) = UCase(catchname$) Then Exit Function
Do While Not EOF(free)
DoEvents
Input #free, a$
If UCase(a$) <> UCase(catchname$) Then list.AddItem a$
Loop
Close #free
Call KillDupes(list)
End Function

Function SaveToList(catchname As String, list As ListBox, filename As String)
Dim free, i
free = FreeFile
Open filename$ For Append As #free 'using append will create the file if it is not there
Write #1, catchname$
For i = 0 To list.ListCount - 1
Write #1, list.list(i)
Next i
Close #1
End Function

Function ScanFile(SearchList As ListBox, FoundList As ListBox, filename As String)
Dim fileL, i, X, Text$
If SearchList.ListCount = 0 Then Exit Function
FoundList.Clear
Open filename$ For Binary Access Read As #1
fileL = LOF(1)
For i = 0 To SearchList.ListCount - 1
For X = 1 To fileL Step 32000
Text$ = Space(32000)
Get #1, X, Text$
If InStr(1, LCase(Text$), LCase(SearchList.list(i)), 1) <> 0 Then
FoundList.AddItem SearchList.list(i)
GoTo nexti
Exit Function
End If
Next X
nexti:
Next i
Close #1
MsgBox "search complete."
End Function

Function IsWAOL(filename As String) As Boolean
Dim fileL, X, Text$, splash, idb, build, AOL, _
norem, forcerem, appid, a
Open filename$ For Binary Access Read As #1
fileL = LOF(1)
For X = 1 To fileL Step 32000
Text$ = Space(32000)
Get #1, X, Text$
splash = InStr(1, Text$, "SPLASH256", 1)
idb = InStr(1, Text$, "IDBSPLASH", 1)
build = InStr(1, Text$, "Build", 1)
AOL = InStr(1, Text$, "America Online", 1)
norem = InStr(1, Text$, "NoRemove", 1)
forcerem = InStr(1, Text$, "ForceRemove", 1)
appid = InStr(1, Text$, "AppID", 1)
If splash <> 0 Then a = a + 1
If idb <> 0 Then a = a + 1
If build <> 0 Then a = a + 1
If AOL <> 0 Then a = a + 1
If norem <> 0 Then a = a + 1
If forcerem <> 0 Then a = a + 1
If appid <> 0 Then a = a + 1
If a = 7 Then IsWAOL = True: Close #1: Exit Function
Next X
Close #1
End Function


Function SerialNumberGen(prefix As String, maxm As String)
Dim i, letter$
Randomize
For i = 1 To Len(maxm)
letter$ = Mid(maxm, i, 1)
If IsNumeric(letter$) = False Then Exit Function
Next i
SerialNumberGen = prefix & Val(Int(Rnd * maxm + 1))
End Function

Function IsNumeric(number As String)
If number$ = "1" Or number$ = "2" Or number$ = "3" Or number$ = "4" Or number$ = "5" Or number$ = "6" Or number$ = "7" Or number$ = "8" Or number$ = "9" Or number$ = "0" Then IsNumeric = True Else IsNumeric = False
End Function

Function GetNumPrinterJobs() As Long
    Dim hPrinter As Long, lNeeded As Long, lReturned As Long
    Dim lJobCount As Long
    OpenPrinter Printer.DeviceName, hPrinter, ByVal 0&
    EnumJobs hPrinter, 0, 99, 1, ByVal 0&, 0, lNeeded, lReturned
    If lNeeded > 0 Then
        ReDim byteJobsBuffer(lNeeded - 1) As Byte
        EnumJobs hPrinter, 0, 99, 1, byteJobsBuffer(0), lNeeded, lNeeded, lReturned
        If lReturned > 0 Then
            lJobCount = lReturned
        Else
            lJobCount = 0
        End If
    Else
        lJobCount = 0
    End If
    ClosePrinter hPrinter
    GetNumPrinterJobs = CStr(lJobCount)

End Function

Function ScanForURL(filename As String, list As ListBox)
'this fuckin thing. the only reason i put this here
'is probably when im thinking clearly, i can debug this
'thing. some files it gets the urls just fine
', the others, it has a bunch of gibberish at the
'end of the url. if you can help, email me at: killz@n2.com

Static thewww, thecom, thenet, theorg
thewww = 1: thecom = 1: thenet = 1: theorg = 1
Dim free
free = FreeFile
list.Clear
Open filename$ For Binary Access Read As #free
Dim fileL, i
fileL = LOF(free)
Dim Text$
For i = 1 To fileL Step 32000
Text$ = Space(32000)
Get #free, i, Text$
thewww = InStr(1, LCase(Text$), LCase("www."), 1)
thecom = InStr(1, LCase(Text$), LCase(".com"))
thenet = InStr(1, LCase(Text$), LCase(".net"))
theorg = InStr(1, LCase(Text$), LCase(".org"))
If thewww <> 0 And thecom <> 0 Then
list.AddItem Mid(Text$, thewww, thecom + 3)
GoTo again
End If
If thewww <> 0 And thenet <> 0 Then
list.AddItem Mid(Text$, thewww, thenet + 3)
GoTo again
End If
If thewww <> 0 And theorg <> 0 Then
list.AddItem Mid(Text$, thewww, theorg + 3)
GoTo again
End If
again:
Next i
Close #free
End Function

Private Sub RC4Initialize(strPwd)
'=========================================================================
' This routine called by EnDeCrypt function. Initializes the
' sbox and the key array)
'=========================================================================

    Dim tempSwap As String
    Dim a As Long
    Dim b As Long
    Dim intLength As Integer

    intLength = Len(strPwd)
    For a = 0 To 255
        Key(a) = Asc(Mid(strPwd, (a Mod intLength) + 1, 1))
        sbox(a) = a
    Next

    b = 0
    For a = 0 To 255
        b = (b + sbox(a) + Key(a)) Mod 256
        tempSwap = sbox(a)
        sbox(a) = sbox(b)
        sbox(b) = tempSwap
    Next

End Sub

Public Function EnDeCrypt(plaintxt As String, psw As String)
'=========================================================================
' This routine does all the work. Call it both to Encrypt
' and to Decrypt your data.
'=========================================================================
    Dim temp As String
    Dim a As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim cipherby As Long
    Dim cipher As String

    i = 0
    j = 0

    RC4Initialize psw

    For a = 1 To Len(plaintxt)
        i = (i + 1) Mod 256
        j = (j + sbox(i)) Mod 256
        temp = sbox(i)
        sbox(i) = sbox(j)
        sbox(j) = temp

        k = sbox((sbox(i) + sbox(j)) Mod 256)

        cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
        cipher = cipher & Chr(cipherby)
    Next

    EnDeCrypt = cipher

End Function


Function PlayMusic(cdfilename As String)
Dim r%
    r% = mciSendString("OPEN " + cdfilename$ + " TYPE SEQUENCER ALIAS " + cdfilename$, 0&, 0, 0)
    r% = mciSendString("PLAY " + cdfilename$ + " FROM 0", 0&, 0, 0)
    r% = mciSendString("CLOSE ANIMATION", 0&, 0, 0)
End Function

Function StopMusic(cdfilename As String)
Dim r%
   r% = mciSendString("OPEN " + cdfilename$ + " TYPE SEQUENCER ALIAS " + cdfilename$, 0&, 0, 0)
    r% = mciSendString&("STOP " + cdfilename$, 0&, 0, 0)
    r% = mciSendString&("CLOSE ANIMATION", 0&, 0, 0)
End Function

Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function

Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
'***************************************
'* This part of the codes i got from an*
'* example (dos-shape) by dos. it lets *
'* your form me any shape.             *
'* Thanks!                             *
'* web site:  http://www.hider.com/dos *
'***************************************

  Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
  Dim lngRgnFinal As Long, lngRgnTmp As Long
  Dim lngStart As Long, lngRow As Long
  Dim lngCol As Long
  If lngTransColor& < 1 Then
    lngTransColor& = GetPixel(picSource.hDC, 0, 0)
  End If
  lngHeight& = picSource.Height / Screen.TwipsPerPixelY
  lngWidth& = picSource.Width / Screen.TwipsPerPixelX
  lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
  For lngRow& = 0 To lngHeight& - 1
    lngCol& = 0
    Do While lngCol& < lngWidth&
      Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) = lngTransColor&
        lngCol& = lngCol& + 1
      Loop
      If lngCol& < lngWidth& Then
        lngStart& = lngCol&
        Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) <> lngTransColor&
          lngCol& = lngCol& + 1
        Loop
        If lngCol& > lngWidth& Then lngCol& = lngWidth&
        lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
        lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
        DeleteObject (lngRgnTmp&)
      End If
    Loop
  Next
  RegionFromBitmap& = lngRgnFinal&
End Function

Sub ChangeMask(mask As PictureBox, frmdd As Form)
On Error Resume Next ' In case of error
' This is also part of Dos's Dos-Shape example. To update if the skin is changed
  Dim lngRetr As Long
  lngRegion& = RegionFromBitmap(mask)
  lngRetr& = SetWindowRgn(frmdd.hwnd, lngRegion&, True)
End Sub

Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$)
'by monk-e-god
Dim textlen%, fstlen%, part1$, part2$, i, textdone$, lastchr$, colorx
Dim colorx2, faded1$, faded2$
    textlen% = Len(TheText)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, textlen% - fstlen%)
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        textdone$ = Left(part1$, i)
        lastchr$ = Right(textdone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        faded1$ = faded1$ + "<Font Color=#" & colorx2 & ">" + "" + lastchr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        textdone$ = Left(part2$, i)
        lastchr$ = Right(textdone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        faded2$ = faded2$ + "<Font Color=#" & colorx2 & ">" + "" + lastchr$
    Next i
    FadeThreeColor = faded1$ + faded2$

End Function

Function FadeByColor3(Colr1, Colr2, Colr3, TheText$)
Dim dacolor1$, dacolor2$, dacolor3$, rednum1%, rednum2%, rednum3%
Dim greennum1%, greennum2%, greennum3%, bluenum1%, bluenum2%
Dim bluenum3%
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, TheText)
End Function


Function LastChatLine() As String
Dim chatroomrich$, chatroom&, richcntl&, i, letter$, txt$
chatroom& = FindChatRoom()
richcntl& = FindChildByClass(chatroom&, "RICHCTNL")
chatroomrich$ = GetText(richcntl&)
For i = 1 To Len(chatroomrich$)
letter$ = Mid(chatroomrich$, i, 1)
If letter$ = "" Then GoTo nexti
If Asc(letter$) = 13 Then
chatroomrich$ = Mid(chatroomrich$, i, Len(chatroomrich$))
End If
nexti:
Next i
txt$ = Mid(chatroomrich$, 3, Len(chatroomrich$))
LastChatLine$ = txt$
End Function

Function IsOnline() As Boolean
Dim aolchild&, AOL&
aolchild& = FindAOLChild()
If InStr(LCase(GetText(aolchild&)), LCase("Welcome, ")) Then
IsOnline = True
Exit Function
End If
While aolchild&
aolchild& = GetWindow(aolchild&, 2)
If InStr(GetText(aolchild&), "Welcome, ") Then
IsOnline = True
Exit Function
End If
DoEvents
Wend
IsOnline = False
End Function

Function GetSignOnSN() As Long
Dim aolchild&
aolchild& = FindAOLChild()
If InStr(LCase(GetText(aolchild&)), LCase("Goodbye from america online")) Then
GetSignOnSN& = aolchild&
Exit Function
End If
While aolchild&
DoEvents
aolchild& = GetWindow(aolchild&, 2)
If InStr(LCase(GetText(aolchild&)), LCase("goodbye from america online")) Then
GetSignOnSN& = aolchild&
Exit Function
End If
Wend
aolchild& = FindAOLChild()
If InStr(LCase(GetText(aolchild&)), LCase("sign on")) Then
GetSignOnSN& = aolchild&
Exit Function
End If
While aolchild&
DoEvents
aolchild& = GetWindow(aolchild&, 2)
If InStr(LCase(GetText(aolchild&)), LCase("sign on")) Then
GetSignOnSN& = aolchild&
Exit Function
End If
Wend
GetSignOnSN& = 0
End Function

Function InStrFindChildTitle(wnd As Long, Window As String) As Long
Dim TheText$
TheText$ = GetText(wnd&)
If InStr(UCase(TheText$), UCase(Window$)) Then
InStrFindChildTitle& = wnd&
Exit Function
End If
While wnd&
wnd& = GetWindow(wnd&, 2)
TheText$ = GetText(wnd&)
If InStr(UCase(TheText$), UCase(Window$)) Then
InStrFindChildTitle& = wnd&
Exit Function
End If
DoEvents
Wend
InStrFindChildTitle& = 0
End Function


Function WriteINI(applicationtitle As String, keyname As String, values As String, filename As String)
WritePrivateProfileString applicationtitle$, keyname$, values$, filename$
End Function

Function GetINI(applicationtitle As String, keyname As String, filename As String) As String
Dim thestring$, nc
thestring$ = String(255, 0)
nc = GetPrivateProfileString(applicationtitle$, keyname$, "", thestring$, 255, filename$)
If nc <> 0 Then thestring$ = Left$(thestring$, nc)
GetINI$ = thestring$
End Function

Function PrimeNumber(Numb As Long)
'if the number is a prime number then it returns "Prime"
'if not it returns 0
Dim k As Integer, i, r, o
Dim n As Boolean
n = True
For i = 2 To Numb
For r = 2 To Numb
o = i * r
If o = Numb Then
n = False
End If
Next r
Next i
If n = True Then
PrimeNumber = "Prime"
Else
PrimeNumber = 0
End If
End Function

Function ListPrimeNumbersFrom(FromNumber As Long)
'this lists all numbers that you declare the fromnumber
'as, if you do something like 1000, it will take a while
'ex: x = listprimenumbersfrom(1000): Text1.text = x
Dim a As Long
a = 0
Do While Not FromNumber = a
a = a + 1
If PrimeNumber(a) = "Prime" Then
ListPrimeNumbersFrom = ListPrimeNumbersFrom & a & ", "
End If
DoEvents
Loop
End Function

Public Function ConvTime(Seconds As Double, Level As Byte, fullreturn As Boolean, _
   display_text As Boolean, delimiter As String)

' Example: Msgbox ConvTime(2147483647, 5, true, true, "")
'any amount of seconds over 2147483647 will result in overflow

If ((Seconds > 2147483647) Or (Level > 5) Or _
   (Len(delimiter) > 1)) Then
ConvTime = 0
Exit Function
End If

Dim Minutes As Double, Hours As Double, Days As Double
Dim Weeks As Double, Years As Double

Dim Minutes_ As Double, Seconds_ As Double, Hours_ As Double, Days_ As Double
Dim Weeks_ As Double, Years_ As Double

Minutes = Seconds \ 60
Hours = Minutes \ 60
Days = Hours \ 24
Weeks = Days \ 7
Years = Days \ 365

Seconds_ = Seconds - (60 * Minutes)
Minutes_ = Minutes - (60 * Hours)
Hours_ = Hours - (24 * Days)
Days_ = Days - (7 * Weeks)
Weeks_ = Weeks - (52 * Years)


If (fullreturn = True) Then

    Select Case Level
    
    Case 0
        If (display_text = True) Then
        ConvTime = Seconds & " seconds."
        Else
        ConvTime = Seconds
        End If
    Case 1
        If (display_text = True) Then
        ConvTime = Minutes & " minutes, " & Seconds_ & " seconds"
        Else
        ConvTime = Minutes & delimiter & Seconds_
        End If
    Case 2
        If (display_text = True) Then
        ConvTime = Hours & " hours, " & Minutes_ & " minutes, " _
           & Seconds_ & " seconds"
        Else
        ConvTime = Hours & delimiter & Minutes_ & delimiter & Seconds_
        End If
    Case 3
        If (display_text = True) Then
        ConvTime = Days & " days, " & Hours_ & " hours, " & _
           Minutes_ & " minutes, " & Seconds_ & " seconds"
        Else
        ConvTime = Days & delimiter & Hours_ & delimiter & _
             Minutes_ & delimiter & Seconds_
        End If
    Case 4
        If (display_text = True) Then
        ConvTime = Weeks & " weeks, " & Days_ & " days, " & Hours_ & _
               " hours, " & Minutes_ & " minutes, " & Seconds_ & " seconds"
        Else
        ConvTime = Weeks & delimiter & Days_ & delimiter & Hours_ & _
           delimiter & Minutes_ & delimiter & Seconds_
        End If
    Case 5
        If (display_text = True) Then
        ConvTime = Years & " years, " & Weeks_ & " weeks, " & Days_ & " days, " & Hours_ & " hours, " & Minutes_ & " minutes, " & Seconds_ & " seconds"
        Else
        ConvTime = Years & delimiter & Weeks_ & delimiter & Days_ & delimiter & Hours_ & delimiter & Minutes_ & delimiter & Seconds_
        End If
    Case Else
            ConvTime = 0
    End Select

Else

    Select Case Level
    
    Case 0
        If (display_text = True) Then
        ConvTime = Seconds & " seconds."
        Else
        ConvTime = Seconds
        End If
    Case 1
        If (display_text = True) Then
        ConvTime = Minutes & " minutes."
        Else
        ConvTime = Minutes
        End If
    Case 2
        If (display_text = True) Then
        ConvTime = Hours & " hours."
        Else
        ConvTime = Hours
        End If
    Case 3
        If (display_text = True) Then
        ConvTime = Days & " days."
        Else
        ConvTime = Days
        End If
    Case 4
        If (display_text = True) Then
        ConvTime = Weeks & " weeks."
        Else
        ConvTime = Weeks
        End If
    Case 5
        If (display_text = True) Then
        ConvTime = Years & " years."
        Else
        ConvTime = Years
        End If
    Case Else
            ConvTime = 0
    End Select

End If

End Function
Public Sub ComplexAddition(a, b, c, d, suma, sumb)
    suma = a + c
    sumb = b + d
End Sub
Public Sub ComplexSubtraction(a, b, c, d, diffa, diffb)
    diffa = a - c
    diffb = b - d
End Sub
Public Sub ComplexCosine(a, b, AnswerA, AnswerB)
    Call ComplexMultiplication(a, b, 0, 1, FirstMultA, FirstMultB)
    Call ComplexExponentiation(FirstMultA, FirstMultB, FirsteA, FirsteB)
    Call ComplexMultiplication(a, b, 0, -1, SecondMultA, SecondMultB)
    Call ComplexExponentiation(SecondMultA, SecondMultB, SecondeA, SecondeB)
    Call ComplexAddition(FirsteA, FirsteB, SecondeA, SecondeB, SumeA, SumeB)
    Call ComplexDivision(SumeA, SumeB, 2, 0, AnswerA, AnswerB)
End Sub
Public Sub ComplexDivision(a, b, c, d, DivisionA, DivisionB)
    Call ComplexMultiplication(a, b, c, -d, multiplicationa, multiplicationb)
    Divisor = c ^ 2 + d ^ 2
    DivisionA = multiplicationa / Divisor
    DivisionB = multiplicationb / Divisor
End Sub
Public Sub ComplexExponentiation(a, b, AnswerA, AnswerB)
    AnswerA = Exp(a) * Cos(b)
    AnswerB = Exp(a) * Sin(b)
End Sub
Public Sub ComplexMagnitude(a, b, Magnitude)
    Magnitude = Sqr(a ^ 2 + b ^ 2)
End Sub
Public Sub ComplexMultiplication(a, b, c, d, AnswerA, AnswerB)
    AnswerA = a * c - b * d
    AnswerB = a * d + b * c
End Sub
Public Sub ComplexSine(a, b, AnswerA, AnswerB)
    Call ComplexMultiplication(a, b, 0, 1, FirstMultA, FirstMultB)
    Call ComplexExponentiation(FirstMultA, FirstMultB, FirsteA, FirsteB)
    Call ComplexMultiplication(a, b, 0, -1, SecondMultA, SecondMultB)
    Call ComplexExponentiation(SecondMultA, SecondMultB, SecondeA, SecondeB)
    Call ComplexAddition(FirsteA, FirsteB, -SecondeA, -SecondeB, SumeA, SumeB)
    Call ComplexDivision(SumeA, SumeB, 2, 1, AnswerA, AnswerB)
End Sub
Public Sub ComplexTangent(a, b, AnswerA, AnswerB)
    Call ComplexSine(a, b, SineA, SineB)
    Call ComplexCosine(a, b, CosineA, CosineB)
    Call ComplexDivision(SineA, SineB, CosineA, CosineB, AnswerA, AnswerB)
End Sub
Public Sub ComplexLogWithSpecialBase(a, b, BaseA, BaseB, AnswerA, AnswerB)
    Call ComplexLog(a, b, Answer1A, Answer1B)
    Call ComplexLog(BaseA, BaseB, Answer2A, Answer2B)
    Call ComplexDivision(Answer1A, Answer1B, Answer2A, Answer2B, AnswerA, AnswerB)
End Sub
Public Sub ComplexLog(a, b, AnswerA, AnswerB)
    PI = 4 * Atn(1)
    Call ComplexMagnitude(a, b, Magnitude)
    If Magnitude = 0 Then
        AnswerA = 0
    Else
        AnswerA = Log(Magnitude)
    End If
    If (a >= 0 And b = 0) Or (a = 0 And b = 0) Then
        AnswerB = 0
    ElseIf a = 0 And b >= 0 Then
        AnswerB = PI / 2
    ElseIf a <= 0 And b = 0 Then
        AnswerB = PI
    ElseIf a = 0 And b <= 0 Then
        AnswerB = 3 * PI / 2
    Else
        AnswerB = PI / 2 * (-a / Abs(a) + 1) + PI * a / Abs(a) * (-b / Abs(b) + 1) + a * b / Abs(a * b) * Atn(Abs(b / a))
    End If
End Sub
Public Sub ComplexPower(a, b, c, d, AnswerA, AnswerB)
    Call ComplexLog(a, b, LogA, LogB)
    Call ComplexMultiplication(c, d, LogA, LogB, MultA, MultB)
    Call ComplexExponentiation(MultA, MultB, AnswerA, AnswerB)
End Sub
Public Function MathDecFrac(wholenumber, nonrepeating, repeating As Double) As String
Dim thefirst$, thesecond, thethird, mathfrac
Dim thetop, thebottom
thefirst$ = Trim(Str(nonrepeating) + Trim(Str(repeating)))
thesecond = 10 ^ (Len(thefirst$))
thethird = 10 ^ (Len(nonrepeating))
mathfrac = Trim(Str((Val(thefirst$) - nonrepeating))) + "/" + Trim(Str((thesecond - thethird)))
wholenumber = wholenumber * Val(Trim(Str((thesecond - thethird))))
thetop = Val(Trim(Str((Val(thefirst$) - nonrepeating)))) + Val(wholenumber)
thebottom = Trim(Str((thesecond - thethird)))
MathDecFrac = Trim(Str(thetop)) + "/" + Trim(Str(thebottom))
End Function
Function CRad(AngleMeasure As Double) As Double
    'converts degrees to redians
Const PI = 3.1415
CRad = (AngleMeasure * (PI / 180)) / PI
End Function
Function IsOdd(zNumber As Integer) As Boolean
    If (zNumber / 2) <> Int((zNumber / 2)) Then IsOdd = True Else IsOdd = False
End Function
Function CDeg(AngleMeasure As Double) As Double
'converts radians to degrees
Const PI = 3.1415
CDeg = (AngleMeasure / (PI / 180)) * PI
End Function
Function Distance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    'Finds the distance between two points,
    '     given their coordinates.
Distance = Sqr(Abs(x1 - x2) ^ 2 + Abs(y1 - y2) ^ 2)
End Function
Public Function GETLcm(aInt As Integer, bInt As Integer)
    'This Function will find the LCM
    '(Least Common Multiple) of any
    'two whole numbers
Dim X As Integer, a1, q
For X = 1 To bInt
Let a1 = aInt * X
Let q = a1 / bInt
If q = Int(q) Then GoTo Anwser
Next X
Anwser:
GETLcm = a1
End Function

