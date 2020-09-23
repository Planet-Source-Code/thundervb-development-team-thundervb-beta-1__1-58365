Attribute VB_Name = "declares"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

'Declaration module , any Dll import /const are declared here
'It partialy based on code from Compiler Controller
'                                   \/
'********************************************************
'This code module is from Compiler Controller by John Chamberlain
'http://archive.devx.com/upload/free/features/vbpj/1999/11nov99/jc1199/jc1199.asp
'All the many modifications, additions and deletions were done by me. <John Sugas>2002
'********************************************************

'Revision history:
'19/8/2004[dd/mm/yyyy] : Imported by Raziel
'Module Imported , intial version
'Many declares added
Option Explicit


'Module signatures from Randy Kath's article on PE file format
Public Const IMAGE_DOS_SIGNATURE = &H5A4D       'MZ    short
Public Const IMAGE_OS2_SIGNATURE = &H454E       'NE    short
Public Const IMAGE_OS2_SIGNATURE_LE = &H454C    'LE    short
Public Const IMAGE_NT_SIGNATURE = &H4550        '--PE  long

'Memory-Related public constants from WinNT.H
Public Const PAGE_NOACCESS = &H1
Public Const PAGE_READONLY = &H2
Public Const PAGE_READWRITE = &H4
Public Const PAGE_WRITECOPY = &H8
Public Const PAGE_EXECUTE = &H10
Public Const PAGE_EXECUTE_READ = &H20
Public Const PAGE_EXECUTE_READWRITE = &H40
Public Const PAGE_EXECUTE_WRITECOPY = &H80
Public Const PAGE_GUARD = &H100
Public Const PAGE_NOCACHE = &H200
Public Const PAGE_WRITECOMBINE = &H400
Public Const MEM_COMMIT = &H1000
Public Const MEM_RESERVE = &H2000
Public Const MEM_DECOMMIT = &H4000
Public Const MEM_RELEASE = &H8000
Public Const MEM_FREE = &H10000
Public Const MEM_PRIVATE = &H20000
Public Const MEM_MAPPED = &H40000
Public Const MEM_RESET = &H80000
Public Const MEM_TOP_DOWN = &H100000
Public Const MEM_4MB_PAGES = &H80000000
Public Const SEC_FILE = &H800000
Public Const SEC_IMAGE = &H1000000
Public Const SEC_VLM = &H2000000
Public Const SEC_RESERVE = &H4000000
Public Const SEC_COMMIT = &H8000000
Public Const SEC_NOCACHE = &H10000000
Public Const MEM_IMAGE = SEC_IMAGE

Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_LEFT = &H1
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = 15     '//(BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_TOPLEFT = &H3
Public Const BF_TOPRIGHT = &H6
Public Const BF_BOTTOMLEFT = &H9
Public Const BF_BOTTOMRIGHT = &HC
Public Const BF_MIDDLE = &H800    '//' Fill in the middle.
Public Const BF_FLAT = &H4000
Public Const EDGE_BUMP = 9
Public Const EDGE_ETCHED = 6
Public Const EDGE_RAISED = 5
Public Const EDGE_SUNKEN = 10

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Public Const ERROR_MORE_DATA = 234 '  dderror
Public Const ERROR_NO_MORE_ITEMS = &H103
Public Const ERROR_KEY_NOT_FOUND = 2

Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const SYNCHRONIZE = &H100000
Public Const KEY_QUERY_VALUE = 1
Public Const KEY_SET_VALUE = 2
Public Const KEY_CREATE_SUB_KEY = 4
Public Const KEY_ENUMERATE_SUB_KEYS = 8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = &H20019
Public Const KEY_WRITE = &H20006
Public Const KEY_EXECUTE = &H20019
Public Const KEY_ALL_ACCESS = &HF003F

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)

Public Const WH_CBT = 5
Public Const HCBT_ACTIVATE = 5
Public Const HCBT_CREATEWND = 3
Public Const HCBT_DESTROYWND = 4
Public Const HCBT_MINMAX = 1
Public Const HCBT_MOVESIZE = 0
Public Const HCBT_SYSCOMMAND = 8


Public Const STARTF_USESTDHANDLES = &H100
Public Const STARTF_USESHOWWINDOW = &H1
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const INFINITE = &HFFFF      '  Infinite timeout
Public Const WAIT_TIMEOUT = &H102

Public Const cMaxPath = 260
Public Const cMaxFile = 260

Public Const OPEN_EXISTING = 3
Public Const GENERIC_READ = &H80000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20

'// ShowWindow/WinExec constants
Public Enum ESW
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_MAX = 10
End Enum

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Enum PATHS
    AddInPath
    DebugPath
    LinkerPath
    MidlPath
    PackerPath
    ProjectPath
    TextEditorPath
End Enum

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Public Declare Function CreateProcessLong Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CreateProcessLong2 Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateProcessBynum Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Long, lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function lenCString Lib "kernel32" Alias "lstrlenA" (lpString As Long) As Long
Public Declare Function CopyCString Lib "kernel32" Alias "lstrcpynA" (ByVal lpStringDestination As String, lpStringSource As Long, ByVal lngMaxLength As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Declare Function GetProcessVersion Lib "kernel32.dll" (ByVal ProcessId As Long) As Long
Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Public Declare Function CreateProcess2 Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
                                    ByVal lpProcessAttributes As Long, _
                                    ByVal lpThreadAttributes As Long, _
                                    ByVal bInheritHandles As Long, _
                                    ByVal dwCreationFlags As Long, _
                                    ByVal lpEnvironment As Long, _
                                    ByVal lpCurrentDirectory As Long, _
                                    ByVal lpStartupInfo As Long, _
                                    ByVal lpProcessInformation As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function TextOutW Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" ( _
     ByVal lpExistingFileName As String, _
     ByVal lpNewFileName As String, _
     ByVal bFailIfExists As Long) As Long

Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
     ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) As Long
     
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal x As Long, _
                         ByVal y As Long, ByVal wOptions As Long, _
                         ByVal lpRect As Long, ByVal lpString As Long, _
                         ByVal nCount As Long, ByVal lpDx As Long) As Long
                         
Public Declare Function ExtTextOutW Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
                         ByVal y As Long, ByVal wOptions As Long, _
                         ByVal lpRect As Long, ByVal lpString As Long, _
                         ByVal nCount As Long, ByVal lpDx As Long) As Long
                         
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" ( _
     ByVal lpClassName As String, _
     ByVal lpWindowName As String) As Long
     
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFilename As String) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetTextAlign Lib "gdi32.dll" ( _
     ByVal hdc As Long) As Long
Declare Function GetCaretPos Lib "user32.dll" ( _
     ByRef lpPoint As POINTAPI) As Long
Declare Function SetCaretPos Lib "user32.dll" ( _
     ByVal x As Long, _
     ByVal y As Long) As Long
     
Declare Function MessageBoxA Lib "user32.dll" ( _
     ByVal hWnd As Long, _
     ByVal lpText As Long, _
     ByVal lpCaption As Long, _
     ByVal wType As Long) As Long
     
Declare Function MessageBoxW Lib "user32.dll" ( _
     ByVal hWnd As Long, _
     ByVal lpText As Long, _
     ByVal lpCaption As Long, _
     ByVal wType As Long) As Long

Declare Function MessageBoxExA Lib "user32.dll" ( _
     ByVal hWnd As Long, _
     ByVal lpText As Long, _
     ByVal lpCaption As Long, _
     ByVal uType As Long, _
     ByVal wLanguageId As Long) As Long

Declare Function MessageBoxExW Lib "user32.dll" ( _
     ByVal hWnd As Long, _
     ByVal lpText As Long, _
     ByVal lpCaption As Long, _
     ByVal uType As Long, _
     ByVal wLanguageId As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long

Public Enum HookState
    hooked = 1
    unhooked = 2
End Enum

Global Const strIncFileData As String = ";; LISTING.INC" & vbNewLine & ";; This file contains assembler macros and is included by the files created" & vbNewLine & ";; with the -FA compiler switch to be assembled by MASM (Microsoft Macro" & vbNewLine & ";; Assembler)." & vbNewLine & ";; Copyright (c) 1993, Microsoft Corporation. All rights reserved." & vbNewLine & _
";; non destructive nops" & vbNewLine & "npad macro size" & vbNewLine & "if size eq 1" & vbNewLine & "  nop" & vbNewLine & "else" & vbNewLine & " if size eq 2" & vbNewLine & "   mov edi, edi" & vbNewLine & " else" & vbNewLine & "  if size eq 3" & vbNewLine & "    ; lea ecx, [ecx+00]" & vbNewLine & "    DB 8DH, 49H, 00H" & vbNewLine & "  else" & vbNewLine & "   if size eq 4" & vbNewLine & "     ; lea esp, [esp+00]" & vbNewLine & "     DB 8DH, 64H, 24H, 00H" & vbNewLine & _
"   else" & vbNewLine & "    if size eq 5" & vbNewLine & "      add eax, DWORD PTR 0" & vbNewLine & "    else" & vbNewLine & "     if size eq 6" & vbNewLine & "       ; lea ebx, [ebx+00000000]" & vbNewLine & "       DB 8DH, 9BH, 00H, 00H, 00H, 00H" & vbNewLine & "     else" & vbNewLine & "      if size eq 7" & vbNewLine & "    ; lea esp, [esp+00000000]" & vbNewLine & "    DB 8DH, 0A4H, 24H, 00H, 00H, 00H, 00H" & vbNewLine & _
"      else" & vbNewLine & "    %out error: unsupported npad size" & vbNewLine & "    .Err" & vbNewLine & "      endif" & vbNewLine & "     endif" & vbNewLine & "    endif" & vbNewLine & "   endif" & vbNewLine & "  endif" & vbNewLine & " endif" & vbNewLine & "endif" & vbNewLine & "endm" & vbNewLine & _
";; destructive nops" & vbNewLine & "dpad macro size, reg" & vbNewLine & "if size eq 1" & vbNewLine & "  inc reg" & vbNewLine & "else" & vbNewLine & "  %out error: unsupported dpad size" & vbNewLine & "  .Err" & vbNewLine & "endif" & vbNewLine & "endm" & vbNewLine

