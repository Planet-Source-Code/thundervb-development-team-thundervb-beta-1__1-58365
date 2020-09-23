Attribute VB_Name = "modLowLevel"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

'this module will contatin low level API codes

'Revision history:
'
'19/8/2004[dd/mm/yyyy] : Creted by Libor
'Module created , intial version
'
'
'19/8/2004 [dd/mm/yyyy] : Modifyed by Raziel
'All the delcarations were made public and moved to the declaration module (\misc code\declares.bas)
'Many things added here , everything is marked
'
'
'22/9/2004[dd/mm/yyyy] : Edited by Raziel
'Fixes hre and there , mainly on string convertion

'Code from VbInlineASM
Public Function ExecuteCommand(ByVal CommandLine As String, ByRef sOutputText As String, Optional workdir As String, Optional ByVal eWindowState As ESW = SW_HIDE) As Boolean
'***I have patched CreateProcessLong declaration (it was wrong) - lpEnvironment As Any - should be - Byval lpEnvironment as Long *** Libor - 2004
'***I've modified this function too....JS-2002
'*** Edited by Raziel for workdir
'DOSOutpus
'Capture the outputs of a DOS command
'Author: Marco Pipino
'marcopipino@libero.it
'28/02/2002
Dim proc As PROCESS_INFORMATION     'Process info filled by CreateProcessA
Dim ret As Long                     'long variable for get the return value of the API functions
Dim start As STARTUPINFO           'StartUp Info passed to the CreateProceeeA function
Dim sa As SECURITY_ATTRIBUTES       'Security Attributes passeed to the CreateProcessA function
Dim hReadPipe As Long               'Read Pipe handle created by CreatePipe
Dim hWritePipe As Long              'Write Pite handle created by CreatePipe
Dim lngBytesread As Long            'Amount of byte read from the Read Pipe handle
Dim strBuff As String * 256         'String buffer reading the Pipe
        
    If Get_Compile(ModifyCmdLine) Then
        frmViewer.ShowViewer "Command Line Edit", CommandLine, False, False
    End If

    If Len(CommandLine) = 0 Then
        ExecuteCommand = False
        Exit Function
    End If
    LogMsg CommandLine, "modLowLevel", "ExecuteCommand"
    On Error Resume Next
    
    'Create the Pipe
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)

    If ret = 0 Then
        'If an error occur during the Pipe creation exit
        'msgboxx "CreatePipe failed. Error: " & Err.LastDllError, vbCritical
        Exit Function
    End If

    'Launch the command line application
    start.cb = Len(start)
    start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    start.wShowWindow = eWindowState
    
    'set the StdOutput and the StdError output to the same Write Pipe handle
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
       
    'Execute the command
    If Len(workdir) > 0 Then
        ret& = CreateProcessLong2(0&, CommandLine, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, ByVal workdir, start, proc)
    Else
        ret& = CreateProcessLong(0&, CommandLine, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, 0&, start, proc)
    End If
    
    If ret <> 1 Then
        'if the command is not found ....
        Exit Function
    End If

    'Now We can ... must close the hWritePipe
    ret = CloseHandle(hWritePipe)
    sOutputText = ""                             '*** patched LIBOR
    
    'Read the ReadPipe handle
    Do
        ret = ReadFile(ByVal hReadPipe, strBuff, ByVal 256, lngBytesread, ByVal 0&)
        sOutputText = sOutputText & Left(strBuff, lngBytesread)
        'Send data to the object via ReceiveOutputs event
    Loop While ret <> 0

    'Close the opened handles
    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)

    ExecuteCommand = True
    
End Function


'***Added by Raziel [18/9/2004]
'hmm not all of em are so lowlevel , maybe change the name to common and
'move here all helper funcs??

'waits for a proccess to end
Public Sub WaitToEnd(proc As Long)

    Do
        Sleep 100
        DoEvents
    Loop While GetProcessVersion(proc) <> 0

End Sub

'Get VB version
Function getVBVersion() As Long
Dim temp As Long

    If GetModuleHandle("VBA5.dll") > 0 Then temp = 5
    If GetModuleHandle("VBA6.dll") > 0 Then temp = 6
    getVBVersion = temp
    LogMsg "VB version=" & temp, "modLowLevel", "getVBVersion"
    
End Function

'From VBInlineAsm
'Copies a Cstring to a VB string - > will be replaced..
Function CStringZero(lpCString As Long) As String
Dim lenString As Long, sBuffer As String, lpBuffer As Long, lngStringPointer As Long, refStringPointer As Long

    If lpCString = 0 Then
        CStringZero = vbNullString
    Else
        lenString = lenCString(lpCString)
        sBuffer = String$(lenString + 1, 0) 'buffer has one extra byte for terminator
        lpBuffer = CopyCString(sBuffer, lpCString, lenString + 1)
        Mid(sBuffer, lenString + 1, 1) = " " ' to fix the 0 at the end
        CStringZero = sBuffer
    End If
    
End Function


'From ansi String Pointer to vb string
Function Cstring(ByVal lpString As Long, ByVal nCount As Long) As String
Dim s As String, temp() As Byte

    If nCount > 0 Then
        ReDim temp(nCount)
        CopyMemory temp(0), ByVal lpString, nCount
        temp(nCount) = 0
        s = StrConv(temp, vbUnicode)
    End If
    Cstring = s
    
End Function

'From Wide String Pointer to vb string
Function CstringW(ByVal lpString As Long, ByVal nCount As Long) As String
Dim s As String, temp() As Byte
    
    nCount = nCount * 2 - 1
    If nCount > 0 Then
        ReDim temp(nCount)
        CopyMemory temp(0), ByVal lpString, nCount
        temp(nCount) = 0
        s = temp
    End If
    CstringW = s
    
End Function

'Form VB str to Ansi Byte array
Public Sub BstrToAnsi(str As String, ba() As Byte)

    If Len(str) = 0 Then
        ReDim ba(0)
    Else
        ba = StrConv(str, vbFromUnicode)
        ReDim Preserve ba(UBound(ba) + 1)
    End If
    
End Sub
