Attribute VB_Name = "modHook"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit
'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Raziel
'Module Created , intial version
'Coded Hook and TrogleHook functions
'
'26/8/2004[dd/mm/yyyy] : Edited by Raziel
'Added EnumerateModules
'
'22/9/2004[dd/mm/yyyy] : Edited by Raziel
'Added HookDll_entry and HookDll_List
'and the helper functions
'

Type module_entry
    name As String
    id As Long
End Type

Type module_list
    modules() As module_entry
    count As Long
End Type

Private Declare Function OpenProcess Lib "kernel32.dll" ( _
     ByVal dwDesiredAccess As Long, _
     ByVal bInheritHandle As Long, _
     ByVal dwProcessId As Long) As Long
     
Private Declare Function EnumProcessModules Lib "psapi.dll" ( _
     ByVal hProcess As Long, _
     ByRef lphModule As Long, _
     ByVal cb As Long, _
     ByRef lpcbNeeded As Long) As Long
     
Private Declare Function CloseHandle Lib "kernel32.dll" ( _
     ByVal hObject As Long) As Long

Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" ( _
     ByVal hProcess As Long, _
     ByVal hModule As Long, _
     ByVal lpFilename As String, _
     ByVal nSize As Long) As Long
     
Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Type DllHook_entry
    ToModule As String
    DllName As String
    FunctionName As String
    FunctionAddress As Long
    HookAddress As Long
    State As HookState
End Type

Public Type DllHook_list
    items() As DllHook_entry
    count As Long
End Type


     
Public Function TrogleHook(ToModule As String, DllName As String, _
                     EntryName As String, newFunction As Long, _
                     oldFunction As Long, errstring As String) As HookState
Dim temp As Long
    temp = Hook(ToModule, DllName, EntryName, newFunction, errstring)
    TrogleHook = hooked
    
    If temp = 0 Then 'failed to hook , try to unhook then
        temp = Hook(ToModule, DllName, newFunction, oldFunction, errstring)
        TrogleHook = unhooked
    End If
    If temp Then oldFunction = temp
    
End Function
'return OldAddress on succes and 0 on fail
Public Function Hook(ToModule As String, DllName As String, _
                     EntryName As Variant, newFunction As Long, _
                     errstring As String) As Long
                     
    Dim oldFunction As Long

    If Not HookDLLImport(ToModule, DllName, EntryName, newFunction, _
                                            oldFunction, errstring) Then
    oldFunction = 0
    End If
    
    Hook = oldFunction
    
End Function

'This code is taken from CompilerControler

'Hooking DLL Calls
'by John Chamberlain
'
'You can use the logic in this function to hook imports in most DLLs and EXEs
'(not just in VB). It will work for most normal Win32 modules. If you use this
'function in your own code please credit its author (me!) and include this
'descriptive header so future users will know what it does.
'
'The call addresses for all implicitly linked DLLs are located in a table
'called the "Import Address Table (IAT)" (or the "Thunk" table). This table is
'generally located at module offset 0x1000 in both DLLs and EXEs and contains
'the addresses of all imported calls in a continuous list with exports from
'different modules separated by NULL (0x 0000 0000). When each DLL is loaded
'the operating system's loader patches this table with the correct addresses.
'In most PE file types an offset to the entry point (which is just past the
'IAT) is located at offset 0xDC from the PE file header which has a signature
'of 0x00004550 (="PE"). Thus the function finds the end of the IAT by scanning
'for this signature and locating the offset.
'
'This function hooks a DLL call by first getting the proc address for the
'specified call and then scanning the IAT for the address. If it is found
'the function substitutes the hook address into the table and returns the
'original address to the caller by reference (in case the caller wants to
'restore the IAT entry to its original state at a later time). If the
'return value was false then the hook could not be set and the reason will
'be returned by reference in the string sError.
'
'When you want to restore the hooked address pass the hook address as
'vCallNameOrAddress and the original address (to be restored) as lpHook.
'The function will find the hooked address in the table and replace it with
'the original address (see UnhookCreateProcess for an example).
'
Public Function HookDLLImport(sImportingModuleName As String, _
                        sExportingModuleName As String, vCallNameOrAddress As Variant, _
                        lpHook As Long, ByRef lpOriginalAddress As Long, _
                        ByRef sError As String) As Boolean
    
    Dim sCallName As String, lpPEHeader As Long
    Dim lpImportingModuleHandle As Long, lpExportingModuleHandle As Long, lpProcAddress As Long
    Dim vectorIAT As Long, lenIAT As Long, lpEndIAT As Long, lpIATCallAddress As Long
    Dim lpflOldProtect As Long, lpflOldProtect2 As Long
    
    On Error GoTo EH

    'Validate the hook
    If lpHook = 0 Then sError = "Hook is null.": Exit Function

    'Get handle (address) of importing module
    lpImportingModuleHandle = GetModuleHandle(sImportingModuleName)
    If lpImportingModuleHandle = 0 Then sError = "Unable to obtain importing module handle for """ & sImportingModuleName & """.": Exit Function

    'Get the proc address of the IAT entry to be changed
    If VarType(vCallNameOrAddress) = vbString Then
    
        sCallName = CStr(vCallNameOrAddress)    'user is hooking an import
    
        'Get handle (address) of exporting module
        lpExportingModuleHandle = GetModuleHandle(sExportingModuleName)
        If lpExportingModuleHandle = 0 Then sError = "Unable to obtain exporting module handle for """ & sExportingModuleName & """.": Exit Function
    
        'Get address of call
        lpProcAddress = GetProcAddress(lpExportingModuleHandle, sCallName)
        If lpProcAddress = 0 Then sError = "Unable to obtain proc address for """ & sCallName & """.": Exit Function
    
    Else
        lpProcAddress = CLng(vCallNameOrAddress) 'user is restoring a hooked import
    End If

    'Beginning of the IAT is located at offset 0x1000 in most PE modules
    vectorIAT = lpImportingModuleHandle + &H1000

    'Scan module to find PE header by looking for header signature
    lpPEHeader = lpImportingModuleHandle
    Do
        If lpPEHeader > vectorIAT Then  'this is not a PE module
            sError = "Module """ & sImportingModuleName & """ is not a PE module."
            Exit Function
        Else
            If Deref(lpPEHeader) = IMAGE_NT_SIGNATURE Then  'we have located the module's PE header
                Exit Do
            Else
                lpPEHeader = lpPEHeader + 1 'keep searching
            End If
        End If
    Loop
    
    'Determine and validate length of the IAT. The length is at offset 0xDC in the PE header.
    lenIAT = Deref(lpPEHeader + &HDC)
    If lenIAT = 0 Or lenIAT > &HFFFFF Then 'its too big or too small to be valid
        sError = "The calculated length of the Import Address Table in """ & sImportingModuleName & """ is not valid: " & lenIAT
        Exit Function
    End If

    'Scan Import Address Table for proc address
    lpEndIAT = lpImportingModuleHandle + &H1000 + lenIAT
    Do
        If vectorIAT > lpEndIAT Then 'we have reached the end of the table
            sError = "Proc address " & Hex(lpProcAddress) & " not found in Import Address Table of """ & sImportingModuleName & """."
            Exit Function
        Else
            lpIATCallAddress = Deref(vectorIAT)
            If lpIATCallAddress = lpProcAddress Then  'we have found the entry
                Exit Do
            Else
                vectorIAT = vectorIAT + 4   'try next entry in table
            End If
        End If
    Loop
    
    'Substitute hook for existing call address and return existing address by ref
    'We must make this memory writable to make the entry in the IAT
    If VirtualProtect(ByVal vectorIAT, 4, PAGE_EXECUTE_READWRITE, lpflOldProtect) = 0 Then
        sError = "Unable to change IAT memory to execute/read/write."
        Exit Function
    Else
        lpOriginalAddress = Deref(vectorIAT)    'save original address
        CopyMemory ByVal vectorIAT, lpHook, 4    'set the hook
        VirtualProtect ByVal vectorIAT, 4, lpflOldProtect, lpflOldProtect2  'restore memory protection
    End If

    HookDLLImport = True 'mission accomplished
Exit Function
    
EH:
    sError = "Unexpected error: " & Err.Description

End Function

Function Deref(lngPointer As Long) As Long  'Equivalent of *lngPointer (returns the value pointed to)
Dim lngValueAtPointer As Long

    CopyMemory lngValueAtPointer, ByVal lngPointer, 4
    Deref = lngValueAtPointer
    
End Function

Function GetAddress(adrof As Long) As Long

    GetAddress = adrof
    
End Function

'Enumerate all modules...
'Usefull for hooking all module's import tables ;)

Function EnumerateModules(proc As Long) As module_list
Dim hMods(1024) As Long, hProcess As Long, cbNeeded As Long, i As Long

    'Get a list of all the modules in this process.
    hProcess = OpenProcess(1040, False, proc)
    If hProcess = 0 Then Exit Function

    If EnumProcessModules(hProcess, hMods(0), 1024 * 4, cbNeeded) Then
     For i = 0 To (cbNeeded / 4)
            Dim pth As String

            'Get the full path to the module's file.
            pth = Space(1024)
            If GetModuleFileNameEx(hProcess, hMods(i), pth, 1024) Then
                With EnumerateModules
                    ReDim Preserve .modules(.count)
                    .modules(.count).id = hMods(i)
                    .modules(.count).name = pth
                    .count = .count + 1
                End With
            End If
        Next i
    End If
    
    CloseHandle hProcess

End Function

Sub AddDllHookEntryToList(list As DllHook_list, item As DllHook_entry)
    
    ReDim Preserve list.items(list.count)
    
    list.items(list.count) = item
    list.count = list.count + 1

End Sub

Function NewDllHook_entry(ToModule As String, DllName As String, FunctionName As String, _
                          FunctionAddress As Long, HookAddress As Long, _
                          State As HookState) As DllHook_entry
                          
    Dim temp As DllHook_entry
    
        With temp
            .ToModule = ToModule
            .DllName = DllName
            .FunctionName = FunctionName
            .FunctionAddress = FunctionAddress
            .HookAddress = HookAddress
            .State = State
        End With
        
    NewDllHook_entry = temp

End Function

Public Function SetHookAndRet(ByVal Todll As String, ByVal dll As String, ByVal funct As String, _
                              ByVal Address As Long, ByRef ErrorString As String) As DllHook_entry
Dim temp As Long

    temp = Hook(Todll, dll, funct, Address, ErrorString)
    If temp = 0 Then
        LogMsg "Unable to set " & funct & " Hook (" & Trim(Todll) & ") ", "modAsmColor", "SetHookAndRet"
    Else
        LogMsg funct & " Hook was set (" & Trim(Todll) & ")", "modAsmColor", "SetHookAndRet"
        SetHookAndRet = NewDllHook_entry(Todll, dll, funct, temp, Address, hooked)
    End If
    
End Function


Sub KillHookList(Hooks As DllHook_list)
Dim temp As Long, strtemp As String, i As Long

    For i = 0 To Hooks.count - 1
        With Hooks.items(i)
            
            If .State = hooked Then
                
                temp = .FunctionAddress
                .State = TrogleHook(.ToModule, .DllName, .FunctionName, .HookAddress, temp, strtemp)
                If temp = 0 Then
                    WarnBox "KillHookList:" & vbNewLine & strtemp, "modHook", "KillHookList"
                    LogMsg "Unable to unset " & .FunctionName & " Hook (" & Trim(.ToModule) & ")", "modHook", "KillHookList"
                Else
                    LogMsg "Killed " & .FunctionName & " Hook (" & Trim(.ToModule) & ")", "modHook", "KillHookList"
                End If
                
            End If
            
        End With
    Next i
    
    Hooks.count = 0

End Sub

Sub CreateHookList(Hooks As DllHook_list, DllName As String, _
                   FunctionName As String, HookAddress As Long, Optional inModule As String)

Dim temp As DllHook_entry, strtemp As String
Dim dlls As module_list, i As Long
    If Len(inModule) = 0 Then
    
        dlls = EnumerateModules(GetCurrentProcessId)
        For i = 0 To dlls.count - 1
            With dlls.modules(i)
            
                'hook ExtTextOutA
                temp = SetHookAndRet(.name, DllName, FunctionName, HookAddress, strtemp)
                If temp.FunctionAddress > 0 Then
                    AddDllHookEntryToList Hooks, temp
                End If

            End With
        Next i
    Else
                temp = SetHookAndRet(inModule, DllName, FunctionName, HookAddress, strtemp)
                If temp.FunctionAddress > 0 Then
                    AddDllHookEntryToList Hooks, temp
                End If
    End If
End Sub

