Attribute VB_Name = "modCreateProcHook"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit
'Revision history:
'20/8/2004[dd/mm/yyyy] : Created by Raziel
'Module Created , intial version
'
'22/9/2004[dd/mm/yyyy] : Edited by Raziel
'Cahnges for new hooking code
'

Public Cancel_compile As Boolean, ProjectSaved As Boolean, LinkerHookState As HookState
Global sName As String, sLine As String, sDir As String, file_vb As String
Public old_CreateProc As Long

'hooked function
Public Function CreateProcess_Hook(lpApplicationName As Long, lpCommandLine As Long, _
                                    lpProcessAttributes As SECURITY_ATTRIBUTES, _
                                    lpThreadAttributes As SECURITY_ATTRIBUTES, _
                                    ByVal bInheritHandles As Long, _
                                    ByVal dwCreationFlags As Long, _
                                    lpEnvironment As Long, _
                                    ByVal lpCurrentDirectory As Long, _
                                    lpStartupInfo As STARTUPINFO, _
                                    lpProcessInformation As PROCESS_INFORMATION) As Long
                                    
    If Not (VBI Is Nothing) Then
        If Not (VBI.ActiveVBProject Is Nothing) Then
            ProjectSaved = VBI.ActiveVBProject.Saved
        End If
    End If
       
    'get the needed info
    sName = CStringZero(lpApplicationName)
    sLine = CStringZero(lpCommandLine)
    file_vb = getS("-f " & Chr(34), Chr(34), sLine)
    sDir = CStringZero(lpCurrentDirectory)
    
    LogMsg "CreateProcess_Hook Called(" & sName & "," & sLine & "," & file_vb & "," & sDir & ")", "modCreateProcHook", "CreateProcess_Hook"
    
    If UCase(Left(sLine, 4)) = "LINK" Then
        linker_edit sLine 'if we link, edit the linker command line ( std dll, exports,base)
        If Get_Compile(PauseBeforeLinking) = True And Get_Compile(ModifyCmdLine) = False Then
            MsgBoxX "We are going to Link"
        End If
    End If
    
    'well if we must skip this..(user canceled compile)
    If Cancel_compile = True Then GoTo nocall
    
    If UCase(Left(sLine, 2)) = "C2" Then
        c2_edit sLine 'if we compile, edit the c2 command line
        If has_asm Then LogMsg "File Has Asm/C", "modCreateProcHook", "CreateProcess_Hook"
    End If
    
    'User - modify cmd line
    If Get_Compile(ModifyCmdLine) Then
        frmViewer.ShowViewer "Command Line Edit", sLine, False
        'MsgBoxX "Edit command line here " & vbNewLine & sLine
        LogMsg "new Command line [user edit] : " & sLine, "modCreateProcHook", "CreateProcess_Hook"
    End If
    
    CreateProcess_Hook = CreateProcess(sName, sLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, sDir, lpStartupInfo, lpProcessInformation)
    WaitToEnd lpProcessInformation.dwProcessId
    
    'if needed,we assemble asm code (the generated asm listing)
    If UCase(Left(sLine, 2)) = "C2" Then
        c2_edit_after sLine
    End If
        
    If UCase(Left(sLine, 4)) = "LINK" Then
        linker_edit_after 'if we linked , do any after link editing needed (Packer , ect)
    End If
        
        
    Exit Function
nocall:
    Cancel_compile = False 'new compile after that

End Function



'edit the compiler command line if needed (the module has Asm/C code)
Sub c2_edit(str As String)
Dim Fix_unnamed As Boolean
    
    Fix_unnamed = CBool(Get_ASM(FixASMListings))
    file_obj = getS("-Fo" & Chr(34), Chr(34), str) 'obj file
    file_asm = Replace(file_obj, ".obj", ".asm", , , vbTextCompare) 'output asm file
    If InStr(1, LoadFile(file_vb), "'#asm'", vbTextCompare) Or InStr(1, LoadFile(file_vb), "#asm_start", vbTextCompare) Or InStr(1, LoadFile(file_vb), "'#c'", vbTextCompare) Then  'module has asm code
    If CBool(Get_ASM(CompileASMCode)) = False Then GoTo noasm
        'set the cmd line to create a asm file instead of an object file
        If Fix_unnamed Then
            str = Replace(str, "-Fo" & Add34(file_obj), "-FAsc -Fa" & Add34(file_asm))
        Else
            str = Replace(str, "-Fo" & Add34(file_obj), "-FAs -Fa" & Add34(file_asm))
        End If
        If ProjectSaved = False Then
        'MsgBoxX "Project Must be saved for asm changes to have effect", , vbOKOnly Or vbInformation
        If MsgBoxX("Project Must be saved for asm changes to have effect" & vbNewLine & _
               "You want to save it now?", , vbYesNo Or vbInformation) = vbYes Then
                SaveProjects True
        End If
                   
        End If
        has_asm = 1
    Else
noasm:
        has_asm = 0
        If CBool(Get_C(CompileCCode)) = True And CBool(Get_ASM(CompileASMCode)) = False Then
            WarnBox "Asm MUST be enabled to compile C code..", "modCreateProcHook", "c2_edit"
        End If
    End If

End Sub

'Assemble the asm listing (if it was created from the c2_edit)
Sub c2_edit_after(str As String)
    
    If has_asm Then 'well we have asm
        ProcInlineAsm str
    End If
    
End Sub

'Init the CreateProcessHook
Sub InitLinkerHook()
Dim temp As Long, strtemp As String

    temp = Hook("VBA" & vb_Dll_version & ".DLL", "kernel32.dll", "CreateProcessA", AddressOf CreateProcess_Hook, strtemp)
    If temp = 0 Then
        MsgBox "InitLinkerHook:" & vbNewLine & strtemp
        LogMsg "Unable to set CreateProc Hook", "modCreateProcHook", "InitLinkerHook"
    Else
        old_CreateProc = temp
        LogMsg "CreateProc Hook was set", "modCreateProcHook", "InitLinkerHook"
    End If
    LinkerHookState = hooked
    
End Sub

'Trogle Hook on/off
Function TrogleLinkerHook() As HookState
Dim temp As Long, strtemp As String

    temp = old_CreateProc
    TrogleLinkerHook = TrogleHook("VBA" & vb_Dll_version & ".DLL", "kernel32.dll", "CreateProcessA", AddressOf CreateProcess_Hook, temp, strtemp)
    LinkerHookState = TrogleLinkerHook
    If temp = 0 Then
        MsgBox "TrogleLinkerHook:" & vbNewLine & strtemp
        LogMsg "Unable to trogle CreateProc Hook", "modCreateProcHook", "TrogleLinkerHook"
    Else
        old_CreateProc = temp
        LogMsg "CreateProc Hook was trogled", "modCreateProcHook", "TrogleLinkerHook"
    End If
    
End Function

'Turn off hook
Sub KillLinkerHook()

    If LinkerHookState = hooked Then TrogleLinkerHook
    LogMsg "CreateProc Hook was unset", "modCreateProcHook", "KillLinkerHook"
    
End Sub


