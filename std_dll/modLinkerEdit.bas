Attribute VB_Name = "modLinkerEdit"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

'Revision history:
'23/8/2004[dd/mm/yyyy] : Created by Raziel
'Module Created , intial version
'
'25/8/2004[dd/mm/yyyy] : Code Edited by Raziel
'Now code uses the Get_Dll function to get setings
'and has more error checking
'
'
'29/8/2004[dd/mm/yyyy] : Edited by Raziel
'WarnBox , ErrorBox and logmsg are used
'
'
'2/9/2004[dd/mm/yyyy] : Edited by Raziel
'Upx , some fixes , log

Dim exe_name As String

Sub linker_edit(ByRef strCLret As String)
Dim strCL As String, i As Long, exports() As String, expnum As Long
    
    LogMsg "Linker Edit", "modLinkerEdit", "linker_edit"
    
    If Get_General(PopUpExportsWindow) Then
        LogMsg "Showning Export Window", "modLinkerEdit", "linker_edit"
        frmSettings.sstSet.Tab = 6
        frmSettings.show vbModal
    End If
    
    LogMsg "exports strings seperation", "modLinkerEdit", "linker_edit"
    
    exports = Split(Get_DLL(ExportedSymbols), "**@@split_@@**")
    If Len(Get_DLL(ExportedSymbols)) > 0 Then
        expnum = UBound(exports)
    Else
        expnum = -1
    End If
    LogMsg expnum & " Exports", "modLinkerEdit", "linker_edit"
    
    strCL = strCLret
    
    If Get_DLL(LinkAsDll) Then  'create dll -> add /dll
        LogMsg "Creating Dll", "modLinkerEdit", "linker_edit"
        strCL = strCL & " /DLL "
    End If
    
    If CBool(Get_DLL(UsePreLoader)) = True Then
    
        If Len(Get_DLL(EntryPointName)) > 0 Then 'change entry point
            WarnBox "User Selected Entrypoint is overwriten by the" & vbNewLine & _
                    "use preloader option.The entry point will be DllMain function", "modLinkerEdit", "linker_edit"
        End If
                
        LogMsg "Changed entry point,With PreLoader", "modLinkerEdit", "linker_edit"
        strCL = Replace(strCL, "/ENTRY:__vbaS", "/ENTRY:PreLoader", , , vbTextCompare)
            
    Else
        If Len(Get_DLL(EntryPointName)) > 0 Then 'change entry point
            LogMsg "Changed entry point", "modLinkerEdit", "linker_edit"
            strCL = Replace(strCL, "/ENTRY:__vbaS", "/ENTRY:" & Get_DLL(EntryPointName), , , vbTextCompare)
        End If
    End If
    
    If Get_DLL(BaseAddress) Then  'change base address
        LogMsg "Changed dll base", "modLinkerEdit", "linker_edit"
        strCL = Replace(strCL, "/BASE:0x400000", "/BASE:0x" & Hex$(Get_DLL(BaseAddress)), , , vbTextCompare)
    End If
    
    If CBool(Get_DLL(ExportSymbols)) = True Then    'export symbols
        LogMsg "Exporting symbols", "modLinkerEdit", "linker_edit"
        If expnum <> -1 Then
            For i = 0 To expnum
                If Len(exports(i)) Then strCL = strCL & " /Export:" & exports(i) & " "
            Next i
        Else
            WarnBox "ExportSymbols is set but No exports are defined", "modLinkerEdit", "LinkerEdit"
        End If
    End If
    
    
    If Len(Get_Paths(LIBFiles_Directory)) > 0 Then
        LogMsg "Using Lib directory : " & Get_Paths(LIBFiles_Directory), "modLinkerEdit", "linker_edit"
        strCL = strCL & "/LIBPATH:" & Add34(Get_Paths(LIBFiles_Directory))
    End If
    
    
    
    exe_name = getS("/OUT:" & Chr$(34), Chr$(34), strCL)
    
    strCLret = strCL
    
    LogMsg "Command line for link : " & strCLret, "modLinkerEdit", "linker_edit"
    
End Sub

Sub linker_edit_after()
    Dim p_cmdline As String, sout As String
    
    If Get_Packer(UsePacker) Then
        LogMsg "Using Packer", "modLinkerEdit", "linker_edit_after"
        If Len(Get_Packer(CommandLine)) > 0 Then
            LogMsg "Command line unproced : " & Get_Packer(CommandLine), "modLinkerEdit", "linker_edit_after"
            p_cmdline = Replace(Get_Packer(CommandLine), "%exename%", exe_name)
            
            LogMsg "Command line proced : " & p_cmdline, "modLinkerEdit", "linker_edit_after"
            
            LogMsg "Packer Exe : " & Get_Paths(Packer) & " " & p_cmdline, "modLinkerEdit", "linker_edit_after"
            
            If Len(Get_Paths(Packer)) > 0 Then
                           
                If ExecuteCommand(Get_Paths(Packer) & " " & p_cmdline, sout) = False Then
                    LogMsg "Packer Failed to run", "modLinkerEdit", "linker_edit_after"
                    WarnBox "Packer Failed to run" & vbNewLine & Get_Paths(Packer) & " " & p_cmdline, "modLinkerEdit", "linker_edit_after"
                Else
                    LogMsg "Packer Ouput" & vbNewLine & sout, "modLinkerEdit", "linker_edit_after"
                End If
        
            Else
                WarnBox "Use packer is Set but packer exe is not set", "modLinkerEdit", "linker_edit_after"
                LogMsg "Error : Use packer is Set but packer exe is not set", "modLinkerEdit", "linker_edit_after"
            End If
        Else
            WarnBox "Use packer is Set but no packer command line is set", "modLinkerEdit", "linker_edit_after"
            LogMsg "Error : Use packer is Set but no packer command line is set", "modLinkerEdit", "linker_edit_after"
        End If
    End If
End Sub
