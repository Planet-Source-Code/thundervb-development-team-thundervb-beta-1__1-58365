Attribute VB_Name = "modInlineAsm"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit
Declare Function disasm_vb Lib "ndisasm_dll.dll" (ByRef dat As Byte, ByRef strout As Byte) As Long

'Revision history:
'23/8/2004[dd/mm/yyyy] : Created by Raziel
'Module Created , intial version
'Coded Hook and TrogleHook functions
'
'25/8/2004[dd/mm/yyyy] : Edited by Raziel
'Many bug fixes , code has full error checking
'Changes For InlineC
'
'26/8/2004[dd/mm/yyyy] : Edited by Raziel
'Assembly Listing Unnamed variables fixing , partial
'Using a Customized version of nDisAsm (in dll) to disassemble
'
'29/8/2004[dd/mm/yyyy] : Edited by Raziel
'Many bugfixes to prevent crashes, even better error handling
'Added logmsg calls to log if enbaled
'Incude file needed is copied to the directory [strIncFileData]
'
'29/8/2004[dd/mm/yyyy] : Edited by Raziel
'More crashes fixed [numberAsmLines]
'
'12/10/2004[dd/mm/yyyy] : Edited by Raziel
'GPF exeption code
'

Global file_asm As String, file_obj As String, has_asm As String, asm As Boolean, file_include As String
Sub ProcInlineAsm(str As String)
Dim sMasmOut As String, masm_exe As String, temp_s As String
On Error GoTo errH:
GPF_Set GPF_RaiseErr, "modInlineAsm", "ProcInlineAsm"
        masm_exe = Get_Paths(ml)
        LogMsg "Inserting Asm Code", "modInlineAsm", "ProcInlineAsm"
        If FileExist(file_asm) = False Then
            If FileLen(file_asm) = False Then
                ErrorBox "Asm Listing Does not exist", "modInlineAsm", "ProcInlineAsm"
                LogMsg "Error:Asm Listing Does not exist", "modInlineAsm", "ProcInlineAsm"
                GoTo CleanUp
            End If
        End If
        
        file_include = GetPath(file_asm) & "listing.inc"
        SaveFile file_include, strIncFileData
        DoEvents ' to save the file
        temp_s = LoadFile(file_asm)
        'fix the asm file (compitable with masm >v5.10)
        'so it can be compiled,do any other fixing is needed
        FixAsm temp_s
        'proccess the asm blocks
        ProcAsmBlocks temp_s
        'Insert the asm code
        temp_s = Replace(temp_s, "'#asm'", vbNewLine, , , vbTextCompare)
        'ProcAsmLines temp_s
        'output the processed asm listing
        SaveFile file_asm, temp_s
        DoEvents
        'inlineC
        procInlineC file_asm
        If Cancel_compile = True Then GoTo CleanUp
        'reload the file..
        temp_s = LoadFile(file_asm)
        'number the lines for easy debuging
        LogMsg "Numbering lines..", "modInlineAsm", "ProcInlineAsm"
        'NumberAsmLines temp_s
        LogMsg "ReSaving", "modInlineAsm", "ProcInlineAsm"
        'now save it for masm..
        SaveFile file_asm, temp_s
        DoEvents
        LogMsg "Assebling", "modInlineAsm", "ProcInlineAsm"
        'execute masm if it exists
        Dim tempM As String
        'Pasue if needed
        If Get_Compile(PauseBeforeAssembly) = True And Get_Compile(ModifyCmdLine) = False Then
            MsgBoxX "We are going to Assemble"
        End If
        
        If (Len(masm_exe) > 0) And FileExist(masm_exe) Then
            tempM = Add34(masm_exe) & " /c /Cp /coff " & Add34(file_asm)
            If ExecuteCommand(tempM, sMasmOut, GetPath(masm_exe)) = False Then GoTo error_masm_exe
            
        Else
error_masm_exe:
            ErrorBox "Masm Not found" & vbNewLine & _
                   "make sure that the masm Path is corect on the addin settings" & vbNewLine & _
                   "The curect masm path is : " & masm_exe, "modInlineAsm", "ProcInlineAsm"
            If MsgBoxX("Cancel Compile?", "Masm Error", vbYesNo Or vbQuestion) = vbYes Then Cancel_compile = True
            GoTo CleanUp
        End If

        If FileExist(GetPath(masm_exe) & GetFilename(file_obj)) = False Then
            'oh well masm din't created the .obj file , so assembling error
            frmMasmError.ShowError tempM & vbNewLine & sMasmOut, LoadFile(file_asm)
        End If
        
        If Cancel_compile = True Then GoTo CleanUp
        
        If CopyFile(GetPath(masm_exe) & GetFilename(file_obj), file_obj, True) <> 0 Then
            kill2 GetPath(masm_exe) & GetFilename(file_obj)
        Else
            ErrorBox "Can't copy file to target directory" & vbNewLine & _
                   "From : " & GetPath(masm_exe) & GetFilename(file_obj) & vbNewLine & _
                   "To    : " & file_obj, "modInlineAsm", "ProcInlineAsm"
        End If
CleanUp:        'cleanup and exit
        kill2 file_asm
        kill2 file_include
GPF_Reset
Exit Sub
errH:
    ErrorBox Err.Description, "modInlineAsm", "ProcInlineAsm"
    GoTo CleanUp

End Sub

'Do any fixes required to make the listing assemblamle corectly
Sub FixAsm(ByRef temp As String)
Dim Fix_unnamed As Boolean
    
    Fix_unnamed = CBool(Get_ASM(FixASMListings))
    If Fix_unnamed Then
        LogMsg "Warning:Asm Listing Fix is on , this should work most times but it is in BETA stage", "modInlineAsm", "ProcInlineAsm"
        WarnBox "Asm Listing Fix is on , this should work most times but it is in BETA stage", "modInlineAsm", "ProcInlineAsm"
        FixAsm_unnamed_vars temp
    End If
    'make it comp with mas >510
    FixAsm_masm temp

End Sub

'add this :
'if @Version gt 510...
'else....
'endif <- we look for this and we add this :
'OPTION NOSCOPE         ;Added By ThunVB to fix label problems
'ASSUME  CS: FLAT, DS: FLAT, ES: FLAT, FS: FLAT, GS: FLAT, SS: FLAT ;Added by ThunVB
Sub FixAsm_masm(ByRef str As String) 'Make the file compitable with masm > than v5.1
Dim i As Long, lines() As String

    lines = Split(str, vbNewLine)
    'for all the lines
    For i = 0 To UBound(lines)
        If InStr(1, lines(i), "endif", vbTextCompare) = 1 Then 'we must do it here
        lines(i) = lines(i) & vbNewLine & _
                   "OPTION NOSCOPED ;Added By ThunVB to fix label problems" & vbNewLine & _
                   "ASSUME  CS: FLAT, DS: FLAT, ES: FLAT, FS: FLAT, GS: FLAT, SS: FLAT ;Added by ThunVB"
                   
            Exit For
        End If
    Next i
    str = Join$(lines, vbNewLine)

End Sub

'Not totaly corect , relies on nDisAsm dll
'fix unamed vars compiler bug
'This works by striping out the compile code bytes generated from the -FAsc
'and dissasebling them , to get the corect asm instuctions
Sub FixAsm_unnamed_vars(ByRef str As String)
Dim i As Long, lines() As String, temp As String, oldlen As Long, temp2 As String

    lines = Split(str, vbNewLine)
    'for all the lines
    For i = 0 To UBound(lines)
        oldlen = Len(lines(i))
        temp2 = GetFirstWord$(RTrim$(Trim$(Replace(lines(i), vbTab, " "))))
        
        If (Len(temp2) = 5) And IsNumeric("&h" & temp2) Then
            temp = FixAsm_hexs_line(lines(i))
            'well if the codebytes are extended and to the next line..
            If Len(lines(i)) = 0 And (oldlen > 0) Then
                i = i + 1
                temp = temp & " " & FixAsm_hexs_line(lines(i))
            End If
            If InStr(1, lines(i), "unnamed") Then
                lines(i) = GetAsmFromOpcode(temp) & "  ;" & lines(i)
            End If
        End If
    Next i
    str = Join$(lines, vbNewLine)

End Sub

'removes the hexs numbers before instruction and returns them
Function FixAsm_hexs_line(ByRef str As String) As String 'fix one line
Dim out As String_B

    str = Replace(str, vbTab, "    ")
    'we want to remove the numbers that the vb compiler
    'adds on the listings..
    'they are on the format :
    '[hhhhh] [hh] [hh] [hh] [hh] ... [asm command]
    'where h means a hex digit [ ]  means that it is optional
    If Mid$(str, 1, 1) = ";" Then GoTo ext 'well this line is a comment
    str = Trim(str) 'remove any space in the start/end of the string
    Do While IsNumeric("&h" & GetFirstWord$(str)) = True And Len(str) > 0 'this is a valid hex number and the string still exists
        If Len(GetFirstWord$(str)) = 3 Then GoTo ext 'no hex is in three numbers, but posssibly a "add","daa" ect
        AppendString out, GetFirstWord$(str) & " "
        RemFisrtWord str 'remove the fisrt word
        str = Trim(str) 'remove any space
    Loop
ext:
    If out.str_index > 0 Then
        FixAsm_hexs_line = GetString(out)
    End If
    
End Function

Sub ProcAsmBlocks(ByRef str As String) 'edit the asm blocks so that they are recognised
Dim i As Long, lines() As String, temp As String, asm As Boolean
    
    lines = Split(str, vbNewLine)
    'we look for lines in the folowing format ; number : other data
    For i = 0 To UBound(lines)
        temp = lines(i): temp = Trim(temp)
        If GetFirstWord(temp) = ";" Then 'starts with a ";"
            RemFisrtWord temp: temp = Trim(temp)
            If IsNumeric(GetFirstWord(temp)) Then 'the second word is  number
                RemFisrtWord temp: temp = Trim(temp)
                If GetFirstWord(temp) = ":" Then ' and the third word is a ":"
                    RemFisrtWord temp: temp = Trim(temp)
                    ProcAsmBlocks_line temp, asm, lines(i), i 'proccess it
                End If
            End If
        End If
    Next i
    str = Join$(lines, vbNewLine)
    
End Sub

Sub ProcAsmLines(ByRef str As String) 'uncoment '#asm' lines
Dim i As Long, lines() As String, temp As String, asm As Boolean
    
    lines = Split(str, vbNewLine)
    'we look for lines in the folowing format ; number : other data
    For i = 0 To UBound(lines)
        temp = lines(i): temp = Trim(temp)
        If GetFirstWord(temp) = ";" Then 'starts with a ";"
            RemFisrtWord temp: temp = Trim(temp)
            If IsNumeric(GetFirstWord(temp)) Then 'the second word is  number
                RemFisrtWord temp: temp = Trim(temp)
                If GetFirstWord(temp) = ":" Then ' and the third word is a ":"
                    RemFisrtWord temp: temp = Trim(temp)
                    If InStr(1, temp, "'#asm'", vbTextCompare) = 1 Then ' if asm line
                        lines(i) = Right(temp, Len(temp) - 6) & vbTab & lines(i)
                    End If
                End If
            End If
        End If
    Next i
    str = Join$(lines, vbNewLine)
    
End Sub

'Covert a asm block to many '#asm' prefixed instructions
Sub ProcAsmBlocks_line(ByRef str As String, ByRef asm As Boolean, ByRef sout As String, i As Long)

    'if it is a coment then
    If Mid(str, 1, 1) <> "'" Then Exit Sub
    'remove the coment and any space after it
    str = Right(str, Len(str) - 1)
    str = Trim(str)
    'if asm_start/end command
    If InStr(1, str, "#asm_start", vbTextCompare) = 1 Then asm = True: Exit Sub
    If InStr(1, str, "#asm_end", vbTextCompare) = 1 Then asm = False: Exit Sub
    'if inside asm block
    If asm = False Then Exit Sub
    'and if the line is not a comand (eg '#asm')
    If Mid(str, 1, 1) = "#" Then Exit Sub
    'expand the block to many '#asm' lines so that they can be recognised
    str = "; asm block expanded  : '#asm' " & str
    'write the output
    sout = str

End Sub

Sub NumberAsmLines(ByRef strg As String) 'number the lines for easy bug finding

    Dim i As Long
    Dim lines() As String
    lines = Split(strg, vbNewLine)
    LogMsg UBound(lines) & " lines", "modInlineAsm", "NumberAsmLines"
    'number the lines for easy bug finding '104 leters space ->104/8 -> 13 tabs
    For i = 0 To UBound(lines)
        lines(i) = lines(i) & String$(IIf(Len(lines(i)) > 104, Len(lines(i)) + 10, 104) - Len(lines(i)), " ") & " ; line number : " & i + 1
        'LogMsg i & " : " & lines(i), "modInlineAsm", "NumberAsmLines"
    Next i
    
    strg = Join$(lines, vbNewLine)

End Sub


'retry to assmeble the asm code
Sub retryAsm(temp_s As String)
Dim sMasmOut As String, masm_exe As String

    masm_exe = Get_Paths(ml)
    SaveFile file_asm, temp_s
        
    If ExecuteCommand(Add34(masm_exe) & " /c /Cp /coff " & Add34(file_asm) _
                                    , sMasmOut, GetPath(masm_exe)) Then
    
    Else
        ErrorBox "Masm Not found", "modInlineAsm", "retryAsm"
    End If
    
    If FileExist(GetPath(masm_exe) & GetFilename(file_obj)) = False Then
        frmMasmError.ShowError Add34(masm_exe) & " /c /Cp /coff " & Add34(file_asm) & vbNewLine & sMasmOut, LoadFile(file_asm)
    End If
    
End Sub


Function GetAsmFromOpcode(ByVal opcode As String) As String
Dim ff As Long, sout As String, sopcode As String, temp As Long

    If Len(RTrim$(Trim$(opcode))) = 0 Then Exit Function
    'ff = FreeFile
    RemFisrtWord opcode ' the 5 digit offset..
    sopcode = opcode
    'Using Nasm for now..
    GetAsmFromOpcode = GetAsmFromOpcodeNasm(sopcode)
    Exit Function
    'Ok this is the vb6 verion but it is incomplete...
    'Nasm can do the work much better and with full compatibilyty..
    'and very good speed (when in dll)
    temp = val("&H" & GetFirstWord$(opcode))
    RemFisrtWord opcode
    Select Case temp
    
        '884424  #1|#|        mov [esp+<#1>],al
        '884C24  #1|#|        mov [esp+<#1>],cl
        '885424  #1|#|        mov [esp+<#1>],dl
        '885C24  #1|#|        mov [esp+<#1>],bl
        Case &H88
        
            temp = val("&H" & GetFirstWord$(opcode)) * 255
            RemFisrtWord opcode
            temp = temp + val("&H" & GetFirstWord$(opcode))
            RemFisrtWord opcode
                
            Select Case temp
                
            '4424  #1|#|        mov [esp+<#1>],al
            Case 17444
                sout = "mov [esp+" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],al"
                
            '4C24  #1|#|        mov [esp+<#1>],cl
            Case 19492
                sout = "mov [esp+" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],cl"
                
            '5424  #1|#|        mov [esp+<#1>],dl
            Case 21540
                sout = "mov [esp+" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],dl"
                
            '5C24  #1|#|        mov [esp+<#1>],bl
            Case 23588
                sout = "mov [esp+" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],bl"
                
            Case Else
            GoTo lble
            
            End Select
            
        '894424  #1|#|        mov [esp+<#1>],eax
        '8945    #1|#|        mov [ebp-<#1>],eax
        '894C24  #1|#|        mov [esp+<#1>],ecx
        '894D    #1|#|        mov [ebp-<#1>],ecx
        '895424  #1|#|        mov [esp+<#1>],edx
        '895C24  #1|#|        mov [esp+<#1>],ebx
        '897424  #1|#|        mov [esp+<#1>],esi
        '8975    #1|#|        mov [ebp-<#1>],esi
        Case &H89
        
               
            temp = val("&H" & GetFirstWord$(opcode))
            RemFisrtWord opcode

            Select Case temp
                
            '4424  #1|#|        mov [esp+<#1>],eax
            Case &H44
                If GetFirstWord$(opcode) = "24" Then
                    RemFisrtWord opcode
                    sout = "mov [esp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],eax"
                Else
                    GoTo lble:
                End If
                    
            '45    #1|#|        mov [ebp-<#1>],eax
            Case &H45
                sout = "mov [ebp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],eax"
                
            '4C24  #1|#|        mov [esp+<#1>],ecx
            Case &H44
                If GetFirstWord$(opcode) = "24" Then
                    RemFisrtWord opcode
                    sout = "mov [esp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],ecx"
                Else
                    GoTo lble:
                End If
                
            '4D    #1|#|        mov [ebp-<#1>],ecx
            Case &H44
                sout = "mov [ebp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],ecx"
                
            '5424  #1|#|        mov [esp+<#1>],edx
            Case &H44
                If GetFirstWord$(opcode) = "24" Then
                    RemFisrtWord opcode
                    sout = "mov [esp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],edx"
                Else
                    GoTo lble:
                End If
                
            '5C24  #1|#|        mov [esp+<#1>],ebx
            Case &H44
                If GetFirstWord$(opcode) = "24" Then
                    RemFisrtWord opcode
                    sout = "mov [esp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],ebx"
                Else
                    GoTo lble:
                End If
                
            '7424  #1|#|        mov [esp+<#1>],esi
            Case &H44
                If GetFirstWord$(opcode) = "24" Then
                    RemFisrtWord opcode
                    sout = "mov [esp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],esi"
                Else
                    GoTo lble:
                End If
                
            '75    #1|#|        mov [ebp-<#1>],esi
            Case &H44
                sout = "mov [ebp+" & SignedB(val("&H" & GetFirstWord$(opcode))) & "],esi"
                
            Case Else
            GoTo lble
            
            End Select
            
        
        '8B45    #1|#|        mov eax,[ebp-<#1>]
        Case &H8B
        '45    #1|#|        mov eax,[ebp-<#1>]
        If GetFirstWord$(opcode) = "45" Then
            RemFisrtWord opcode
            sout = "mov eax,[ebp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "]"
        Else
            GoTo lble:
        End If
        
        '8D4424  #1|#|        lea eax,[esp+<#1>]
        '8D45    #1|#|        lea eax,[ebp-<#1>]
        '8D4C24  #1|#|        lea ecx,[esp+<#1>]
        '8D4D    #1|#|        lea ecx,[ebp-<#1>]
        '8D5424  #1|#|        lea edx,[esp+<#1>]
        '8D55    #1|#|        lea edx,[ebp-<#1>]
        Case &H8D
        temp = val("&H" & GetFirstWord$(opcode))
            RemFisrtWord opcode

            Select Case temp
                
            '4424  #1|#|        lea eax,[esp+<#1>]
            Case &H44
                If GetFirstWord$(opcode) = "24" Then
                    RemFisrtWord opcode
                    sout = "lea eax,[esp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "]"
                Else
                    GoTo lble:
                End If
            '45    #1|#|        lea eax,[ebp-<#1>]
            Case &H45
                sout = "lea eax,[ebp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "]"
            '4C24  #1|#|        lea ecx,[esp+<#1>]
            Case &H4C
                If GetFirstWord$(opcode) = "24" Then
                    RemFisrtWord opcode
                    sout = "lea ecx,[esp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "]"
                Else
                    GoTo lble:
                End If
            '4D    #1|#|        lea ecx,[ebp-<#1>]
            Case &H4D
                sout = "lea ecx,[ebp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "]"
            '5424  #1|#|        lea edx,[esp+<#1>]
            Case &H54
                If GetFirstWord$(opcode) = "24" Then
                    RemFisrtWord opcode
                    sout = "lea edx,[esp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "]"
                Else
                    GoTo lble:
                End If
            '55    #1|#|        lea edx,[ebp-<#1>]
            Case &H55
                sout = "lea edx,[ebp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "]"
            Case Else
            GoTo lble
            
            End Select
            
        'C64424  #1 #1|#|     mov byte ptr [esp+<#1>],<@2>
        Case &HC6
        'C64424  #1 #1|#|     mov byte ptr [esp+<#1>],<@2>
        GoTo lble
        
        'C74424  #1 #4|#|     mov dword ptr [esp+<#1>],<@2>
        'C745    #1 #4|#|     mov dword ptr [ebp-<#1>],<@2>
        Case &HC7
        temp = val("&H" & GetFirstWord$(opcode))
            RemFisrtWord opcode

            Select Case temp
                
            '4424  #1 #4|#|     mov dword ptr [esp+<#1>],<@2>
            Case &H44
                If GetFirstWord$(opcode) = "24" Then
                    RemFisrtWord opcode
                    sout = "mov dword ptr [esp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "]"
                    RemFisrtWord opcode
                    sout = sout & "," & val("&H" & Split(opcode, " ")(3) & Split(opcode, " ")(2) & Split(opcode, " ")(1) & Split(opcode, " ")(0))
                Else
                    GoTo lble:
                End If
            
            '45    #1 #4|#|     mov dword ptr [ebp-<#1>],<@2>
            Case &H45
             
                    sout = "mov dword ptr [ebp" & SignedB(val("&H" & GetFirstWord$(opcode))) & "]"
                    RemFisrtWord opcode
                    sout = sout & "," & val("&H" & Split(opcode, " ")(3) & Split(opcode, " ")(2) & Split(opcode, " ")(1) & Split(opcode, " ")(0))
                
            Case Else
            GoTo lble
            
            End Select

        Case Else
lble:
        sout = GetAsmFromOpcodeNasm(sopcode)
        GoTo nmsg
    End Select
    
    'MsgBox sout & vbNewLine & GetAsmFromOpcodeNasm(sopcode) & vbNewLine & sopcode
nmsg:
    GetAsmFromOpcode = sout
End Function

'Use nasm to dissassemble
Function GetAsmFromOpcodeNasm(ByVal opcode As String) As String
Dim temp(23) As Byte, sout() As Byte, tind As Long, sout2 As String

    
    'write the data
    ReDim sout(255)
    Do
        temp(tind) = CByte(val("&h" & GetFirstWord(opcode)))
        tind = tind + 1
        RemFisrtWord opcode: opcode = Trim$(opcode)
    Loop While Len(opcode)
    disasm_vb temp(0), sout(0)
    ReDim Preserve sout(Find0(sout) - 1)
    sout2 = StrConv(sout, vbUnicode)
    sout2 = replace0xwith0yyyh(Trim$(sout2))
    
    If Len(sout2) = 0 Then ErrorBox "Can't fix asm listing.." & vbNewLine & sout2, "modInlineAsm", "GetAsmFromOpcodeNasm"

    
    GetAsmFromOpcodeNasm = sout2
End Function

Function replace0xwith0yyyh(str As String) As String
Dim temp As String, pos As Long, l As Long, tmp As String, s As Long
    s = 1
rech:
    pos = InStr(s, str, "0x", vbTextCompare)
    l = 1
    Do While IsNumeric("&h" & Mid$(str, pos + 2, l)) And (pos + 2 + l) < (Len(str) + 2)
        l = l + 1
    Loop
    l = l - 1
    If pos > 0 Then
        tmp = val("&H" & Mid$(str, pos + 2, l))
        If Len(tmp) <= Len(Mid$(str, pos, l + 2)) Then
            tmp = Space((Len(Mid$(str, pos, l)) - Len(tmp)) + 2) & tmp
        Else
            ErrorBox tmp & vbNewLine & str, "modInlineAsm", "replace0xwith0yyyh"
        End If
        Mid$(str, pos) = tmp
        s = pos + 2 + l
        GoTo rech
    End If
    str = Replace(str, "dword", "dword PTR")
    replace0xwith0yyyh = str
    
End Function

'sign extend a byte
Function SignedB(val As Long) As String

    If val > 127 Then val = val - 256
    
    If val >= 0 Then
        SignedB = "+" & val
    Else
        SignedB = val
    End If
    
End Function

'Find 0 in a asni C string (as a byte array)
Function Find0(str() As Byte) As Long
Dim i As Long
    
    For i = 0 To UBound(str)
        If str(i) = 0 Then Exit For
    Next i
    Find0 = i
    
End Function

