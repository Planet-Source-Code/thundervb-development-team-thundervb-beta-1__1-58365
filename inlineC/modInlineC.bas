Attribute VB_Name = "modInlineC"
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
'Basic functionality , code designed
'
'29/8/2004[dd/mm/yyyy] : Edited by Raziel
'Now it uses logmsg to log messages if log is enabled..
'
'6/9/2004[dd/mm/yyyy]  : Edited by Raziel
'Code Changed to be more compitable, now it realy folows the code design
'Better logging , Crash Free
'
'30/10/2004[dd/mm/yyyy]  : Edited by Raziel
'Fixed some bugs here and there
'
'Notes:
'hmm not so inline , rather in function
'it seems that the best way to do this is not to merge it with vb code on the
'same function...

Public Type c_code_asm
    fName As String 'vb name
    fnameC As String 'C name
    AsmCode As String ' asm listing
End Type


Dim file_asm_in As String, file_c As String, file_asm_out As String
'Proccess inline C
Sub procInlineC(asmfile As String)
Dim cl_exe As String

    If CBool(Get_C(CompileCCode)) = True Then
    
        cl_exe = Get_Paths(CCompiler)
        file_asm_in = asmfile                                    'name.asm
        file_asm_out = Left(asmfile, Len(asmfile) - 4) & "C.asm" 'so nameC.asm
        file_c = Left(file_asm_out, Len(file_asm_out) - 3) & "c" 'so nameC.c
        InsertCcode
        kill2 file_c
        kill2 file_asm_out
        
    Else
    
        LogMsg "C is not enabled", "modInlineC", "procInlineC"
        
    End If
    
End Sub

'Insert the C code on a vb module listing
Sub InsertCcode()
Dim asm() As String, i As Long, temp As String, temp2 As c_code_asm
Dim l1 As Long, tmp_l As String
    
    asm = Split(LoadFile(file_asm_in), vbNewLine)
    LogMsg "Inserting C Code", "modInlineC", "InsertCcode"
    LogMsg UBound(asm) & " lines", "modInlineC", "InsertCcode"
    For i = 0 To UBound(asm)
        If Len(asm(i)) > 0 Then
            temp = GetLineFromLising(asm(i))
            If InStr(1, temp, "'#c'", vbTextCompare) = 1 Then
                LogMsg i & " : " & temp, "modInlineC", "InsertCcode"
                For l1 = i To 0 Step -1
                    If Len(asm(l1)) Then
                        If InStr(1, asm(l1), "PUBLIC") Then
                            tmp_l = Replace(Trim(asm(l1)), vbTab, "    ")
                            RemFisrtWord tmp_l: tmp_l = Trim(tmp_l)
                            temp2.fName = GetFirstWord(tmp_l)
                            Exit For
                        End If
                    End If
                Next l1
                If Len(temp2.fName) = 0 Then
                    ErrorBox "Error while looking for Public directive", "modInlineC", "InsertCcode"
                    Exit Sub
                End If
                RemFisrtWord temp: temp = Trim(temp)
                temp2.fnameC = GetFirstWord(Replace(temp, "(", " "))
                temp2.AsmCode = GetRemAsmCode(asm, i)
                asm(i) = AssembleCCode(temp2)
            End If
        End If
    Next i
    
    LogMsg "saving result", "modInlineC", "InsertCcode"
    
    SaveFile file_asm_in, Join$(asm, vbNewLine)
    DoEvents

    LogMsg "saved", "modInlineC", "InsertCcode"
    
End Sub

'Complie to assebly the C code
Function AssembleCCode(from As c_code_asm) As String
Dim c_code As String, s_out As String, temp As String, cl_exe As String, fl_ As String

    cl_exe = Get_Paths(CCompiler)
    SaveFile file_c, GetCCode(from.AsmCode)
    LogMsg "Assembling C Code", "modInlineC", "AssembleCCode"
    If Len(cl_exe) > 0 And FileExist(cl_exe) Then
    
        
        If ExecuteCommand(Add34(cl_exe) & " " & Add34(file_c) & " /G6  /Ox  /Gz /GA /FAs /Fa" & Add34(file_asm_out) & " /c /I" & Add34(Get_Paths(INCFiles_Directory)), s_out, GetPath(cl_exe)) = False Then
            ErrorBox "Can't Execute C compiler (cl.exe):" & vbNewLine & cl_exe, "modInlineC", "AssembleCCode"
            GoTo ext:
        End If
        
        
        If FileExist(file_asm_out) = False Then
            frmCError.ShowError s_out, LoadFile(file_c)
        Else
            If FileLen(file_asm_out) = 0 Then
                frmCError.ShowError s_out, LoadFile(file_c)
            End If
        End If
        
        If FileExist(file_asm_out) = False Then
            GoTo ext:
        Else
            If FileLen(file_asm_out) = 0 Then
                GoTo ext:
            End If
        End If
        
        fl_ = LoadFile(file_asm_out)
        temp = GetAsmCodeFromCListing(fl_)
        AssembleCCode = Replace(temp, getCFunctName(temp, from.fnameC), from.fName)
        kill2 file_asm_out
        kill2 file_c
    Else
        ErrorBox "cl.exe Not found" & vbNewLine & _
                   "make sure that the cl.exe path is corect on the addin settings" & vbNewLine & _
                   "The curect cl.exe path is : " & cl_exe, "modInlineC", "AssembleCCode"
ext:
        If MsgBoxX("Cancel Compile?", "cl.exe Error", vbYesNo Or vbQuestion) = vbYes Then
                   Cancel_compile = True
                   Exit Function
        End If
    End If

    
End Function

'Get's the c code from a processed vb asm listing
Function GetCCode(asm As String) As String
Dim lines() As String, out As String_B, i As Long, temp As String

    lines = Split(asm, vbNewLine)
    For i = 0 To UBound(lines)
        temp = Trim(lines(i))
        'proccess it
        If InStr(1, temp, "'#c'", vbTextCompare) = 1 Then ' c code line
            AppendString out, Right(temp, Len(temp) - Len("'#c'")) & vbNewLine
        End If
    Next i
    
    GetCCode = GetString(out)
End Function

'Gets asm code from c listing
Function GetAsmCodeFromCListing(asm As String) As String
Dim i As Long, temp As String, tmp2() As String
        
    i = InStr(1, asm, "INCLUDELIB", vbTextCompare)
    temp = Mid$(asm, i, Len(asm) - i)
    tmp2 = Split(temp, vbNewLine)
    ReDim Preserve tmp2(UBound(tmp2) - 1)
    GetAsmCodeFromCListing = Join$(tmp2, vbNewLine)
    
End Function

'Get's the function's name in C (adding the _name@#)
Function getCFunctName(asm As String, nameold As String) As String
Dim i As Long, nam As String, i1 As Long

    i = InStr(1, asm, "_" & nameold & "@", vbTextCompare) + Len("_" & nameold)
    
    nam = "_" & nameold & Mid$(asm, i, 1): i = i + 1
    Do
        nam = nam & Mid$(asm, i, 1): i = i + 1
    Loop While IsNumeric(Mid$(asm, i, 1))
    getCFunctName = nam
    
End Function

'Get a line from a asm listing
Function GetLineFromLising(str As String) As String
Dim temp As String

    temp = str
  If GetFirstWord(temp) = ";" Then 'starts with a ";"
        RemFisrtWord temp: temp = Trim(temp)
        If IsNumeric(GetFirstWord(temp)) Then 'the second word is  number
            RemFisrtWord temp: temp = Trim(temp)
            If GetFirstWord(temp) = ":" Then ' and the third word is a ":"
                RemFisrtWord temp: temp = Trim(temp)
                GetLineFromLising = temp ' return it :)
            End If
        End If
    End If
    
End Function

'Gets and removes a function that contains C from VB listing
Function GetRemAsmCode(str() As String, ByVal i As Long) As String
Dim curseg As String, i2 As Long, out As String_B, seg_s As Long
    'This  must be chnged too
    'Look for word "SEGMENT" and in the line containing it , take it's first word [segname SEGMENT]
    
    For i2 = i To 0 Step -1
        If Len(str(i2)) Then
            If InStr(1, str(i2), "SEGMENT") Then
                curseg = GetFirstWord(Replace(Trim(str(i2)), vbTab, "    "))
                seg_s = i2
                Exit For
            End If
        End If
    Next i2
    
    If Len(curseg) = 0 Then
        ErrorBox "Error while looking for SEGMENT directive", "modInlineC", "GetRemAsmCode"
        Exit Function
    End If
    
    For i2 = i To UBound(str)
        If InStr(1, str(i2), curseg & vbTab & "ENDS", vbTextCompare) = 1 Then
            str(i2) = ""
            Exit For
        Else
            AppendString out, GetLineFromLising(str(i2)) & vbNewLine
            str(i2) = ""
        End If
    Next i2
    
    For i2 = seg_s - 2 To i - 1
        str(i2) = ""
    Next i2
    
    GetRemAsmCode = GetString(out)

End Function

'Retry to compile C code
Sub retryC(temp_s As String)
Dim c_code As String, s_out As String, temp As String, cl_exe As String, fl_ As String

    cl_exe = Get_Paths(CCompiler)
    SaveFile file_c, temp_s
    
    If Len(cl_exe) > 0 And FileExist(cl_exe) Then
    
        
        If ExecuteCommand(Add34(cl_exe) & " " & Add34(file_c) & " /G6  /Ox  /Gz /GA /FAs /Fa" & Add34(file_asm_out) & " /c /I" & Add34(Get_Paths(INCFiles_Directory)), s_out, GetPath(cl_exe)) = False Then
            ErrorBox "Can't Execute C compiler (cl.exe):" & vbNewLine & cl_exe, "modInlineC", "RetryC"
        End If
        
        If FileExist(file_asm_out) = False Then
            frmCError.ShowError s_out, LoadFile(file_c)
        Else
            If FileLen(file_asm_out) = 0 Then
                frmCError.ShowError s_out, LoadFile(file_c)
            End If
        End If
    
    End If
    
End Sub

'Info / How this code works
'
'asm listing :
'
'
'PUBLIC  ?Killed@Module1@@AAGXXZ             ; Module1::Killed        ; line number : 195
';   COMDAT ?Killed@Module1@@AAGXXZ       ; line number : 196
'text$1  SEGMENT      ; line number : 197
'?Killed@Module1@@AAGXXZ PROC NEAR           ; Module1::Killed, COMDAT        ; line number : 198
'         ; line number : 199
'; 94   : '#c'int Killed(int i,int i2){       ; line number : 200
'; 95   : '#c'i*=i2;      ; line number : 201
'; 96   : '#c'i2*=i;      ; line number : 202
'; 97   : '#c'return i*i2;        ; line number : 203
'; 98   : '#c'}       ; line number : 204
'; 99   : End Function        ; line number : 205
'         ; line number : 206
'    xor eax, eax         ; line number : 207
'    ret 8        ; line number : 208
'?Killed@Module1@@AAGXXZ ENDP                ; Module1::Killed        ; line number : 209
';replace end;
'text$1  ENDS         ; line number : 210
'END      ; line number : 211
'
'
'we want to replace ?Killed@Module1@@AAGXXZ with the c code..
'
'so at fisrt we will select this code and kill coments
'
'
'
'PUBLIC  ?Killed@Module1@@AAGXXZ
'
'text$1  SEGMENT
'?Killed@Module1@@AAGXXZ PROC NEAR
'
'; 94   : '#c'int Killed(int i,int i2){
'; 95   : '#c'i*=i2;
'; 96   : '#c'i2*=i;
'; 97   : '#c'return i*i2;
'; 98   : '#c'}
'; 99   : End Function
'
'    xor eax, eax
'    ret 8
'?Killed@Module1@@AAGXXZ ENDP
'
'text$1  ENDS
'
'
'
' then we will find the function name and boundires
' Fist Public before '#c' code is function name(or 5 lines up)
'
'PUBLIC  [funct name in VB]
' we mark this as replace start
'
' then at the 3rd line up we have SEGMENT name:
' [segname] SEGMENT
' so we need [segname] ENDS to mark replace end
'c function name is the
' second word of th first c line
'; 94   : '#c'int Killed(int i,int i2){
'; 94   : '#c'[functype] [funcname][shtis]
'we compile all c code to a listing
'cl "c:\main.c" /G6 /Gz /GA /FAs /Fa"C:\AsmOut" /c
'meaning
'compile [infile] /ppro /stdcall /OptimizeForWindowsApplication /Asmwithsource /asmfile[file] /compileonly
'we skip everything til the first include lib
'
'
'Result:
'
'INCLUDELIB LIBC
'INCLUDELIB OLDNAMES

'PUBLIC  _foo
'_DATA   SEGMENT
'_foo    DD  022H
'_DATA   ENDS
'PUBLIC  _smaple@0
'EXTRN   _printf:NEAR
'; Function compile flags: /Ods
'; File c:\main.c
'_TEXT   SEGMENT
'_i2$ = -8
'_i$ = -4
'_smaple@0 PROC NEAR
'
'; 9    : {
'
'    push ebp
'    mov ebp, esp
'    push ecx
'    push ecx
'
'; 10   :    int i=foo;
'
'    mov eax, DWORD PTR _foo
'    mov DWORD PTR _i$[ebp], eax
'
'; 11   :    int i2=foo*i;
'
'    mov eax, DWORD PTR _foo
'    imul    eax, DWORD PTR _i$[ebp]
'    mov DWORD PTR _i2$[ebp], eax
'
'; 12   :    i++;
'
'    mov eax, DWORD PTR _i$[ebp]
'    inc eax
'    mov DWORD PTR _i$[ebp], eax
'
'; 13   :    i*=100;
'
'    mov eax, DWORD PTR _i$[ebp]
'    imul    eax, 100                ; 00000064H
'    mov DWORD PTR _i$[ebp], eax
'
'; 14   :    printf(i);
'
'    push    DWORD PTR _i$[ebp]
'    call    _printf
'    pop ecx
'
'; 15   :    return 0;
'
'    xor eax, eax
'
'; 16   : }
'
'    leave
'    ret 0
'_smaple@0 ENDP
'_TEXT   ENDS
'
'here we find the _functname string and read the next chars until we have a non numeric char
'after we replace everywhere "_[functname][foundchars] with [funct name in VB]
'finaly we replace the selected replace reange in vb asm listing with the text
'that was genereated from the c listing, removing the last line from the c listing (the "end")

'INCLUDELIB LIBC
'INCLUDELIB OLDNAMES

'PUBLIC  _foo
'_DATA   SEGMENT
'_foo    DD  022H
'_DATA   ENDS
'PUBLIC  ?Killed@Module1@@AAGXXZ
'EXTRN   _printf:NEAR
'; Function compile flags: /Ods
'; File c:\main.c
'_TEXT   SEGMENT
'_i2$ = -8
'_i$ = -4
'?Killed@Module1@@AAGXXZ PROC NEAR
'
'; 9    : {
'
'    push ebp
'    mov ebp, esp
'    push ecx
'    push ecx
'
'; 10   :    int i=foo;
'
'    mov eax, DWORD PTR _foo
'    mov DWORD PTR _i$[ebp], eax
'
'; 11   :    int i2=foo*i;
'
'    mov eax, DWORD PTR _foo
'    imul    eax, DWORD PTR _i$[ebp]
'    mov DWORD PTR _i2$[ebp], eax
'
'; 12   :    i++;
'
'    mov eax, DWORD PTR _i$[ebp]
'    inc eax
'    mov DWORD PTR _i$[ebp], eax
'
'; 13   :    i*=100;
'
'    mov eax, DWORD PTR _i$[ebp]
'    imul    eax, 100                ; 00000064H
'    mov DWORD PTR _i$[ebp], eax
'
'; 14   :    printf(i);
'
'    push    DWORD PTR _i$[ebp]
'    call    _printf
'    pop ecx
'
'; 15   :    return 0;
'
'    xor eax, eax
'
'; 16   : }
'
'    leave
'    ret 0
'?Killed@Module1@@AAGXXZ ENDP
'_TEXT   ENDS

