Attribute VB_Name = "modAsmColor"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

'Revision history:
'20/8/2004[dd/mm/yyyy] : Created by Raziel
'Module Created , intial version
'Coloring of Asm Code on the VB ide
'Hooks ExtTextOut and then uses SetTextColor to colorise the asm line
'
'Note :Many fixes are made but not listed here..
'
'22/9/2004[dd/mm/yyyy] : Edited by Raziel
'Major Changes on the hooking code , now we hook A and W versions
'on all modules loaded , in a safe way
'coloring is done on any string drawen from any control..
'exept ones drawing transparent Text (like textbox ) cause it is not possible to
'draw the string corectly..
'
'23/9/2004[dd/mm/yyyy] : Edited by Raziel
'Fixed some small bugs
'
'24/10/2004[dd/mm/yyyy] : Edited by Raziel
'Fixed some more small bugs
'Added WHookAlso option to disable the W hooks if they are not wanted ..

Option Explicit

#Const nonExHook = True ' hook and TextOut
#Const WHookAlso = False ' do not hook and ExTextOutW,TextOutW

Dim Hooks As DllHook_list, errstring As String
Dim oldY As Long, inblock As Long

Public AsmColorHookState As HookState

Dim cword As ColorInfo_list
Dim b_color_asm As Boolean

Dim s_full As String, s_offset As Long, curline As Long


Sub InitAsmColorHook()

    CreateHookList Hooks, "gdi32", "ExtTextOutA", AddressOf ExtTextOutAHook
    
    #If WHookAlso Then
        CreateHookList Hooks, "gdi32", "ExtTextOutW", AddressOf ExtTextOutWHook
    #End If
    
    #If nonExHook Then
        CreateHookList Hooks, "gdi32", "TextOutA", AddressOf TextOutAHook
        #If WHookAlso Then
            CreateHookList Hooks, "gdi32", "TextOutW", AddressOf TextOutWHook
        #End If
    #End If
    
    LogMsg Hooks.count & " hooks were set", "modAsmColor", "InitAsmColorHook"
    
    AsmColorHookState = hooked
    
End Sub

'They are still here for "backup" purposes...
#If 0 Then
    
    Sub InitAsmColorHook_o()
    Dim temp As Long, strtemp As String
        
        temp = Hook("VBA" & vb_Dll_version & ".DLL", "gdi32", "ExtTextOutA", AddressOf ExtTextOutAHook, strtemp)
        If temp = 0 Then
            MsgBox "InitAsmColorHook:" & vbNewLine & strtemp
            LogMsg "Unable to set ExtTextOut Hook", "modAsmColor", "InitAsmColorHook"
        Else
            LogMsg "ExtTextOut Hook was set", "modAsmColor", "InitAsmColorHook"
            oldExtTextOutA = temp
        End If
        
        #If nonExHook Then
            temp = Hook("msvbvm60.dll", "gdi32", "TextOutA", AddressOf TextOutHook, strtemp)
            If temp = 0 Then
                MsgBox "InitAsmColorHook:" & vbNewLine & strtemp
                LogMsg "Unable to set TextOut Hook", "modAsmColor", "InitAsmColorHook"
            Else
                LogMsg "TextOut Hook was set", "modAsmColor", "InitAsmColorHook"
                oldTextOutA = temp
            End If
        #End If
        
        AsmColorHookState = hooked
        
    End Sub
    
    Function TrogleAsmColorHook_o() As HookState
    Dim temp As Long, strtemp As String
    
        temp = oldExtTextOutA
        TrogleAsmColorHook = TrogleHook("VBA" & vb_Dll_version & ".DLL", "gdi32", "ExtTextOutA", AddressOf ExtTextOutAHook, temp, strtemp)
        AsmColorHookState = TrogleAsmColorHook
        If temp = 0 Then
            MsgBox "TrogleAsmColorHook:" & vbNewLine & strtemp
            LogMsg "Unable to trogle ExtTextOut Hook", "modAsmColor", "TrogleAsmColorHook"
        Else
            LogMsg "ExtTextOut Hook was Trogled", "modAsmColor", "TrogleAsmColorHook"
            oldExtTextOutA = temp
        End If
        
        #If nonExHook Then
            temp = oldTextOutA
            Call TrogleHook("msvbvm60.dll", "gdi32", "TextOutA", AddressOf TextOutHook, temp, strtemp)
            If temp = 0 Then
                MsgBox "TrogleAsmColorHook:" & vbNewLine & strtemp
                LogMsg "Unable to trogle TextOut Hook", "modAsmColor", "TrogleAsmColorHook"
            Else
                LogMsg "TextOut Hook was Trogled", "modAsmColor", "TrogleAsmColorHook"
                oldTextOutA = temp
            End If
        #End If
        
    End Function
    
    Sub killAsmColorHook_o()
    
        If AsmColorHookState = hooked Then TrogleAsmColorHook
        LogMsg "ExtTextOut/TextOut Hook was unset", "modAsmColor", "KillAsmColorHook"
            
    End Sub
#End If

Sub KillAsmColorHook()

    KillHookList Hooks
    AsmColorHookState = unhooked
    
End Sub
'ExtTextOutA Hook - Asm Coloring
Function ExtTextOutAHook(ByVal hdc As Long, ByVal x As Long, _
                         ByVal y As Long, ByVal wOptions As Long, _
                         ByVal lpRect As Long, ByVal lpString As Long, _
                         ByVal nCount As Long, ByVal lpDx As Long) As Long
'  HDC hdc,          // handle to DC
'  int X,            // x-coordinate of reference point
'  int Y,            // y-coordinate of reference point
'  UINT fuOptions,   // text-output options
'  CONST RECT* lprc, // optional dimensions
'  LPCTSTR lpString, // string
'  UINT cbCount,     // number of characters in string
'  CONST INT* lpDx   // array of spacing values (not used as far as i know here)
Dim s As String, cpos As Long, cmode As Long, oldCol As Long, oldTA As Long, oldP As POINTAPI
    'If Not (VBI Is Nothing) Then _
        'If Not (VBI.ActiveWindow Is Nothing) Then _
            'If VBI.ActiveWindow.Type <> vbext_wt_CodeWindow Then GoTo draw_direct
    'If Not (VBI Is Nothing) Then _
    '    If Not (VBI.ActiveCodePane Is Nothing) Then _
    '        If Not (VBI.ActiveCodePane.CodeModule Is Nothing) Then _
    '            If y > 0 Then _
    '                curline = VBI.ActiveCodePane.TopLine + ((y - 30) / 16) + 1: _
    '                If curline > 0 And curline <= VBI.ActiveCodePane.CodeModule.CountOfLines Then _
    '                    s_full = VBI.ActiveCodePane.CodeModule.lines(curline, 1)
    'If MustNotBeColored(hDC) Then
    '    ExtTextOutAHook = ExtTextOut(hDC, x, y, wOptions, lpRect, lpString, nCount, lpDx)
    '    Exit Function
    'End If
    
    s = Cstring(lpString, nCount)
    
    's_offset = 1 - InStr(1, s, s, vbTextCompare)
    
    cmode = (GetTextColor(hdc) = RGB(255, 255, 255))
    
    'LogMsg s & " " & lpString & " " & nCount, "modAsmColor", "ExtTextOutAHook"
    
    oldCol = GetTextColor(hdc)
    oldTA = GetTextAlign(hdc)
    
    If (oldTA And 1) = 0 Then
        MoveToEx hdc, x, y, VarPtr(oldP)
        SetTextAlign hdc, oldTA Or 1
    End If
    'ExtTextOutAHook = ExtTextOut(hDC, x, y, wOptions, lpRect, lpString, nCount, lpDx)
    'if opaque then draw with no coloring
    If ((wOptions And 2) = 2) Then GoTo Draw_NoColors

    If Len(Trim(s)) = 0 Then GoTo Draw_NoColors
    
    If InStr(1, Trim(s), "'#ASM'", vbTextCompare) = 1 And Len(s) > 5 And b_color_asm = True Then 'line contains asm code
        Dim sof As Long
        sof = InStr(1, s, "'#ASM'", vbTextCompare) + 5
        If cmode Then
            SetTextColor hdc, (RGB(255, 255, 255))
        Else
            SetTextColor hdc, RGB(0, 140, 0)
        End If
    
        ExtTextOutAHook = ExtTextOut(hdc, x, y, wOptions, lpRect, lpString, sof, lpDx)
        lpString = lpString + sof: nCount = nCount - sof
        s = Right(s, Len(s) - sof)
        ExtTextOutAHook = Draw_Asm(hdc, x, y, wOptions, lpRect, lpString, nCount, lpDx, s)
    ElseIf Inline(curline - 1) = True And Len(s) > 1 Then ' tooo slow for now..
        If cmode Then
            SetTextColor hdc, (RGB(255, 255, 255))
        Else
            SetTextColor hdc, RGB(0, 140, 0)
        End If
        ExtTextOutAHook = ExtTextOut(hdc, x, y, wOptions, lpRect, lpString, 1, lpDx)
        lpString = lpString + 1: nCount = nCount - 1
        s = Right(s, Len(s) - 1)
        ExtTextOutAHook = Draw_Asm(hdc, x, y, wOptions, lpRect, lpString, nCount, lpDx, s)
    Else
        ExtTextOutAHook = CExtTextOut(hdc, x, y, wOptions, lpRect, lpString, nCount, lpDx, s)
    End If
    
    SetTextColor hdc, oldCol
        
    If (oldTA And 1) = 0 Then
        SetTextAlign hdc, oldTA
        MoveToEx hdc, oldP.x, oldP.y, 0
    End If
    
    Exit Function
Draw_NoColors:
    
    'SetTextColor hDC, RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
    ExtTextOutAHook = ExtTextOut(hdc, x, y, wOptions, lpRect, lpString, nCount, lpDx)
    
    SetTextColor hdc, oldCol
        
    If (oldTA And 1) = 0 Then
        SetTextAlign hdc, oldTA
        MoveToEx hdc, oldP.x, oldP.y, 0
    End If
End Function

 Function ExtTextOutWHook( _
     ByVal hdc As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal wOptions As Long, _
     ByVal lpRect As Long, _
     ByVal lpString As Long, _
     ByVal nCount As Long, _
     ByVal lpDx As Long) As Long
     Dim tempA() As Byte

    If nCount Then
     
        tempA = StrConv(CstringW(lpString, nCount), vbFromUnicode)
        'Dim temp As String
        'temp = CstringW(lpString, nCount)
        'temp = "text " & temp & " text"
        'MsgBox temp
        ExtTextOutWHook = ExtTextOutAHook(hdc, x, y, wOptions, lpRect, VarPtr(tempA(0)), nCount, lpDx)
    Else
        ExtTextOutWHook = ExtTextOutW(hdc, x, y, wOptions, lpRect, lpString, nCount, lpDx)
    End If

End Function

Function TextOutWHook(ByVal hdc As Long, _
                      ByVal x As Long, _
                      ByVal y As Long, _
                      ByVal lpString As Long, _
                      ByVal nCount As Long) As Long
                      
    Dim tempA() As Byte

    If nCount Then
     
        tempA = StrConv(CstringW(lpString, nCount), vbFromUnicode)
        'Dim temp As String
        'temp = CstringW(lpString, nCount)
        'temp = "text " & temp & " text"
        'MsgBox temp
        TextOutWHook = TextOutAHook(hdc, x, y, VarPtr(tempA(0)), nCount)
    Else
        TextOutWHook = TextOutW(hdc, x, y, lpString, nCount)
    End If

End Function

Function Draw_Asm(ByRef hdc As Long, ByRef x As Long, _
                         ByRef y As Long, ByRef wOptions As Long, _
                         ByRef lpRect As Long, ByRef lpString As Long, _
                         ByRef nCount As Long, ByRef lpDx As Long, ByVal s As String) As Long
    Dim cpos As Long, cmode As Long, out As Col_String, lps As Long
    Dim temp As String, i As Long, temp_l As Long
    cmode = (GetTextColor(hdc) = RGB(255, 255, 255))
    If Len(s) = 0 Then Exit Function
    s = Replace(s, ",", " ")
    s = Replace(s, "[", " ")
    s = Replace(s, "]", " ")
    s = Replace(s, "{", " ")
    s = Replace(s, "}", " ")
    s = Replace(s, "+", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, ";", ";")
    s = ProcStrings(s)
    Do
        temp = GetFirstWordWithSpace(s): RemFirstWordWithSpace s
        temp_l = Len(temp)
        temp = Trim(temp)
        temp_l = temp_l - Len(temp)
        
        AppendColString out, Len(temp), GetAsmWordColor(temp)
        If temp_l Then AppendColString out, temp_l, GetAsmWordColor("")

        
        If Mid$(s, 1, 1) = ";" Then
            out.str(out.str_index - 1).strlen = out.str(out.str_index - 1).strlen - 1
            AppendColString out, Len(s) + 1, GetAsmWordColor(";")
            s = ""
        End If
    Loop While Len(s)
    
    lps = lpString
    For i = 0 To out.str_index - 2
        If cmode = 0 Then SetTextColor hdc, out.str(i).col
        Draw_Asm = ExtTextOut(hdc, x, y, wOptions, lpRect, lps, out.str(i).strlen, 0)
        lps = lps + out.str(i).strlen
    Next i
    If cmode = 0 Then SetTextColor hdc, out.str(i).col
    Draw_Asm = ExtTextOut(hdc, x, y, wOptions, lpRect, lps, out.str(i).strlen - 1, 0)
    
End Function

Function Draw_Asm_old(ByRef hdc As Long, ByRef x As Long, _
                         ByRef y As Long, ByRef wOptions As Long, _
                         ByRef lpRect As Long, ByRef lpString As Long, _
                         ByRef nCount As Long, ByRef lpDx As Long, ByVal s As String) As Long
    Dim cpos As Long, cmode As Long

    cmode = (GetTextColor(hdc) = RGB(255, 255, 255))
    
    If Len(s) > 0 Then
        cpos = InStr(1, s, ";")
        If cpos = 0 Then
            If cmode Then
                SetTextColor hdc, (RGB(255, 255, 255))
            Else
                SetTextColor hdc, RGB(0, 0, 140)
            End If
            Draw_Asm_old = ExtTextOut(hdc, x, y, wOptions, lpRect, lpString, nCount, 0)
        Else
            If cmode Then
                SetTextColor hdc, (RGB(255, 255, 255))
            Else
                SetTextColor hdc, RGB(0, 0, 140)
            End If
            Draw_Asm_old = ExtTextOut(hdc, x, y, wOptions, lpRect, lpString, cpos - 1, 0)
            If cmode Then
                SetTextColor hdc, (RGB(255, 255, 255))
            Else
                SetTextColor hdc, RGB(0, 140, 0)
            End If
            Draw_Asm_old = ExtTextOut(hdc, x, y, wOptions, lpRect, lpString + cpos - 1, nCount + 1 - cpos, 0)
        End If
    End If

End Function

'if we are in asm code block
Function Inline(ByVal line As Long) As Boolean
Dim curl As Long, sC As String, s_pos As Long, e_pos As Long

    'tooo slooowww for now...
    Exit Function
    If (VBI Is Nothing) Then Exit Function
    If (VBI.ActiveCodePane Is Nothing) Then Exit Function
    If (VBI.ActiveCodePane.codeModule Is Nothing) Then Exit Function
    If line < 1 Then Exit Function
    If line > VBI.ActiveCodePane.codeModule.CountOfLines Then Exit Function
    
    sC = VBI.ActiveCodePane.codeModule.lines(1, line)
    
    s_pos = InStrRev(sC, "#asm_start", , vbTextCompare)
    e_pos = InStrRev(sC, "#asm_end", , vbTextCompare)
    If s_pos > 0 And s_pos > e_pos Then
        Inline = True
    Else
        Inline = False
    End If
    
End Function

Function GetAsmWordColor(word As String) As Long
Dim temp As String, col As Long, i As Long
    
    If Len(word) > 0 Then
        If Mid$(word, 1, 1) = Chr$(34) And Mid$(word, Len(word), 1) = Chr$(34) Then
            temp = "*" & Add34("string") & "*"
            GoTo nochange
        End If
        
        If Mid$(word, 1, 1) = "'" And Mid$(word, Len(word), 1) = "'" Then
            temp = "*'string'*"
            GoTo nochange
        End If
        
        If IsNumeric(word) Then
            temp = "*Number*"
            GoTo nochange
        End If
    End If
    
    temp = " " & Trim(word) & " "
nochange:
    If Len(temp) = 2 Then temp = "*default*"
    For i = 0 To cword.count - 1
        If InStr(1, cword.ColorInfo(i).str, temp, vbTextCompare) Then col = cword.ColorInfo(i).Color: Exit For
    Next i
    
    If col = 0 Then
        temp = "*default*"
        For i = 0 To cword.count - 1
            If InStr(1, cword.ColorInfo(i).str, temp, vbTextCompare) Then col = cword.ColorInfo(i).Color
        Next i
    End If
    
    GetAsmWordColor = col

End Function

Sub initAsmColors(FromStr As String)
Dim str() As String, str2() As String, i As Long
    
    LogMsg "Initing Asm colors", "modAsmColor", "initAsmcolors"
    cword.count = 0
    str = Split(FromStr, "_@#slst@_")
    For i = 0 To UBound(str)
        If Len(str(i)) Then
            str2 = Split(str(i), "_@#sent@_")
            'AddColor str2(0), Val(str2(1))
            ReDim Preserve cword.ColorInfo(cword.count)
            cword.ColorInfo(cword.count).str = str2(0)
            cword.ColorInfo(cword.count).Color = val(str2(1))
            cword.count = cword.count + 1
        End If
    Next i
    
    If Not (VBI Is Nothing) Then
        If Not (VBI.ActiveCodePane Is Nothing) Then
        Dim old_ As Long
        
            'temporary solution..
            old_ = VBI.ActiveCodePane.TopLine
            If VBI.ActiveCodePane.CountOfVisibleLines > 0 Then
                VBI.ActiveCodePane.TopLine = VBI.ActiveCodePane.CountOfVisibleLines
            End If
            VBI.ActiveCodePane.TopLine = old_
            'hmm this seems not to work...
            'HWnd is hiden but exists (f2 , show hiden members)
            SendMessage VBI.ActiveCodePane.Window.hWnd, 16, 0, ByVal 0
            
        End If
    End If
    
End Sub

Sub AsmColoringEn(bEn As Boolean)

    b_color_asm = bEn
    'If b_color_asm = False Then KillAsmColorHook
    
End Sub

Public Function TextOutAHook(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
'  HDC hdc,          // handle to DC
'  int X,            // x-coordinate of reference point
'  int Y,            // y-coordinate of reference point
'  LPCTSTR lpString, // string
'  UINT cbCount,     // number of characters in string

    TextOutAHook = ExtTextOutAHook(hdc, x, y, 0, 0, lpString, nCount, 0)
    
End Function
