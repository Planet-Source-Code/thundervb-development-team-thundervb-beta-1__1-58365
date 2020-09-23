Attribute VB_Name = "modCColor"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

'Revision history:
Option Explicit
'24/8/2004[dd/mm/yyyy] : Created by Raziel
'Module Created , intial version
'Coloring Of C code on the vb Ide
'Uses hooked ExtTextOut from AsmColoring

Dim cword As ColorInfo_list, cmode As Long
Dim b_color_c As Boolean

'Called from ExtTextOutHook when text is not asm text..
Function CExtTextOut(ByRef hdc As Long, ByRef x As Long, _
                         ByRef y As Long, ByRef wOptions As Long, _
                         ByRef lpRect As Long, ByRef lpString As Long, _
                         ByRef nCount As Long, ByRef lpDx As Long, ByVal s As String) As Long
                         
    If InStr(1, Trim(s), "'#c'", vbTextCompare) = 1 And Len(s) > 4 And b_color_c = True Then 'line contains asm code
        Dim sof As Long
        cmode = (GetTextColor(hdc) = RGB(255, 255, 255))
        sof = InStr(1, s, "'#c'", vbTextCompare) + 3
        If cmode Then
            SetTextColor hdc, (RGB(255, 255, 255))
        Else
            SetTextColor hdc, RGB(0, 140, 0)
        End If
    
        CExtTextOut = ExtTextOut(hdc, x, y, wOptions, lpRect, lpString, sof, lpDx)
        lpString = lpString + sof: nCount = nCount - sof
        s = Right(s, Len(s) - sof)
        
        CExtTextOut = Draw_C(hdc, x, y, wOptions, lpRect, lpString, nCount, lpDx, s)
    Else
        CExtTextOut = ExtTextOut(hdc, x, y, wOptions, lpRect, lpString, nCount, lpDx)
    End If
    
End Function

'Draws with Coloring a C string
Function Draw_C(ByRef hdc As Long, ByRef x As Long, _
                         ByRef y As Long, ByRef wOptions As Long, _
                         ByRef lpRect As Long, ByRef lpString As Long, _
                         ByRef nCount As Long, ByRef lpDx As Long, ByVal s As String) As Long
    Dim cpos As Long, cmode As Long, out As Col_String, lps As Long
    Dim temp As String, i As Long, temp_l As Long
    cmode = (GetTextColor(hdc) = RGB(255, 255, 255))
    If Len(s) = 0 Then Exit Function
    'What should be considered white space for coloring..
    s = Replace(s, ",", " ") ',
    s = Replace(s, "[", " ") '[
    s = Replace(s, "]", " ") ']
    s = Replace(s, "(", " ") '(
    s = Replace(s, ")", " ") ')
    s = Replace(s, "{", " ") '{
    s = Replace(s, "}", " ") '}
    s = Replace(s, "+", " ") '+
    s = Replace(s, "-", " ") '-
    s = Replace(s, "*", " ") '*
    s = Replace(s, "<", " ") '<,<<
    s = Replace(s, ">", " ") '>,>>
    's = Replace(s, "/", " ") '/ damit this takes comments too
    s = Replace(s, "\", " ") '\
    s = Replace(s, "=", " ") '=
    s = Replace(s, ";", " ") 'and ;
    'possibly more..
    s = ProcStrings(s)
    Do
        temp = GetFirstWordWithSpace(s): RemFirstWordWithSpace s
        temp_l = Len(temp)
        temp = Trim(temp)
        temp_l = temp_l - Len(temp)
        AppendColString out, Len(temp), GetCWordColor(temp)
        If temp_l Then AppendColString out, temp_l, GetCWordColor("")
        If Mid$(s, 1, 2) = "//" Then AppendColString out, Len(s), GetCWordColor("//"): s = ""
    Loop While Len(s)
    
    lps = lpString
    For i = 0 To out.str_index - 2
        If cmode = 0 Then SetTextColor hdc, out.str(i).col
        Draw_C = ExtTextOut(hdc, x, y, wOptions, lpRect, lps, out.str(i).strlen, 0)
        lps = lps + out.str(i).strlen
    Next i
    If cmode = 0 Then SetTextColor hdc, out.str(i).col
    Draw_C = ExtTextOut(hdc, x, y, wOptions, lpRect, lps, out.str(i).strlen - 1, 0)
    
End Function

'No comment...
Function GetCWordColor(word As String) As Long
Dim temp As String, col As Long, i As Long
    

    If Len(word) > 0 Then
        If Mid$(word, 1, 1) = Chr$(34) And Mid$(word, Len(word), 1) = Chr$(34) Then
            temp = "*" & Add34("string") & "*"
            GoTo nochange
        End If
        
        If IsNumeric(word) Then
            temp = "*Number*"
            GoTo nochange
        End If
        
        If Mid$(word, 1, 2) = "0x" Then
            Mid$(word, 1, 2) = "&H"
            If IsNumeric(word) Then
                temp = "*HexNumber*"
                GoTo nochange
            End If
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

    GetCWordColor = col

End Function

'Init C string Coloring Color Table
Sub initCcolors(FromStr As String)
Dim str() As String, str2() As String, i As Long

    LogMsg "Initing C colors", "modCColor", "initCcolors"
    cword.count = 0
    str = Split(FromStr, "_@#slst@_")
    For i = 0 To UBound(str)
        If Len(str(i)) Then
            str2 = Split(str(i), "_@#sent@_")
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
            SendMessage VBI.ActiveCodePane.Window.hWnd, &HF&, 0, 0
            
        End If
    End If
    
End Sub

Sub CColoringEn(bEn As Boolean)

    b_color_c = bEn
    
End Sub
