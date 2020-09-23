Attribute VB_Name = "modIntelliAsm"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit
'Revision history:
'6/10/2004[dd/mm/yyyy] : Created by Raziel
'Module Created , intial version
'Added most function prototypes
'Basic functionality
'
'8/10/2004 using isml files for tips and lists
'10/10/2004 Fixed a lot of things , added intelliasm* subs..
'20-25/10/2004 Many changes , acuracy is much better..
'              isml code is imroved a lot too , to suport all the needed features...
'30/10/2004 Fixes , Fixes and Bug fixes ;)
'           Also , comments ;D


'TODO : Local Vars, Labels..
'       Test it a lot
'       Create a big isml file..[90% done]

Public dat As Isml_File

Public Type LTI
    ttInfo As String
    ListInfo As String
    curLVL As Long
End Type

Dim style As Isml_kw_types

'Complete a text , by replacing the curent word with the strtest word
Sub CompleteText(strTest As String)
Dim temp As Long, slin As Long, scol As Long, tlin As String
Dim we As Long, ws As Long, b1 As String, b2 As String

    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveCodePane Is Nothing Then Exit Sub
    If VBI.ActiveCodePane.codeModule Is Nothing Then Exit Sub
    
    'If we must complete a number ... we must not do anything ..
    If strTest = "#" Then Exit Sub
    'GEt hte curent cursor pos
    VBI.ActiveCodePane.GetSelection slin, scol, 0, 0
    tlin = VBI.ActiveCodePane.codeModule.lines(slin, 1)
    tlin = Replace$(Replace$(tlin, "dword ptr ", "dword_ptr_"), "byte ptr ", "byte_ptr_")
    tlin = Replace$(Replace$(tlin, "word ptr ", "word_ptr_"), "qword ptr ", "qword_ptr_")
    GetWordFormPos slin, scol, ws, we 'Get the curent word ...
    ws = ws - 1
    'And edit the line so that it is replaced by the new word
    With VBI.ActiveCodePane.codeModule
        
        If (ws) > 0 Then
            b1 = Left(tlin, ws)
        End If
        If (Len(tlin) - we) > 0 Then
            b2 = Right(tlin, Len(tlin) - we)
        End If
        b1 = b1 & strTest
        .ReplaceLine slin, b1 & b2 'write the line back :)
    End With
    
    'Seet the new cursos position
    VBI.ActiveCodePane.SetSelection slin, Len(b1) + 1, slin, Len(b1) + 1

End Sub

'Get the curent Code context , for asm tooltips ans asm lists...
Function GetCurContext() As String
Dim curline As String, slin As Long, scol As Long
'Error handling
On Error GoTo errH
GPF_Set GPF_RaiseErr, "modIntelliAsm", "GetCurContext"

    'Get the curent line
    VBI.ActiveCodePane.GetSelection slin, scol, 0, 0
    scol = scol - 1
    If scol > 0 Then
        curline = LCase(VBI.ActiveCodePane.codeModule.lines(slin, 1))
        If scol < Len(curline) Then
            curline = Left(curline, scol)
        End If
    Else
        curline = ""
    End If
    
    'Check if we are on asm code
    If InStr(1, Trim$(curline), "'#asm'") = 0 Then
        GetCurContext = "none" 'no in asm so no popup :)
        GoTo Exit_Function
    End If
    
    ''Check if we are on asm code with a " " after the text
    'If StrComp(Trim$(curline), "'#asm'", vbTextCompare) Then
    '    If (scol + 1) < Len(curline) Then
    '        GetCurContext = "none" 'no space no popup :)
    '        goto Exit_Function
    '    End If
    'End If
    
    'well , we are just in asm context , so show a list of everything
    If (Len(curline) - (InStr(1, curline, "'#asm'") + 6)) < 1 Then
        GetCurContext = "asml1" 'display all kw's
        GoTo Exit_Function
    End If
    
    'Remove spaces , #asm' ect...
    curline = RTrim$(Right$(curline, Len(curline) - (InStr(1, curline, "'#asm'") + 6)) & "_")
    
    'If Mid$(curline, Len(curline) - 1, 1) <> " " Then
    
    curline = Left$(curline, Len(curline) - 1)
    
    'End If
    'Proccess strings so that we cannot misstake them for anything else..
    ProcStrings curline
    'expand comas
    curline = Replace(curline, ",", " , ")
    
    If Len(Trim$(curline)) = 0 Then
        
        GetCurContext = "asml1" ' nothing typped just spaces..
        GoTo Exit_Function
        
    End If
    
    If InStrRev(curline, ";") Then 'too bad .. we are in comment..
        GetCurContext = "comment" 'comment - nothing to show
        GoTo Exit_Function
    End If
    
    'frmDConsole.AppendLog Add34(curline)
    'Ohh here is the action (and the bugs :D)
    'We are on asm code , and the user has typed something ..
    'We must find out what and give a tooltip and/or a list ...
    Dim wa() As String, i As Long, wordC As Long, l As Long, out As String_B

    'To detect ptr's as single words...
    curline = Replace$(Replace$(curline, "dword ptr ", "dword_ptr_"), "byte ptr ", "byte_ptr_")
    curline = Replace$(Replace$(curline, "word ptr ", "word_ptr_"), "qword ptr ", "qword_ptr_")
    'Split the string on words...
    wa = GetAllWordsToArr(curline)
    wordC = UBound(wa) + 1 'Word count
    
        
    If wordC < 2 Then 'nah.. not much type so propably still on the first word
        GetCurContext = "asml1" 'default to this case the user is still on teh first word
    Else
        GetCurContext = "none" 'defalt to none since the user has allready typed the command
    End If
    
    'find the best match from our database...
    For i = 0 To dat.kw_count - 1
        With dat.kw(i)
            If ((.count > wordC And (.Isml_kw_Type And kw_PopUpList) > 0) Or _
               ((.count >= wordC) And ((.Isml_kw_Type And kw_Tipform) > 0))) And (wordC > 1) Then
                                                ' if it is possible to match this..
                Dim lp As Long
                Dim lps As Long
                For l = 0 To wordC - 1 'check all teh words
                    With .Matches(l)
                        lp = lps 'shits for <any><..><..> matching
                        If isMatch(wa(l), dat.kw(i).Matches(l), dat.kw(i).Isml_kw_Type And kw_Tipform, lps) = 0 Then
                            GoTo NextOne 'well we din't mached , goto next one..
                        Else 'shits for <any><..><..> matching
                            If lp <> 0 Then lps = 0
                            lp = 0
                        End If
                    End With
                Next l
                
                'if it reached here , all the things matched up :)
                If l >= 1 Then 'we compared something...
                    'so we can must take on count the result...
                    AppendString out, CStr(i)
                End If
            End If
        End With
NextOne:
    Next i
    If out.str_index > 0 Then 'if any matches at all then
        FinaliseString out    'procces
        GetCurContext = CStr(wordC + lps) & "|" & Join$(out.str, "|")
        If lps < 0 Then ' if a <any><prv> macth , then no list should be visible
            'the " " will tell that on the GetListFromContext routine ;)
            GetCurContext = " " & GetCurContext
        End If
    End If
    
'Error handling..
Exit_Function:
GPF_Reset
Exit Function
errH:
    MsgBox Err.Description & vbNewLine & _
           Err.Source
GPF_Reset
End Function

'Creates a list from teh curent Context (that is found using GetCurContext above ;))
Function GetListFromContext(con As String, style As Isml_kw_types) As LTI
Dim tl() As Long, tls() As String, i As Long, t As Isml_File, bhl As Boolean
'Error handling
On Error GoTo errH
GPF_Set GPF_RaiseErr, "modIntelliAsm", "GetListFromContext"
    
    'if this context is meant not to be for a list..
    'set a flag so that we  know this at the end of the routine :)
    If Mid$(con, 1, 1) = " " Then
        bhl = True
        con = Right$(con, Len(con) - 1)
    End If
    
    'default to this ..
    style = kw_PopUpList
    Select Case con
 
        Case "asml1" 'asm normal keywords- all the isml list...
            GetListFromContext.ListInfo = kwListToString(dat, 0, 0)  'olny first string (the name of the command)

        Case "none", "comment", "" 'comment -> nothing
            GetListFromContext.ListInfo = ""
            
        Case Else 'we have a numeric list with the possible tooltips/lists
            tls = Split(con, "|") 'get the list
            ReDim tl(UBound(tls)) 'and proccess it
            tl(0) = tls(0)
            For i = 1 To UBound(tls)
                tl(i) = tls(i)
                AddkwToFile t, dat.kw(tl(i))
                If dat.kw(tl(i)).Isml_kw_Type And kw_Tipform Then
                    style = style Or kw_Tipform ' if we can display a tip then set the flag
                End If
                If dat.kw(tl(i)).Isml_kw_Type And kw_PopUpList Then
                    style = style Or kw_PopUpList ' if we can display a list then set the flag
                End If
            Next i
            'Tip text
            If (style And kw_Tipform) > 0 And tl(0) > 1 Then
                GetListFromContext.ttInfo = kwListToString(t, , , kw_Tipform)   'the hole string
            End If
            'List text
            If style And kw_PopUpList Then
                GetListFromContext.ListInfo = kwListToString(t, tl(0) - 1, 0, kw_PopUpList)   'olny the curent level
            End If
            GetListFromContext.curLVL = tl(0)
            
            If bhl = True Then ' if this context is meant not to be for a list..
                'eg , a <any><prv> mach
                GetListFromContext.ListInfo = ""
            End If
    End Select
'error handling
GPF_Reset
Exit Function
errH:
    MsgBox Err.Description & vbNewLine & _
           Err.Source
GPF_Reset
End Function

'return the word on witch the caret is (eg "ster abc|d dfg dfg" will return abcd)
Public Function GetCurCarWord() As String
    Dim slin As Long, scol As Long
    
    VBI.ActiveCodePane.GetSelection slin, scol, 0, 0
    GetCurCarWord = GetWordFormPos(slin, scol)
    
End Function

'return the word on witch the caret is (eg "ster abc|d dfg dfg" will return abcd)
Public Function GetWordFormPos(ByVal slin As Long, ByVal scol As Long, Optional ByRef ws As Long, Optional ByRef we As Long) As String
On Error GoTo errH
GPF_Set GPF_RaiseErr, "modIntelliAsm", "GetCurCarWord"

    Dim tlin As String
    'do some replacements for word seperators
    tlin = Replace(VBI.ActiveCodePane.codeModule.lines(slin, 1), "(", " ")
    tlin = Replace(Replace(Replace(Replace(tlin, "-", " "), "+", " "), "*", " "), ")", " ")
    tlin = Replace(Replace(Replace(Replace(tlin, "/", " "), "\", " "), "^", " "), "&", " ")
    tlin = Replace$(Replace$(tlin, "dword ptr ", "dword_ptr_"), "byte ptr ", "byte_ptr_")
    tlin = Replace$(Replace$(tlin, "word ptr ", "word_ptr_"), "qword ptr ", "qword_ptr_")
    tlin = Replace(tlin, ",", " ")
    
    'if nothing to return then..
    If Len(tlin) = 0 Or Len(Trim$(tlin)) = 0 Or (scol - 1) < 1 Then
        GetWordFormPos = ""
        GoTo Exit_Function
    End If
    
    'Just a few calulations to find the string ...
    'tlin = Replace(Replace(Replace(tlin, "]", " "), ",", " "), "[", " ")
    tlin = Replace(tlin, ",", " ")
    
    ws = InStrRev(tlin, " ", scol - 1, vbTextCompare)
    we = InStr(scol, tlin, ChrW$(32))
    
    ws = ws + 1
    If we = 0 Then we = Len(tlin) + 1
    
    'and to return it..
    If (we - ws) <= 0 Then
        GetWordFormPos = ""
    Else
        GetWordFormPos = Mid$(tlin, ws, we - ws)
    End If
    
Exit_Function:
GPF_Reset
Exit Function
errH:
    MsgBox Err.Description & vbNewLine & _
           Err.Source
GPF_Reset
End Function

'Text changed , so intelliAsm must be updated
Public Sub IntelliAsmChange(ByVal hWnd As Long, ByVal bNotMakenew As Boolean)
Dim tP As POINTAPI, tp2 As WinApiForVb.POINTAPI, temp As LTI
Dim bMakenewTT As Boolean, bMakenewIA As Boolean

    GetCaretPos tP
    
    If bNotMakenew Then bMakenewTT = asmTT.visible Else bMakenewTT = True
    If bNotMakenew Then bMakenewIA = IntelliAsm.visible Else bMakenewIA = True
    'Get the curent Intelli Asm state
    temp = GetListFromContext(GetCurContext(), style)

    'Exit Sub
    'And display/hide tooltip/list as needed
    If Len(temp.ttInfo) = 0 Then asmTT.HideToolTip
    
    If style And Isml_kw_types.kw_PopUpList Then
        If bMakenewIA Then
            IntelliAsm.ShowIntelliAsm hWnd, tP.x - 15, tP.y + 16, _
                                      GetCurCarWord(), temp.ListInfo
        End If
    End If
    If style And kw_Tipform Then
        If bMakenewTT Then
            tp2.x = tP.x: tp2.y = tP.y
            ClientToScreen hWnd, tp2
            ScreenToClient GetParent(asmTT.hWnd), tp2
            asmTT.ShowTooltip tp2.x - 15, tp2.y - 16, Replace$(temp.ttInfo, "|", " "), False, , True
            Dim t() As String, t2 As String, i As Long
            'Make the current match BOLD
            If Len(temp.ttInfo) Then
                
                t = Split(temp.ttInfo, "|")
                asmTT.pbc.Cls
                For i = 0 To temp.curLVL - 2
                    t2 = t2 & t(i) & " "
                Next i
                asmTT.CurrentX = 0
                asmTT.CurrentY = 0
                asmTT.pbc.Print t2;
                asmTT.pbc.FontBold = True
                asmTT.pbc.Print t(temp.curLVL - 1) & " ";
                asmTT.pbc.FontBold = False
                t2 = ""
                For i = temp.curLVL To UBound(t)
                    t2 = t2 & t(i) & " "
                Next i
                asmTT.pbc.Print t2;
            End If
        End If
    End If

End Sub

'Hide Everything..
Public Sub IntelliAsmHideAll()

    IntelliAsm.HideIntelliAsm
    asmTT.HideToolTip

End Sub

'Send command to the list...
Public Sub IntelliAsmListSend(ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long)
    
    If IntelliAsm.visible Then
        SendMessage IntelliAsm.iSe.memb_list.hWnd, msg, wParam, lParam
    End If
    
End Sub

'Expand asm , ccc to '#asm' and '#ccc'
Public Sub AutoExpandAsm_C()
On Error Resume Next
GPF_Set GPF_RaiseErr, "modIntelliAsm", "AutoExpandAsm"
Dim sL As Long, sC As Long, eL As Long, eC As Long

    'get the text
    VBI.ActiveCodePane.GetSelection sL, sC, eL, eC
    If (sL = eL) And (eC = sC) Then
        Dim t As String
        t = VBI.ActiveCodePane.codeModule.lines(sL, 1)
        'expand if needed
        If InStr(1, Trim$(t), "asm ", vbTextCompare) = 1 Then
            t = Replace$(t, "asm ", "'#asm' ", 1, 1, vbTextCompare)
            VBI.ActiveCodePane.codeModule.ReplaceLine sL, t
            VBI.ActiveCodePane.SetSelection sL, sC + Len("'#asm' ") - Len("asm "), eL, sC + Len("'#asm' ") - Len("asm ")
        ElseIf InStr(1, Trim$(t), "ccc ", vbTextCompare) = 1 Then
            t = Replace$(t, "ccc ", "'#c' ", 1, 1, vbTextCompare)
            VBI.ActiveCodePane.codeModule.ReplaceLine sL, t
            VBI.ActiveCodePane.SetSelection sL, sC + Len("'#c' ") - Len("ccc "), eL, sC + Len("'#c' ") - Len("ccc ")
        End If
        
    End If


GPF_Reset
End Sub
