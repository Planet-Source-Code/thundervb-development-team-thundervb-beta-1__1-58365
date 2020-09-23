Attribute VB_Name = "modCTCol"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

'Based on a method found on psc..
'there are much better ways to do this
'but .. i do not need neither flicker free (cause it is not displayed)
'neither fast (cause it is used rarely ...) but i did needed accurate ,
'bug free and easy to extend...

'Non real time coloring for copy's..
'Casue i liked VB .net copies with coloring info ...

'ToDo : hmm what to do??

Const COL_KEYWORD As String = &H800000    ' dark blue
Const COL_COMMENT As String = &H8000&    ' middle green
Const CHAR_COMMENT As String = "'"        ' comment line char

Type WORD_TYPE
    Text As String  ' word to be colored
    Color As Long   ' color to color the word :)
End Type



Dim Words() As WORD_TYPE
Dim wc As Long

Dim CWords() As WORD_TYPE
Dim CWc As Long

Dim AsmWords() As WORD_TYPE
Dim AsmWc As Long

Dim inited As Boolean




Public Sub InitKeyWords()

    If inited Then Exit Sub
    inited = True
    ' initialize the array of words
    'VisualBasic
    AddWord "Option", COL_KEYWORD
    AddWord "Set", COL_KEYWORD
    AddWord "Explicit", COL_KEYWORD
    AddWord "Type", COL_KEYWORD
    AddWord "as", COL_KEYWORD
    AddWord "End", COL_KEYWORD
    AddWord "Dim", COL_KEYWORD
    AddWord "ReDim", COL_KEYWORD
    AddWord "Public", COL_KEYWORD
    AddWord "Private", COL_KEYWORD
    AddWord "Sub", COL_KEYWORD
    AddWord "ByVal", COL_KEYWORD
    AddWord "Byref", COL_KEYWORD
    AddWord "If", COL_KEYWORD
    AddWord "Then", COL_KEYWORD
    AddWord "Else", COL_KEYWORD
    AddWord "For", COL_KEYWORD
    AddWord "Next", COL_KEYWORD
    
    AddWord "To", COL_KEYWORD
    AddWord "Exit", COL_KEYWORD
    AddWord "Do", COL_KEYWORD
    AddWord "Loop", COL_KEYWORD
    AddWord "While", COL_KEYWORD
    AddWord "Until", COL_KEYWORD
    AddWord "DoEvents", COL_KEYWORD
    AddWord "Long", COL_KEYWORD
    AddWord "Byte", COL_KEYWORD
    AddWord "Single", COL_KEYWORD
    AddWord "Double", COL_KEYWORD
    AddWord "Integer", COL_KEYWORD
    AddWord "Function", COL_KEYWORD
    AddWord "And", COL_KEYWORD
    AddWord "Event", COL_KEYWORD
    AddWord "LBound", COL_KEYWORD
    AddWord "Xor", COL_KEYWORD

    AddWord "Const", COL_KEYWORD
    AddWord "Boolean", COL_KEYWORD
    AddWord "Lib", COL_KEYWORD
    AddWord "Alias", COL_KEYWORD
    AddWord "UBound", COL_KEYWORD
    AddWord "VarPtr", COL_KEYWORD
    AddWord "Declare", COL_KEYWORD
    AddWord "On Error", COL_KEYWORD
    AddWord "Resume", COL_KEYWORD
    AddWord "Select", COL_KEYWORD
    AddWord "Case", COL_KEYWORD

    'For C
    AddCWord "and"
    AddCWord "and_eq"
    AddCWord "bitand"
    AddCWord "bitor"
    AddCWord "bool"
    AddCWord "break"
    AddCWord "break"
    AddCWord "case"
    AddCWord "catch"
    AddCWord "char"
    AddCWord "class"
    AddCWord "compl"
    AddCWord "const"
    AddCWord "const_cast"
    AddCWord "continue"
    AddCWord "default"
    AddCWord "delete"
    AddCWord "do"
    AddCWord "double"
    AddCWord "dynamic_cast"
    
    AddCWord "else"
    AddCWord "enum"
    AddCWord "explicit"
    AddCWord "export"
    AddCWord "extern"
    AddCWord "false"
    AddCWord "float"
    AddCWord "for"
    AddCWord "friend"
    AddCWord "goto"
    AddCWord "if"
    AddCWord "inline"
    AddCWord "dynamic_cast"
    AddCWord "int"
    AddCWord "long"
    AddCWord "mutable"
    AddCWord "namespace"
    AddCWord "new"
    AddCWord "not"
    AddCWord "not_eq"
    AddCWord "operator"
    AddCWord "or"
    AddCWord "or_eq"
    AddCWord "private"
    AddCWord "protected"
    AddCWord "public"
    AddCWord "register"
    AddCWord "reinterpret_cast"
    
    AddCWord "return"
    AddCWord "short"
    AddCWord "signed"
    AddCWord "sizeof"
    AddCWord "static"
    AddCWord "static_cast"
    AddCWord "struct"
    AddCWord "switch"
    AddCWord "template"
    AddCWord "this"
    AddCWord "throw"
    AddCWord "true"
    AddCWord "try"
    AddCWord "typedef"
    AddCWord "typeid"
    AddCWord "typename"
    AddCWord "union"
    AddCWord "unsigned"
    AddCWord "using"
    AddCWord "virtual"
    AddCWord "void"
    AddCWord "volatile"
    AddCWord "wchar_t"
    AddCWord "while"
    AddCWord "xor"
    AddCWord "xor_eq"
    
    
    'For Asm
    AddAsmWord "EAX"
    AddAsmWord "EBX"
    AddAsmWord "ECX"
    AddAsmWord "EDX"
    AddAsmWord "AH"
    AddAsmWord "AL"
    AddAsmWord "BH"
    AddAsmWord "BL"
    AddAsmWord "CH"
    AddAsmWord "CL"
    AddAsmWord "DH"
    AddAsmWord "DL"
    AddAsmWord "CS"
    AddAsmWord "DS"
    AddAsmWord "ES"
    AddAsmWord "FS"
    AddAsmWord "GS"
    AddAsmWord "SS"
    AddAsmWord "AX"
    AddAsmWord "BX"
    AddAsmWord "CX"
    AddAsmWord "DX"
    AddAsmWord "ESI"
    AddAsmWord "EDI"
    AddAsmWord "EBP"
    AddAsmWord "EIP"
    AddAsmWord "ESP"
    AddAsmWord "EFLAGS"
    
End Sub
            
            
Public Function DoNonRealTimeColor(rtb As RichTextBox) As String
Dim i As Long
Dim p1 As Long, p2 As Long
Dim Text As String
Dim sTmp As String

    InitKeyWords
    rtb.Text = " " & rtb.Text & " "
    ' cache the text - speeds things up a bit
    Text = LCase(ReplWSwithSpace(Replace(ProcStringsUnderAll(rtb.Text), vbNewLine, "  ")))
    
    ' go through each item in the Words array
    For i = LBound(Words) To UBound(Words)
    
        ' find each instance of the word in the rtb
        p1 = InStr(1, Text, Words(i).Text)
        Do While p1 > 0
        
            ' color it to the appropriate color
            rtb.SelStart = p1
            rtb.SelLength = Len(Words(i).Text) - 2
            rtb.SelColor = Words(i).Color
            
            ' go on to the next word
            p1 = InStr(p1 + 1, Text, Words(i).Text)

        Loop

    Next i
    
    ' go through and color all the comments '#asm' as well as '#c'..
    p1 = 1
    Text = rtb.Text
    Do While p1 <> 2 And p1 < Len(Text)
        
        ' find the next eol character
        p2 = InStr(p1 + 1, Text, vbCrLf)
        If p2 = 0 Then p2 = Len(Text)
        
        ' grab this line into a temp variable
        sTmp = LCase(Mid$(Text, p1, p2 - p1))
        
        ' if it's a comment line - color it
        If InStr(1, Trim$(sTmp), "'#asm'") = 1 Then '#asm' line..
            rtb.SelStart = (p1 + InStr(1, sTmp, "'#asm'") - 2)
            rtb.SelLength = Len("'#asm'")
            rtb.SelColor = COL_COMMENT
            rtb.SelStart = rtb.SelStart + Len("'#asm'")
            rtb.SelLength = p2 - rtb.SelStart
            rtb.SelColor = GetAsmWordColor("*default*")
            Dim l As Long, ls As Long
            sTmp = Replace(ReplWSwithSpace(sTmp) & " ", "'#asm'", "      ")
            For l = 0 To AsmWc - 1
                ls = InStr(1, sTmp, AsmWords(l).Text)
                Do While ls
                    rtb.SelStart = p1 + ls - 1
                    rtb.SelLength = Len(AsmWords(l).Text) - 2
                    rtb.SelColor = AsmWords(l).Color
                    ls = InStr(ls + 1, sTmp, AsmWords(l).Text)
                Loop
            Next l
            ls = InStr(1, sTmp, ";")
            If ls Then
                rtb.SelStart = (p1 + ls - 2)
                rtb.SelLength = p2 - (p1 + ls - 2)
                rtb.SelColor = GetAsmWordColor(";")
                rtb.SelItalic = True
            End If
        ElseIf InStr(1, Trim$(sTmp), "'#c'") = 1 Then 'no #asm' line ?? '#c' perhaps???
            rtb.SelStart = (p1 + InStr(1, sTmp, "'#c'") - 2)
            rtb.SelLength = Len("'#c'")
            rtb.SelColor = COL_COMMENT
            rtb.SelStart = rtb.SelStart + Len("'#c'")
            rtb.SelLength = p2 - rtb.SelStart
            rtb.SelColor = GetCWordColor("*default*")
            sTmp = Replace(ReplWSwithSpace(sTmp) & " ", "'#c'", "    ")
            For l = 0 To CWc - 1
                ls = InStr(1, sTmp, CWords(l).Text)
                Do While ls
                    rtb.SelStart = p1 + ls - 1
                    rtb.SelLength = Len(CWords(l).Text) - 2
                    rtb.SelColor = CWords(l).Color
                    ls = InStr(ls + 1, sTmp, CWords(l).Text)
                Loop
            Next l
            ls = InStr(1, sTmp, "//")
            If ls Then
                rtb.SelStart = (p1 + ls - 2)
                rtb.SelLength = p2 - (p1 + ls - 2)
                rtb.SelColor = GetCWordColor("//")
                rtb.SelItalic = True
            End If
        ElseIf InStr(1, sTmp, "'") Then 'ohh well disapointed ..this was just a comment..
            rtb.SelStart = (p1 + InStr(1, sTmp, "'") - 2)
            rtb.SelLength = p2 - (p1 + InStr(1, sTmp, "'") - 2)
            rtb.SelColor = COL_COMMENT
            rtb.SelItalic = True
        End If
        
        ' move onto the next line
        p1 = p2 + 2
        
    Loop
    
    DoNonRealTimeColor = rtb.TextRTF
    
End Function

Sub AddWord(strname As String, lCol As Long)

    ReDim Preserve Words(wc)
    Words(wc).Color = lCol
    Words(wc).Text = " " & LCase(strname) & " "
    wc = wc + 1
    
End Sub

Sub AddCWord(strname As String)

    ReDim Preserve CWords(CWc)
    CWords(CWc).Color = GetAsmWordColor(LCase(strname))
    CWords(CWc).Text = " " & LCase(strname) & " "
    CWc = CWc + 1
    
End Sub

Sub AddAsmWord(strname As String)

    ReDim Preserve AsmWords(AsmWc)
    AsmWords(AsmWc).Color = GetAsmWordColor(LCase(strname))
    AsmWords(AsmWc).Text = " " & LCase(strname) & " "
    AsmWc = AsmWc + 1
    
End Sub
