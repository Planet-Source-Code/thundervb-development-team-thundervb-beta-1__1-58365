Attribute VB_Name = "modAsmToolTip"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

Public cpTip As New frmTip
Public TipVisible As Boolean
Public TipOffset As Long
Public LastWord As String

Public Sub CheckToolTip()
Dim sLine As Long, scol As Long, ecol As Long, temp As String, tP As Long

On Error GoTo exterr:

    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveCodePane Is Nothing Then Exit Sub
    
    With VBI.ActiveCodePane
        .GetSelection sLine, scol, 0, ecol
        If scol > ecol Then
            Dim tmpa As Long
            tmpa = scol: scol = ecol: ecol = tmpa
        End If
        temp = Mid$(.codeModule.lines(sLine, 1), scol, ecol - scol)
        tP = DeclaredAt(temp, sLine)
        
        If Len(temp) And tP > 0 Then
            cpTip.EnToolTipA
            If TipVisible = False Then
                'VBI.ActiveCodePane.Window.SetFocus
                TipVisible = True
                TipOffset = tP
            End If
            TipReSetText
            
        Else
            If Len(LastWord) > 0 Then
                    tP = DeclaredAt(LastWord, sLine)
                    If tP > 0 Then
                        cpTip.EnToolTipA
                        If TipVisible = False Then

                            TipVisible = True
                            TipOffset = tP
                        End If
                        TipReSetText
                        Exit Sub
                    End If
            End If
            TipVisible = False
            cpTip.HideToolTip
            
        End If
        
    End With
Exit Sub
exterr:
    If Len(LastWord) > 0 Then
            tP = DeclaredAt(LastWord, sLine)
            If tP > 0 Then
                cpTip.EnToolTipA
                If TipVisible = False Then

                    TipVisible = True
                    TipOffset = tP
                End If
                TipReSetText
                Exit Sub
            End If
    Else
        TipVisible = False
        cpTip.HideToolTip
    End If
    
End Sub

Sub TipReSetText()
Dim soffset As Long, slen As Long, MaxLines As Long

    If TipVisible Then
    
        If VBI Is Nothing Then Exit Sub
        If VBI.ActiveCodePane Is Nothing Then Exit Sub
        
        MaxLines = VBI.ActiveCodePane.codeModule.CountOfLines
        
        soffset = TipOffset
        If soffset < 1 Then soffset = 1
        If soffset > MaxLines Then soffset = MaxLines
        slen = 9
        If (soffset + slen) >= MaxLines Then slen = MaxLines - soffset - 1
        If slen = 0 Then slen = 1: soffset = soffset - 1
        cpTip.ShowTooltip 0, 0, VBI.ActiveCodePane.codeModule.lines(soffset, slen), True
        TipOffset = soffset
    Else
        TipOffset = 0
    End If
    
End Sub


Function DeclaredAt(ByVal sText As String, sLine As Long) As Long
Dim temp As Long, s As String, cLine As String, pcount As Long, pk As vbext_ProcKind, pline As Long
Dim sLines() As String
    If Len(sText) = 0 Then Exit Function
    If VBI Is Nothing Then Exit Function
    If VBI.ActiveCodePane Is Nothing Then Exit Function
    cLine = Trim(VBI.ActiveCodePane.codeModule.lines(sLine, 1))
    If Len(cLine) < Len("'#asm'") Then Exit Function
    If LCase(Left(cLine, Len("'#asm'"))) <> LCase("'#asm'") Then Exit Function
    
    With VBI.ActiveCodePane.codeModule
        s = .ProcOfLine(sLine, pk)
        pline = .ProcBodyLine(s, pk)
        pcount = .ProcCountLines(s, pk)
        'pcount = pcount - pline
        sLines = Split(.lines(pline, pcount), vbNewLine)
    End With
    'asm constans/label find here
    'They are declared as: name equ value
    '                      name db  value
    '                      name dw  value
    '                      name dd  value
    '                      name:
    Dim i As Long
    For i = 0 To UBound(sLines)
        
        sLines(i) = Trim(sLines(i))
        If Len(sLines(i)) < Len("'#asm'") Then GoTo NextOne
        If LCase(Left(sLines(i), Len("'#asm'"))) <> LCase("'#asm'") Then GoTo NextOne
        sLines(i) = Right(sLines(i), Len(sLines(i)) - Len("'#asm'"))
        sLines(i) = Trim(sLines(i))
        If Len(sLines(i)) = 0 Then GoTo NextOne
        
        If InStr(sLines(i), ";") Then
            sLines(i) = Left(sLines(i), InStr(sLines(i), ";") - 1)
        End If
        Select Case GetFirstWord(sLines(i))
            Case Is = sText
                RemFisrtWord sLines(i): sLines(i) = Trim(sLines(i))
                Select Case GetFirstWord(sLines(i))
                
                    Case "dd", "db", "dw", "equ"
                        DeclaredAt = pline + i
                        Exit Function
                        
                End Select
            
            Case Is = (sText & ":")
                DeclaredAt = pline + i
                Exit Function
                            
            Case "extern", "externdef"
                RemFisrtWord sLines(i): sLines(i) = Trim(sLines(i))
                
                If GetFirstWord(sLines(i)) = sText & ":near" Or GetFirstWord(sLines(i)) = sText Or GetFirstWord(sLines(i)) = sText & ":" Then
                    
                    DeclaredAt = pline + i
                    Exit Function
                
                End If
                
        End Select
        
NextOne:
    Next i
    
    temp = InStr(1, s, sText & ":", vbTextCompare)
    'pcount
    If temp = 0 Then Exit Function
    
    DeclaredAt = UBound(Split(Left(s, temp), vbNewLine))
    

    
End Function

' Return the word the mouse is over.
Public Function RichWordOver(hWnd As Long, x As Long, y As Long) As String
Dim curline As Long, curChar As Long, t As String, pw As Long, nw As Long
On Error Resume Next

    ' Convert the position to pixels.
    Dim tm As A_TEXTMETRIC, hdc As Long
    hdc = GetDC(hWnd)
    A_GetTextMetrics hdc, tm
    ReleaseDC hWnd, hdc
    curline = VBI.ActiveCodePane.TopLine + ((y - 34) \ tm.tmHeight)
    curChar = ((x - 34) \ tm.tmAveCharWidth) + 1
    If curline > VBI.ActiveCodePane.codeModule.CountOfLines Then
        curline = VBI.ActiveCodePane.codeModule.CountOfLines
    End If
    
    t = VBI.ActiveCodePane.codeModule.lines(curline, 1)
    
    If (curChar > Len(t)) Or (curChar < 0) Then
        LastWord = ""
    Else
        LastWord = GetWordFormPos(curline, curChar)
    End If

End Function

