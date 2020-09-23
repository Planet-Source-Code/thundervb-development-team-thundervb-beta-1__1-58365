Attribute VB_Name = "modSubCH"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

'Revision history:
'13/9/2004[dd/mm/yyyy] : Created by Raziel
'This is where all the subclassed messanges proccesing is done...
'Any code that is inovled with the sublcass init / update/redirect
'is on the modSubClasser.bas , this code just proccess the messanges
'
'Mouse Wheel works on MDI enviroment
'
'1/10/2004[dd/mm/yyyy] : Edited by Raziel
'Works on SDI and MDI, mainly due to improvement on the subclasser
'
'
'6/10/2004[dd/mm/yyyy] : Edited by Raziel
'Base code for intelliAsm
'IntelliAsm form show/hide + tip changes
'
'10/10/2004[dd/mm/yyyy] : Edited by Raziel
'Many code fixes , intelliAsm is ready for use...
'
'22/10/2004[dd/mm/yyyy] : Edited by Raziel
'Many many fixes , code now implemetns and teh ctrl+i and ctrl+space
'(they do the same thing in our case)
'ToDo : we need a good asm definition file..
'
Option Explicit


Private Declare Function GetCaretPos Lib "user32" _
               (lpPoint As POINTAPI) As Long

Dim ptemp As POINTAPI
Dim ctrl_down As Boolean
Public asmTT As New frmTip
Public bDoNotHideOfFL As Boolean
Public wm_tthide As Boolean

'Called when the curent's window proc is called , before VB's one
Public Sub WindowProcBef(ByRef hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef PrevProc As Long, ByRef skipVB As Boolean, ByRef skipAft As Boolean)
               
    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveCodePane Is Nothing Then Exit Sub
    
    Select Case uMsg
           
        Case WM_LBUTTONDOWN, WM_RBUTTONDOWN
            IntelliAsm.HideIntelliAsm
            asmTT.HideToolTip
            
        Case WM_KILLFOCUS
            IntelliAsm.HideIntelliAsm
            If asmTT.visible Then
                asmTT.HideToolTip
                wm_tthide = True
            Else
                wm_tthide = False
            End If

            
        Case 522 'Mouse wheel event
            If IntelliAsm.visible = True Then
                SendMessage IntelliAsm.iSe.memb_list.hWnd, uMsg, wParam, lParam
                Exit Sub
            End If
            asmTT.HideToolTip
            Dim Top As Long
            If wParam < 0 Then
                Top = VBI.ActiveCodePane.TopLine + 3
                If Top > VBI.ActiveCodePane.codeModule.CountOfLines Then
                    Top = VBI.ActiveCodePane.codeModule.CountOfLines
                End If
                If TipVisible = False Then VBI.ActiveCodePane.TopLine = Top
                TipOffset = TipOffset + 1
             Else
                Top = VBI.ActiveCodePane.TopLine - 3
                If Top < 1 Then
                    Top = 1
                End If
                If TipVisible = False Then VBI.ActiveCodePane.TopLine = Top
                TipOffset = TipOffset - 1
            End If
            TipReSetText
            
        Case WM_MOUSEMOVE
            CheckToolTip
            TipReSetText
            RichWordOver hWnd, lParam And 65535, lParam \ 65536
        
        Case WM_KEYDOWN ', WM_CHAR
            
            If ((wParam = VK_UP) Or (wParam = VK_DOWN) Or _
                (wParam = vbKeyPageUp) Or (wParam = vbKeyPageDown)) And IntelliAsm.visible = True Then
                
                IntelliAsmListSend uMsg, wParam, lParam
                skipVB = True
                skipAft = True
                
            ElseIf (wParam <> VK_DELETE) And (InStrRev(" ,.];" & vbLf & vbTab, ChrW$(wParam))) And (IntelliAsm.visible = True) And (Len(IntelliAsm.iSe.memb_list.Text) > 0) Then
                
                IntelliAsm.iSe.list_DblClick
                'skipVB = True
                
            ElseIf ((wParam = VK_ESCAPE)) And (IntelliAsm.visible = True) Then
                IntelliAsmHideAll
                skipVB = True
                skipAft = True
                
            End If
        Case WM_KEYUP
            If wParam = VK_CONTROL Then
                ctrl_down = False
            End If
            
    End Select
    
    If GetAsyncKeyState(VK_CONTROL) = True Then
        If ((GetAsyncKeyState(VK_SPACE) <> 0) Or (GetAsyncKeyState(VK_I) <> 0)) = True Then
            IntelliAsmChange hWnd, False
        End If
    End If

End Sub

'Called when the curent's window proc is called , After VB's one
Public Sub WindowProcAft(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, PrevProc As Long, ByRef RetValue As Long)
Dim temp As String
    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveCodePane Is Nothing Then Exit Sub
    
    Select Case uMsg
        
        Case WM_CHAR
        'If wParam = 3 Then MsgBox "copy"
            AutoExpandAsm_C
            If wParam <> 27 Then
                IntelliAsmChange hWnd, False
            End If
            
        Case WM_KEYDOWN
        '    MsgBox wParam
            IntelliAsmChange hWnd, True
            
        'Case WM_HOTKEY
        '    Select Case wParam
        '        Case hk_CopyHotKeyID
        '            'MsgBox "Ctrl+C"
        '            SendMessage hWnd, WM_KEYDOWN, &H11, &H1D0001
        '            SendMessage hWnd, WM_KEYDOWN, &H43, &H2E0001
        '            SendMessage hWnd, WM_CHAR, &H3, &H2E0001
        '            SendMessage hWnd, WM_KEYUP, &H43, &H2E0001
        '            SendMessage hWnd, WM_KEYUP, &H11, &H1D0001
        '            SendMessage hWnd, WM_COPY, 0, 0
        '            Dim rtf As RichTextBox
        '            Set rtf = frmImages.rtb
        '            rtf.TextRTF = ""
        '            rtf.Text = "'Colored by ThunderVB " & GetThunVBVer & vbNewLine & _
        '                                              Clipboard.GetText(vbCFText)
        '            Clipboard.SetText DoNonRealTimeColor(rtf), vbCFRTF
        '        Case hk_CutHotKeyID
        '            'MsgBox "Ctrl+X"
        '        Case hk_CtrlIHotKeyID
        '            'MsgBox "Ctrl+I"
        '        Case hk_CtrlSpaceHotKeyID
        '            'MsgBox "Ctrl+Space"
        '        Case Else
        '            MsgBox "Other hotkey ?? id=" & wParam
        '    End Select
            
    End Select
    
    
        If ((GetAsyncKeyState(VK_C) <> 0) Or (GetAsyncKeyState(VK_X) <> 0)) = True Then
            Dim rtf As RichTextBox
            Set rtf = frmImages.rtb
            rtf.TextRTF = ""
            rtf.Text = "'Colored by ThunderVB " & GetThunVBVer & vbNewLine & _
                                              Clipboard.GetText(vbCFText)
            Clipboard.SetText DoNonRealTimeColor(rtf), vbCFRTF
        End If
    
    
End Sub
