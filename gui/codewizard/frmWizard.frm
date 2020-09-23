VERSION 5.00
Begin VB.Form frmCodeWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASM Code Generator"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview code"
      Height          =   375
      Left            =   4320
      TabIndex        =   39
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtNoOfParam 
      Height          =   285
      Left            =   1920
      TabIndex        =   37
      Top             =   195
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   34
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCursor 
      Caption         =   "Cursor"
      Height          =   375
      Left            =   1560
      TabIndex        =   33
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Clipboard"
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame fraDebug 
      Caption         =   "Debug"
      Height          =   735
      Left            =   2520
      TabIndex        =   30
      Top             =   4800
      Width           =   3015
      Begin VB.CheckBox chbINT3 
         Caption         =   "debug code (""int 3"")"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraLocalVar 
      Caption         =   "Variables/Parameters"
      Height          =   735
      Left            =   240
      TabIndex        =   26
      Top             =   5640
      Width           =   5295
      Begin VB.CheckBox chbAddRet 
         Caption         =   "Add return (""ret"")"
         Height          =   195
         Left            =   3120
         TabIndex        =   40
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chbJumpOverBlock 
         Caption         =   "jump over block (""jmp"")"
         Height          =   195
         Left            =   3120
         TabIndex        =   29
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chbAccessToParam 
         Caption         =   "access to parameters"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox chbLocalVar 
         Caption         =   "local variables (""local"")"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraGenCode 
      Caption         =   "Code"
      Height          =   735
      Left            =   240
      TabIndex        =   22
      Top             =   3960
      Width           =   5295
      Begin VB.CheckBox chbAddComments 
         Caption         =   "Add comments ("";"")"
         Height          =   195
         Left            =   3360
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optClearText 
         Caption         =   "Clear text"
         Height          =   195
         Left            =   1320
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optInLineASM 
         Caption         =   "Inline ASM"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraSaveESP 
      Caption         =   "Saving ESP"
      Height          =   735
      Left            =   240
      TabIndex        =   19
      Top             =   4800
      Width           =   2055
      Begin VB.OptionButton optNoESP 
         Caption         =   "No"
         Height          =   195
         Left            =   1320
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optStandartESP 
         Caption         =   "Standart"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraSaveRegisters 
      Caption         =   "Preserve registers"
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   5295
      Begin VB.CommandButton cmdSaveNo 
         Caption         =   "no"
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdSaveAll 
         Caption         =   "all"
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chbSaveECX 
         Caption         =   "ECX"
         Height          =   195
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chbSaveEDI 
         Caption         =   "EDI"
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chbSaveEDX 
         Caption         =   "EDX"
         Height          =   195
         Left            =   2400
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chbSaveESI 
         Caption         =   "ESI"
         Height          =   195
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chbSaveEBX 
         Caption         =   "EBX"
         Height          =   195
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chbSaveEAX 
         Caption         =   "EAX"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CheckBox chbPasteToProcedure 
      Caption         =   "Paste code to Procedure"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   2175
   End
   Begin VB.Frame fraProcedure 
      Caption         =   "Procedure"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5295
      Begin VB.CheckBox chbDeclareParams 
         Caption         =   "Declare parameters"
         Height          =   195
         Left            =   2880
         TabIndex        =   38
         Top             =   405
         Width           =   1695
      End
      Begin VB.Frame fraScope 
         Caption         =   "Scope"
         Height          =   735
         Left            =   2760
         TabIndex        =   4
         Top             =   840
         Width           =   2295
         Begin VB.OptionButton optPublic 
            Caption         =   "Public"
            Height          =   375
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optPrivate 
            Caption         =   "Private"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fraType 
         Caption         =   "Type"
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2295
         Begin VB.OptionButton optSub 
            Caption         =   "Sub"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optFunc 
            Caption         =   "Function"
            Height          =   375
            Left            =   960
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtProcName 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   405
         Width           =   420
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Number of parameters"
      Height          =   195
      Left            =   240
      TabIndex        =   36
      Top             =   240
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Paste to"
      Height          =   195
      Left            =   240
      TabIndex        =   35
      Top             =   6600
      Width           =   585
   End
End
Attribute VB_Name = "frmCodeWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

'30.8. 2004 - GUI created
'5.9.  2004 - code
'16.9. 2004 - fixed "Local" bug, improved output asm code, added Return option

Private Const MSG_TITLE As String = "ASM Code Generator"

'asm tag
Private Const ASM_TAG As String = "'#asm'"

'debug code
Private Const DEBUG_INT3  As String = "int 3"

'save registers
Private Const SAVE_EAX As String = "push eax"
Private Const SAVE_EBX As String = "push ebx"
Private Const SAVE_ECX As String = "push ecx"
Private Const SAVE_EDX As String = "push edx"
Private Const SAVE_EDI As String = "push edi"
Private Const SAVE_ESI As String = "push esi"

'restore registers
Private Const LOAD_EAX As String = "pop eax"
Private Const LOAD_EBX As String = "pop ebx"
Private Const LOAD_ECX As String = "pop ecx"
Private Const LOAD_EDX As String = "pop edx"
Private Const LOAD_EDI As String = "pop edi"
Private Const LOAD_ESI As String = "pop esi"

'parameters
Private Const PARAMS_START As String = "push ebp" & vbCrLf & "mov ebp, esp"
Private Const PARAMS_END As String = "mov esp, ebp" & vbCrLf & "pop ebp"
Private Const PARAMS_LOCAL As String = "local ???"
Private Const PARAMS_EBP As String = "param* equ [ebp+**]"
Private Const PARAMS_ESP As String = "param* equ [esp+**]"
Private Const PARAMS_JMP_OVER As String = "jmp Over" & vbCrLf & "Over:"

Private Const REPLACE_NAME As String = "*"
Private Const REPLACE_NUM As String = "**"

'add code to function/sub
Private Const PROC_SCOPE_PUB As String = "Public"
Private Const PROC_SCOPE_PRIV As String = "Private"
Private Const PROC_TYPE_SUB As String = "Sub"
Private Const PROC_TYPE_FUNC As String = "Function"
Private Const PROC_TYPE_RET As String = "as ???"
Private Const PROC_PARAM As String = "ByVal|ByRef param** as ???"
Private Const PROC_END_SUB As String = "End Sub"
Private Const PROC_END_FUNC As String = "End Function"

Private sCode As String

'--------------
'--- Events ---
'--------------

'close form
Private Sub cmdClose_Click()
    Unload Me
End Sub

'preview code
Private Sub cmdPreview_Click()
Dim s As String
    s = GenCode
    If Len(s) <> 0 Then
        LogMsg "Preview of code", Me.name, "cmdPreview_Click", True, True
        frmViewer.ShowViewer MSG_TITLE, GenCode, True
    End If
End Sub

'save all registers
Private Sub cmdSaveAll_Click()

    chbSaveEAX.value = 1
    chbSaveEBX.value = 1
    chbSaveECX.value = 1
    chbSaveEDX.value = 1
    chbSaveEDI.value = 1
    chbSaveESI.value = 1
    
End Sub

'save no registers
Private Sub cmdSaveNo_Click()
    
    chbSaveEAX.value = 0
    chbSaveEBX.value = 0
    chbSaveECX.value = 0
    chbSaveEDX.value = 0
    chbSaveEDI.value = 0
    chbSaveESI.value = 0
    
End Sub

'initialize
Private Sub Form_Initialize()
    LogMsg "Loading " & Add34(Me.caption) & " window", Me.name, "Form_Initialize", True, True
End Sub


Private Sub Form_Load()

    optPrivate.value = True
    optSub.value = True
    optStandartESP.value = True
    optInLineASM.value = True
    
    Call chbPasteToProcedure_Click
    
    txtNoOfParam.Text = 0
    
End Sub

Private Sub Form_Terminate()
    LogMsg "Unloading " & Add34(Me.caption) & " window", Me.name, "Form_Terminate", True, True
End Sub

Private Sub chbPasteToProcedure_Click()
    fraProcedure.Enabled = chbPasteToProcedure.value
End Sub

'check number - number of parameters
Private Sub txtNoOfParam_Validate(Cancel As Boolean)
    
    Cancel = True
    
    If IsNumeric(txtNoOfParam.Text) = False Then
        MsgBox "Write number of parameters.", vbInformation, MSG_TITLE
    ElseIf CLng(txtNoOfParam.Text) < 0 Then
        MsgBox "Number must be greater than zero.", vbInformation, MSG_TITLE
    ElseIf InStr(1, txtNoOfParam.Text, ",") > 0 Or InStr(1, txtNoOfParam.Text, ".") > 0 Then
        MsgBox "Write integer number.", vbInformation, MSG_TITLE
    Else
        Cancel = False
    End If
    
End Sub

'save ASM code to clipbord
Private Sub cmdClipboard_Click()
Dim sCode As String
    
    LogMsg "Pasting ASM code to clipboard", Me.name, "cmdClipboard_Click", True, True
    
    'get ASM code
    sCode = GenCode
    If Len(sCode) = 0 Then Exit Sub
    
    'save it to clipboard
    Clipboard.Clear
    Clipboard.SetText sCode
    
End Sub

'paste ASM code to cursor
Private Sub cmdCursor_Click()
Dim sCode As String

    sCode = GenCode
    If Len(sCode) = 0 Then Exit Sub
    
    LogMsg "Pasting ASM code to cursor location", Me.name, "cmdCursor_Click", True, True
    PutStringToCurCursor sCode

End Sub

'------------------------
'--- Helper functions ---
'------------------------

'generate code
'return - string - generated ASM code
Private Function GenCode() As String
Dim s As String, i As Long

    'check LOCAL "bug"
    If optStandartESP.value = True And chbLocalVar.value = 1 Then
        MsgBox "When using Local variables, MASM will generate necessary code to save EBP and ESP." & vbCrLf & _
               "Set Saving ESP option to No or do not use Local variables.", vbInformation, MSG_TITLE
        Exit Function
    End If
    
    'check procedure name
    If chbPasteToProcedure.value = 1 Then
        
        If Len(txtProcName.Text) = 0 Then
            MsgBox "Procedure name is zero-length.", vbInformation, MSG_TITLE
            Exit Function
        End If
        
    End If

    sCode = ""
    
    'local variables
    If chbLocalVar.value = 1 Then AddCode PARAMS_LOCAL
    'add "int 3"
    If chbINT3.value = 1 Then AddCode DEBUG_INT3, True
    'jump over block
    If chbJumpOverBlock.value = 1 Then AddCode PARAMS_JMP_OVER, True
    
    'access to parameters via constants
    If chbAccessToParam.value = 1 And CLng(txtNoOfParam.Text) > 0 Then
        
        AddCode ""
        
        'enum all parameters
        For i = 1 To CLng(txtNoOfParam.Text)
            If optStandartESP.value = True Then
                'create variable names
                s = Replace(PARAMS_EBP, REPLACE_NAME, i, , 1)
                s = Replace(s, REPLACE_NUM, i * 4 + 4, , 1)
                AddCode s
            Else
                'create variable names
                s = Replace(PARAMS_ESP, REPLACE_NAME, i, , 1)
                s = Replace(s, REPLACE_NUM, i * 4, , 1)
                AddCode s
            End If
        Next i
    
    End If

    'common procedure entry-point
    If optStandartESP.value = True Then AddCode PARAMS_START, True
    
    If chbSaveEAX.value = 1 Or chbSaveEBX.value = 1 Or chbSaveECX.value = 1 Or chbSaveEDX.value = 1 Or chbSaveEDI.value = 1 Or chbSaveESI.value = 1 Then AddCode ""
    
    'save registers
    If chbSaveEAX.value = 1 Then AddCode SAVE_EAX
    If chbSaveEBX.value = 1 Then AddCode SAVE_EBX
    If chbSaveECX.value = 1 Then AddCode SAVE_ECX
    If chbSaveEDX.value = 1 Then AddCode SAVE_EDX
    If chbSaveEDI.value = 1 Then AddCode SAVE_EDI
    If chbSaveESI.value = 1 Then AddCode SAVE_ESI
    
    If chbSaveEAX.value = 1 Or chbSaveEBX.value = 1 Or chbSaveECX.value = 1 Or chbSaveEDX.value = 1 Or chbSaveEDI.value = 1 Or chbSaveESI.value = 1 Then AddCode ""
    
    'restore registers
    If chbSaveESI.value = 1 Then AddCode LOAD_ESI
    If chbSaveEDI.value = 1 Then AddCode LOAD_EDI
    If chbSaveEDX.value = 1 Then AddCode LOAD_EDX
    If chbSaveECX.value = 1 Then AddCode LOAD_ECX
    If chbSaveEBX.value = 1 Then AddCode LOAD_EBX
    If chbSaveEAX.value = 1 Then AddCode LOAD_EAX
    
    'common procedure entry-point
    If optStandartESP.value = True Then AddCode PARAMS_END, True
    
    'RET instrunction
    If chbAddRet.value = 1 Then
        If CLng(txtNoOfParam.Text) > 0 Then
            AddCode "ret " & 4 * CLng(txtNoOfParam.Text), True
        Else
            AddCode "ret", True
        End If
    End If

    'add asm tag
    If optInLineASM.value = True Then
        'adjust first line
        If Len(sCode) > 0 Then sCode = ASM_TAG & " " & sCode
        sCode = Replace(sCode, vbCrLf, vbCrLf & ASM_TAG & " ")
    End If
    
    'add comment tag - ;
    If chbAddComments.value = 1 Then
Dim asLine() As String, lMax As Long, lCur As Long
        
        'initialize
        lMax = 0
        asLine = Split(sCode, vbCrLf)
            
        'look for the longest line
        For i = LBound(asLine) To UBound(asLine)
            If Len(asLine(i)) > lMax Then lMax = Len(asLine(i))
        Next i
        
        'align line and add ";"
        For i = LBound(asLine) To UBound(asLine)
            lCur = Len(asLine(i))
            asLine(i) = asLine(i) & Space(lMax - lCur + 5) & ";"
        Next i
        
        sCode = Join$(asLine, vbCrLf)
    
    End If
    
    'paste code to procedure
    If chbPasteToProcedure.value = 1 Then
Dim sHeader As String
        
        'private/public
        sHeader = IIf(optPrivate.value = True, PROC_SCOPE_PRIV, PROC_SCOPE_PUB)
        'Sub/Fnction
        sHeader = sHeader & " " & IIf(optSub.value = True, PROC_TYPE_SUB, PROC_TYPE_FUNC)
        'procedure name
        sHeader = sHeader & " " & txtProcName.Text
        sHeader = sHeader & "("
        
        'add parameters
        If CLng(txtNoOfParam.Text) > 0 And chbDeclareParams.value = 1 Then
            For i = 1 To CLng(txtNoOfParam.Text)
                sHeader = sHeader & Replace(PROC_PARAM, REPLACE_NUM, i)
                'add ","
                If i <> CLng(txtNoOfParam.Text) Then sHeader = sHeader & ", "
            Next i
        End If
        
        sHeader = sHeader & ")"
        'if it is function then add "as ???"
        If optSub.value = False Then sHeader = sHeader & " " & PROC_TYPE_RET
    
        'End Sub/End function
        sCode = sHeader & vbCrLf & sCode & vbCrLf & IIf(optSub.value = True, PROC_END_SUB, PROC_END_FUNC)
        
    End If
    
    GenCode = sCode

End Function

'add text to variable sCode
'sASM - text to add
'bBlankLine - between lines will be blank line

Private Sub AddCode(sASM As String, Optional bBlankLine As Boolean = False)

    If Len(sCode) = 0 Then
        sCode = sASM
    Else
        If bBlankLine = True Then
            If Right(sCode, 2) <> vbCrLf Then
                sCode = sCode & vbCrLf & vbCrLf & sASM
            Else
                sCode = sCode & vbCrLf & sASM
            End If
        Else
            sCode = sCode & vbCrLf & sASM
        End If
    End If

End Sub
