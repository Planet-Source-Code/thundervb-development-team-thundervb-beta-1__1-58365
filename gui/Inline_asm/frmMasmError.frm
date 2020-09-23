VERSION 5.00
Begin VB.Form frmMasmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asm Error"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin ThunderVB.ScintillaEdit txtErr 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   7695
      _extentx        =   13573
      _extenty        =   5530
      1               =   10000000
      2               =   3959
      3               =   12632256
      4               =   0
      5               =   0
      6               =   0
      7               =   0
      8               =   1
      9               =   0
      10              =   0
      11              =   0
      12              =   1
      13              =   0
      14              =   0
      15              =   0
      16              =   0   'False
      17              =   0   'False
      18              =   -1  'True
      19              =   0
      20              =   -1  'True
      21              =   8
      22              =   0
      23              =   -1  'True
      24              =   -1  'True
      25              =   0   'False
      26              =   1
      27              =   1
      28              =   0
      29              =   1
      30              =   500
      31              =   65535
      32              =   0   'False
      33              =   0
      34              =   0   'False
      35              =   -1  'True
      36              =   0
      37              =   -1  'True
      38              =   2000
      39              =   0
      40              =   -1  'True
      41              =   -1  'True
      42              =   0   'False
      43              =   0
      44              =   0
      45              =   0
      46              =   0
      47              =   -1  'True
      48              =   0
      49              =   0   'False
      50              =   0
      51              =   0
      52              =   5
      53              =   0   'False
      54              =   $"frmMasmError.frx":0000
      55              =   512
   End
   Begin VB.CommandButton cmd_can 
      Caption         =   "Cancel Compile"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdIng 
      Caption         =   "Ingore"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "Retry"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox txtMasmOut 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmMasmError.frx":017F
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmMasmError"
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


'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Raziel
'Form Created , intial version


Dim hf As Long

Private Sub cmd_can_Click()

    hf = 1
    Cancel_compile = True
    Me.hide
    
End Sub

Private Sub cmdIng_Click()

    hf = 1
    Me.hide
    
End Sub

Private Sub cmdRetry_Click()

    hf = 2
    Me.hide
    
End Sub

Private Sub Form_Initialize()
    'init com controls peharps?? (for xp styles)
    'init sciedit
    Dim i As Long
    hf = 0
    txtErr.MarginWidth(0) = 35
    txtErr.MarginWidth(1) = 0
    txtErr.MarginType(0) = SC_MARGIN_NUMBER
    txtErr.Lexer = SCLex.asm
    txtErr.StyleClearAll
       
    'WordList &cpuInstruction = *keywordlists[0];
    'WordList &mathInstruction = *keywordlists[1];
    'WordList &registers = *keywordlists[2];
    'WordList &directive = *keywordlists[3];
    'WordList &directiveOperand = *keywordlists[4];
    'WordList &extInstruction = *keywordlists[5];
    
    txtErr.KEYWORDS(0) = LCase("EAX EBX ECX EDX")
    txtErr.KEYWORDS(1) = LCase("AX BX CX DX")
    txtErr.KEYWORDS(2) = LCase("AH AL BH BL CH CL DH DL")
    txtErr.KEYWORDS(3) = LCase("CS DS ES FS GS SS")
    txtErr.KEYWORDS(4) = LCase("ESI EDI EBP EIP ESP")
    txtErr.KEYWORDS(5) = LCase("EFLAGS")
    
    For i = 0 To 31
        txtErr.StyleSetFore CInt(i), GetAsmWordColor("*default*")
    Next i
    
    'standar
    txtErr.StyleSetFore SCE_ASM_NUMBER, GetAsmWordColor("1234")
    txtErr.StyleSetFore SCE_ASM_STRING, GetAsmWordColor("'this is a string'")
    txtErr.StyleSetFore SCE_ASM_COMMENT, GetAsmWordColor(";")
    'remaped
    '
    txtErr.StyleSetFore SCE_ASM_CPUINSTRUCTION, GetAsmWordColor("eax")
    txtErr.StyleSetFore SCE_ASM_MATHINSTRUCTION, GetAsmWordColor("ax")
    txtErr.StyleSetFore SCE_ASM_REGISTER, GetAsmWordColor("ah")
    txtErr.StyleSetFore SCE_ASM_DIRECTIVE, GetAsmWordColor("cs")
    txtErr.StyleSetFore SCE_ASM_DIRECTIVEOPERAND, GetAsmWordColor("esi")
    txtErr.StyleSetFore SCE_ASM_EXTINSTRUCTION, GetAsmWordColor("EFLAGS")

End Sub

Public Sub ShowError(masmErr As String, errText As String, Optional title As String = "Error on Asm code:")
    
    hf = 0
    Me.caption = title
    Me.txtMasmOut = masmErr
    Me.txtErr.Text = errText
    Me.show ' vbModal
    Do
        Sleep 10
        DoEvents
    Loop While hf = 0
    Me.hide
    If hf = 2 Then
        hf = 0
        retryAsm Me.txtErr.Text
    End If
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    hf = 1
    Cancel = 1
    Me.hide
    
End Sub

