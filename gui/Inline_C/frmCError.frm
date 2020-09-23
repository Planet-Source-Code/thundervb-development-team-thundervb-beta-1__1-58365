VERSION 5.00
Begin VB.Form frmCError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C Error"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtclOut 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmCError.frx":0000
      Top             =   120
      Width           =   7575
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "Retry"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdIng 
      Caption         =   "Ingore"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmd_can 
      Caption         =   "Cancel Compile"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   6240
      Width           =   1575
   End
   Begin ThunderVB.ScintillaEdit txtErr 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5530
      0               =   0   'False
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
      54              =   $"frmCError.frx":0014
      55              =   512
   End
End
Attribute VB_Name = "frmCError"
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
'29/8/2004[dd/mm/yyyy] : Created by Raziel
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
    Dim i As Long
    
    hf = 0
    txtErr.MarginWidth(0) = 35
    txtErr.MarginWidth(1) = 0
    txtErr.MarginType(0) = SC_MARGIN_NUMBER
    txtErr.Lexer = SCLex.CppNoCase
    txtErr.StyleClearAll
       
    '"Primary keywords and identifiers",
    '"Secondary keywords and identifiers",
    '"Documentation comment keywords",
    '"Unused",
    '"Global classes and typedefs",
    
    txtErr.KEYWORDS(0) = LCase("and and_eq asm auto bitand bitor bool break case catch char " & _
                               "class compl const const_cast continue default delete do double " & _
                               "dynamic_cast else enum explicit export extern false float for " & _
                               "friend goto if inline int long mutable namespace new not not_eq " & _
                               "operator or or_eq private protected public register reinterpret_cast " & _
                               "return short signed sizeof static static_cast struct switch " & _
                               "template this throw true try typedef typeid typename union " & _
                               "unsigned using virtual void volatile wchar_t while xor xor_eq")

    
    For i = 0 To 31
        txtErr.StyleSetFore CInt(i), GetCWordColor("*default*")
    Next i
    
    'standar
    txtErr.StyleSetFore SCE_C_NUMBER, GetCWordColor("1234")
    txtErr.StyleSetFore SCE_C_STRING, GetCWordColor(Add34("'this is a string'"))
    txtErr.StyleSetFore SCE_C_COMMENT, GetCWordColor("//")
    'remaped
    '
    txtErr.StyleSetFore SCE_C_WORD, GetCWordColor("if")


End Sub

Public Sub ShowError(clErr As String, errText As String, Optional title As String = "Error on C code:")
    
    hf = 0
    Me.caption = title
    Me.txtclOut.Text = clErr
    Me.txtErr.Text = errText
    Me.show ' vbModal
    Do
        Sleep 10
        DoEvents
    Loop While hf = 0
    Me.hide
    If hf = 2 Then
        hf = 0
        retryC Me.txtErr.Text
    End If
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    hf = 1
    Cancel = 1
    Me.hide
    
End Sub



