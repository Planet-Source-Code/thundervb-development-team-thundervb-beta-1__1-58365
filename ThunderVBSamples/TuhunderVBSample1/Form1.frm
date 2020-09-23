VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThunderVB Sample 1"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Call VB code from C code"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Call a cdecl function"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Call a function by pointer"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "[asm/c on/off]"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    MsgBox "Result is : " & CallBP(AddressOf CalledByPtr)
    
End Sub

Private Sub Command2_Click()

    MsgBox "Result is : " & Callcdecl
    
End Sub

Private Sub Command3_Click()
    
    CallVBF
    
End Sub

Private Sub Form_Load()
    
    Label2.Caption = IIf(IsAsmEnabled, "Asm is On", "Asm is Off") & vbNewLine
    Label2.Caption = Label2.Caption & IIf(IsCEnabled, "C is On", "C is Off")
    
End Sub

