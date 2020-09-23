VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ThunderVB Sample 3"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "C_copy"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Api_CopyMem"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ASM-MMX"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Dim testmem1(104857600 / 128) As Long
Dim testmem2(104857600 / 128) As Long
Dim tmrS As Double
Dim i As Long
Const redo As Long = 128
Dim tmrE As Double

Private Sub Command1_Click()
tmrS = Timer
For i = 0 To redo
MemCopyMMX_PreFetch VarPtr(testmem1(0)), VarPtr(testmem2(0)), (UBound(testmem2) + 1) * 4
Next i
tmrE = Timer
On Error Resume Next
Label1.Caption = ((1 / (tmrE - tmrS)) * (UBound(testmem1) + 1) * 4 * (redo + 1)) / 1024 / 1024
End Sub

Private Sub Command2_Click()
tmrS = Timer
For i = 0 To redo
CopyMemory VarPtr(testmem1(0)), VarPtr(testmem2(0)), (UBound(testmem2) + 1) * 4
Next i
tmrE = Timer
On Error Resume Next
Label1.Caption = ((1 / (tmrE - tmrS)) * (UBound(testmem1) + 1) * 4 * (redo + 1)) / 1024 / 1024
End Sub

Private Sub Command3_Click()
tmrS = Timer
For i = 0 To redo
cpy VarPtr(testmem1(0)), VarPtr(testmem2(0)), (UBound(testmem2)) * 4
Next i
tmrE = Timer
On Error Resume Next
Label1.Caption = ((1 / (tmrE - tmrS)) * (UBound(testmem1) + 1) * 4 * (redo + 1)) / 1024 / 1024
End Sub
