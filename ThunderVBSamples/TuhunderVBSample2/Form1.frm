VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThunderVB sample 2"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Do not use Unicode Text"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   1680
      Width           =   3615
   End
   Begin VB.HScrollBar nrot 
      Height          =   255
      LargeChange     =   10
      Left            =   960
      Max             =   100
      TabIndex        =   5
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "Rot times"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "output"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Input"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "VB->C and back String Convertions ...."
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As Boolean

Private Sub Check1_Click()

    If Check1.Value Then
        A = True
    Else
        A = False
    End If
    
End Sub

Private Sub nrot_Change()

    Text1_Change

End Sub

Private Sub nrot_Scroll()

    nrot_Change
    
End Sub

Private Sub Text1_Change()

    If Len(Text1.Text) > 0 Then
        If A Then
            Dim st() As Byte, outp() As Byte
            st = ConvToCString(Text1.Text)
            outp = st
            StrRot_A st(0), UBound(st), outp(0), nrot.Value
            Label2.Caption = ConvToVBString(outp)
        Else
            Dim temp As String
            temp = Text1.Text
            StrRot_W StrPtr(Text1.Text), Len(temp), StrPtr(temp), nrot.Value
            Label2.Caption = temp
        End If
    End If
    
End Sub
