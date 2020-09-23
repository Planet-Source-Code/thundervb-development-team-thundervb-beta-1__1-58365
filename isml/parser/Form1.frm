VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_DblClick()
Dim t As Long, t2 As Isml_File
    t = FreeFile
    
    Open "c:\test.isml" For Output As t
    Close t
    
    Open "c:\test.isml" For Binary As t
        Put t, , Me.Text1.Text
    Close t
    
    t2 = LoadIsmlFile("c:\test.isml")
    
End Sub

