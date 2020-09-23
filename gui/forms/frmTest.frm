VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2280
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   2280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "change text"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "show text"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
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

    MsgBox frmViewer.ShowViewer("Viewer", "only show text")

End Sub

Private Sub Command2_Click()

    MsgBox frmViewer.ShowViewer("Modify", "change text", False)
    
End Sub

