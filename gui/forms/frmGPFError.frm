VERSION 5.00
Begin VB.Form frmGPFError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "An unhandled error ocured...."
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSav 
      Caption         =   "Save Project(s)"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop execution"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCont 
      Caption         =   "Continue"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdRasieErr 
      Caption         =   "RaiseError"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "[Note the VBide is/is not locked]"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "[ext err str info]"
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Extended error info :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   $"frmGPFError.frx":0000
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmGPFError"
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

Dim hf As GPF_actions

Public Function ShowGPF(str As String) As GPF_actions
On Error GoTo errH
    
    Me.Label3.caption = str
    Me.Label4 = "Note : The VB IDE IS NOT freesed"
    Me.show
    Do
        Sleep 10
        DoEvents
    Loop While Me.visible = True
    ShowGPF = hf
Exit Function

errH:
    Me.Label4 = "Note : The VB IDE is FREESED"
    Me.show vbModal
    Resume Next
    
End Function

Private Sub cmdCont_Click()

    hf = GPF_Cont
    Me.hide
    
End Sub

Private Sub cmdRasieErr_Click()

    hf = GPF_RaiseErr
    Me.hide
    
End Sub

Private Sub cmdSav_Click()
    
    SaveProjects True
    
End Sub

Private Sub cmdStop_Click()

    hf = GPF_Stop
    Me.hide
    
End Sub

