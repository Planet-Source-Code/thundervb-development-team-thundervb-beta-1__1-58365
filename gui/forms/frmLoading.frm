VERSION 5.00
Begin VB.Form frmLoading 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   DrawMode        =   6  'Mask Pen Not
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00C00000&
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loading ThunVB , Pelase wait"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   590
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmLoading"
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

'3/10/2004[dd/mm/yyyy] : Created by Raziel
'Intial code
'
'5/10/2004[dd/mm/yyyy] : Edited by Raziel
'now , VB IDE does not losses focus
'

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Sub Form_Load()
    
    TopMost True
    SetWindowLong Me.hWnd, GWL_STYLE, GetWindowLong(Me.hWnd, GWL_STYLE) Or WS_CHILDWINDOW
    SetParent Me.hWnd, GetCodeWinPar()  'GetDesktopWindow()
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    TopMost False
    
End Sub

Private Function TopMost(bTopMost As Boolean) As Boolean

   If bTopMost = True Then 'Make the window topmost
      TopMost = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
   Else
      TopMost = SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
      TopMost = False
   End If

End Function
