VERSION 5.00
Begin VB.Form frmTip 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "frmTip"
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   0
      ScaleHeight     =   1290
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   15
      Width           =   15
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4095
      TabIndex        =   3
      Top             =   1305
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4095
      TabIndex        =   2
      Top             =   0
      Width           =   4095
   End
   Begin VB.PictureBox Picture4 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   4080
      ScaleHeight     =   1290
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   15
      Width           =   15
   End
   Begin VB.PictureBox pbc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   30
      ScaleHeight     =   1095
      ScaleWidth      =   3975
      TabIndex        =   5
      Top             =   30
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label code 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "[code]"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frmTip"
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
'15/9/2004[dd/mm/yyyy] : Created by Raziel
'Works + coloring..
'
'5/10/2004[dd/mm/yyyy] : Edited by Raziel
'now , VB IDE does not losses focus
'
'22/10/2004 , using PictureBox to make text bold ect..

Private Declare Function GetCursorPos Lib "user32.dll" ( _
     ByRef lpPoint As POINTAPI) As Long
     

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE

Private Type nmhdr
  hwndFrom As Long
  idFrom As Long
  code As Long
End Type

Private Type REQSIZE
    nmhdr As nmhdr
    RECT As RECT
End Type

Dim ptemp As POINTAPI
      
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Function TopMost(bTopMost As Boolean) As Boolean

   If bTopMost = True Then 'Make the window topmost
      TopMost = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
   Else
      TopMost = SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
      TopMost = False
   End If

End Function

Sub EnToolTipA()
        
        GetCursorPos ptemp
        ScreenToClient GetParent(Me.hWnd), ptemp
        ShowTooltip ptemp.x + 35, ptemp.y + 20, "", False, True
        
End Sub

Sub ShowTooltip(x As Long, y As Long, str As String, bstrOlny As Boolean, Optional bXYolny As Boolean = False, Optional pb As Boolean = False)

    If Me.visible = False Then
        TopMost True
        Me.visible = True
    End If
    
    If bstrOlny = False Then
        Me.Top = y * 15
        Me.Left = x * 15
    End If
    
    If bXYolny Then
        If Len(Trim$(code.caption)) = 0 Then HideToolTip
    Else
        SetText str, pb
    End If
    
End Sub

Sub HideToolTip()

        If Me.visible = True Then
            TopMost False
            Me.visible = False
            wm_tthide = False
        End If
        
End Sub

Sub ShowToolTip_l()

        If Me.visible = False Then
            TopMost True
            Me.visible = True
        End If
        
End Sub

Private Sub SetText(strText As String, Optional pb As Boolean = False)
    
    If Len(Trim$(strText)) = 0 Then HideToolTip: Exit Sub
    If Replace(strText, vbTab, "    ") <> code.caption Then
        code.caption = Replace(strText, vbTab, "    ")
        Me.Height = code.Height + 45
        Me.Width = code.Width + 45
        If pb Then
            code.visible = False
            pbc.visible = True
            pbc.Cls
            pbc.Height = pbc.TextHeight(strText)
            pbc.Width = pbc.TextWidth(strText)
            Me.Height = pbc.Height + 60
            Me.Width = pbc.Width + 60
        Else
            pbc.visible = False
            code.visible = True
        End If
    End If
    
End Sub

Private Sub code_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_GotFocus
End Sub

Private Sub code_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_GotFocus
End Sub

Private Sub Form_Load()

    SetWindowLong Me.hWnd, GWL_STYLE, GetWindowLong(Me.hWnd, GWL_STYLE) Or WS_CHILDWINDOW
    SetParent Me.hWnd, GetCodeWinPar() 'GetDesktopWindow()

End Sub

Private Sub Form_GotFocus()
On Error Resume Next

    VBI.ActiveWindow.SetFocus
    
End Sub


Private Sub pbc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_GotFocus
End Sub

Private Sub pbc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_GotFocus
End Sub
