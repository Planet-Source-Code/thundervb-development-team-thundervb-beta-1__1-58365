VERSION 5.00
Begin VB.Form IntelliAsm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   234
   ShowInTaskbar   =   0   'False
   Begin ThunderVB.IntelliSense iSe 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2820
      _ExtentX        =   4763
      _ExtentY        =   3387
   End
End
Attribute VB_Name = "IntelliAsm"
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
'someday[dd/mm/yyyy] : Created by Raziel
'Module Created , intial version
'just a empty file
'
'
'5/10/2004[dd/mm/yyyy] : edited by Raziel
'form show/hide with, form style set with winapi
'form resize handling
'
'
'6/10/2004[dd/mm/yyyy] : Created by Raziel
'Found a workaround on form positioning problems
'main sceleton is finished , started work on
'modIntelliAsm as it will be the core for
'the intelliAsm code

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE


Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                                    ByVal hWnd As Long, ByVal nIndex As Long, _
                                    ByVal dwNewLong As Long) As Long
                                    
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                                    ByVal hWnd As Long, _
                                    ByVal nIndex As Long) As Long
Dim oldlist As String, bDoNotHide As Boolean

Private Function TopMost(bTopMost As Boolean) As Boolean

   If bTopMost = True Then 'Make the window topmost
      TopMost = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
   Else
      TopMost = SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
      TopMost = False
   End If

End Function

Private Sub Form_GotFocus()
On Error Resume Next
    
    Dim t As String
    t = oldlist
    VBI.ActiveWindow.SetFocus
    If wm_tthide Then
        asmTT.ShowToolTip_l
    End If
    If Len(t) Then
        SetList t
        Me.visible = True
    End If
    
End Sub

Private Sub Form_Load()
    
    SetWindowLong Me.hWnd, GWL_STYLE, GetWindowLong(Me.hWnd, GWL_STYLE) Or WS_CHILDWINDOW
    SetParent Me.hWnd, GetCodeWinPar() 'GetDesktopWindow() ' i think it is not needed but i like it :D
    
End Sub

Private Sub Form_Resize()

    iSe.Width = Me.ScaleWidth
    iSe.Height = Me.ScaleHeight
    
End Sub

Sub SetTip(strTip As String)
Dim i As Long, ml As Double
    
    If Me.visible Then
        If IsNumeric(strTip) Then strTip = "#"
        strTip = Replace$(Replace$(Replace$(strTip, "[", " "), "]", " "), "_", " ")
        With iSe.memb_list
            For i = 0 To .ListCount - 1
                If fstrc(.list(i), strTip) > ml Then
                    ml = fstrc(.list(i), strTip)
                    .ListIndex = i
                End If
            Next i
        End With
    End If
    
End Sub

Function fstrc(strm As String, strTip As String) As Double
Dim mc As Long, i As Long
    mc = IIf(Len(strm) > Len(strTip), Len(strTip), Len(strm))
    
    For i = 1 To mc
        If AscW(Mid$(strm, i, 1)) <> AscW(Mid$(strTip, i, 1)) Then
            Exit For
        End If
    Next i
    On Error Resume Next
    fstrc = i - 1
    
End Function

Sub ShowIntelliAsm(mhWnd As Long, x As Long, y As Long, strTip As String, strList As String)
    Dim tlp As POINTAPI, st() As String, i As Long, wrect As RECT, cp As POINTAPI
'On Error Resume Next
    If Len(strList) > 0 Then
        
        SetList strList
        tlp.x = x: tlp.y = y
        ClientToScreen mhWnd, tlp
        ScreenToClient GetParent(Me.hWnd), tlp
        Call GetWindowRect(GetParent(Me.hWnd), wrect)
        If (tlp.x + 180) > wrect.Right Then
            tlp.x = wrect.Right - 180
        End If
        If (tlp.y + 140) > wrect.Bottom Then
            GetCaretPos cp
            tlp.y = cp.y - 68
        End If
        SetWindowPos Me.hWnd, 0, tlp.x, tlp.y, 180, 128, 0
        
        If Me.visible = False Then
            Me.show
            TopMost True
        End If
        
        SetTip strTip
    Else
        HideIntelliAsm
    End If
    
End Sub

Sub SetList(strnewlist As String)
Dim d1 As String_B
    
    If strnewlist <> oldlist Then
        Dim st() As String, i As Long
        
        oldlist = strnewlist
        iSe.memb_list.Clear
        If Len(strnewlist) > 0 Then
            st = Split(strnewlist, "|")
            
            With iSe.memb_list
                .Clear
                For i = 0 To UBound(st)
                    AddToLBNoDup Trim$(st(i)), d1
                Next i
                SendMessage .hWnd, WM_SETFOCUS, .hWnd, .hWnd
            End With
        Else
            HideIntelliAsm
        End If
    End If

End Sub

Sub MoveTo(mhWnd As Long, x As Long, y As Long)
    Dim tlp As POINTAPI
    
    If Me.visible = True Then
        
        tlp.x = x: tlp.y = y
        ClientToScreen mhWnd, tlp
        ScreenToClient GetParent(Me.hWnd), tlp
        SetWindowPos Me.hWnd, 0, tlp.x, tlp.y, 180, 128, 0
        
    End If
    
End Sub

Sub HideIntelliAsm()
    
    If bDoNotHide = False Then
        TopMost False
        'oldlist = ""
        Me.hide
    End If
    
End Sub

Private Sub iSe_GotFocus()
    
    Form_GotFocus

End Sub

Private Sub iSe_GotFocusList()
    
    Form_GotFocus
    
End Sub

Private Sub iSe_UserSelected(ByVal strd As String)

    CompleteText strd
    Me.HideIntelliAsm
    Form_GotFocus
    
End Sub

Private Sub AddToLBNoDup(strItem As String, d1 As String_B)

    If d1.str_index > 0 Then
        If InStr(1, "|" & GetString(d1) & "|", "|" & strItem & "|") = 0 Then
            AppendString d1, strItem & "|"
            iSe.memb_list.AddItem strItem
        End If
    Else
        AppendString d1, strItem
        iSe.memb_list.AddItem strItem
    End If
    
End Sub
