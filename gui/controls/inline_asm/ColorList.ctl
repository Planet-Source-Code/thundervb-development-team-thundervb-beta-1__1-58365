VERSION 5.00
Begin VB.UserControl ColorList 
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   ScaleHeight     =   3540
   ScaleWidth      =   4935
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   1800
      Top             =   2040
   End
   Begin VB.PictureBox ppb 
      Height          =   495
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picCol 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   0
      ScaleHeight     =   3510
      ScaleWidth      =   210
      TabIndex        =   2
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox lnEnd 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   4920
      ScaleHeight     =   3510
      ScaleWidth      =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15
   End
   Begin VB.ListBox lstCol 
      Appearance      =   0  'Flat
      Height          =   1230
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "ColorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit
'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Raziel
'Control created , intial version
'
'20/8/2004[dd/mm/yyyy] : Edited by Raziel
'Added colors before the names..
'
'26/8/2004[dd/mm/yyyy] : Edited by Raziel
'Added text edit [when the user clicks on the list , to edit it]
'

Public Event ChangeColor(oldCol As Long, newcol As Long, bCancel As Boolean, bHandled As Boolean)
Public Event ColorSelected(Index As Long)
Public Event LineEdited(Index As Long)
Dim oldindx As Long, nested As Long, lstclick As Double
Dim edit_ind As Long, last_ind As Long
Private Sub lstCol_Click()
    
    If (Timer - lstclick) < 6 And (Timer - lstclick) > 0.5 And last_ind = lstCol.ListIndex Then
        last_ind = lstCol.ListIndex
        EditText lstCol.ListIndex
    Else
        last_ind = lstCol.ListIndex
        lstclick = Timer
        RaiseEvent ColorSelected(lstCol.ListIndex)
        'lblCol.BackColor = lstCol.ItemData(lstCol.ListIndex)
    End If
    
End Sub

Private Sub lstCol_DblClick()
Dim newcol As Long, Cancel As Boolean, bHandled As Boolean
    
    Text1_LostFocus
    RaiseEvent ChangeColor(lstCol.ItemData(lstCol.ListIndex), newcol, Cancel, bHandled)
    If Cancel Or (newcol = lstCol.ItemData(lstCol.ListIndex)) Or (bHandled = False) Then Exit Sub
    lstCol.ItemData(lstCol.ListIndex) = newcol
    lstCol_Click
    Text1_LostFocus
    
End Sub

Private Sub lstCol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    DrawCols

End Sub

Private Sub lstCol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    DrawCols
    
End Sub

Private Sub lstCol_Scroll()

    DrawCols

End Sub

Private Sub lstCol_Validate(Cancel As Boolean)

    lstCol_Scroll
    
End Sub

Private Sub picCol_Paint()

    DrawCols

End Sub



Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then KeyCode = 0: Text1_LostFocus

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then KeyAscii = 0: Text1_LostFocus

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then KeyCode = 0: Text1_LostFocus
    
End Sub

Private Sub Text1_LostFocus()
    
    lstCol.list(edit_ind) = Text1.Text
    Text1.visible = False
    RaiseEvent LineEdited(lstCol.ListIndex)
    
End Sub

Private Sub Timer1_Timer()

    If App.LogMode = 0 Then Timer1.Enabled = False 'in ide..
    If lstCol.TopIndex <> oldindx Then
        oldindx = lstCol.TopIndex
        lstCol_Scroll
    End If
    
End Sub

Private Sub UserControl_Initialize()
    
    Me.SetDefaultsC
    
    UserControl.ScaleMode = 3
    
End Sub

Private Sub UserControl_LostFocus()

    Text1_LostFocus
    
End Sub

Private Sub UserControl_Paint()

    picCol_Paint

End Sub

Private Sub UserControl_Resize()
Dim i As Long

    picCol.Cls
    lstCol.Width = UserControl.Width / 15 - 16
    lstCol.Height = UserControl.Height / 15
    DrawCols
    
End Sub

Public Property Get listdata() As ListBox

    Set listdata = lstCol

End Property


Public Property Let listdata(ld As ListBox)

    lstCol = ld

End Property


Public Function AddColor(str As String, Color As Long) As Long
Dim temp As ColorInfo_entry

    With lstCol
        .AddItem str
        .ItemData(.ListCount - 1) = Color
        AddColor = .ListCount - 1
    End With
    DrawCols
End Function

Public Function RemoveColor(Index As Long)
    With lstCol
        .RemoveItem Index
    End With
    DrawCols
End Function

Public Property Get Color(Index As Long) As String()
Dim temp(1) As String
    
    With lstCol
        temp(0) = .list(Index)
        temp(1) = .ItemData(Index)
    End With
    Color = temp
End Property

Public Property Let Color(Index As Long, value() As String)
    
    With lstCol
      .list(Index) = value(0)
      .ItemData(Index) = value(1)
    End With
    
End Property


Public Property Get ColorInfo() As String
Dim i As Long, out As String_B
If lstCol.ListCount = 0 Then Exit Sub
    With lstCol
        For i = 0 To .ListCount - 1
            AppendString out, .list(i) & "_@#sent@_" & .ItemData(i) & "_@#slst@_"
        Next i
        out.str(out.str_index - 1) = Left(out.str(out.str_index - 1), Len(out.str(out.str_index - 1)) - Len("_@#slst@_"))
    End With
    FinaliseString out
    ColorInfo = GetString(out)
End Property

'regs_@#sent@_1234_@#slst@_
Public Property Let ColorInfo(ld As String)
Dim str() As String, str2() As String, i As Long

    lstCol.Clear
    str = Split(ld, "_@#slst@_")
    For i = 0 To UBound(str)
        If Len(str(i)) Then
            str2 = Split(str(i), "_@#sent@_")
            AddColor str2(0), val(str2(1))
        End If
    Next i
    
End Property

Sub DrawCols()
Dim temp As Long, max As Long, fs As Long, ccx As Long, i As Long

    picCol.Cls
    ccx = 15
    fs = lstCol.FontSize * (12 / 8.25)
    max = lstCol.ListCount
    For i = 0 To (lstCol.ListCount - lstCol.TopIndex)
        temp = (lstCol.TopIndex + i)
        If temp < max Then
            picCol.FillColor = lstCol.ItemData(temp)
            ppb.BackColor = lstCol.ItemData(temp)
            picCol.PaintPicture ppb.Image, ccx, (i * fs) * 15 + (i + 1) * 15, fs * 15, fs * 15
            'picCol.Circle (ccx, 1 * 15 + (i * fs) * 15 + i * 15 + (fs / 2 / 1.2) * 15), (fs / 2 / 1.4) * 15
        End If
    Next i
    picCol.Refresh
    
End Sub

Sub SetDefaultsAsm()

        lstCol.Clear
        AddColor " EAX EBX ECX EDX ", RGB(120, 120, 150)
        AddColor " AX BX CX DX ", RGB(120, 120, 170)
        AddColor " AH AL BH BL CH CL DH DL ", RGB(120, 120, 190)
        AddColor " CS DS ES FS GS SS ", RGB(120, 120, 110)
        AddColor " ESI EDI EBP EIP ESP ", RGB(120, 120, 150)
        AddColor " EFLAGS ", RGB(120, 120, 210)
        AddColor " ; ", RGB(0, 128, 0)
        AddColor "*default*", RGB(0, 0, 140)
        
        AddColor "*" & Add34("string") & "*", RGB(105, 105, 105)
        AddColor "*'string'*", RGB(105, 105, 105)
        AddColor "*Number*", RGB(120, 100, 255)
        
End Sub

'We realy need more here..
Sub SetDefaultsC()
        
        lstCol.Clear
        AddColor " and and_eq asm auto bitand bitor bool break case catch char " & _
                 "class compl const const_cast continue default delete do double " & _
                 "dynamic_cast else enum explicit export extern false float for " & _
                 "friend goto if inline int long mutable namespace new not not_eq " & _
                 "operator or or_eq private protected public register reinterpret_cast " & _
                 "return short signed sizeof static static_cast struct switch " & _
                 "template this throw true try typedef typeid typename union " & _
                 "unsigned using virtual void volatile wchar_t while xor xor_eq ", _
                                                                      RGB(0, 0, 140)
        
        AddColor " // ", RGB(0, 128, 0)
        AddColor "*default*", RGB(0, 0, 0)
        
        AddColor "*" & Add34("string") & "*", RGB(105, 105, 105)
        AddColor "*Number*", RGB(120, 100, 255)
        AddColor "*HexNumber*", RGB(105, 85, 240)
        
End Sub

Sub EditText(ind As Long)

    edit_ind = ind
    Text1.Top = GetY(ind - lstCol.TopIndex)
    Text1.Height = GetHeight
    Text1.Left = lstCol.Left
    Text1.Width = lstCol.Width
    Text1.Text = lstCol.list(ind)
    Text1.visible = True
    Text1.SetFocus
    
End Sub

Function GetY(i As Long) As Long
Dim fs As Long, temp As Long

    fs = lstCol.FontSize * (12 / 8.25)
    temp = (i * fs) + (i + 1)  ' fs * 15
    GetY = temp
    
End Function

Function GetHeight() As Long
Dim fs As Long, temp As Long

    fs = lstCol.FontSize * (12 / 8.25)
    temp = fs
    GetHeight = temp
    
End Function
