VERSION 5.00
Begin VB.UserControl IntelliSense 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FFFF&
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   ScaleHeight     =   2655
   ScaleWidth      =   3885
   Begin VB.PictureBox picBord2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   3885
      TabIndex        =   5
      Top             =   0
      Width           =   3885
   End
   Begin VB.PictureBox picbord1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   0
      ScaleHeight     =   2610
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   15
      Width           =   15
   End
   Begin VB.PictureBox picbord5 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H0099A8AC&
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   3870
      ScaleHeight     =   2580
      ScaleWidth      =   0
      TabIndex        =   3
      Top             =   15
      Width           =   15
   End
   Begin VB.PictureBox picbord3 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H0099A8AC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   3855
      ScaleHeight     =   2610
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   15
      Width           =   15
   End
   Begin VB.PictureBox picbord4 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   3885
      TabIndex        =   1
      Top             =   2640
      Width           =   3885
   End
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H0099A8AC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   3885
      TabIndex        =   0
      Top             =   2625
      Width           =   3885
   End
   Begin VB.ListBox list 
      Appearance      =   0  'Flat
      Height          =   1950
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "IntelliSense"
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

Public memb_list As ListBox
Public Event UserSelected(ByVal strd As String)
Public Event GotFocusList()

Private Sub list_Click()
    
    RaiseEvent GotFocusList
    
End Sub

Public Sub list_DblClick()

    RaiseEvent UserSelected(list.Text)

End Sub

Private Sub list_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent GotFocusList
    
End Sub

Private Sub UserControl_Resize()

    UserControl.list.Width = UserControl.Width - 15
    UserControl.list.Height = UserControl.Height - 15
    Set memb_list = UserControl.list
    
End Sub


Sub SendKey(key As String)

End Sub

