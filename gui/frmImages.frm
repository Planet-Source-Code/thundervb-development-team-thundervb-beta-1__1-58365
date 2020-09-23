VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmImages 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtb 
      Height          =   345
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmImages.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrToolTip 
      Interval        =   345
      Left            =   120
      Top             =   1080
   End
   Begin VB.PictureBox picImg1 
      Height          =   345
      Left            =   480
      Picture         =   "frmImages.frx":00A7
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox picImg 
      AutoSize        =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   120
      Picture         =   "frmImages.frx":1087
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   720
      Width           =   345
   End
   Begin VB.Label Label1 
      Caption         =   "Form to Store image resources needed , timer and any other controls that are needed on non form  code [like rtf coloring]"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmImages"
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

Private Sub tmrToolTip_Timer()

    'modSubCH.WindowProcAft 0, WM_CHAR, 0, 0, 0, 0
    'dat = LoadIsmlFile("c:\asmdefs.isml")
    CheckToolTip
    'If Get_General(SetTopMost) Then
    '    If frmSettings.visible Then
    '        SetWindowPos frmSettings.hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    '    End If
    'End If
    
End Sub
