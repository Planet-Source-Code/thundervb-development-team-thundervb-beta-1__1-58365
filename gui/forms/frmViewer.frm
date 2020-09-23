VERSION 5.00
Begin VB.Form frmViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "caption"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   500
      Left            =   1750
      TabIndex        =   3
      Top             =   4900
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   500
      Left            =   3300
      TabIndex        =   2
      Top             =   4900
      Width           =   1800
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   500
      Left            =   100
      TabIndex        =   1
      Top             =   4920
      Width           =   1800
   End
   Begin VB.TextBox txtData 
      Height          =   4700
      Left            =   100
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   100
      Width           =   5000
   End
End
Attribute VB_Name = "frmViewer"
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

Private sData As String, sOData As String

'sCaption - form's caption
'sText    - TextBox's text
'bViewer  - True  - only show text
'         - False - user could change text
'bCanRetNull - True - If cancel/close then return null
'            - Fasle - If cancel/close then return original value
'
'return   - ""   - when bViewer = False or user hits Cancel button
'         - text - when bViewer = True and user hits Save button

Public Function ShowViewer(sCaption As String, sText As String, Optional bViewer As Boolean = True, Optional bCanRetNull As Boolean) As String
    
    'change caption
    frmViewer.caption = sCaption
    
    'un/lock text box
    txtData.Locked = bViewer
    'set text
    txtData.Text = sText
    
    If bCanRetNull = False Then
        sOData = sText
    Else
        sOData = vbNullString
    End If
    
    'if we only want to show text
    If bViewer = True Then
    
        'show close button
        cmdClose.visible = True
        
        'other make unvisible
        cmdSave.visible = False
        cmdCancel.visible = False
        
    Else
    
        'make unvisible
        cmdClose.visible = False
        
        'make visible
        cmdSave.visible = True
        cmdCancel.visible = True
        
    End If
    sData = sOData
    Me.show
    DoEvents
    
    Do While Me.visible
        Dim tmsg As msg
                
        
        If A_PeekMessage(tmsg, 0&, 0&, 0&, PM_REMOVE) Then
        
        'If tmsg.hwnd = Me.hwnd Or tmsg.hwnd = txtData.hwnd Or _
        '   tmsg.hwnd = cmdSave.hwnd Or tmsg.hwnd = cmdClose.hwnd Or tmsg.hwnd = cmdCancel.hwnd Then
           
            Call TranslateMessage(tmsg)
            Call A_DispatchMessage(tmsg)
            SetWindowPos Me.hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            
        End If
    Loop
    
    ShowViewer = sData
    sText = ShowViewer
    
    Unload Me
    
End Function

'user does not want to change text
Private Sub cmdCancel_Click()

    sData = sOData
    Me.hide
    
End Sub

'user close window
Private Sub cmdClose_Click()

    sData = sOData
    Me.hide
    
End Sub

'user changes text
Private Sub cmdSave_Click()

    sData = txtData.Text
    Me.hide
    
End Sub

Private Sub Form_Activate()
    LogMsg "Loading " & Add34(Me.caption) & " window", Me.name, "Form_Activate", True, True
End Sub

'catch unload event
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'check unloading event
   ' If UnloadMode <> vbFormCode Then
   '
    '    cmdCancel_Click
    '    Cancel = False
    '
    'End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LogMsg "Unloading " & Add34(Me.caption) & " window", Me.name, "Form_Unload", True, True
End Sub
