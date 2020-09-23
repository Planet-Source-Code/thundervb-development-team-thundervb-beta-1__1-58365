VERSION 5.00
Begin VB.Form frmDllMainTemp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DllMain template"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdToCursor 
      Caption         =   "Cursor"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdToClipboard 
      Caption         =   "Clipboard"
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox chbDllMain 
      Caption         =   "Add DllMain function"
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CheckBox chbConst 
      Caption         =   "Add constants"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Paste to"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   585
   End
End
Attribute VB_Name = "frmDllMainTemp"
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

'13.09. 2004 - initial version - GUI, code

'constants
Private Const CONST_DLL As String = "Private Const DLL_PROCESS_ATTACH As Long = 1" & vbCrLf & _
                                    "Private Const DLL_PROCESS_DETACH As Long = 0" & vbCrLf & _
                                    "Private Const DLL_THREAD_ATTACH As Long = 2" & vbCrLf & _
                                    "Private Const DLL_THREAD_DETACH As Long = 3"

'DllMain
Private Const DLLMAIN_TEMPLATE As String = "Public Function DllMain(ByVal hInstDLL As Long, ByVal fdwReason As Long, ByVal lpvReserved As Long) As Long" & vbCrLf _
                                          & vbTab & "Select Case fdwReason" & vbCrLf _
                                          & vbTab & vbTab & "Case DLL_PROCESS_ATTACH" & vbCrLf _
                                          & vbTab & vbTab & "Case DLL_PROCESS_DETACH" & vbCrLf _
                                          & vbTab & vbTab & "Case DLL_THREAD_ATTACH" & vbCrLf _
                                          & vbTab & vbTab & "Case DLL_THREAD_DETACH" & vbCrLf _
                                          & vbTab & "End Select" & vbCrLf _
                                          & "End Function"

'dllmain name - default
Private Const DLLMAIN_NAME As String = "DllMain"

'----------------------
'--- CONTROL EVENTS ---
'----------------------

Private Sub Form_Initialize()
    LogMsg "Loading " & Add34(Me.caption) & " window", Me.name, "Form_Initialize", True, True
End Sub

Private Sub Form_Terminate()
    LogMsg "Unloading " & Add34(Me.caption) & " window", Me.name, "Form_Terminate", True, True
End Sub

'close
Private Sub cmdClose_Click()
    Unload Me
End Sub

'paste code to clipboard
Private Sub cmdToClipboard_Click()
Dim sCode As String
    
    sCode = GetCode
    'replace DllMain with user defined
    If Len(frmSettings.DLL_EntryPoint.Text) <> 0 Then sCode = Replace(sCode, DLLMAIN_NAME, frmSettings.DLL_EntryPoint.Text)
    
    LogMsg "Pasting code to clipboard", Me.name, "cmdToClipboard_Click", True, True
    'save to clipboard
    Clipboard.Clear
    Clipboard.SetText sCode
    
End Sub

'paste template to cursor
Private Sub cmdToCursor_Click()
Dim sCode As String

    sCode = GetCode
    'replace DllMain with user defined
    If Len(frmSettings.DLL_EntryPoint.Text) <> 0 Then sCode = Replace(sCode, DLLMAIN_NAME, frmSettings.DLL_EntryPoint.Text)

    LogMsg "Pasting code to cursor location", Me.name, "cmdToCursor_Click", True, True
    'paste to cursor
    PutStringToCurCursor GetCode
    
End Sub

'------------------------
'--- HELPER FUNCTIONS ---
'------------------------

'return template
Private Function GetCode() As String
    
    'add constants
    If chbConst.value = 1 Then GetCode = CONST_DLL
    
    'add dllmain
    If chbDllMain.value = 1 Then
        
        If Len(GetCode) <> 0 Then
            'add new line
            GetCode = GetCode & CrLf(2) & DLLMAIN_TEMPLATE
        Else
            GetCode = DLLMAIN_TEMPLATE
        End If
        
    End If

End Function
