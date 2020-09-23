Attribute VB_Name = "modMain"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

'Load/Unload/ addin things halding

'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Raziel
'Module created , intial version
'
'
'2/9/2004[dd/mm/yyyy] : Edited by Raziel
'Many changes , more stable code
'
'
'Notes.. This file is edeited here and there all the time to
'        Support more things/better loading ect..
'        most of em are not logged cause they are too many/non important

Option Explicit
Declare Function GlobalAddAtom Lib "kernel32.dll" Alias "GlobalAddAtomA" ( _
     ByVal lpString As String) As Integer
     
Global VBI As VBIDE.VBE

Global hk_CopyHotKeyID As Long
Global hk_CutHotKeyID As Long
Global hk_CtrlIHotKeyID As Long
Global hk_CtrlSpaceHotKeyID As Long

Dim CpyHK As New clsHotKey
Dim CutHK As New clsHotKey
Dim C_I As New clsHotKey
Dim C_sp As New clsHotKey

Sub AddinLoaded()

    
    'If App.LogMode = 1 Then
    '    GetOldAddress MainhWnd
    'End If
    
    LogMsg "Startup Finished , Loading settings", "modMain", "AddinLoaded"
    LoadSettings
     
End Sub

Sub AddinUnload(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    
    'if not in ide
    If App.LogMode = 1 Then
        KillLinkerHook
        KillAsmColorHook
        StopGPFHandler
    End If
    
    FlushLog
    
End Sub

Sub AddinLoad(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
Dim sbuff As String * 260
    StartGPFHandler
    
    GetModuleFileName GetModuleHandle("ThunderVB.dll"), sbuff, 260
        
    Let_Paths AddIn_Directory, GetPath(sbuff)
    
    LogMsg "Starting Up", "modMain", "AddinLoad"
    

        
    
    
    LogMsg Get_Paths(AddIn_Directory), "modMain", "ProjectActivated"
    
    vb_Dll_version = getVBVersion
    'if not in ide
    If App.LogMode = 1 Then
        LogMsg "Setting Hooks", "modMain", "AddinLoad"
        InitLinkerHook
        InitAsmColorHook
        dat = LoadIsmlFile(Get_Paths(AddIn_Directory) & "asmdefs.isml")
    Else
        MsgBox "Please Compile First"
        'dat = LoadIsmlFile("c:\asmdefs.isml")
    End If
    LogMsg "Isml file kw count : " & dat.kw_count, "", ""
    
'    MsgBox Replace(kwListToString(dat), "|", vbNewLine)
    
    If ConnectMode = ext_cm_AfterStartup Then
        AddinLoaded
    End If
    
End Sub

Sub MenuClicked()
'On Error Resume Next
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Const flags = SWP_NOMOVE Or SWP_NOSIZE
    

    If CBool(Get_General(SetTopMost)) Then
    
        frmSettings.show 'vbModal
        Call SetWindowPos(frmSettings.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
        'frmSettings.hide
        'Unload frmSettings
    Else
        frmSettings.show
        Call SetWindowPos(frmSettings.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    End If
    
End Sub

Sub ButtonClicked(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    MenuClicked 'frmSettings.show
    
End Sub

Public Sub ProjectActivated(VBProject As VBProject)

    LogMsg "Loading new project settings", "modMain", "ProjectActivated"
    
    If Len(VBProject.filename) Then
        Let_Paths Project_Directory, GetPath(VBProject.filename)
        Let_Paths Debug_Directory, GetPath(VBProject.filename) & "debug\"
    Else
        WarnBox "Canot find project Directory" & vbNewLine & _
        "if new project then save it and restart ThunderVB", "modMain", "ProjectActivated"
    End If
    
    LoadSettings
    
    LogMsg Get_Paths(Project_Directory), "modMain", "ProjectActivated"
    LogMsg Get_Paths(Debug_Directory), "modMain", "ProjectActivated"
    
End Sub

Public Sub ProjectAdded(VBProject As VBProject)

    ProjectActivated VBProject
    
End Sub

Public Sub ProjectRemoved(VBProject As VBProject)

    
    
End Sub

Public Sub ProjectRenamed(VBProject As VBProject, OldName As String)

    
    
End Sub
