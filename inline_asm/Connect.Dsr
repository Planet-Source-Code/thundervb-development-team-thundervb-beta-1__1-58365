VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7560
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11835
   _ExtentX        =   20876
   _ExtentY        =   13335
   _Version        =   393216
   Description     =   ".....Nothing yet...."
   DisplayName     =   "STD Dll Creator"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
'Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmAddIn                 As New frmAddIn
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents vbProj As VBIDE.VBProjects
Attribute vbProj.VB_VarHelpID = -1

Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    End If

    FormDisplayed = True
    mfrmAddIn.Show
   
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    'save the vb instance
    Set VBInstance = Application
    Set Connect = Me
    Set vbProj = VBInstance.VBProjects
    'this is a good place to set a breakpoint and
    LoadAddinSettings
    'test various addin objects, properties and methods
    'Debug.Print VBInstance.FullName
    ReDim exports(0)
    If App.LogMode <> 0 Then
        Hook "VBA6.DLL", AddressOf CreateProcess_Hook, 0
        colorASM_init
    End If
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar(AddInName_full)
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    If App.LogMode <> 0 Then
        Hook "VBA6.DLL", AddressOf CreateProcess_Hook, 1
        colorASM_Close
    End If
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    Set vbProj = VBInstance.VBProjects
    frmTimer.tmrSaved.Interval = 100
    frmTimer.tmrSaved.Enabled = True
    
    If FileExist(masm_exe) = False Then
        MsgBox "Masm Path is wrong" & vbNewLine & _
               "Please Set it corectly" & vbNewLine & _
               "The curect path is : " & masm_exe, vbOKOnly
        frmAddIn.Show
    End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

Private Sub vbProj_ItemActivated(ByVal VBProject As VBIDE.VBProject)
modSettings.LoadProjectSettings VBInstance.ActiveVBProject
End Sub

Private Sub vbProj_ItemAdded(ByVal VBProject As VBIDE.VBProject)
modSettings.LoadProjectSettings VBInstance.ActiveVBProject
End Sub
