VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11325
   _ExtentX        =   19976
   _ExtentY        =   17701
   _Version        =   393216
   Description     =   "ThunderVB (aka ThunVB) adiin for VB6"
   DisplayName     =   "ThunderVB_npl"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ThunderVB"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

Public mcbMenuCommandBar As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents prjevents As VBProjects
Attribute prjevents.VB_VarHelpID = -1
Public WithEvents BuildEvents As VBIDE.VBBuildEvents
Attribute BuildEvents.VB_VarHelpID = -1
Dim buts() As clsComBut
Dim SettingsID As Long
Dim GenerateID As Long
Dim TemplateID As Long
Dim ConvertID As Long
Dim ExplorerID As Long
Dim RunTimeID As Long

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
    
    ApiTimer True

End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    MainhWnd = 0
    ReDim buts(0)

    'save the vb instance
    Set VBI = Application
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    If ConnectMode <> ext_cm_External Then
        Set mcbMenuCommandBar = AddToAddInCommandBar("ThunderVB")
        'sink the event
        Set Me.MenuHandler = VBI.Events.CommandBarEvents(mcbMenuCommandBar)
        SettingsID = AddButtonToToolbar("ThunVB Settings", "None", frmImages.picImg(0).Picture)
        GenerateID = AddButtonToToolbar("ThunVB ASM Code Generator", "None", frmImages.picImg1.Picture)
        TemplateID = AddButtonToToolbar("ThunVB ASM Templates", "None", frmImages.picImg1.Picture)
        ExplorerID = AddButtonToToolbar("ThunVB ASM Code Explorer", "None", frmImages.picImg1.Picture)
        RunTimeID = AddButtonToToolbar("ThunVB ASM RunTime engine Hooking", "None", frmImages.picImg1.Picture)
        AddButtonToMenu "ThunVB", "ASM Code Generator", Nothing, GenerateID
        buts(UBound(buts) - 1).cbarObj.BeginGroup = True
        AddButtonToMenu "ThunVB", "ASM Templates", Nothing, TemplateID
        ConvertID = AddButtonToMenu("ThunVB", "Expand asm/c blocks", Nothing)
    End If
  

    MainhWnd = GetCodeWinPar()
    ApiTimer False
        
    Call AddinLoad(Application, ConnectMode, AddInInst, custom())
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    ApiTimer True
    
    KillAllSubClasses
    
    'On Error Resume Next
    Call AddinUnload(RemoveMode, custom())
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    Dim i As Long
    
    For i = 0 To UBound(buts)
        Set buts(i) = Nothing
    Next i
        
End Sub


Private Sub AddinInstance_OnStartupComplete(custom() As Variant)

    Set prjevents = VBI.VBProjects
    'this is a nice class ., but how to get a haldle for it?
    'Set BuildEvents = Nothing
    Call AddinLoaded
    
    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveVBProject Is Nothing Then Exit Sub
    MainhWnd = GetCodeWinPar()
    modMain.ProjectAdded VBI.ActiveVBProject
    
End Sub

Public Sub cmd_Click(id As Long, CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
    Select Case id
        
        Case SettingsID
            ButtonClicked CommandBarControl, handled, CancelDefault
        Case TemplateID
            frmTemplates.show
        Case GenerateID
            frmCodeWizard.show
        Case ConvertID
            If (VBI Is Nothing) Then Exit Sub
            If (VBI.ActiveCodePane Is Nothing) Then Exit Sub
            Dim i As Long
            AsmBlocks_ConvertLine_init
            With VBI.ActiveCodePane.codeModule
                For i = 1 To .CountOfLines
                    .ReplaceLine i, CBlocks_ConvertLine(AsmBlocks_ConvertLine(.lines(i, 1)))
                Next i
            End With
        Case ExplorerID
            frmASMExplorer.show
        Case RunTimeID
            frmRunTime.show
    End Select

End Sub

Private Sub BuildEvents_BeginCompile(ByVal VBProject As VBIDE.VBProject)

    'BuildStarted
    
End Sub

Private Sub BuildEvents_EnterDesignMode()
    
    'DesignStarted
    
End Sub

Private Sub BuildEvents_EnterRunMode()
    
    'RunStarted
    
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    MenuClicked
End Sub


'------------------------
'--- Helper Functions ---
'------------------------
Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBI.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

Private Sub prjevents_ItemActivated(ByVal VBProject As VBIDE.VBProject)

    ProjectActivated VBProject
    
End Sub

Private Sub prjevents_ItemAdded(ByVal VBProject As VBIDE.VBProject)

    ProjectAdded VBProject
    
End Sub

Private Sub prjevents_ItemRemoved(ByVal VBProject As VBIDE.VBProject)

    ProjectRemoved VBProject
    
End Sub

Private Sub prjevents_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)

    ProjectRenamed VBProject, OldName
    
End Sub

Function AddButtonToToolbar(tip As String, capt As String, defimage As IPictureDisp, Optional id As Long = -1) As Long
Dim ClipBoardBackup As IPictureDisp, temp As CommandBarControl


    Set temp = VBI.CommandBars(2).Controls.Add(msoControlButton)
    temp.caption = capt
    On Error Resume Next
    If Not (defimage Is Nothing) Then
        If Clipboard.GetFormat(vbCFBitmap) Then
            Set ClipBoardBackup = Clipboard.GetData(vbCFBitmap)
        Else
            Set ClipBoardBackup = New StdPicture
        End If
        Clipboard.SetData defimage, vbCFBitmap
        temp.PasteFace
        Clipboard.SetData ClipBoardBackup, vbCFBitmap
    End If
    On Error GoTo 0
    temp.ToolTipText = tip
    
    Set buts(UBound(buts)) = New clsComBut
    If id = -1 Then
        AddButtonToToolbar = buts(UBound(buts)).init(temp, Me, UBound(buts))
    Else
        AddButtonToToolbar = buts(UBound(buts)).init(temp, Me, id)
    End If
    
    ReDim Preserve buts(UBound(buts) + 1)
    
End Function

Function AddButtonToMenu(tip As String, capt As String, defimage As IPictureDisp, Optional id As Long = -1) As Long
Dim oldclip As Clipboard, temp As CommandBarControl

    Set temp = VBI.CommandBars("Code Window").Controls.Add(msoControlButton)
    temp.caption = capt
    On Error Resume Next
    If Not (defimage Is Nothing) Then
        'Set oldclip = Clipboard
        Clipboard.SetData defimage
        temp.PasteFace
        'Clipboard = oldclip
    End If
    On Error GoTo 0
    temp.ToolTipText = tip
    
    Set buts(UBound(buts)) = New clsComBut
    If id = -1 Then
        AddButtonToMenu = buts(UBound(buts)).init(temp, Me, UBound(buts))
    Else
        AddButtonToMenu = buts(UBound(buts)).init(temp, Me, id)
    End If
    ReDim Preserve buts(UBound(buts) + 1)

End Function
