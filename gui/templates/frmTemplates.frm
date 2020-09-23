VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTemplates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASM Templates"
   ClientHeight    =   4815
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwTemp 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6165
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdToCursor 
      Caption         =   "Cursor"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdToClipBoard 
      Caption         =   "Clipboard"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Paste to"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   585
   End
   Begin VB.Menu mnuMainTemplate 
      Caption         =   "Template"
      Begin VB.Menu mnuMainNew 
         Caption         =   "New template"
         Begin VB.Menu mnuEmpty 
            Caption         =   "Empty"
         End
         Begin VB.Menu mnuBasedOn 
            Caption         =   "Based on"
         End
         Begin VB.Menu mnuFromVB 
            Caption         =   "From VB IDE"
         End
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddTags 
         Caption         =   "Add ASM tags"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmTemplates"
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

'TEMPLATES
'---------
'templates are stored in "add-in directory\templates"
'so only create in add-in directory directory templates
'and save there these files - empty, module.asm, class.bas

'TODO - extracting code from VB IDE

'20.8. 2004 - initial version - created GUI
'21.8. 2004 - patching code
'22.8. 2004 - added menu
'06.9. 2004 - new menu, fixing code
'17.9. 2004 - now code uses VB IO file statements instead of FSO - functions in modPublic.bas, now using MsgBoxX
'           - code improved

Private Const MSG_TITLE As String = "ASM Templates"
Private Const EMPTY_TEMP As String = "empty"
Private Const temp As String = ".asm"
Private Const ASM_TAG As String = "'#asm'"

'directory where templates are stored
Private Const directory As String = "templates\"

Private Enum EXTRACT
    code
    Description
End Enum

'--------------
'--- Events ---
'--------------

Private Sub Form_Initialize()
    LogMsg "Loading " & Add34(Me.caption) & " window", Me.name, "Form_Initialize", True, True
End Sub

Private Sub Form_Terminate()
    LogMsg "Unloading " & Add34(Me.caption) & " window", Me.name, "Form_Terminate", True, True
End Sub

Private Sub Form_Load()
    
    With lvwTemp
    
        'settings
        .view = lvwReport
        .HideSelection = False
        .LabelEdit = lvwAutomatic
        
        'add columns
        .ColumnHeaders.Add 1, , "Template name", .Width * 1 / 3
        .ColumnHeaders.Add 2, , "Template description", .Width * 2 / 3 - 60
        
    End With
    
    'load templates to listview
    If DoInit(False) = False Then Call cmdClose_Click

End Sub

'change template name
Private Sub lvwTemp_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    If RenameTemp(Trim(NewString)) = True Then
        'reload listview
        Call DoInit
    Else
        Cancel = True
    End If
    
End Sub

'edit template
Private Sub lvwTemp_DblClick()
Dim sCode As String, sDesc As String

    'check selected item
    If lvwTemp.SelectedItem Is Nothing Then Exit Sub

    sCode = ExtractFromTemplate(GetFullPath, code, True, True)
    If Len(sCode) = 0 Then Exit Sub
    
    sDesc = ExtractFromTemplate(GetFullPath, Description, True)
    If Len(sDesc) = 0 Then Exit Sub
       
    sCode = frmViewer.ShowViewer("Template editor", sCode, False, False)
    
    If Len(sCode) = 0 Then Exit Sub
    MsgBox GetFullPath
    SaveFile GetFullPath, sDesc & vbCrLf & sCode

End Sub

'delete template
Private Sub lvwTemp_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        Call DeleteTemp
        Call DoInit
    End If

End Sub

'change description
Private Sub lvwTemp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        Call ChangeDesc
        Call DoInit
    End If

End Sub

'----------------
'--- menu NEW ---
'----------------

'create new empty template
Private Sub mnuEmpty_Click()
Dim sName As String

    'new template name
    sName = Trim(InputBox("Enter template name", MSG_TITLE))
    If Len(sName) = 0 Then
        MsgBoxX "New template name is zero-length.", MSG_TITLE
        Exit Sub
    End If
        
    If FileExist(GetDir & sName & temp) = True Then
        MsgBoxX "Template name has been in the list yet.", MSG_TITLE
        Exit Sub
    End If
    
On Error Resume Next
    
    'rename template
    FileCopy GetDir & EMPTY_TEMP, GetDir & sName & temp
    
    'check error
    Call TestFileErrors(Err.Number)
    If Err.Number <> 0 Then Exit Sub
    
On Error GoTo 0
        
    'reload list-view
    Call DoInit

End Sub

'extract asm code from cursor location (from VB IDE)
Private Sub mnuFromVB_Click()
Dim sCode As String, sName As String, sDesc As String
    
    MsgBox "You know , it hard to get all this to work.."
    'PATCH-RAZIEL
    'cursor is in some location in VB IDE
    'suppouse it is in asm block
    'e.g.
    
    '...VB statement
    '#asm' aaa
    '#asm' bbb   --cursor is here
    '#asm' ccc
    '...VB statement
    
    'I want to extract asm code (without asm prefix)
    'aaa
    'bbb
    'ccc
    
    'and store it in variable sCode
    'so sCode will contain "aaa" & vbCrLf & "bbb" & vbCrLf & "ccc"
    
    'if it is cursor not in ASM line/block then sCode will be empty - ""
       
    'check code
    If Len(sCode) = 0 Then
        MsgBoxX "Cursor is not in ASM block.", MSG_TITLE
        Exit Sub
    End If
    
    'get template name
    sName = Trim(InputBox("Enter template name", MSG_TITLE))
    If Len(sName) = 0 Then
        MsgBoxX "Template name is zero-length.", MSG_TITLE
        Exit Sub
    End If
    
    If FileExist(GetDir & sName & temp) = True Then
        MsgBoxX "Template name has been in the list yet.", MSG_TITLE
        Exit Sub
    End If
    
    'get description
    sDesc = Trim(InputBox("Write template description.", MSG_TITLE))
    If Len(sDesc) = 0 Then
        MsgBoxX "Template description is zero-length.", MSG_TITLE
        Exit Sub
    End If
    
On Error Resume Next
    
    SaveFile GetDir & sName & temp, sDesc & vbCrLf & sCode
    
    Call TestFileErrors(Err.Number)
    If Err.Number <> 0 Then Exit Sub
    
On Error GoTo 0
    
    Call DoInit
    
End Sub

'create new template based on another (existing) template
Private Sub mnuBasedOn_Click()
Dim sName As String

    'some existing template should be selected
    If lvwTemp.SelectedItem Is Nothing Then
        MsgBoxX "Select existing template that will be base for new template.", MSG_TITLE
        Exit Sub
    End If
    
    'get new template name
    sName = Trim(InputBox("Enter template name", MSG_TITLE))
    If Len(sName) = 0 Then
        MsgBoxX "New template name is zero-length.", MSG_TITLE
        Exit Sub
    End If
    
    If FileExist(GetDir & sName & temp) = True Then
        MsgBoxX "Template name has been in the list yet.", MSG_TITLE
        Exit Sub
    End If
        
On Error Resume Next
    
    'rename template
    FileCopy GetFullPath, GetDir & sName & temp
    
    Call TestFileErrors(Err.Number)
    If Err.Number <> 0 Then Exit Sub
    
On Error GoTo 0
       
    'reload list-view
    Call DoInit

End Sub

'---------------------------
'--- other items in menu ---
'---------------------------

Private Sub mnuAddTags_Click()
    mnuAddTags.Checked = Not mnuAddTags.Checked
End Sub

Private Sub mnuHelp_Click()
Dim s As String
    s = "To rename template - double-click (slowly) on template name" & vbCrLf
    s = "To edit template - double-click (fast) on template name" & vbCrLf
    s = s & "To delete template - use " & Add34("Delete") & " key" & vbCrLf
    s = s & "To change template description - use right mouse button"
    
    MsgBoxX s, MSG_TITLE
    
End Sub

'-----------------------------
'--- CommandButtons events ---
'-----------------------------

'close Button
Private Sub cmdClose_Click()
    Unload Me
End Sub

'save asm code from template to clipboard
Private Sub cmdToClipboard_Click()
Dim sCode As String

    If lvwTemp.SelectedItem Is Nothing Then Exit Sub

    'get ASM code
    sCode = ExtractFromTemplate(GetFullPath, code, True)
    If Len(sCode) = 0 Then Exit Sub

    LogMsg "Pasting code of template to clipboard", Me.name, "cmdToClipboard_Click", True, True

    'save it
    Clipboard.Clear
    Clipboard.SetText sCode

End Sub

'preview code
Private Sub cmdPreview_Click()
Dim sCode As String

    If lvwTemp.SelectedItem Is Nothing Then Exit Sub

    'get ASM code
    sCode = ExtractFromTemplate(GetFullPath, code, True)
    If Len(sCode) = 0 Then Exit Sub

    LogMsg "Preview of template", Me.name, "cmdPreview_Click", True, True
    'view it
    frmViewer.ShowViewer "Template", sCode, True

End Sub

'paste code to cursor location
Private Sub cmdToCursor_Click()
Dim sCode As String

    If lvwTemp.SelectedItem Is Nothing Then Exit Sub
    
    'get ASM code
    sCode = ExtractFromTemplate(GetFullPath, code, True)
    If Len(sCode) = 0 Then Exit Sub
    
    LogMsg "Pasting code of template to cursor location", Me.name, "cmdToCursor_Click", True, True
    PutStringToCurCursor sCode
    
End Sub

'------------------------
'--- Helper functions ---
'------------------------

'fill ListView with Templates
'parameter - bWarning - true -> if Description does not exits show warning otherwise not
'return - true/false

Private Function DoInit(Optional bWarning As Boolean = False) As Long
Dim sDesc As String, sFileName As String

    'at first check folder
    If DirExist(GetDir) = False Then
        MsgBox "Add-In template directory does not exist ?!?!", vbCritical, MSG_TITLE
        DoInit = False
        Exit Function
    End If
   
    With lvwTemp
    
        'cler list-view
        .ListItems.Clear
    
        sFileName = Dir(GetDir & "*.*", vbNormal)
        Do While sFileName <> ""
            
            'check file type
            If StrComp(Right(sFileName, Len(temp)), temp, vbTextCompare) <> 0 Then GoTo NextFile
            
            'add file-name to the textbox
            .ListItems.Add , , Left(sFileName, Len(sFileName) - Len(temp))
            
            'extract info from template
            sDesc = ExtractFromTemplate(GetDir & sFileName, Description, bWarning)
            
            If Len(sDesc) = 0 Then
                .ListItems.item(.ListItems.count).ListSubItems.Add , , "<no description>"
            Else
                .ListItems.item(.ListItems.count).ListSubItems.Add , , sDesc
            End If

NextFile:
            sFileName = Dir
            
        Loop
    
    End With
    
    DoInit = True

End Function

'return full path to file that is selected in listview
Private Function GetFullPath() As String
    If lvwTemp.SelectedItem Is Nothing Then Exit Function
    GetFullPath = GetDir & lvwTemp.SelectedItem.Text & temp
End Function

'return path "add-in directory\templates"
Private Function GetDir() As String
    GetDir = Get_Paths(AddIn_Directory) & directory
End Function

'extract asm code/template description from template
'parameters - sTempPath - full path to template
'           - eData - code/description
'bWarning   - show warning dialogs

Private Function ExtractFromTemplate(sTempPath As String, eData As EXTRACT, Optional bWarning As Boolean = False, Optional bDoNotAddASMTags As Boolean = False) As String
Dim sData As String, lFirstLine As Long, d As Long, sLine As String
    
    LogMsg "Extracting " & IIf(eData = code, "code", "description") & " from template file " & sTempPath, Me.name, "ExtractFromTemplate", True, True
    
On Error Resume Next
    
    'try to open file
    d = FreeFile
    Open sTempPath For Input As d
    
    If Err.Number <> 0 Then
        MsgBox "Error during openning template " & sTempPath, vbExclamation, MSG_TITLE
        Exit Function
    End If

On Error GoTo 0

    'check end of file
    If EOF(d) = True Then
        If bWarning = True Then MsgBoxX "Template is empty.", MSG_TITLE
        Exit Function
    End If
    
    'read description
    Line Input #d, ExtractFromTemplate
    
    If eData = Description Then
        '---description ---
        Close d
        Exit Function
    Else
        '--- code ---
    
        'read whole file
        While EOF(d) = False
            Line Input #d, sLine
            sData = sData & IIf(Len(sData) = 0, "", vbCrLf) & sLine
        Wend
               
        Close d
        
        'get asm code
        If Len(sData) = 0 Then
            If bWarning = True Then MsgBoxX "Template is empty.", MSG_TITLE
            Exit Function
        End If
            
        If mnuAddTags.Checked = True And bDoNotAddASMTags = False Then
            'add ASM tags
            ExtractFromTemplate = ASM_TAG & " " & sData
            ExtractFromTemplate = Replace(ExtractFromTemplate, vbCrLf, vbCrLf & ASM_TAG & " ")
        Else
            ExtractFromTemplate = sData
        End If
            
    End If
    
End Function

'change template name
Private Function RenameTemp(ByVal sNewName As String) As Boolean

    LogMsg "Renaming template", Me.name, "RenameTemp", True, True

    RenameTemp = False
    
    If FileExist(GetDir & sNewName & ".asm") = True Then
        MsgBoxX "Template name has been in the list yet.", MSG_TITLE
        Exit Function
    End If

On Error Resume Next
    
    'try to rename file
    Name GetFullPath As GetDir & sNewName & ".asm"
            
    'invalid file name
    If Err.Number = 5 Then
        MsgBoxX "Invalid Template name.", MSG_TITLE
    
    'other error
    ElseIf Err.Number <> 0 Then
        MsgBoxX "Error during renaming template (file)", MSG_TITLE
    End If
    
On Error GoTo 0
    
    RenameTemp = True

End Function

'change template description
Private Sub ChangeDesc()
Dim sDesc As String, d As Long

    LogMsg "Changind description of template", Me.name, "ChangeDesc", True, True

    'check selected item
    If lvwTemp.SelectedItem Is Nothing Then Exit Sub
    
    'get new description
    sDesc = Trim(InputBox("Write new template description.", MSG_TITLE, lvwTemp.SelectedItem.ListSubItems(1).Text))
    If Len(sDesc) = 0 Then
        MsgBoxX "New template description is zero-length.", MSG_TITLE
        Exit Sub
    End If

On Error Resume Next

    d = FreeFile

    Open GetFullPath For Output As d

        If Err.Number <> 0 Then
            MsgBox "Error during opening file " & GetFullPath, vbExclamation, MSG_TITLE
            Exit Sub
        End If
    
On Error GoTo 0
    
        'write new description
        Print #d, sDesc
    Close d
    
End Sub

'delete Template
Private Sub DeleteTemp()
    
    LogMsg "Deleting template", Me.name, "DeleteTemp", True, True
    
    'check selected item
    If lvwTemp.SelectedItem Is Nothing Then Exit Sub

    'question
    If MsgBox("Are you shure to delete selected template?", vbYesNo + vbQuestion, MSG_TITLE) = vbNo Then Exit Sub

    'try to delete file
On Error Resume Next
    kill2 GetFullPath
    
End Sub

'check err object
Private Sub TestFileErrors(ByVal lNumber As Long)

    If lNumber = 5 Then
        MsgBoxX "Invalid template name.", MSG_TITLE
    ElseIf lNumber <> 0 Then
        MsgBox "Error during creating new template.", vbExclamation, MSG_TITLE
    End If

End Sub

