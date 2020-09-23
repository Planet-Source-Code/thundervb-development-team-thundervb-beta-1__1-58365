Attribute VB_Name = "modPublic"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

'this module will contain public functions for general use

'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Libor
'Module created , initial version
'
'21/8/2004[dd/mm/yyyy] : Code Edited by Libor
'added new Settings for StdCall DLL, new Save* and Read* functions: Libor
'NOTE : SaveVBP and ReadVBP needs patching
'
'26/8/2004[dd/mm/yyy]  : Code Edited by Raziel
'Added WarnBox,ErrorBox and MsgBoxX
'
'27/8/2004[dd/mm/yyyy]
'Save* and Read* functions moved to frmSettings : Libor
'
'26/8/2004[dd/mm/yyy]  : Code Edited by Raziel
'Added LogMsg
'
'7/9/2004[dd/mm/yyyy]
'Added new settings - ASM/C
'New function LoadSettings    :Libor
'
'9/9/2004[dd/mm/yyy]  : Code Edited by Raziel
'Added Putstringtocurpos
'fixed ProcStrings , now it works as it should
'
'10/9/2004[dd/mm/yyyy] :Code edited by Raziel
'Code to convert #asm_start .. #asm_end to '#asm' lines..
'
'12/9/2004[dd/mm/yyyy]
'changed settings in StdCall Tab : Libor
'
'13/9/2004[dd/mm/yyyy]
'added new option to Get_Packer, new function CrLf : Libor
'
'16/9/2004[dd/mm/yyyy]
'added "Show packer output", "Add to menu" option, better CrLf function : Libor
'added function DirExist, LoadFile function patched
'
'19/9/2004[dd/mm/yyyy]
'patched function LogMsg function, all constants are private : Libor
'
'22/9/2004[dd/mm/yyyy] : Raziel
'Minor Changes on the logging code
'
'
'1/10/2004[dd/mm/yyyy] : Raziel
'GetFunctionCode ,EnumFunctionNames,EnumModuleNames
'SetFunctionLine ,SetCurLine
'6/10/2004[dd/mm/yyyy] : Raziel
'Sligth mods for GPF handling..
'Name
'7/10/2004[dd/mm/yyyy] : Raziel
'added save projects
'
'10/10/2004 [dd/mm/yyyy] : Raziel
'added procStringUnderAll , GetThunVBVer and ReplWSwithSpace

Public Const APP_NAME As String = "ThunderVB"
Public Const APP_NAMEs As String = "ThunVB"    ' s = short version

'Other
Public vb_Dll_version As Long '5/6

'*** General ***
'form Settings/Tab General

Private bGeneral_LoadOnStart As Boolean
Private bGeneral_UnLoadNotEXE As Boolean
Private bGeneral_PopUpExportWindow As Boolean
Private bGeneral_SetTopMost As Boolean
Private bGeneral_ListForAllModules As Boolean
Private bGeneral_HookCompiler As Boolean
Private bGeneral_SaveOBJ As Boolean
Private bGeneral_AddTLBIfNeeded As Boolean
Private bGeneral_HideErrorDialogs As Boolean
Private bGeneral_AddToMenu As Boolean
Private bGeneral_LogThunVBTools As Boolean

Public Enum GENERAL_
    LoadOnStartUp
    UnLoadIfProjectIsNotStandartExe
    PopUpExportsWindow
    SetTopMost
    ListingsForAllModules
    HookCompiler
    SaveObjFiles
    AddTlbToReferencesIfNeeded
    HideErrorDialogs
    AddToMenu
    LogThunVBTools
End Enum

'*** PATHS ***
'form Settings/Tab Paths

'->will set Addin
Private sPaths_AddIn As String        'path to addin directory
Private sPaths_Debug  As String       'path to debug directory
Private sPaths_Project As String      'project directory

'->will be set in Settings form
Private sPaths_Midl  As String        'path to midl.exe
Private sPaths_ML  As String          'path to ml.exe
Private sPaths_Packer As String       'path to packer
Private sPaths_TextEditor As String   'path to text-editor
Private sPaths_INCFiles As String     'path to .INC files (directory)
Private sPaths_LIBFiles As String     'path to .LIB files (directory)
Private sPaths_CCompiler As String    'path to C Compiler

Public Enum PATHS_
    AddIn_Directory
    Debug_Directory
    Midl
    ml
    Packer
    Project_Directory
    TextEditor
    INCFiles_Directory
    LIBFiles_Directory
    CCompiler
End Enum

'*** Packer ***
'form Settings/Tab Packer

Private bPacker_UsePacker As Boolean    'use packer
Private bPacker_ShowPackerOutPut As Boolean    'packer output
Private sPacker_CommandLine As String   'packer command line
Private sPacker_CmdLineDesc As String   'description of cmdline

Public Enum PACKER_
    UsePacker
    CommandLine
    CmdLineDescription
    ShowPackerOutPut
End Enum

'*** Debug ***
'form Settings/Tab Debug

Private bDebug_EnableDebugMsgBox As Boolean
Private bDebug_EnableOutPutToDebugLog As Boolean
Private bDebug_DelDebugLogBeforeCompiling As Boolean
Private bDebug_OutPutAssemblerMessToLog As Boolean
Private bDebug_OutPutMapFiles As Boolean
Private bDebug_ForceLog As Boolean
Private bDebug_DeleteLST As Boolean
Private bDebug_DeleteASM As Boolean

Public Enum DEBUG_
    EnableDebugMsgBox
    EnableOutPutToDebugLog
    DeleteDebugLogBeforeCompiling
    OutPutAssemblerMessagesToLog
    OutPutMapFiles
    ForceLog
    DeleteLST
    DeleteASM
End Enum

'*** Compile ***
'form Settings/Tab Compile

Private bCompile_PauseBeforeAsm As Boolean
Private bCompile_PauseBeforeLink As Boolean
Private bCompile_ModifyCmdLine As Boolean
Private bCompile_SkipLinking As Boolean

Public Enum COMPILE_
    PauseBeforeAssembly
    PauseBeforeLinking
    ModifyCmdLine
    SkipLinking
End Enum

'*** StdCallDLL ***
'form Settings/Tab StdCall DLL

Private bDLL_LinkAsDLL As Boolean
Private lBaseAddress As Long
Private bExportSymbols As Boolean
Private sExportedSymbols As String
Private sEntryPointName As String
Private bDLL_UsePreLoader As Boolean
Private bDLL_DebugPreLoader As Boolean
Private bDLL_FullLoading As Boolean

Public Enum DLL_
    LinkAsDll
    BaseAddress
    ExportSymbols
    ExportedSymbols
    EntryPointName
    UsePreLoader
    DebugPreLoader
    FullLoading
End Enum

'*** ASM ***
'form Settings/Tab ASM

Private bASM_UseASMColoring As Boolean  'do ASM coloring
Private sASM_ASMColors As String        'colors
Private bASM_FixASMListings As Boolean
Private bASM_CompileASMCode As Boolean

Public Enum ASM_
    UseASMColoring
    ASMColors
    FixASMListings
    CompileASMCode
End Enum

'*** C ***
'form Settings/Tab C

Private bC_UseCColoring As Boolean    'do C coloring
Private sC_CColors As String          'colors
Private bC_CompileCCode As Boolean

Public Enum C_
    UseCColoring
    CColors
    CompileCCode
End Enum

'Code Coloring
Public Type ColorInfo_entry
    str As String
    Color As Long
End Type

Public Type ColorInfo_list
    ColorInfo() As ColorInfo_entry
    count As Long
End Type

'Loging
Dim log As file_b
'misc things
Dim in_asm_block As Boolean 'if we are in asm code block , used when expanding #asm_start blocks
Dim in_C_block As Boolean 'if we are in c code block , used when expanding #c_start blocks
'---------------
'--- GENERAL ---
'---------------

'return General settings
'parameter - eGeneral - General flag
'          - True/False - flag

Public Function Get_General(ByVal eGeneral As GENERAL_) As Boolean
   
    Select Case eGeneral
        Case GENERAL_.AddTlbToReferencesIfNeeded
            Get_General = bGeneral_AddTLBIfNeeded
        Case GENERAL_.HideErrorDialogs
            Get_General = bGeneral_HideErrorDialogs
        Case GENERAL_.HookCompiler
            Get_General = bGeneral_HookCompiler
        Case GENERAL_.ListingsForAllModules
            Get_General = bGeneral_ListForAllModules
        Case GENERAL_.LoadOnStartUp
            Get_General = bGeneral_LoadOnStart
        Case GENERAL_.PopUpExportsWindow
            Get_General = bGeneral_PopUpExportWindow
        Case GENERAL_.SaveObjFiles
            Get_General = bGeneral_SaveOBJ
        Case GENERAL_.SetTopMost
            Get_General = bGeneral_SetTopMost
        Case GENERAL_.UnLoadIfProjectIsNotStandartExe
            Get_General = bGeneral_UnLoadNotEXE
        Case GENERAL_.AddToMenu
            Get_General = bGeneral_AddToMenu
        Case GENERAL_.LogThunVBTools
            Get_General = bGeneral_LogThunVBTools
    End Select

End Function

'change General flag
'parameters - eGeneral - general flag to change
'           - bNewValue - new flag

Public Sub Let_General(ByVal eGeneral As GENERAL_, bNewValue As Boolean)
   
    Select Case eGeneral
        Case GENERAL_.AddTlbToReferencesIfNeeded
            bGeneral_AddTLBIfNeeded = bNewValue
        Case GENERAL_.HideErrorDialogs
            bGeneral_HideErrorDialogs = bNewValue
        Case GENERAL_.HookCompiler
            bGeneral_HookCompiler = bNewValue
        Case GENERAL_.ListingsForAllModules
            bGeneral_ListForAllModules = bNewValue
        Case GENERAL_.LoadOnStartUp
            bGeneral_LoadOnStart = bNewValue
        Case GENERAL_.PopUpExportsWindow
            bGeneral_PopUpExportWindow = bNewValue
        Case GENERAL_.SaveObjFiles
            bGeneral_SaveOBJ = bNewValue
        Case GENERAL_.SetTopMost
            bGeneral_SetTopMost = bNewValue
        Case GENERAL_.UnLoadIfProjectIsNotStandartExe
            bGeneral_UnLoadNotEXE = bNewValue
        Case GENERAL_.AddToMenu
            bGeneral_AddToMenu = bNewValue
        Case GENERAL_.LogThunVBTools
            bGeneral_LogThunVBTools = bNewValue
    End Select

End Sub

'-------------
'--- PATHS ---
'-------------

'return paths
'parameter - ePath - special path
'          - bWarning - when path is not specified, alert will appear
'return - ""     - path is not set
'       - string - path

Public Function Get_Paths(ByVal ePath As PATHS_, Optional bWarning As Boolean = False) As String
Dim sText As String
   
    Select Case ePath
        Case PATHS_.AddIn_Directory
            Get_Paths = sPaths_AddIn
            sText = "Path to Add-in directory is not set."
        Case PATHS_.Debug_Directory
            Get_Paths = sPaths_Debug
            sText = "Path to Debug directory is not set."
        Case PATHS_.Packer
            Get_Paths = sPaths_Packer
            sText = "Path to Packer is not set." & vbCrLf & "Setting/Paths"
        Case PATHS_.Project_Directory
            Get_Paths = sPaths_Project
            sText = "Path to this Project directory is not set."
        Case PATHS_.Midl
            Get_Paths = sPaths_Midl
            sText = "Path to MIDL (midl.exe) is not set." & vbCrLf & "Setting/Paths"
        Case PATHS_.ml
            Get_Paths = sPaths_ML
            sText = "Path to ML (ml.exe) is not set." & vbCrLf & "Setting/Paths"
        Case PATHS_.TextEditor
            Get_Paths = sPaths_TextEditor
            sText = "Path to your Text-Editor is not set." & vbCrLf & "Setting/Paths"
        Case PATHS_.INCFiles_Directory
            Get_Paths = sPaths_INCFiles
            sText = "Path to .INC files is not set." & vbCrLf & "Setting/Paths"
        Case PATHS_.LIBFiles_Directory
            Get_Paths = sPaths_LIBFiles
            sText = "Path to .LIB files is not set." & vbCrLf & "Setting/Paths"
        Case PATHS_.CCompiler
            Get_Paths = sPaths_CCompiler
            sText = "Path to C compiler is not set." & vbCrLf & "Setting/Paths"
    End Select

    'check path
    If Len(Get_Paths) = 0 And bWarning = True Then
        MsgBox sText, vbExclamation, APP_NAME
    End If

End Function

'change Paths
'parameters - ePath - path
'           - sNewValue - new path

Public Sub Let_Paths(ByVal ePath As PATHS_, sNewValue As String)
  
    Select Case ePath
        Case PATHS_.AddIn_Directory
            sPaths_AddIn = sNewValue
        Case PATHS_.Debug_Directory
            sPaths_Debug = sNewValue
        Case PATHS_.Packer
            sPaths_Packer = sNewValue
        Case PATHS_.Project_Directory
            sPaths_Project = sNewValue
        Case PATHS_.Midl
            sPaths_Midl = sNewValue
        Case PATHS_.ml
            sPaths_ML = sNewValue
        Case PATHS_.TextEditor
            sPaths_TextEditor = sNewValue
        Case PATHS_.CCompiler
            sPaths_CCompiler = sNewValue
        Case PATHS_.INCFiles_Directory
            sPaths_INCFiles = sNewValue
        Case PATHS_.LIBFiles_Directory
            sPaths_LIBFiles = sNewValue
    End Select
    
End Sub

'--------------
'--- PACKER ---
'--------------

'get Packer settings
'parameter - ePacker - packer setting
'return -  "False"/"True" or string

Public Function Get_Packer(ByVal ePacker As PACKER_) As String

    Select Case ePacker
        Case PACKER_.UsePacker
            Get_Packer = bPacker_UsePacker
        Case PACKER_.ShowPackerOutPut
            Get_Packer = bPacker_ShowPackerOutPut
        Case PACKER_.CommandLine
            Get_Packer = sPacker_CommandLine
        Case PACKER_.CmdLineDescription
            Get_Packer = sPacker_CmdLineDesc
    End Select
    
End Function

'change Packer flags
'parameters - ePacker - flags
'           - sNewValue - new flag/path

Public Sub Let_Packer(ByVal ePacker As PACKER_, sNewValue As String)

    Select Case ePacker
        Case PACKER_.UsePacker
            bPacker_UsePacker = CBool(sNewValue)
        Case PACKER_.ShowPackerOutPut
            bPacker_ShowPackerOutPut = CBool(sNewValue)
        Case PACKER_.CommandLine
            sPacker_CommandLine = sNewValue
        Case PACKER_.CmdLineDescription
            sPacker_CmdLineDesc = sNewValue
    End Select
    
End Sub

'-----------
'--- ASM ---
'-----------

'get ASM settings
'parameter - eASM - asm setting
'return -  string/true/false

Public Function Get_ASM(ByVal eASM As ASM_) As String

    Select Case eASM
        Case ASM_.ASMColors
            Get_ASM = sASM_ASMColors
        Case ASM_.CompileASMCode
            Get_ASM = bASM_CompileASMCode
        Case ASM_.FixASMListings
            Get_ASM = bASM_FixASMListings
        Case ASM_.UseASMColoring
            Get_ASM = bASM_UseASMColoring
    End Select
    
End Function

'change ASM flags
'parameters - eASM - flags
'           - sNewValue - new setting

Public Sub Let_ASM(ByVal eASM As ASM_, sNewValue As String)

    Select Case eASM
        Case ASM_.ASMColors
            sASM_ASMColors = sNewValue
        Case ASM_.CompileASMCode
            bASM_CompileASMCode = CBool(sNewValue)
        Case ASM_.FixASMListings
            bASM_FixASMListings = CBool(sNewValue)
        Case ASM_.UseASMColoring
            bASM_UseASMColoring = CBool(sNewValue)
    End Select
    
End Sub

'---------
'--- C ---
'---------

'get C settings
'parameter - eC - C setting
'return -  string/true/false

Public Function Get_C(ByVal eC As C_) As String

    Select Case eC
        Case C_.CColors
            Get_C = sC_CColors
        Case C_.CompileCCode
            Get_C = bC_CompileCCode
        Case C_.UseCColoring
            Get_C = bC_UseCColoring
    End Select
    
End Function

'change Code coloring flags
'parameters - eC - flags
'           - sNewValue - new value

Public Sub Let_C(ByVal eC As C_, sNewValue As String)

    Select Case eC
        Case C_.CColors
            sC_CColors = sNewValue
        Case C_.CompileCCode
            bC_CompileCCode = CBool(sNewValue)
        Case C_.UseCColoring
            bC_UseCColoring = CBool(sNewValue)
    End Select
    
End Sub

'-------------
'--- DEBUG ---
'-------------

'get Debug flag
'parameter - eDebug_ - debug flag
'return - TRUE/FALSE

Public Function Get_Debug(ByVal eDebug As DEBUG_) As Boolean

    Select Case eDebug
        Case DEBUG_.DeleteASM
            Get_Debug = bDebug_DeleteASM
        Case DEBUG_.DeleteDebugLogBeforeCompiling
            Get_Debug = bDebug_DelDebugLogBeforeCompiling
        Case DEBUG_.DeleteLST
            Get_Debug = bDebug_DeleteLST
        Case DEBUG_.EnableDebugMsgBox
            Get_Debug = bDebug_EnableDebugMsgBox
        Case DEBUG_.EnableOutPutToDebugLog
            Get_Debug = bDebug_EnableOutPutToDebugLog
        Case DEBUG_.ForceLog
            Get_Debug = bDebug_ForceLog
        Case DEBUG_.OutPutAssemblerMessagesToLog
            Get_Debug = bDebug_OutPutAssemblerMessToLog
        Case DEBUG_.OutPutMapFiles
            Get_Debug = bDebug_OutPutMapFiles
    End Select
        
End Function


'change Debug flag
'parameters - eDebug - flag
'           - bNewValue - new value

Public Sub Let_Debug(ByVal eDebug As DEBUG_, bNewValue As Boolean)

    Select Case eDebug
        Case DEBUG_.DeleteASM
            bDebug_DeleteASM = bNewValue
        Case DEBUG_.DeleteDebugLogBeforeCompiling
            bDebug_DelDebugLogBeforeCompiling = bNewValue
        Case DEBUG_.DeleteLST
            bDebug_DeleteLST = bNewValue
        Case DEBUG_.EnableDebugMsgBox
            bDebug_EnableDebugMsgBox = bNewValue
        Case DEBUG_.EnableOutPutToDebugLog
            bDebug_EnableOutPutToDebugLog = bNewValue
        Case DEBUG_.ForceLog
            bDebug_ForceLog = bNewValue
        Case DEBUG_.OutPutAssemblerMessagesToLog
            bDebug_OutPutAssemblerMessToLog = bNewValue
        Case DEBUG_.OutPutMapFiles
            bDebug_OutPutMapFiles = bNewValue
    End Select
        
End Sub

'---------------
'--- COMPILE ---
'---------------

'get Compile flag
'parameter - eCompile - compile flag
'return - TRUE/FALSE

Public Function Get_Compile(ByVal eCompile As COMPILE_) As Boolean

    Select Case eCompile
        Case COMPILE_.ModifyCmdLine
            Get_Compile = bCompile_ModifyCmdLine
        Case COMPILE_.PauseBeforeAssembly
            Get_Compile = bCompile_PauseBeforeAsm
        Case COMPILE_.PauseBeforeLinking
            Get_Compile = bCompile_PauseBeforeLink
        Case COMPILE_.SkipLinking
            Get_Compile = bCompile_SkipLinking
    End Select
    
End Function

'change Compile flag
'parameters - eCompile  - flag
'           - bNewValue - new flag

Public Sub Let_Compile(ByVal eCompile As COMPILE_, bNewValue As Boolean)

    Select Case eCompile
        Case COMPILE_.ModifyCmdLine
            bCompile_ModifyCmdLine = bNewValue
        Case COMPILE_.PauseBeforeAssembly
            bCompile_PauseBeforeAsm = bNewValue
        Case COMPILE_.PauseBeforeLinking
            bCompile_PauseBeforeLink = bNewValue
        Case COMPILE_.SkipLinking
            bCompile_SkipLinking = bNewValue
    End Select
    
End Sub

'-------------------
'--- STDCALL DLL ---
'-------------------

'get StdCall DLL settings
'parameter - eDLL - DLL setting
'return -  False/True, long, string

Public Function Get_DLL(ByVal eDLL As DLL_) As String

    Select Case eDLL
        Case DLL_.BaseAddress
            Get_DLL = lBaseAddress
        Case DLL_.ExportedSymbols
            Get_DLL = sExportedSymbols
        Case DLL_.ExportSymbols
            Get_DLL = bExportSymbols
        Case DLL_.LinkAsDll
            Get_DLL = bDLL_LinkAsDLL
        Case DLL_.EntryPointName
            Get_DLL = sEntryPointName
        Case DLL_.DebugPreLoader
            Get_DLL = bDLL_DebugPreLoader
        Case DLL_.FullLoading
            Get_DLL = bDLL_FullLoading
        Case DLL_.UsePreLoader
            Get_DLL = bDLL_UsePreLoader
    End Select

End Function

'change StdCall flags
'parameters - eDLL - flag
'           - bNewValue - new flag

Public Sub Let_DLL(ByVal eDLL As DLL_, sNewValue As String)

    Select Case eDLL
        Case DLL_.BaseAddress
            lBaseAddress = CLng(sNewValue)
        Case DLL_.ExportedSymbols
            sExportedSymbols = sNewValue
        Case DLL_.ExportSymbols
            bExportSymbols = CBool(sNewValue)
        Case DLL_.LinkAsDll
            bDLL_LinkAsDLL = CBool(sNewValue)
        Case DLL_.EntryPointName
            sEntryPointName = sNewValue
        Case DLL_.DebugPreLoader
            bDLL_DebugPreLoader = CBool(sNewValue)
        Case DLL_.FullLoading
            bDLL_FullLoading = CBool(sNewValue)
        Case DLL_.UsePreLoader
            bDLL_UsePreLoader = CBool(sNewValue)
    End Select

End Sub

'------------------------
'--- Helper Functions ---
'------------------------

'load settings from registry and VBP
Public Sub LoadSettings()

    'load and kill form
    Load frmSettings
    Unload frmSettings

End Sub


'get the fisrt word of a string , with the space folowing it
Public Function GetFirstWordWithSpace(str As String) As String
Dim temp As Long
    
    temp = InStrWS(1, str, " ", vbBinaryCompare)
    If temp Then
        GetFirstWordWithSpace = Left$(str, temp - 1)
    Else
        GetFirstWordWithSpace = str
    End If
    
End Function

'removes the first word from a string and the space after it
Public Sub RemFirstWordWithSpace(str As String) 'remove fisrt word
Dim temp As Long

    temp = InStrWS(1, str, " ", vbBinaryCompare)
    If temp Then
        str = Right$(str, Len(str) - temp + 1)
    Else
        str = ""
    End If
    
End Sub

'get the fisrt word of a string
Public Function GetFirstWord(str As String) As String
Dim temp As Long

    temp = InStr(1, str, " ")
    If temp Then
        GetFirstWord = Left$(str, temp - 1)
    Else
        GetFirstWord = str
    End If
    
End Function

'removes the first word from a string
Public Sub RemFisrtWord(str As String) 'remove fisrt word
Dim temp As Long

    temp = InStr(1, str, " ")
    If temp Then
        str = Right$(str, Len(str) - temp)
    Else
        str = ""
    End If
    
End Sub

Public Function GetAllWordsToArr(ByVal str As String) As String()
Dim toS As String_B, temp As Long, t2 As Long
    
    If Mid$(str, Len(str), 1) = " " Then
        t2 = 1
    End If
    
    str = Trim$(str)
    Do
        temp = InStr(1, str, " ")
        
        If temp > 0 Then
            AppendString toS, Left$(str, temp - 1)
            str = Trim$(Right$(str, Len(str) - temp))
        Else
            AppendString toS, str
            str = ""
        End If
    Loop While temp > 0
    
    If t2 = 1 Then
        AppendString toS, " "
        'frmDConsole.AppendLog str & "|"
    End If
    
    FinaliseString toS
    
    GetAllWordsToArr = toS.str
    
End Function
'loads a text file ands returns its contents as String
Public Function LoadFile(file As String, Optional bCheckFile As Boolean = True) As String
Dim ff As Long
    
    If bCheckFile = False Then GoTo 10
    
    If FileExist(file) = False Then
        ErrorBox "File dows not exist (" & file & ")", "modPublic", "LoadFile"
        Exit Function
    End If
    
10:
    
    ff = FreeFile
    Open file For Binary As ff
    LoadFile = Space(LOF(ff))
    Get ff, , LoadFile
    Close ff
    
End Function

'saves a text file
Public Sub SaveFile(file As String, data As String)
Dim ff As Long
    
    ff = FreeFile
    Open file For Output As ff
    Close ff
    Open file For Binary As ff
    Put ff, , data
    Close ff
    
End Sub
'checks if a file exists
Public Function FileExist(file As String) As Boolean

    If GetFileAttributes(file) <> -1 Then FileExist = True
    
End Function

'checks if a file exists
Public Function DirExist(directory As String) As Boolean

    If Dir(directory, vbDirectory) <> "" Then DirExist = True
    
End Function

'gets the filename from a full file path (eg "c:\windows\notepad.exe"->"notepad.exe")
Public Function GetFilename(filepath As String) As String

    GetFilename = Split(filepath, "\")(UBound(Split(filepath, "\")))
    
End Function

'gets the path from a full file path (eg "c:\windows\notepad.exe" ->"c:\windows\")
Public Function GetPath(filepath As String) As String

    GetPath = Replace(filepath, GetFilename(filepath), "")

End Function

'deletes a file if it exists..
Public Sub kill2(file As String)

    If FileExist(file) Then Kill file

End Sub

'adds "  on the start and end of a string if they do not exist
Public Function Add34(sText As String) As String
Dim temp As String

    Add34 = sText
    If Len(Add34) > 1 Then
        If Mid$(Add34, 1, 1) <> Chr$(34) Then Add34 = Chr$(34) & Add34
        If Mid$(Add34, Len(Add34), 1) <> Chr$(34) Then Add34 = Add34 & Chr$(34)
    End If
    
End Function
'removes "  on the start and end of a string if they do not exist
Public Function Rem34(sText As String) As String
Dim temp As String

    Rem34 = sText
    If Len(Rem34) > 1 Then
        If Mid$(Rem34, 1, 1) = Chr$(34) Then Rem34 = Left(Rem34, Len(Rem34))
        If Mid$(Rem34, Len(Rem34), 1) = Chr$(34) Then Rem34 = Right(Rem34, Len(Rem34))
    End If
    
End Function

'get the string between the find1 and find2 strings (without containing them)
Public Function getS(find1 As String, find2 As String, str As String, Optional start As Long = 1) As String
Dim i As Long, i2 As Long

    i = InStr(start, str, find1, vbTextCompare) + Len(find1)
    i2 = InStr(i, str, find2, vbTextCompare)
    If i2 > i Then
        getS = Mid$(str, i, i2 - i)
        start = i2
    Else
        start = 0
    End If
    
End Function

Private Function InStrWS(start As Long, string1 As String, string2 As String, cm As VbCompareMethod) As Long
Dim temp As Long, i As Long

    temp = InStr(start, string1, string2, cm)
    If temp Then
        Do
            If Mid$(string1, temp + i, 1) = " " Then
                i = i + 1
                If i > Len(string1) Then i = Len(string1): Exit Do
            Else
                Exit Do
            End If
        Loop
    temp = temp + i
    End If
    InStrWS = temp
    
End Function

'Added by Raziel , 25/8/2004
'Used to display messages to the user
'According to the pluing settings (Log msg's , hide them ect)

Public Function WarnBox(str As String, codeModule As String, codePosition As String, Optional style As VbMsgBoxStyle = vbOKOnly Or vbExclamation) As VbMsgBoxResult
    
    If Get_General(HideErrorDialogs) = False Then
        WarnBox = MsgBox("Warning : " & vbNewLine & str, style, codeModule & ":" & codePosition)
    End If
    
    LogMsg "(From " & codeModule & ":" & codePosition & ") " & str, "modPublic", "WarnBox"
    
End Function

Public Function ErrorBox(str As String, codeModule As String, codePosition As String, Optional style As VbMsgBoxStyle = vbOKOnly Or vbCritical) As VbMsgBoxResult
    
    If Get_General(HideErrorDialogs) = False Then
        ErrorBox = MsgBox("Error : " & vbNewLine & str, style, codeModule & ":" & codePosition)
    End If
    
    LogMsg "(From " & codeModule & ":" & codePosition & ") " & str, "modPublic", "ErrorBox"
    
End Function

Public Function MsgBoxX(str As String, Optional caption As String = APP_NAME, Optional style As VbMsgBoxStyle = vbOKOnly Or vbInformation) As VbMsgBoxResult

    MsgBoxX = MsgBox(str, style, caption)
    
End Function

'Changes all strings spaces in the s string to "_"
'eg string "this is 'a simple' "string cont"aining a string" is changed to
'          "this is 'a_simple' "string_cont"aining a string"
Function ProcStrings(s As String) As String
Dim temp As String, st As Long, stold As Long
    
    st = 1
    stold = 0
    Do While st > stold
        stold = st
        temp = Add34(getS(Chr$(34), Chr$(34), s, st))
        
        If st > stold And st > 0 Then
            s = Replace$(s, temp, Replace$(Replace$(Replace$(temp, " ", "_"), ";", "_"), "'", "_"))
            st = st + 1
        End If
    Loop

    st = 1
    stold = 0
    Do While st > stold
        stold = st
        temp = "'" & (getS("'", "'", s, st)) & "'"
        
        If st > stold And st > 0 Then
            s = Replace$(s, temp, Replace$(Replace$(temp, " ", "_"), ";", "_"))
            st = st + 1
        End If
    Loop
    
    ProcStrings = s
    
End Function
Function ProcStringsUnderAll(s As String) As String
Dim temp As String, st As Long, stold As Long
    
    st = 1
    stold = 0
    Do While st > stold
        stold = st
        temp = Add34(getS(Chr$(34), Chr$(34), s, st))
        
        If st > stold And st > 0 Then
            s = Replace$(s, temp, String$(Len(temp), "_"))
            st = st + 1
        End If
    Loop
    
    ProcStringsUnderAll = s
    
End Function
'Log format : [Time : ]\[codemodule:codePosition\] str
'upgrade - log ThunVB tools (like Code Generator, Code Wizard...) - set bThunVBToolMsg to True and set in General Tab option Log ThunVB tools : Libor
Public Sub LogMsg(str As String, codeModule As String, codePosition As String, Optional bLogTime As Boolean = True, Optional bThunVBToolMsg As Boolean = False)
Dim temp As String
    
    
    'If Get_Debug(EnableOutPutToDebugLog) = False Then Exit Sub
    
    'Libor - patch
    'If bThunVBToolMsg = True And Get_General(LogThunVBTools) = False Then Exit Sub
    
    'Get_Paths(Debug_Directory)
    If log.maxbuflen = 0 Then
        log = OpenFile(APP_NAME & "_" & ".txt", 512)
        Seek log.filenum, LOF(log.filenum) + 1
        AppendToFile log, "*******************************************************" & vbNewLine
    End If
    
    If bLogTime Then
        AppendToFile log, Time & " : "
    End If
    
    AppendToFile log, "[" & codeModule & ":" & codePosition & "]" & vbTab & str & vbNewLine
    FlushFile log
    
End Sub

Public Sub FlushLog()
    
    If log.maxbuflen > 0 Then
        CloseFile log
    End If
    
End Sub

Function PutStringToCurCursor(str As String) As Boolean
Dim curline As Long

    If VBI Is Nothing Then Exit Function
    If VBI.ActiveCodePane Is Nothing Then Exit Function
    If VBI.ActiveCodePane.codeModule Is Nothing Then Exit Function
    
    With VBI.ActiveCodePane
        .Window.SetFocus
        .GetSelection curline, 0, 0, 0
        .codeModule.InsertLines curline, str
    End With
    PutStringToCurCursor = True
    
End Function

Sub AsmBlocks_ConvertLine_init()
    
    in_asm_block = False
    
End Sub

Function AsmBlocks_ConvertLine(ByRef str As String) As String
Dim cmstr As String
    
    AsmBlocks_ConvertLine = str
    'if it is a coment then
    str = Trim(str)
    If Mid(str, 1, 1) <> "'" Then Exit Function
    'remove the coment and any space after it
    str = Right(str, Len(str) - 1)
    cmstr = Trim(Replace(str, vbTab, " "))
    'if asm_start/end command
    If InStr(1, cmstr, "#asm_start", vbTextCompare) = 1 Then AsmBlocks_ConvertLine = "'ASM START": in_asm_block = True: Exit Function
    If InStr(1, cmstr, "#asm_end", vbTextCompare) = 1 Then in_asm_block = False: AsmBlocks_ConvertLine = "'ASM END": Exit Function
    'if inside asm block
    If in_asm_block = False Then Exit Function
    'and if the line is not a comand (eg '#asm')
    If Mid(str, 1, 1) = "#" Then Exit Function
    'expand the block to many '#asm' lines
    'and write the output
    AsmBlocks_ConvertLine = vbTab & "'#asm' " & str

End Function

Function CBlocks_ConvertLine(ByRef str As String) As String
Dim cmstr As String
    
    CBlocks_ConvertLine = str
    'if it is a coment then
    str = Trim(str)
    'If Mid(str, 1, 1) <> "'" Then Exit Function
    If Len(str) = 0 Then Exit Function
    'remove the coment and any space after it
    'str = Right(str, Len(str) - 1)
    cmstr = Trim(Replace(str, vbTab, " "))
    'if c_start/end command
    If InStr(1, cmstr, "'#c_start", vbTextCompare) = 1 Then CBlocks_ConvertLine = "'c START": in_C_block = True: Exit Function
    If InStr(1, cmstr, "'#c_end", vbTextCompare) = 1 Then in_C_block = False: CBlocks_ConvertLine = "'c END": Exit Function
    'if inside c block
    If in_C_block = False Then Exit Function
    'and if the line is not a comand (eg '#c')
    If Mid(str, 1, 1) = "#" Then Exit Function
    'expand the block to many '#c' lines
    'and write the output
    CBlocks_ConvertLine = vbTab & "'#c' " & str

End Function

'make string that contains vbCrLf characters
'parameter - lCount - number of vbCrLfs
'return - string - string of vbCrLfs

Public Function CrLf(Optional ByVal lCount As Long = 1) As String
Dim i As Long
    For i = 1 To lCount
        CrLf = CrLf & vbCrLf
    Next i
End Function


Function GetFunctionCode(inMod As String, funct As String) As String
Dim pk As vbext_ProcKind, s As String
Dim pcount As Long, pline As Long, sLines As String
    Dim nams() As String, namsC As Long
    Dim objComponent As VBComponent
    Dim objMember As Member
    Dim strtemp As String
    Dim intTemp As Integer

    If VBI Is Nothing Then Exit Function
    If VBI.ActiveVBProject Is Nothing Then Exit Function
    If VBI.ActiveVBProject.VBComponents Is Nothing Then Exit Function
    
    For Each objComponent In VBI.ActiveVBProject.VBComponents
        If (objComponent.Type = vbext_ct_StdModule Or _
            objComponent.Type = vbext_ct_ClassModule Or _
            objComponent.Type = vbext_ct_VBForm) And objComponent.name = inMod Then
            For Each objMember In objComponent.codeModule.Members
                If objMember.Type = vbext_mt_Method And objMember.name = funct Then
                    With objComponent.codeModule
                        pline = .ProcBodyLine(funct, pk)
                        pcount = .ProcCountLines(funct, pk)
                        sLines = .lines(pline, pcount)
                    End With
                End If
            Next
        End If
    Next

    GetFunctionCode = sLines
    
End Function

Function EnumModuleNames() As String()
    Dim nams() As String, namsC As Long
    Dim objComponent As VBComponent
    Dim objMember As Member
    Dim strtemp As String
    Dim intTemp As Integer

    If VBI Is Nothing Then Exit Function
    If VBI.ActiveVBProject Is Nothing Then Exit Function
    If VBI.ActiveVBProject.VBComponents Is Nothing Then Exit Function
    
    'enumerate the procedures in every module file within
    'the current project
    For Each objComponent In VBI.ActiveVBProject.VBComponents
        If objComponent.Type = vbext_ct_StdModule Or _
           objComponent.Type = vbext_ct_ClassModule Or _
           objComponent.Type = vbext_ct_VBForm Then
           
           ReDim Preserve nams(namsC)
           nams(namsC) = objComponent.name
           namsC = namsC + 1
           
        End If
    Next
    
    EnumModuleNames = nams
    
End Function
    

Function EnumFunctionNames(inMod As String) As String()

    Dim nams() As String, namsC As Long
    Dim objComponent As VBComponent
    Dim objMember As Member
    Dim strtemp As String
    Dim intTemp As Integer

    If VBI Is Nothing Then Exit Function
    If VBI.ActiveVBProject Is Nothing Then Exit Function
    If VBI.ActiveVBProject.VBComponents Is Nothing Then Exit Function
    
    For Each objComponent In VBI.ActiveVBProject.VBComponents
        If (objComponent.Type = vbext_ct_StdModule Or _
            objComponent.Type = vbext_ct_ClassModule Or _
            objComponent.Type = vbext_ct_VBForm) And objComponent.name = inMod Then
            For Each objMember In objComponent.codeModule.Members
                If objMember.Type = vbext_mt_Method Then
                    ReDim Preserve nams(namsC)
                    nams(namsC) = objMember.name
                    namsC = namsC + 1
                End If
            Next
        End If
    Next
    
    EnumFunctionNames = nams

End Function

Sub SetCurLine(inMod As String, funct As String, NewTopLine As Long)
Dim pk As vbext_ProcKind, s As String
Dim pcount As Long, pline As Long, sLines As String
    Dim nams() As String, namsC As Long
    Dim objComponent As VBComponent
    Dim objMember As Member
    Dim strtemp As String
    Dim intTemp As Integer

    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveVBProject Is Nothing Then Exit Sub
    If VBI.ActiveVBProject.VBComponents Is Nothing Then Exit Sub
    
    For Each objComponent In VBI.ActiveVBProject.VBComponents
        If (objComponent.Type = vbext_ct_StdModule Or _
            objComponent.Type = vbext_ct_ClassModule Or _
            objComponent.Type = vbext_ct_VBForm) And objComponent.name = inMod Then
            For Each objMember In objComponent.codeModule.Members
                If objMember.Type = vbext_mt_Method And objMember.name = funct Then
                    With objComponent.codeModule
                        pline = .ProcBodyLine(funct, pk)
                        .CodePane.TopLine = pline + NewTopLine
                    End With
                End If
            Next
        End If
    Next

    
End Sub

Sub SetFunctionLine(inMod As String, funct As String, LineNum As Long, newLineString As Long, Optional bReplace As Boolean = True)
Dim pk As vbext_ProcKind, s As String
Dim pcount As Long, pline As Long, sLines As String
    Dim nams() As String, namsC As Long
    Dim objComponent As VBComponent
    Dim objMember As Member
    Dim strtemp As String
    Dim intTemp As Integer

    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveVBProject Is Nothing Then Exit Sub
    If VBI.ActiveVBProject.VBComponents Is Nothing Then Exit Sub
    
    For Each objComponent In VBI.ActiveVBProject.VBComponents
        If (objComponent.Type = vbext_ct_StdModule Or _
            objComponent.Type = vbext_ct_ClassModule Or _
            objComponent.Type = vbext_ct_VBForm) And objComponent.name = inMod Then
            For Each objMember In objComponent.codeModule.Members
                If objMember.Type = vbext_mt_Method And objMember.name = funct Then
                    With objComponent.codeModule
                        pline = .ProcBodyLine(funct, pk)
                        If bReplace Then
                            .ReplaceLine pline + LineNum, newLineString
                        Else
                            .InsertLines pline + LineNum, newLineString
                        End If
                    End With
                End If
            Next
        End If
    Next
    
End Sub

'Save all projects ...
Public Sub SaveProjects(binf As Boolean)

Dim i As Long, i2 As Long
    
    For i = 1 To VBI.VBProjects.count
    
        With VBI.VBProjects(i)
        
            .SaveAs .filename
            If binf Then
                MsgBoxX "Saving project:" & .name
            End If
            For i2 = 1 To .VBComponents.count
            
                If .VBComponents(i2).FileCount > 1 Then
                    Dim i3 As Long
                        For i3 = 1 To .VBComponents(i2).FileCount
                        
                            If MsgBoxX("Save " & Add34(.VBComponents(i2).name) & " as " & Add34(.VBComponents(i2).FileNames(i3)) & " ??", _
                                       , vbQuestion Or vbYesNo) = vbYes Then
                                .VBComponents(i2).SaveAs .VBComponents(i2).FileNames(i3)
                            Else
                                
                            End If
                            
                        Next i3
                    Else
                        .VBComponents(i2).SaveAs .VBComponents(i2).FileNames(1)
                End If
                
            Next i2
            If binf Then
                MsgBoxX "Project " & Add34(.name) & " was saved", , vbInformation Or vbOKOnly
            End If
        End With
        
    Next i
    
End Sub

Public Function GetThunVBVer() As String

    GetThunVBVer = " Version " & App.Major & "." & App.Minor & "." & App.Revision
    
End Function

'replace all chars that count as word seperators with " "
Public Function ReplWSwithSpace(strin As String) As String

    ReplWSwithSpace = Replace(strin, "//", "__@?comment?__")
    ReplWSwithSpace = Replace(Replace(Replace(Replace(ReplWSwithSpace, ",", " "), "(", " "), ")", " "), "=", " ")
    ReplWSwithSpace = Replace(Replace(Replace(Replace(ReplWSwithSpace, "+", " "), "-", " "), "*", " "), "/", " ")
    ReplWSwithSpace = Replace(Replace(Replace(Replace(ReplWSwithSpace, "[", " "), "]", " "), "{", " "), "}", " ")
    ReplWSwithSpace = Replace(Replace(Replace(Replace(ReplWSwithSpace, "^", " "), "|", " "), "<", " "), ">", " ")
    ReplWSwithSpace = Replace(Replace(Replace(ReplWSwithSpace, "\", " "), "__@?comment?__", "//"), vbTab, "    ")
    
End Function

Function GetCodeWinPar() As Long
Dim ret As Long
    
    On Error Resume Next

    If VBI.MainWindow Is Nothing Then
        ret = GetDesktopWindow
    Else
        
        If VBI.MainWindow.hWnd <> 0 Then
            ret = VBI.MainWindow.hWnd
        Else
            ret = GetDesktopWindow
        End If
    End If
    
    GetCodeWinPar = ret

End Function


