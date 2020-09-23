VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRunTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASM - C RunTime"
   ClientHeight    =   3945
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToCursor 
      Caption         =   "Cursor"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdToClipboard 
      Caption         =   "Clipboard"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwRT 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Paste to"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   585
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "Select"
      Begin VB.Menu mnuAll 
         Caption         =   "All"
      End
      Begin VB.Menu mnuNo 
         Caption         =   "No"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnlyASM 
         Caption         =   "Only ASM"
      End
      Begin VB.Menu mnuOnlyC 
         Caption         =   "Only C"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuFileType 
         Caption         =   "File type"
         Begin VB.Menu mnuASM 
            Caption         =   "ASM"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuC 
            Caption         =   "C"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuGenerate 
         Caption         =   "Generate"
         Begin VB.Menu mnuRunTime 
            Caption         =   "RunTime engine"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuProcedures 
            Caption         =   "Procedures"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuScope 
         Caption         =   "Procedures scope"
         Begin VB.Menu mnuPublic 
            Caption         =   "Public"
            Checked         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmRunTime"
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

'ASM - C RunTime
'---------------

'08.10. 2004 - basic GUI and code
'09.10. 2004 - code improvement
'10.10. 2004 - supports even C language, better GUI
'16.10. 2004 - code improved


Private Const ENUM_ As String = "#ENUM#"        'members in enum
Private Const NUMBER_ As String = "#NUMBER#"    'dimension of the array of pointers
Private Const INIT_ As String = "#INIT#"        'place for init. code

Private Const VB_ As String = "#VBNAME#"        'VB name
Private Const RT_ As String = "#RTNAME#"        'RunTime name

'fill the array with pointers
Private Const INIT_POINTER As String = "apVB(eVB.#VBNAME#) = GetPointer(AddressOf #RTNAME#)"
Private Const INIT_IAT As String = "apIAT(eVB.#VBNAME#) = EnumIAT(App.hInstance, DLL, Enum2String(#VBNAME#))"

Private Const ENUM2VB_ As String = "#ENUM2VB#"  'convert enum to VB name
Private Const ENUM2RT_ As String = "#ENUM2RT#"  'convert enum to RunTime name

'convert enum to names
Private Const ENUM_VB As String = "Enum2String = Choose(eProcedure,#VBNAME#)"
Private Const ENUM_RT As String = "Enum2String = Choose(eProcedure,#RTNAME#)"

'unknow header
Private Const UNKNOWN_HEADER As String = "??? unknown header ???"

Private Const INCLUDE_ASM As String = "'#asm' include "  'include - asm
Private Const INCLUDE_C As String = "'#c' include "      'include - c

Private Enum eFileType
    asm = 1
    C = 2
End Enum

Private Const SEPARATOR_FILE As String = "_"

Private Const RUNTIME_FILE As String = "modRunTimeEngine.bas"  'name of module where runtime is stored
Private Const RUNTIME_DIR As String = "runtime"                'name of directory where asm/c files are stored
Private Const RUNTIME_HEADERS As String = "headers.txt"        'name of file where are stored headers of functions

'----------------------
'--- CONTROL EVENTS ---
'----------------------

Private Sub Form_Initialize()
    LogMsg "Loading " & Add34(Me.caption) & " window", Me.name, "Form_Initialize", True, True
End Sub

Private Sub Form_Terminate()
    LogMsg "Unloading " & Add34(Me.caption) & " window", Me.name, "Form_Terminate", True, True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim sCode As String
    
    'get all code
    sCode = GenerateAllCode
    If Len(sCode) = 0 Then Exit Sub
    
    'show it
    frmViewer.ShowViewer "RunTime engine", "->" & sCode & "<-"
    
End Sub

Private Sub Form_Load()
    'init list view
    Call InitListView
    
    'add to the listview C files
    If mnuC.Checked = True Then Call AddFilesToListView(C)
    'add to the listview ASM files
    If mnuASM.Checked = True Then Call AddFilesToListView(asm)
End Sub

'List View sorting
Private Sub lvwRT_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwRT
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

'------------
'--- MENU ---
'------------

'select all
Private Sub mnuAll_Click()
Dim i As Long
    For i = 1 To lvwRT.ListItems.count
        lvwRT.ListItems.item(i).Checked = True
    Next i
End Sub

'select no
Private Sub mnuNo_Click()
Dim i As Long
    For i = 1 To lvwRT.ListItems.count
        lvwRT.ListItems.item(i).Checked = False
    Next i
End Sub

'select only ASM files
Private Sub mnuOnlyASM_Click()
Dim i As Long
    With lvwRT.ListItems
        For i = 1 To .count
            .item(i).Checked = (.item(i).ListSubItems(2).Text = Enum2String(asm))
        Next i
    End With
End Sub

'select only C files
Private Sub mnuOnlyC_Click()
Dim i As Long
    With lvwRT.ListItems
        For i = 1 To .count
            .item(i).Checked = (.item(i).ListSubItems(2).Text = Enum2String(C))
        Next i
    End With
End Sub

'file filter - ASM
Private Sub mnuASM_Click()
    mnuASM.Checked = Not mnuASM.Checked
    'reload list view
    Call Form_Load
End Sub

'file filter - C
Private Sub mnuC_Click()
    mnuC.Checked = Not mnuC.Checked
    'reload list view
    Call Form_Load
End Sub

'add procedures
Private Sub mnuProcedures_Click()
    mnuProcedures.Checked = Not mnuProcedures.Checked
End Sub

'add runtime
Private Sub mnuRunTime_Click()
    mnuRunTime.Checked = Not mnuRunTime.Checked
End Sub

'scope of procedures
Private Sub mnuPublic_Click()
    mnuPublic.Checked = Not mnuPublic.Checked
End Sub

'------------------------
'--- HELPER FUNCTIONS ---
'------------------------

'fill list view
'parameters - eType - C or ASM files
Private Sub AddFilesToListView(eType As eFileType)
Dim sPath As String, sVB As String, sRT As String
Dim lPos As Long, sFileType As String

    'enum all files in the directory
    sFileType = "*." & Enum2String(eType)
    sPath = Dir(Get_Paths(AddIn_Directory) & RUNTIME_DIR & "\*." & Enum2String(eType), vbNormal)
    
    Do While Len(sPath) <> 0
        
        'look for separator
        lPos = InStr(1, sPath, SEPARATOR_FILE, vbTextCompare)
        If lPos <> 0 Then
        
            'extract VB and RT name from file name
            sVB = Left(sPath, lPos - 1)
            sRT = Mid(sPath, lPos + 1, Len(sPath) - lPos - Len(sFileType) + 1)
            
            With lvwRT
                
                'add VB name
                .ListItems.Add , , sVB
                'add RT name
                .ListItems.item(.ListItems.count).ListSubItems.Add , , sRT
                'add type of language (C or ASM)
                .ListItems.item(.ListItems.count).ListSubItems.Add , , Enum2String(eType)
                
                'in Tag save file name
                .ListItems.item(.ListItems.count).Tag = sPath
            
            End With
            
        End If
        
        sPath = Dir
        
    Loop

End Sub

'enum to string
Private Function Enum2String(eType As eFileType)
    Enum2String = Choose(eType, "ASM", "C")
End Function

'return declare of RunTime function
'parameters - sRTName - run time name of function
Private Function GetHeader(sRTName As String) As String
Dim asHeaders() As String, i As Long
    
    'load file
    asHeaders = Split(LoadFile(Get_Paths(AddIn_Directory) & RUNTIME_DIR & "\" & RUNTIME_HEADERS, True), vbCrLf)
    
On Error Resume Next
    i = LBound(asHeaders)
    If Err.Number <> 0 Then Exit Function
On Error GoTo 0
    
    'enum lines
    For i = LBound(asHeaders) To UBound(asHeaders)
        
        'try to find line where function is declared
        If InStr(1, asHeaders(i), sRTName, vbTextCompare) <> 0 Then
            
            'OK
            GetHeader = asHeaders(i)
            Exit Function
            
        End If
        
    Next i
    
End Function

'init. list view
Private Sub InitListView()

    With lvwRT
    
        'clear all
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        .view = lvwReport
        .LabelEdit = lvwManual
        .Checkboxes = True
        
        'add columns
        .ColumnHeaders.Add 1, , "VB name", .Width * 1 / 3
        .ColumnHeaders.Add 2, , "RunTime name", .Width * 1 / 3
        .ColumnHeaders.Add 3, , "Language", .Width * 1 / 3 - 60
        
    End With

End Sub

'create basic code for hooking engine
Private Function CreateBasicCode() As String
Dim sMod As String, i As Long, sEnum As String, lCount As Long, sInit As String
Dim sEnum2VB As String, sEnum2RT As String
    
    'load bas file
    sMod = LoadFile(Get_Paths(AddIn_Directory) & RUNTIME_DIR & "\" & RUNTIME_FILE, True)
    If Len(sMod) = 0 Then Exit Function
    
    lCount = 0
    
    With lvwRT.ListItems
        For i = 1 To .count
            If .item(i).Checked = True Then
                 
                'create enum
                lCount = lCount + 1
                sEnum = sEnum & vbCrLf & Space(4) & .item(i).Text & "_" & .item(i).ListSubItems(2).Text & " = " & lCount
                 
                'create init (fill array with pointers)
                sInit = sInit & vbCrLf & Space(4) & INIT_POINTER & vbCrLf & Space(4) & INIT_IAT & vbCrLf
                sInit = Replace(sInit, VB_, .item(i).Text & "_" & .item(i).ListSubItems(2).Text)
                sInit = Replace(sInit, RT_, .item(i).ListSubItems(1).Text)
                 
                'create enum2string function
                sEnum2VB = sEnum2VB & ", " & .item(i).Text & "_" & .item(i).ListSubItems(2).Text
                sEnum2RT = sEnum2RT & ", " & .item(i).ListSubItems(1).Text
                 
            End If
        Next i
    End With
    
    If lCount = 0 Then Exit Function
    
    'add enum
    sMod = Replace(sMod, ENUM_, Mid(sEnum, Len(vbCrLf & Space(4)) + 1), , 1, vbTextCompare)
    'add constant
    sMod = Replace(sMod, NUMBER_, lCount, , 1, vbTextCompare)
    'add init
    sMod = Replace(sMod, INIT_, Mid(sInit, Len(vbCrLf & Space(4)) + 1), , 1)
    
    'path enum to string line
    sEnum2VB = Replace(ENUM_VB, VB_, Mid(sEnum2VB, Len(" ,")), , 1)
    sEnum2RT = Replace(ENUM_RT, RT_, Mid(sEnum2RT, Len(" ,")), , 1)
    
    'add enum 2 string
    sMod = Replace(sMod, ENUM2VB_, sEnum2VB, , 1)
    sMod = Replace(sMod, ENUM2RT_, sEnum2RT, , 1)
    
    CreateBasicCode = sMod
    
End Function

'create procedures for hooking
Private Function CreateHookProcedures(bPublic As Boolean) As String
Dim i As Long, sCode As String, sHeader As String
Dim lPos As Long, sNewHeader As String, bSub As Boolean
Dim lCount As Long

    lCount = 0

    With lvwRT.ListItems
        For i = 1 To .count
            If .item(i).Checked = True Then
            
                'number of selected items
                lCount = lCount + 1
                
                'get declare of function
                sHeader = GetHeader(.item(i).ListSubItems(1).Text)
                
                If Len(sHeader) = 0 Then
                    'declare was not found - use "unknown" header
NOHEADER:
                    sNewHeader = UNKNOWN_HEADER & vbCrLf & IIf(.item(i).ListSubItems(2).Text = Enum2String(asm), INCLUDE_ASM, INCLUDE_C) & Get_Paths(AddIn_Directory) & RUNTIME_DIR & "\" & .item(i).Tag & vbCrLf & UNKNOWN_HEADER
                Else
                    
                    'add public/name
                    sNewHeader = IIf(bPublic = True, "Public", "Private") & " Function " & .item(i).ListSubItems(1).Text & "_" & .item(i).ListSubItems(2).Text
                    
                    'add parameters
                    lPos = InStr(1, sHeader, "(")
                    If lPos <> 0 Then sNewHeader = sNewHeader & Mid(sHeader, lPos) Else GoTo NOHEADER
                    
                    'add "include" line
                    sNewHeader = sNewHeader & vbCrLf & IIf(.item(i).ListSubItems(2).Text = Enum2String(asm), INCLUDE_ASM, INCLUDE_C) & Get_Paths(AddIn_Directory) & RUNTIME_DIR & "\" & .item(i).Tag

                    'add end function
                    sNewHeader = sNewHeader & vbCrLf & "End Function"
                    
                    'if it is Sub replace string "Function" with string "Sub"
                    If InStr(1, sHeader, " Sub ", vbTextCompare) <> 0 Then sNewHeader = Replace(sNewHeader, " Function", " Sub")
    
                End If
            
                sCode = sCode & vbCrLf & vbCrLf & sNewHeader
            
            End If
        Next i
    End With
    
    'check counter and trim vbcrlfs
    If lCount <> 0 Then CreateHookProcedures = Mid(sCode, Len(vbCrLf & vbCrLf) + 1)

End Function

'generate all (engine, procedures) code
Private Function GenerateAllCode() As String
    
    'generate engine code
    If mnuRunTime.Checked = True Then GenerateAllCode = CreateBasicCode
    'generate hooks
    If mnuProcedures.Checked = True Then GenerateAllCode = GenerateAllCode & vbCrLf & vbCrLf & CreateHookProcedures(mnuPublic.Checked)
    
    If GenerateAllCode = vbCrLf & vbCrLf Then GenerateAllCode = ""
    If Left(GenerateAllCode, Len(vbCrLf & vbCrLf)) = vbCrLf & vbCrLf Then GenerateAllCode = Mid(GenerateAllCode, Len(vbCrLf & vbCrLf) + 1)

End Function
