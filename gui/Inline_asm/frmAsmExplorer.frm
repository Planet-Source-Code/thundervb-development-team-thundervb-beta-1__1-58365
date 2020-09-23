VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmASMExplorer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASM Explorer"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbMod 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2415
   End
   Begin MSComctlLib.ListView lvwInfo 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cmbFilter 
      Height          =   315
      Left            =   5400
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.ComboBox cmbProc 
      Height          =   315
      Left            =   2760
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      Caption         =   "Module"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   525
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      Caption         =   "Filter"
      Height          =   195
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   330
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "Procedure"
      Height          =   195
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmASMExplorer"
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

'13.09. 2004 - initial version, GUI
'18.09. 2004 - better GUI, code for filter
'19.09. 2004 - better filter and code, filter for labels
'14.10. 2004 - implemented filtering code, added new filter "All ASM code"

'TODO - editing listview, use LogMsg

'beginning of the code
'*********************
Private Const F_LOCAL As String = "local"
Private Const F_RET As String = "ret"
Private Const F_INT As String = "int"
Private Const F_CALL As String = "call"

'middle of the code
'******************
Private Const F_EQU As String = " equ "
Private Const F_DD As String = " dd "
Private Const F_DW As String = " dw "
Private Const F_DB As String = " db "

'end of the code
'***************
Private Const F_LABEL = ":"

Private Const COMMENT As String = ";"

'filter type
Private Enum eFilterType
    LOCAL_ = 1
    RET_ = 2
    Label_ = 3
    EQU_ = 4
    DWORD_ = 5
    BYTE_ = 6
    WORD_ = 7
    CALL_ = 8
    INT_ = 9
    All_ = 10
End Enum

'place of filter
Private Enum eCode
    code_beginning
    code_middle
    code_end
    code_all
End Enum

Private Enum eGetAsmInfo
    ASM_Code
    ASM_Comment
End Enum

'structure for storing asm info about asm code
Private Type tInfoASM
    lLine As Long
    sASM As String
    sComment As String
End Type

Private Const ASM_PREFIX As String = "'#asm'"
Private bLoaded As Boolean

'--------------
'--- EVENTS ---
'--------------

Private Sub lvwInfo_DblClick()
    'check selected item
    If lvwInfo.SelectedItem Is Nothing Or cmbProc.ListIndex = -1 Or cmbMod.ListIndex = -1 Or cmbFilter.ListIndex = -1 Then Exit Sub
    'set current line
    SetCurLine cmbMod.Text, cmbProc.Text, lvwInfo.SelectedItem.Text
End Sub

Private Sub cmbFilter_Click()

    If bLoaded = False Or cmbMod.ListIndex = -1 Or cmbProc.ListIndex = -1 Or cmbFilter.ListIndex = -1 Then Exit Sub
    Call InitListView
    
    Select Case cmbFilter.ItemData(cmbFilter.ListIndex)
        Case eFilterType.LOCAL_
            Call AddToListView(Filter(LOCAL_, code_beginning))
        Case eFilterType.CALL_
            Call AddToListView(Filter(CALL_, code_beginning))
        Case eFilterType.RET_
            Call AddToListView(Filter(RET_, code_beginning))
        Case eFilterType.INT_
            Call AddToListView(Filter(INT_, code_beginning))
        Case eFilterType.EQU_
            Call AddToListView(Filter(EQU_, code_middle))
        Case eFilterType.BYTE_
            Call AddToListView(Filter(BYTE_, code_middle))
        Case eFilterType.DWORD_
            Call AddToListView(Filter(DWORD_, code_middle))
        Case eFilterType.WORD_
            Call AddToListView(Filter(WORD_, code_middle))
        Case eFilterType.Label_
            Call AddToListView(Filter(Label_, code_end))
        Case eFilterType.All_
            Call AddToListView(Filter(All_, code_all))
    End Select
    
End Sub

'combobox - with list of all Modules
Private Sub cmbMod_Click()
    Call InitProc
    Call cmbFilter_Click
End Sub

'combobox - with list of all Procedures
Private Sub cmbProc_Click()
    Call cmbFilter_Click
End Sub

'initialize
Private Sub Form_Initialize()
    LogMsg "Loading " & Add34(Me.caption) & " window", Me.name, "Form_Initialize", True, True
    Call InitAll
End Sub

Private Sub Form_Terminate()
    LogMsg "Unloading " & Add34(Me.caption) & " window", Me.name, "Form_Terminate", True, True
End Sub

Private Sub Form_Load()
    Call InitAll
End Sub

'------------------------
'--- HELPER FUNCTIONS ---
'------------------------

'extract info from line of code
'parameters - eInfo - type of info
'           - sAsmLine - line of code
'retunt - string

Private Function GetInfo(eInfo As eGetAsmInfo, sAsmLine As String) As String
Dim lComment As Long

    'look for comment
    lComment = InStr(sAsmLine, COMMENT)
    
    'choose type of info
    If eInfo = ASM_Code Then
        
        If lComment = 0 Then
            'no comment
            GetInfo = sAsmLine
        Else
            'extract line of code
            GetInfo = Left(sAsmLine, lComment - 1)
        End If
    
    Else

        If lComment = 0 Then
            'no comment
            GetInfo = ""
        Else
            'extract comment
            GetInfo = Right(sAsmLine, Len(sAsmLine) - lComment)
        End If
        
    End If

    GetInfo = Trim(GetInfo)

End Function

'add tInfoASM structure to listview
'parameters - tInfoASM() - array of tInfoASM structures

Private Sub AddToListView(tInfoASM() As tInfoASM)
Dim i As Long
    
    With lvwInfo
    
        'clear listview
        .ListItems.Clear
        
        'check array
On Error Resume Next
        i = LBound(tInfoASM)
        If Err.Number <> 0 Then Exit Sub
On Error GoTo 0
    
        'add items
        For i = LBound(tInfoASM) To UBound(tInfoASM)
            .ListItems.Add , , tInfoASM(i).lLine
            .ListItems.item(.ListItems.count).ListSubItems.Add , , tInfoASM(i).sASM
            .ListItems.item(.ListItems.count).ListSubItems.Add , , tInfoASM(i).sComment
        Next i

    End With

End Sub

'filter
'parameters - asLines() - array of lines of code
'           - eFilter - type of filter
'           - eType - type of code
'return - array of tInfoASM structures

Private Function Filter(eFilter As eFilterType, eType As eCode) As tInfoASM()
Dim i As Long, sLine As String, lCount As Long, sFind As String
Dim bContinue As Boolean, atASM() As tInfoASM
Dim asLines() As String

    'check items
    If cmbMod.ListIndex = -1 Or cmbProc.ListIndex = -1 Then Exit Function
    
    'get code of function and split it
    asLines = Split(GetFunctionCode(cmbMod.Text, cmbProc.Text), vbCrLf)
    'check the array
    On Error Resume Next
        i = LBound(asLines)
        If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    
    lCount = 0

    'choose filter text
    Select Case eFilter
        Case eFilterType.CALL_
            sFind = F_CALL
        Case eFilterType.INT_
            sFind = F_INT
        Case eFilterType.LOCAL_
            sFind = F_LOCAL
        Case eFilterType.RET_
            sFind = F_RET
        Case eFilterType.EQU_
            sFind = F_EQU
        Case eFilterType.BYTE_
            sFind = F_DB
        Case eFilterType.DWORD_
            sFind = F_DD
        Case eFilterType.WORD_
            sFind = F_DW
        Case eFilterType.Label_
            sFind = F_LABEL
        Case eFilterType.All_
            sFind = "*."
    End Select
    
    'check filter
    If sFind = "" Then Exit Function
    
    'enum all lines
    For i = LBound(asLines) To UBound(asLines)
    
        sLine = Trim(asLines(i))
        'is it line with ASM code?
        If StrComp(Left(sLine, Len(ASM_PREFIX)), ASM_PREFIX, vbTextCompare) <> 0 Then GoTo 10
            
        'trim ASM prefix
        sLine = Trim(Mid(sLine, Len(ASM_PREFIX) + 1))
        bContinue = False
        
        'type of code
        Select Case eType
            Case eCode.code_beginning
                If StrComp(Left(sLine, Len(sFind)), sFind, vbTextCompare) = 0 Then bContinue = True
            Case eCode.code_middle
                If InStr(1, sLine, sFind, vbTextCompare) > 0 Then bContinue = True
            Case eCode.code_end
                If StrComp(Right(GetInfo(ASM_Code, sLine), Len(sFind)), sFind, vbTextCompare) = 0 Then bContinue = True
            Case eCode.code_all
                bContinue = True
        End Select
            
        If bContinue = True Then

            'adjust array
            lCount = lCount + 1
            ReDim Preserve atASM(1 To lCount)
        
            'save info
            With atASM(lCount)
                .lLine = i                        'LBound of the array is 0 so add 1
                .sASM = GetInfo(ASM_Code, sLine)
                .sComment = GetInfo(ASM_Comment, sLine)
            End With
        
        End If

10:
    Next i
    
    'return
    Filter = atASM

End Function

'init listview
Private Sub InitListView()
    
    With lvwInfo
    
        'settings
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        .view = lvwReport
        .HideSelection = False
        .LabelEdit = lvwManual
        
        'add columns
        .ColumnHeaders.Add 1, , "line", .Width * 1 / 19
        .ColumnHeaders.Add 2, , "ASM code", .Width * 9 / 19
        .ColumnHeaders.Add 3, , "comment", .Width * 9 / 19 - 70
        
    End With
    
End Sub

'initialize filter (combobox)
Private Sub InitFilter()
    
    bLoaded = False
    
    'add filters
    With cmbFilter
    
        .Clear
    
        .AddItem Enum2String(LOCAL_)
        .ItemData(.NewIndex) = eFilterType.LOCAL_
        
        .AddItem Enum2String(INT_)
        .ItemData(.NewIndex) = eFilterType.INT_
        
        .AddItem Enum2String(RET_)
        .ItemData(.NewIndex) = eFilterType.RET_
        
        .AddItem Enum2String(Label_)
        .ItemData(.NewIndex) = eFilterType.Label_
        
        .AddItem Enum2String(EQU_)
        .ItemData(.NewIndex) = eFilterType.EQU_
        
        .AddItem Enum2String(DWORD_)
        .ItemData(.NewIndex) = eFilterType.DWORD_
        
        .AddItem Enum2String(BYTE_)
        .ItemData(.NewIndex) = eFilterType.BYTE_
        
        .AddItem Enum2String(BYTE_)
        .ItemData(.NewIndex) = eFilterType.WORD_
        
        .AddItem Enum2String(CALL_)
        .ItemData(.NewIndex) = eFilterType.CALL_
        
        .AddItem Enum2String(All_)
        .ItemData(.NewIndex) = eFilterType.All_
        
        .ListIndex = 0
        
    End With
    
    bLoaded = True
    
End Sub

'init name of modules (combobox)
Private Sub InitMod()
Dim i As Long, asMod() As String

    'init
    bLoaded = False
    cmbMod.Clear
    
    'get list of all modules
    asMod = EnumModuleNames()
    
    'check array
On Error Resume Next
    i = LBound(asMod)
    If Err.Number <> 0 Then Exit Sub
On Error GoTo 0

    'add names
    For i = LBound(asMod) To UBound(asMod)
        cmbMod.AddItem asMod(i)
    Next i

    'exit
    cmbMod.ListIndex = 0
    bLoaded = True

End Sub

'init name of procedures (combobox)
Private Sub InitProc()
Dim i As Long, asProc() As String
    
    'init
    bLoaded = False
    cmbProc.Clear
    
    'check combo and get list of all functions in module
    If cmbMod.ListIndex = -1 Then Exit Sub
    asProc = EnumFunctionNames(cmbMod.Text)
    
    'check array
On Error Resume Next
    i = LBound(asProc)
    If Err.Number <> 0 Then Exit Sub
On Error GoTo 0
    
    'add names
    For i = LBound(asProc) To UBound(asProc)
        cmbProc.AddItem asProc(i)
    Next i

    'exit
    cmbProc.ListIndex = 0
    bLoaded = True

End Sub

'convert enum to string
Private Function Enum2String(Filter As eFilterType) As String
    Enum2String = Choose(Filter, "Local variables (LOCAL)", "RET", "Labels", "Constants (EQU)", "Variables - DWORD", "Variables - BYTE", "Variables - WORD", "CALL", "Interrupts (INT)", "All ASM code")
End Function

Private Sub InitAll()

    Call InitMod
    Call InitProc
    Call InitListView
    Call InitFilter

End Sub
