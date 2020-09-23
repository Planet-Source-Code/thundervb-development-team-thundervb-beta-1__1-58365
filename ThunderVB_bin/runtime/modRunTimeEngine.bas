Option Explicit

Public Enum eVB
    #ENUM#
End Enum

Private Const DLL As String = "msvbvm60.dll"
Private Const NUMBER as Long = #NUMBER#

Private apVB(1 To NUMBER) As Long
Private apIAT(1 To NUMBER) As Long
Private apOld(1 To NUMBER) As Long

'-----------
'--- API ---
'-----------

Private Declare Function lstrlen Lib "kernel32.dll" (ByVal pString As Long) As Long

Private Declare Function VirtualQuery Lib "kernel32.dll" (ByVal lpAddress As Long, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Private Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type


Private Declare Function VirtualProtect Lib "kernel32.dll" (ByVal lBaseAdresa As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const PAGE_READWRITE As Long = &H4

Private Const IMAGE_DOS_SIGNATURE = &H5A4D
Private Const IMAGE_NT_SIGNATURE = &H4550

Private Type IMAGEDATADIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Private Type IMAGE_DOS_HEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long 'raw adresa NT hlavicky
End Type

Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Reserved1 As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    SubSystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(1 To 16) As IMAGEDATADIRECTORY
End Type

Private Type IMAGE_NT_HEADERS
    Signature As Long
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

Private Type IMAGE_IMPORT_DESCRIPTOR
    OriginalFirstThunk As Long
    TimeDateStamp As Long
    ForwarderChain As Long
    Name As Long
    FirstThunk As Long
End Type

Private Type IMAGE_THUNK_DATA32
    FunctionOrOrdinalOrAddress As Long
End Type

Private Const IMAGE_DIRECTORY_ENTRY_IMPORT = 1
Private Const IMAGE_ORDINAL_FLAG32 As Long = &H80000000

Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

'---------------------
'--- IAT FUNCTIONS ---
'---------------------

Private Function EnumIAT(ByVal hBase As Long, sDllName As String, ByVal sApiName As String) As Long
Dim tDosHeader As IMAGE_DOS_HEADER, tNTHeader As IMAGE_NT_HEADERS
Dim tImpDesc As IMAGE_IMPORT_DESCRIPTOR, tAddrThunk As IMAGE_THUNK_DATA32
Dim Pointer As Long, hLib As Long, pProc As Long

    EnumIAT = 0
    
    'base address of dll
    hLib = GetModuleHandle(sDllName)
    If hLib = 0 Then Exit Function
    
    'proc. address
    pProc = GetProcAddress(hLib, sApiName)
    If pProc = 0 Then Exit Function
    
    'copy DOS header
    CopyMemory VarPtr(tDosHeader), hBase, Len(tDosHeader)
    If tDosHeader.e_magic <> IMAGE_DOS_SIGNATURE Then Exit Function
    
    'copy NT header
    CopyMemory VarPtr(tNTHeader), hBase + tDosHeader.e_lfanew, Len(tNTHeader)
    If tNTHeader.Signature <> IMAGE_NT_SIGNATURE Then Exit Function
    
    'RVA of import descriptor
    Pointer = tNTHeader.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_IMPORT + 1).VirtualAddress
    If Pointer = 0 Then Exit Function
    
    Do While True
    
        'copy import descriptor
        CopyMemory VarPtr(tImpDesc), Pointer + hBase, Len(tImpDesc)
        
        'check last import descriptor
        If tImpDesc.Name = 0 Then Exit Function
        
        'check name of imported DLL
        If StrComp(PointerToStringA(hBase + tImpDesc.Name), sDllName, vbTextCompare) = 0 Then GoTo 10
    
        'go to the next image import. desc
        Pointer = Pointer + Len(tImpDesc)
    
    Loop
    
10:
    
    'VA of API
    Pointer = tImpDesc.FirstThunk
    If Pointer = 0 Then Exit Function
    
    Do While True
    
        'get address thunk
        CopyMemory VarPtr(tAddrThunk), hBase + Pointer, Len(tAddrThunk)
        
        'check last API
        If tAddrThunk.FunctionOrOrdinalOrAddress = 0 Then Exit Function
        
        If tAddrThunk.FunctionOrOrdinalOrAddress = pProc Then
            EnumIAT = Pointer + hBase
            Exit Function
        End If

        Pointer = Pointer + Len(tAddrThunk)
        
    Loop
    
End Function

Private Function OverWriteIAT(ByVal VA As Long, ByVal lNewValue As Long) As Long
Dim tMem As MEMORY_BASIC_INFORMATION, lOldProtect As Long, lVP As Long
    
    OverWriteIAT = 0

    'get info about memory
    VirtualQuery VA, tMem, Len(tMem)
    If tMem.BaseAddress = 0 Or tMem.RegionSize = 0 Then Exit Function
    
    'set new memory protection
    lVP = VirtualProtect(tMem.BaseAddress, tMem.RegionSize, PAGE_EXECUTE_READWRITE, lOldProtect)
    If lVP = 0 Then OverWriteIAT = 0

    'save pointer to API
    CopyMemory VarPtr(OverWriteIAT), VA, Len(VA)
    'set new value
    CopyMemory VA, VarPtr(lNewValue), Len(lNewValue)
    
    Call VirtualProtect(tMem.BaseAddress, tMem.RegionSize, lOldProtect, lVP)
    
End Function

'------------------------
'--- HELPER FUNCTIONS ---
'------------------------

Private Function PointerToStringA(ByVal lpStringA As Long) As String
Dim bBuffer() As Byte, lDelka As Long
   
   'check pointer
    If lpStringA <> 0 Then
      
        'lenght of string
        lDelka = lstrlen(lpStringA)
        If lDelka <> 0 Then
            
            'create buffer
            ReDim bBuffer(0 To (lDelka - 1)) As Byte
            
            'copy buffer
            CopyMemory VarPtr(bBuffer(0)), ByVal lpStringA, lDelka
            
            'to unicode
            PointerToStringA = StrConv(bBuffer, vbUnicode)
            Exit Function
            
        End If
    End If
    
    'error
    PointerToStringA = ""

End Function

Private Function Enum2String(eProcedure As eVB, Optional rtc As Boolean = True) As String
    If rtc = True Then
        #ENUM2RT#
    Else
        #ENUM2VB#
    End If
End Function

Private Function GetPointer(ByVal pFunction) As Long
    GetPointer = pFunction
End Function

'----------------------
'--- PUBLIC MEMBERS ---
'----------------------

Public Function IsHooked(eProcedure As eVB) As Boolean
    If apOld(eProcedure) = 0 Then IsHooked = False Else IsHooked = True
End Function

Public Function InitRunTimeEngine() As Boolean
Dim i As Long
    
    #INIT#
    For i = LBound(apVB) To UBound(apVB)
        If apVB(i) = 0 Or apIAT(i) = 0 Then
            InitRunTimeEngine = False
            Exit Function
        End If
    Next i

    InitRunTimeEngine = True

End Function

Public Sub HookAll()
Dim i As Long

    For i = LBound(apVB) To UBound(apVB)
        Hook i
    Next i

End Sub

Public Sub UnHookAll()
Dim i As Long

    For i = LBound(apVB) To UBound(apVB)
        UnHook i
    Next i

End Sub

Public Sub Hook(eProcedure As eVB)
    If IsHooked(eProcedure) = False Then apOld(eProcedure) = OverWriteIAT(apIAT(eProcedure), apVB(eProcedure))
End Sub

Public Sub UnHook(eProcedure As eVB)
    If IsHooked(eProcedure) = True Then
        Call OverWriteIAT(apIAT(eProcedure), apOld(eProcedure))
        apOld(eProcedure) = 0
    End If
End Sub

'-------
' HOOKS
'-------