VERSION 5.00
Begin VB.UserControl ScintillaEdit 
   BackColor       =   &H0080FFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "ScintillaEdit.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ScintillaEdit.ctx":0342
End
Attribute VB_Name = "ScintillaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

' Using Form PlanetSourceCode :
' Self SubClasser by
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................

'==================================================================================================
'Subclasser declarations
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hWnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Other Decs
'======================

Dim SelVisible As Boolean
Const strDefText = "Scintilla Control Warper Made by drk||Raziel" & vbNewLine & _
                   "This is a VB6 User Control + WinApi" & vbNewLine & _
                   "The Control was made from the" & vbNewLine & _
                   "Documentation of the v1.6.1 Release" & vbNewLine & _
                   "For the latest version of Scintilla visit " & vbNewLine & _
                   "www.scintilla.org" & vbNewLine & _
                   "Please note that this control is" & vbNewLine & _
                   "available for download on www.pscode.com" & vbNewLine & _
                   "and that it has nothing to do" & vbNewLine & _
                   "with the author of Scintilla" & vbNewLine & _
                   "So don't BUG him for my bugz"
           
Dim WndPointer As Long
Dim FunPointer As Long
Dim sci As Long
Dim lppw As Long
Dim ctr_style As Long 'control's style..
Dim bDoNotInit As Boolean

Public Event SciNotify(param As Long)

'=======================
'Subclasser :
'=======================

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data
'Notes:
  'If you really know what you're doing, it's possible to change the values of the
  'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
  'values get passed to the default handler.. and optionaly, the 'after' callback
  
  Select Case uMsg

        ' ======================================================
        ' Hide the TextBox when it loses focus (its LostFocus event it not fired
        ' when losing focus to a window outside the app).
    
        Case WM_NOTIFY
            If lParam <> 0 Then RaiseEvent SciNotify(lParam)

    
    End Select
    
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And 2 Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, 2, .nAddrSub)
    End If
    If When And 1 Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, 1, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                          'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hWnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = A_SetWindowLong(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call CopyMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                                 'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
      End If
    End With
    
    i = i - 1                                                                           'Next element
  Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call A_SetWindowLong(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = 2 Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(A_GetModuleHandle(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

'=================
'Rest Control Code
'=================

Public Function Send_SCI_message(ByVal msg As Long, ByVal par1 As Long, ByVal par2 As Long) As Long
    
    If (FunPointer = 0) Or (WndPointer = 0) Then
        Send_SCI_message = A_SendMessage(sci, msg, par1, par2)
    Else
        Send_SCI_message = CallBP(WndPointer, msg, par1, par2, FunPointer)
    End If
    
End Function

Public Function Send_SCI_messageStr(ByVal msg As Long, ByVal par1 As Long, ByVal par2 As String) As Long
    
    If (FunPointer = 0) Or (WndPointer = 0) Then
        Send_SCI_messageStr = A_SendMessageStr(sci, msg, par1, par2)
    Else
        Dim temp() As Byte
        sciBstrToAnsi par2, temp
        Send_SCI_messageStr = CallBP(WndPointer, msg, par1, VarPtr(temp(0)), FunPointer)
    End If
    
End Function

Private Sub UserControl_Initialize()
    If A_LoadLibrary("SciLexer.DLL") = 0 Then GoTo ErrExit
    If bDoNotInit = False Then
        ctr_style = WS_EX_CLIENTEDGE
    End If
    sci = A_CreateWindowEx(BorderStyle, "Scintilla", _
                           UserControl.name, WS_CHILD Or WS_VISIBLE, 0, 0, 200, 200, _
                           UserControl.hWnd, 0, App.hInstance, 0)
    If sci = 0 Then GoTo ErrExit
    'DebugerRaise
    If CCode() Then
        FunPointer = Send_SCI_message(SCI_GETDIRECTFUNCTION, 0, 0)
        WndPointer = Send_SCI_message(SCI_GETDIRECTPOINTER, 0, 0)
    End If

    Subclass_Start UserControl.hWnd
    Subclass_AddMsg UserControl.hWnd, WM_NOTIFY
    If bDoNotInit = False Then
        Text = strDefText
        bDoNotInit = True
    End If
    '// now a wrapper to call Scintilla directly
    'sptr_t CallScintilla(unsigned int iMessage, uptr_t wParam, sptr_t lParam){
    '    return pSciMsg(pSciWndData, iMessage, wParam, lParam);
    '}
    Exit Sub
    
ErrExit:
    'Error Recovery
    
End Sub

Private Sub UserControl_Resize()
    SetWindowPos sci, 0, 0, 0, UserControl.Width / 15, _
                 UserControl.Height / 15, 0
End Sub

Private Sub UserControl_Terminate()
    Subclass_StopAll
    DestroyWindow sci
    FunPointer = 0
    WndPointer = 0
End Sub

Public Property Get BorderStyle() As Long
    BorderStyle = ctr_style
End Property

Public Property Let BorderStyle(newS As Long)
    Dim oldS As Long
    If ctr_style <> newS Then
        Dim tbp As New PropertyBag, tbp2 As New PropertyBag, t As RECT
        A_SendMessage UserControl.hWnd, WM_SETREDRAW, False, 0
        oldS = ctr_style
        UserControl_WriteProperties tbp2
        ctr_style = newS
        UserControl_WriteProperties tbp
        UserControl_Terminate
        UserControl_Initialize
        If sci = 0 Then
            ctr_style = oldS
            UserControl_Initialize
            Set tbp = tbp2
        End If
        A_SendMessage UserControl.hWnd, WM_SETREDRAW, True, 0
        UserControl_Resize
        UserControl_ReadProperties tbp
        t.Top = 0: t.Left = 0: t.Right = -UserControl.Width / 15: t.Bottom = UserControl.Height / 15
        
    End If
    
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'WriteToPropBag PropBag, HideSelection
    'WriteToPropBag PropBag, SelectionEnd
    'WriteToPropBag PropBag, SelectionStart
    'WriteToPropBag PropBag, Anchor
    'WriteToPropBag PropBag, CurrentPos
    'WriteToPropBag PropBag, UndoCollection
    'WriteToPropBag PropBag, ErrStatus
    'WriteToPropBag PropBag, EnableOverType
    'WriteToPropBag PropBag, TargetEnd
    'WriteToPropBag PropBag, TargetStart
    'WriteToPropBag PropBag, StyleBits
    'WriteToPropBag PropBag, ReadOlny
    'WriteToPropBag PropBag, Text
    'WriteToPropBag PropBag, SelectionMode
    
    If lppw = -1 Then
        'load defaults
        Text = strDefText
    End If
    
    lppw = 0
    WriteToPropBag PropBag, False
    WriteToPropBag PropBag, MOUSEDWELLTIME
    WriteToPropBag PropBag, MODEVENTMASK
    WriteToPropBag PropBag, EDGECOLOUR
    WriteToPropBag PropBag, EDGECOLUMN
    WriteToPropBag PropBag, edgeMode
    WriteToPropBag PropBag, ZOOM
    WriteToPropBag PropBag, wrapVisualFlagsLocation
    WriteToPropBag PropBag, LAYOUTCACHE
    WriteToPropBag PropBag, WRAPSTARTINDENT
    WriteToPropBag PropBag, wrapVisualFlags
    WriteToPropBag PropBag, wrapMode
    'WriteToPropBag PropBag, FOLDEXPANDED
    'WriteToPropBag PropBag, FOLDLEVEL
    'WriteToPropBag PropBag, DOCPOINTER
    WriteToPropBag PropBag, PRINTWRAPMODE
    WriteToPropBag PropBag, PRINTCOLOURMODE
    WriteToPropBag PropBag, PRINTMAGNIFICATION
    WriteToPropBag PropBag, HIGHLIGHTGUIDE
    WriteToPropBag PropBag, INDENTATIONGUIDES
    'WriteToPropBag PropBag, LINEINDENTATION
    WriteToPropBag PropBag, BACKSPACEUNINDENTS
    WriteToPropBag PropBag, tabIndents
    WriteToPropBag PropBag, indent
    WriteToPropBag PropBag, useTabs
    WriteToPropBag PropBag, TABWIDTH
    'WriteToPropBag PropBag, Focus
    WriteToPropBag PropBag, CodePage
    WriteToPropBag PropBag, TwoPhaseDraw
    WriteToPropBag PropBag, BufferEdDraw
    WriteToPropBag PropBag, UsePalette
    WriteToPropBag PropBag, MarginRight
    WriteToPropBag PropBag, MarginLeft
    'WriteToPropBag PropBag, MarginSensitive
    'WriteToPropBag PropBag, MarginMask
    'WriteToPropBag PropBag, MarginWidth
    'WriteToPropBag PropBag, MarginType
    WriteToPropBag PropBag, ControlCharSymbol
    WriteToPropBag PropBag, CaretWidth
    WriteToPropBag PropBag, CaretPeriod
    WriteToPropBag PropBag, CaretLineBack
    WriteToPropBag PropBag, CaretLineVisible
    WriteToPropBag PropBag, CaretFore
    'WriteToPropBag PropBag, LineState
    WriteToPropBag PropBag, ViewEOL
    WriteToPropBag PropBag, MouseDownCaptures
    WriteToPropBag PropBag, ViewWhiteSpace
    WriteToPropBag PropBag, EndAtLastLine
    WriteToPropBag PropBag, ScrollWidth
    WriteToPropBag PropBag, xOffset
    WriteToPropBag PropBag, VScrollBar
    WriteToPropBag PropBag, HScrollBar
    WriteToPropBag PropBag, HideSelection
    WriteToPropBag PropBag, SelectionEnd
    WriteToPropBag PropBag, SelectionStart
    WriteToPropBag PropBag, Anchor
    WriteToPropBag PropBag, CurrentPos
    WriteToPropBag PropBag, UndoCollection
    WriteToPropBag PropBag, ErrStatus
    WriteToPropBag PropBag, EnableOverType
    WriteToPropBag PropBag, TargetEnd
    WriteToPropBag PropBag, TargetStart
    WriteToPropBag PropBag, StyleBits
    WriteToPropBag PropBag, ReadOlny
    WriteToPropBag PropBag, Text
    WriteToPropBag PropBag, BorderStyle
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lppw = 0
    'HideSelection = ReadfromPropBag(PropBag, False)
    'SelectionEnd = ReadfromPropBag(PropBag, 0)
    'SelectionStart = ReadfromPropBag(PropBag, 0)
    'Anchor = ReadfromPropBag(PropBag, 0)
    'CurrentPos = ReadfromPropBag(PropBag, 0)
    'UndoCollection = ReadfromPropBag(PropBag, True)
    'ErrStatus = ReadfromPropBag(PropBag, 0)
    'EnableOverType = ReadfromPropBag(PropBag, False)
    'TargetEnd = ReadfromPropBag(PropBag, 0)
    'TargetStart = ReadfromPropBag(PropBag, 0)
    'StyleBits = ReadfromPropBag(PropBag, 5)
    'ReadOlny = ReadfromPropBag(PropBag, False)
    'Text = ReadfromPropBag(PropBag, strDefText)
    'SelectionMode = ReadfromPropBag(PropBag, sci_SelectionMode.Sel_Stream)
    
    'If ReadfromPropBag(PropBag, True) Then lppw = -1: UserControl_WriteProperties PropBag
    
    lppw = 1
    MOUSEDWELLTIME = ReadfromPropBag(PropBag, 10000000)
    MODEVENTMASK = ReadfromPropBag(PropBag, 3959)
    EDGECOLOUR = ReadfromPropBag(PropBag, 12632256)
    EDGECOLUMN = ReadfromPropBag(PropBag, 0)
    edgeMode = ReadfromPropBag(PropBag, 0)
    ZOOM = ReadfromPropBag(PropBag, 0)
    wrapVisualFlagsLocation = ReadfromPropBag(PropBag, 0)
    LAYOUTCACHE = ReadfromPropBag(PropBag, 1)
    WRAPSTARTINDENT = ReadfromPropBag(PropBag, 0)
    wrapVisualFlags = ReadfromPropBag(PropBag, 0)
    wrapMode = ReadfromPropBag(PropBag, 0)
    'FOLDEXPANDED( = ReadfromPropBag(PropBag)
    'FOLDLEVEL = ReadfromPropBag(PropBag)
    'DOCPOINTER = ReadfromPropBag(PropBag)
    PRINTWRAPMODE = val(ReadfromPropBag(PropBag, 0))
    PRINTCOLOURMODE = ReadfromPropBag(PropBag, 1)
    PRINTMAGNIFICATION = ReadfromPropBag(PropBag, 0)
    HIGHLIGHTGUIDE = ReadfromPropBag(PropBag, 0)
    INDENTATIONGUIDES = ReadfromPropBag(PropBag, False)
    'LINEINDENTATION = ReadfromPropBag(PropBag)
    BACKSPACEUNINDENTS = ReadfromPropBag(PropBag, False)
    tabIndents = ReadfromPropBag(PropBag, True)
    indent = ReadfromPropBag(PropBag, 0)
    useTabs = ReadfromPropBag(PropBag, True)
    TABWIDTH = ReadfromPropBag(PropBag, 8)
    'Focus = ReadfromPropBag(PropBag)
    CodePage = ReadfromPropBag(PropBag, 0)
    TwoPhaseDraw = ReadfromPropBag(PropBag, True)
    BufferEdDraw = ReadfromPropBag(PropBag, True)
    UsePalette = ReadfromPropBag(PropBag, False)
    MarginRight = ReadfromPropBag(PropBag, 1)
    MarginLeft = ReadfromPropBag(PropBag, 1)
    'MarginSensitive = ReadfromPropBag(PropBag)
'    MarginMask = ReadfromPropBag(PropBag)
'    MarginWidth = ReadfromPropBag(PropBag)
'    MarginType = ReadfromPropBag(PropBag)
    ControlCharSymbol = ReadfromPropBag(PropBag, 0)
    CaretWidth = ReadfromPropBag(PropBag, 1)
    CaretPeriod = ReadfromPropBag(PropBag, 500)
    CaretLineBack = ReadfromPropBag(PropBag, 65535)
    CaretLineVisible = ReadfromPropBag(PropBag, False)
    CaretFore = ReadfromPropBag(PropBag, 0)
'    LineState = ReadfromPropBag(PropBag)
    ViewEOL = ReadfromPropBag(PropBag, False)
    MouseDownCaptures = ReadfromPropBag(PropBag, True)
    ViewWhiteSpace = ReadfromPropBag(PropBag, 0)
    EndAtLastLine = ReadfromPropBag(PropBag, True)
    ScrollWidth = ReadfromPropBag(PropBag, 200)
    xOffset = ReadfromPropBag(PropBag, 0)
    VScrollBar = ReadfromPropBag(PropBag, True)
    HScrollBar = ReadfromPropBag(PropBag, True)
    HideSelection = ReadfromPropBag(PropBag, False)
    SelectionEnd = ReadfromPropBag(PropBag, 0)
    SelectionStart = ReadfromPropBag(PropBag, 0)
    Anchor = ReadfromPropBag(PropBag, 0)
    CurrentPos = ReadfromPropBag(PropBag, 0)
    UndoCollection = ReadfromPropBag(PropBag, True)
    ErrStatus = ReadfromPropBag(PropBag, 0)
    EnableOverType = ReadfromPropBag(PropBag, False)
    TargetEnd = ReadfromPropBag(PropBag, 0)
    TargetStart = ReadfromPropBag(PropBag, 0)
    StyleBits = ReadfromPropBag(PropBag, 5)
    ReadOlny = ReadfromPropBag(PropBag, False)
    Text = ReadfromPropBag(PropBag, strDefText)
    BorderStyle = ReadfromPropBag(PropBag, 512)
End Sub


'SCI_GETTEXT(int length, char *text)
Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = "MainSettings"
  Dim temp As Long, tstr() As Byte

    temp = Send_SCI_message(SCI_GETTEXT, 0, 0)
    If temp < 2 Then Text = vbNullString: Exit Property
    ReDim tstr(temp - 1)
    Send_SCI_message SCI_GETTEXT, temp, VarPtr(tstr(0))
    ReDim Preserve tstr(temp - 2)
    Text = StrConv(tstr, vbUnicode)
  
End Property

'SCI_SETTEXT(<unused>, const char *text)
Public Property Let Text(newText As String)
  
  Send_SCI_messageStr SCI_SETTEXT, 0, newText
  
End Property

'SCI_SETSAVEPOINT
Public Function SetSavePoint()
  
  Send_SCI_message SCI_SETSAVEPOINT, 0, 0
  
End Function

'SCI_GETLINE(int line, char *text)
Public Function GetLine(numLine As Long) As String
    Dim tstr() As Byte, ln As Long
    
        ln = GetLineLen(numLine)
        If ln < 1 Then GetLine = vbNullString: Exit Function
        
        ReDim Preserve tstr(ln - 1)
        Send_SCI_message SCI_GETLINE, numLine, VarPtr(tstr(0))
        GetLine = StrConv(tstr, vbUnicode)
  
End Function


'SCI_REPLACESEL(<unused>, const char *text)
Public Sub SetSelectedText(tstr As String)

    Send_SCI_messageStr SCI_REPLACESEL, 0, tstr

End Sub

'SCI_SETREADONLY(bool readOnly)
Public Property Let ReadOlny(setRO As Boolean)
    
    Send_SCI_message SCI_SETREADONLY, setRO And 1, 0
    
End Property

'SCI_GETREADONLY
Public Property Get ReadOlny() As Boolean
Attribute ReadOlny.VB_ProcData.VB_Invoke_Property = "MainSettings"
    
    ReadOlny = Send_SCI_message(SCI_GETREADONLY, 0, 0)
    
End Property

'SCI_GETTEXTRANGE(<unused>, TextRange *tr)
Public Function GetTextRange(FromChar As Long, ToChar As Long) As String
    Dim temp As sci_TextRange
    
    If (ToChar - FromChar) < 1 Then Exit Function
    
    temp.chrg.cpMin = FromChar
    temp.chrg.cpMax = ToChar
    temp.lpstrText = StrConv(Space(ToChar - FromChar + 1), vbFromUnicode)
    
    Send_SCI_message SCI_GETTEXTRANGE, 0, VarPtr(temp)
    
    GetTextRange = Left(StrConv(temp.lpstrText, vbUnicode), ToChar - FromChar)
    
End Function

'SCI_GETSTYLEDTEXT(<unused>, TextRange *tr)
Public Function GetStyledTextRange(FromChar As Long, ToChar As Long) As Byte()
    Dim temp As sci_TextRange_Arr, tat() As Byte
    
    If (ToChar - FromChar) < 1 Then Exit Function
    
    temp.chrg.cpMin = FromChar
    temp.chrg.cpMax = ToChar
    ReDim tat((ToChar - FromChar) * 2 + 1)
    temp.lpstrText = VarPtr(tat(0))
    
    Send_SCI_message SCI_GETSTYLEDTEXT, 0, VarPtr(temp)
    
    ReDim Preserve tat((ToChar - FromChar) * 2 - 1)
    
    GetStyledTextRange = tat
    
End Function

'SCI_ALLOCATE(int bytes, <unused>)
Public Sub Allocate(bufSize As Long)
    
    Send_SCI_message SCI_ALLOCATE, bufSize, 0
    
End Sub

'SCI_ADDTEXT(int length, const char *s)
Public Sub AddText(txt As String, Optional txtLen As Long = -1)
    
    If txtLen = -1 Then txtLen = Len(txt)
    
    Send_SCI_messageStr SCI_ADDTEXT, txtLen, txt
    
End Sub

'SCI_ADDSTYLEDTEXT(int length, cell *s)
Public Sub AddStyledText(StyledTxt() As Byte, Optional txtLen As Long = -1)
    
    If txtLen = -1 Then txtLen = UBound(StyledTxt) + 1
    
    Send_SCI_message SCI_ADDSTYLEDTEXT, txtLen, VarPtr(StyledTxt(0))
    
End Sub

'SCI_APPENDTEXT(int length, const char *s)
Public Sub AppendText(txt As String, Optional txtLen As Long = -1)
    
    If txtLen = -1 Then txtLen = Len(txt)
    
    Send_SCI_messageStr SCI_APPENDTEXT, txtLen, txt
    
End Sub

'SCI_INSERTTEXT(int pos, const char *text)
Public Sub InsertText(txt As String, Optional posLen As Long = -1)
    
    Send_SCI_messageStr SCI_INSERTTEXT, posLen, txt
    
End Sub

'SCI_CLEARALL
Public Sub ClearText()
    
    Send_SCI_message SCI_CLEARALL, 0, 0
    
End Sub

'SCI_CLEARDOCUMENTSTYLE
Public Sub ClearStyle()
    
    Send_SCI_message SCI_CLEARDOCUMENTSTYLE, 0, 0
    
End Sub

'SCI_GETCHARAT(int position)
Public Function GetCharAt(pos As Long) As String

    GetCharAt = Chr(Send_SCI_message(SCI_GETCHARAT, pos, 0))
    
End Function

'SCI_GETSTYLEAT(int position)
Public Function GetStyleAt(pos As Long) As Long

    GetStyleAt = Send_SCI_message(SCI_GETSTYLEAT, pos, 0)
    
End Function

'SCI_SETSTYLEBITS(int bits)
Public Property Let StyleBits(bits As Long)

    Send_SCI_message SCI_SETSTYLEBITS, bits, 0
    
End Property

'SCI_GETSTYLEBITS
Public Property Get StyleBits() As Long
Attribute StyleBits.VB_ProcData.VB_Invoke_Property = "MainSettings"

    StyleBits = Send_SCI_message(SCI_GETSTYLEBITS, 0, 0)
    
End Property

'Searching:
'

'SCI_FINDTEXT(int flags, TextToFind *ttf)
Public Function FindText(flags As sci_SearchFlags, cpMin As Long, cpMax As Long, TextToFind As String) As Long()
Dim temp As sci_TextToFind, tmp() As Byte, rez(1) As Long
    
    tmp = StrConv(TextToFind, vbFromUnicode)
    ReDim Preserve tmp(UBound(tmp) + 1)
    
    tmp(UBound(tmp)) = 0
    temp.chrg.cpMin = cpMin
    temp.chrg.cpMax = cpMax
    temp.lpstrText = VarPtr(tmp(0))

    Send_SCI_message SCI_FINDTEXT, flags, VarPtr(temp)
    rez(0) = temp.chrgText.cpMin
    rez(1) = temp.chrgText.cpMax
    FindText = rez
    
End Function

'SCI_SEARCHANCHOR
Public Sub SearchAnchor()

    Send_SCI_message SCI_SEARCHANCHOR, 0, 0
    
End Sub

'SCI_SEARCHNEXT(int searchFlags, const char *text)
Public Function SearchNext(flags As sci_SearchFlags, TextToFind As String) As Long
Dim tmp() As Byte
    
    tmp = StrConv(TextToFind, vbFromUnicode)
    ReDim Preserve tmp(UBound(tmp) + 1)
    tmp(UBound(tmp)) = 0
    SearchNext = Send_SCI_message(SCI_SEARCHNEXT, flags, VarPtr(tmp(0)))
    
End Function

'SCI_SEARCHPREV(int searchFlags, const char *text)
Public Function SearchPrev(flags As sci_SearchFlags, TextToFind As String) As Long
Dim tmp() As Byte
    
    tmp = StrConv(TextToFind, vbFromUnicode)
    ReDim Preserve tmp(UBound(tmp) + 1)
    tmp(UBound(tmp)) = 0
    SearchPrev = Send_SCI_message(SCI_SEARCHPREV, flags, VarPtr(tmp(0)))

End Function

'Search and replace using the target :
'

'SCI_GETTARGETSTART
Public Property Get TargetStart() As Long
Attribute TargetStart.VB_ProcData.VB_Invoke_Property = "MainSettings"

    TargetStart = Send_SCI_message(SCI_GETTARGETSTART, 0, 0)
    
End Property

'SCI_SETTARGETSTART(int pos)
Public Property Let TargetStart(pos As Long)
    
    Send_SCI_message SCI_SETTARGETSTART, pos, 0
    
End Property

'SCI_GETTARGETEND
Public Property Get TargetEnd() As Long
Attribute TargetEnd.VB_ProcData.VB_Invoke_Property = "MainSettings"

    TargetEnd = Send_SCI_message(SCI_GETTARGETEND, 0, 0)
    
End Property

'SCI_SETTARGETEND(int pos)
Public Property Let TargetEnd(pos As Long)
    
    Send_SCI_message SCI_SETTARGETEND, pos, 0
    
End Property

'SCI_TARGETFROMSELECTION
Public Sub TargetFromSelection()

    Send_SCI_message SCI_TARGETFROMSELECTION, 0, 0

End Sub


'SCI_GETSEARCHFLAGS
Public Property Get SearchFlags() As sci_SearchFlags

    SearchFlags = Send_SCI_message(SCI_GETSEARCHFLAGS, 0, 0)
    
End Property

'SCI_SETSEARCHFLAGS(int searchFlags)
Public Property Let SearchFlags(flags As sci_SearchFlags)
    
    Send_SCI_message SCI_SETSEARCHFLAGS, flags, 0
    
End Property


'SCI_SEARCHINTARGET(int length, const char *text)
Public Function SearchinTarget(strToFind As String) As Long

    SearchinTarget = Send_SCI_messageStr(SCI_SEARCHINTARGET, Len(strToFind), strToFind)

End Function

'SCI_REPLACETARGET(int length, const char *text)
Public Function ReplaceTarget(strWithText As String) As Long

    ReplaceTarget = Send_SCI_messageStr(SCI_REPLACETARGET, Len(strWithText), strWithText)

End Function

'SCI_REPLACETARGETRE(int length, const char *text)
Public Function ReplaceTargetRegEx(strWithText As String) As Long

    ReplaceTargetRegEx = Send_SCI_messageStr(SCI_REPLACETARGETRE, Len(strWithText), strWithText)

End Function

'Overtype:
'

'SCI_GETOVERTYPE
Public Property Get EnableOverType() As Boolean
Attribute EnableOverType.VB_ProcData.VB_Invoke_Property = "MainSettings"

    EnableOverType = Send_SCI_message(SCI_GETOVERTYPE, 0, 0)

End Property

'SCI_SETOVERTYPE(bool overType)
Public Property Let EnableOverType(val As Boolean)

    Send_SCI_message SCI_SETOVERTYPE, val And 1, 0

End Property

'Cut, copy and paste:
'

'Std clipboard funct
'Clip*
'SCI_CUT
Public Sub ClipCut()

    Send_SCI_message SCI_CUT, 0, 0
    
End Sub

'SCI_COPY
Public Sub ClipCopy()

    Send_SCI_message SCI_COPY, 0, 0

End Sub

'SCI_PASTE
Public Sub ClipPase()

    Send_SCI_message SCI_PASTE, 0, 0

End Sub

'SCI_CLEAR
Public Sub ClipClear()

    Send_SCI_message SCI_CLEAR, 0, 0
    
End Sub

'SCI_CANPASTE
Public Sub ClipCanPaste()
    
        Send_SCI_message SCI_CANPASTE, 0, 0

End Sub

'SCI_COPYRANGE(int start, end)
Public Sub ClipCopyRange(FromChar As Long, ToChar As Long)
    
        Send_SCI_message SCI_COPYRANGE, FromChar, ToChar

End Sub

'SCI_COPYTEXT(int length, const char *text)
Public Sub ClipCopyText(strText As String)
    
        Send_SCI_messageStr SCI_COPYTEXT, Len(strText), strText

End Sub

'Error handling:
'

'SCI_GETSTATUS
Public Property Get ErrStatus() As Long
Attribute ErrStatus.VB_ProcData.VB_Invoke_Property = "MainSettings"

    ErrStatus = Send_SCI_message(SCI_GETSTATUS, 0, 0)

End Property

'SCI_SETSTATUS(int status)
Public Property Let ErrStatus(newStat As Long)

    Send_SCI_message SCI_SETSTATUS, newStat, 0

End Property

'From here , most of the code body  is generated using a tool
'That parsed the SCI help and created most of the code (i just did type fixing and
'ansi/unicode/pointer convertions ;) )
'
'

'SCI_UNDO
Public Function Undo() As Long
  
  Undo = Send_SCI_message(SCI_UNDO, 0, 0)
  
End Function

'SCI_CANUNDO
Public Function CanUndo() As Boolean
  
  CanUndo = Send_SCI_message(SCI_CANUNDO, 0, 0)
  
End Function

'SCI_EMPTYUNDOBUFFER
Public Function EmptyUndoBuffer() As Long
  
  EmptyUndoBuffer = Send_SCI_message(SCI_EMPTYUNDOBUFFER, 0, 0)
  
End Function

'SCI_REDO
Public Function Redo() As Long
  
  Redo = Send_SCI_message(SCI_REDO, 0, 0)
  
End Function

'SCI_CANREDO
Public Function CanRedo() As Boolean
  
  CanRedo = Send_SCI_message(SCI_CANREDO, 0, 0)
  
End Function

'SCI_SETUNDOCOLLECTION(bool collectUndo)
Public Property Let UndoCollection(CollectUndo As Boolean)
  
  Send_SCI_message SCI_SETUNDOCOLLECTION, CollectUndo And 1, 0
  
End Property

'SCI_GETUNDOCOLLECTION
Public Property Get UndoCollection() As Boolean
Attribute UndoCollection.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  UndoCollection = Send_SCI_message(SCI_GETUNDOCOLLECTION, 0, 0)
  
End Property

'SCI_BEGINUNDOACTION
Public Function BeginUndoAction() As Long
  
  BeginUndoAction = Send_SCI_message(SCI_BEGINUNDOACTION, 0, 0)
  
End Function

'SCI_ENDUNDOACTION
Public Function EndUndoAction() As Long
  
  EndUndoAction = Send_SCI_message(SCI_ENDUNDOACTION, 0, 0)
  
End Function

'SCI_GETTEXTLENGTH
Public Function TextLength() As Long
  
  TextLength = Send_SCI_message(SCI_GETTEXTLENGTH, 0, 0)
  
End Function

'SCI_GETLENGTH
Public Function length() As Long
  
  length = Send_SCI_message(SCI_GETLENGTH, 0, 0)
  
End Function

'SCI_GETLINECOUNT
Public Function LineCount() As Long
  
  LineCount = Send_SCI_message(SCI_GETLINECOUNT, 0, 0)
  
End Function

'SCI_GETFIRSTVISIBLELINE
Public Function FirstVisibleLine() As Long
  
  FirstVisibleLine = Send_SCI_message(SCI_GETFIRSTVISIBLELINE, 0, 0)
  
End Function

'SCI_LINESONSCREEN
Public Function LinesOnScreen() As Long
  
  LinesOnScreen = Send_SCI_message(SCI_LINESONSCREEN, 0, 0)
  
End Function

'SCI_GETMODIFY
Public Function Modify() As Boolean
  
  Modify = Send_SCI_message(SCI_GETMODIFY, 0, 0)
  
End Function

'SCI_SETSEL(int anchorPos, currentPos)
Public Sub Selection(anchorPos As Long, CurrentPos As Long)
  
  Send_SCI_message SCI_SETSEL, anchorPos, CurrentPos
  
End Sub

'SCI_GOTOPOS(int position)
Public Function GotoPos(position As Long) As Long
  
  GotoPos = Send_SCI_message(SCI_GOTOPOS, position, 0)
  
End Function

'SCI_GOTOLINE(int line)
Public Function GotoLine(line As Long) As Long
  
  GotoLine = Send_SCI_message(SCI_GOTOLINE, line, 0)
  
End Function

'SCI_SETCURRENTPOS(int position)
Public Property Let CurrentPos(position As Long)
  
  Send_SCI_message SCI_SETCURRENTPOS, position, 0
  
End Property

'SCI_GETCURRENTPOS
Public Property Get CurrentPos() As Long
Attribute CurrentPos.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  CurrentPos = Send_SCI_message(SCI_GETCURRENTPOS, 0, 0)
  
End Property

'SCI_SETANCHOR(int position)
Public Property Let Anchor(position As Long)
  
  Send_SCI_message SCI_SETANCHOR, position, 0
  
End Property

'SCI_GETANCHOR
Public Property Get Anchor() As Long
Attribute Anchor.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  Anchor = Send_SCI_message(SCI_GETANCHOR, 0, 0)
  
End Property

'SCI_SETSELECTIONSTART(int position)
Public Property Let SelectionStart(position As Long)
  
  Send_SCI_message SCI_SETSELECTIONSTART, position, 0
  
End Property

'SCI_GETSELECTIONSTART
Public Property Get SelectionStart() As Long
Attribute SelectionStart.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  SelectionStart = Send_SCI_message(SCI_GETSELECTIONSTART, 0, 0)
  
End Property

'SCI_SETSELECTIONEND(int position)
Public Property Let SelectionEnd(position As Long)
  
  Send_SCI_message SCI_SETSELECTIONEND, position, 0
  
End Property

'SCI_GETSELECTIONEND
Public Property Get SelectionEnd() As Long
Attribute SelectionEnd.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  SelectionEnd = Send_SCI_message(SCI_GETSELECTIONEND, 0, 0)
  
End Property

'SCI_SELECTALL
Public Sub SelectAll()
  
  Send_SCI_message SCI_SELECTALL, 0, 0
  
End Sub

'SCI_LINEFROMPOSITION(int position)
Public Function LineFromPosition(position As Long) As Long
  
  LineFromPosition = Send_SCI_message(SCI_LINEFROMPOSITION, position, 0)
  
End Function

'SCI_POSITIONFROMLINE(int line)
Public Function PositionFromLine(line As Long) As Long
  
  PositionFromLine = Send_SCI_message(SCI_POSITIONFROMLINE, line, 0)
  
End Function

'SCI_GETLINEENDPOSITION(int line)
Public Function LineEndPosition(line As Long) As Long
  
  LineEndPosition = Send_SCI_message(SCI_GETLINEENDPOSITION, line, 0)
  
End Function

'heeh this is made by me not auto converted .. heheh lol
'SCI_LINELENGTH(int line).
Public Function GetLineLen(numLine As Long) As Long
    
    GetLineLen = Send_SCI_message(SCI_GETLINE, numLine, 0)
    
End Function

'SCI_GETCOLUMN(int position)
Public Function GetColumn(position As Long) As Long
  
  GetColumn = Send_SCI_message(SCI_GETCOLUMN, position, 0)
  
End Function

'SCI_POSITIONFROMPOINT(int x, y)
Public Function PositionFromPoint(x As Long, y As Long) As Long
  
  PositionFromPoint = Send_SCI_message(SCI_POSITIONFROMPOINT, x, y)
  
End Function

'SCI_POSITIONFROMPOINTCLOSE(int x, y)
Public Function PositionFromPointClose(x As Long, y As Long) As Long
  
  PositionFromPointClose = Send_SCI_message(SCI_POSITIONFROMPOINTCLOSE, x, y)
  
End Function

'SCI_POINTXFROMPOSITION(<unused>, position)
Public Function PointXFromPosition(position As Long) As Long
  
  PointXFromPosition = Send_SCI_message(SCI_POINTXFROMPOSITION, 0, position)
  
End Function

'SCI_POINTYFROMPOSITION(<unused>, position)
Public Function PointYFromPosition(position As Long) As Long
  
  PointYFromPosition = Send_SCI_message(SCI_POINTYFROMPOSITION, 0, position)
  
End Function

'SCI_HIDESELECTION(bool hide)
Public Property Let HideSelection(hide As Boolean)
  
   Send_SCI_message SCI_HIDESELECTION, hide And 1, 0
   SelVisible = hide
   
End Property

'This is an enhasement ... it does not exist on the real control
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  HideSelection = SelVisible
  
End Property


'SCI_GETSELTEXT(<unused>, char *text)
Public Function GetSelText() As String
    Dim temp As Long, tstr() As Byte
    
    temp = Send_SCI_message(SCI_GETSELTEXT, 0, 0)
    If temp < 2 Then GetSelText = vbNullString: Exit Function
    ReDim tstr(temp - 1)
    Send_SCI_message SCI_GETSELTEXT, 0, VarPtr(tstr(0))
    ReDim Preserve tstr(temp - 2)
    GetSelText = StrConv(tstr, vbUnicode)
  
End Function

'SCI_GETCURLINE(int textLen, char *text)
Public Function GetCurLine() As String
    Dim temp As Long, tstr() As Byte
    
    temp = Send_SCI_message(SCI_GETCURLINE, 0, 0)
    If temp < 2 Then GetCurLine = vbNullString: Exit Function
    ReDim tstr(temp - 1)
    Send_SCI_message SCI_GETCURLINE, temp, VarPtr(tstr(0))
    ReDim Preserve tstr(temp - 2)
    GetCurLine = StrConv(tstr, vbUnicode)
  
End Function

'SCI_SELECTIONISRECTANGLE
Public Function SelectionIsRectangle() As Boolean
  
  SelectionIsRectangle = Send_SCI_message(SCI_SELECTIONISRECTANGLE, 0, 0)
  
End Function

'A note here :  the docs are wrong..
'This one takes one parameter..
'it should be SCI_SETSELECTIONMODE(int mode) instead of
'SCI_SETSELECTIONMODE
Public Property Let SelectionMode(mode As sci_SelectionMode)

  Send_SCI_message SCI_SETSELECTIONMODE, mode, 0

End Property

'SCI_GETSELECTIONMODE
Public Property Get SelectionMode() As sci_SelectionMode
  
  SelectionMode = Send_SCI_message(SCI_GETSELECTIONMODE, 0, 0)
  
End Property

'SCI_GETLINESELSTARTPOSITION(int line)
Public Function GetLineSelStartPosition(line As Long) As Long
  
  GetLineSelStartPosition = Send_SCI_message(SCI_GETLINESELSTARTPOSITION, line, 0)
  
End Function

'SCI_GETLINESELENDPOSITION(int line)
Public Function GetLineSelEndPosition(line As Long) As Long
  
  GetLineSelEndPosition = Send_SCI_message(SCI_GETLINESELENDPOSITION, line, 0)
  
End Function

'SCI_MOVECARETINSIDEVIEW
Public Function MoveCaretInsideView() As Long
  
  MoveCaretInsideView = Send_SCI_message(SCI_MOVECARETINSIDEVIEW, 0, 0)
  
End Function

'SCI_WORDENDPOSITION(int position, bool onlyWordCharacters)
Public Function WordEndPosition(position As Long, onlyWordCharacters As Boolean) As Long
  
  WordEndPosition = Send_SCI_message(SCI_WORDENDPOSITION, position, onlyWordCharacters And 1)
  
End Function

'SCI_WORDSTARTPOSITION(int position, bool onlyWordCharacters)
Public Function WordStartPosition(position As Long, onlyWordCharacters As Boolean) As Long
  
  WordStartPosition = Send_SCI_message(SCI_WORDSTARTPOSITION, position, onlyWordCharacters And 1)
  
End Function

'SCI_POSITIONBEFORE(int position)
Public Function PositionBefore(position As Long) As Long
  
  PositionBefore = Send_SCI_message(SCI_POSITIONBEFORE, position, 0)
  
End Function

'SCI_POSITIONAFTER(int position)
Public Function POSITIONAFTER(position As Long) As Long
  
  POSITIONAFTER = Send_SCI_message(SCI_POSITIONAFTER, position, 0)
  
End Function

'SCI_TEXTWIDTH(int styleNumber, const char *text)
Public Function TextWidth(styleNumber As Long, Text As String) As Long
  
  TextWidth = Send_SCI_messageStr(SCI_TEXTWIDTH, styleNumber, Text)
  
End Function

'SCI_TEXTHEIGHT(int line)
Public Function TextHeight(line As Long) As Long
  
  TextHeight = Send_SCI_message(SCI_TEXTHEIGHT, line, 0)
  
End Function

'SCI_CHOOSECARETX
Public Function ChooseCaretX() As Long
  
  ChooseCaretX = Send_SCI_message(SCI_CHOOSECARETX, 0, 0)
  
End Function



'SCI_LINESCROLL(int column, line)
Public Function LineScroll(column As Long, line As Long) As Long
  
  LineScroll = Send_SCI_message(SCI_LINESCROLL, column, line)
  
End Function

'SCI_SCROLLCARET
Public Function SCROLLCARET() As Long
  
  SCROLLCARET = Send_SCI_message(SCI_SCROLLCARET, 0, 0)
  
End Function

'SCI_SETXCARETPOLICY(int caretPolicy, caretSlop)
Public Sub SetXCaretPolicy(caretPolicy As Long, caretSlop As Long)
  
  Send_SCI_message SCI_SETXCARETPOLICY, caretPolicy, caretSlop
  
End Sub

'SCI_SETYCARETPOLICY(int caretPolicy, caretSlop)
Public Sub SetYCaretPolicy(caretPolicy As Long, caretSlop As Long)
  
  Send_SCI_message SCI_SETYCARETPOLICY, caretPolicy, caretSlop
  
End Sub

 
'SCI_SETVISIBLEPOLICY(int caretPolicy, caretSlop)
Public Sub SetVisiblePolicy(caretPolicy As Long, caretSlop As Long)
  
  Send_SCI_message SCI_SETVISIBLEPOLICY, caretPolicy, caretSlop
  
End Sub

'SCI_SETHSCROLLBAR(bool visible)
Public Property Let HScrollBar(bvisible As Boolean)
  
  Send_SCI_message SCI_SETHSCROLLBAR, bvisible And 1, 0
  
End Property

'SCI_GETHSCROLLBAR
Public Property Get HScrollBar() As Boolean
Attribute HScrollBar.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  HScrollBar = Send_SCI_message(SCI_GETHSCROLLBAR, 0, 0)
  
End Property

'SCI_SETVSCROLLBAR(bool visible)
Public Property Let VScrollBar(visible As Boolean)
  
  Send_SCI_message SCI_SETVSCROLLBAR, visible And 1, 0
  
End Property

'SCI_GETVSCROLLBAR
Public Property Get VScrollBar() As Boolean
Attribute VScrollBar.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  VScrollBar = Send_SCI_message(SCI_GETVSCROLLBAR, 0, 0)
  
End Property

'SCI_GETXOFFSET
Public Property Get xOffset() As Long
Attribute xOffset.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  xOffset = Send_SCI_message(SCI_GETXOFFSET, 0, 0)
  
End Property

'SCI_SETXOFFSET(int xOffset)
Public Property Let xOffset(xOffset As Long)
  
  Send_SCI_message SCI_SETXOFFSET, xOffset, 0
  
End Property

'SCI_SETSCROLLWIDTH(int pixelWidth)
Public Property Let ScrollWidth(pixelWidth As Long)
  
  Send_SCI_message SCI_SETSCROLLWIDTH, pixelWidth, 0
  
End Property

'SCI_GETSCROLLWIDTH
Public Property Get ScrollWidth() As Long
Attribute ScrollWidth.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  ScrollWidth = Send_SCI_message(SCI_GETSCROLLWIDTH, 0, 0)
  
End Property

'SCI_SETENDATLASTLINE(bool endAtLastLine)
Public Property Let EndAtLastLine(bEndAtLastLine As Boolean)
  
  Send_SCI_message SCI_SETENDATLASTLINE, bEndAtLastLine And 1, 0
  
End Property

'SCI_GETENDATLASTLINE
Public Property Get EndAtLastLine() As Boolean
Attribute EndAtLastLine.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  EndAtLastLine = Send_SCI_message(SCI_GETENDATLASTLINE, 0, 0)
  
End Property

'SCI_SETVIEWWS(int wsMode)
Public Property Let ViewWhiteSpace(wsMode As Long)
  
  Send_SCI_message SCI_SETVIEWWS, wsMode, 0
  
End Property

'SCI_GETVIEWWS
Public Property Get ViewWhiteSpace() As Long
Attribute ViewWhiteSpace.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  ViewWhiteSpace = Send_SCI_message(SCI_GETVIEWWS, 0, 0)
  
End Property

'SCI_SETWHITESPACEFORE(bool useWhitespaceForeColour, colour)
Public Property Let WhiteSpaceFore(useWhitespaceForeColour As Boolean, colour As Long)

  MsgBox "add get"
  Send_SCI_message SCI_SETWHITESPACEFORE, useWhitespaceForeColour And 1, colour
  
End Property

'SCI_SETWHITESPACEBACK(bool useWhitespaceBackColour, colour)
Public Property Let WhiteSpaceBack(useWhitespaceBackColour As Boolean, colour As Long)
  
  MsgBox "add get"
  Send_SCI_message SCI_SETWHITESPACEBACK, useWhitespaceBackColour And 1, colour
  
End Property

'SCI_SETCURSOR(int curType)
Public Property Let Cursor(curType As sci_CursorStyle)
  
  Send_SCI_message SCI_SETCURSOR, curType, 0
  
End Property

'SCI_GETCURSOR
Public Property Get Cursor() As sci_CursorStyle
  
  Cursor = Send_SCI_message(SCI_GETCURSOR, 0, 0)
  
End Property

'SCI_SETMOUSEDOWNCAPTURES(bool captures)
Public Property Let MouseDownCaptures(captures As Boolean)
  
  Send_SCI_message SCI_SETMOUSEDOWNCAPTURES, 1 And captures, 0
  
End Property

'SCI_GETMOUSEDOWNCAPTURES
Public Property Get MouseDownCaptures() As Boolean
Attribute MouseDownCaptures.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  MouseDownCaptures = Send_SCI_message(SCI_GETMOUSEDOWNCAPTURES, 0, 0)
  
End Property

'SCI_SETEOLMODE(int eolMode)
Public Property Let EOLMode(neweolMode As sci_EOLModes)
  
  Send_SCI_message SCI_SETEOLMODE, EOLMode, 0
  
End Property

'SCI_GETEOLMODE
Public Property Get EOLMode() As sci_EOLModes
  
  EOLMode = Send_SCI_message(SCI_GETEOLMODE, 0, 0)
  
End Property

'SCI_CONVERTEOLS(int eolMode)
Public Sub ConvertEOLs(neolMode As sci_EOLModes)
  
   Send_SCI_message SCI_CONVERTEOLS, EOLMode, 0
  
End Sub

'SCI_SETVIEWEOL(bool visible)
Public Property Let ViewEOL(visible As Boolean)
  
  Send_SCI_message SCI_SETVIEWEOL, 1 And visible, 0
  
End Property

'SCI_GETVIEWEOL
Public Property Get ViewEOL() As Boolean
Attribute ViewEOL.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  ViewEOL = Send_SCI_message(SCI_GETVIEWEOL, 0, 0)
  
End Property

'SCI_GETENDSTYLED
Public Property Get EndStyled() As Long
  
  EndStyled = Send_SCI_message(SCI_GETENDSTYLED, 0, 0)
  
End Property

'SCI_STARTSTYLING(int position, mask)
Public Sub StartStyling(position As Long, mask As Long)
  
   Send_SCI_message SCI_STARTSTYLING, position, mask
  
End Sub

'SCI_SETSTYLING(int length, style)
Public Sub SetStyling(length As Long, style As Long)
  
  Send_SCI_message SCI_SETSTYLING, length, style
  
End Sub

'SCI_SETSTYLINGEX(int length, const char *styles)
Public Sub SetStylengEx(nam As String)
  
  Send_SCI_messageStr SCI_SETSTYLINGEX, Len(nam), nam
  
End Sub

'SCI_SETLINESTATE(int line, value )
Public Property Let LineState(line As Long, value As Long)
  
  Send_SCI_message SCI_SETLINESTATE, line, value
  
End Property

'SCI_GETLINESTATE(int line)
Public Property Get LineState(line As Long) As Long
Attribute LineState.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  LineState = Send_SCI_message(SCI_GETLINESTATE, line, 0)
  
End Property

'SCI_GETMAXLINESTATE
Public Function GetMaxLineState() As Long
  
  GetMaxLineState = Send_SCI_message(SCI_GETMAXLINESTATE, 0, 0)
  
End Function


'SCI_STYLERESETDEFAULT
Public Sub StyleResetDefault()
  
   Send_SCI_message SCI_STYLERESETDEFAULT, 0, 0
  
End Sub

'SCI_STYLECLEARALL
Public Sub StyleClearAll()
  
  Send_SCI_message SCI_STYLECLEARALL, 0, 0
  
End Sub


'SCI_STYLESETFONT(int styleNumber, char *fontName)
Public Sub StyleSetFont(styleNumber As Long, fName As String)
  
  Send_SCI_messageStr SCI_STYLESETFONT, styleNumber, fName
  
End Sub

'SCI_STYLESETSIZE(int styleNumber, sizeInPoints)
Public Sub STYLESETSIZE(styleNumber As Long, sizeInPoints As Long)
  
  Send_SCI_message SCI_STYLESETSIZE, styleNumber, sizeInPoints
  
End Sub

'SCI_STYLESETBOLD(int styleNumber, bool bold)
Public Sub StyleSetBold(styleNumber As Long, bold As Boolean)
  
  Send_SCI_message SCI_STYLESETBOLD, styleNumber, 1 And bold
  
End Sub

'SCI_STYLESETITALIC(int styleNumber, bool italic)
Public Sub StyleSetItalic(styleNumber As Long, italic As Boolean)
  
  Send_SCI_message SCI_STYLESETITALIC, styleNumber, 1 And italic
  
End Sub

'SCI_STYLESETUNDERLINE(int styleNumber, bool underline)
Public Sub StyleSetUnderline(styleNumber As Long, underline As Boolean)
  
    Send_SCI_message SCI_STYLESETUNDERLINE, styleNumber, 1 And underline
  
End Sub

'SCI_STYLESETFORE(int styleNumber, colour)
Public Sub StyleSetFore(styleNumber As Long, colour As Long)
  
  Send_SCI_message SCI_STYLESETFORE, styleNumber, colour
  
End Sub

'SCI_STYLESETBACK(int styleNumber, colour)
Public Sub StyleSetBack(styleNumber As Long, colour As Long)
  
  Send_SCI_message SCI_STYLESETBACK, styleNumber, colour
  
End Sub

'SCI_STYLESETEOLFILLED(int styleNumber, bool eolFilled)
Public Sub STYLESETEOLFILLED(styleNumber As Long, eolFilled As Boolean)
  
  Send_SCI_message SCI_STYLESETEOLFILLED, styleNumber, eolFilled And 1
  
End Sub

'SCI_STYLESETCHARACTERSET(int styleNumber, charSet)
Public Sub StyleSetCharacterset(styleNumber As Long, charSet As Long)
  
  Send_SCI_message SCI_STYLESETCHARACTERSET, styleNumber, charSet
  
End Sub

'SCI_STYLESETCASE(int styleNumber, caseMode)
Public Sub StyleSetCase(styleNumber As Long, caseMode As Long)
     
    Send_SCI_message SCI_STYLESETCASE, styleNumber, caseMode
  
End Sub

'SCI_STYLESETVISIBLE(int styleNumber, bool visible)
Public Sub StyleSetVisible(styleNumber As Long, visible As Boolean)
  
   Send_SCI_message SCI_STYLESETVISIBLE, styleNumber, visible And 1
  
End Sub

'SCI_STYLESETCHANGEABLE(int styleNumber, bool changeable)
Public Sub StyleSetChangeable(styleNumber As Long, changeable As Boolean)
  
  Send_SCI_message SCI_STYLESETCHANGEABLE, styleNumber, 1 And changeable
  
End Sub

'SCI_STYLESETHOTSPOT(int styleNumber, bool hotspot)
Public Sub StyleSetHotSpot(styleNumber As Long, hotspot As Boolean)
  
  Send_SCI_message SCI_STYLESETHOTSPOT, styleNumber, 1 And hotspot
  
End Sub



'SCI_SETSELFORE(bool useSelectionForeColour, colour)
Public Property Let SelcectedForeColor(useSelectionForeColour As Boolean, colour As Long)

  MsgBox "Add Get"
  Send_SCI_message SCI_SETSELFORE, 1 And useSelectionForeColour, colour
  
End Property

'SCI_SETSELBACK(bool useSelectionBackColour, colour)
Public Property Let SelBack(useSelectionBackColour As Boolean, colour As Long)
  
  MsgBox "Add Get"
  Send_SCI_message SCI_SETSELBACK, 1 And useSelectionBackColour, colour
  
End Property

'SCI_SETCARETFORE(int colour)
Public Property Let CaretFore(colour As Long)
  
  Send_SCI_message SCI_SETCARETFORE, colour, 0
  
End Property

'SCI_GETCARETFORE
Public Property Get CaretFore() As Long
Attribute CaretFore.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  CaretFore = Send_SCI_message(SCI_GETCARETFORE, 0, 0)
  
End Property

'SCI_SETCARETLINEVISIBLE(bool show)
Public Property Let CaretLineVisible(show As Boolean)
  
  Send_SCI_message SCI_SETCARETLINEVISIBLE, 1 And show, 0
  
End Property

'SCI_GETCARETLINEVISIBLE
Public Property Get CaretLineVisible() As Boolean
Attribute CaretLineVisible.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  CaretLineVisible = Send_SCI_message(SCI_GETCARETLINEVISIBLE, 0, 0)
  
End Property



'SCI_SETCARETLINEBACK(int colour)
Public Property Let CaretLineBack(colour As Long)
  
  Send_SCI_message SCI_SETCARETLINEBACK, colour, 0
  
End Property

'SCI_GETCARETLINEBACK
Public Property Get CaretLineBack() As Long
Attribute CaretLineBack.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  CaretLineBack = Send_SCI_message(SCI_GETCARETLINEBACK, 0, 0)
  
End Property

'SCI_SETCARETPERIOD(int milliseconds)
Public Property Let CaretPeriod(milliseconds As Long)
  
  Send_SCI_message SCI_SETCARETPERIOD, milliseconds, 0
  
End Property

'SCI_GETCARETPERIOD
Public Property Get CaretPeriod() As Long
Attribute CaretPeriod.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  CaretPeriod = Send_SCI_message(SCI_GETCARETPERIOD, 0, 0)
  
End Property

'SCI_SETCARETWIDTH(int pixels)
Public Property Let CaretWidth(pixels As Long)
  
  Send_SCI_message SCI_SETCARETWIDTH, pixels, 0
  
End Property

'SCI_GETCARETWIDTH
Public Property Get CaretWidth() As Long
Attribute CaretWidth.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  CaretWidth = Send_SCI_message(SCI_GETCARETWIDTH, 0, 0)
  
End Property

'SCI_SETHOTSPOTACTIVEFORE
Public Property Let HotSpotActiveFore(val As Long)
  
  Send_SCI_message SCI_SETHOTSPOTACTIVEFORE, val, 0
  
End Property

'SCI_SETHOTSPOTACTIVEBACK
Public Property Let HOTSPOTACTIVEBACK(val As Long)
  
  Send_SCI_message SCI_SETHOTSPOTACTIVEBACK, val, 0
  
End Property

'SCI_SETHOTSPOTACTIVEUNDERLINE
Public Property Let HOTSPOTACTIVEUNDERLINE(val As Long)
  
  Send_SCI_message SCI_SETHOTSPOTACTIVEUNDERLINE, val, 0
  
End Property

'SCI_SETHOTSPOTSINGLELINE
Public Property Let HOTSPOTSINGLELINE(val As Long)
  
  Send_SCI_message SCI_SETHOTSPOTSINGLELINE, val, 0
  
End Property

'SCI_SETCONTROLCHARSYMBOL(int symbol)
Public Property Let ControlCharSymbol(symbol As Long)
  
  Send_SCI_message SCI_SETCONTROLCHARSYMBOL, symbol, 0
  
End Property

'SCI_GETCONTROLCHARSYMBOL
Public Property Get ControlCharSymbol() As Long
Attribute ControlCharSymbol.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  ControlCharSymbol = Send_SCI_message(SCI_GETCONTROLCHARSYMBOL, 0, 0)
  
End Property



'SCI_SETMARGINTYPEN(int margin, type)
Public Property Let MarginType(margin As Long, mType As Long)
  
  Send_SCI_message SCI_SETMARGINTYPEN, margin, mType
  
End Property

'SCI_GETMARGINTYPEN(int margin)
Public Property Get MarginType(margin As Long) As Long
Attribute MarginType.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  MarginType = Send_SCI_message(SCI_GETMARGINTYPEN, margin, 0)
  
End Property


'SCI_SETMARGINWIDTHN(int margin, pixelWidth)
Public Property Let MarginWidth(margin As Long, pixelWidth As Long)
  
  Send_SCI_message SCI_SETMARGINWIDTHN, margin, pixelWidth
  
End Property

'SCI_GETMARGINWIDTHN(int margin)
Public Property Get MarginWidth(margin As Long) As Long
Attribute MarginWidth.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  MarginWidth = Send_SCI_message(SCI_GETMARGINWIDTHN, margin, 0)
  
End Property

'SCI_SETMARGINMASKN(int margin, mask)
Public Property Let MarginMask(margin As Long, mask As Long)
  
  Send_SCI_message SCI_SETMARGINMASKN, margin, mask
  
End Property

'SCI_GETMARGINMASKN(int margin)
Public Property Get MarginMask(margin As Long) As Long
Attribute MarginMask.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  MarginMask = Send_SCI_message(SCI_GETMARGINMASKN, margin, 0)
  
End Property

'SCI_SETMARGINSENSITIVEN(int margin, bool sensitive)
Public Property Let MarginSensitive(margin As Long, sensitive As Boolean)
  
  Send_SCI_message SCI_SETMARGINSENSITIVEN, margin, 1 And sensitive
  
End Property

'SCI_GETMARGINSENSITIVEN(int margin)
Public Property Get MarginSensitive(margin As Long) As Boolean
Attribute MarginSensitive.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  MarginSensitive = Send_SCI_message(SCI_GETMARGINSENSITIVEN, margin, 0)
  
End Property

'SCI_SETMARGINLEFT(<unused>, pixels)
Public Property Let MarginLeft(pixels As Long)
  
  Send_SCI_message SCI_SETMARGINLEFT, 0, pixels
  
End Property

'SCI_GETMARGINLEFT
Public Property Get MarginLeft() As Long
Attribute MarginLeft.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  MarginLeft = Send_SCI_message(SCI_GETMARGINLEFT, 0, 0)
  
End Property

'SCI_SETMARGINRIGHT(<unused>, pixels)
Public Property Let MarginRight(pixels As Long)
  
  Send_SCI_message SCI_SETMARGINRIGHT, 0, pixels
  
End Property

'SCI_GETMARGINRIGHT
Public Property Get MarginRight() As Long
Attribute MarginRight.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  MarginRight = Send_SCI_message(SCI_GETMARGINRIGHT, 0, 0)
  
End Property

'SCI_SETFOLDMARGINCOLOUR(bool useSetting, colour)
Public Property Let FoldMarginColour(useSetting As Boolean, colour As Long)
  
  Send_SCI_message SCI_SETFOLDMARGINCOLOUR, 1 And useSetting, colour
  
End Property

'SCI_SETFOLDMARGINHICOLOUR(bool useSetting, colour)
Public Property Let FoldMarginHiColour(useSetting As Boolean, colour As Long)
  
  Send_SCI_message SCI_SETFOLDMARGINHICOLOUR, 1 And useSetting, colour
  
End Property

'SCI_SETMARGINTYPEN(int margin, iType)
Public Property Let MarginTypem(margin As Long, iType As Long)
  
  Send_SCI_message SCI_SETMARGINTYPEN, margin, iType
  
End Property

'SCI_SETUSEPALETTE(bool allowPaletteUse)
Public Property Let UsePalette(allowPaletteUse As Boolean)
  
  Send_SCI_message SCI_SETUSEPALETTE, 1 And allowPaletteUse, 0
  
End Property

'SCI_GETUSEPALETTE
Public Property Get UsePalette() As Boolean
Attribute UsePalette.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  UsePalette = Send_SCI_message(SCI_GETUSEPALETTE, 0, 0)
  
End Property

'SCI_SETBUFFEREDDRAW(bool isBuffered)
Public Property Let BufferEdDraw(isBuffered As Boolean)
  
  Send_SCI_message SCI_SETBUFFEREDDRAW, 1 And isBuffered, 0
  
End Property

'SCI_GETBUFFEREDDRAW
Public Property Get BufferEdDraw() As Boolean
Attribute BufferEdDraw.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  BufferEdDraw = Send_SCI_message(SCI_GETBUFFEREDDRAW, 0, 0)
  
End Property

'SCI_SETTWOPHASEDRAW(bool twoPhase)
Public Property Let TwoPhaseDraw(twoPhase As Boolean)
  
  Send_SCI_message SCI_SETTWOPHASEDRAW, 1 And twoPhase, 0

End Property

'SCI_GETTWOPHASEDRAW
Public Property Get TwoPhaseDraw() As Boolean
Attribute TwoPhaseDraw.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  TwoPhaseDraw = Send_SCI_message(SCI_GETTWOPHASEDRAW, 0, 0)
  
End Property

'SCI_SETCODEPAGE(int codePage)
Public Property Let CodePage(CodePage As Long)
  
  Send_SCI_message SCI_SETCODEPAGE, CodePage, 0
  
End Property

'SCI_GETCODEPAGE
Public Property Get CodePage() As Long
Attribute CodePage.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  CodePage = Send_SCI_message(SCI_GETCODEPAGE, 0, 0)
  
End Property

'SCI_SETWORDCHARS(<unused>, const char *chars)
Public Property Let WORDCHARS(wChars As String)
  
  Send_SCI_messageStr SCI_SETWORDCHARS, 0, wChars
  
End Property

'SCI_SETWHITESPACECHARS(<unused>, const char *chars)
Public Property Let WhiteSpaceChars(wsChars As String)
  
  Send_SCI_messageStr SCI_SETWHITESPACECHARS, 0, wsChars
  
End Property

'SCI_SETCHARSDEFAULT
Public Sub CharsDefault()
  
  Send_SCI_message SCI_SETCHARSDEFAULT, 0, 0
  
End Sub

'SCI_GRABFOCUS
Public Sub GrabFocus()
  
    Send_SCI_message SCI_GRABFOCUS, 0, 0
  
End Sub

'SCI_SETFOCUS(bool focus)
Public Property Let Focus(Focus As Boolean)
  
  Send_SCI_message SCI_SETFOCUS, 1 And Focus, 0
  
End Property

'SCI_GETFOCUS
Public Property Get Focus() As Boolean
Attribute Focus.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  Focus = Send_SCI_message(SCI_GETFOCUS, 0, 0)
  
End Property


'SCI_BRACEHIGHLIGHT(int pos1, pos2)
Public Function BRACEHIGHLIGHT(pos1 As Long, pos2 As Long) As Long
  
  BRACEHIGHLIGHT = Send_SCI_message(SCI_BRACEHIGHLIGHT, pos1, pos2)
  
End Function

'SCI_BRACEBADLIGHT(int pos1)
Public Function BRACEBADLIGHT(pos1 As Long) As Long
  
  BRACEBADLIGHT = Send_SCI_message(SCI_BRACEBADLIGHT, pos1, 0)
  
End Function

'SCI_BRACEMATCH(int position, maxReStyle)
Public Function BRACEMATCH(position As Long, maxReStyle As Long) As Long
  
  BRACEMATCH = Send_SCI_message(SCI_BRACEMATCH, position, maxReStyle)
  
End Function



'SCI_SETTABWIDTH(int widthInChars)
Public Property Let TABWIDTH(widthInChars As Long)

  Send_SCI_message SCI_SETTABWIDTH, widthInChars, 0
  
End Property

'SCI_GETTABWIDTH
Public Property Get TABWIDTH() As Long
Attribute TABWIDTH.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  TABWIDTH = Send_SCI_message(SCI_GETTABWIDTH, 0, 0)
  
End Property

'SCI_SETUSETABS(bool useTabs)
Public Property Let useTabs(useTabs As Boolean)
  
  Send_SCI_message SCI_SETUSETABS, 1 And useTabs, 0
  
End Property

'SCI_GETUSETABS
Public Property Get useTabs() As Boolean
Attribute useTabs.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  useTabs = Send_SCI_message(SCI_GETUSETABS, 0, 0)
  
End Property

'SCI_SETINDENT(int widthInChars)
Public Property Let indent(widthInChars As Long)
  
  Send_SCI_message SCI_SETINDENT, widthInChars, 0
  
End Property

'SCI_GETINDENT
Public Property Get indent() As Long
Attribute indent.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  indent = Send_SCI_message(SCI_GETINDENT, 0, 0)
  
End Property

'SCI_SETTABINDENTS(bool tabIndents)
Public Property Let tabIndents(tabIndents As Boolean)
  
  Send_SCI_message SCI_SETTABINDENTS, 1 And tabIndents, 0
  
End Property

'SCI_GETTABINDENTS
Public Property Get tabIndents() As Boolean
Attribute tabIndents.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  tabIndents = Send_SCI_message(SCI_GETTABINDENTS, 0, 0)
  
End Property

'SCI_SETBACKSPACEUNINDENTS(bool bsUnIndents)
Public Property Let BACKSPACEUNINDENTS(bsUnIndents As Boolean)
  
  Send_SCI_message SCI_SETBACKSPACEUNINDENTS, 1 And bsUnIndents, 0
  
End Property

'SCI_GETBACKSPACEUNINDENTS
Public Property Get BACKSPACEUNINDENTS() As Boolean
Attribute BACKSPACEUNINDENTS.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  BACKSPACEUNINDENTS = Send_SCI_message(SCI_GETBACKSPACEUNINDENTS, 0, 0)
  
End Property

'SCI_SETLINEINDENTATION(int line, indentation)
Public Property Let LINEINDENTATION(line As Long, indentation As Long)
  
  Send_SCI_message SCI_SETLINEINDENTATION, line, indentation
  
End Property

'SCI_GETLINEINDENTATION(int line)
Public Property Get LINEINDENTATION(line As Long) As Long
Attribute LINEINDENTATION.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  LINEINDENTATION = Send_SCI_message(SCI_GETLINEINDENTATION, line, 0)
  
End Property

'SCI_GETLINEINDENTPOSITION(int line)
Public Property Get LINEINDENTPOSITION(line As Long) As Long
  
  LINEINDENTPOSITION = Send_SCI_message(SCI_GETLINEINDENTPOSITION, line, 0)
  
End Property

'SCI_SETINDENTATIONGUIDES(bool view)
Public Property Let INDENTATIONGUIDES(view As Boolean)
  
  Send_SCI_message SCI_SETINDENTATIONGUIDES, 1 And view, 0
  
End Property

'SCI_GETINDENTATIONGUIDES
Public Property Get INDENTATIONGUIDES() As Boolean
Attribute INDENTATIONGUIDES.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  INDENTATIONGUIDES = Send_SCI_message(SCI_GETINDENTATIONGUIDES, 0, 0)
  
End Property

'SCI_SETHIGHLIGHTGUIDE(int column)
Public Property Let HIGHLIGHTGUIDE(column As Long)
  
  Send_SCI_message SCI_SETHIGHLIGHTGUIDE, column, 0
  
End Property

'SCI_GETHIGHLIGHTGUIDE
Public Property Get HIGHLIGHTGUIDE() As Long
Attribute HIGHLIGHTGUIDE.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  HIGHLIGHTGUIDE = Send_SCI_message(SCI_GETHIGHLIGHTGUIDE, 0, 0)
  
End Property

'SCI_MARKERDEFINE(int markerNumber, markerSymbols)
Public Function MARKERDEFINE(markerNumber As Long, markerSymbols As Long) As Long
  
  MARKERDEFINE = Send_SCI_message(SCI_MARKERDEFINE, markerNumber, markerSymbols)
  
End Function

'SCI_MARKERDEFINEPIXMAP(int markerNumber, const char *xpm)
Public Function MARKERDEFINEPIXMAP(markerNumber As Long, xpm As String) As Long
  
  MARKERDEFINEPIXMAP = Send_SCI_messageStr(SCI_MARKERDEFINEPIXMAP, markerNumber, xpm)
  
End Function

'SCI_MARKERSETFORE(int markerNumber, colour)
Public Function MARKERSETFORE(markerNumber As Long, colour As Long) As Long
  
  MARKERSETFORE = Send_SCI_message(SCI_MARKERSETFORE, markerNumber, colour)
  
End Function

'SCI_MARKERSETBACK(int markerNumber, colour)
Public Function MARKERSETBACK(markerNumber As Long, colour As Long) As Long
  
  MARKERSETBACK = Send_SCI_message(SCI_MARKERSETBACK, markerNumber, colour)
  
End Function

'SCI_MARKERADD(int line, markerNumber)
Public Function MARKERADD(line As Long, markerNumber As Long) As Long
  
  MARKERADD = Send_SCI_message(SCI_MARKERADD, line, markerNumber)
  
End Function

'SCI_MARKERDELETE(int line, markerNumber)
Public Function MARKERDELETE(line As Long, markerNumber As Long) As Long
  
  MARKERDELETE = Send_SCI_message(SCI_MARKERDELETE, line, markerNumber)
  
End Function

'SCI_MARKERDELETEALL(int markerNumber)
Public Function MARKERDELETEALL(markerNumber As Long) As Long
  
  MARKERDELETEALL = Send_SCI_message(SCI_MARKERDELETEALL, markerNumber, 0)
  
End Function

'SCI_MARKERGET(int line)
Public Function MARKERGET(line As Long) As Long
  
  MARKERGET = Send_SCI_message(SCI_MARKERGET, line, 0)
  
End Function

'SCI_MARKERNEXT(int lineStart, markerMask)
Public Function MARKERNEXT(lineStart As Long, markerMask As Long) As Long
  
  MARKERNEXT = Send_SCI_message(SCI_MARKERNEXT, lineStart, markerMask)
  
End Function

'SCI_MARKERPREVIOUS(int lineStart, markerMask)
Public Function MARKERPREVIOUS(lineStart As Long, markerMask As Long) As Long
  
  MARKERPREVIOUS = Send_SCI_message(SCI_MARKERPREVIOUS, lineStart, markerMask)
  
End Function

'SCI_MARKERLINEFROMHANDLE(int handle)
Public Function MARKERLINEFROMHANDLE(handle As Long) As Long
  
  MARKERLINEFROMHANDLE = Send_SCI_message(SCI_MARKERLINEFROMHANDLE, handle, 0)
  
End Function

'SCI_MARKERDELETEHANDLE(int handle)
Public Function MARKERDELETEHANDLE(handle As Long) As Long
  
  MARKERDELETEHANDLE = Send_SCI_message(SCI_MARKERDELETEHANDLE, handle, 0)
  
End Function




'SCI_INDICSETSTYLE(int indicatorNumber, indicatorStyle)
Public Function INDICSETSTYLE(indicatorNumber As Long, indicatorStyle As Long) As Long
  
  INDICSETSTYLE = Send_SCI_message(SCI_INDICSETSTYLE, indicatorNumber, indicatorStyle)
  
End Function

'SCI_INDICGETSTYLE(int indicatorNumber)
Public Function INDICGETSTYLE(indicatorNumber As Long) As Long
  
  INDICGETSTYLE = Send_SCI_message(SCI_INDICGETSTYLE, indicatorNumber, 0)
  
End Function

'SCI_INDICSETFORE(int indicatorNumber, colour)
Public Function INDICSETFORE(indicatorNumber As Long, colour) As Long
  
  INDICSETFORE = Send_SCI_message(SCI_INDICSETFORE, indicatorNumber, colour)
  
End Function

'SCI_INDICGETFORE(int indicatorNumber)
Public Function INDICGETFORE(indicatorNumber As Long) As Long
  
  INDICGETFORE = Send_SCI_message(SCI_INDICGETFORE, indicatorNumber, 0)
  
End Function


'SCI_AUTOCSHOW(int lenEntered, const char *list)
Public Function AUTOCSHOW(lenEntered As Long, list As String) As Long
  
  AUTOCSHOW = Send_SCI_messageStr(SCI_AUTOCSHOW, lenEntered, list)
  
End Function

'SCI_AUTOCCANCEL
Public Function AUTOCCANCEL() As Long
  
  AUTOCCANCEL = Send_SCI_message(SCI_AUTOCCANCEL, 0, 0)
  
End Function

'SCI_AUTOCACTIVE
Public Function AUTOCACTIVE() As Long
  
  AUTOCACTIVE = Send_SCI_message(SCI_AUTOCACTIVE, 0, 0)
  
End Function

'SCI_AUTOCPOSSTART
Public Function AUTOCPOSSTART() As Long
  
  AUTOCPOSSTART = Send_SCI_message(SCI_AUTOCPOSSTART, 0, 0)
  
End Function

'SCI_AUTOCCOMPLETE
Public Function AUTOCCOMPLETE() As Long
  
  AUTOCCOMPLETE = Send_SCI_message(SCI_AUTOCCOMPLETE, 0, 0)
  
End Function

'SCI_AUTOCSTOPS(<unused>, const char *chars)
Public Function AUTOCSTOPS(chars As String) As Long
  
  AUTOCSTOPS = Send_SCI_messageStr(SCI_AUTOCSTOPS, 0, chars)
  
End Function

'SCI_AUTOCSETSEPARATOR(char separator)
Public Function AUTOCSETSEPARATOR(separator As Byte) As Long
  
  AUTOCSETSEPARATOR = Send_SCI_message(SCI_AUTOCSETSEPARATOR, separator, 0)
  
End Function

'SCI_AUTOCGETSEPARATOR
Public Function AUTOCGETSEPARATOR() As Long
  
  AUTOCGETSEPARATOR = Send_SCI_message(SCI_AUTOCGETSEPARATOR, 0, 0)
  
End Function

'SCI_AUTOCSELECT(<unused>, const char *select)
Public Function AUTOCSELECT(sselect As String) As Long
  
  AUTOCSELECT = Send_SCI_messageStr(SCI_AUTOCSELECT, 0, sselect)
  
End Function

'SCI_AUTOCGETCURRENT
Public Function AUTOCGETCURRENT() As Long
  
  AUTOCGETCURRENT = Send_SCI_message(SCI_AUTOCGETCURRENT, 0, 0)
  
End Function

'SCI_AUTOCSETCANCELATSTART(bool cancel)
Public Function AUTOCSETCANCELATSTART(Cancel As Boolean) As Long
  
  AUTOCSETCANCELATSTART = Send_SCI_message(SCI_AUTOCSETCANCELATSTART, 1 And Cancel, 0)
  
End Function

'SCI_AUTOCGETCANCELATSTART
Public Function AUTOCGETCANCELATSTART() As Boolean
  
  AUTOCGETCANCELATSTART = Send_SCI_message(SCI_AUTOCGETCANCELATSTART, 0, 0)
  
End Function

'SCI_AUTOCSETFILLUPS(<unused>, const char *chars)
Public Function AUTOCSETFILLUPS(chars As String) As Long
  
  AUTOCSETFILLUPS = Send_SCI_messageStr(SCI_AUTOCSETFILLUPS, 0, chars)
  
End Function

'SCI_AUTOCSETCHOOSESINGLE(bool chooseSingle)
Public Function AUTOCSETCHOOSESINGLE(chooseSingle As Boolean) As Long
  
  AUTOCSETCHOOSESINGLE = Send_SCI_message(SCI_AUTOCSETCHOOSESINGLE, 1 And chooseSingle, 0)
  
End Function

'SCI_AUTOCGETCHOOSESINGLE
Public Function AUTOCGETCHOOSESINGLE() As Boolean
  
  AUTOCGETCHOOSESINGLE = Send_SCI_message(SCI_AUTOCGETCHOOSESINGLE, 0, 0)
  
End Function

'SCI_AUTOCSETIGNORECASE(bool ignoreCase)
Public Function AUTOCSETIGNORECASE(ignoreCase As Boolean) As Long
  
  AUTOCSETIGNORECASE = Send_SCI_message(SCI_AUTOCSETIGNORECASE, 1 And ignoreCase, 0)
  
End Function

'SCI_AUTOCGETIGNORECASE
Public Function AUTOCGETIGNORECASE() As Long
  
  AUTOCGETIGNORECASE = Send_SCI_message(SCI_AUTOCGETIGNORECASE, 0, 0)
  
End Function

'SCI_AUTOCSETAUTOHIDE(bool autoHide)
Public Function AUTOCSETAUTOHIDE(autoHide As Boolean) As Long
  
  AUTOCSETAUTOHIDE = Send_SCI_message(SCI_AUTOCSETAUTOHIDE, 1 And autoHide, 0)
  
End Function

'SCI_AUTOCGETAUTOHIDE
Public Function AUTOCGETAUTOHIDE() As Boolean
  
  AUTOCGETAUTOHIDE = Send_SCI_message(SCI_AUTOCGETAUTOHIDE, 0, 0)
  
End Function

'SCI_AUTOCSETDROPRESTOFWORD(bool dropRestOfWord)
Public Function AUTOCSETDROPRESTOFWORD(dropRestOfWord As Boolean) As Long
  
  AUTOCSETDROPRESTOFWORD = Send_SCI_message(SCI_AUTOCSETDROPRESTOFWORD, 1 And dropRestOfWord, 0)
  
End Function

'SCI_AUTOCGETDROPRESTOFWORD
Public Function AUTOCGETDROPRESTOFWORD() As Boolean
  
  AUTOCGETDROPRESTOFWORD = Send_SCI_message(SCI_AUTOCGETDROPRESTOFWORD, 0, 0)
  
End Function

'SCI_REGISTERIMAGE
Public Function REGISTERIMAGE() As Long
  
  REGISTERIMAGE = Send_SCI_message(SCI_REGISTERIMAGE, 0, 0)
  
End Function

'SCI_CLEARREGISTEREDIMAGES
Public Function CLEARREGISTEREDIMAGES() As Long
  
  CLEARREGISTEREDIMAGES = Send_SCI_message(SCI_CLEARREGISTEREDIMAGES, 0, 0)
  
End Function

'SCI_AUTOCSETTYPESEPARATOR(char separatorCharacter)
Public Function AUTOCSETTYPESEPARATOR(separatorCharacter As Byte) As Long
  
  AUTOCSETTYPESEPARATOR = Send_SCI_message(SCI_AUTOCSETTYPESEPARATOR, separatorCharacter, 0)
  
End Function

'SCI_AUTOCGETTYPESEPARATOR
Public Function AUTOCGETTYPESEPARATOR() As Long
  
  AUTOCGETTYPESEPARATOR = Send_SCI_message(SCI_AUTOCGETTYPESEPARATOR, 0, 0)
  
End Function

'SCI_USERLISTSHOW(int listType, const char *list)
Public Function USERLISTSHOW(listType As Long, list As String) As Long
  
  USERLISTSHOW = Send_SCI_messageStr(SCI_USERLISTSHOW, listType, list)
  
End Function

'SCI_CALLTIPSHOW(int posStart, const char *definition)
Public Function CALLTIPSHOW(posStart As Long, definition As String) As Long
  
  CALLTIPSHOW = Send_SCI_messageStr(SCI_CALLTIPSHOW, posStart, definition)
  
End Function

'SCI_CALLTIPCANCEL
Public Function CALLTIPCANCEL() As Long
  
  CALLTIPCANCEL = Send_SCI_message(SCI_CALLTIPCANCEL, 0, 0)
  
End Function

'SCI_CALLTIPACTIVE
Public Function CALLTIPACTIVE() As Long
  
  CALLTIPACTIVE = Send_SCI_message(SCI_CALLTIPACTIVE, 0, 0)
  
End Function

'SCI_CALLTIPPOSSTART
Public Function CALLTIPPOSSTART() As Long
  
  CALLTIPPOSSTART = Send_SCI_message(SCI_CALLTIPPOSSTART, 0, 0)
  
End Function

'SCI_CALLTIPSETHLT(int highlightStart, highlightEnd)
Public Function CALLTIPSETHLT(highlightStart As Long, highlightEnd As Long) As Long
  
  CALLTIPSETHLT = Send_SCI_message(SCI_CALLTIPSETHLT, highlightStart, highlightEnd)
  
End Function

'SCI_CALLTIPSETBACK(int colour)
Public Function CALLTIPSETBACK(colour As Long) As Long
  
  CALLTIPSETBACK = Send_SCI_message(SCI_CALLTIPSETBACK, colour, 0)
  
End Function

'SCI_CALLTIPSETFORE(int colour)
Public Function CALLTIPSETFORE(colour As Long) As Long
  
  CALLTIPSETFORE = Send_SCI_message(SCI_CALLTIPSETFORE, colour, 0)
  
End Function

'SCI_CALLTIPSETFOREHLT(int colour)
Public Function CALLTIPSETFOREHLT(colour As Long) As Long
  
  CALLTIPSETFOREHLT = Send_SCI_message(SCI_CALLTIPSETFOREHLT, colour, 0)
  
End Function


'SCI_ASSIGNCMDKEY(int keyDefinition, sciCommand)
Public Function ASSIGNCMDKEY(keyDefinition As Long, sciCommand As Long) As Long
  
  ASSIGNCMDKEY = Send_SCI_message(SCI_ASSIGNCMDKEY, keyDefinition, sciCommand)
  
End Function

'SCI_CLEARCMDKEY(int keyDefinition)
Public Function CLEARCMDKEY(keyDefinition As Long) As Long
  
  CLEARCMDKEY = Send_SCI_message(SCI_CLEARCMDKEY, keyDefinition, 0)
  
End Function

'SCI_CLEARALLCMDKEYS
Public Function CLEARALLCMDKEYS() As Long
  
  CLEARALLCMDKEYS = Send_SCI_message(SCI_CLEARALLCMDKEYS, 0, 0)
  
End Function

'SCI_NULL
Public Function DoNull() As Long
  
  DoNull = Send_SCI_message(SCI_NULL, 0, 0)
  
End Function

'SCI_USEPOPUP(bool bEnablePopup)
Public Function USEPOPUP(bEnablePopup As Boolean) As Long
  
  USEPOPUP = Send_SCI_message(SCI_USEPOPUP, 1 And bEnablePopup, 0)
  
End Function

'SCI_STARTRECORD
Public Function STARTRECORD() As Long
  
  STARTRECORD = Send_SCI_message(SCI_STARTRECORD, 0, 0)
  
End Function

'SCI_STOPRECORD
Public Function STOPRECORD() As Long
  
  STOPRECORD = Send_SCI_message(SCI_STOPRECORD, 0, 0)
  
End Function

'SCI_FORMATRANGE(bool bDraw, RangeToFormat *pfr)
Public Function FORMATRANGE(bDraw As Boolean, RangeToF As Long) As Long
  
  FORMATRANGE = Send_SCI_message(SCI_FORMATRANGE, 1 And bDraw, RangeToF)
  
End Function

'SCI_SETPRINTMAGNIFICATION(int magnification)
Public Property Let PRINTMAGNIFICATION(magnification As Long)
  
  Send_SCI_message SCI_SETPRINTMAGNIFICATION, magnification, 0
  
End Property

'SCI_GETPRINTMAGNIFICATION
Public Property Get PRINTMAGNIFICATION() As Long
Attribute PRINTMAGNIFICATION.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  PRINTMAGNIFICATION = Send_SCI_message(SCI_GETPRINTMAGNIFICATION, 0, 0)
  
End Property

'SCI_SETPRINTCOLOURMODE(int mode)
Public Property Let PRINTCOLOURMODE(mode As Long)
  
  Send_SCI_message SCI_SETPRINTCOLOURMODE, mode, 0
  
End Property

'SCI_GETPRINTCOLOURMODE
Public Property Get PRINTCOLOURMODE() As Long
Attribute PRINTCOLOURMODE.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  PRINTCOLOURMODE = Send_SCI_message(SCI_GETPRINTCOLOURMODE, 0, 0)
  
End Property

'SCI_GETPRINTWRAPMODE
Public Property Get PRINTWRAPMODE() As Long
Attribute PRINTWRAPMODE.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  PRINTWRAPMODE = Send_SCI_message(SCI_GETPRINTWRAPMODE, 0, 0)
  
End Property

'SCI_SETPRINTWRAPMODE(int wrapMode)
Public Property Let PRINTWRAPMODE(wrapMode As Long)
  
  Send_SCI_message SCI_SETPRINTWRAPMODE, wrapMode, 0
  
End Property

'SCI_GETDIRECTFUNCTION
Public Property Get DIRECTFUNCTION() As Long
  
  DIRECTFUNCTION = Send_SCI_message(SCI_GETDIRECTFUNCTION, 0, 0)
  
End Property

'SCI_GETDIRECTPOINTER
Public Property Get DIRECTPOINTER() As Long
  
  DIRECTPOINTER = Send_SCI_message(SCI_GETDIRECTPOINTER, 0, 0)
  
End Property

'SCI_GETDOCPOINTER
Public Property Get DOCPOINTER() As Long
Attribute DOCPOINTER.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  DOCPOINTER = Send_SCI_message(SCI_GETDOCPOINTER, 0, 0)
  
End Property

'SCI_SETDOCPOINTER(<unused>, document *pDoc)
Public Property Let DOCPOINTER(pDoc As Long)
  
  Send_SCI_message SCI_SETDOCPOINTER, 0, pDoc
  
End Property

'SCI_CREATEDOCUMENT
Public Function CREATEDOCUMENT() As Long
  
  CREATEDOCUMENT = Send_SCI_message(SCI_CREATEDOCUMENT, 0, 0)
  
End Function

'SCI_ADDREFDOCUMENT(<unused>, document *pDoc)
Public Function ADDREFDOCUMENT(pDoc As Long) As Long
  
  ADDREFDOCUMENT = Send_SCI_message(SCI_ADDREFDOCUMENT, 0, pDoc)
  
End Function

'SCI_RELEASEDOCUMENT(<unused>, document *pDoc)
Public Function RELEASEDOCUMENT(pDoc As Long) As Long
  
  RELEASEDOCUMENT = Send_SCI_message(SCI_RELEASEDOCUMENT, 0, pDoc)
  
End Function

'SCI_VISIBLEFROMDOCLINE(int docLine)
Public Function VISIBLEFROMDOCLINE(docLine As Long) As Long
  
  VISIBLEFROMDOCLINE = Send_SCI_message(SCI_VISIBLEFROMDOCLINE, docLine, 0)
  
End Function

'SCI_DOCLINEFROMVISIBLE(int displayLine)
Public Function DOCLINEFROMVISIBLE(displayLine As Long) As Long
  
  DOCLINEFROMVISIBLE = Send_SCI_message(SCI_DOCLINEFROMVISIBLE, displayLine, 0)
  
End Function

'SCI_SHOWLINES(int lineStart, lineEnd)
Public Function SHOWLINES(lineStart As Long, lineEnd As Long) As Long
  
  SHOWLINES = Send_SCI_message(SCI_SHOWLINES, lineStart, lineEnd)
  
End Function

'SCI_HIDELINES(int lineStart, lineEnd)
Public Function HIDELINES(lineStart As Long, lineEnd As Long) As Long
  
  HIDELINES = Send_SCI_message(SCI_HIDELINES, lineStart, lineEnd)
  
End Function

'SCI_GETLINEVISIBLE(int line)
Public Property Get LINEVISIBLE(line As Long) As Long
  
  LINEVISIBLE = Send_SCI_message(SCI_GETLINEVISIBLE, line, 0)
  
End Property

'SCI_SETFOLDLEVEL(int line, level)
Public Property Let FOLDLEVEL(line As Long, level As Long)
  
  Send_SCI_message SCI_SETFOLDLEVEL, line, level
  
End Property

'SCI_GETFOLDLEVEL(int line)
Public Property Get FOLDLEVEL(line As Long) As Long
Attribute FOLDLEVEL.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  FOLDLEVEL = Send_SCI_message(SCI_GETFOLDLEVEL, line, 0)
  
End Property

'SCI_SETFOLDFLAGS(int flags)
Public Property Let FOLDFLAGS(flags As Long)
  
  Send_SCI_message SCI_SETFOLDFLAGS, flags, 0
  
End Property





'SCI_SETFOLDEXPANDED(int line, bool expanded)
Public Property Let FOLDEXPANDED(line As Long, expanded As Boolean)
  
  Send_SCI_message SCI_SETFOLDEXPANDED, line, 1 And expanded
  
End Property

'SCI_GETFOLDEXPANDED(int line)
Public Property Get FOLDEXPANDED(line As Long) As Boolean
Attribute FOLDEXPANDED.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  FOLDEXPANDED = Send_SCI_message(SCI_GETFOLDEXPANDED, line, 0)
  
End Property

'SCI_TOGGLEFOLD(int line)
Public Function TOGGLEFOLD(line As Long) As Long
  
  TOGGLEFOLD = Send_SCI_message(SCI_TOGGLEFOLD, line, 0)
  
End Function

'SCI_ENSUREVISIBLE(int line)
Public Function ENSUREVISIBLE(line As Long) As Long
  
  ENSUREVISIBLE = Send_SCI_message(SCI_ENSUREVISIBLE, line, 0)
  
End Function

'SCI_ENSUREVISIBLEENFORCEPOLICY(int line)
Public Function ENSUREVISIBLEENFORCEPOLICY(line As Long) As Long
  
  ENSUREVISIBLEENFORCEPOLICY = Send_SCI_message(SCI_ENSUREVISIBLEENFORCEPOLICY, line, 0)
  
End Function

'SCI_GETLASTCHILD(int startLine, level)
Public Property Get LASTCHILD(startLine As Long, level As Long) As Long
  
  LASTCHILD = Send_SCI_message(SCI_GETLASTCHILD, startLine, level)
  
End Property

'SCI_GETFOLDPARENT(int startLine)
Public Property Get FOLDPARENT(startLine As Long) As Long
  
  FOLDPARENT = Send_SCI_message(SCI_GETFOLDPARENT, startLine, 0)
  
End Property

'SCI_SETWRAPMODE(int wrapMode)
Public Property Let wrapMode(wrapMode As Long)
  
  Send_SCI_message SCI_SETWRAPMODE, wrapMode, 0
  
End Property

'SCI_GETWRAPMODE
Public Property Get wrapMode() As Long
Attribute wrapMode.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  wrapMode = Send_SCI_message(SCI_GETWRAPMODE, 0, 0)
  
End Property

'SCI_SETWRAPVISUALFLAGS(int wrapVisualFlags)
Public Property Let wrapVisualFlags(wrapVisualFlags As Long)
  
  Send_SCI_message SCI_SETWRAPVISUALFLAGS, wrapVisualFlags, 0
  
End Property

'SCI_GETWRAPVISUALFLAGS
Public Property Get wrapVisualFlags() As Long
Attribute wrapVisualFlags.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  wrapVisualFlags = Send_SCI_message(SCI_GETWRAPVISUALFLAGS, 0, 0)
  
End Property

'SCI_SETWRAPSTARTINDENT(int indent)
Public Property Let WRAPSTARTINDENT(indent As Long)
  
  Send_SCI_message SCI_SETWRAPSTARTINDENT, indent, 0
  
End Property

'SCI_GETWRAPSTARTINDENT
Public Property Get WRAPSTARTINDENT() As Long
Attribute WRAPSTARTINDENT.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  WRAPSTARTINDENT = Send_SCI_message(SCI_GETWRAPSTARTINDENT, 0, 0)
  
End Property

'SCI_SETLAYOUTCACHE(int cacheMode)
Public Property Let LAYOUTCACHE(cacheMode As Long)
  
  Send_SCI_message SCI_SETLAYOUTCACHE, cacheMode, 0
  
End Property

'SCI_GETLAYOUTCACHE
Public Property Get LAYOUTCACHE() As Long
Attribute LAYOUTCACHE.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  LAYOUTCACHE = Send_SCI_message(SCI_GETLAYOUTCACHE, 0, 0)
  
End Property


'SCI_LINESJOIN
Public Function LINESJOIN() As Long
  
  LINESJOIN = Send_SCI_message(SCI_LINESJOIN, 0, 0)
  
End Function

'SCI_SETWRAPVISUALFLAGSLOCATION(int wrapVisualFlagsLocation)
Public Property Let wrapVisualFlagsLocation(wrapVisualFlagsLocation As Long)
  
  Send_SCI_message SCI_SETWRAPVISUALFLAGSLOCATION, wrapVisualFlagsLocation, 0
  
End Property

'SCI_GETWRAPVISUALFLAGSLOCATION
Public Property Get wrapVisualFlagsLocation() As Long
Attribute wrapVisualFlagsLocation.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  wrapVisualFlagsLocation = Send_SCI_message(SCI_GETWRAPVISUALFLAGSLOCATION, 0, 0)
  
End Property

'SCI_LINESSPLIT(int pixelWidth)
Public Function LINESSPLIT(pixelWidth As Long) As Long
  
  LINESSPLIT = Send_SCI_message(SCI_LINESSPLIT, pixelWidth, 0)
  
End Function

'SCI_ZOOMIN
Public Function ZOOMIN() As Long
  
  ZOOMIN = Send_SCI_message(SCI_ZOOMIN, 0, 0)
  
End Function

'SCI_ZOOMOUT
Public Function ZOOMOUT() As Long
  
  ZOOMOUT = Send_SCI_message(SCI_ZOOMOUT, 0, 0)
  
End Function

'SCI_SETZOOM(int zoomInPoints)
Public Property Let ZOOM(zoomInPoints As Long)
  
  Send_SCI_message SCI_SETZOOM, zoomInPoints, 0
  
End Property

'SCI_GETZOOM
Public Property Get ZOOM() As Long
Attribute ZOOM.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  ZOOM = Send_SCI_message(SCI_GETZOOM, 0, 0)
  
End Property



'SCI_GETEDGEMODE
Public Property Get edgeMode() As Long
Attribute edgeMode.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  edgeMode = Send_SCI_message(SCI_GETEDGEMODE, 0, 0)
  
End Property

'SCI_SETEDGECOLUMN(int column)
Public Property Let EDGECOLUMN(column As Long)
  
  Send_SCI_message SCI_SETEDGECOLUMN, column, 0
  
End Property

'SCI_GETEDGECOLUMN
Public Property Get EDGECOLUMN() As Long
Attribute EDGECOLUMN.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  EDGECOLUMN = Send_SCI_message(SCI_GETEDGECOLUMN, 0, 0)
  
End Property

'SCI_SETEDGECOLOUR(int colour)
Public Property Let EDGECOLOUR(colour As Long)
  
  Send_SCI_message SCI_SETEDGECOLOUR, colour, 0
  
End Property

'SCI_GETEDGECOLOUR
Public Property Get EDGECOLOUR() As Long
Attribute EDGECOLOUR.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  EDGECOLOUR = Send_SCI_message(SCI_GETEDGECOLOUR, 0, 0)
  
End Property

'SCI_SETEDGEMODE(int edgeMode)
Public Property Let edgeMode(edgeMode As Long)
  
  Send_SCI_message SCI_SETEDGEMODE, edgeMode, 0
  
End Property

'SCI_SETLEXER(int lexer)
Public Property Let Lexer(llexer As SCLex)
  
  Send_SCI_message SCI_SETLEXER, llexer, 0
  
End Property

'SCI_GETLEXER
Public Property Get Lexer() As SCLex
  
  Lexer = Send_SCI_message(SCI_GETLEXER, 0, 0)
  
End Property




'SCI_COLOURISE(int start, end)
Public Function COLOURISE(start As Long, ends As Long) As Long
  
  COLOURISE = Send_SCI_message(SCI_COLOURISE, start, ends)
  
End Function

'SCI_SETPROPERTY(const char *key, const char *value)
Public Property Let PROPERTY(key As String, value As String)
  
  Send_SCI_messageStr SCI_SETPROPERTY, StrPtr(StrConv(key, vbFromUnicode)), value
  
End Property

'SCI_SETKEYWORDS(int keyWordSet, const char *keyWordList)
Public Property Let KEYWORDS(keyWordSet As Long, keyWordList As String)
  
  Send_SCI_messageStr SCI_SETKEYWORDS, keyWordSet, keyWordList
  
End Property

'SCI_SETLEXERLANGUAGE(<unused>, const char *name)
Public Property Let LEXERLANGUAGE(name As String)
  
  Send_SCI_messageStr SCI_SETLEXERLANGUAGE, 0, name
  
End Property

'SCI_LOADLEXERLIBRARY(<unused>, const char *path)
Public Function LOADLEXERLIBRARY(path As String) As Long
  
  LOADLEXERLIBRARY = Send_SCI_messageStr(SCI_LOADLEXERLIBRARY, 0, path)
  
End Function


'SCI_SETMODEVENTMASK(int eventMask)
Public Property Let MODEVENTMASK(eventMask As Long)
  
  Send_SCI_message SCI_SETMODEVENTMASK, eventMask, 0
  
End Property

'SCI_GETMODEVENTMASK
Public Property Get MODEVENTMASK() As Long
Attribute MODEVENTMASK.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  MODEVENTMASK = Send_SCI_message(SCI_GETMODEVENTMASK, 0, 0)
  
End Property

'SCI_SETMOUSEDWELLTIME
Public Property Let MOUSEDWELLTIME(tm As Long)
  
  Send_SCI_message SCI_SETMOUSEDWELLTIME, tm, 0

End Property

'SCI_GETMOUSEDWELLTIME
Public Property Get MOUSEDWELLTIME() As Long
Attribute MOUSEDWELLTIME.VB_ProcData.VB_Invoke_Property = "MainSettings"
  
  MOUSEDWELLTIME = Send_SCI_message(SCI_GETMOUSEDWELLTIME, 0, 0)
  
End Property



Private Sub WriteToPropBag(toB As PropertyBag, data As Variant)

    toB.WriteProperty CStr(lppw), data
    lppw = lppw + 1

End Sub

Private Function ReadfromPropBag(toB As PropertyBag, Optional def As Variant = "none") As Variant
    
    If def = "none" Then
        ReadfromPropBag = toB.ReadProperty(CStr(lppw), "")
    Else
        ReadfromPropBag = toB.ReadProperty(CStr(lppw), def)
    End If
    lppw = lppw + 1

End Function
