Attribute VB_Name = "modGPFException"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit

'This code is not  mine..
'It is just exteneded a bit by me [drkIIRaziel]
' -------------------------------------------------------------- '
' module to handle unhandled exceptions (GPFs)
' created 25/11/02
' modified  10/12/02
' will barden
'
' 10/12/02 - added a more descriptive error message, and
'            setup the error handler to use VBs's internal
'            error bubbling to raise it.
' 5/10/2004 -Modifyed for ThunVB by Raziel
' -------------------------------------------------------------- '
'
' -------------------------------------------------------------- '
' apis
' -------------------------------------------------------------- '

' used to set and remove our callback
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long

' to raise a GPF (for testing)
Private Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

' to get the last GPF code
Private Declare Function GetExceptionInformation Lib "kernel32" () As Long

' -------------------------------------------------------------- '
' consts
' -------------------------------------------------------------- '

' return values from our callback
Public Const EXCEPTION_CONTINUE_EXECUTION = -1
Public Const EXCEPTION_CONTINUE_SEARCH = 0

' length field in the EXCEPTION_RECORD struct
Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15

' to describe the violation - defined in windows.h
Public Const EXCEPTION_CONTINUABLE              As Long = &H0
Public Const EXCEPTION_NONCONTINUABLE           As Long = &H1

Public Const EXCEPTION_ACCESS_VIOLATION         As Long = &HC0000005 ' The thread tried to read from or write to a virtual address for which it does not have the appropriate access
Public Const EXCEPTION_BREAKPOINT               As Long = &H80000003 ' A breakpoint was encountered.
Public Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED    As Long = &HC000008C ' The thread tried to access an array element that is out of bounds and the underlying hardware supports bounds checking.
Public Const EXCEPTION_FLT_DIVIDE_BY_ZERO       As Long = &HC000008E ' The thread tried to divide a floating-point value by a floating-point divisor of zero.
Public Const EXCEPTION_FLT_INVALID_OPERATION    As Long = &HC0000090 ' This exception represents any floating-point exception not included in this list
Public Const EXCEPTION_FLT_OVERFLOW             As Long = &HC0000091 ' The exponent of a floating-point operation is greater than the magnitude allowed by the corresponding type
Public Const EXCEPTION_INT_DIVIDE_BY_ZERO       As Long = &HC0000094 ' The thread tried to divide an integer value by an integer divisor of zero.
Public Const EXCEPTION_INT_OVERFLOW             As Long = &HC0000095 ' The result of an integer operation caused a carry out of the most significant bit of the result
Public Const EXCEPTION_ILLEGAL_INSTRUCTION      As Long = &HC000001D ' The thread tried to execute an invalid instruction
Public Const EXCEPTION_PRIV_INSTRUCTION         As Long = &HC0000096 ' The thread tried to execute an instruction whose operation is not allowed in the current machine mode

' -------------------------------------------------------------- '
' structs
' -------------------------------------------------------------- '

' holds info about a specific eception
Public Type EXCEPTION_RECORD
  ExceptionCode      As Long  ' type of exception - defined above
  ExceptionFlags     As Long  ' whether the exception is continuable or not
  pExceptionRecord   As Long  ' pointer to another EXCEPTION_RECORD struct (for nested exceptions)
  ExceptionAddress   As Long  ' the address at which the exception occurred
  NumberParameters   As Long  ' number of params in the following array
  Information(EXCEPTION_MAXIMUM_PARAMETERS - 1) As Long ' extra info.. not really needed.
End Type

' processor specific - not really needed anyway
Public Type CONTEXT
  Null               As Long
End Type

' wrapper for the above types
Public Type EXCEPTION_POINTERS
  pExceptionRecord   As EXCEPTION_RECORD
  ContextRecord      As CONTEXT
End Type

'GPF Interface stuff
Public Enum GPF_actions
    GPF_None
    GPF_RaiseErr
    GPF_Cont
    GPF_Stop
End Enum

Public GPF_action As GPF_actions
Public GPF_CodeProc As String
Public GPF_CodeMod As String
Public GPF_Last_Exeption As EXCEPTION_POINTERS

Public Type gpf_pb_e

    GPF_action As GPF_actions
    GPF_CodeProc As String
    GPF_CodeMod As String
    GPF_Last_Exeption As EXCEPTION_POINTERS

End Type
' -------------------------------------------------------------- '
' private variables
' -------------------------------------------------------------- '
Private mlpOldProc As Long
Private pb() As gpf_pb_e, pbl As Long, pbli As Long
' -------------------------------------------------------------- '
' methods
' -------------------------------------------------------------- '

' setup the new handler
Public Function StartGPFHandler() As Boolean
   LogMsg "Seting up GPF handler", "modGPFException", "StartGPFHandler"
   ' assume success
   StartGPFHandler = True
   
   ' if we're already handling, there's no point
   If mlpOldProc = 0 Then
   
      ' set up the handler
      mlpOldProc = SetUnhandledExceptionFilter(AddressOf ExceptionHandler)
      ' not all systems will return a handle
      If mlpOldProc = 0 Then mlpOldProc = 1
      
   End If
   
End Function

' release the new handler
Public Sub StopGPFHandler()
   LogMsg "killing GPF handler", "modGPFException", "StopGPFHandler"
   ' release the handler
   SetUnhandledExceptionFilter vbNull
   
   ' reset the variable
   mlpOldProc = 0
   
End Sub

' just for debugging - test the handler by firing a GPF
Public Sub TestGPFHandler()

   ' raise a GPF
   RaiseException EXCEPTION_ARRAY_BOUNDS_EXCEEDED, 0, 0, 0
   
End Sub

' altered on 10/12/02 by request - this function now simply raises
' an error so that VB can handle it properly, via On Error.
Public Function ExceptionHandler(ByRef uException As EXCEPTION_POINTERS) As Long
Dim lTmp       As Long
Dim sType      As String
Dim lAddress   As Long
Dim sContinue  As String

   ' let's get some information about the error in order
   ' to raise a nicely defined, and explanatory error via VB
   CopyMemory lTmp, ByVal uException.pExceptionRecord.ExceptionCode, 4
   Select Case lTmp
      Case EXCEPTION_ACCESS_VIOLATION
         sType = "EXCEPTION_ACCESS_VIOLATION"
      Case EXCEPTION_BREAKPOINT
         sType = "EXCEPTION_BREAKPOINT"
      Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
         sType = "EXCEPTION_ARRAY_BOUNDS_EXCEEDED"
      Case EXCEPTION_FLT_DIVIDE_BY_ZERO
         sType = "EXCEPTION_FLT_DIVIDE_BY_ZERO"
      Case EXCEPTION_FLT_INVALID_OPERATION
         sType = "EXCEPTION_FLT_INVALID_OPERATION"
      Case EXCEPTION_FLT_OVERFLOW
         sType = "EXCEPTION_FLT_OVERFLOW"
      Case EXCEPTION_INT_DIVIDE_BY_ZERO
         sType = "EXCEPTION_INT_DIVIDE_BY_ZERO"
      Case EXCEPTION_INT_OVERFLOW
         sType = "EXCEPTION_INT_OVERFLOW"
      Case EXCEPTION_ILLEGAL_INSTRUCTION
         sType = "EXCEPTION_ILLEGAL_INSTRUCTION"
      Case EXCEPTION_PRIV_INSTRUCTION
         sType = "EXCEPTION_PRIV_INSTRUCTION"
      Case Else
         sType = "Unknown exception type 0x" & Hex(uException.pExceptionRecord.ExceptionCode) & _
                 ".Possibly VB6 exeption that was not handled"
   End Select

   ' check for a couple of other important points..
   With uException.pExceptionRecord
      ' can we continue from this error?
      If .ExceptionFlags = EXCEPTION_CONTINUABLE Then
         sContinue = "Ok to continue."
      ElseIf .ExceptionFlags = EXCEPTION_NONCONTINUABLE Then
         sContinue = "NOT ok to continue."
      Else
         sContinue = "Probably safe to continue, but better not."
      End If
      ' and lastly, where the error occurred.
      lAddress = .ExceptionAddress
   End With
   
   GPF_Last_Exeption = uException
    ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
    Select Case GPF_action
        Case GPF_Cont
            LogMsg "An handled GPF error (" & sType & ") " & _
                   "occurred at: " & lAddress & ". " & sContinue, _
                   GPF_CodeMod, GPF_CodeProc
            LogMsg "Trying to continue", GPF_CodeMod, GPF_CodeProc
            ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
            
        Case GPF_actions.GPF_RaiseErr
        LogMsg "An handled GPF error (" & sType & ") " & _
                   "occurred at: " & lAddress & ". " & sContinue, _
                   GPF_CodeMod, GPF_CodeProc
        LogMsg "Raising Error", GPF_CodeMod, GPF_CodeProc
        Err.Raise vbObjectError + 513, _
                 "Exception Handler", _
                 "An unhandled error (" & sType & ") " & vbCrLf & _
                 "occurred at: " & lAddress & ". " & sContinue
                 ' continue with execution
                ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
                
        Case GPF_actions.GPF_Stop
        LogMsg "An handled GPF error (" & sType & ") " & _
                   "occurred at: " & lAddress & ". " & sContinue, _
                   GPF_CodeMod, GPF_CodeProc
            LogMsg "Killing VB proccess", GPF_CodeMod, GPF_CodeProc
            ExceptionHandler = EXCEPTION_CONTINUE_SEARCH
            
        Case GPF_actions.GPF_None
             LogMsg "An unhandled error (" & sType & ") " & _
                    "occurred at: " & lAddress & ". " & sContinue, _
                    "modGPFException", "ExeptionHandler"
             Select Case frmGPFError.ShowGPF("An unhandled error (" & sType & ") " & _
                         "occurred at: " & lAddress & ". " & sContinue)
              Case GPF_actions.GPF_RaiseErr
                  Err.Raise vbObjectError + 513, _
                           "Exception Handler", _
                           "An unhandled error (" & sType & ") " & vbCrLf & _
                           "occurred at: " & lAddress & ". " & sContinue
                           ' continue with execution
                           ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
              Case GPF_actions.GPF_Cont
                  'continue with execution
                  ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
              Case GPF_actions.GPF_Stop
                  'stop execution
                  ExceptionHandler = EXCEPTION_CONTINUE_SEARCH
            End Select
    End Select
    
End Function

'Set the current gpf Hnadling mode
Public Sub GPF_Set(nAct As GPF_actions, fromMod As String, fromProc As String)
Dim gpfnull As EXCEPTION_POINTERS

    If pbli >= pbl Then
        ReDim Preserve pb((pbli + 1) * 2)
        pbl = UBound(pb)
    End If
    
    With pb(pbli)
        .GPF_action = GPF_action
        '.GPF_Last_Exeption = GPF_Last_Exeption
        .GPF_CodeMod = GPF_CodeMod
        .GPF_CodeProc = GPF_CodeProc
    End With
    
    pbli = pbli + 1
    
    GPF_Last_Exeption = gpfnull
    GPF_action = nAct
    GPF_CodeMod = fromMod
    GPF_CodeProc = fromProc
    
End Sub

Public Sub GPF_Reset()
Dim gpfnull As EXCEPTION_POINTERS
    
    If pbli >= 1 Then
        pbli = pbli - 1
        With pb(pbli)
            GPF_Last_Exeption = gpfnull
            GPF_action = .GPF_action
            GPF_CodeMod = .GPF_CodeMod
            GPF_CodeProc = .GPF_CodeProc
        End With
    Else
        GPF_Last_Exeption = gpfnull
        GPF_action = GPF_None
        GPF_CodeMod = ""
        GPF_CodeProc = ""
    End If
    
    'unset the gpf handling data and the VB error handler


End Sub
