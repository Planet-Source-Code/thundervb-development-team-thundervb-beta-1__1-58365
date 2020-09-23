Attribute VB_Name = "modAsm"
Option Explicit

'NOTE :
'Here we define some aliases so that c code can call vb functions..
'This code must be on an asm/c olny code module , diferent than the one that the
'function that we want to call is defined and dieferent than the module that
'we call teh function
'Eg here , The function is defined on modOther and called from modC , while this i modAsm..

'What we do?? we create a C decorated name for the VB decorated functions
'for each function that you want to do this you must
'write the folowing
'#asm' EXTRN   ?<functioname>@<modulename>@@AAGXXZ:NEAR
'#asm' PUBLIC  _<functioname>@<size_of_params_in_bytes>
'#asm' _<functioname>@<size_of_params_in_bytes>: jmp ?<functioname>@<modulename>@@AAGXXZ
'Where :
'<functionname> is the name of the function that you want to call
'<modulename> is the module that this function is writen
'<size_of_params_in_bytes> is the size of the parameters that the function takes
'Here is a table to help you with the calulations..
'VB      C     size
'Long    Int   4
'Integer Short 2
'byte    char  1
'Byref   *     4
'
'All params are passed byref by default [if no ByVal is defined]
'Strings ect must be converted manualy..
'

Private Sub AsmAliases()
'#asm' EXTRN   ?CalledByPtr@modOther@@AAGXXZ:NEAR
'#asm' PUBLIC  _CalledByPtr@4
'#asm' _CalledByPtr@4: jmp ?CalledByPtr@modOther@@AAGXXZ
End Sub

'Detect if asm is on
Public Function IsAsmEnabled() As Boolean
'#asm' mov eax,-1
'#asm' ret
End Function
