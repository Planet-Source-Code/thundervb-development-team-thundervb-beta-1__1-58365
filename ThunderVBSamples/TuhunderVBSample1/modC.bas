Attribute VB_Name = "modC"
Option Explicit

'here we tell the assembler to use scoped labels..
'this is needed because the C compiler generates code that needs to be
'scoped in order to assemble
Private Function AsmSettings()
'#asm' OPTION SCOPED
End Function

'Detect if c is on
Public Function IsCEnabled() As Boolean
'#c'int nnn(){
'#c'return -1;
'#c'}
End Function

'Simple call function by pointer
Public Function CallBP(ByVal p1 As Long) As Long
'#c'int CallBP(int (__stdcall*fn)(int)){
'#c'if (fn==0) return 0;
'#c'return fn(0);
'#c'}
End Function

'Simple call function by pointer , a cdecl function this time ;)
'This i just to show you how this can be done , it is not used by the demo ;)
Public Function CallBP_cdecl(ByVal p1 As Long) As Long
'#c'int CallBP_cdecl(int (__cdecl*fn)()){
'#c'if (fn==0) return 0;
'#c'return fn();
'#c'}
End Function

'Here we demonstrate how a vb function cal inlude many c functions..
'The first one is called by VB as Callcdecl
'Also , here we call a cdecl function from C showing how we can call cdecl functions ;)
Public Function Callcdecl() As Long
'#c'int mmm(){
'#c'return Callmycfunct();
'#c'}
'#c'
'#c'int __cdecl mycfunct(){
'#c'return 100;
'#c'}
'#c'
'#c'int Callmycfunct(){
'#c'return mycfunct();
'#c'}
End Function


'Calls a VB function from C code
'Note , some code must be writen on an other module to alow this..
'Look on modAsm, asmAliases sub
Public Function CallVBF() As Long
'#c'int CallBP_cdecl(){
'#c'CalledByPtr(1);
'#c'}
End Function

