Attribute VB_Name = "modSciAsm"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

'Asm-C functions for The sciEdit control ..
'The control compiles with or without inlineAsm/c enabled but if enabled it is much faster
Public Function DebugerExeption() As Long
    
    '#asm'  int 3
    DebugerExeption = 0
    
End Function

Public Function CallBP(ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long, ByVal p5 As Long) As Long
'#c'int CallBP(int p1,int p2,int p3,int p4,int (__cdecl*fn)(int,int,int,int)){
'#c'if (fn==0) return 0;
'#c'return fn(p1,p2,p3,p4);
'#c'}
End Function

Public Function CCode() As Boolean
'#c'int CCode(){
'#c'return 1 ;
'#c'}
End Function
