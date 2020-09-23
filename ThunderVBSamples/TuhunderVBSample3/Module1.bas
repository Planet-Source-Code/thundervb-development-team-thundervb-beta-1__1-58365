Attribute VB_Name = "Module1"
'Sample project
'Asm code can be freely mixed with vb code
'but it is sugested not to use asm on big modules/forms cause vb6 compiler
'generates code using unamed vars witch are not outputed corectly on the asm listing
'The listing can be fixed automaticaly , but this is still in beta stage..
'DAMIT , VB6 Compiler is full with BUGS even the SP6 release...
'


Public Function MemCopyMMX_PreFetch(ByVal Dest As Long, ByVal Source As Long, ByVal ln As Long) As Long
''MMX bits from SGI web site
'ASM START
    '#asm' Option Scoped
    '#asm' .MMX
    '#asm' .XMM
    '#asm'  push ebp ;save registers
    '#asm'  mov ebp, esp
    '#asm'  push ebx
    '#asm'  push esi
    '#asm'  push edi
    '#asm'  mov eax, [ebp+16] ; put byte count in eax
    '#asm'  mov esi, [ebp+12] ; copy source pointer into source index
    '#asm'  mov edi, [ebp+8] ; copy dest pointer into destination index
    '#asm'  cld ; copy bytes forward
    '#asm'  cmp eax, 64 ; if under 64 bytes long
    '#asm'  jl Under64PreFetch ; jump
    '#asm'  push eax ; place a copy of eax on the stack
    '#asm'  shr eax, 6 ; integer divide eax by 64
    '#asm'  shl eax, 6 ; multiply eax by 64 to get dividend
    '#asm'  mov ecx, eax ; copy it into variable
    '#asm'  pop eax ; retrieve length in eax off the stack
    '#asm'  sub eax, ecx ; subtract dividend from length to get remainder
    '#asm'  mov ebx, eax ; copy remainder into variable
    '#asm'  shr ecx, 6 ; divide by 64 for DWORD data size
    '#asm'  ;shr ecx, 6 ;// 64 bytes per iteration
    '#asm'
    '#asm' loop1PreFetch:
    '#asm'
    '#asm'  prefetchnta 64[ESI] ;// Prefetch next loop, non-temporal
    '#asm'  prefetchnta 96[ESI]
    '#asm'  movq mm1 , 0[ESI] ;// Read in source data
    '#asm'  movq mm2, 8[ESI]
    '#asm'  movq mm3, 16[ESI]
    '#asm'  movq mm4, 24[ESI]
    '#asm'  movq mm5, 32[ESI]
    '#asm'  movq mm6, 40[ESI]
    '#asm'  movq mm7, 48[ESI]
    '#asm'  movq mm0, 56[ESI]
    '#asm'  movntq 0[EDI], mm1 ;// Non-temporal stores
    '#asm'  movntq 8[EDI], mm2
    '#asm'  movntq 16[EDI], mm3
    '#asm'  movntq 24[EDI], mm4
    '#asm'  movntq 32[EDI], mm5
    '#asm'  movntq 40[EDI], mm6
    '#asm'  movntq 48[EDI], mm7
    '#asm'  movntq 56[EDI], mm0
    '#asm'  Add esi, 64
    '#asm'  Add edi, 64
    '#asm'  dec ecx
    '#asm'  jnz loop1PreFetch
    '#asm'   fdiv
    '#asm'  emms
    '#asm'  mov eax, ebx ; put remainder in ecx
    '#asm'  Under64PreFetch:
    '#asm'  push eax ; place a copy of eax on the stack
    '#asm'  shr eax, 2 ; integer divide eax by 4
    '#asm'  shl eax, 2 ; multiply eax by 4 to get dividend
    '#asm'  mov ecx, eax ; copy it into variable
    '#asm'  pop eax ; retrieve length in eax off the stack
    '#asm'  sub eax, ecx ; subtract dividend from length to get remainder
    '#asm'  mov ebx, eax ; copy remainder into variable
    '#asm'  shr ecx, 2 ; divide by 4 for DWORD data size
    '#asm'  rep movsd ; repeat while not zero, move string DWORD
    '#asm'  mov ecx, ebx ; put remainder in ecx
    '#asm'  Under4PreFetch:
    '#asm'  rep movsb ; copy remaining BYTES from source to dest
    '#asm'
    '#asm'  pop edi
    '#asm'  pop esi
    '#asm'  pop ebx
    '#asm'  mov esp, ebp
    '#asm'  pop ebp
    '#asm'  ret 12
    '#asm'
'ASM END
End Function

'You CAN'T mix vb6 code with C code...
'It is just imposible .. [at least for now and the near future]
'Also , i don't know how if it is possible to use extenral libs (like the std C lib)
'ect..
'Also , everything writen on a single Procedure is compiled as a single C file
'C code on one procedure can't call another C procedure defined on an other procedure..
'You can however write more that one C procedures on a single VB procedure
'With come "tricks" you can make c/asm code to call vb/asm/c code ..
'Look on ThunderVBSample1 for more info on how to do this ;)

'here we kill some bytes at the start of the buffer.. ohh well i think we will survive ;)
'If you don't have a c compiler then delete any line containing '#c'
Public Function cpy(ByVal from As Long, ByVal toa As Long, ByVal count As Long) As Long
'#c'int cpy(char* from,char* tar,int i2){
'#c'int* p=(int*)((int)from & 0xFFFFFFF0);
'#c'int* p2=(int*)((int)tar & 0xFFFFFFF0); // a comment
'#c'int i3;
'#c'
'#c'i2=(i2>>2)+1;
'#c'
'#c'for(i3=0;i3<i2;i3+=8)
'#c'{
'#c'p[i3]=p2[i3];
'#c'p[i3+1]=p2[i3+1];
'#c'p[i3+2]=p2[i3+2];
'#c'p[i3+3]=p2[i3+3];
'#c'p[i3+4]=p2[i3+4];
'#c'p[i3+5]=p2[i3+5];
'#c'p[i3+6]=p2[i3+6];
'#c'p[i3+7]=p2[i3+7];
'#c'}
'#c'return (i2-1)<<2;
'#c'}
End Function


