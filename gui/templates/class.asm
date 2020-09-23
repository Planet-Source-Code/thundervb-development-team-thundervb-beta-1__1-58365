function in Class
push ebp
mov ebp, esp

push ebx
push esi
push edi

mov  DWORD PTR _ClsFunc2$[ebp], 12345678  ;return this number

mov eax, DWORD PTR _ClsFunc1$[ebp]
mov ecx, DWORD PTR _ClsFunc2$[ebp]
mov DWORD PTR [eax], ecx

pop edi
pop esi
pop ebx

mov esp, ebp
pop ebp

ret 8
