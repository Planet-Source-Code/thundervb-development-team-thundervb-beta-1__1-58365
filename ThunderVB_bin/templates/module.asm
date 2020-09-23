function in Module
;long1 equ[ebp+8]
;long2 equ[ebp+12]

push ebp
mov ebp, esp

mov esp, ebp
pop ebp

ret 8