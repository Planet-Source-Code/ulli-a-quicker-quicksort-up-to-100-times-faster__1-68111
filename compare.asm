
  use32

	push   edi		; save edi
	push   esi		; save esi
	mov    edi, [esp + 16]	; get 1st param
	mov    edi, [edi]
	mov    esi, [esp + 20]	; get 2nd param
	xor    eax, eax 	; clear return value
	cmp    edi, esi 	; is same string or both nullstring
	je     Equal		; then equal
	cmp    edi, eax 	; is 1st nullstring
	je     Less		; then less
	cmp    esi, eax 	; is second nullstring
	je     Greater		; then greater
	mov    ecx, [esi - 4]	; length of 2nd string in bytes
	cmp    ecx, [edi - 4]	; compare with length of 1st string
	cmovg  ecx, [edi - 4]	; get shorter one for count
	shr    ecx,1		; divide by 2 for unicode
	inc    ecx		; plus 1 for term
	cld			; forward scan
	repe   cmpsw		; compare strings
	je     Equal		; equal - exit
	jl     Less		; [edi] is less: jump

  Greater:
	dec    eax		; [edi] is greater: eax=-2; will be incremented
	dec    eax

  Less:
	inc    eax		; eax = 1 or -1

  Equal:
	pop    esi		; restore esi
	pop    edi		; restore edi
	mov    edx, [esp + 16]
	mov    [edx], eax
	ret    16


