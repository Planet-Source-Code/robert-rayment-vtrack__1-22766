;VTrack2.asm  by Robert Rayment  1/5/01

;VB

;   res& = CallWindowProc(ptMCode&, ptBITDATA&, irand&, 2&, OpCode&)
;                                   8           12      16  20        

;' OpCode& = 0 MC FloorCASTER
;' OpCode& = 1 MC Big Plasmas
;' OpCode& = 2 MC Plus Smoothing

;'=========================================================
;' Structure for input to mcode for RayCASTER & Plasma
;' PS = ProjSurf(), FS = FloorSurf()
; Public Type BITDATA
;    
;   PSWH As Long            ' Width=Height of PS         0 PSWH
;   ptrProjSurf As Long     ' Pointer to PS              4 ptPS
;   FSWH As Long            ' Width=Height of FS         8 FSWH
;   ptrFloorSurf As Long    ' Pointer to FS             12 ptFS
;   
;   LAH As Long             ' Look ahead length L       16 LAH        'Input
;   ex As Long              ' eye x-coord               20 ex         'Input
;   ey As Long              ' eye y-coord               24 ey         'Input
;   TANG As Long            ' Track angle 0-359         28 TANG       'Input
;   Collision As Long       ' 0/1 no/yes collision      32 Collision  'Return
;   
;   zvp As Single           ' eye z-coord               36 zvp        'Input
;   zdslope As Single       ' Ray slope                 40 zdslope    'Input
;   zslopemult As Single    ' Ray slope multiplier      44 zslopemult 'Input
;   zhtmult As Single       ' Objects height multiplier 48 zhtmult    'Input
;	hlim As Long			' Height limiter            52 hlim       'Input
;
; End Type
; Public bmp As BITDATA
;'=========================================================

%macro movab 2		;name & num of parameters
  push dword %2		;2nd param
  pop dword %1		;1st param
%endmacro			;use  movab %1,%2
;Allows eg	movab bmW,[ebx+4]

; INPUT DATA STORE
; Long integers
%define PSWH        [ebp-4]
%define ptPS        [ebp-8]
%define FSWH        [ebp-12]
%define ptFS        [ebp-16]
%define LAH         [ebp-20]
%define ex          [ebp-24]
%define ey          [ebp-28]
%define TANG        [ebp-32]
%define Collision   [ebp-36]

; Singles
%define zvp         [ebp-40]
%define zdslope     [ebp-44]
%define zslopemult  [ebp-48]
%define zhtmult     [ebp-52]
%define hlim        [ebp-56]
; ----------------------------

;	TEMPORARY VARIABLES FOR FLOOR_CASTER
%define i2      [ebp-60]       ; integer 2
%define i360    [ebp-64]       ; integer 360
%define PSWH2   [ebp-68]       ; = PSWH/2

%define zrayang	[ebp-72]       ; =(TANG*2*pi)-PSWH/2
%define xvp     [ebp-76]       ; single x	
%define yvp     [ebp-80]       ; single y
%define C       [ebp-84]       ; Column number

%define xray    [ebp-88]       ; singles
%define yray    [ebp-92]
%define zray    [ebp-96]
%define izray   [ebp-100]       ; integer

%define zdx     [ebp-104]      ; singles 
%define zdy     [ebp-108]
%define zdz     [ebp-112]

%define zvoxscale  [ebp-116]
%define row        [ebp-120]   ; integers

%define ixr     [ebp-124]      
%define iyr     [ebp-128]

%define cul     [ebp-132]      ; color		
%define iht     [ebp-136]      ; height integer

%define Collide [ebp-140]      ; Collision value

%define irand   [ebp-144]      ; Random seed
%define ir255   [ebp-148]      ; Random seed

[bits 32]

	push ebp
	mov ebp,esp
	sub esp,148
	push edi
	push esi
	push ebx

;	LOAD INPUT DATA

	mov ebx,[ebp+8]		; ->BITDATA
	
	movab PSWH,[ebx]
	movab ptPS,[ebx+4]
	movab FSWH,[ebx+8]
	movab ptFS,[ebx+12]
	movab LAH, [ebx+16]
	
	movab ex,       [ebx+20]
	movab ey,       [ebx+24]
	movab TANG,     [ebx+28]
	movab Collision,[ebx+32]
	
	movab zvp,       [ebx+36]
	movab zdslope,   [ebx+40]
	movab zslopemult,[ebx+44]
	movab zhtmult,   [ebx+48]
	movab hlim,      [ebx+52]

; ----------------------------
	; GET OpCode&
	mov eax,[ebp+20]
; ----------------------------
	
	cmp eax,0
	jne TestFor1
	
	mov Collide,eax		; No Collision
	
	CALL FLOOR_CASTER

	jmp GETOUT
	
TestFor1:
	rcr eax,1
	jnc TestFor2
	
	CALL PLASMA
	jmp GETOUT

TestFor2:
	rcr eax,1
	jnc TestFor4

	CALL BYTESMOOTHING
	jmp GETOUT

TestFor4:

GETOUT:
	
	mov eax,irand
	
	mov ecx,ebx
	pop ebx
	pop esi
	pop edi
	mov esp,ebp
	pop ebp
	RET 16

;###########################################
FLOOR_CASTER:
	; Store some constants
	mov eax,2
	mov i2,eax
	mov eax,360
	mov i360,eax
	mov eax,PSWH
	shr eax,1
	mov PSWH2,eax		; PSWH/2
	
	; Calculate zrayang
	fild dword TANG		; TANG
	fild dword i2		; 2,TANG
	fmulp st1			; 2*TANG
	fldpi				; pi,2*TANG
	fmulp st1			; pi*2*TANG
	fild dword PSWH2	; PSWH/2,pi*2*TANG
	fsubp ST1			; st1-st0
	fstp dword zrayang  ; =(TANG*2*pi)-PSWH/2
	
	; Store ex & ey as singles ie xvp, yvp
	fild dword ex		; xvp=ex
	fstp dword xvp
	fild dword ey		; yvp=ey
	fstp dword yvp
	
	; OUTER COLUMN LOOP
	mov ecx,0			; For C = 0 To PSWH - 1

NEXT_COLUMN:

	PUSH ecx
	mov dword C,ecx		; Store C

	; Calc xray,yray,zray
	mov eax,xvp			; xray = xvp
	mov xray,eax
	mov eax,yvp			; yray = yvp
	mov yray,eax
	mov eax,zvp			; zray = zvp
	mov zray,eax

	; Store zray as integer izray
	fld dword zray
	fistp dword izray
		
	; Calc zdy,zdx
	fld dword zrayang	; zrayang
	fild dword C		; C,zrayang
	faddp st1			; zrayang+C
	fild dword i360
	fdivp st1			; ((zrayang+C)/360)
	fsincos				; cos | sin
	fstp dword zdy		; = Cos((zrayang+C)/360)
	fstp dword zdx		; = Sin((zrayang+C)/360)
	
	; Calc zdz
	fild dword i2		; 2
	fld dword zdslope	; zdslope,2
	fld dword zslopemult; zslopemult,zdslope,2 
	fmulp st1			; zslopemult*zdslope,2
	fmulp st1			; zslopemult*zdslope*2
	fstp dword zdz		; zdz = zslopemult*zdslope*2
	
	; Init zvoxscale
	fldz				; 0
	fstp dword zvoxscale; zvoxscale = 0
	
	; Init row
	mov eax,1			; row = 1
	mov row,eax

	; LOOK AHEAD STEPS
	mov ecx,0			; For cstep = 0 To LAH

NEXT_STEP:

	; Find ixr & iyr
	fld dword xray		; xray
	fistp dword ixr		; ixr = Int(xray)
	mov eax,ixr			
	mov edx,FSWH
	mov ebx,17
	sub edx,ebx			; edx = FSWH-17
	and eax,edx
	add eax,9
	mov ixr,eax			; = (xray AND (FSWH - 17) + 9
	
	fld dword yray		; yray
	fistp dword iyr		; iyr = Int(yray)
	mov eax,iyr
	and eax,edx
	add eax,9
	mov iyr,eax			; = (yray AND (FSWH - 17) + 9

	; Get color/height = FloorSurf(ixr,iyr) ->cul
	; and multiply by zhtmult
	mov esi,ptFS		; -> FloorSurf(1,1)
	mov eax,iyr
	dec eax
	mov edx,FSWH
	mul edx		
	add esi,eax 		; ptFS + (iyr-1) * FSWH
	mov eax,ixr
	dec eax
	add esi,eax			; -> FloorSurf(ixr,iyr)
	
	; Get cul
	xor eax,eax
	mov al,[esi]		; cul-ht
	mov cul,eax			; cul in AL
	mov iht,eax			; ht
	
	cmp eax,hlim
	jle NoHtChange
	
	mov eax,hlim
	mov iht,eax			; plateau height
	
NoHtChange:
	; Multiply by zhtmult
	fild dword iht		; iht
	fld dword zhtmult	; zhtmult,iht
	fmulp ST1
	fistp dword iht		; iht = iht * zhtmult
	
	; Compare iht & izray
	mov eax,iht
	cmp eax,izray
	jle near UPDATE_RAY ; iht <= izray

	; zht > zray  -  above drawn area

DRAW_PIX:

	mov edi,ptPS		; ->ProjSurf(1,1)	
	mov eax,row
	dec eax
	mov edx,PSWH
	mul edx
	add edi,eax			; ptPS + (row-1) * PSWH
	mov eax,C
	;inc eax
	add edi,eax			; ->ProjSurf(C+1,row)
	mov eax,cul
	mov [edi],AL		; ProjSurf(C+1,row) = cul
	
	; Update zray
	fld dword zdz
	fld dword zdslope
	faddp st1
	fstp dword zdz		; zdz = zdz + zdslope
	
	fld dword zray
	fld dword zvoxscale
	faddp st1
	fstp dword zray		; zray = zray + zvoxscale
	
	mov eax,row
	inc eax
	mov row,eax			; row = row + 1
	
	cmp eax,PSWH		; row-PSWH
	ja near UPDATE_COLUMN	;row > PSWH
	
	; Test collision
	;---------------
	mov eax,Collide
	cmp eax,1
	je End_Collision    ;Collision already found
	
	mov eax,ex			; eax=ex
	add eax,9			; match +9 for ixr
	mov edx,ixr
	sub edx,2			; edx=ixr-2
	cmp eax,edx			; ex-(ixr-2)
	jl End_Collision	; ex < ixr-2

	add edx,4			; edx=ixr+2
	cmp eax,edx			; ex-(ixr+2)
	jg End_Collision	; ex > ixr+2
	
	mov eax,ey			; eax=ey
	add eax,9			; match +9 for iyr
	mov edx,iyr
	sub edx,2			; edx = iyr-2
	cmp eax,edx			; ey-(iyr-2)
	jl End_Collision	; ey < iyr-2

	add edx,4			; edx=iyr+2
	cmp eax,edx			; ey-(iyr+2)
	jg End_Collision	; ey > iyr+2

	mov eax,1
	mov Collide,eax		; A Collision
End_Collision:
	;---------------
	
	fld dword zray	
	fistp dword izray
	mov eax,izray
	cmp eax,iht			; izray-iht
	jle near DRAW_PIX	; izray <= iht
	
	; izray > iht   column done

UPDATE_RAY:
	fld dword zdx
	fld dword xray
	faddp st1
	fstp dword xray		;xray=xray+zdx
	
	fld dword zdy
	fld dword yray
	faddp st1
	fstp dword yray		;yray=yray+zdy
	
	fld dword zdz
	fld dword zray
	faddp st1
	fstp dword zray		;zray=zray+zdz

	; Store zray as integer izray
	fld dword zray
	fistp dword izray

	fld dword zdslope
	fld dword zvoxscale
	faddp st1
	fstp dword zvoxscale	;zvoxscale=zvoxscale+zdslope
	
	inc ecx
	cmp ecx,LAH			; ecx-LAH
	jle near NEXT_STEP	; ecx <= LAH	Next cstep

UPDATE_COLUMN:
	POP ecx
	inc ecx
	mov eax,PSWH
	dec eax				; PSWH-1
	cmp eax,ecx			; (PSWH-1)-C
	jg near NEXT_COLUMN	; (PSWH-1) > C  ie C < (PSWH-1)  Next C
	
	; C = PSWH-1  all columns done

SaveCollision:
	; Return Collision 0/1 no/yes collision
	mov ebx,[ebp+8]		; ->BITDATA
	mov eax,Collide
	mov [ebx+32],eax	; bmp.Collision = Collide
RET						; END FLOOR_CASTER

;########################################################################
;###### PLASMA ##########################################################
;###### 2048+16 x 2048+16 FloorSurf  Plasma 9->FSWH-8 2048x2048 square ##
;########################################################################

;	TEMPORARY VARIABLES FOR PLASMA

%define FSWHm8		[ebp-56]       ; FSWH-8 For loops
%define FSWHm16		[ebp-60]		; FSWH-16 for Stepsize
%define NoSteps		[ebp-64]       ; Init 2 then x 2
%define StepSize	[ebp-68]       ; StepSize (FSWH-16)/2
%define IY			[ebp-72]
%define IX			[ebp-76]
%define ix1			[ebp-80]
%define iy1			[ebp-84]
%define ix2			[ebp-88]
%define iy2			[ebp-92]
%define ixa 		[ebp-96]
%define iya 		[ebp-100]
%define cul1 		[ebp-104]
%define cul2 		[ebp-108]
%define cul3 		[ebp-112]
%define cul4 		[ebp-116]
%define ixm			[ebp-120]
%define iym			[ebp-124]

%define culr		[ebp-128]    


PLASMA:

	mov eax,[ebp-12]
	mov irand,eax
	
	mov eax,FSWH
	sub eax,8
	mov FSWHm8,eax		; FSWH-8
	
	sub eax,8			; FSWH-16
	mov FSWHm16,eax

	; Seed corners
	
	CALL RandRough		; eax random
	mov ebx,eax			; BL random color
	mov eax,9
	mov ixa,eax
	mov iya,eax
	CALL PutCul			; FloorSurf(9,9) = cul
	
	CALL RandRough		; eax random
	mov ebx,eax			; BL random color
	mov eax,FSWHm8
	mov ixa,eax
	mov eax,9
	mov iya,eax
	CALL PutCul			; FloorSurf(FSWHm8,9) = cul
	
	CALL RandRough		; eax random
	mov ebx,eax			; BL random color
	mov eax,9
	mov ixa,eax
	mov eax,FSWHm8
	mov iya,eax
	CALL PutCul			; FloorSurf(9,FSWHm8) = cul
	
	CALL RandRough		; eax random
	mov ebx,eax			; BL random color
	mov eax,FSWHm8
	mov ixa,eax
	mov iya,eax
	CALL PutCul			; FloorSurf(FSWHm8,FSWHm8) = cul
	
	
	mov eax,2
	mov NoSteps,eax		; Start NoSteps = 2
	
NewStepSize:

	mov eax,FSWHm16		; Calc StepSize
	mov ebx,NoSteps
	cmp ebx,0
	ja ook
	RET
ook:
	xor edx,edx			; CRITICAL because it's  edx:eax/ebx
	div ebx
	mov StepSize,eax	; StepSize = (FSWH-16)/NoSteps

	cmp eax,2			; Check StepSize
	jge Continue
	RET
	
Continue:

	mov ecx,9			; For IY = 9 To FSWH-8
ForIY:
	PUSH ecx
	mov IY,ecx
	
	mov ecx,9			; For IX = 9 To FSWH-8
ForIX:
	PUSH ecx
	mov IX,ecx

	;----------------------------------------
	mov eax,IX			; Set ix1,iy1,ix2,iy2
	cmp eax,FSWHm8
	jle IXok
	mov eax,FSWHm8
IXok:
	mov ix1,eax
	
	mov eax,IY
	cmp eax,FSWHm8
	jle IYok
	mov eax,FSWHm8
IYok:
	mov iy1,eax
	
	mov eax,IX
	add eax,StepSize
	cmp eax,FSWHm8
	jle IXStepok
	mov eax,FSWHm8
IXStepok:
	mov ix2,eax
	
	mov eax,IY
	add eax,StepSize
	cmp eax,FSWHm8
	jle IYStepok
	mov eax,FSWHm8
IYStepok:
	mov iy2,eax
	
	;----------------------------------------
	CALL RandRough		;edx = 16 or 8 * (Rnd - 0.5)
	mov culr,edx
	;----------------------------------------
	
	; Get 4 colors
	mov eax,ix1			; cul1 = FloorSurf(ix1,iy1)
	mov ixa,eax
	mov eax,iy1
	mov iya,eax
	CALL GetCul
	mov cul1,ecx
	
	mov eax,ix1			; cul2 = FloorSurf(ix1,iy2)
	mov ixa,eax
	mov eax,iy2
	mov iya,eax
	CALL GetCul
	mov cul2,ecx
	
	mov eax,ix2			; cul3 = FloorSurf(ix2,iy1)
	mov ixa,eax
	mov eax,iy1
	mov iya,eax
	CALL GetCul
	mov cul3,ecx
	
	mov eax,ix2			; cul4 = FloorSurf(ix2,iy2)
	mov ixa,eax
	mov eax,iy2
	mov iya,eax
	CALL GetCul
	mov cul4,ecx
	;----------------------------------------
	
	;Use  PutCul 5 times Put cul number in BL into FloorSurf(ixa,iya)
	
	mov eax,cul1		; 1 LEFT
	add eax,cul2
	shr eax,1			; (cul1+cul2)/2
	add eax,culr
	and eax,255
	mov ebx,eax			; cul in BL
	
	mov eax,ix1			; FloorSurf(ix1,iym)= cul
	mov ixa,eax
	
	mov eax,iy1
	add eax,iy2
	shr eax,1
	mov iya,eax
	mov iym,eax			; iym=(iy1+iy2)/2
	CALL PutCul
	;-------------------------------------

	mov eax,cul3		; 2 RIGHT
	add eax,cul4
	shr eax,1			; (cul3+cul4)/2
	add eax,culr
	and eax,255
	mov ebx,eax			; cul in BL
	
	mov eax,ix2			; FloorSurf(ix2,iym)= cul
	mov ixa,eax
	
	mov eax,iym			; iym
	mov iya,eax
	CALL PutCul
	;-------------------------------------
	
	mov eax,cul1		; 3 BOTTOM
	add eax,cul3
	shr eax,1			; (cul1+cul3)/2
	add eax,culr
	and eax,255
	mov ebx,eax			; cul in BL
	
	mov eax,ix1
	add eax,ix2
	shr eax,1
	mov ixa,eax
	mov ixm,eax			; ixm=(ix1+ix2)/2
	
	mov eax,iy1			; FloorSurf(ixm,iy1)= cul
	mov iya,eax
	CALL PutCul
	;-------------------------------------

	mov eax,cul2		; 4 TOP
	add eax,cul4
	shr eax,1			; (cul2+cul4)/2
	add eax,culr
	and eax,255
	mov ebx,eax			; cul in BL
	
	mov eax,ixm			; ixm
	mov ixa,eax
	
	mov eax,iy2			; FloorSurf(ixm,iy2)= cul
	mov iya,eax
	CALL PutCul
	;-------------------------------------

	mov eax,cul1		; 5 MIDDLE
	add eax,cul2
	add eax,cul3
	add eax,cul4
	shr eax,2			; (cul1+cul2+cul3+cul4)/4
	add eax,culr
	and eax,255
	mov ebx,eax			; cul in BL
	
	mov eax,ixm			; ixm
	mov ixa,eax
	
	mov eax,iym			; iym
	mov iya,eax
	CALL PutCul			; FloorSurf(ixm,iym)= cul
	;-------------------------------------
	
	POP ecx
	mov eax,StepSize
	add ecx,eax
	mov eax,FSWHm8
	cmp ecx,eax
	jle near ForIX

	POP ecx
	mov eax,StepSize
	add ecx,eax
	mov eax,FSWHm8
	cmp ecx,eax
	jle near ForIY

	mov eax,NoSteps
	shl eax,1	
	mov NoSteps,eax		; NoSteps=NoSteps x 2
	jmp NewStepSize

RET

;########################################################################
;######  RandRough                          #############################
;######  Uses irand, eax, edx               ############################
;########################################################################
RandRough:

;	Produces 	random number in eax &
;				16 or 8 * (Rnd - 0.5) in edx

	mov eax,01180Bh		;71699
	imul dword irand
	add eax,0AB209h		;700937
	; Make odd
	rcr eax,1
	jc ok
	stc
ok:
	rcl eax,1
	mov irand,eax
	
	xor edx,edx
	mov dl,al			;0-255
	
	sub dx,128			;-128 -> +128
	sar dx,3			;3 edx( div 16) = +/-8

RET

;########################################################################
;######  GetCul  Get cul number to ecx                      #############
;######  ecx = [FloorSurf(ixa,iya)                          #############
;######  Uses eax, ebx, ecx, edi  nb edx not used           #############
;########################################################################
GetCul:				; Input ixa,iya  Output: Cul in ecx
	mov eax,FSWH	; Get offset to ixa,iya
	mov ebx,iya
	dec ebx
	mul ebx
	mov ebx,ixa
	dec ebx
	add eax,ebx
	mov edi,ptFS
	add edi,eax
	
	movzx ecx,byte[edi]
	RET
	
;########################################################################
;######  PutCul  Put cul number in BL into FloorSurf(ixa,iya) ###########
;######  Uses eax, edx, edi, ebx input  nb ecx not used     #############
;########################################################################
PutCul:				; Input ixa,iya, cul in BL
	mov eax,FSWH	; Get offset to ixa,iya
	mov edx,iya
	dec edx
	mul edx
	mov edx,ixa
	dec edx
	add eax,edx
	mov edi,ptFS
	add edi,eax
	
	mov [edi],BL
	RET

;########################################################################
;######  BYTESMOOTHING  #################################################
;########################################################################

BYTESMOOTHING:

MOV ecx,4

NSmoothers:

PUSH ecx

	
	mov edi,ptFS
	mov eax,9
	dec eax
	mov ebx,FSWH
	mul ebx
	add edi,eax
	mov eax,9
	dec eax
	add edi,eax
	
	mov ecx,4227072		; 2048 x 2064
	xor eax,eax

ByteSmooth:
	xor edx,edx
	mov AL,[edi-1]
	add edx,eax
	mov AL,[edi+1]
	add edx,eax
	mov AL,[edi-2048]
	add edx,eax
	mov AL,[edi+2048]
	add edx,eax
	shr edx,2
	cmp edx,3
	ja culok
	mov DL,128
culok:
	mov [edi],DL
	inc edi
	dec ecx
	jnz ByteSmooth

POP ecx
dec ecx
jnz near NSmoothers


RET
;###########################################
;;;;;;;;;;;;;;;;;;;;
;pop ecx
;mov eax,ebx 
;RET
;;;;;;;;;;;;;;;;;;;;

	