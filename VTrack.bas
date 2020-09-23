Attribute VB_Name = "Module1"
' VTrack2.bas by Robert Rayment  1/5/01

Option Base 1  ' Arrays' base 1

DefLng A-W     ' Long integers
DefSng X-Z     ' Singles

' To shift cursor out of the way
Public Declare Sub SetCursorPos Lib "user32" (ByVal IX As Long, ByVal IY As Long)

' To get user keying
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function GetTickCount Lib "kernel32" () As Long
Global TickDifference As Long 'timing
Global LastTick As Long
Global CurrentTick As Long


' Copy one array to another of same number of bytes
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)


' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   Colors(0 To 255) As RGBQUAD
End Type
Public bm As BITMAPINFO

' For transferring drawing in byte array to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcX As Long, ByVal SrcY As Long, _
ByVal SrcW As Long, ByVal SrcH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

' Constants for StretchDIBits
Public Const DIB_PAL_COLORS = 1 '  color table in palette indices
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046

' Set alternative actions for StretchDIBits (not used here)
Public Declare Function SetStretchBltMode Lib "gdi32" _
(ByVal hdc As Long, ByVal nStretchMode As Long) As Long
' nStretchMode
Public Const STRETCH_ANDSCANS = 1    'default
Public Const STRETCH_ORSCANS = 2
Public Const STRETCH_DELETESCANS = 3
Public Const STRETCH_HALFTONE = 4

' For calling machine code
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpMCode As Long, _
ByVal Long1 As Long, ByVal Long2 As Long, _
ByVal Long3 As Long, ByVal OpCode As Long) As Long

'=========================================================
' Structure for input to mcode for RayCASTER & Plasma
' PS = ProjSurf(), FS = FloorSurf()
Public Type BITDATA
   
   PSWH As Long            ' Width=Height of PS         0 PSWH
   ptrProjSurf As Long     ' Pointer to PS              4 ptPS
   FSWH As Long            ' Width=Height of FS         8 FSWH
   ptrFloorSurf As Long    ' Pointer to FS             12 ptFS
   
   LAH As Long             ' Look ahead length L       16 LAH        'Input
   
   ex As Long              ' eye x-coord               20 ex         'Input
   ey As Long              ' eye y-coord               24 ey         'Input
   TANG As Long            ' Track angle 0-359         28 TANG       'Input
   Collision As Long       ' 0/1 no/yes collision      32 Collision  'Return
   
   zvp As Single           ' eye z-coord               36 zvp        'Input
   zdslope As Single       ' Ray slope                 40 zdslope    'Input
   zslopemult As Single    ' Ray slope multiplier      44 zslopemult 'Input
   zhtmult As Single       ' Objects height multiplier 48 zhtmult    'Input
   
   hlim As Long            ' Height limit for plateaux 52 hlim       'Input Default(255)

End Type
Public bmp As BITDATA

Global ptMCode&   ' = VarPtr(InCode(1))
Global ptBITDATA& ' = VarPtr(bmp.PSWH)

'=========================================================

Global ControlAutoManual As Boolean    ' Automatic or Manual
Global CodeType As Boolean             ' VB or Machine code
Global SurfaceType                     ' 0,1,2  Pyramids, Plasma & BigPlasma


Global PSWH, FSWH   ' ProjSurf & FloorSurf sizes
Global LAH          ' Look ahead distance from front to back clipping plane
Global ex, ey       ' Viewer coords - nb. integers

' User key actions
Global TANG    ' Angle picked up from Track or USER
Global keyLeft As Boolean, keyRight As Boolean
Global keyGoUP As Boolean, keyGoDN As Boolean
Global zEL           ' +/- movement length from keying
Global Collision     ' 0/1 no/yes collision
Global Rand&

' Track factors
Global MAXTrackPts, Marg, Rad   ' Track #points, margin & corner radius
Global TrackXC() As Long, TrackYC() As Long, TrackANG() As Long

' Surfaces
Global ProjSurf() As Byte
Global FloorSurf() As Byte
Global ZeroSurf() As Byte

Global PathSpec$        ' App path
Global PalSpec$         ' PAL file held in bm.Colors(0 To 255) As RGBQUAD

'----------------------------------------------------------------------
' Global FloorCASTER parameters
Global zdslope    ' = 0.05    ' Starting slope of ray for each column scanned.  Larger
                  ' values flatten & smaller values exaggerate heights.

Global zhtmult    ' = 2       ' Voxel ht multiplier.

Global zslopemult ' = -100    ' Slope multiplier.  Larger -ve values look
                  ' more vertically down.

Global zvp        ' = 256     ' Starting height of viewer.  Higher values lift
                  ' the viewer & give a more pointed perspective.

Global hlim       ' Plateaux Surface type 7 height limiter
'----------------------------------------------------------------------

Global zHt, ixr, iyr
Global Smooth

Global InCode() As Byte ' To hold mcode

Global Const pi# = 3.1415927
Global Const dtr# = pi# / 180

Public Sub Pyramids()

' Draw pyramids to FloorSurf()

Dim maxcul As Byte
Dim cul As Byte

Randomize

' Floor surf size = FSWH

SEPSTEP = 96   ' Colored pyramids SEPSTEP apart
HTSTEP = 48    ' Size of pyramids
For J = 1 To FSWH - SEPSTEP Step SEPSTEP
For I = 1 To FSWH - SEPSTEP Step SEPSTEP
   maxcul = 125 + Rnd * 126
   culstep = 2 * maxcul / HTSTEP
   cul = 0
   For VSTEP = 0 To HTSTEP \ 2
      ' Mov start point diagonally & shorten sides
      jlo = J + VSTEP
      jup = J + HTSTEP - 1 - VSTEP
      ilo = I + VSTEP
      iup = I + HTSTEP - 1 - VSTEP
      ' 2 horizontal lines
      For jj = jlo To jup Step (jup - jlo)
      For ii = I + VSTEP To iup
         FloorSurf(ii, jj) = cul
      Next ii
      Next jj
      ' 2 vertical lines
      For ii = ilo To iup Step (iup - ilo)
      For jj = J + VSTEP To jup
         FloorSurf(ii, jj) = cul
      Next jj
      Next ii
      If cul < (255 - culstep) Then
         cul = cul + culstep
      Else
         cul = 50
      End If
   Next VSTEP
Next I
Next J

' OUTSIDE WALLS

InDis = 400
cul = 246
cul = 220
For I = 1 To FSWH
   J = InDis
   FloorSurf(I, J) = cul
   
   J = FSWH - InDis
   FloorSurf(I, J) = cul
Next I

For J = 1 To FSWH
   I = InDis
   FloorSurf(I, J) = cul
   
   
   I = FSWH - InDis
   FloorSurf(I, J) = cul
Next J

End Sub

Public Sub Plasma()

' Non-recursive plasm generator for a FSWxFSWH (ie 2^11 x 2^11)
' byte array

' NB PlasmaTile used instead which copies a smaller plasma tile
' over some of the surface

Randomize

' Set 4 outer corners to random color numbers
FloorSurf(1, 1) = 255 * Rnd
FloorSurf(FSWH, 1) = 255 * Rnd
FloorSurf(1, FSWH) = 255 * Rnd
FloorSurf(FSWH, FSWH) = 255 * Rnd

Roughness = 20   ' Larger values increase randomness
                 ' If too large just get a speckly random pattern
FSWH1 = FSWH - 1 ' For ANDing

For Lev = 2 To 11  ' 11 levels needed to cover the 2048x2048 surface
   NoSteps = 2 ^ Lev
   StepSize = FSWH / NoSteps
   For IY = 1 To FSWH + 1 Step StepSize
   For IX = 1 To FSWH + 1 Step StepSize
      ix1 = (IX And FSWH1) + 1
      iy1 = (IY And FSWH1) + 1
      ix2 = ((IX + StepSize) And FSWH1) + 1
      iy2 = ((IY + StepSize) And FSWH1) + 1
      cul1 = FloorSurf(ix1, iy1)
      cul2 = FloorSurf(ix1, iy2)
      cul3 = FloorSurf(ix2, iy1)
      cul4 = FloorSurf(ix2, iy2)
      'Left ix1
      iym = (iy1 + iy2) \ 2
      cula = (cul1 + cul2) \ 2 + Roughness * (Rnd - 0.5)
      FloorSurf(ix1, iym) = cula And &HFF
      'Right ix2
      cula = (cul3 + cul4) \ 2 + Roughness * (Rnd - 0.5)
      FloorSurf(ix2, iym) = cula And &HFF
      'Bottom  iy1
      ixm = (ix1 + ix2) \ 2
      cula = (cul1 + cul3) \ 2 + Roughness * (Rnd - 0.5)
      FloorSurf(ixm, iy1) = cula And &HFF
      'Top iy2
      cula = (cul2 + cul4) \ 2 + Roughness * (Rnd - 0.5)
      FloorSurf(ixm, iy2) = cula And &HFF
      'Center ixm,iym
      cula = (cul1 + cul2 + cul3 + cul4) \ 4 + Roughness * (Rnd - 0.5)
      FloorSurf(ixm, iym) = cula And &HFF
   Next IX
   Next IY
Next Lev

End Sub

Public Sub PlasmaTile()

' Non-recursive plasma generator for a FSWH x FSWH (ie 2^11 x 2^11)
' byte array (FloorSurf())

' USING TILING of first 256x256 part of FloorSurf
' NB this assumes FSWH is 2048 x 2048

Randomize

' Set 4 outer corners to random color numbers
FloorSurf(1, 1) = 255 * Rnd
FloorSurf(256, 1) = 255 * Rnd
FloorSurf(256, 256) = 255 * Rnd
FloorSurf(256, 256) = 255 * Rnd

Roughness = 20    ' Larger values increase randomness
                  ' If too large just get a speckly random pattern
For Lev = 2 To 8  ' 8 levels needed to cover the 256x256 surface
   NoSteps = 2 ^ Lev
   StepSize = 256 / NoSteps
   For IY = 1 To 256 + 1 Step StepSize
   For IX = 1 To 256 + 1 Step StepSize
      ix1 = (IX And 255) + 1
      iy1 = (IY And 255) + 1
      ix2 = ((IX + StepSize) And 255) + 1
      iy2 = ((IY + StepSize) And 255) + 1
      cul1 = FloorSurf(ix1, iy1)
      cul2 = FloorSurf(ix1, iy2)
      cul3 = FloorSurf(ix2, iy1)
      cul4 = FloorSurf(ix2, iy2)
      ' Left ix1
      iym = (iy1 + iy2) \ 2
      cula = (cul1 + cul2) \ 2 + Roughness * (Rnd - 0.5)
      FloorSurf(ix1, iym) = cula And &HFF
      ' Right ix2
      cula = (cul3 + cul4) \ 2 + Roughness * (Rnd - 0.5)
      FloorSurf(ix2, iym) = cula And &HFF
      ' Bottom  iy1
      ixm = (ix1 + ix2) \ 2
      cula = (cul1 + cul3) \ 2 + Roughness * (Rnd - 0.5)
      FloorSurf(ixm, iy1) = cula And &HFF
      ' Top iy2
      cula = (cul2 + cul4) \ 2 + Roughness * (Rnd - 0.5)
      FloorSurf(ixm, iy2) = cula And &HFF
      ' Center ixm,iym
      cula = (cul1 + cul2 + cul3 + cul4) \ 4 + Roughness * (Rnd - 0.5)
      FloorSurf(ixm, iym) = cula And &HFF
   Next IX
   Next IY
Next Lev

' TILE PATTERN over 2048 x 2048 surface
' leaving some gaps (sea?)

For J = 256 To 1792 Step 256
For I = 256 To 1792 Step 256

   For jj = 1 To 256
   For ii = 1 To 256
      FloorSurf(I + ii - 1, J + jj - 1) = FloorSurf(ii, jj)
   Next ii
   Next jj
   
Next I
Next J


End Sub

Public Function OutSideRect(ex, ey, ByVal xTL, ByVal yTL, ByVal xBR, ByVal yBR)
' Rectangle is 256,256 -> 2048-256,2048-256 1792,1792
If ex < xTL Then OutSideRect = True: Exit Function
If ex > xBR Then OutSideRect = True: Exit Function
If ey > yTL Then OutSideRect = True: Exit Function
If ey < yBR Then OutSideRect = True: Exit Function
OutSideRect = False
End Function

Public Sub Loadmcode(InFile$)
' Load machine code into InCode() byte array
On Error GoTo InFileErr
If Dir$(InFile$) = "" Then
   MsgBox (InFile$ & " missing")
   Erase ProjSurf, FloorSurf
   DoEvents
   Unload Form1
   End
End If
Open InFile$ For Binary As #1
MCSize& = LOF(1)
If MCSize& = 0 Then
InFileErr:
   MsgBox (InFile$ & " missing")
   Erase ProjSurf, FloorSurf
   DoEvents
   Unload Form1
   End
End If
ReDim InCode(MCSize&)
Get #1, , InCode
Close #1
On Error GoTo 0
End Sub

Public Sub ReadPAL(PalSpec$)
' Read JASC-PAL palette file
' Any error shown by PalSpec$ = ""
' Else RGB into Colors(i) Long

Dim RED As Byte, Green As Byte, BLUE As Byte
On Error GoTo palerror
Open PalSpec$ For Input As #1
Line Input #1, a$
p = InStr(1, a$, "JASC")
If p = 0 Then PalSpec$ = "": Exit Sub
   
   'JASC-PAL
   '0100
   '256
   Line Input #1, a$
   Line Input #1, a$

   For N = 0 To 255
      If EOF(1) Then Exit For
      Line Input #1, a$
      ParsePAL a$, RED, Green, BLUE
      bm.Colors(N).rgbBlue = BLUE
      bm.Colors(N).rgbGreen = Green
      bm.Colors(N).rgbRed = RED
      bm.Colors(N).rgbReserved = 0
   Next N
   Close #1

Exit Sub
'===========
palerror:
PalSpec$ = ""
Exit Sub
End Sub

Public Sub ParsePAL(ain$, RED As Byte, Green As Byte, BLUE As Byte)
' Input string ain$, with 3 numbers(R G B) with
' space separators and then any text
ain$ = LTrim(ain$)
lena = Len(ain$)
R$ = ""
g$ = ""
b$ = ""
num = 0 'R
nt = 0
For I = 1 To lena
   C$ = Mid$(ain$, I, 1)
   
   If C$ <> " " Then
      If nt = 0 Then num = num + 1
      nt = 1
      If num = 4 Then Exit For
      If Asc(C$) < 48 Or Asc(C$) > 57 Then Exit For
      If num = 1 Then R$ = R$ + C$
      If num = 2 Then g$ = g$ + C$
      If num = 3 Then b$ = b$ + C$
   Else
      nt = 0
   End If
Next I
RED = Val(R$): Green = Val(g$): BLUE = Val(b$)
End Sub
