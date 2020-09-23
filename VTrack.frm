VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   Caption         =   "VTrack2"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8430
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   Icon            =   "VTrack.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   562
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSteer 
      Enabled         =   0   'False
      Height          =   780
      Left            =   30
      Picture         =   "VTrack.frx":0442
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   136
      TabIndex        =   0
      Top             =   0
      Width           =   2100
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   120
         Left            =   1650
         Picture         =   "VTrack.frx":54FC
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   2
         Top             =   15
         Width           =   300
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VTrack by Robert Rayment"
      BeginProperty Font 
         Name            =   "ScriptC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   2865
      TabIndex        =   1
      Top             =   1275
      Width           =   3825
   End
   Begin VB.Menu mnuStart 
      Caption         =   "&START"
   End
   Begin VB.Menu mnuStop 
      Caption         =   "S&TOP"
   End
   Begin VB.Menu zdum1 
      Caption         =   "Control"
      Begin VB.Menu mnuAutomatic 
         Caption         =   "Automatic"
      End
      Begin VB.Menu mnuManual 
         Caption         =   "Manual"
      End
   End
   Begin VB.Menu zdum3 
      Caption         =   "Surface type"
      Begin VB.Menu mnuPyramids 
         Caption         =   "Pyramids"
      End
      Begin VB.Menu mnuPlasma 
         Caption         =   "Plasma"
      End
      Begin VB.Menu brk1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGreyClouds 
         Caption         =   "Grey clouds (MC)"
      End
      Begin VB.Menu mnuMoonlight 
         Caption         =   "Moonlight (MC)"
      End
      Begin VB.Menu mnuDayClouds 
         Caption         =   "Day Clouds (MC)"
      End
      Begin VB.Menu mnuHills 
         Caption         =   "Hills (MC)"
      End
      Begin VB.Menu mnuDunes 
         Caption         =   "Dunes (MC)"
      End
      Begin VB.Menu mnuPlateaux 
         Caption         =   "Plateaux (MC)"
      End
   End
   Begin VB.Menu p2 
      Caption         =   "Smoother"
      Begin VB.Menu DoSmoothing 
         Caption         =   "Smooth"
      End
      Begin VB.Menu NoSmoothing 
         Caption         =   "No smoothing"
      End
   End
   Begin VB.Menu zdum5 
      Caption         =   "Palettes"
      Begin VB.Menu HalfSine 
         Caption         =   "Half sine"
      End
      Begin VB.Menu Grey 
         Caption         =   "Grey"
      End
      Begin VB.Menu Sine 
         Caption         =   "Sine"
      End
      Begin VB.Menu Blues 
         Caption         =   "Blues"
      End
      Begin VB.Menu Green 
         Caption         =   "Green"
      End
      Begin VB.Menu Dunes 
         Caption         =   "Dunes"
      End
      Begin VB.Menu Earth 
         Caption         =   "Earth"
      End
   End
   Begin VB.Menu zdum7 
      Caption         =   "Code"
      Begin VB.Menu mnuVB6 
         Caption         =   "VB6"
      End
      Begin VB.Menu mnuMachineCode 
         Caption         =   "Machine code"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  VTrack2  by Robert Rayment 1/5/01

' NB PSWH-1 caused inteference problems??? for
' StretchDIBIts.  PSWH  corrects!!

Option Base 1  ' Arrays' base 1

DefLng A-W     ' Long integers
DefSng X-Z     ' Singles


' This program uses floor ray-casting from the eye onto a color-height
' byte array (ie for simplicity avoiding separate height & color maps).
' The color-heights are then projected back to a front surface byte
' array which blitted to the screen using the API StretchDIBits.  Using
' byte arrays for color is fast but blocky.  The extension to 16-bit
' color is straightforward in theory, the byte arrays being replaced
' by integer arrays, StretchDIBits would then pick up the color directly
' from the array, but 16-bit color squeezes the RGB into 15 bits with
' 1 alpha bit and is more complicated to work with.  NLS, on PSC, shows
' an alternative method for Voxels with 24-bit color.

' Controls are the cursor keys to move forward, backward, turn to the left
' and turn to the right.  With Shift the up & down arrows change the height
' of objects and Ctrl with the up & down arrows moves the viewer up & down.
' The SPACE bar brings the motion to an immediate halt.

' Motion control can be Automatic or Manual.  With Automatic the motion
' proceeds along one circuit of a track by itself and gives the timimg
' in seconds. Whereas Manual allows the User to control the motion.  With
' the Pyramids a wall forms a boundary around the surface but with Plasmas
' the surface rolls over giving continuous motion.  Different palettes can
' be selected and they take effect immediately.  Changing the motion control
' surface type or smoothing requires the START button to be pressed.  Can
' change the surface appearance by changing the palette but could have
' Tracks in the sky - which doesn't really matter!


' The casting routine is modified from that of D Brebner's
' of Unlimited Realities.  In particular a spherical correction
' is applied and the scanning is from -ve to +ve angles placing the
' viewer in the middle of the screen.  This FloorCASTER routine is
' commented in detail - the effects are very subtle.

' The main jump-off routine is STARTACTION

'----------------------------------------------------------------------
' MACHINE CODE
' Using NASM. Netwide Assembler freeware from
' www.phoenix.gb/net
' www.cryogen.com/Nasm
' & other sites
' see also
' www.geocities.com/SunsetStrip/Stage/8513/assembly.html
'----------------------------------------------------------------------

Dim Done As Boolean  'To exit loop





Private Sub Form_Load()

Picture1.Visible = False
DoSmoothing_Click

' Default global parameters for FloorCASTER
' - all these markedly affect the resulting display.

zdslope = 0.05    ' Starting slope of ray for each column scanned.  Larger
                  ' values flatten & smaller values exaggerate heights.

zhtmult = 2       ' Voxel ht multiplier.

zslopemult = -100 ' Slope multiplier.  Larger -ve values look
                  ' more vertically down.

zvp = 256         ' Starting height of viewer.  Higher values lift
                  ' the viewer & give a more pointed perspective.

'LAH = 512        ' Look ahead distance - the larger it is the more
                  ' of the FloorSurf is scanned and the slower it will
                  ' be.  Set in STARTACTION

hlim = 255        ' Default height limiter
'-----------------

' Starting Track factors
Marg = 512 + 128
Rad = 45
'-----------------

' Align controls PictureBox

With picSteer
   .Width = 80
   .Top = 0
   .Left = 0
   .Height = 52
   .Width = 140
End With

'-----------------------------
' Set Starting Menu Options

  SetOptions

'-----------------------------


KeyPreview = True    ' Allows Form to get keying first

' Default - no keys pressed
keyRight = False
keyLeft = False
keyGoUP = False
keyGoDN = False

' Start movement increment keying
zEL = 0
Collision = 0

' Set form up
ScaleMode = vbPixels
WindowState = vbNormal
AutoRedraw = False
Top = 1000
Left = 1000

' Set form size

Width = 640 * Screen.TwipsPerPixelX    'ie * 15
Height = 480 * Screen.TwipsPerPixelY   'ie * 15

'Width = 800 * Screen.TwipsPerPixelX    'ie * 15
'Height = 600 * Screen.TwipsPerPixelY   'ie * 15

'Width = 256 * Screen.TwipsPerPixelX    'ie * 15
'Height = 256 * Screen.TwipsPerPixelY   'ie * 15


Show
DoEvents

' Get app path
PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

' Starting palette
PalSpec$ = PathSpec$ & "HalfSine.PAL"
ReadPAL PalSpec$

' Selected dimensions

' PSWH & FSWH are the ProjSurf & FloorSurf sizes and they
' must be squares.

' NOTE PSHW can be changed but NOT FSWH
' this would be OK for floor-casting but not
' for m/code smoothing which assumes a 2048 x 2048
' floor surface plus an 8 byte border.

' If its all too fast PSWH can be increased to
' say 384 or 512.

PSWH = 256 '384 '512          ' ProjSurf Width & Height
FSWH = 2048 + 16              ' FloorSurf Width & Height + 8 pixel border
ReDim ProjSurf(PSWH, PSWH)
ReDim ZeroSurf(PSWH, PSWH)   ' ProjSurf background
ReDim FloorSurf(FSWH, FSWH)
   
' FIXED VALUES IN BITDATA for machine code
'   PSWH As Long            'Width=Height of PS         0 PSWH
'   ptrProjSurf As Long     'Pointer to PS              4 ptPS
'   FSWH As Long            'Width=Height of FS         8 FSWH
'   ptrFloorSurf As Long    'Pointer to FS             12 ptFS
bmp.PSWH = PSWH
bmp.ptrProjSurf = VarPtr(ProjSurf(1, 1))
bmp.FSWH = FSWH
bmp.ptrFloorSurf = VarPtr(FloorSurf(1, 1))

' Store pointer to BITDATA
ptBITDATA& = VarPtr(bmp.PSWH)

' Fill BITMAPINFO.BITMAPINFOHEADER FOR StretchDIBits
bm.bmiH.biSize = 40
bm.bmiH.biwidth = PSWH
bm.bmiH.biheight = PSWH
bm.bmiH.biPlanes = 1
bm.bmiH.biBitCount = 8
bm.bmiH.biCompression = 0
bm.bmiH.biSizeImage = PSWH * PSWH '0 '0 not needed here
bm.bmiH.biXPelsPerMeter = 0
bm.bmiH.biYPelsPerMeter = 0
bm.bmiH.biClrUsed = 0
bm.bmiH.biClrImportant = 0

'-------------------------
' Set Track arrays
ReDim TrackXC(6000), TrackYC(6000), TrackANG(6000)

'=========================================================

' Load mcode
InFile$ = PathSpec$ & "VTrack.bin"
Loadmcode (InFile$)

' Store pointer to machine code array
ptMCode& = VarPtr(InCode(1))

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

Case vbKeyLeft
   
   keyLeft = True
   keyRight = False
   keyGoUP = False
   keyGoDN = False
   
Case vbKeyRight
   
   keyRight = True
   keyLeft = False
   keyGoUP = False
   keyGoDN = False
   
Case vbKeyUp
   
   If Shift = 1 Then 'SHIFT Increase object heights
      
      ' Default zhtmult = 2
      zhtmult = zhtmult + 0.2
      If zhtmult > 4 Then zhtmult = 4
   
   ElseIf Shift = 2 Then  'CTRL Move up
      
      ' Default zdslope = 0.05
      ' Decreasing zdslope point down further
      ' rem  zdz =2 * zdslope * -ve value
      zdslope = zdslope - 0.008
      If zdslope < 0.01 Then zdslope = 0.01
   
      ' Default zvp=256
      ' Increase zvp raises viewer
      zvp = zvp + 32
      If zvp > 416 Then zvp = 416
      
      ' Default 2 * PSWH = 512
      ' Increase LAH to look further ahead (also slows)
      LAH = LAH + 48
      If LAH > 752 Then LAH = 752
      
      
   Else
      If keyGoUP = True Then
         zEL = zEL + 1
         If zEL > 256 Then zEL = 256
      End If
      keyGoUP = True
      keyGoDN = False
      keyLeft = False
      keyRight = False
   End If
   
Case vbKeyDown
   
   If Shift = 1 Then 'SHIFT Decrease object heights
      
      ' Default zhtmult = 2
      zhtmult = zhtmult - 0.2
      If zhtmult < 0# Then zhtmult = 0#
   
   ElseIf Shift = 2 Then 'CTRL Move down
      
      ' Default value zdslope = 0.05
      ' Increasing zdslope point up further
      ' rem  zdz = 2 * zdslope * -ve value
      zdslope = zdslope + 0.008
      If zdslope > 0.09 Then zdslope = 0.09
      
      ' Default zvp=256
      ' Decrease zvp lowers viewer
      zvp = zvp - 32
      If zvp < 96 Then zvp = 96
      
      ' Default 2 * PSWH = 512
      ' Decrease LAH back towards default
      LAH = LAH - 48
      If LAH < 272 Then LAH = 272
      
   Else
      If keyGoDN = True Then
         zEL = zEL - 1
         If zEL < -256 Then zEL = -256
      End If
      keyGoUP = False
      keyGoDN = True
      keyLeft = False
      keyRight = False
   End If

Case vbKeySpace    ' HALT
   zEL = 0


End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

   keyLeft = False
   keyRight = False
   keyGoUP = False
   keyGoDN = False

End Sub

Private Sub Form_Resize()

'ShowFloorSurf
'DoEvents
'ShowFloorSurf

End Sub


Private Sub FloorCASTER()

' Cast onto the FloorSurf and draw the perspective
' result onto the display array (ProjSurf)
' ex,ey viewer coords on FloorSurf

'     ^ Y   increasing row numbers in FloorSurf (1->FSWH)
'     |
'     |
'     |
'     |
'     |
'      -----------> X  increasing column numbers (1->FSWH)


' PSWH is ProjSurf width & height
' FSWH is FloorSurf width & height
' LAH Look ahead - distance to back clipping plane  [[Changed by CTRL + Up & down cursors]]


' Default parameters - all these markedly affect the resulting display.

' zdslope = 0.05    ' Starting slope of ray for each column scanned.  Larger
                    ' values flatten & smaller values exaggerate heights.
                    ' [Changed by CTRL + Up & down cursors]
                    
' zhtmult = 2       ' Voxel ht multiplier.  [[Changed by SHIFT + Up & down cursors]]

' zslopemult = -100 ' Slope multiplier.  Larger -ve values look
                    ' more vertically down.

' zvp = 256         ' Starting height of viewer.  Higher values lift
                    ' the viewer & give a more pointed perspective.
                    ' [[Changed by CTRL + Up & down cursors]]

' hlim              ' height limiter
'-----------------

' Values picked up from Auto Track or User
zrayang = (TANG * 2 * pi#) - PSWH / 2  ' TANG = angle in degrees
xvp = ex
yvp = ey
'-----------------

For C = 0 To PSWH - 1
   xray = xvp:  yray = yvp:  zray = zvp
   ' ie start at same viewpoint for each column
   
   ' x & y components to move a unit distance along ray
   zdy = Cos((zrayang + C) / 360)
   zdx = Sin((zrayang + C) / 360)
   
   ' Starting ray height below floor
   zdz = 2 * zdslope * zslopemult ' / (Cos((C - PSWH / 2) / 360) ^ 3)  'dish
   
   ' Factor 2 * removes spherical distortion
   ' provided zvoxscale = zvoxscale + zdslope
   ' & * 0.5 omitted before Next cstep in Brebner's code
   
   ' zslopemult - larger -ve values look more vertically down
   
   ' / (Cos((C - PSWH / 2) / 360) ^ 3) produces dished track
   
   zvoxscale = 0  ' Moves zray up as each pixel is drawn so that
                  ' only higher pixels are drawn thereafter.
                  ' zray is also modified according to the distance
                  ' away from the viewer by the line zdz = zdz + zdslope
   
   row = 1        ' First screen row
   
   For cstep = 0 To LAH
      
      'Ensure rows and columns stay on FloorSurf
      ixr = (xray And (2047)) + 9      ' +9 to allow for Border
      iyr = (yray And (2047)) + 9
      
      FSHt = FloorSurf(ixr, iyr)
      
      If SurfaceType = 7 And FSHt > hlim Then
         FSHt = hlim 'Plateax
      End If
      
      
      zHt = FSHt * zhtmult
      ' * zhtmult simply scales heights without distortion
      ' this parameter is changed by vbKeyUp + Shift or vbKeyDown + Shift
      ' when in manual mode.
      
      
      If zHt > zray Then
         
         Do
            ' Copy color from floor to screen
            ProjSurf(C + 1, row) = FloorSurf(ixr, iyr)
            zdz = zdz + zdslope
            zray = zray + zvoxscale
            row = row + 1
            If row > PSWH Then
               cstep = LAH
               Exit Do
            End If
        
            ' Collision detection when eye point is roughly
            ' vertically above floor point being drawn to screen
            ' ex + 9, ey + 9 to match ixr,iyr.
            
            If Not OutSideRect(ex + 9, ey + 9, ixr - 2, iyr + 2, ixr + 2, iyr - 2) Then
               Collision = 1
            Else
               Collision = 0
            End If
         
        Loop Until zray > zHt
         
         ' Fill remainder of column
         'If zray < PSWH Then
         '   For zr = zrow To PSWH
         '    ProjSurf(C + 1, zr) = 200 'cul And 255
         '   Next zr
         'End If
         ' NB CopyMemory is quicker
         
      End If
      xray = xray + zdx
      yray = yray + zdy
      zray = zray + zdz
      zvoxscale = zvoxscale + zdslope
   
   Next cstep
Next C

End Sub


Private Sub MakeFloorSurf()

Caption = "Track   LOADING FLOOR"

Done = True    'end RunCode looping

MousePointer = vbHourglass


Select Case SurfaceType
Case 0
      Pyramids
Case 1
   If CodeType = True Then ' VB
      PlasmaTile  ' VB
   Else           ' MC
      PlasmaTile  ' but VB PlasmaTile
   End If
Case Is >= 2      ' Big Plasma surfaces  (too slow in VB)
   If CodeType = True Then    ' VB
      PlasmaTile     'IF NOT MC will do PlasmaTile for MC Surfaces
   Else              ' MC
      RunMCode 1     'Machine code Big Plasmas
   
      If Smooth = True Then
         RunMCode 2  ' Machine code smoother
      End If
   
   End If
End Select

MousePointer = vbDefault

End Sub

Private Sub MakeTrack()

' Overlay FloorSurf with a rounded rectangular track &
' store track x,y centers and angle of track

Caption = "Track   LOADING TRACK"


Randomize

'LAY TRACK

iww = 15                   ' Half width of Track
Marg = 256 + 256 + 128 + 8 'Track margin
'----------------------
'Left 0 deg
I = Marg
For J = Marg To FSWH - Marg + iww
   iw = iww + 2 * Sin(CDbl(J) / 32)    ' Make wavy edges
   For k = I - iw To I + iw
      FloorSurf(k, J) = 4 + 3 * (0.5 - Rnd)
   Next k
Next J
'Center line
For J = Marg To FSWH - Marg + iww Step 6
   For jj = J To J + 4
      FloorSurf(I, jj) = 1
   Next jj
Next J
'----------------------

'Top 90 deg
J = FSWH - Marg
For I = Marg To FSWH - Marg + iww
   iw = iww + 2 * Sin(CDbl(I) / 32)
   For k = J - iw To J + iw
      FloorSurf(I, k) = 6 + 3 * (0.5 - Rnd)
   Next k
Next I
'Center line
For I = Marg To FSWH - Marg + iww Step 6
   For ii = I To I + 4
      FloorSurf(ii, J) = 1
   Next ii
Next I
'----------------------

'Right 180 deg
I = FSWH - Marg
For J = FSWH - Marg To Marg - iww Step -1
   iw = iww + 2 * Sin(CDbl(J) / 64)
   For k = I - iw To I + iw
      FloorSurf(k, J) = 8 + 3 * (0.5 - Rnd)
   Next k
Next J
'Center line
For J = FSWH - Marg To Marg - iww Step -6
   For jj = J To J - 4 Step -1
      FloorSurf(I, jj) = 1
   Next jj
Next J
'----------------------

'Bottom 270 deg
J = Marg
For I = FSWH - Marg To Marg - iww Step -1
   iw = iww + 2 * Sin(CDbl(I) / 64)
   For k = J - iw To J + iw
      FloorSurf(I, k) = 10 + 3 * (0.5 - Rnd)
   Next k
Next I
'Center line
For I = FSWH - Marg To Marg - iww Step -6
   For ii = I To I - 4 Step -1
      FloorSurf(ii, J) = 1
   Next ii
Next I

'============================================

'SET TRACK COORDS & TRACK ANGLE FOR AUTOMATIC CIRCUIT WITH TIMING

NP = 1 'Number of track points

Rad = 45   'Radius of corners

'----------------------
'Left 0 deg
I = Marg
For J = Marg + Rad To FSWH - Marg - Rad
   TrackXC(NP) = I
   TrackYC(NP) = J
   TrackANG(NP) = 0
   NP = NP + 1
Next J
'TopLeft   0 - 90 deg
ex1 = TrackXC(NP - 1)
ey1 = TrackYC(NP - 1)
For zalpha = 0 To 90
   TrackYC(NP) = ey1 + Rad * Sin(zalpha * dtr#)
   TrackXC(NP) = ex1 + Rad * (1 - Cos(zalpha * dtr#))
   TrackANG(NP) = zalpha
   NP = NP + 1
Next zalpha

'----------------------
'Top 90 deg
J = FSWH - Marg
For I = Marg + Rad To FSWH - Marg - Rad  '''''''
   TrackXC(NP) = I
   TrackYC(NP) = J
   TrackANG(NP) = 90
   NP = NP + 1
Next I
'Top right 90 - 180 deg
ex1 = TrackXC(NP - 1)
ey1 = TrackYC(NP - 1)
For zalpha = 90 To 180
   TrackYC(NP) = ey1 - Rad * (1 - Cos((zalpha - 90) * dtr#))
   TrackXC(NP) = ex1 + Rad * Sin((zalpha - 90) * dtr#)
   TrackANG(NP) = zalpha
   NP = NP + 1
Next zalpha

'----------------------
'Right 180 deg
I = FSWH - Marg
For J = FSWH - Marg - Rad To Marg + Rad Step -1
   TrackXC(NP) = I
   TrackYC(NP) = J
   TrackANG(NP) = 180
   NP = NP + 1
Next J
'Bottom right   180 - 270 deg
ex1 = TrackXC(NP - 1)
ey1 = TrackYC(NP - 1)
For zalpha = 180 To 270
   TrackYC(NP) = ey1 - Rad * Sin((zalpha - 180) * dtr#)
   TrackXC(NP) = ex1 - Rad * (1 - Cos((zalpha - 180) * dtr#))
   TrackANG(NP) = zalpha
   NP = NP + 1
Next zalpha

'----------------------
'Bottom 270 deg
J = Marg
For I = FSWH - Marg - Rad To Marg + Rad Step -1
   TrackXC(NP) = I
   TrackYC(NP) = J
   TrackANG(NP) = 270
   NP = NP + 1
Next I
'Bottom left 270 - 360 deg
ex1 = TrackXC(NP - 1)
ey1 = TrackYC(NP - 1)
For zalpha = 270 To 360
   TrackYC(NP) = ey1 + Rad * (1 - Cos((zalpha - 270) * dtr#))
   TrackXC(NP) = ex1 - Rad * Sin((zalpha - 270) * dtr#)
   TrackANG(NP) = zalpha
   If zalpha = 360 Then TrackANG(NP) = 0
   NP = NP + 1
Next zalpha


MAXTrackPts = NP - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Over-done exit !!?
Done = True    'end loop
Erase FloorSurf, ProjSurf
Erase TrackXC, TrackYC, TrackANG
Unload Me
End
End Sub

Private Sub ShowProjSurf()
      
FormWidth& = Me.Width \ Screen.TwipsPerPixelX
FormHeight& = Me.Height \ Screen.TwipsPerPixelY
   
bm.bmiH.biwidth = PSWH
bm.bmiH.biheight = PSWH

'Stretch byte-array to Form
'NB The ByVal is critical in this! Otherwise big memory leak!
succ& = StretchDIBits(Me.hdc, _
0, 0, _
FormWidth& - 8, FormHeight& - 40, _
0, 0, _
PSWH, PSWH, _
ByVal bmp.ptrProjSurf, bm, _
0&, SRCCOPY)

'Public Const DIB_PAL_COLORS = 1 '  color table in palette indices
'Public Const DIB_RGB_COLORS = 0 '  color table in RGBs


End Sub

Private Sub SetOptions()
'Global ControlAutoManual As Boolean
'Global CodeType As Boolean
'Global SurfaceType

ControlAutoManual = False  ' Manual
mnuAutomatic.Checked = False
mnuManual.Checked = True

ClrSurfaceChecks
SurfaceType = 0            ' Pyramids
mnuPyramids.Checked = True

CodeType = True            ' VB code
mnuVB6.Checked = True
mnuMachineCode.Checked = False

ClrPaletteChecks
HalfSine.Checked = True

End Sub


'### PALETTES ###########################################################

Private Sub ClrPaletteChecks()
HalfSine.Checked = False
Grey.Checked = False
Sine.Checked = False
Blues.Checked = False
Green.Checked = False
Dunes.Checked = False
Earth.Checked = False
End Sub

Private Sub Halfsine_Click()
ClrPaletteChecks
PalSpec$ = PathSpec$ & "Halfsine.PAL"
ReadPAL PalSpec$
HalfSine.Checked = True
End Sub
Private Sub Grey_Click()
ClrPaletteChecks
PalSpec$ = PathSpec$ & "Grey.PAL"
ReadPAL PalSpec$
Grey.Checked = True
End Sub
Private Sub Sine_Click()
ClrPaletteChecks
PalSpec$ = PathSpec$ & "Sine.PAL"
ReadPAL PalSpec$
Sine.Checked = True
End Sub
Private Sub Blues_Click()
ClrPaletteChecks
PalSpec$ = PathSpec$ & "Blues.PAL"
ReadPAL PalSpec$
Blues.Checked = True
End Sub
Private Sub Green_Click()
ClrPaletteChecks
PalSpec$ = PathSpec$ & "Green.PAL"
ReadPAL PalSpec$
Green.Checked = True
End Sub
Private Sub Dunes_Click()
ClrPaletteChecks
PalSpec$ = PathSpec$ & "Dunes.PAL"
ReadPAL PalSpec$
Dunes.Checked = True
End Sub
Private Sub Earth_Click()
ClrPaletteChecks
PalSpec$ = PathSpec$ & "Earth.PAL"
ReadPAL PalSpec$
Earth.Checked = True
End Sub

'#### MENU ITEMS #######################################

Private Sub mnuStart_Click()

STARTACTION

End Sub

Private Sub mnuStop_Click()
Done = Not Done
Do
   DoEvents
Loop Until Done = False
End Sub

Private Sub mnuAutomatic_Click()
Done = True
ControlAutoManual = True
mnuAutomatic.Checked = True
mnuManual.Checked = False
End Sub
Private Sub mnuManual_Click()
ControlAutoManual = False
mnuAutomatic.Checked = False
mnuManual.Checked = True
End Sub


' SURFACES

Private Sub ClrSurfaceChecks()
mnuPyramids.Checked = False
mnuPlasma.Checked = False
mnuGreyClouds.Checked = False
mnuMoonlight.Checked = False
mnuDayClouds.Checked = False
mnuHills.Checked = False
mnuDunes.Checked = False
mnuPlateaux.Checked = False
hlim = 255  ' Default
End Sub

Private Sub mnuPyramids_Click()
Done = True
SurfaceType = 0
ClrSurfaceChecks
mnuPyramids.Checked = True
' FloorCASTER defaults
zdslope = 0.05    ' Starting slope of ray for each column scanned.
zhtmult = 2       ' Voxel ht multiplier.
zslopemult = -100 ' Slope multiplier.  Larger -ve values look
zvp = 256         ' Starting height of viewer.  Higher values lift

End Sub
Private Sub mnuPlasma_Click()
Done = True
SurfaceType = 1
ClrSurfaceChecks
mnuPlasma.Checked = True
' FloorCASTER defaults
zdslope = 0.05    ' Starting slope of ray for each column scanned.
zhtmult = 2       ' Voxel ht multiplier.
zslopemult = -100 ' Slope multiplier.  Larger -ve values look
zvp = 256         ' Starting height of viewer.  Higher values lift
End Sub

' MACHINE CODE SURFACES
Private Sub mnuGreyClouds_Click()
Done = True
SurfaceType = 2
ClrSurfaceChecks
mnuGreyClouds.Checked = True
Grey_Click
zhtmult = 0.5
End Sub
Private Sub mnuMoonlight_Click()
Done = True
SurfaceType = 3
ClrSurfaceChecks
mnuMoonlight.Checked = True
Sine_Click
zhtmult = 0.5
End Sub
Private Sub mnuDayClouds_Click()
Done = True
SurfaceType = 4
ClrSurfaceChecks
mnuDayClouds.Checked = True
Blues_Click
zhtmult = 0.5
End Sub
Private Sub mnuHills_Click()
Done = True
SurfaceType = 5
ClrSurfaceChecks
mnuHills.Checked = True
Green_Click
zhtmult = 0.5
End Sub
Private Sub mnuDunes_Click()
Done = True
SurfaceType = 6
ClrSurfaceChecks
mnuDunes.Checked = True
Dunes_Click
zhtmult = 1#
End Sub
Private Sub mnuPlateaux_Click()
Done = True
SurfaceType = 7
ClrSurfaceChecks
mnuPlateaux.Checked = True
Earth_Click
hlim = 180
zhtmult = 1#
End Sub

' SMOOTHING
Private Sub DoSmoothing_Click()
   Smooth = True
   DoSmoothing.Checked = True
   NoSmoothing.Checked = False
End Sub
Private Sub NoSmoothing_Click()
   Smooth = False
   DoSmoothing.Checked = False
   NoSmoothing.Checked = True
End Sub

' CODE
Private Sub mnuVB6_Click()
CodeType = True            ' VB code
mnuVB6.Checked = True
mnuMachineCode.Checked = False
End Sub
Private Sub mnuMachineCode_Click()
CodeType = False           ' Machine code
mnuVB6.Checked = False
mnuMachineCode.Checked = True
End Sub


Private Sub STARTACTION()

SetCursorPos 200, 200   ' To shift cursor out of the way

Me.MousePointer = vbHourglass

' Set up in Form Load
' PSWH = 256 etc
' FSWH = 2048 + 16

LAH = 2 * PSWH   ' Look ahead start distance eg 512, etc

' ReDim ZeroSurf(PSWH, PSWH)
' Set up ZeroSurf to be used to set background
' color of ProjSurf
For J = 1 To PSWH
For I = 1 To PSWH
   ZeroSurf(I, J) = 2      ' Color number 2 in palette
Next I
Next J

ReDim FloorSurf(FSWH, FSWH)
   
Randomize
Rand& = 255 * Rnd

MakeFloorSurf  ' This checks if VB or
               ' MC with RunMCode 1 (Plasmas) and RunMCode 2 (Smoothing)

' Surfaces Pyramids, Plasma, Hills, Dunes & Plateau have a Track
' others (ie Grey clouds, Moonlight & Day clouds don't).

If SurfaceType < 2 Or SurfaceType > 4 Then
   MakeTrack  ' Also fills TrackXC(),TrackYC() & TrackANG()
End If

MousePointer = vbDefault
Caption = "TRACKING"

'----------------------------------------------------------
If ControlAutoManual = True Then       ' AUTOMATIC TRACKING NO KEYS
                                       ' with Timer
   zT! = Timer
   For nn = 1 To MAXTrackPts Step 5
   
      ' Pick up values from Track
      ex = TrackXC(nn)
      ey = TrackYC(nn)
      TANG = TrackANG(nn)
   
      ' Show angle
      Caption = TANG & " deg"
      
      CopyMemory ProjSurf(1, 1), ZeroSurf(1, 1), PSWH * PSWH
      
      '-------------------------
      If CodeType = True Then
         
         FloorCASTER    ' Caster routine
      
      Else
         
         RunMCode 0     'Machine code RayCASTER
      
      End If
      '-------------------------
      
      ShowProjSurf   ' Show ProjSurf on screen
   
      DoEvents
      
      If ControlAutoManual = False Then Exit For
   
   Next nn
  
  ' Show time to go round Track
   Caption = "  Time = " & Str$(Int(Timer - zT!))
   zT! = 0

'----------------------------------------------------------
ElseIf ControlAutoManual = False Then    'USER TRACKING

   Done = False
   
   ' Starting values for USER control
   TANG = 0
   ex = Marg
   ey = Marg + Rad
    
    ' Test if any keying
   Do
      
      Form1.SetFocus
      
      If keyLeft = True Then
         TANG = TANG - 4
         If TANG < 0 Then TANG = 360
      End If
      If keyRight = True Then
         TANG = TANG + 4
         If TANG > 360 Then TANG = 0
      End If
      
      ' Give info   (nb will go a bit fasrer if omitted)
      zS = Round(zdslope, 3)
      zH = Round(zhtmult, 3)
      'zH = Round(zHt, 3)
      
      a$ = Str$(TANG) & " deg  Speed=" & Str$(Int(zEL)) & " exy=" & Str$(ex) & Str$(ey)
      a$ = a$ + " Ht=x" & Str$(zH) & " Slp=" & Str$(zS) & " Vp=" & Str$(zvp) & " LAH=" & Str$(LAH)
      'a$ = a$ + " ixyr=" & Str$(ixr) & Str$(iyr)
      Caption = a$
      
      ex = ex + zEL * Sin(TANG * dtr#)
      ey = ey + zEL * Cos(TANG * dtr#)
      
      Select Case SurfaceType
      Case 0    'Pyramids - halts at edge of FloorSurf()
         'Pyramid walls - proximity  Trial & error
         Bdr = 450 '511
         If OutSideRect(ex, ey, Bdr, FSWH - Bdr, FSWH - Bdr, Bdr) Then
            ex = ex - zEL * Sin(TANG * dtr#)
            ey = ey - zEL * Cos(TANG * dtr#)
            zEL = 0
         End If
      
      Case Else     'Plasma, Clouds, Hills etc - this rolls over edges of FloorSurf()
      
         'Keep ex,ey on FloorSurf()
         If ex > (FSWH - 9) Then ex = 9: ey = ey - 1.41 * Cos(TANG * dtr#)
         If ex < 9 Then ex = (FSWH - 9): ey = ey - 1.41 * Cos(TANG * dtr#)
         
         If ey > (FSWH - 9) Then ey = 9: ex = ex - 1.41 * Sin(TANG * dtr#)
         If ey < 9 Then ey = (FSWH - 9): ex = ex - 1.41 * Sin(TANG * dtr#)
         
      End Select
            
      CopyMemory ProjSurf(1, 1), ZeroSurf(1, 1), PSWH * PSWH

      '-------------------------
      If CodeType = True Then
      
         FloorCASTER    ' Caster routine
      
      Else
         
         RunMCode 0      'Machine code RayCASTER
      
      End If
      '-------------------------
   
      If SurfaceType < 2 Or SurfaceType > 4 Then   'ie NOT Clouds
      If Collision = 1 Then
         'Back off and stop  (NB the 2 * gives a slight sideways motion)
         ex = ex - 2 * zEL * Sin(TANG * dtr#)
         ey = ey - 2 * zEL * Cos(TANG * dtr#)
         zEL = 0
         Collision = 0
         Picture1.Visible = True    'Show RED
      Else
         Picture1.Visible = False
      End If
      End If
      
      ShowProjSurf   ' Show ProjSurf on screen
   
      DoEvents
      
   Loop Until Done
   
End If
'----------------------------------------------------------

End Sub

Private Sub RunMCode(OpCode&)

' OpCode& = 0 MC FloorCASTER
' OpCode& = 1 MC Big Plasmas
' OpCode& = 2 MC Plus Smoothing

'=========================================================
' INPUT DATA for BITDATA
'   LAH As Long             ' Look ahead length L       16 LAH        'Input
'   ex As Long              ' eye x-coord               20 ex         'Input
'   ey As Long              ' eye y-coord               24 ey         'Input
'   TANG As Long            ' Track angle 0-359         28 TANG       'Input
'   Collision As Long       ' 0/1 no/yes collision      32 Collision  'Return
'   zvp As Single           ' eye z-coord               36 zvp        'Input
'   zdslope As Single       ' Ray slope                 40 zdslope    'Input
'   zslopemult As Single    ' Ray slope multiplier      44 zslopemult 'Input
'   zhtmult As Single       ' Objects height multiplier 48 zhtmult    'Input
'   hlim As Long            ' Height limit for plateaux 52 hlim       'Input Default(255)
'=========================================================
  
   bmp.LAH = LAH
   bmp.ex = ex
   bmp.ey = ey
   bmp.TANG = TANG
   bmp.zvp = zvp
   bmp.zdslope = zdslope
   bmp.zslopemult = zslopemult
   bmp.zhtmult = zhtmult
   bmp.hlim = hlim
   
   res& = CallWindowProc(ptMCode&, ptBITDATA&, Rand&, 2&, OpCode&)

   Rand& = res&
   
   Collision = bmp.Collision
   
End Sub



