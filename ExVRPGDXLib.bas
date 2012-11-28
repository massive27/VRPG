Attribute VB_Name = "DXLib"
'Note:  For any DirectX applications, you need to go into 'References' under project and add
'the DirectX 7 for Visual Basic thingy.

'Option Explicit

'Main Variables
Public dX As New DirectX7
Public DD As DirectDraw7

'Picture box replacements
Public picDD1 As DirectDrawSurface7
Public picDD2 As DirectDrawSurface7
Public SysSurface As DirectDrawSurface7 'For fasty recoloring

'The first buffer holds a bitmap. The second "Primary" surface
'represents what appears on screen.
Public picBuffer As DirectDrawSurface7
Public Primary As DirectDrawSurface7

Public hwcaps As DDCAPS
Public helcaps As DDCAPS

Public picBuffer2 As DirectDrawSurface7
Public Primary2 As DirectDrawSurface7

'The first desciptor describes the screen
'The second describes the surface that the bitmap will go into.
Dim ddsd1 As DDSURFACEDESC2
Dim ddsd2 As DDSURFACEDESC2
Public display As DDSURFACEDESC2

'Dim sprites As DirectDrawSurface7

'The clipper handles obscured surfaces and stops our application
'from drawing over the top of other windows.
Dim ddClipper As DirectDrawClipper
Dim bufferClipper As DirectDrawClipper

Dim ddClipper2 As DirectDrawClipper
Dim bufferClipper2 As DirectDrawClipper

Public r3 As RECT

Public ddinput As DirectInput
Public ddkeyboard As DirectInputDevice



'A simple initialization flag.
Public bInit As Boolean

Function createDXsurface(X, Y) As DirectDrawSurface7

    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lWidth = Form1.Picture1.ScaleWidth
    ddsd2.lHeight = Form1.Picture1.ScaleHeight

'createDXsurface = DD.CreateSurface(ddsd2)

End Function

Function makesurfdesc(X, Y) As DDSURFACEDESC2

    Dim hooker As DDSURFACEDESC2

    hooker.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    hooker.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    hooker.lWidth = X
    hooker.lHeight = Y
    makesurfdesc = hooker

End Function


Private Sub Form_Load()
    'Start the ball rolling......
    'Init
End Sub

Sub DXinit(picbox As PictureBox, Optional picbox2 As PictureBox)
'On Error GoTo ErrHandler:
    'Initialization procedure
    dbmsg "Initializing DirectX"
    Set DD = dX.DirectDrawCreate("")
    dbmsg "Setting Cooperative Level"
    DD.SetCooperativeLevel Form1.hwnd, DDSCL_NORMAL 'DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
      
    dbmsg "Setting up input objects"
    'Set up keyboard crap
    Set ddinput = dX.DirectInputCreate
    Set ddkeyboard = ddinput.CreateDevice("GUID_SysKeyboard")
    ddkeyboard.SetCommonDataFormat DIFORMAT_KEYBOARD
    ddkeyboard.SetCooperativeLevel Bridge.hwnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    ddkeyboard.Acquire
    
    
    
    'The empty string parameter means to use the active display driver
    
        
    'Indicate this app will be a normal windowed app - not fullscreen.
    'with the same display depth as the current display. This can
    'be very limiting - The end-user can have the desktop in anything
    'from 16 colours to 16 million. To find out what the current depth is
    'use:
    'dim DeskTopBpp as long
    'DesktopBpp = DX.SystemBpp

    
    
    
        
    
    'Indicate that the ddsCaps member is valid in this type
    ddsd1.lFlags = DDSD_CAPS
    'This surface is the primary surface (what is visible to the user)
    'ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    'ddsd1.lBackBufferCount = 2
    'DD.GetCaps hwcaps, helcaps
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE 'Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
    'You're now creating the primary surface with the surface description you just set
    dbmsg "Creating Surface"
    Set Primary = DD.CreateSurface(ddsd1)
    
    'Set Primary2 = DD.CreateSurface(ddsd1)
    
    'Now let's set the second surface description
    ddsd2.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    'This is going to be a plain off-screen surface - ie, to hold a bitmap
    
    ddsd2.lWidth = 1200
    ddsd2.lHeight = 1000
    
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    'Now we create the off-screen surface from the pre-rendered picture
    dbmsg "Creating backbuffers"
    Set picBuffer = DD.CreateSurface(ddsd2) '  DD.CreateSurfaceFromFile("dirtback3.bmp", ddsd2)
    'ddsd2.lWidth = 1200
    'ddsd2.lHeight = 1000
    'Set picBuffer = DD.CreateSurface(ddsd2)
    'Set picBuffer2 = DD.CreateSurfaceFromFile("dirtback2.bmp", ddsd2)
    Set picDD1 = DD.CreateSurface(ddsd2)
    Dim r3 As RECT
    picDD1.BltColorFill r3, 0
    Set picDD2 = DD.CreateSurface(ddsd2)
    'loadsprites
    'Creates the clipper, and attaches it to the picturebox and
    'the primary surface. This is all that has to be done - the clipper
    'itself handles everything else.
    dbmsg "Creating surface clippers"
    Set ddClipper = DD.CreateClipper(0)
    Set bufferClipper = DD.CreateClipper(0)
    
    Dim crect(0) As RECT
    
    crect(0).Top = 0 '0
    crect(0).Left = 0 '0
    crect(0).Right = 800 '800
    crect(0).Bottom = 600 '600
    
    crect(0).Right = 1200 '1200
    crect(0).Bottom = 1000 '1000
    
    bufferClipper.SetClipList 1, crect()
    
    ddClipper.SetHWnd picbox.hwnd
    
        
    Primary.SetClipper ddClipper
    picBuffer.SetClipper bufferClipper
    
    'If Not IsMissing(picbox2) Then
    'bufferClipper2.SetClipList 1, crect()
    'ddClipper2.SetHWnd picbox2.hwnd
    'Primary2.SetClipper ddClipper2
    'picBuffer2.SetClipper bufferClipper2
    'End If
    
    'Yes it has been initialized and is ready to blit
    bInit = True
    
    'Ok now were ready to blit this thing, call the blt procedure
    'One huge advantage of Windowed mode is that you don't
    'have to have a loop, you just call "blt" when you need the
    'picture to be updated.
    dbmsg "Blitting to surface"
    blt picbox
    'InitDXSound Form1.hwnd
    
    Debugger.MoreInit 'm'' for compatibility...
    
    
Exit Sub
ErrHandler:
MsgBox "Unable to initialize DirectDraw - Closing program", vbInformation, "error"
End
End Sub

Sub blt(picbox As PictureBox)
On Error GoTo ErrHand:
    'Has it been initialized? If not let's get out of this procedure
    Dim r1 As RECT 'The screen size
    Dim r2 As RECT 'The bitmap size
    If nodraw = 1 Then Exit Sub
'
    If bInit = False Then Exit Sub
    ddClipper.SetHWnd picbox.hwnd
    'Some local variables
    Dim ddrval As Long

    'Gets the bounding rect for the entire window handle, stores in r1
    Call dX.GetWindowRect(picbox.hwnd, r1)
    
    r2.Top = 0: r2.Left = 0
    r2.Bottom = ddsd2.lHeight
    r2.Right = ddsd2.lWidth
    
    r2.Top = 400: r2.Left = 400
    
    'r2.Right = 400: r2.Bottom =
    
    'Using Blt instead of Bltfast is essential - do not try and use bltfast.
    'The advantage of using blt is that it resizes the picture to be the same as
    'the picture box, this means that we can resize the window and the code
    'will adapt to fit the new size - even though it will look really ugly
    'when stretched.
    ddrval = Primary.blt(r1, picBuffer, r2, DDBLT_WAIT)
    'Primary.Flip Nothing, DDFLIP_WAIT  'DDFLIP_DONOTWAIT
    'Primary.Flip Nothing, DDFLIP_DONOTWAIT
    
Exit Sub
ErrHand:
'MsgBox "There was an error whilst redrawing the screen.", vbCritical, "error"
Err.clear
End Sub

Sub blt2(picbox As PictureBox, surf As DirectDrawSurface7)
On Error GoTo ErrHand:
    'Has it been initialized? If not let's get out of this procedure
    Dim r1 As RECT 'The screen size
    Dim r2 As RECT 'The bitmap size
'
    If bInit = False Then Exit Sub
    
    'Some local variables
    Dim ddrval As Long

    'Gets the bounding rect for the entire window handle, stores in r1
    Call dX.GetWindowRect(picbox.hwnd, r1)
    
    r2.Top = 0: r2.Left = 0
    r2.Bottom = ddsd2.lHeight
    r2.Right = ddsd2.lWidth
    
    'Using Blt instead of Bltfast is essential - do not try and use bltfast.
    'The advantage of using blt is that it resizes the picture to be the same as
    'the picture box, this means that we can resize the window and the code
    'will adapt to fit the new size - even though it will look really ugly
    'when stretched.
    ddrval = surf.blt(r1, picBuffer, r2, DDBLT_WAIT)

Exit Sub
ErrHand:
MsgBox "There was an error while redrawing the screen.", vbCritical, "error"
End Sub


Function clssurface(surface As DirectDrawSurface7)

'Dim sdesc As DDSURFACEDESC2
Dim r1 As RECT
'surface.GetSurfaceDesc sdesc
surface.BltColorFill r1, 0

End Function

