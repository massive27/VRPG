VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpriteBitmaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

' ======================================================================
' Class:    cSpriteBitmaps
' Filename: cSpriteBitmaps.cls
' Author:   SP McMahon
' Date:     14 July 1998
'
' A class to store a picture clip of sprite bitmaps and an associated
' mask to allow transparent drawing onto a DC.
' ======================================================================

' Memory DC & Bitmap to hold the sprite
Private m_hdcSprite As Long
Private m_hBmpSprite As Long
Private m_hBmpSpriteOld As Long

Public DXS As DirectDrawSurface7

Private DXS2() As DirectDrawSurface7

' Memory DC & Bitmap to hold the mask element of the sprite
Private m_hdcMask As Long
Private m_hBmpMask As Long
Private m_hBmpMaskOld As Long

' Memory DC to build up the sprite and draw it:
Private m_hDCBuild As Long
Private m_hBmpBuild As Long
Private m_hBmpBuildOld As Long

' Variables to store sprite frames, current cell and cell size:
Private m_lDX As Long 'X frame pixel width
Private m_lDY As Long
Private m_lCX As Long 'Number of X frames
Private m_lCY As Long
Private mwidth As Long
Private mheight As Long
Public FileName As String

Private celllookup() As Integer

Dim cKey As DDCOLORKEY
Dim Multframe As Byte

#Const HYDEBUG = 1 'm''

Property Get cells()
cells = m_lCX * m_lCY
End Property

Public Sub TransparentDraw( _
        ByVal destsurface As DirectDrawSurface7, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lCell As Long, _
        Optional ByVal bUseBuildDC As Boolean = True, _
        Optional raster As Long, _
        Optional translevel As Single = 0, _
        Optional x2 As Long = 0, Optional y2 As Long = 0, _
        Optional notrans As Byte = 0, _
        Optional rotation As Long = 0 _
    )
Dim lDX As Long
Dim lDY As Long
Dim lDC As Long
Dim loX As Long
Dim loY As Long

If lCell < 1 Or lCell > m_lCX * m_lCY Then Exit Sub

'On Error GoTo errorhandler:
Dim sdesc As DDSURFACEDESC2
    'Dim mcjunk As DDSURFACEDESC2
    'destsurface.GetSurfaceDesc mcjunk
destsurface.GetSurfaceDesc sdesc
'sdesc.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
'sdesc.lBackBufferCount = 2

If sdesc.lBackBufferCount > 0 Then destsurface.GetAttachedSurface sdesc.ddsCaps

'Stop

If DXS.isLost > 0 Then DXS.restore: MsgBox "Error #1 in Transparentdraw function."

    If nodraw = 1 Then Exit Sub

    If bInit = False Then DXinit Form1.Picture7, Bridge.Picture1

    If raster = 1 Then translevel = 2

    If translevel > 0 Then TranslucentDraw destsurface, lX, lY, lCell, bUseBuildDC, translevel: Exit Sub

    

    ' Get the position of the sprite cell within the data:
    pGetXYForCell lCell, lDX, lDY
    
    Dim ddrval As Long
    Dim r1 As RECT 'The screen size
    Dim r2 As RECT 'The bitmap size

    'Dim cKey As DDCOLORKEY
    'cKey.high = DX.CreateColorRGB(0, 0, 0)
    'cKey.low = cKey.high
    'sprites.SetColorKey DDCKEY_SRCBLT, cKey
    
    'Gets the bounding rect for the entire window handle, stores in r1
    'Call DX.GetWindowRect(picbox.hWnd, r1)

    'xframe = frame Mod 5: If xframe = 0 Then xframe = 5
    'yframe = Int(frame / 5) + 1
    
    'r1=destination coords, r2=source coords
    
    If x2 > m_lDX Then x2 = 0
    If y2 > m_lDY Then y2 = 0
    If y2 > 0 Then y2 = m_lDY - y2
    If x2 > 0 Then x2 = m_lDX - x2
    
    
    
    r2.Top = lDY '+ 400
    r2.Left = lDX ' + 400
    r2.Bottom = lDY + m_lDY - y2 '+ 400
    r2.Right = lDX + m_lDX - x2 '+ 400

    Dim mcjunk As DDSURFACEDESC2
    destsurface.GetSurfaceDesc mcjunk
    If mcjunk.lWidth > 800 Then xadd = mcjunk.lWidth - 800
    If mcjunk.lHeight > 600 Then yadd = mcjunk.lHeight - 600
    
    r1.Left = lX + xadd
    r1.Top = lY + yadd
    r1.Right = lX + m_lDX + xadd - x2
    r1.Bottom = lY + m_lDY + yadd - y2
    
    'If r1.Bottom > mcjunk.lHeight Then r1.Bottom = mcjunk.lHeight
    'If r1.Right > mcjunk.lWidth Then r1.Right = mcjunk.lWidth
    'r1.Bottom = r1.Bottom + 144
    'r1.Right = r1.Right + 96
    'Using Blt instead of Bltfast is essential - do not try and use bltfast.
    'The advantage of using blt is that it resizes the picture to be the same as
    'the picture box, this means that we can resize the window and the code
    'will adapt to fit the new size - even though it will look really ugly
    'when stretched.
    
    'ddrval =
    'Dim doink As DDBLTFX
    'doink.lROP = SRCCOPY
    'doink.lAlphaSrcConst = 100
    
    Dim rec3 As RECT
    
    Dim FX As DDBLTFX
    
    FX.lRotationAngle = rotation * 100
    
    'destsurface.BltFx r1, DXS, r2, DDBLT_WAIT Or DDBLT_KEYSRC Or DDBLT_ROP, doink
    If Multframe = 0 Then
    If notrans = 0 Then blarg = destsurface.blt(r1, DXS, r2, DDBLT_DONOTWAIT Or DDBLT_KEYSRC) 'Or doink
    If notrans = 1 Then destsurface.blt r1, DXS, r2, DDBLT_WAIT
    'Dim bfx As DDBLTFX
    'bfx.lROP = 4096 'SRCPAINT
    'destsurface.BltFx r1, DXS, r2, DDBLT_WAIT Or DDBLT_ROP, bfx
    
    'If notrans = 0 Then blarg = destsurface.BltFast(lX + xadd, lY + yadd, DXS, r2, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY) 'Or doink
    'If notrans = 0 Then blarg = destsurface.BltFast(900, 500, DXS, r2, DDBLTFAST_WAIT Or DDBLTFAST_KEYSRC) 'Or doink
    If blarg = DDERR_UNSUPPORTED Then Debug.Print "Unsupported Error"
    If blarg = DDERR_GENERIC Then Debug.Print "Generic Error"
    If blarg = DDERR_INVALIDRECT Then Debug.Print "Invalid Rectangle Error" ': Stop
    If blarg = DDERR_INVALIDPARAMS Then Debug.Print "Invalid Parameters Error"
    If blarg = DDERR_INVALIDOBJECT Then Debug.Print "Invalid Object Error"
    If blarg = DDERR_WASSTILLDRAWING Then Debug.Print "Was Still Drawing Error"
    
    'If blarg < 0 Then Debug.Print "Error: " & blarg
    If notrans = 1 Then destsurface.BltFast lX + xadd, lY + yadd, DXS, r2, DDBLTFAST_WAIT
    
    Else:
    If notrans = 0 Then destsurface.blt r1, DXS2(lCell), rec3, DDBLT_DONOTWAIT Or DDBLT_KEYSRC
    If notrans = 1 Then destsurface.blt r1, DXS2(lCell), rec3, DDBLT_DONOTWAIT
    
    End If
    'destsurface.blt r1, DXS2(lCell), rec3, DDBLT_WAIT Or DDBLT_KEYSRC
    
    'destsurface.BltFast lX, lY, DXS, r2, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    'destsurface.blttodc hdc, DXS, r2, DDBLT_WAIT Or DDBLT_KEYSRC

    Exit Sub
    
errorhandler:
    'MsgBox "Error in sprite"
Exit Sub
    
    ' Determine where to draw to:
    If (bUseBuildDC) Then
        ' Now we take a copy of the background we are drawing on
        ' into the work DC:
        lDC = m_hDCBuild
        BitBlt lDC, 0, 0, m_lDX, m_lDY, hdcDest, lX, lY, SRCCOPY
        loX = 0
        loY = 0
    Else
        lDC = hdcDest
        loX = lX
        loY = lY
    End If
    
    ' We OR the mask into place, i.e. wherever the mask is black
    ' there is no effem_cTile, but wherever it is white the background
    ' goes completely white:
    BitBlt lDC, loX, loY, m_lDX, m_lDY, m_hdcMask, lDX, lDY, SRCPAINT
    ' Now we AND the sprite into position: wherever the sprite is white,
    ' there is no effem_cTile, but over the whitened mask area the sprite
    ' shows as normal:
    BitBlt lDC, loX, loY, m_lDX, m_lDY, m_hdcSprite, lDX, lDY, SRCAND
    
    If (bUseBuildDC) Then
        ' Now we can blit the combination of the background with the
        ' transparently drawn area onto the Destination:
        BitBlt hdcDest, lX, lY, m_lDX, m_lDY, m_hDCBuild, 0, 0, SRCCOPY
    End If
    
End Sub

Public Function CreateFromPicture( _
        ByRef oPic As PictureBox, _
        ByVal lXImageCount As Long, _
        ByVal lYImageCount As Long, _
        Optional ByRef oMaskPic As StdPicture = Nothing, _
        Optional ByVal lTransColor As Long = &HFFFFFF _
    ) As Boolean
Dim tBM As Bitmap
Dim bContinue As Boolean
Dim ddsd2 As DDSURFACEDESC2
'    On Error GoTo 5
'    GoTo 10
'5     Debug.Print "Picture creation failed, probably out of memory...": Exit Function
'10
    
    'For error trapping
    Open "createerr.txt" For Output As #69
    Write #69, "BLARG!  Currently creating from picture..."
    Close #69
    
    If bInit = False Then DXinit Form1.Picture7, Bridge.Picture1
    
    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lWidth = oPic.ScaleWidth ' / Screen.TwipsPerPixelX
    ddsd2.lHeight = oPic.ScaleHeight '/ Screen.TwipsPerPixelY
    'Dim cKey As DDCOLORKEY
    cKey.high = dX.CreateColorRGB(1, 1, 1)
    cKey.low = dX.CreateColorRGB(0, 0, 0)
    'Set DXS = DX.CreateSurfaceFromFile("C:\VB\ExpansionVRPG\overlays2pink.bmp", ddsd2)
    
    Set DXS = DD.CreateSurface(ddsd2)
    
    zark = DXS.GetDC
    
    'flarg = GDILoadPictureIntoDC( _
    '        False, _
    '        oPic, _
    '        m_hdcSprite, _
    '        m_hBmpSprite, _
    '        m_hBmpSpriteOld _
    '        )
    
    BitBlt zark, 0, 0, oPic.ScaleWidth, oPic.ScaleHeight, oPic.Picture.Handle, 0, 0, SRCCOPY
    'oPic.Render zark, 0, 0, oPic.Width, oPic.Height
    DXS.ReleaseDC zark
    
    'Dim cKey As DDCOLORKEY
    'cKey.high = dX.CreateColorRGB(0, 0, 0)
    'cKey.lOW = cKey.high
    LockMe
    
    'If sSpriteBitmap = "temp.bmp" Then
    gcol = DXS.GetLockedPixel(2, 2) 'Compensate for off-black BG colors in different file formats
    
    
    UnlockMe
    
    'getrgb gcol, r1, g1, b1
    
    cKey.low = transcolor
    cKey.high = transcolor 'dX.CreateColorRGB(r1 + 115, g1 + 115, b1 + 115)

    DXS.SetColorKey DDCKEY_SRCBLT, cKey
    
                ' Set the x and y count for image cells in the sprite:
            m_lCX = lXImageCount
            m_lCY = lYImageCount
            ' The width & Height of the sprite:
            m_lDX = oPic.ScaleWidth \ m_lCX
            m_lDY = oPic.ScaleHeight \ m_lCY
            mwidth = oPic.ScaleWidth
            mheight = oPic.ScaleHeight
            
    splitsurface lXImageCount, lYImageCount
    
    Kill "createerr.txt"
    'ReDim celllookup(1 To lXImageCount * lYImageCount)
    
    Exit Function
    
    
End Function

Public Function CreateFromFile( _
        ByVal sSpriteBitmap As String, _
        ByVal lXImageCount As Long, _
        ByVal lYImageCount As Long, _
        Optional ByVal sMaskBitmap As String = "", _
        Optional ByVal lTransColor As Long = 0 _
    ) As Boolean
Dim tBM As Bitmap
Dim bContinue As Boolean
Dim ddsd2 As DDSURFACEDESC2

    
    On Error GoTo 5
    GoTo 10
5     MsgBox "Error in 'CreateFromFile' function"
10

#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.CreateFromFile, " & sSpriteBitmap) 'm''
#End If 'm''
    
    'm'' Duam use a picturebox for loading and retrieving size of pictures.
    'm'' I added With to reduce the timecall of the object
    
    If sSpriteBitmap = "" Then MsgBox "Error #1 in CreateFromFile function.": Exit Function
    'If Left(sSpriteBitmap, 4) = "kari" Then Stop
    
#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.CreateFromFile try DXinit") 'm''
#End If 'm''

    If bInit = False Then DXinit Form1.Picture7, Bridge.Picture1
    
#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.CreateFromFile getfile") 'm''
#End If 'm''

    origfile = sSpriteBitmap
    If Dir(sSpriteBitmap) = "" Then sSpriteBitmap = getfile(origfile, "Data.pak", , , 1)
    If sSpriteBitmap = "" Then sSpriteBitmap = getfile(Left(origfile, Len(origfile) - 4) & ".gif", "Data.pak")
    
    If Dir(sSpriteBitmap) = "" Or sSpriteBitmap = "" Then gamemsg "Warning:  File not found: " & sSpriteBitmap: sSpriteBitmap = getfile("flail2.gif")
    If sSpriteBitmap = "" Then MsgBox "Error #2 in Transparentdraw function.": Exit Function
    If Not Dir(Left(sSpriteBitmap, Len(sSpriteBitmap) - 4) & ".gif") = "" Then sSpriteBitmap = Left(sSpriteBitmap, Len(sSpriteBitmap) - 4) & ".gif" ': Debug.Print "Using GIF"
    'm'' below is Duam's special code for update. removed for faster loading.
    'm''If Not Dir("c:\vb\expansionvrpg\gifs\" & Left(sSpriteBitmap, Len(sSpriteBitmap) - 4) & ".gif") = "" Then sSpriteBitmap = "c:\vb\expansionvrpg\gifs\" & Left(sSpriteBitmap, Len(sSpriteBitmap) - 4) & ".gif" ': Debug.Print "Using GIF"
    
#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.CreateFromFile loadpicture") 'm''
#End If 'm''

    FileName = sSpriteBitmap
    
    With Form1.Picture1 'm''

    'ChDir App.Path
    'Dim orcpic As StdPicture
    If Dir(sSpriteBitmap) = "" Or sSpriteBitmap = "" Then gamemsg "Warning:  File not found: " & sSpriteBitmap: origbitmap = sSpriteBitmap: sSpriteBitmap = getfile("flail2.gif")
    'm'' Form1.Picture1.Picture = LoadPicture(sSpriteBitmap)
    .Picture = LoadPicture(sSpriteBitmap) 'm''

    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lWidth = .ScaleWidth 'm'' Form1.Picture1.ScaleWidth
    ddsd2.lHeight = .ScaleHeight 'm'' Form1.Picture1.ScaleHeight
    Dim cKey As DDCOLORKEY
    'cKey.high = dX.CreateColorRGB(1, 1, 1)
    'cKey.lOW = dX.CreateColorRGB(0, 0, 0)
    
    'm'' If Not Right(sSpriteBitmap, 4) = ".bmp" Then SavePicture Form1.Picture1.image, "temp.bmp": sSpriteBitmap = "temp.bmp"
    If Not Right(sSpriteBitmap, 4) = ".bmp" Then SavePicture .image, "temp.bmp": sSpriteBitmap = "temp.bmp"
    
#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.CreateFromFile actually call DXCreate") 'm''
#End If 'm''

    Set DXS = DD.CreateSurfaceFromFile(sSpriteBitmap, ddsd2)
    
    LockMe
    
    'If sSpriteBitmap = "temp.bmp" Then
    gcol = DXS.GetLockedPixel(2, 2) 'Compensate for off-black BG colors in different file formats
    
#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.CreateFromFile get transparency and try cleaning") 'm''
#End If 'm''

    #If USELEGACY <> 1 Then
    'm'' experimental : to get rid of the ugly non transparent border, clear it with trans color.
    If sSpriteBitmap <> "dirtback2.bmp" Then 'm'' test fix for heady9 crash
    'Debugger.GD_CleanPic DXS, ddsd2.lWidth, ddsd2.lHeight 'm''
    End If 'm''
    #End If
    
    UnlockMe
    
    'getrgb gcol, r1, g1, b1
    
    cKey.low = transcolor
    cKey.high = transcolor 'dX.CreateColorRGB(r1 + 115, g1 + 115, b1 + 115)

#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.CreateFromFile recoloring") 'm''
#End If 'm''

    If gcol > 0 Then recolor 0, 0, 0, 0, 0, 0, 1

#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.CreateFromFile kill temp.bmp") 'm''
#End If 'm''
    
    If sSpriteBitmap = "temp.bmp" Then Kill "temp.bmp"

    DXS.SetColorKey DDCKEY_SRCBLT, cKey

                ' Set the x and y count for image cells in the sprite:
            m_lCX = lXImageCount
            m_lCY = lYImageCount
            ' The width & Height of the sprite:
            

            mwidth = .ScaleWidth 'm'' Form1.Picture1.ScaleWidth
            mheight = .ScaleHeight 'm'' Form1.Picture1.ScaleHeight

            m_lDX = mwidth \ m_lCX
            m_lDY = mheight \ m_lCY
            
            'DXS.GetSurfaceDesc ddsd2
            'm_lDX = ddsd2.lWidth \ m_lCX
            'm_lDY = ddsd2.lHeight \ m_lCY
            
    End With 'm''
    
#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.CreateFromFile splitsurface") 'm''
#End If 'm''

    splitsurface lXImageCount, lYImageCount
    
    'killbinfiles
    
    'ReDim celllookup(1 To lXImageCount * lYImageCount)
    Exit Function
    
End Function
Private Function pbCreateMaskFromSprite( _
        ByVal bmWidth As Long, _
        ByVal bmHeight As Long, _
        ByVal lTransColor As Long _
    ) As Boolean
Dim lOrigColor As Long
Dim bTransColorIsWhite As Boolean
    
    ' Now create a mask Object.  Here we want
    ' to just copy the bitmap into a monochrome
    ' DC, then invert it:
    If (GDIMakeDCAndBitmap( _
        True, _
        m_hdcMask, m_hBmpMask, m_hBmpMaskOld, _
        bmWidth, bmHeight)) Then
        
        If (lTransColor = &HFFFFFF) Then
            bTransColorIsWhite = True
        End If
        
        If (bTransColorIsWhite) Then
            ' Copy the Bitmap from the source, inverting it
            ' as we go:
            BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, &HFF0062 ' Whiteness
            BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, SRCINVERT
            pbCreateMaskFromSprite = True
        Else
            ' We need to make an non-inverse mask first.  This
            ' is ORed with the original sprite to make the area
            ' with transparent colour = white, then it the mask is
            ' inverted.
            Dim hDCTemp As Long, hBmpTemp As Long, hBmpTempOld As Long
            
            If (GDIMakeDCAndBitmap( _
                True, _
                hDCTemp, hBmpTemp, hBmpTempOld, _
                bmWidth, bmHeight)) Then
                
                ' Set the back colour for the sprite to the
                ' transparent colour (this means it is set to
                ' black when we bitblt to the mono DC):
                lOrigColor = SetBkColor(m_hdcSprite, lTransColor)
                ' Make non-inverse mask:
                BitBlt hDCTemp, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, SRCCOPY
                ' Return back colour of sprite
                SetBkColor m_hdcSprite, lOrigColor
                ' Make sprite white in lTransColor region:
                BitBlt m_hdcSprite, 0, 0, bmWidth, bmHeight, hDCTemp, 0, 0, SRCPAINT
                ' Invert the mask:
                BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, &HFF0062 ' Whiteness
                BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, hDCTemp, 0, 0, SRCINVERT
                
                ' Clear up temporary bitmap:
                GDIClearDCBitmap hDCTemp, hBmpTemp, hBmpTempOld
            
                pbCreateMaskFromSprite = True
            End If
        End If
    End If

End Function
Property Get CellWidth() As Long
    CellWidth = m_lDX
End Property
Property Get CellHeight() As Long
    CellHeight = m_lDY
End Property
Public Sub DirectBltSprite( _
        ByVal lHDC As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lCell As Long _
    )
Dim lDX As Long
Dim lDY As Long
    pGetXYForCell lCell, lDX, lDY
    BitBlt lHDC, lX, lY, m_lDX, m_lDY, m_hdcSprite, lDX, lDY, SRCCOPY
End Sub
Private Sub pGetXYForCell( _
        ByVal lCell As Long, _
        ByRef lX As Long, _
        ByRef lY As Long _
    )
    
    If lCell > 1 And celllookup(lCell, 0) = 0 And celllookup(lCell, 0) = 0 Then celllookup(lCell, 0) = ((lCell - 1) Mod m_lCX) * m_lDX: celllookup(lCell, 1) = ((lCell - 1) \ m_lCX) * m_lDY
    lX = celllookup(lCell, 0)
    lY = celllookup(lCell, 1)
    'lX = ((lCell - 1) Mod m_lCX) * m_lDX
    'lY = ((lCell - 1) \ m_lCX) * m_lDY
End Sub

Private Sub ClearUp()
    ' Here we ensure all our Objects are destroyed:

    'GDIClearDCBitmap m_hdcSprite, m_hBmpSprite, m_hBmpSpriteOld
    'GDIClearDCBitmap m_hdcMask, m_hBmpMask, m_hBmpMaskOld
    'GDIClearDCBitmap m_hDCBuild, m_hBmpBuild, m_hBmpBuildOld
End Sub

Private Sub Class_Initialize()
    If bInit = False Then DXinit Form1.Picture7 ''', Bridge.Picture1
    'Debug.Print "m_cSpriteBitmaps:Initialise"
End Sub

Private Sub Class_Terminate()
    'Debug.Print "m_cSpriteBitmaps:Terminate"
    ClearUp
End Sub

Sub recolor(ByVal r As Byte, ByVal g As Byte, ByVal bl As Byte, Optional ByVal lightness As Single = 0.5, Optional ByVal transparency As Long = 0, Optional ByVal shiny As Long = 0, Optional transonly As Long = 0, Optional tolerance = 8)
'm'' recoloring algo to change the grey part of picture into something shiny
'm'' painfully slow...

Dim a As Long, b As Long 'm'' declaring for faster exec
Dim total As Long, transc As Long, basecol As Long 'm''
Dim redd As Long, green As Long, blue As Long 'm''
Dim targ As Single, sr As Single, sg As Single, sb As Single 'm''



'recolor2 r, g, bl, lightness, transparency, shiny, transonly

'79 If nodraw > 0 Then DoEvents: GoTo 79

'Call makered
'Exit Sub

'If r = g And g = bl And r > 0 Then Stop: Exit Sub

'r = 255
'g = 120
'bl = 10

#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.recolor ... getdiplaymode") 'm''
#End If 'm''

DD.GetDisplayMode display
#If USELEGACY = 1 Then 'm'' duam's old debug stuff
Open "lastcolored.txt" For Output As #69
Write #69, "RGB:", r, g, bl
'Write #69, "Width: " & sdesc.lWidth & " Height:" & sdesc.lHeight
Write #69, "Lightness:" & lightness
Close #69
#End If 'm''
lightness = lightness + 0.2

If Multframe = 1 Then recolorsplit r, g, bl, lightness, transparency, shiny: Exit Sub

If lightness = 0 Then lightness = 0.5

If lightness > 0.5 Then
    cdif = lightness - 0.5
    r = lesser(255, r * (1 + cdif))
    g = lesser(255, g * (1 + cdif))
    bl = lesser(255, bl * (1 + cdif))
End If

If lightness < 0.5 Then
    cdif = lightness + 0.5
    r = greater(0, r * cdif)
    g = greater(0, g * cdif)
    bl = greater(0, bl * cdif)
End If


'If DXS.isLost > 0 Then DXS.restore

If nocoloring = 1 Then Exit Sub 'Mostly for making girl graphics
If r = 0 And g = 0 And bl = 0 And transonly = 0 Then Exit Sub
nodraw = 1
'shiny = 1
'On Error GoTo errorhandler:

'Replaces all straight greys with the specified color

'Shiny 1 means opalescent shiny (The best, generally) whereas 2 means metallic shiny

'Test Quickrangecolor
'quickrangecolor r, g, bl, pc: nodraw = 0
'Exit Sub

'pc.ScaleMode = vbPixels
'pc.AutoRedraw = True
'If pc.Width > 500 Or pc.ScaleWidth > 500 Then Stop

'Lock the surface
Dim elrectum As RECT
Dim sdesc As DDSURFACEDESC2
Dim coller As Long
Dim col As Long

whitecol = GenRGB(255, 255, 255)

'DXS.GetSurfaceDesc sdesc
DXS.Lock elrectum, sdesc, DDLOCK_WAIT, 0 'DDLOCK_NOSYSLOCK or


basecol = DXS.GetLockedPixel(2, 2)

'm'' blue2 = dX.ColorGetBlue(basecol) 'm'' looks useless, so : comment !
'm'' green2 = dX.ColorGetGreen(basecol) 'm'' also
'm'' redd2 = dX.ColorGetRed(basecol) 'm'' also
sr = r 'm'' converts variant, will helps calc faster
sg = g 'm''
sb = bl 'm''

'getrgb basecol, redd2, green2, blue2

If sdesc.lWidth <= 0 Or sdesc.lHeight <= 0 Then MsgBox "Error #1 in recolor function.": Exit Sub

#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.recolor start the main loop") 'm''
#End If 'm''

With DXS 'm'' tweak : implicit call of DXS object in the huge loop below

For a = 0 To sdesc.lWidth - 1
    If transparency = 1 Then transc = (a + 1) Mod 2
    For b = 0 To sdesc.lHeight - 1
    
        If transparency = 1 Then If transc = 1 Then transc = 0 Else transc = 1
    
        col = .GetLockedPixel(a, b) 'm'' DXS. implicit
        
        If b = 1 And a = 1 Then basecol = col
        'getrgb col, redd, green, blue
        getrgb col, blue, green, redd
        'blue = dX.ColorGetBlue(col)
        'green = dX.ColorGetGreen(col)
        'redd = dX.ColorGetRed(col)

        
2        If col <= 0 Then GoTo 5
        If transonly = 1 Then
            If col = basecol Then .SetLockedPixel a, b, 0 'm'' DXS. implicit
            'If diff(redd, redd2) < 3 And diff(green, green2) < 3 And diff(blue, blue2) < 3 Then DXS.SetLockedPixel a, b, 0
            GoTo 5
        End If
        
        'redd = col Mod 256
        'green = Int((col Mod 65536) / 256)
        'blue = Int(col / 65536)
        
        'If redd = green And green = blue And blue = redd Then
        'm'' If diff(redd, green) < tolerance And diff(green, blue) < tolerance And diff(redd, green) < tolerance Then
        If Abs(redd - green) < tolerance And Abs(green - blue) < tolerance And Abs(redd - green) < tolerance Then 'm'' make diff() function inline to avoid asm call jumps
        total = redd
        
        If shiny = 2 Then
            total = total * 2
134            If total > 250 Then total = total - 250: GoTo 134
        End If
        targ = CSng(total / (250) * (0.5 + lightness))
        'm'' redd = r * targ 'm'' see below the 3 lines replacing those
        'm'' green = g * targ
        'm'' blue = bl * targ
        
        redd = sr * targ 'm'' using declared var for faster computation
        green = sg * targ 'm''
        blue = sb * targ 'm''
        
        'redd = 255
        'green = 120
        'blue = 10
        
        If shiny = 1 Then
        If redd > 140 Or green > 140 Or blue > 140 Then
135     redd = redd * 2: If redd > 255 Then redd = redd - 255
136     green = green * 2: If green > 255 Then green = green - 255
137     blue = blue * 2: If blue > 255 Then blue = blue - 255
        Else:
     redd = redd * 5: If redd > 255 Then redd = redd - 255
    green = green * 5: If green > 255 Then green = green - 255
     blue = blue * 5: If blue > 255 Then blue = blue - 255
        'redd = redd - 40: If redd < 0 Then redd = redd + 255 - total
        'green = green - 40: If green < 0 Then green = green + 255 - total
        'blue = blue - 40: If blue < 0 Then blue = blue + 255 - total
        End If
        End If
        'If redd > 80 Then Stop
        'redd = Int(total * (r / 255) * (lightness + 1)): If redd > 255 Then redd = 255
        'green = Int(total * (g / 255) * (lightness + 1)): If green > 255 Then green = 255
        'blue = Int(total * (b / 255) * (lightness + 1)): If blue > 255 Then blue = 255
        
        If redd < 3 Then redd = 3
        If green < 3 Then green = 3
        If blue < 3 Then blue = 3
        If redd > 255 Then redd = 255
        If green > 255 Then green = 255
        If blue > 255 Then blue = 255
        'And green = 0 And blue = 0 Then redd = 3: green = 3: blue = 3
        
        If transc = 1 Then redd = 0: green = 0: blue = 0 'Transparency
        
        'NOTE:  Directdraw surfaces take BGR rather than RGB...though they seem to return RGB...go figure...
        'If blue > 20 Then Stop
        'If display.ddpfPixelFormat.lRGBAlphaBitMask > 0 Then
        'coller = dX.CreateColorRGBA(redd / 16, green / 16, blue / 16, 0)
        'coller = dX.ColorGetBlue(blue) 'Or dX.ColorGetRed(red) Or dX.ColorGetGreen(green)
        coller = GenRGB(redd, green, blue)
        'coller = dX.CreateColorRGB(redd , green, blue)
        'Else:
        'If coller <= 0 Then coller = dX.CreateColorRGB(redd, green, blue)
        'End If
        
        If coller < 0 Then coller = 0
        
        If coller > whitecol Then coller = whitecol
        If transc = 0 Then .SetLockedPixel a, b, coller Else .SetLockedPixel a, b, gcol 'dX.CreateColorRGB(redd, green, blue)  'm'' DXS. implicit
        'If redd >= 0 And green >= 0 And blue >= 0 Then pc.PSet (a, b), RGB(redd, green, blue)
        
        End If
        
'        Form1.Command1.Caption = "R" & redd & "G" & green & "B" & blue
'        If pc.Point(a, b) = color1 Then pc.PSet (a, b), color2
5    Next b
Next a

'If transparency = 1 Then
'    cKey.lOW = gcol
'    cKey.high = gcol 'dX.CreateColorRGB(r1 + 115, g1 + 115, b1 + 115)'

'    DXS.SetColorKey DDCKEY_SRCBLT, cKey
'End If
.Unlock elrectum 'Unlock the surface 'm'' DXS. implicit

End With

nodraw = 0

#If USELEGACY = 1 Then 'm'' old duam debug code
Kill "lastcolored.txt"
#End If 'm''

#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.recolor ended") 'm''
#End If 'm''

Exit Sub
errorhandler: MsgBox "Recolor fucked i'self. "
DXS.Unlock elrectum
End Sub

Public Sub DrawtoDC( _
        ByVal hDC As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lCell As Long, _
        Optional ByVal bUseBuildDC As Boolean = True, _
        Optional raster As Long, _
        Optional sizemult As Single = 1 _
    )

   ' Exit Sub
Dim lDX As Long
Dim lDY As Long
Dim lDC As Long
Dim loX As Long
Dim loY As Long

    If bInit = False Then DXinit Form1.Picture7, Bridge.Picture1

    ' Get the position of the sprite cell within the data:
    pGetXYForCell lCell, lDX, lDY
    
    Dim ddrval As Long
    Dim r1 As RECT 'The screen size
    Dim r2 As RECT 'The bitmap size

    'Dim cKey As DDCOLORKEY
    'cKey.high = DX.CreateColorRGB(0, 0, 0)
    'cKey.low = cKey.high
    'sprites.SetColorKey DDCKEY_SRCBLT, cKey
    
    'Gets the bounding rect for the entire window handle, stores in r1
    'Call DX.GetWindowRect(picbox.hWnd, r1)

    'xframe = frame Mod 5: If xframe = 0 Then xframe = 5
    'yframe = Int(frame / 5) + 1
    
    'r1=destination coords, r2=source coords
    
    r2.Top = lDY
    r2.Left = lDX
    r2.Bottom = lDY + m_lDY
    r2.Right = lDX + m_lDX

    
    r1.Top = lY
    r1.Left = lX
    r1.Bottom = (lY + m_lDY) * sizemult
    r1.Right = (lX + m_lDX) * sizemult
    'r1.Bottom = r1.Bottom + 144
    'r1.Right = r1.Right + 96
    'Using Blt instead of Bltfast is essential - do not try and use bltfast.
    'The advantage of using blt is that it resizes the picture to be the same as
    'the picture box, this means that we can resize the window and the code
    'will adapt to fit the new size - even though it will look really ugly
    'when stretched.
    
    'ddrval =
    DXS.BltToDC hDC, r2, r1
    'destsurface.blt r1, DXS, r2, DDBLT_WAIT Or DDBLT_KEYSRC
    'destsurface.blttodc hdc, DXS, r2, DDBLT_WAIT Or DDBLT_KEYSRC

    Exit Sub
End Sub


Public Function CreateFromSurface( _
        ByRef oPic As DirectDrawSurface7, _
        ByVal lXImageCount As Long, _
        ByVal lYImageCount As Long, _
        Optional ByRef oMaskPic As StdPicture = Nothing, _
        Optional ByVal lTransColor As Long = &HFFFFFF, _
        Optional gscale As Single = 1 _
    ) As Boolean
Dim tBM As Bitmap
Dim bContinue As Boolean
Dim sdesc As DDSURFACEDESC2
Dim sdesc2 As DDSURFACEDESC2
'    On Error GoTo 5
'    GoTo 10
'5     Debug.Print "Picture creation failed, probably out of memory...": Exit Function
'10
    
    If bInit = False Then DXinit Form1.Picture7, Bridge.Picture1
        
    oPic.GetSurfaceDesc sdesc2
    
    sdesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    sdesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    sdesc.lWidth = sdesc2.lWidth
    sdesc.lHeight = sdesc2.lHeight
    Dim cKey As DDCOLORKEY
    cKey.high = dX.CreateColorRGB(1, 1, 1)
    cKey.low = dX.CreateColorRGB(0, 0, 0)
    'Set DXS = DX.CreateSurfaceFromFile("C:\VB\ExpansionVRPG\overlays2pink.bmp", sdesc)
    
    Set DXS = DD.CreateSurface(sdesc)
    
    
    Dim r1 As RECT 'Source
    Dim r2 As RECT 'Destination
    
    DXS.BltColorFill r1, 0
    r1.Left = 0
    r1.Top = 0
    r1.Right = sdesc2.lWidth
    r1.Bottom = sdesc2.lHeight
    
    r2.Left = 0
    r2.Top = 0
    r2.Right = sdesc.lWidth * gscale
    r2.Bottom = sdesc.lHeight * gscale
    
    DXS.blt r2, oPic, r1, DDBLT_WAIT
    
    'Dim cKey As DDCOLORKEY
    'cKey.high = dX.CreateColorRGB(0, 0, 0)
    'cKey.lOW = cKey.high
    LockMe
    
    'If sSpriteBitmap = "temp.bmp" Then
    gcol = DXS.GetLockedPixel(2, 2) 'Compensate for off-black BG colors in different file formats
    
    
    
    UnlockMe
    
    'getrgb gcol, r1, g1, b1
    
    cKey.low = transcolor
    cKey.high = transcolor 'dX.CreateColorRGB(r1 + 115, g1 + 115, b1 + 115)

    DXS.SetColorKey DDCKEY_SRCBLT, cKey
    
                ' Set the x and y count for image cells in the sprite:
            m_lCX = lXImageCount
            m_lCY = lYImageCount
            ' The width & Height of the sprite:
            m_lDX = (sdesc.lWidth * gscale) \ m_lCX
            m_lDY = (sdesc.lHeight * gscale) \ m_lCY
        
            mheight = sdesc.lWidth * gscale
            mwidth = sdesc.lHeight * gscale
        
    splitsurface lXImageCount, lYImageCount
    'ReDim celllookup(1 To lXImageCount * lYImageCount)
        
    Exit Function

End Function

Public Sub TranslucentDraw( _
        ByVal destsurface As DirectDrawSurface7, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lCell As Long, _
        Optional ByVal bUseBuildDC As Boolean = True, _
        Optional translevel As Single _
    )
Dim lDX As Long
Dim lDY As Long
Dim lDC As Long
Dim loX As Long
Dim loY As Long

    If bInit = False Then DXinit Form1.Picture7, Bridge.Picture1

    ' Get the position of the sprite cell within the data:
    pGetXYForCell lCell, lDX, lDY
    
    Dim ddrval As Long
    Dim r1 As RECT 'The screen size
    Dim r2 As RECT 'The bitmap size
    
    r2.Top = lDY
    r2.Left = lDX
    r2.Bottom = lDY + m_lDY
    r2.Right = lDX + m_lDX

    
    r1.Top = lY
    r1.Left = lX
    r1.Bottom = lY + m_lDY
    r1.Right = lX + m_lDX
    
    'ddrval =
    'destsurface.blt r1, DXS, r2, DDBLT_WAIT Or DDBLT_KEYSRC
    
    Dim destdesc As DDSURFACEDESC2
    Dim srcdesc As DDSURFACEDESC2
    
    destsurface.GetSurfaceDesc destdesc
    DXS.GetSurfaceDesc srcdesc
    destsurface.Lock r1, destdesc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
    DXS.Lock r2, srcdesc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
    
    For a = r2.Left To r2.Right
        For b = r2.Top To r2.Bottom
            col = DXS.GetLockedPixel(a, b)
            If col = 0 Then GoTo 5
            a2 = r1.Left + a: b2 = r1.Top + b
            
            col2 = destsurface.GetLockedPixel(a2, b2)
            getrgb col, rd1, gr1, bl
            If rd1 < 0 Then rd1 = 1
            If gr1 < 0 Then gr1 = 1
            If bl1 < 0 Then bl1 = 1
            getrgb col2, rd2, gr2, bl2
            
            rd3 = ((rd1 * translevel) + rd2) / (1 + translevel)
            gr3 = ((gr1 * translevel) + gr2) / (1 + translevel)
            bl3 = ((bl1 * translevel) + bl2) / (1 + translevel)
            
            destsurface.SetLockedPixel a2, b2, GenRGB(rd3, gr3, bl3)
            
5         Next b
    Next a
    
destsurface.Unlock r1
DXS.Unlock r2
    
    'destsurface.blttodc hdc, DXS, r2, DDBLT_WAIT Or DDBLT_KEYSRC

End Sub

Public Sub OrDraw( _
        ByVal destsurface As DirectDrawSurface7, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lCell As Long, _
        Optional ByVal bUseBuildDC As Boolean = True, _
        Optional translevel As Single _
    )
Dim lDX As Long
Dim lDY As Long
Dim lDC As Long
Dim loX As Long
Dim loY As Long

    If bInit = False Then DXinit Form1.Picture7, Bridge.Picture1

    ' Get the position of the sprite cell within the data:
    pGetXYForCell lCell, lDX, lDY
    
    Dim ddrval As Long
    Dim r1 As RECT 'The screen size
    Dim r2 As RECT 'The bitmap size
    
    r2.Top = lDY
    r2.Left = lDX
    r2.Bottom = lDY + m_lDY
    r2.Right = lDX + m_lDX

    
    r1.Top = lY
    r1.Left = lX
    r1.Bottom = lY + m_lDY
    r1.Right = lX + m_lDX
    
    'ddrval =
    'destsurface.blt r1, DXS, r2, DDBLT_WAIT Or DDBLT_KEYSRC
    
    Dim destdesc As DDSURFACEDESC2
    Dim srcdesc As DDSURFACEDESC2
    
    destsurface.GetSurfaceDesc destdesc
    DXS.GetSurfaceDesc srcdesc
    destsurface.Lock r1, destdesc, DDLOCK_WAIT, 0
    DXS.Lock r2, srcdesc, DDLOCK_WAIT, 0
    
    For a = r2.Left To r2.Right
        For b = r2.Top To r2.Bottom
        
            col = DXS.GetLockedPixel(a, b)
            If col = 0 Then GoTo 5
            a2 = r1.Left + a: b2 = r1.Top + b
            
            col2 = destsurface.GetLockedPixel(a2, b2)
            
            destsurface.SetLockedPixel a2, b2, col Or col2
            
5         Next b
    Next a
    
destsurface.Unlock r1
DXS.Unlock r2
    
    'destsurface.blttodc hdc, DXS, r2, DDBLT_WAIT Or DDBLT_KEYSRC

End Sub

Function digjunk(amt, Optional size = 2)

Open "lastdigged.txt" For Output As #69
Write #69, "Something died while digging, baby."
'Write #69, "Width: " & sdesc.lWidth & " Height:" & sdesc.lHeight
Write #69, "Filename:" & FileName
Close #69

'orange = RGB(238, 156, 0)
'brown = RGB(138, 78, 0)

orange = GenRGB(138, 56, 0)
brown = GenRGB(58, 38, 0)

Dim r1 As RECT
Dim sdesc As DDSURFACEDESC2

'DXS.GetSurfaceDesc sdesc
DXS.Lock r1, sdesc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0

zurked = 0

For a = 1 To amt
3   wx = roll(sdesc.lWidth) - 1
    wy = roll(sdesc.lHeight) - 1
    
    If wx >= sdesc.lWidth Or wx <= 0 Or wy >= sdesc.lHeight Or wy <= 0 Then GoTo 3
    col = DXS.GetLockedPixel(wx, wy)
    
    If col = 0 Then GoTo 3
    'getrgb col, r, g, b
    'pc1.DrawMode = vbMaskPen
    curx = wx
    cury = wy
    zurked = zurked + 1
    For b = 1 To size + 2
        'curx = curx + roll(3) - 2
        'cury = cury + roll(3) - 2
        
        direc = roll(4)
        If direc = 1 Then curx = curx + 1
        If direc = 2 Then curx = curx - 1
        If direc = 3 Then cury = cury + 1
        If direc = 4 Then cury = cury - 1
        
        If curx > sdesc.lWidth - 1 Or curx <= 0 Or cury > sdesc.lHeight - 1 Or cury <= 0 Then GoTo 9
        
        DXS.SetLockedPixel curx, cury, 0
        
            nx = curx
            ny = cury
        
        For c = 1 To 3
            nx = nx + roll(3) - roll(3)
            ny = ny + roll(3) - roll(3)
            If nx >= sdesc.lWidth Or nx <= 0 Or ny >= sdesc.lHeight Or ny <= 0 Then GoTo 5
            If DXS.GetLockedPixel(nx, ny) > 0 Then DXS.SetLockedPixel nx, ny, orange
5         Next c
        
            nx = curx
            ny = cury
        
        For c = 1 To 2
            nx = nx + roll(2) - roll(2)
            ny = ny + roll(2) - roll(2)
            If nx >= sdesc.lWidth Or nx <= 0 Or ny >= sdesc.lHeight Or ny <= 0 Then GoTo 7
            If DXS.GetLockedPixel(nx, ny) > 0 Then DXS.SetLockedPixel nx, ny, brown
7        Next c
        
        'For c = curx - (size + roll(3)) To curx + (size + roll(3))
        '    For d = cury - (size + roll(3)) To cury + (size + roll(3))
        '    pc1.PSet (c, d), orange
        '    Next d
        'Next c
9    Next b


10 Next a

DXS.Unlock r1


Kill "lastdigged.txt"
'pc1.DrawMode = vbCopyPen

End Function

Sub recolorall(ByVal r, ByVal g, ByVal bl, Optional ByVal lightness As Single = 0.5, Optional ByVal transparency = 0, Optional shiny = 0, Optional ByVal translucency As Single = 1)

If nocoloring = 1 Then Exit Sub 'Mostly for making girl graphics
nodraw = 1

'Replaces all straight greys with the specified color

'Shiny 1 means opalescent shiny (The best, generally) whereas 2 means metallic shiny

'Test Quickrangecolor
'quickrangecolor r, g, bl, pc: nodraw = 0
'Exit Sub

'pc.ScaleMode = vbPixels
'pc.AutoRedraw = True
'If pc.Width > 500 Or pc.ScaleWidth > 500 Then Stop

'Lock the surface
Dim r1 As RECT
Dim sdesc As DDSURFACEDESC2

DXS.GetSurfaceDesc sdesc
DXS.Lock r1, sdesc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0

For a = 1 To sdesc.lWidth - 1
    If transparency = 1 Then transc = (a + 1) Mod 2
    For b = 1 To sdesc.lHeight - 1
    
        If transparency = 1 Then If transc = 1 Then transc = 0 Else transc = 1
    
        col = DXS.GetLockedPixel(a, b)
2        If col <= 0 Then GoTo 5
        'getrgb col, redd, green, blue
        blue = col Mod 256
        green = Int((col Mod 65536) / 256)
        redd = Int(col / 65536)
        
        'If redd = green And green = blue Then
        total2 = greater(redd, green)
        total2 = greater(blue, total2)
        
        total = (redd + green + blue) / 3 ') + total2 + total2) / 3
        
        If total * 2 < total2 Then total = total2
        
        If shiny = 2 Then
            total = total * 2
134            If total > 250 Then total = total - 250: GoTo 134
        End If
        
        'targ = total / 255
        total3 = (255 - total) / 2
        redd = (r + redd + redd - total3) / 3 ' - total3) '/ 2
        green = (g + green + green - total3) / 3 '- total3) '/ 2
        blue = (bl + blue + blue - total3) / 3 ' - total3) '/ 2
        If redd > 255 Then redd = 255
        If redd < 1 Then redd = 1
        If green > 255 Then green = 255
        If green < 1 Then green = 1
        If blue > 255 Then blue = 255
        If blue < 1 Then blue = 1
        
        
        'redd = ((redd * (translucency - 1)) + (Int((total * ((r * (lightness + 1)) / 255)) + (total * lightness) / 2))) / translucency: If redd > 255 Then redd = 255
        'green = ((green * (translucency - 1)) + (Int((total * ((g * (lightness + 1)) / 255)) + (total * lightness) / 2))) / translucency: If green > 255 Then green = 255
        'blue = ((blue * (translucency - 1)) + (Int((total * ((bl * (lightness + 1)) / 255)) + (total * lightness) / 2))) / translucency: If blue > 255 Then blue = 255
        
        
        If shiny = 1 Then
        If redd > 140 Or green > 140 Or blue > 140 Then
135     redd = redd * 2: If redd > 255 Then redd = redd - 255
136     green = green * 2: If green > 255 Then green = green - 255
137     blue = blue * 2: If blue > 255 Then blue = blue - 255
        Else:
     redd = redd * 5: If redd > 255 Then redd = redd - 255
    green = green * 5: If green > 255 Then green = green - 255
     blue = blue * 5: If blue > 255 Then blue = blue - 255
        'redd = redd - 40: If redd < 0 Then redd = redd + 255 - total
        'green = green - 40: If green < 0 Then green = green + 255 - total
        'blue = blue - 40: If blue < 0 Then blue = blue + 255 - total
        End If
        End If
        'If redd > 80 Then Stop
        'redd = Int(total * (r / 255) * (lightness + 1)): If redd > 255 Then redd = 255
        'green = Int(total * (g / 255) * (lightness + 1)): If green > 255 Then green = 255
        'blue = Int(total * (b / 255) * (lightness + 1)): If blue > 255 Then blue = 255
        
        If redd = 0 And green = 0 And blue = 0 Then redd = 1: green = 1: blue = 1
        
        If transc = 1 Then redd = 0: green = 0: blue = 0 'Transparency
        
        'NOTE:  Directdraw surfaces take BGR rather than RGB...though they seem to return RGB...go figure...
        coller = RGB(blue, green, redd)
        DXS.SetLockedPixel a, b, coller 'dX.CreateColorRGB(redd, green, blue)
        'If redd >= 0 And green >= 0 And blue >= 0 Then pc.PSet (a, b), RGB(redd, green, blue)
        
        'End If
        
'        Form1.Command1.Caption = "R" & redd & "G" & green & "B" & blue
'        If pc.Point(a, b) = color1 Then pc.PSet (a, b), color2
5    Next b
Next a

'pc.Picture = pc.image
'selfassign pc

DXS.Unlock r1 'Unlock the surface

nodraw = 0

End Sub

Sub distort(centerx As Integer, centery As Integer, Width As Integer, Height As Integer)
Dim a As Long 'm'' declare

On Local Error GoTo 110 'm'' will try to handle when no bulge is made (compatibility...)

ReDim pixooz(Width) As Long
Dim Empt As Long 'm'' empty handler - a try for bug fixing

'pc.ScaleMode = vbPixels
'pc.AutoRedraw = True
'If pc.Width > 500 Or pc.ScaleWidth > 500 Then Stop

Dim r1 As RECT
Dim sdesc As DDSURFACEDESC2

DXS.GetSurfaceDesc sdesc
DXS.Lock r1, sdesc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, Empt

If Form1.menu_mod(3).Checked = False Then 'm''
    Debugger.GD_BulgeIt DXS, centerx, centery, Height, Width 'm''
    GoTo 100 'm''
End If 'm''

For b = centery To centery + Height
    
    wd00d = 0
    For a = centerx To centerx - Width Step -1
        'If wd00d = 16 Then Stop
        col = DXS.GetLockedPixel(a, b)
        pixooz(wd00d) = col
        wd00d = wd00d + 1
        'If wd00d = 16 Then Stop
        'If wd00d >= UBound(pixooz()) Then Exit For
    Next a
    
    widdo = (b - centery) / Height * diff(b - centery - Height, 0) / 12
    curx = 0
    For a = 0 To UBound(pixooz())

        snarg = snarg + widdo * widdo
6         DXS.SetLockedPixel curx + centerx, b, pixooz(a): curx = curx - 1
          If snarg > 0 Then snarg = snarg - 1: GoTo 6
          If pixooz(a) = 0 Then Exit For
    Next a
    
    For a = centerx To centerx + Width

        col = DXS.GetLockedPixel(a, b)
        pixooz(a - centerx) = col
        'If col <= 0 Then GoTo 5
        
    Next a
    
    'widdo = -((b - centery) / (Height - centery)) + 0.2
    widdo = (b - centery) / Height * diff(b - centery - Height, 0) / 12
    'widdo=
    'If b - centery > height / 2 Then widdo = (b - centery) / height + (b / 2)
    curx = 0
    For a = 0 To UBound(pixooz())
    
        snarg = snarg + widdo * widdo
7         DXS.SetLockedPixel curx + centerx, b, pixooz(a): curx = curx + 1
          If snarg > 0 Then snarg = snarg - 1: GoTo 7
          If pixooz(a) = 0 Then Exit For
    Next a
    
    'For a = centerx To centerx + width
        
    '    pc.PSet ((a - centerx) * 2 + centerx - 1, b), pixooz(a - centerx)
    '    pc.PSet ((a - centerx) * 2 + centerx, b), pixooz(a - centerx)
        'redd = col Mod 256
        'green = Int((col Mod 65536) / 256)
        'blue = Int(col / 65536)
        'pc.PSet (Int(a * 2 - 1), b), RGB(255, 0, 0)
        'pc.PSet (Int(a * 2), b), RGB(255, 0, 0)

        
        'pc.PSet (a, b), RGB(255, 255, 0)
    'Next a
        

        'If redd >= 0 And green >= 0 And blue >= 0 Then pc.PSet (a, b), RGB(redd, green, blue)
        
        'End If
        
'        Form1.Command1.Caption = "R" & redd & "G" & green & "B" & blue
'        If pc.Point(a, b) = color1 Then pc.PSet (a, b), color2

'5   Next a
 Next b

'pc.Picture = pc.image
100: 'm''
DXS.Unlock r1

Exit Sub 'm''
110: 'm''
'm'' special handler or "no bulge" bug
#If HYDEBUG = 1 Then 'm''
dbmsg ("VRPG2.distort : error happened " & Err.Number & " : " & Err.Description)  'm''
#End If 'm''
Err.clear 'm''
Resume Next 'm''

End Sub

'Property Get Assmonkey() As DirectDrawSurface7
'    Assmonkey = DXS
'End Property

Sub LockMe()

Dim elrectum As RECT
Dim sdesc As DDSURFACEDESC2

'DXS.GetSurfaceDesc sdesc
DXS.Lock elrectum, sdesc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0


End Sub

Sub UnlockMe()

Dim elrectum As RECT
DXS.Unlock elrectum

End Sub

Function splitsurface(xfr, yfr) As Integer
'Returns number of frames

ReDim celllookup(1 To greater(1, xfr * yfr), 0 To 1)

Exit Function

fr = xfr * yfr

splitsurface = fr

ReDim DXS2(1 To fr)


Dim ddsd2 As DDSURFACEDESC2
Dim origdesc As DDSURFACEDESC2

    Dim dest As RECT
    Dim src As RECT

DXS.GetSurfaceDesc origdesc

If fr = 1 Then Multframe = 0: Exit Function
'    Set DXS2(1) = DD.CreateSurface(origdesc)
'    DXS2(1).blt dest, DXS, src, DDBLT_WAIT'

'    DXS2(1).SetColorKey DDCKEY_SRCBLT, cKey
    'Set DXS2(1) = DXS
'    Exit Function
'End If

    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lWidth = origdesc.lWidth / xfr
    ddsd2.lHeight = origdesc.lHeight / yfr
    'Dim cKey As DDCOLORKEY
    cKey.high = dX.CreateColorRGB(1, 1, 1)
    cKey.low = dX.CreateColorRGB(0, 0, 0)

For a = 1 To fr

    X = (a - 1) Mod xfr
    Y = Int((a - 1) / xfr)

    Set DXS2(a) = DD.CreateSurface(ddsd2)
    
    src.Left = ddsd2.lWidth * X
    src.Top = ddsd2.lHeight * Y
    src.Right = src.Left + ddsd2.lWidth
    src.Bottom = src.Top + ddsd2.lHeight
    
    
    'dest.Top = 0
    'dest.Left = 0
    'dest.Right = ddsd2.lWidth
    'dest.Bottom = ddsd2.lHeight
    
    DXS2(a).blt dest, DXS, src, DDBLT_WAIT

    DXS2(a).SetColorKey DDCKEY_SRCBLT, cKey
    
Next a

'Set DXS = Nothing

End Function

Sub recolorsplit(ByVal r As Byte, ByVal g As Byte, ByVal bl As Byte, Optional ByVal lightness As Single = 0.5, Optional ByVal transparency = 0, Optional shiny = 0)

'79 If nodraw > 0 Then DoEvents: GoTo 79

'If DXS.isLost > 0 Then DXS.restore

#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.recolorsplit ... ") 'm''
#End If 'm''

If nocoloring = 1 Then Exit Sub 'Mostly for making girl graphics
If r = 0 And g = 0 And bl = 0 Then Exit Sub
nodraw = 1
'shiny = 1
'On Error GoTo errorhandler:

'Replaces all straight greys with the specified color

'Shiny 1 means opalescent shiny (The best, generally) whereas 2 means metallic shiny

'Test Quickrangecolor
'quickrangecolor r, g, bl, pc: nodraw = 0
'Exit Sub

'pc.ScaleMode = vbPixels
'pc.AutoRedraw = True
'If pc.Width > 500 Or pc.ScaleWidth > 500 Then Stop

'Lock the surface
Dim elrectum As RECT
Dim sdesc As DDSURFACEDESC2
Dim coller As Long
Dim col As Long

For zab = 1 To UBound(DXS2())

DXS2(zab).GetSurfaceDesc sdesc
DXS2(zab).Lock elrectum, sdesc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0

#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.recolorsplit write temp file") 'm''
#End If 'm''

Open "lastcolored.txt" For Output As #69
Write #69, "RGB:", r, g, bl
Write #69, "Width: " & sdesc.lWidth & " Height:" & sdesc.lHeight
Write #69, "Lightness:" & lightness
Close #69

If sdesc.lWidth <= 0 Or sdesc.lHeight <= 0 Then MsgBox "Error #1 in Recolorsplit function.": Exit Sub

For a = 1 To sdesc.lWidth - 1
    If transparency = 1 Then transc = (a + 1) Mod 2
    For b = 1 To sdesc.lHeight - 1
    
        If transparency = 1 Then If transc = 1 Then transc = 0 Else transc = 1
    
        col = DXS2(zab).GetLockedPixel(a, b)
2        If col <= 0 Then GoTo 5
        redd = col Mod 256
        green = Int((col Mod 65536) / 256)
        blue = Int(col / 65536)
        
        If redd = green And green = blue Then
        total = redd
        
        If shiny = 2 Then
            total = total * 2
134            If total > 250 Then total = total - 250: GoTo 134
        End If
        
        redd = Int((total * ((r * (lightness + 1)) / 255)) + (total * lightness) / 2): If redd > 255 Then redd = 255
        green = Int((total * ((g * (lightness + 1)) / 255)) + (total * lightness) / 2): If green > 255 Then green = 255
        blue = Int((total * ((bl * (lightness + 1)) / 255)) + (total * lightness) / 2): If blue > 255 Then blue = 255
        
        If shiny = 1 Then
        If redd > 140 Or green > 140 Or blue > 140 Then
135     redd = redd * 2: If redd > 255 Then redd = redd - 255
136     green = green * 2: If green > 255 Then green = green - 255
137     blue = blue * 2: If blue > 255 Then blue = blue - 255
        Else:
     redd = redd * 5: If redd > 255 Then redd = redd - 255
    green = green * 5: If green > 255 Then green = green - 255
     blue = blue * 5: If blue > 255 Then blue = blue - 255
        'redd = redd - 40: If redd < 0 Then redd = redd + 255 - total
        'green = green - 40: If green < 0 Then green = green + 255 - total
        'blue = blue - 40: If blue < 0 Then blue = blue + 255 - total
        End If
        End If
        'If redd > 80 Then Stop
        'redd = Int(total * (r / 255) * (lightness + 1)): If redd > 255 Then redd = 255
        'green = Int(total * (g / 255) * (lightness + 1)): If green > 255 Then green = 255
        'blue = Int(total * (b / 255) * (lightness + 1)): If blue > 255 Then blue = 255
        
        If redd = 0 And green = 0 And blue = 0 Then redd = 1: green = 1: blue = 1
        
        If transc = 1 Then redd = 0: green = 0: blue = 0 'Transparency
        
        'NOTE:  Directdraw surfaces take BGR rather than RGB...though they seem to return RGB...go figure...
        coller = RGB(blue, green, redd)
        If coller < 0 Then coller = 0
        If coller > RGB(255, 255, 255) Then coller = RGB(255, 255, 255)
        DXS2(zab).SetLockedPixel a, b, coller 'dX.CreateColorRGB(redd, green, blue)
        'If redd >= 0 And green >= 0 And blue >= 0 Then pc.PSet (a, b), RGB(redd, green, blue)
        
        End If
        
'        Form1.Command1.Caption = "R" & redd & "G" & green & "B" & blue
'        If pc.Point(a, b) = color1 Then pc.PSet (a, b), color2
5    Next b
Next a

'pc.Picture = pc.image
'selfassign pc

DXS2(zab).Unlock elrectum 'Unlock the surface

nodraw = 0

Kill "lastcolored.txt"

Next zab

#If HYDEBUG = 1 Then 'm''
dbmsg ("cSpriteBitmaps.recolorsplit ended.") 'm''
#End If 'm''

Exit Sub
errorhandler: MsgBox "Recolor fucked i'self. "

End Sub


