Attribute VB_Name = "Graphicslib"


'Function GDILoadBitmapIntoDC

' **********************************************************
' GDI Helper function: Loads a bitmap from file and selects
' it into a memory DC.
' **********************************************************

'Sub GDIClearDCBitmap

' **********************************************************
' GDI Helper function: Goes through the steps required
' to clear up a bitmap within a DC.
' **********************************************************

'Sub TransparentDraw

'Function CreateFromFile



Public Type spritemapthingy
   cmap As cSpriteBitmaps
  'Actual frames for sprites
   FileName As String
End Type

Public sprite() As cSprite
Public spritefile() As String
'which file the sprite uses

Public spritemaps(500) As spritemapthingy

Public sprt As cSpriteBitmaps
Public sprt2 As cSprite
Public cStage As cBitmap

Public uprdown As Integer
Public bill As Integer
Public lastsprite As Long
Public lastmap As Integer
Public spritesloaded As Byte
Public backgr As cSpriteBitmaps
'Public backgrmap As cSpriteBitmaps

Dim pixooz() As Long
'Dim pixooz2() As Long

Private Sub Command1_Click()
draw
End Sub

'Private Sub Form_Load()

'Set sprt = New cSpriteBitmaps

'Set sprt2 = New cSprite

'CreateSpriteResource sprt, App.Path & "\Fighter.bmp", 1, 1, RGB(0, 0, 0)
'CreateSprite sprt, sprt2

'Set cStage = New cBitmap
'public bux As Long
'bux = 1
'sprt2.Cell = bux

'cStage.CreateAtSize 40, 40
'cStage.CreateFromFile App.Path & "\Stuff1.bmp"
'cStage.RenderBitmap Form1.hDC, 0, 0


'End Sub



Private Sub CreateSpriteResource( _
        ByRef cR As cSpriteBitmaps, _
        ByVal sFile As String, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal lTransColor As Long _
    )
    Set cR = New cSpriteBitmaps
    cR.CreateFromFile sFile, cX, cY, , lTransColor
End Sub
Private Sub CreateSprite( _
        ByRef cR As cSpriteBitmaps, _
        ByRef cS As cSprite, _
        hDC As Long _
    )
    Set cS = New cSprite
    cS.SpriteData = cR
    cS.Create hDC
End Sub

Sub draw()
'For a = 0 To 100


'Dim lHDC As Long
'lHDC = Me.hDC

cStage.RenderBitmap Form1.hDC, 0, 0

bill = bill + 1




'sprt2.TransparentDraw Picture1.hDC, sprt2.X, sprt2.Y, sprt2.Cell, False

'For a = 0 To 5
sprt2.RestoreBackground cStage.hDC
sprt2.x = 15
sprt2.y = sprt2.y + uprdown
If sprt2.y > 200 Then uprdown = -15
If sprt2.y < 5 Then uprdown = 15
sprt2.StoreBackground cStage.hDC, sprt2.x, sprt2.y
'm''sprt2.TransparentDraw picBuffer, sprt2.X, sprt2.Y, sprt2.cell, False
sprt2.TransparentDraw picBuffer.GetDC, sprt2.x, sprt2.y, sprt2.cell, False
'cShip.StageToScreen lHDC, cStage.hDC
'sprt2.RestoreBackground cStage.hDC
'Next a

sprt2.StageToScreen Form1.hDC, cStage.hDC
DoEvents
'Next a


End Sub

Function newsprite(hDC As Long, FileName As String, x As Long, y As Long, Optional Xframes = 1, Optional Yframes = 1)
ChDir App.Path
MsgBox "Obsolete function newsprite called"
Stop
spritesloaded = 0
If lastsprite <= 1 Then lastsprite = 1: ReDim sprite(1 To lastsprite) As cSprite: ReDim spritefile(1 To lastsprite) As String
5
For a = 1 To 500
    If spritemaps(a).FileName = FileName Then
'    totalsprites = totalsprites + 1
    CreateSprite spritemaps(a).cmap, sprite(lastsprite), hDC: spritefile(lastsprite) = FileName: newsprite = lastsprite: sprite(lastsprite).cell = 1: sprite(lastsprite).x = x: sprite(lastsprite).y = y
    lastsprite = lastsprite + 1
    ReDim Preserve sprite(1 To lastsprite) As cSprite
    ReDim Preserve spritefile(1 To lastsprite) As String
    spritesloaded = 1
    Exit Function
    End If
Next a

CreateSpriteResource spritemaps(lastmap).cmap, FileName, Xframes, Yframes, RGB(0, 0, 0)
spritemaps(lastmap).FileName = FileName
lastmap = lastmap + 1
GoTo 5

End Function

Function drawmain(desthDC As Long)
'cStage.CreateFromPicture backgr
'backgrmap.TransparentDraw cStage.hdc, 0, 0, 0
Static drawing As Byte
If drawing = 1 Then Exit Function
drawing = 1
'backgr.DirectBltSprite cStage.hdc, 0, 0, 0, True
'backgr.TransparentDraw cStage.hdc, 0, 0, 0, True
For a = 1 To lastsprite
    If Not spritefile(a) = "" Then

'    sprite(a).RestoreBackground cStage.hdc
'    sprite(a).StoreBackground cStage.hdc, sprite(a).x, sprite(a).y
    sprite(a).TransparentDraw cStage.hDC, sprite(a).x, sprite(a).y, sprite(a).cell, True
'''    sprite(a).TransparentDraw picBuffer, sprite(a).X, sprite(a).Y, sprite(a).Cell, True
'    cShip.StageToScreen desthHDC, cStage.hdc

    End If
Next a

cStage.RenderBitmap desthDC, 1, 1
drawing = 0
End Function

Function makebacksprite(filen)

'CreateSpriteResource backgrmap, filen, 0, 0, RGB(0, 0, 0)
'CreateSprite backgrmap, backgr, cStage.hdc
End Function

Sub recolor(ByVal color1 As Long, ByVal color2 As Long, pc As PictureBox)

'replaces a color in a picturebox image.
'To use in a sprite, use the following two lines (One of which is already done here)
'Picture1.Picture = Picture1.Image
'spritemaps(1).cmap.CreateFromPicture PC.Picture, 1, 1, , RGB(0, 0, 0)
'The picturebox must autosize.

pc.ScaleMode = vbPixels
'pc.AutoRedraw = True
'If pc.Width > 500 Or pc.ScaleWidth > 500 Then Stop
For a = 0 To pc.Width
    For b = 0 To pc.Height
        If pc.Point(a, b) = color1 Then pc.PSet (a, b), color2
    Next b
Next a

'pc.Picture = pc.image
selfassign pc

End Sub

Sub quickrangecolor(ByVal r, ByVal g, ByVal b, pc As PictureBox)

'Simply ANDs a big box of color over the whole picturebox

'r = Int((r + r + 150) / 3)
'g = Int((g + g + 150) / 3)
'b = Int((b + b + 150) / 3)

If r < 20 Then r = 20
If g < 20 Then g = 20
If b < 20 Then b = 20

'pc.AutoRedraw = True
pc.DrawMode = vbMaskPen
'pc.DrawMode = vbandPen
pc.Line (0, 0)-Step(pc.Width, pc.Height), RGB(r * 2, g * 2, b * 2), BF
'pc.Line (0, 0)-Step(pc.Width, pc.Height), RGB(0, 0, 0), BF
selfassign pc
pc.DrawMode = vbCopyPen
End Sub

Sub rangecolor(ByVal r, ByVal g, ByVal bl, pc As PictureBox, Optional ByVal lightness As Single = 0.5, Optional transparency = 0)

If nocoloring = 1 Then Exit Sub 'Mostly for making girl graphics
nodraw = 1

'Replaces all straight greys with the specified color

'Test Quickrangecolor
'quickrangecolor r, g, bl, pc: nodraw = 0
'Exit Sub

pc.ScaleMode = vbPixels
'pc.AutoRedraw = True
'If pc.Width > 500 Or pc.ScaleWidth > 500 Then Stop
For a = 0 To pc.ScaleWidth
    If transparency = 1 Then transc = (a + 1) Mod 2
    For b = 0 To pc.ScaleHeight
    
        If transparency = 1 Then If transc = 1 Then transc = 0 Else transc = 1
    
        col = pc.Point(a, b)
2        If col <= 0 Then GoTo 5
        redd = col Mod 256
        green = Int((col Mod 65536) / 256)
        blue = Int(col / 65536)
        
        If redd = green And green = blue Then
        total = redd
        redd = Int((total * ((r * (lightness + 1)) / 255)) + (total * lightness) / 2): If redd > 255 Then redd = 255
        green = Int((total * ((g * (lightness + 1)) / 255)) + (total * lightness) / 2): If green > 255 Then green = 255
        blue = Int((total * ((bl * (lightness + 1)) / 255)) + (total * lightness) / 2): If blue > 255 Then blue = 255
        'If redd > 80 Then Stop
        'redd = Int(total * (r / 255) * (lightness + 1)): If redd > 255 Then redd = 255
        'green = Int(total * (g / 255) * (lightness + 1)): If green > 255 Then green = 255
        'blue = Int(total * (b / 255) * (lightness + 1)): If blue > 255 Then blue = 255
        
        If redd = 0 And green = 0 And blue = 0 Then blue = 1
        
        If transc = 1 Then redd = 0: green = 0: blue = 0 'Transparency
        
        If redd >= 0 And green >= 0 And blue >= 0 Then pc.PSet (a, b), RGB(redd, green, blue)
        
        End If
        
'        Form1.Command1.Caption = "R" & redd & "G" & green & "B" & blue
'        If pc.Point(a, b) = color1 Then pc.PSet (a, b), color2
5    Next b
Next a

'pc.Picture = pc.image
selfassign pc

nodraw = 0

End Sub

Sub slowrangecolor(ByVal r, ByVal g, ByVal bl, pc As PictureBox, Optional ByVal lightness As Single = 0.5)

'Replaces all straight greys with the specified color
'Slowrangecolor is made to run in the background to prevent heinous skippage

pc.ScaleMode = vbPixels
'pc.AutoRedraw = True
'If pc.Width > 500 Or pc.ScaleWidth > 500 Then Stop
For a = 0 To pc.ScaleWidth
    For b = 0 To pc.ScaleHeight
        col = pc.Point(a, b)
        If col <= 0 Then GoTo 5
        redd = col Mod 256
        green = Int((col Mod 65536) / 256)
        blue = Int(col / 65536)
        If redd = green And green = blue Then
        total = redd
        redd = Int((total * ((r * (lightness + 1)) / 255)) + (total * lightness) / 2): If redd > 255 Then redd = 255
        green = Int((total * ((g * (lightness + 1)) / 255)) + (total * lightness) / 2): If green > 255 Then green = 255
        blue = Int((total * ((bl * (lightness + 1)) / 255)) + (total * lightness) / 2): If blue > 255 Then blue = 255
        If red = 0 And green = 0 And blue = 0 Then blue = 1
        If redd >= 0 And green >= 0 And blue >= 0 Then pc.PSet (a, b), RGB(redd, green, blue)
        End If
        
'        Form1.Command1.Caption = "R" & redd & "G" & green & "B" & blue
'        If pc.Point(a, b) = color1 Then pc.PSet (a, b), color2
        DoEvents
5    Next b
Next a

'pc.Picture = pc.image
selfassign pc

End Sub


Sub distort(centerx As Integer, centery As Integer, Width As Integer, Height As Integer, pc As PictureBox, Optional ByVal lightness As Single = 0.5)

ReDim pixooz(Width) As Long

pc.ScaleMode = vbPixels
'pc.AutoRedraw = True
'If pc.Width > 500 Or pc.ScaleWidth > 500 Then Stop
For b = centery To centery + Height
    
    wd00d = 0
    For a = centerx To centerx - Width Step -1
        'If wd00d = 16 Then Stop
        col = pc.Point(a, b)
        pixooz(wd00d) = col
        wd00d = wd00d + 1
        'If wd00d = 16 Then Stop
        'If wd00d >= UBound(pixooz()) Then Exit For
    Next a
    
    widdo = (b - centery) / Height * diff(b - centery - Height, 0) / 12
    curx = 0
    For a = 0 To UBound(pixooz())

        snarg = snarg + widdo * widdo
6         pc.PSet (curx + centerx, b), pixooz(a): curx = curx - 1
          If snarg > 0 Then snarg = snarg - 1: GoTo 6
          If pixooz(a) = 0 Then Exit For
    Next a
    
    For a = centerx To centerx + Width

        col = pc.Point(a, b)
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
7         pc.PSet (curx + centerx, b), pixooz(a): curx = curx + 1
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
selfassign pc

End Sub

Function halffy(pc1 As PictureBox, pc2 As PictureBox)
'Halves a picture in size

Dim col(3, 2) As Long
'ReDim pixooz(pc1.Width, pc1.Height)
'ReDim pixooz2(pc2.Width, pc2.Height)

For a = 0 To pc1.Width
    For b = 0 To pc1.Height
    
    For c = 1 To 4
    'For d = 0 To 1
    wcolor = pc1.Point(a * 2 + (c Mod 2), b * 2 + Int(c / 3))
    If wcolor = 0 Then pc2.PSet (a, b), 0: GoTo 5
    getrgb wcolor, col(c - 1, 0), col(c - 1, 1), col(c - 1, 2)
    'Next d
    Next c
    
    r = 0
    For c = 0 To 3: r = r + col(c, 0): Next c
    r = r / 4
    
    g = 0
    For c = 0 To 3: g = g + col(c, 1): Next c
    g = g / 4
    
    bl = 0
    For c = 0 To 3: bl = bl + col(c, 2): Next c
    bl = bl / 4
    
    If r < 0 Or g < 0 Or b < 0 Then GoTo 5
    pc2.PSet (a, b), RGB(r, g, bl)
        
5     Next b
Next a


End Function

Function digjunk(pc1 As PictureBox, amt, Optional size = 2)

orange = RGB(238, 156, 0)
brown = RGB(138, 78, 0)

For a = 1 To amt
3   wx = roll(pc1.Width)
    wy = roll(pc1.Height)
    col = pc1.Point(wx, wy)
    If col = 0 Then GoTo 3
    'getrgb col, r, g, b
    pc1.DrawMode = vbMaskPen
    curx = wx
    cury = wy
    For b = 1 To size * 6
        curx = curx + roll(3) - 2
        cury = cury + roll(3) - 2
        pc1.PSet (curx, cury), 0
        
        For c = 1 To 5
            pc1.PSet (curx + roll(2) - roll(2), cury + roll(2) - roll(2)), brown
        Next c
        
        For c = 1 To 20
            pc1.PSet (curx + roll(4) - roll(4), cury + roll(4) - roll(4)), orange
        Next c
        
        'For c = curx - (size + roll(3)) To curx + (size + roll(3))
        '    For d = cury - (size + roll(3)) To cury + (size + roll(3))
        '    pc1.PSet (c, d), orange
        '    Next d
        'Next c
    Next b


Next a

pc1.DrawMode = vbCopyPen

End Function

Function gookup(pb As PictureBox, amt, seed)

Rnd (-5)
randword montype(mon(seed).type).name

pb.DrawMode = vbCopyPen

4 colr = roll(180) + 50: colg = roll(130) + 50: If colr > colg * 2 Or colg > colr * 2 Then GoTo 4
stomcol = RGB(colr, colg, 0)
stomcol2 = RGB(colr / 2, colg / 2, 0)
stomcol3 = RGB(colr * 1.5, colg * 1.5, 0)

Randomize seed

For a = 1 To amt
5   zex = roll(pb.Width)
    zey = roll(pb.Height)
    wcolor = pb.Point(zex, zey)
    If wcolor = 0 Then GoTo 5
    aroll = roll(2)
    
    If aroll = 1 Then
6         colr = roll(150) + 70: colg = roll(130) + 50: If colr > colg + 15 Then GoTo 6
            kulr = RGB(colr, colg, 0)
        For b = 1 To roll(5) + 10
        zex = zex + (roll(3) - 2)
        zey = zey + (roll(3) - 2)
        pb.PSet (zex, zey), kulr
        Next b
        
    End If

    If aroll = 2 Then
        zerk = (turncount + a) Mod 12 + 4
        For b = 1 To zerk
        If b = 1 Then pb.Line (zex - 3, zey + b)-Step(9, 1), stomcol
        If b = 2 Then pb.Line (zex - 2, zey + b)-Step(7, -2), stomcol
        If b = 3 Then pb.Line (zex - 1, zey + b)-Step(5, 0), stomcol
        If b > 3 Then pb.Line (zex, zey + b)-Step(3, 0), stomcol
        Next b
        pb.Line (zex + 2, zey + zerk / 2)-Step(0, zerk - 1), stomcol2
        pb.Line (zex, zey + zerk / 2)-Step(0, zerk / 2), stomcol3
    End If


Next a


End Function
