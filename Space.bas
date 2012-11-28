Attribute VB_Name = "Space"
'When you put in new parts, use the setplrsys function
'Do combat overhead armada-style? (Turretted weapons automatically fire at anything in range)
'When ship runs low/out of HP, that's when it is boarded or eaten by monsters

'Public Type startype
'    x As Long
'    y As Long
'    gnum As Byte
'End Type

Public stargraphs(100) As cSpriteBitmaps

Public Type systype
    name As String
    slottype As String
    graphname As String
    desc As String
    eff(1 To 10, 1 To 5) As String
End Type

Public Type plrshipinvtype
    obj(9) As systype
End Type

Public Type sectiontype
    'tiles(11, 11) As Integer
    'ovrlay(11, 11) As Integer
    type As String
    'connectors(1 To 4) As Byte '(What connection nodes are)
    partsobj(1 To 4) As systype  'Object, row
    X As Integer
    Y As Integer
End Type

Public Type planetT
    graphname As String
    X As Long
    Y As Long
    owner As Byte
    secx As Long
    secy As Long 'Bigger X and Y values, for galactic position
    graphnum As Integer
    planetname As String 'Used to find planet data when landing or whatever
    data As objecttype 'Data.name is the planet's type (Starbase, Portal, Planet, Lair etc.)
                       'Owner goes under this data
End Type

Public factions(0 To 30) As factionT

Public Type stationT
    name As String
    chassisfile As String
    roomsX(20) As Byte
    roomsY(20) As Byte
    
End Type

Public planets() As planetT

'Public Type plrshiptype
'    section(10, 10) As sectiontype
'    frametype As systype
'End Type

Public Type weapontype
    name As String
    tohit As Integer
    dice As Integer
    damage As Integer
    Bonus As Long
    shipmult As Single
    biomult As Single
    shieldmult As Single 'Multiplies damage vs. shields
    armormult As Single 'Multiplies target armor
    ammomax As Integer
    ammo As Integer
    special As systype
End Type


Public Type shiptype
    name As String
    owner As Integer '0=?, 1=Player, 2=Enemy
    faction As String 'Faction that owns it
    isplr As Byte
    Class As String
    shieldmax As Long
    shields As Single
    shieldregen As Single
    hp As Long
    hpmax As Long
    armor As Long
    speed As Integer
    turnspeed As Byte 'Higher=slower
    weapons(1 To 12) As systype
    weapondelay(1 To 12) As Byte
    'systems(1 To 48) As systype
    'weapon(1 To 36) As weapontype
    'weaponpref(1 To 3) As Byte
    special As systype
    X As Single
    Y As Single
    xspeed As Single
    yspeed As Single
    graphname As String
    graphnum As Integer
    facing As Integer
    maxspeed As Integer
    swoopcount As Integer 'Used to control swooping behavior
    energy As Single
    energyregen As Single
    energymax As Long
    ismonster As Byte
    instomach As Integer 'What monster ship's stomach this ship is in
    eating As Integer '-1=swooping, -2=player in stomach, other=ship# in stomach
    target As Integer 'Target ship
    r As Byte
    g As Byte
    b As Byte
End Type

Type genroomT 'Station room generator thingy
    seg As String
    rot As Byte
End Type


Public Type factionT
    name As String
    color As Long
    fleet As Single
    size As Single
    resources As Single
    wealth As Single 'How quickly they replenish resources
    diplomacy As Single
    alignment As Integer
    likesplayer As Single 'How much they like the player
    opinions(1 To 50) As Single 'how much they like/dislike the other factions (-1 or more means at war)
    X As Single
    Y As Single
    origx As Single
    origy As Single
    origsize As Single
    conversname As String 'Conversation tree name
    shiptypes(1 To 5) As shiptype '1 is smallest ship, 5 is largest
    NPC As Byte '"NPC" factions will declare war on other factions at random and such (They do not follow the storyline)
End Type


Type shipclass
    map(60, 60) As tiletype
    data As shiptype
End Type

Public Type plrshiptype
    name As String
    Class As String
    'shipdat As shiptype
    'shieldmax As Long
    'shields As Long
    'shieldregen As Single
    'hp As Long
    'hpmax As Long
    armor As Long
    cargomax As Integer
    specsystems(1 To 24) As systype
    'Use the systype name or slottype to denote system type (frame, armor etc.) and systype data fields to
    'determine effects
    weapon() As weapontype
    sections(1 To 48) As sectiontype
    cargoname() As String
    cargoamt() As Long
    weight As Long  'Base weight
    firerate As Single
    fuel As Single
    fuelmax As Single
End Type

Public Type bulletsT
    owner As Single '0 means inactive
    duration As Long
    X As Single
    Y As Single
    movex As Single
    movey As Single
    angle As Integer
    maxspeed As Integer
    graphnum As Byte
    seekingtarget As Integer 'If higher than 0, the weapon will seek
    seekingspeed As Integer
    special As systype
    Maxhits As Integer
End Type

Public Type numbertype
    X As Integer
    Y As Integer
    num As String
    graphnum As Byte
    duration As Byte
End Type

Public Type randomjunkT
    X As Single
    Y As Single
    frame As Single
    animtype As Byte
    graphnum As Integer
    special As String
    animspeed As Single
    age As Integer
End Type

Public rjunk(1 To 100) As randomjunkT 'Random junk--cargo, money, powerups, etc.
'Public rjunkg(1 To 30) As cSpriteBitmaps

Public spacestage As cBitmap
Public plrship As shiptype
Public plrshipdat As plrshiptype
'Public shipsecs() As sectiontype
Public shipnames() As String
'Public randomshipcrap() As String
Public shiptypes() As shiptype

Public news(1 To 10) As newstype
Public newsname As String 'Name of who you're talking to

Public Type randomgraphtype
    g As cSpriteBitmaps
    filen As String
    color As Long
End Type

Public Type startype
    X As Single
    Y As Single
    size As Byte 'Determines how fast it moves away
    graphnum As Integer
End Type

Public Type newstype
    newstype As String
    strings(1 To 5) As String
End Type

Public Type starbasetype
    name As String
    convertsfrom(1 To 5, 1 To 2) As String 'Typename, amount
    convertsto(1 To 5, 1 To 2) As String
    consumes(1 To 5, 1 To 2) As String 'Station's upkeep
    produces(1 To 5, 1 To 2) As String
    cargoname() As String
    cargoamt() As Long
    favoritesell(1 To 3) As String 'What type of saleswomen are most common
    special As systype
End Type

Public starbase As starbasetype

Public stars(1 To 60) As startype

Public rgraphs() As randomgraphtype

Public shipgraphs() As cSpriteBitmaps
Public shipgraphnames() As String
Public shipgraphcolor() As Long
Public projgraphs() As cSpriteBitmaps 'Projectile graphics
Public partstock() As systype

Public ships() As shiptype

Public numbergraphs(1 To 4) As cSpriteBitmaps
Public numbers(1 To 16) As numbertype

'Public cargoname() As String 'Cargo currently being carried
Public smoveX As Single 'Ship's current X movement
Public smoveY As Single 'Ship's current Y movement
Public scurX As Double
Public scurY As Double
Public spaceon As Byte
Public bullets() As bulletsT
Public bulletgraphs(1 To 32) As cSpriteBitmaps
'1: Red Laser, 2: Green Laser, 3: Blue Laser, 4: Violet Laser

Public spacecycle As Byte
Public gamedate As Long '"Days" since the game began
Public firedelay As Byte

Public radarzoom

Public starmapon As Byte
Public starmapx As Integer
Public starmapy As Integer
Public sxoff As Integer
Public syoff As Integer

Public pause As Byte

Function addsys(obj As systype, eff1, Optional eff2 = "", Optional eff3 = "", Optional eff4 = "", Optional eff5 = "", Optional replace = 1)

For a = 1 To 10
    If obj.eff(a, 1) = eff1 Then
        If replace = 0 Then
            obj.eff(a, 2) = Val(eff2) + Val(obj.eff(a, 2))
            obj.eff(a, 3) = Val(eff3) + Val(obj.eff(a, 3))
            obj.eff(a, 4) = Val(eff4) + Val(obj.eff(a, 4))
            obj.eff(a, 5) = Val(eff5) + Val(obj.eff(a, 5))
        Else:
            obj.eff(a, 2) = eff2
            obj.eff(a, 3) = eff3
            obj.eff(a, 4) = eff4
            obj.eff(a, 5) = eff5
        End If
    Exit Function
    End If
    
    If obj.eff(a, 1) = "" Then
            obj.eff(a, 1) = eff1
            obj.eff(a, 2) = eff2
            obj.eff(a, 3) = eff3
            obj.eff(a, 4) = eff4
            obj.eff(a, 5) = eff5
            Exit Function
    End If
    
Next a

End Function

Function getsys(obj As systype, effname, Optional ByVal slot As Integer = 2, Optional default = 0)

For a = 1 To 10
    If obj.eff(a, 1) = effname Then getsys = obj.eff(a, slot): Exit Function
    If obj.eff(a, 1) = "" Then getsys = default: Exit Function
Next a

End Function

Function settohit(ByVal base As Integer, ByRef tohit As Integer, ByRef dambon As Integer, ByRef critbon As Integer, ByRef critmult As Single)
'Takes the base attack bonus and turns it into to-hit percentages, damage bonuses and critical
'hit chance and multiplier

'base = Val(Text2.Text)
dambon = 0
critbon = 0
critmult = 2

tohit = base

'If base > 63 Then dambon = Int(base - 62) / 3
If base > 62 Then dambon = Int(base - 62) / 3
If base > 61 Then critbon = Int(base - 61) / 3
If base > 60 Then tohit = tohit - Int(base - 60) * 0.67
If tohit > 80 Then tohit = tohit - Int(tohit - 80) / 2

If critbon > 0 Then critmult = (critbon / 20 + 4) / 2: critbon = critbon / 2

critbon = (critbon + 10) / 2

tohit = (tohit + 60) / 2

'Text1.Text = "To hit: " & tohit & "%" & vbCrLf & "Damage bonus: +" & dambon & "%" & vbCrLf & "Critical Chance: " & critbon & "%" & vbCrLf & "Critical Multiplier: x" & critmult



End Function

Function getshipamt(valname)
'zarg = 0
'For a = 0 To UBound(plrship.sections())
'    zarg = zarg + getsecamt(plrship.sections(a), valname)
'Next a

'For a = 0 To UBound(plrship.specsystems())
'    zarg = zarg + getsys(plrship.specsystems(a), valname)
'Next a

'getshipamt = zarg
End Function

Function getsecamt(obj As sectiontype, valname)

'zarg = 0
'For a = 1 To 4
'    For b = 1 To 4
'    zarg = zarg + getsys(obj.partsobj(a, b), valname)
'    Next b
'Next a

'getsecamt = zarg

End Function

Function shipshot(weapon As systype, targ As shiptype, Optional ByRef shields) As Long
'Returns damage, and shields if a variable is sent (DO NOT SEND ACTUAL SHIELDS AS THE VARIABLE--Use a placeholder to return the value)

Dim tohit As Integer
Dim dambon As Integer
Dim critbon As Integer
Dim critmult As Single

tohit = 0
dambon = 0
critbon = 0
critmult = 2

'Weapon:  "Weapon", Name, Dice, Damage, Bonus
'settohit getsys(weapon, "Weapon", 3), tohit, dambon, critbon, critmult 'Set tohit and bonuses using weapon accuracy

'If roll(100) < tohit Then Exit Function  'Roll to hit

dmg = rolldice(Val(getsys(weapon, "Weapon", 3)), Val(getsys(weapon, "Weapon", 4))) + Val(getsys(weapon, "Weapon", 5))  'Damage roll
'dmg = dmg * dambon 'Damage bonus
If roll(100) < critbon Then dmg = dmg * critmult  'Critical Hit

shields = 0
If targ.shields > 0 Then
    If dmg * 4 < targ.shields Then shields = lesser(dmg, targ.shields) Else shields = lesser(targ.shields, Int(dmg * 0.75))
    dmg = dmg - shields
End If

targ.shields = targ.shields - shields
targ.hp = targ.hp - dmg
If targ.hp <= 0 Then
'targ.owner = 0
'targ.grapnum = 0
'targ.graphname = 0
killship targ
End If

shipshot = dmg

End Function

Function killship(targ As shiptype, Optional explosion = 1)
If explosion > 0 Then createjunk targ.X, targ.Y, "explosion" & explosion & ".bmp", , , 0.3, 6
targ.hp = targ.hpmax
targ.owner = -1

If targ.eating > 0 Then ships(targ.eating).instomach = 0
If targ.eating = -2 Then plrship.instomach = 0


'if targ.isplr=1 then if targ.instomach>0 then

For a = 1 To roll(6)
createjunk targ.X + roll(50) - 25, targ.Y + roll(50) - 25, "CashIcon.bmp", 1, "Gold:" & Int(targ.hpmax / 10), 0.5
Next a

targ.X = 0
targ.Y = 0

End Function

Function damageship(targ As shiptype, damage As Long)

dmg = damage - targ.armor

shields = 0
If targ.shields > 0 Then
    If dmg * 4 < targ.shields Then shields = lesser(dmg, targ.shields) Else shields = lesser(targ.shields, Int(dmg * 0.75))
    dmg = dmg - shields
End If

targ.shields = targ.shields - shields
targ.hp = targ.hp - dmg
If targ.hp <= 0 Then
'targ.owner = 0
'targ.grapnum = 0
'targ.graphname = 0
killship targ
End If

If dmg > 0 Then addtext dmg, targ.X, targ.Y, 255, 170, 0, , , 250, 0, 0
If shields > 0 Then addtext shields, targ.X, targ.Y - 20, 0, 250, 40, , , 0, 0, 0

End Function

'Sub drawnumbers()

'For a = 1 To UBound(numbers())
'    If numbers(a).duration > 0 Then drawnumber numbers(a).num, numbers(a).graphnum, numbers(a).x, numbers(a).y: numbers(a).duration = numbers(a).duration - 1: numbers(a).y = numbers(a).y - 1
'Next a

'End Sub

'Sub drawnumber(num As String, graphnum, ByVal x, ByVal y)

'If num = "" Then num = "0"
'zug = Len(num)
'x = x - (zug * 5)
'For a = 1 To zug
'    numbergraphs(graphnum).TransparentDraw spacestage.hDC, x, y, Val(Mid(num, a, 1))
'Next a

'End Sub

Sub drawships()

'Run junk that occurs each frame
If pause = 1 Then Exit Sub
If starmapon = 1 Then drawstarmap: Exit Sub


shipmaint plrship
If firedelay > 0 Then firedelay = firedelay - 1

shipAI
'If spacecycle Mod greater(1, 100 - ((plrship.shieldregen * 3) / plrship.shieldmax)) <= 1 Then regainshields plrship

'If plrship.shieldmax >= plrship.shieldregen Then If spacecycle Mod (20 + plrship.shieldmax * 2 / plrship.shieldregen) < 1 Then regainshields plrship
'If plrship.shieldregen > plrship.shieldmax Then If spacecycle Mod greater(1, (15 - plrship.shieldregen / plrship.shieldmax)) < 1 Then regainshields plrship

'On Error Resume Next
Dim r4 As RECT
picBuffer.BltColorFill r4, 0 'RGB(0, 10, 145)
drawstars
drawplanets
drawjunk

If plrship.graphnum = 0 Then loadshipgraph plrship
tries = 0
3    plrship.X = plrship.X + plrship.xspeed
    plrship.Y = plrship.Y + plrship.yspeed
    If tries > 10 Then GoTo 4 'plrship.x = 300: plrship.y = 300
    If plrship.X < 0 Then plrship.xspeed = -plrship.xspeed: plrship.X = plrship.X + 5: tries = tries + 1: GoTo 3
    If plrship.Y < 0 Then plrship.yspeed = -plrship.yspeed: plrship.Y = plrship.Y + 5: tries = tries + 1: GoTo 3
    If plrship.instomach > 0 Then plrship.X = ships(plrship.instomach).X: plrship.Y = ships(plrship.instomach).Y: plrship.facing = ships(plrship.instomach).facing
    drawthingy shipgraphs(plrship.graphnum), plrship.X, plrship.Y, plrship.facing, 1
    
4
For a = 1 To UBound(ships())
    If ships(a).owner = -1 Then GoTo 7
    ships(a).X = ships(a).X + ships(a).xspeed: If ships(a).X < 0 Then ships(a).xspeed = -ships(a).xspeed
    ships(a).Y = ships(a).Y + ships(a).yspeed: If ships(a).Y < 0 Then ships(a).yspeed = -ships(a).yspeed
7 Next a
'

'For a = 1 To 10
    'For b = 1 To 10
    'shipgraphs(plrship.graphnum).TransparentDraw picbuffer, plrship.x + a * 25 - shipgraphs(plrship.graphnum).CellWidth / 2, plrship.y + b * 25 - shipgraphs(plrship.graphnum).CellHeight / 2, plrship.facing
    If plrship.instomach = 0 Then drawthingy shipgraphs(plrship.graphnum), plrship.X, plrship.Y, plrship.facing

    'Next b
'Next a

For a = 1 To UBound(ships())

    If ships(a).owner = -1 Then GoTo 5
    If ships(a).graphnum = 0 Then loadshipgraph ships(a) 'Load ship graphics if they haven't been already
    If diff(ships(a).X, plrship.X) > 1000 Or diff(ships(a).Y, plrship.Y) > 1000 Then GoTo 5
    'shipgraphs(ships(a).graphnum).TransparentDraw picbuffer, ships(a).x - shipgraphs(ships(a).graphnum).CellWidth / 2, ships(a).y - shipgraphs(ships(a).graphnum).CellHeight / 2, greater(ships(a).facing, 1)
    fradd = 0
    If Not ships(a).eating = 0 Then fradd = 18
    If ships(a).eating = -1 Then fradd = 36
    drawthingy shipgraphs(ships(a).graphnum), ships(a).X, ships(a).Y, greater(ships(a).facing + fradd, 1)

5 Next a

drawbullets
checkforhits
drawtexts
'picBuffer.SetForeColor RGB(0, 250, 0)
'picBuffer.drawtext 410, 610, "Shields: " & Int(plrship.shields) & " / " & Int(plrship.shieldmax), False
drawtext "Shields: " & Int(plrship.shields) & " / " & Int(plrship.shieldmax), 10, 210, 0, 250, 0
'picBuffer.SetForeColor RGB(250, 50, 0)
'picBuffer.drawtext 410, 630, "HP: " & Int(plrship.hp) & " / " & Int(plrship.hpmax), False
drawtext "HP: " & Int(plrship.hp) & " / " & Int(plrship.hpmax), 10, 230, 250, 50, 0
drawradar 4, 4, radarzoom '100

picBuffer.drawtext 410, 690, "X: " & Int(plrship.X), False
picBuffer.drawtext 410, 710, "Y: " & Int(plrship.Y), False
drawtext "Energy: " & Int(plrship.energy) & "/" & plrship.energymax, 10, 250, 250, 200, 0
drawtext "Fuel: " & Int(plrshipdat.fuel) & "/" & plrshipdat.fuelmax, 10, 270, 0, 40, 250


drawtext "<M>inimap  <E>ngineering  <X>Exit Bridge  <L>and  <H>yperdrive", 10, 550, 250, 250, 250
drawtext "<C>ommunications", 10, 570, 250, 250, 250

drawwepstats 10, 310

177 blt Bridge.Picture1

End Sub

Sub drawbullets()

For a = 1 To UBound(bullets())
    If bullets(a).owner > 0 Then
    bullets(a).X = bullets(a).X + bullets(a).movex
    bullets(a).Y = bullets(a).Y + bullets(a).movey
    'bullets(a).x = 110000
    'bullets(a).y = 100000
    'bulletgraphs(bullets(a).graphnum).TransparentDraw picbuffer, bullets(a).x - bulletgraphs(bullets(a).graphnum).CellWidth / 2, bullets(a).y - bulletgraphs(bullets(a).graphnum).CellHeight / 2, bullets(a).angle / 20
    drawthingy bulletgraphs(bullets(a).graphnum), bullets(a).X, bullets(a).Y, bullets(a).angle / 20
    bullets(a).duration = bullets(a).duration - 1: If bullets(a).duration < 1 Then bullets(a).owner = 0
    End If
Next a

End Sub

Sub loadshipgraph(ship1 As shiptype)
'Loads ship graphics--assumes 18 directions

If ship1.graphname = "" Then ship1.graphname = "ship1s.bmp"

'Assign pre-existing graphics if already loaded
For a = 1 To UBound(shipgraphs())
    If ship1.graphname = shipgraphnames(a) And RGB(ship1.r, ship1.g, ship1.b) = shipgraphcolor(a) Then ship1.graphnum = a: Exit Sub
Next a

'Else load graphics

zark = UBound(shipgraphs()) + 1
ReDim Preserve shipgraphs(1 To zark)
ReDim Preserve shipgraphnames(1 To zark)
ReDim Preserve shipgraphcolor(1 To zark)

yfr = 1

If ship1.ismonster = 1 Then yfr = 3

'shipgraphs(zark).CreateFromFile ship1.graphname, 18, 1, , 0
makesprite shipgraphs(zark), Form1.Picture1, ship1.graphname, ship1.r, ship1.g, ship1.b, , 18, yfr
shipgraphnames(zark) = ship1.graphname
shipgraphcolor(zark) = RGB(ship1.r, ship1.g, ship1.b)
ship1.graphnum = zark
'shipgraphs(zark).recolorall roll(255), roll(255), roll(255), , , , 2
End Sub

Function loadrandomgraph(filen, Optional r = 0, Optional g = 0, Optional b = 0, Optional cells = 1, Optional ycells = 1)

If filen = "" Then filen = "flail2.bmp"
If Dir(filen) = "" Then filen = "flail2.bmp"

tcolor = RGB(r, g, b)
For a = 1 To UBound(rgraphs())
    If rgraphs(a).filen = filen And rgraphs(a).color = tcolor Then loadrandomgraph = a: Exit Function
Next a

zark = UBound(rgraphs()) + 1
ReDim Preserve rgraphs(1 To zark)
makesprite rgraphs(zark).g, Form1.Picture1, filen, r, g, b, , cells, ycells
rgraphs(a).filen = filen: rgraphs(a).color = tcolor
loadrandomgraph = zark

End Function

Function spaceload()

ReDim ships(1 To 1)
ReDim shipgraphs(1 To 1)
ReDim shipgraphnames(1 To 1)
ReDim shipgraphcolor(1 To 1)
ReDim bullets(1 To 1)
ReDim planets(1 To 1)
ReDim rgraphs(1 To 1)
ReDim partstock(1 To 80)
plrship.graphnum = 0
loadbulletgraphs

End Function


Function loadshipclass(filen)

Open filen For Input As #1

    Do While Not EOF(1)
'        if
            
    Loop

Close #1

End Function

Function getplrsys(effname As String, Optional ByVal slot = 2, Optional multeffname As String = "", Optional multslot = 2)
'Mult goes by percentages

'Get system amount, multiplied by any modifiers from parts on the same circuit
For b = 1 To UBound(plrshipdat.sections())
    For a = 1 To 4
    If Not multeffname = "" Then
        cursys = cursys + (Val(getsys(plrshipdat.sections(b).partsobj(a), effname, slot)) * getsys(plrshipdat.sections(b).partsobj(a), multeffname, multslot, 1)) * greater(getglobal("Global" & multeffname, slot), 1)
    Else:
        cursys = cursys + getsys(plrshipdat.sections(b).partsobj(a), effname, slot)
    End If
    Next a
Next b

For b = 1 To UBound(plrshipdat.specsystems())
    cursys = cursys + getsys(plrshipdat.specsystems(b), effname, slot) * getglobal("Global" & effname, slot)
Next b

getplrsys = cursys

End Function

Function getsection(sec As sectiontype)

'    For a = 1 To 4
'        cursys = cursys + (getsys(sec.partsobj(a), effname, slot) * getsys(sec.partsobj(a), multeffname, multslot, 1)) * getglobal("Global" & multeffname, slot)
'    Next a

End Function

Function getglobal(effname As String, Optional ByVal slot = 2, Optional default = 1)
Static lasteffname As String
Static lastret
If effname = lasteffname Then getglobal = lastret: Exit Function

lasteffname = effname

For b = 1 To UBound(plrshipdat.sections())
    For a = 1 To 4
    cursys = cursys + getsys(plrshipdat.sections(b).partsobj(a), effname, slot)
    Next a
Next b

For b = 1 To UBound(plrshipdat.specsystems())
    cursys = cursys + getsys(plrshipdat.specsystems(b), effname, slot)
Next b

lastred = cursys
getglobal = cursys

End Function

Function clearsys(sysobj As systype, Optional effsonly = 0)
For a = 1 To 10
    For b = 1 To 5
        sysobj.eff(a, b) = ""
    Next b
Next a

If effsonly > 0 Then Exit Function
sysobj.slottype = ""
sysobj.graphname = ""
sysobj.name = ""
End Function

Function setplrsys(partslot)

'For a = 1 To UBound(plrshipdat.sections())

'Next a
'clearsys plrship.systems(a), 1

'For a = 1 To 4
'    combinesys plrship.systems(a), plrshipdat.sections.partsobj(a)
'Next a

End Function

Function combinesys(sys1 As systype, sys2 As systype)
'Adds sys1's stuff to sys2

For a = 1 To 10
addsys sys1, sys2.eff(a, 1), sys2.eff(a, 2), sys2.eff(a, 3), sys2.eff(a, 4), sys2.eff(a, 5), 0
Next a

End Function

Sub accelship(ship1 As shiptype)

calcaccel ship1.facing * 20, 1, 5, ship1.xspeed, ship1.yspeed

End Sub

Sub loadbridge()
Randomize Timer

Static hozer As Byte
If hozer = 1 Then GoTo 15 Else hozer = 1

radarzoom = 100

spaceload
ChDir App.Path & "\Fileholder"
plrship.graphname = "ship4s.bmp"
plrship.X = 300
plrship.Y = 300
plrship.facing = 1
plrship.graphnum = 0
plrship.shieldmax = 500
plrship.shieldregen = 50
plrship.shields = 500
plrship.hpmax = 4000
plrship.hp = 4000
plrship.graphnum = 0
plrship.energymax = 500
plrship.energyregen = 1
plrship.isplr = 1

For a = 1 To 50
    aroll = roll(3)
    Select Case aroll
    'createjunk roll(5000), roll(5000), getgener("CashIcon.bmp", "CargoIcon.bmp", "RedCrystalIcon.bmp", "GreenCrystalIcon.bmp", "PurpleCrystalIcon.bmp"), 1, "Gold:" & roll(500), 0.5
    Case 1: createjunk roll(2000), roll(2000), "CrystalIconRed.bmp", 1, "HP:75", 1, 35
    Case 2: createjunk roll(2000), roll(2000), "CrystalIconGreen.bmp", 1, "Shields:200", 1, 35
    Case 3: createjunk roll(2000), roll(2000), "CashIcon.bmp", 1, "Gold:100", 1
    End Select
Next a

starbase.convertsfrom(1, 1) = "Iron Ore"
starbase.convertsfrom(1, 2) = 15
starbase.consumes(1, 1) = "Xodium Crystals"
starbase.consumes(1, 2) = 30

addcargo "Iron Ore", 500
addcargo "Xodium Crystals", 50
starbasecargo "Iron Ore", 15
starbasecargo "Gold", 50


'For a = 1 To 4
'addsys plrship.weapons(a), "Weapon", "Phaser", 3, 4, 0
'addsys plrship.weapons(a), "WeaponFire", 3, 12, 10
'addsys plrship.weapons(a), "Bulletinfo", 5, 12, 100, 1
'Next a

For a = 1 To 4
plrshipdat.sections(a).type = "Shield Bay"
plrshipdat.sections(a).partsobj(1) = loadpart("Ezotronic Coil")
plrshipdat.sections(a).partsobj(2) = loadpart("Ezotron Regulator")
plrshipdat.sections(a).partsobj(3) = loadpart("Ezotronic Bridge")
plrshipdat.sections(a).partsobj(4) = loadpart("Shield Emitter")
Next a

For a = 5 To 8
plrshipdat.sections(a).type = "Weapon Bay"
plrshipdat.sections(a).partsobj(1) = loadpart("Power Turbine")
plrshipdat.sections(a).partsobj(2) = loadpart("Power Turbine")
plrshipdat.sections(a).partsobj(3) = loadpart("Power Turbine")
If a < 7 Then plrshipdat.sections(a).partsobj(4) = loadpart("Crimson Laser") Else plrshipdat.sections(a).partsobj(4) = loadpart("Heavy Laser")
Next a

setplrship

ReDim ships(1 To 3)

For a = 1 To UBound(ships())
ships(a).graphnum = 0
ships(a).owner = 2
ships(a).graphname = "ship" & roll(24) & "s.bmp"
ships(a).X = roll(800): ships(a).Y = roll(600)
ships(a).speed = 4 + roll(10): ships(a).turnspeed = roll(4) + 1

ships(a).energymax = 500
ships(a).energyregen = 1
ships(a).shieldmax = 50
ships(a).shieldregen = 10
ships(a).hpmax = 50
ships(a).hp = 50
addsys ships(a).weapons(1), "Weapon", "Phaser", 3, 4, 0
addsys ships(a).weapons(1), "WeaponFire", 3, 12, 10
addsys ships(a).weapons(1), "Bulletinfo", roll(16), 12, 100, 1

Next a

ReDim Preserve ships(1 To 8)

For a = 4 To 8
    ships(a).ismonster = 1
    ships(a).graphnum = 0
    ships(a).owner = 2
    'ships(a).graphname = "girlspacemons.gif"
    ships(a).graphname = "spacemon-amoebas.bmp"
    ships(a).X = roll(800): ships(a).Y = roll(600)
    ships(a).speed = 4 + roll(10): ships(a).turnspeed = roll(4) + 1
    
    ships(a).energymax = 500
    ships(a).energyregen = 1
    'ships(a).shieldmax = 50
    'ships(a).shieldregen = 10
    ships(a).hpmax = 500
    ships(a).hp = ships(a).hpmax
    addsys ships(a).weapons(1), "Weapon", "Phaser", 3, 4, 0
    addsys ships(a).weapons(1), "WeaponFire", 3, 12, 10
    addsys ships(a).weapons(1), "Bulletinfo", roll(16), 12, 100, 1

Next a

For a = 1 To UBound(stars())
    stars(a).graphnum = 0
    stars(a).X = roll(800)
    stars(a).Y = roll(600)
    'stars(a).size = roll(2)
    If stars(a).size = 3 Then If Not roll(3) = 1 Then stars(a).size = 1
    stars(a).size = 1
Next a

Call CreateUniverse
GoTo 15

ReDim planets(1 To 1000)

planets(1).X = 800: planets(1).Y = 800
planets(1).graphname = "portal1.bmp"
planets(1).planetname = getplanetname


For a = 2 To UBound(planets())
planets(a).X = roll(500000): planets(a).Y = roll(500000)
planets(a).graphname = getgener("starbase3.bmp", "starbase1.bmp", "starbase2.bmp", "planetfire.bmp", "planetfire2.bmp", "planettwilight.bmp")
planets(a).planetname = getplanetname
'Debug.Print "Planet Name: "; planets(a).planetname
Next a
15
spaceon = 1
Bridge.Show
End Sub

Function spacekeyhandler()

Dim state As DIKEYBOARDSTATE
ddkeyboard.GetDeviceStateKeyboard state

If state.Key(DIK_UP) Then calcaccel plrship.facing * 20, 0.2, 15, plrship.xspeed, plrship.yspeed
If state.Key(DIK_H) Then calcaccel plrship.facing * 20, 0.2 * 15, 15 * 15, plrship.xspeed, plrship.yspeed


'If Not GetTickCount() Mod 100 = 0 Then Exit Function
If state.Key(DIK_LEFT) And spacecycle Mod 3 = 0 Then plrship.facing = plrship.facing + 1
If state.Key(DIK_RIGHT) And spacecycle Mod 3 = 0 Then plrship.facing = plrship.facing - 1
If plrship.facing = 0 Then plrship.facing = 18
If plrship.facing = 19 Then plrship.facing = 1

If state.Key(DIK_L) Then tryland
If state.Key(DIK_M) Then
    starmapon = 1
    'If spacecycle Mod 2 = 0 Then
    'For a = 1 To 10: spaceturn: Next a
    'spaceturn
    spaceturn
    Else: starmapon = 0
End If

If state.Key(DIK_C) Then loadcomms

If state.Key(DIK_LCONTROL) Then firenext
If state.Key(DIK_RCONTROL) And firedelay <= 0 Then firenext

If state.Key(DIK_ADD) Then radarzoom = radarzoom * 1.1
If state.Key(DIK_MINUS) Then radarzoom = radarzoom * 0.9

If state.Key(DIK_E) Then
    cmd = ""
    pause = 1
    Shipsys.setsysnum 1: Shipsys.Show: Shipsys.partsupdat
End If

If state.Key(DIK_X) Then
    cmd = ""
    randword shipclassname
    spaceon = 0
    nodraw = 0
    genmapfromship plrship ', , 0.5
    'checkaccess2  (Is now built into genmapfromship)
    addshipsystomap
    Bridge.Hide
End If

'genstarshipmap "Random", 4, 4, 4, 3, 1, 26

' And spacecycle Mod 5 = 0

'Dim durg As systype
'addsys durg, "Weapon", "Buggery Cannon", 3, 6, 4
'addsys durg, "Bulletinfo", 2, 12, 100, 0
'createbullet plrship.x + roll(50), plrship.y + roll(50), durg, 1, plrship.facing * 20, plrship.xspeed / 2, plrship.yspeed / 2

'If firedelay <= 0 Then firenext

'plrship.energy = plrship.energy - 10
'End If

End Function

Sub firenext()
Static lastwep As Byte

lastwep = lastwep + 1

'Fire next weapon in line if possible
For a = lastwep To 12
If plrship.weapondelay(a) < 1 Then fireweapon plrship, a: firedelay = getplrfirerate: Exit Sub
Next a

lastwep = 0
'Otherwise, fire first free weapon
For a = 1 To 12
If plrship.weapondelay(a) < 1 Then fireweapon plrship, a: lastwep = a: firedelay = getplrfirerate: Exit Sub
Next a

End Sub

Sub firefirst(ship1 As shiptype)

'fire first free weapon
For a = 1 To 12
'If ship1.weapons(a).name = "" Then GoTo 5
'If turretonly = 1 Then If getsys(ship1.weapons(a), "WeaponFire", 5) = "" Then GoTo 5
If ship1.weapondelay(a) <= 1 Then fireweapon ship1, a, ship1.owner, ship1.facing: lastwep = a: Exit Sub
5 Next a

End Sub

Sub fireturret(ship1 As shiptype, target As shiptype, Optional angle = 0)

'fire first free weapon
For a = 1 To 12
'If ship1.weapons(a).name = "" Then GoTo 5
If getsys(ship1.weapons(a), "WeaponFire", 5) = "" Then GoTo 5

If ship1.weapondelay(a) <= 0 Then
    If angle = 0 Then MsgBox "Error #1 in fireturret function.": angle = calcforshot(ship1.X, ship1.Y, target.X, target.Y, target.xspeed, target.yspeed, getsys(ship1.weapons(a), "BulletInfo", 2), X, Y, angle, 18)
    fireweapon ship1, a, ship1.owner, angle ': lastwep = a ': Exit Sub
End If

5 Next a

End Sub

Sub spacetimer()

Static lasttick As Long
'Static zark As Double

If Not GetTickCount() > lasttick + 10 Then Exit Sub
'zark = zark - 1: If zark <= 0 Then nodraw = 0
If lasttick + 20 < GetTickCount() Then nodraw = 1 Else nodraw = 0
lasttick = GetTickCount() ': If zark = 1 Then lasttick = lasttick - zark
spacecycle = spacecycle + 1
If spacecycle > 100 Then spacecycle = 1

spacekeyhandler
nodraw = 0
drawships

If spacecycle = 1 Then If countships < 12 And roll(3) = 1 Then addship

End Sub

Function createbullet(ByVal X, ByVal Y, system As systype, ByVal owner As Byte, Optional ByVal angle As Single, Optional ByVal xmod = 0, Optional ByVal ymod = 0)

'Bulletinfo, Graphnumber, Speed, Duration, Maxhits

For a = 1 To UBound(bullets())
    If bullets(a).owner = 0 Then Exit For
Next a

If UBound(bullets()) < a Then ReDim Preserve bullets(1 To a)

bullets(a).special = system
bullets(a).angle = angle
speed = Val(getsys(system, "Bulletinfo", 3)): If speed < 2 Then speed = 2
bullets(a).movex = 0: bullets(a).movey = 0
calcaccel angle, Int(speed), Int(speed), bullets(a).movex, bullets(a).movey
bullets(a).movex = bullets(a).movex + xmod
bullets(a).movey = bullets(a).movey + ymod
bullets(a).X = X: bullets(a).Y = Y
bullets(a).owner = owner
createbullet = a
bullets(a).graphnum = Val(getsys(system, "Bulletinfo", 2))
bullets(a).duration = Val(getsys(system, "Bulletinfo", 4))
bullets(a).Maxhits = Val(getsys(system, "Bulletinfo", 5))
End Function

Function loadbulletgraphs()

'1 = Red Blaster
'2 = Green Blaster
'3 = Blue Blaster
'4 = Purple Blaster
'5 = Red Phaser
'6 = Green Phaser
'7 = Blue Phaser
'8 = Purple Phaser
'9 = Red Laser
'10 = Green Laser
'11 = Blue Laser
'12 = Purple Laser

'13 = Orange Phaser

makesprite bulletgraphs(1), Form1.Picture1, "lasers.bmp", 255, 50, 50, , 18
makesprite bulletgraphs(2), Form1.Picture1, "lasers.bmp", 50, 255, 50, , 18
makesprite bulletgraphs(3), Form1.Picture1, "lasers.bmp", 50, 50, 255, , 18
makesprite bulletgraphs(4), Form1.Picture1, "lasers.bmp", 155, 25, 255, , 18
makesprite bulletgraphs(5), Form1.Picture1, "laser2s.bmp", 255, 0, 0, , 18
makesprite bulletgraphs(6), Form1.Picture1, "laser2s.bmp", 0, 150, 0, , 18
makesprite bulletgraphs(7), Form1.Picture1, "laser2s.bmp", 0, 0, 250, , 18
makesprite bulletgraphs(8), Form1.Picture1, "laser2s.bmp", 155, 25, 255, , 18
makesprite bulletgraphs(9), Form1.Picture1, "laser3s.bmp", 255, 25, 25, , 18
makesprite bulletgraphs(10), Form1.Picture1, "laser3s.bmp", 25, 150, 25, , 18
makesprite bulletgraphs(11), Form1.Picture1, "laser3s.bmp", 25, 25, 250, , 18
makesprite bulletgraphs(12), Form1.Picture1, "laser3s.bmp", 155, 25, 255, , 18

makesprite bulletgraphs(13), Form1.Picture1, "laser2s.bmp", 155, 125, 0, , 18

End Function

Function createfacship(ship As shiptype, facname, shipname, Optional graphic = "ship1s.bmp", Optional speed = 6, Optional turnspeed = 3, Optional armor = 0, Optional shields = 40, Optional shieldregen = 10, Optional energy = 300, Optional energyregen = 0.5, Optional hp = 40, Optional wepname = "Blaster", Optional dice = 2, Optional damage = 6, Optional Bonus = 2, Optional Graphnumber = 1, Optional wepspeed = 12, Optional duration = 100, Optional Maxhits = 1, Optional unknown = 0, Optional firewait = 12, Optional poweruse = 10, Optional red = 0, Optional green = 0, Optional blue = 0)

With ship
    .faction = facname
    .armor = armor
    .energy = energy
    .energymax = energy
    .energyregen = energyregen
    .graphname = graphic
    .hpmax = hp
    .hp = hp
    .maxspeed = speed
    .speed = speed
    .name = shipname
    .shields = shields
    .shieldmax = shields
    .shieldregen = shieldregen
    .turnspeed = turnspeed
    .r = red
    .g = green
    .b = blue
End With

addsys ship.weapons(1), "Weapon", wepname, dice, damage, Bonus
addsys ship.weapons(1), "WeaponFire", unknown, firewait, poweruse
addsys ship.weapons(1), "Bulletinfo", Graphnumber, wepspeed, duration, Maxhits

'"Weapon", Name, Dice, Damage, Bonus
'"WeaponFire", ??? (Ammo), fire wait, power use,
'"Bulletinfo", Graphnumber, Speed, Duration, Maxhits

End Function

Function addweptoship(ship As shiptype, wepnum, wepname, dice, damage, Bonus, Graphnumber, wepspeed, duration, Maxhits, ammo, firewait, poweruse, Optional turret = "")

addsys ship.weapons(wepnum), "Weapon", wepname, dice, damage, Bonus
addsys ship.weapons(wepnum), "WeaponFire", ammo, firewait, poweruse, turret
addsys ship.weapons(wepnum), "Bulletinfo", Graphnumber, wepspeed, duration, Maxhits

End Function

Function checkforhits()

'Check for bullet impacts
For a = 1 To UBound(bullets())
    If bullets(a).owner = 0 Then GoTo 5
    'Check for collisions with player
    If bullets(a).owner < 2 Then GoTo 9
    If plrship.instomach > 0 Then GoTo 9
    If diff(plrship.X, bullets(a).X) * 2 > shipgraphs(plrship.graphnum).CellWidth Then GoTo 9
    If diff(plrship.Y, bullets(a).Y) * 2 > shipgraphs(plrship.graphnum).CellHeight Then GoTo 9
        
        dmg = shipshot(bullets(a).special, plrship, shields)
        If dmg = 0 Then createjunk bullets(a).X, bullets(a).Y, "explosion3.bmp", , , 0.5, 6 Else createjunk bullets(a).X, bullets(a).Y, "explosion2.bmp", , , 0.5, 6
        If dmg > 0 Then addtext dmg, bullets(a).X, bullets(a).Y, 255, 170, 0, , "plrdamage", 250, 0, 0
        If shields > 0 Then addtext shields, bullets(a).X, bullets(a).Y - 20, 0, 250, 40, , "plrshields", 0, 0, 0
        'addtext shields, bullets(a).x, bullets(a).y - 20, 0, 250, 30, , "shields" & b, 0, 0, 100
        If bullets(a).Maxhits > 0 Then bullets(a).Maxhits = bullets(a).Maxhits - 1
        If bullets(a).Maxhits = 0 Then bullets(a).duration = 0: bullets(a).owner = 0
    
9
    'Check for collisions with ships
    For b = 1 To UBound(ships())
        If ships(b).owner = -1 Then GoTo 7
        If bullets(a).owner = ships(b).owner Or bullets(a).owner = 0 Then GoTo 7
        If diff(ships(b).X, bullets(a).X) * 2 > shipgraphs(ships(b).graphnum).CellWidth Then GoTo 7
        If diff(ships(b).Y, bullets(a).Y) * 2 > shipgraphs(ships(b).graphnum).CellHeight Then GoTo 7
        
        dmg = shipshot(bullets(a).special, ships(b), shields)
        If dmg = 0 Then createjunk bullets(a).X, bullets(a).Y, "explosion3.bmp", , , 0.5, 6 Else createjunk bullets(a).X, bullets(a).Y, "explosion2.bmp", , , 0.5, 6
        
        If dmg > 0 Then addtext dmg, bullets(a).X, bullets(a).Y, 250, 170, 0, , "HP" & b, 250, 0, 0
        If shields > 0 Then addtext shields, bullets(a).X, bullets(a).Y - 20, 0, 250, 40, , "shields" & b, 0, 0, 0
        
        ships(b).swoopcount = ships(b).swoopcount / 2
        
        If bullets(a).Maxhits > 0 Then bullets(a).Maxhits = bullets(a).Maxhits - 1
        If bullets(a).Maxhits = 0 Then bullets(a).duration = 0: bullets(a).owner = 0
7     Next b
    

    
5 Next a

'Check for ships colliding with each other

For a = 1 To UBound(ships())
    If diff(plrship.X, ships(a).X) < 60 And diff(plrship.Y, ships(a).Y) < 60 Then ships(a).swoopcount = 120
    For b = 1 To UBound(ships())
        If diff(ships(a).X, ships(b).X) > shipgraphs(ships(b).graphnum).CellWidth Then GoTo 12
        If diff(ships(a).Y, ships(b).Y) > shipgraphs(ships(b).graphnum).CellHeight Then GoTo 12
        'ships(b).xspeed = -ships(a).xspeed: ships(b).yspeed = -ships(a).yspeed: zark = 1: Exit For
        If diff(ships(a).X, ships(b).X) > diff(ships(a).Y, ships(b).Y) Then
        'dist = shipgraphs(ships(b).graphnum).CellWidth / 2
        dist = 3
        If ships(b).X > ships(a).X Then ships(b).X = ships(b).X + dist: ships(a).X = ships(a).X - dist Else ships(b).X = ships(b).X - dist: ships(a).X = ships(a).X + dist
        Else:
        If ships(b).Y > ships(a).Y Then ships(b).Y = ships(b).Y + dist: ships(a).Y = ships(a).Y - dist Else ships(b).Y = ships(b).Y - dist: ships(a).Y = ships(a).Y + dist
        End If
        'Bounce the ships off of each other
        'calcangle ships(a).x, ships(a).y, ships(b).x, ships(b).y, ships(b).xspeed, ships(b).yspeed, 10, cell, 18
        'calcangle ships(b).x, ships(b).y, ships(a).x, ships(a).y, ships(a).xspeed, ships(a).yspeed, 10, cell, 18
        'Exit For
12
Next b
'If zark = 1 Then Exit For
Next a

End Function

Function shipAI()
Dim durg As systype

'Weapon:  "Weapon", Name, Dice, Damage, Bonus
addsys durg, "Weapon", "Buggery Cannon", 3, 6, 4
addsys durg, "Bulletinfo", 2, 12, 100, 0


'If Not spacecycle Mod 10 = 0 Then Exit Function

For a = 1 To UBound(ships())
        If ships(a).owner = -1 Then GoTo 5
        shipmaint ships(a)
        If spacecycle = 50 Then If calcdist(ships(a).X, ships(a).Y, plrship.X, plrship.Y) > 10000 Or (ships(a).X < 1 Or ships(a).Y < 1) Then killship ships(a), 0: GoTo 5
        'Turn
        If ships(a).swoopcount > 0 Then ships(a).swoopcount = ships(a).swoopcount - 1
        calcforshot ships(a).X, ships(a).Y, plrship.X, plrship.Y, plrship.xspeed, plrship.yspeed, 12, xmove, ymove, angle
        If spacecycle Mod ships(a).turnspeed = 0 And ships(a).swoopcount <= 1 Then
        turntowards angle, ships(a).facing
        End If
        'If spacecycle Mod 10 = 0 Then If diff(ships(a).facing, angle) < 2 Then createbullet ships(a).X, ships(a).Y, durg, ships(a).owner, ships(a).facing * 20, ships(a).xspeed / 2, ships(a).yspeed / 2
        If diff(ships(a).facing, angle) < 2 Then firefirst ships(a) ' ships(a).X, ships(a).Y, durg, ships(a).owner, ships(a).facing * 20, ships(a).xspeed / 2, ships(a).yspeed / 2
        fireturret ships(a), plrship, angle
        
        If ships(a).ismonster = 0 Then
            calcaccel ships(a).facing * 20, ships(a).speed / 50, ships(a).speed, ships(a).xspeed, ships(a).yspeed
        End If
        
        If ships(a).ismonster > 0 Then
            If ships(a).eating = 0 Then
                If spacecycle = 99 Then If roll(6) = 1 Then ships(a).eating = -1: GoTo 13
                If calcdist(ships(a).X, ships(a).Y, plrship.X, plrship.Y) < 300 Then
                calcaccel ships(a).facing * 20, -(ships(a).speed / 50), ships(a).speed, ships(a).xspeed, ships(a).yspeed
                Else: calcaccel ships(a).facing * 20, ships(a).speed / 50, ships(a).speed, ships(a).xspeed, ships(a).yspeed
                End If
                
13         End If
        
        If ships(a).eating = -1 Then
            If spacecycle = 98 Then If roll(2) = 1 Then ships(a).eating = 0: GoTo 5
            calcaccel ships(a).facing * 20, ships(a).speed / 50, ships(a).speed, ships(a).xspeed, ships(a).yspeed
            If calcdist(ships(a).X, ships(a).Y, plrship.X, plrship.Y) < 50 And plrship.instomach = 0 Then ships(a).eating = -2: plrship.instomach = a
        End If
        
        If Not ships(a).eating = -1 And Not ships(a).eating = 0 Then
        
            turntowards roll(360), ships(a).facing
            If spacecycle Mod 10 = 0 Then digestship ships(a), plrship
            
        End If
        
        End If
        
5 Next a

End Function

Function shipmaint(ship1 As shiptype)
'Per-cycle stuff

ship1.energy = ship1.energy + ship1.energyregen
If ship1.energy > ship1.energymax Then ship1.energy = ship1.energymax

For a = 1 To 12
If ship1.weapondelay(a) > 0 Then ship1.weapondelay(a) = ship1.weapondelay(a) - 1
If ship1.weapons(a).eff(1, 1) = "" Then ship1.weapondelay(a) = 1
Next a

If ship1.shieldmax = 0 Then GoTo 5
If ship1.shieldmax >= ship1.shieldregen Then If spacecycle Mod (20 + ship1.shieldmax * 2 / ship1.shieldregen) < 1 Then regainshields ship1
If ship1.shieldregen > ship1.shieldmax Then If spacecycle Mod greater(1, (15 - ship1.shieldregen / ship1.shieldmax)) < 1 Then regainshields ship1
5

End Function

Function drawplanets()

For a = 1 To UBound(planets())
    If planets(a).X = 0 Then GoTo 5
    If diff(planets(a).X, plrship.X) > 1000 Or diff(planets(a).Y, plrship.Y) > 1000 Then GoTo 5
    If planets(a).graphnum = 0 Then planets(a).graphnum = loadrandomgraph(planets(a).graphname)
    drawthingy rgraphs(planets(a).graphnum).g, planets(a).X, planets(a).Y, 1
    'rgraphs(planets(a).graphnum).g.TransparentDraw picbuffer, planets(a).x - rgraphs(gn).g.CellWidth / 2, planets(a).y - rgraphs(gn).g.CellHeight / 2, 1
    'rgraphs(planets(a).graphnum).g.TransparentDraw picbuffer, 100, 100, 1
5 Next a

End Function

Function drawthingy(cS As cSpriteBitmaps, ByVal X, ByVal Y, Optional ByVal cell = 1, Optional ByVal refreshpersp = 0)

Static x3
Static y3

If cS Is Nothing Then Exit Function

x4 = X

X = X - plrship.X - cS.CellWidth / 2 + 400
Y = Y - plrship.Y - cS.CellHeight / 2 + 300

If refreshpersp = 0 Then GoTo 5 'No perspective changes

calcaccel plrship.facing * 20, 250, 250, x2, y2

If Not x4 = plrship.X Then GoTo 5

If x2 > x3 Then x3 = x3 + diff(x2, x3) / 30
If y2 > y3 Then y3 = y3 + diff(y2, y3) / 30
If x2 < x3 Then x3 = x3 - (x3 - x2) / 30
If y2 < y3 Then y3 = y3 - (y3 - y2) / 30

5

X = X - x3 'plrship.xspeed * 5
Y = Y - y3 'plrship.yspeed * 5

sxoff = -x3
syoff = -y3

If X < -cS.CellWidth Then Exit Function
If Y < -cS.CellHeight Then Exit Function

If X > cS.CellWidth + 800 Then Exit Function
If Y > cS.CellHeight + 600 Then Exit Function

If refreshpersp = 1 Then Exit Function

cS.TransparentDraw picBuffer, X, Y, cell

End Function

Sub regainshields(ship1 As shiptype)

Dim ratio As Single

ratio = ship1.shieldregen / ship1.shieldmax

ship1.shields = ship1.shields + (ship1.shieldmax * 0.05)

'ship1.shields = ship1.shields + 1 + ship1.shieldregen * ratio

If ship1.shields > ship1.shieldmax Then ship1.shields = ship1.shieldmax

End Sub

Sub drawradar(X, Y, Optional zoom = 20)

Dim r1 As RECT
Dim sdesc As DDSURFACEDESC2



'Exit Sub

r1.Top = 400 + X
r1.Left = 400 + Y
r1.Right = 600 + X
r1.Bottom = 600 + Y

picBuffer.GetSurfaceDesc sdesc

'picbuffer.Lock r1, sdesc, DDLOCK_DONOTWAIT, 0

picBuffer.SetForeColor RGB(0, 150, 0)

picBuffer.DrawLine X + 400, Y + 400, X + 600, Y + 400
picBuffer.DrawLine X + 400, Y + 400, X + 400, Y + 600
picBuffer.DrawLine X + 400, Y + 600, X + 600, Y + 600
picBuffer.DrawLine X + 600, Y + 400, X + 600, Y + 600

picBuffer.setDrawStyle vbDash

'GoTo 5
For a = 1 To 5
        picBuffer.DrawLine X + 400 + a * 40, Y + 400, X + 400 + a * 40, Y + 600
        picBuffer.DrawLine X + 400, Y + 400 + a * 40, X + 600, Y + 400 + a * 40
Next a
5
picBuffer.setDrawStyle vbSolid
picBuffer.SetForeColor RGB(250, 150, 0)
For a = 1 To UBound(ships())
    If ships(a).owner < 0 Then GoTo 6
    If diff(plrship.X, ships(a).X) / zoom > 99 Then GoTo 6
    If diff(plrship.Y, ships(a).Y) / zoom > 99 Then GoTo 6
    picBuffer.DrawCircle (ships(a).X - plrship.X) / zoom + 500 + X, (ships(a).Y - plrship.Y) / zoom + 500 + Y, 2
    'picbuffer.SetLockedPixel (ships(a).x - plrship.x + 800) / zoom + 460 + x, (ships(a).y - plrship.y + 600) / zoom + 460 + x, RGB(0, 50, 255)
6 Next a

picBuffer.SetForeColor RGB(0, 140, 50)
For a = 1 To UBound(planets())
    'picbuffer.DrawLine planets(a).X / (500000 / 400) + 400, planets(a).Y / (500000 / 400) + 400, planets(a).X / (500000 / 400) + 402, planets(a).Y / (500000 / 400) + 402
    If planets(a).X = 0 Then GoTo 8
    If diff(plrship.X, planets(a).X) / zoom > 99 Then GoTo 8
    If diff(plrship.Y, planets(a).Y) / zoom > 99 Then GoTo 8
    picBuffer.DrawCircle (planets(a).X - plrship.X) / zoom + 500 + X, (planets(a).Y - plrship.Y) / zoom + 500 + Y, 6 * (100 / zoom)
8 Next a

picBuffer.SetForeColor RGB(0, 150, 255)
picBuffer.DrawLine X + 495, Y + 500, X + 505, Y + 500
picBuffer.DrawLine X + 500, Y + 495, X + 500, Y + 505

picBuffer.SetForeColor RGB(0, 250, 65)
picBuffer.DrawLine X + 400, 400 + spacecycle * 2, 600, X + 400 + spacecycle * 2
picBuffer.SetForeColor RGB(0, 200, 0)
picBuffer.DrawLine X + 400, 400 + spacecycle * 2 - 1, 600, X + 400 + spacecycle * 2 - 1
picBuffer.SetForeColor RGB(0, 100, 0)
picBuffer.setDrawStyle vbDot
picBuffer.DrawLine X + 400, 400 + spacecycle * 2, 600, X + 400 + spacecycle * 2
picBuffer.DrawLine X + 400, 400 + spacecycle * 2 - 2, 600, X + 400 + spacecycle * 2 - 2

picBuffer.setDrawStyle vbSolid
'picbuffer.SetLockedPixel (400) / zoom + 500 + x, (300) / zoom + 500 + y, RGB(255, 0, 0)

'picbuffer.Unlock r1

End Sub

Sub drawstars()

For a = 1 To UBound(stars())
    If stars(a).graphnum = 0 Then stars(a).size = roll(3): stars(a).graphnum = loadrandomgraph("star" & stars(a).size & ".bmp", roll(5) * 50, roll(5) * 50, roll(5) * 50)
    If stars(a).X > 800 Then stars(a).X = stars(a).X - 800: stars(a).Y = roll(600)
    If stars(a).Y > 600 Then stars(a).Y = stars(a).Y - 600: stars(a).X = roll(800)
    If stars(a).X < 1 Then stars(a).X = stars(a).X + 800: stars(a).Y = roll(600)
    If stars(a).Y < 1 Then stars(a).Y = stars(a).Y + 600: stars(a).X = roll(800)
    stars(a).X = stars(a).X - (plrship.xspeed / 12) * (stars(a).size * stars(a).size)
    stars(a).Y = stars(a).Y - (plrship.yspeed / 12) * (stars(a).size * stars(a).size)

    'drawthingy rgraphs(stars(a).graphnum).g, stars(a).X + plrship.X - 400, stars(a).Y + plrship.Y - 300

    rgraphs(stars(a).graphnum).g.TransparentDraw picBuffer, stars(a).X, stars(a).Y, 1
    'drawthingy rgraphs(stars(a).graphnum), stars(a).x + plrship.x, stars(a).y + plrship.y
Next a

End Sub

Sub drawjunk()
'Gold:Amt
'Cargo:Name:Amt
'Energy:Amt
'Shields:Amt
'HP:Amt
'Ammo?:???
'Animtype=1 sets for pickups


For a = 1 To UBound(rjunk())
    If rjunk(a).graphnum = 0 Then GoTo 5
    drawthingy rgraphs(rjunk(a).graphnum).g, rjunk(a).X, rjunk(a).Y, Int(rjunk(a).frame)
    If rjunk(a).animtype = 0 Then rjunk(a).frame = rjunk(a).frame + rjunk(a).animspeed: If rjunk(a).frame > rgraphs(rjunk(a).graphnum).g.cells Then rjunk(a).graphnum = 0 'Type 0=Show anim once, then dissappear
    If rjunk(a).animtype = 1 Then
        If spacecycle = 50 Then rjunk(a).age = rjunk(a).age + 1: If rjunk(a).age > 12 Then rjunk(a).graphnum = 0: GoTo 5
        rjunk(a).frame = rjunk(a).frame + greater(rjunk(a).animspeed, 0.1)
        If rjunk(a).frame > rgraphs(rjunk(a).graphnum).g.cells Then rjunk(a).frame = 1 'Type 1=Loop + collision check
        If diff(rjunk(a).X, plrship.X) < 50 And diff(rjunk(a).Y, plrship.Y) < 50 Then
        val1 = getfromstring(rjunk(a).special, 1)
        val2 = getfromstring(rjunk(a).special, 2)
        val3 = getfromstring(rjunk(a).special, 2)
        If val1 = "Gold" Then plr.gp = plr.gp + val2: addtext val2 & " GP", rjunk(a).X, rjunk(a).Y, 240, 240, 10, 60
        If val1 = "Cargo" Then addcargo val2, Val(val3): addtext val2 & " " & val3, rjunk(a).X, rjunk(a).Y, 250, 250, 250
        If val1 = "Energy" Then plrship.energy = plrship.energy + val2: addtext "+" & val2 & " Energy", rjunk(a).X, rjunk(a).Y, 250, 250, 1
        If val1 = "HP" Then plrship.hp = plrship.hp + val2: addtext "+" & val2 & " HP", rjunk(a).X, rjunk(a).Y, 250, 250, 1, 60: If plrship.hp > plrship.hpmax Then plrship.hp = plrship.hpmax
        If val1 = "Shields" Then plrship.shields = plrship.shields + val2: addtext "+" & val2 & " Shields", rjunk(a).X, rjunk(a).Y, 1, 250, 1, 60
        rjunk(a).graphnum = 0
        End If
    End If

5 Next a

End Sub

Sub digestship(ship1 As shiptype, eatenship As shiptype)

If ship1.hp < 1 Then eatenship.instomach = 0: Exit Sub

damageship eatenship, greater(roll(eatenship.hpmax / 100), 1)

If eatenship.hp < 100 And eatenship.isplr = 1 Then

genmapfromship ship1, 4, 7

zark = roll(4) + 2
ReDim montype(1 To zark)

For a = 1 To zark
    montype(a) = randenzymetype(greater(plr.level, 1))
Next a

loadallmongraphs
randommonsters "ALL", Form1.Picture1.Height

spaceon = 0: nodraw = 0: Bridge.Hide: Exit Sub

End If

    redl = roll(150) + 100
    greenl = redl - roll(redl / 2)
    addtext getgener("*gurgle*", "*blorp*", "*groan*"), ship1.X + roll(30) - 15, ship1.Y + roll(80) - 40, redl, greenl, 3


End Sub

Sub createjunk(X, Y, filen, Optional animtype = 0, Optional special = "", Optional animspeed = 1, Optional frames = 17)
For a = 1 To UBound(rjunk())
    If rjunk(a).graphnum = 0 Then
        With rjunk(a)
        .X = X
        .Y = Y
        .frame = 1
        .special = special
        .animtype = animtype
        .graphnum = loadrandomgraph(filen, , , , frames)
        .animspeed = animspeed
        .age = 0
        End With
        Exit Sub
    End If
Next a

End Sub

Function fireweapon(ship1 As shiptype, ByVal slot As Byte, Optional ByVal owner = 1, Optional ByVal angle = 0) As Boolean
'Returns true if weapon can fire and automatically subtracts energy and sets fire wait

If ship1.weapondelay(slot) > 0 Then Exit Function
'If ship1.weapons(slot).name = "" Then Exit Function
ship1.weapondelay(slot) = getsys(ship1.weapons(slot), "WeaponFire", 3)
If ship1.energy < getsys(ship1.weapons(slot), "WeaponFire", 4) Then Exit Function
If angle = 0 Then angle = ship1.facing
createbullet ship1.X + roll(40) - 20, ship1.Y + roll(40) - 20, ship1.weapons(slot), owner, angle * 20, ship1.xspeed / 2, ship1.yspeed / 2
ship1.energy = ship1.energy - getsys(ship1.weapons(slot), "WeaponFire", 4)

End Function

Function addnews(newstype As String, Optional string1 As String = "", Optional string2 As String = "", Optional string3 As String = "", Optional string4 As String = "", Optional string5 As String = "")

'move all news back one step
For a = 10 To 2 Step -1
    news(a) = news(a - 1)
Next a

news(1).newstype = newstype
news(1).strings(1) = string1
news(1).strings(2) = string2
news(1).strings(3) = string3
news(1).strings(4) = string4
news(1).strings(5) = string5

End Function

Function createnews(Optional ByVal suf As String = "") As String

Dim newzes(1 To 20) As String 'News sentences to choose from
Dim newsstr As String

'Randomize Timer

If suf = "" Then wnews = roll(5) Else wnews = roll(10) 'Only the top five will be used as news--later ones tend to be rumors
If news(wnews).newstype = "" Then wnews = 1

Open "spaceconvs.txt" For Input As #1

newson = 0

newstype = news(wnews).newstype
string1 = news(wnews).strings(1)
string2 = news(wnews).strings(2)
string3 = news(wnews).strings(3)
string4 = news(wnews).strings(4)
string5 = news(wnews).strings(5)

If suf = "Rumors" Then
aroll = roll(3)

Select Case aroll
Case 1: newstype = "#RANDOM"
Case 2: newstype = "#MONSTERS": string1 = "Tau Proxima IV": string2 = "Doom Karis"
End Select

End If

Do While Not EOF(1)
    Input #1, gstr
    
    'Switch news adding on or off
    If gstr = "#NEWS" Then Input #1, gstr2
    If gstr = "#NEWS" And gstr2 = newstype & suf Then
    'newson = 1 Else newson = 0
        newsnum = 1
        Do While (1)
        Input #1, gstr
        If gstr = "#NEWSSTR" Then Input #1, newzes(newsnum): newsnum = newsnum + 1 Else Exit Do
        Loop
3         aroll = roll(newsnum)
        If newzes(aroll) = "" Then GoTo 3
        newsstr = newsstr & newzes(aroll) & " "
    End If
    
Loop

Close #1

If newstype = "#WAR" Then
    swaptxt newsstr, "$FACTION1Z", string3
    swaptxt newsstr, "$FACTION2Z", string4
    swaptxt newsstr, "$FACTION1", string1
    swaptxt newsstr, "$FACTION2", string2
    swaptxt newsstr, "$PLANETTAKEN", string5
End If

If newstype = "#MONSTEREAT" Then
    swaptxt newsstr, "$FACTION", string1
    swaptxt newsstr, "$STARSHIPNAME", string2
    swaptxt newsstr, "$STARSHIPTYPE", string3
    swaptxt newsstr, "$SPACEMONSTER", string4
    'swaptxt newsstr, "$PLANETTAKEN", string5
End If

If newstype = "#MONSTERS" Then
    swaptxt newsstr, "$PLANETNAME", string1
    swaptxt newsstr, "$PLANETMONSTER", string2
'    swaptxt newsstr, "", string3
'    swaptxt newsstr, "", string4
'    swaptxt newsstr, "", string5
End If

'If newstype = "#WAR" Then
'    swaptxt newsstr, "", string1
'    swaptxt newsstr, "", string2
'    swaptxt newsstr, "", string3
'    swaptxt newsstr, "", string4
'    swaptxt newsstr, "", string5
'End If

'Standard Replacements
swaptxt newsstr, "$POOP", getpoop("")
swaptxt newsstr, "$EATEN", geteaten
swaptxt newsstr, "$EATING", geteating
swaptxt newsstr, "$EAT", geteat
swaptxt newsstr, "$DIGESTED", getdigest("ed")
swaptxt newsstr, "$DIGESTING", getdigest("ing")
swaptxt newsstr, "$DIGEST", getdigest("")
swaptxt newsstr, "$RANDOMNAME", getname
swaptxt newsstr, "$RANDOMNAME2", getname
swaptxt newsstr, "$RANK", getgener("Ensign", "Lieutenant", "Commander", "Captain", "Admiral", "Commodore", "Lady", "Chief")
swaptxt newsstr, "$BELLIES", getbelly("s")
swaptxt newsstr, "$BELLY", getbelly("")
swaptxt newsstr, "$MYNAME", newsname

createnews = newsstr

End Function



Function getfromset(setname) As String

Dim sets(25) As String

Open "c:\vb\VRPG\spaceconvs.txt" For Input As #1

Do While Not EOF(1)
    Input #1, gstr
    
    If gstr = "#SET" Then Input #1, gstr2
    If gstr = "#SET" And gstr2 = setname Then
        Input #1, gstr
        If gstr = "#SETI" Then Input #1, sets(lastset): lastset = lastset + 1 Else Exit Do
    End If
Loop

getfromset = sets(roll(lastset - 1))

End Function

Function landplanet(planet As planetT)

nodraw = 0
'ChDir App.Path & "\" & plr.name
'If Dir(App.Path & "\" & plr.name, vbDirectory) = "" Then MkDir App.Path & "\" & plr.name

'filen = App.Path & "\" & plr.name & "\" & "Universe.dat"
filen = "curgame.dat"

5 If findfile(planet.planetname & ".dat", filen) = True Then 'Prefer binary files
    
    loadbindata planet.planetname & ".dat", filen

Else:

    If findfile(planet.planetname & ".txt", filen) = True Then 'If no binary is found, load a text file
        
        loaddata planet.planetname & ".txt", filen
        
    Else: 'If no text OR binary is found, generate random planet
    
    If planet.data.name = "Starbase" Then
    loadstation 22, Int(getgener("24", "26", "27", "28", "29", "30"))
    savebindata planet.planetname & ".dat"
    addfile planet.planetname & ".dat", filen
    GoTo 17
    End If
        maplev = Int((diff(250000, plrship.X) + diff(250000, plrship.Y)) / 10000)
        filen2 = genmap2(, , , , , planet.planetname, Int((plr.level + maplev) / 2))
        addfile filen2, filen
        Kill filen2
        GoTo 5
        'planet.planetname, plr.level, 100, 100
        
    End If

End If

17 showform10 scanmap(60), -1, planet.graphname

Bridge.Hide: spaceon = 0: plr.X = mapx / 2: plr.Y = mapy / 2

End Function

Sub tryland()

For a = 1 To UBound(planets())
    If diff(planets(a).X, plrship.X) < 300 And diff(planets(a).Y, plrship.Y) < 300 Then landplanet planets(a): Exit Sub
Next a

End Sub

Sub partswap(ByRef part1 As systype, ByRef part2 As systype)
Dim part3 As systype

part3 = part1
part1 = part2
part2 = part3

End Sub

Sub setplrship()
'Weapon:  "Weapon", Name, Dice, Damage, Bonus
'Shields:  "Shield"/"ShieldBon", Regen, Max, Reflection, ,
'"Extraneous", Weight, , ,
'"Engine", Thruster, Hyper, ,

'Adding "Global", as in "GlobalShieldBon" means it will affect ALL systems of that type.
'Getplrsys automatically looks for this.

'Set Shields
plrship.shieldmax = getplrsys("Shield", 3, "ShieldBon", 3)
plrship.shieldregen = getplrsys("Shield", 2, "ShieldBon", 2)
plrship.shields = plrship.shieldmax

plrship.energyregen = (10 + (getplrsys("Power", 2, "PowerBon", 2) / 5)) / 30
plrship.energymax = getplrsys("Power", 4, "PowerBon", 4)

'Clear Weapons
Dim doink As systype
For a = 1 To UBound(plrship.weapons())
    plrship.weapons(a) = doink
Next a

'Set Weapons
lastwep = 1
For b = 1 To UBound(plrshipdat.sections())
    For a = 1 To 4
    'If b = 5 And a = 4 Then Stop
    If getsys(plrshipdat.sections(b).partsobj(a), "Weapon", 1) = "Weapon" And lastwep < 13 Then
    'ReDim Preserve plrship.weapons(1 To lastwep)
        If a = 4 Then
        wdice = getsys(plrshipdat.sections(b).partsobj(a), "Weapon", 3)
        wdamage = getsys(plrshipdat.sections(b).partsobj(a), "Weapon", 4)
        wbonus = getsys(plrshipdat.sections(b).partsobj(a), "Weapon", 5)
        wname = getsys(plrshipdat.sections(b).partsobj(a), "Weapon", 2)
        
        wenergy = getsys(plrshipdat.sections(b).partsobj(a), "WeaponFire", 4)
        wfire = getsys(plrshipdat.sections(b).partsobj(a), "WeaponFire", 3) 'Firing Delay
        wammo = getsys(plrshipdat.sections(b).partsobj(a), "WeaponFire", 2)
        
        wzarg = 0 'wdice * (wdamage * 0.4) + wbonus
        For c = 1 To 3
            wzarg = wzarg * getsys(plrshipdat.sections(b).partsobj(c), "WeaponMult", 2)
            wzarg = wzarg + Val(getsys(plrshipdat.sections(b).partsobj(c), "WeaponMult", 3))
        Next c
        plrship.weapons(lastwep) = plrshipdat.sections(b).partsobj(a)
        addsys plrship.weapons(lastwep), "Weapon", wname, wdice, wdamage, wbonus + wzarg, 1
        addsys plrship.weapons(lastwep), "WeaponFire", wammo, wfire, wenergy, , 1
        lastwep = lastwep + 1
        
        Else:
        plrship.weapons(lastwep) = plrshipdat.sections(b).partsobj(a)
        lastwep = lastwep + 1
        End If
    
    End If
    Next a
    'If Not multeffname = "" Then
    '    cursys = cursys + (getsys(plrshipdat.sections(b).partsobj(a), effname, slot) * getsys(plrshipdat.sections(b).partsobj(a), multeffname, multslot, 1)) * getglobal("Global" & multeffname, slot)
    'Else:
    '    cursys = cursys + getsys(plrshipdat.sections(b).partsobj(a), effname, slot)
    'End If
Next b

weight = (plrshipdat.weight + getplrsys("Extraneous", 2)) * getplrsys("ExtraneousBon", 2)
plrship.speed = getplrsys("Engine", 2, "EngineBon", 2)

End Sub

Function addcargo(ByVal cargoname, ByVal amt)

Static dink As Byte
If dink = 0 Then ReDim plrshipdat.cargoname(1 To 1): ReDim plrshipdat.cargoamt(1 To 1): dink = 1

For a = 1 To UBound(plrshipdat.cargoname())
    If plrshipdat.cargoname(a) = cargoname Then plrshipdat.cargoamt(a) = plrshipdat.cargoamt(a) + amt: Exit Function
    If plrshipdat.cargoname(a) = "" Then plrshipdat.cargoname(a) = cargoname: plrshipdat.cargoamt(a) = amt: Exit Function
Next a

ReDim Preserve plrshipdat.cargoname(1 To UBound(plrshipdat.cargoname()) + 1): ReDim Preserve plrshipdat.cargoamt(1 To UBound(plrshipdat.cargoname()))
plrshipdat.cargoname(a) = cargoname: plrshipdat.cargoamt(a) = amt

End Function

Function starbasecargo(ByVal cargoname, ByVal amt)

Static dink As Byte
If dink = 0 Then ReDim starbase.cargoname(1 To 1): ReDim starbase.cargoamt(1 To 1): dink = 1

For a = 1 To UBound(starbase.cargoname())
    If starbase.cargoname(a) = cargoname Then starbase.cargoamt(a) = starbase.cargoamt(a) + amt: Exit Function
    If starbase.cargoname(a) = "" Then starbase.cargoname(a) = cargoname: starbase.cargoamt(a) = amt: Exit Function
Next a

ReDim Preserve starbase.cargoname(1 To UBound(starbase.cargoname()) + 1): ReDim Preserve starbase.cargoamt(1 To UBound(starbase.cargoname()))
starbase.cargoname(a) = cargoname: starbase.cargoamt(a) = amt

End Function

Function scanmap(Optional scanlev = 30) As String

txt = mapjunk.terstr & vbCrLf
txt = txt & "--- [[ PLANETARY SCAN ]] ---" & vbCrLf & vbCrLf
txt = txt & "Temperature: " & mapjunk.planetjunk.temperature & " -- "
txt = txt & "Moisture: " & mapjunk.planetjunk.moisture & " -- "
txt = txt & "Rockyness: " & mapjunk.planetjunk.rockyness & " -- "
txt = txt & "Vegetation: " & mapjunk.planetjunk.vegetation & vbCrLf & vbCrLf

scurv = Int((100 - scanlev) / 10)
lowlev = greater(lowestlevel - roll(scurv), 1)
highlev = lowestlevel + roll(scurv)

txt = txt & "Threat Level: " & lowlev & "-" & highlev & vbCrLf

txt = txt & "Detected Ores: "

'lastzark = 1
Dim zark(1 To 3)
Dim zarkname(1 To 3)
For a = 0 To scanlev * (scanlev / 10) + 50
    X = roll(mapx)
    Y = roll(mapy)
    
    If map(X, Y).object > 0 Then
        If Not geteff(objtypes(objs(map(X, Y).object).type), "Cargo", 2) = "" Then
            For b = 1 To 3
                If zarkname(b) = geteff(objtypes(objs(map(X, Y).object).type), "Cargo", 2) Then zark(b) = zark(b) + objs(map(X, Y).object).string: Exit For
                
                If zark(b) = 0 Then
                    zarkname(b) = geteff(objtypes(objs(map(X, Y).object).type), "Cargo", 2)
                    zark(b) = zark(b) + objs(map(X, Y).object).string
                    Exit For
                End If
            Next b
        End If
    End If
    If map(X, Y).monster > 0 Then moncount = moncount + 1
Next a

For a = 1 To 3
    If zark(a) > 0 Then
        
        'Multiply ore reading according to what percentage of surface was read (So inferior scanners don't always read low)
        
        zark(a) = zark(a) * (1 / lesser((mapx * mapy) / (scanlev * (scanlev / 10) + 50), 1))
        size = "Trace"
        If zark(a) > 150 Then size = "Low"
        If zark(a) > 400 Then size = "Medium"
        If zark(a) > 700 Then size = "High"
        If zark(a) > 1600 Then size = "Ultra-High"
        
        txt = txt & zarkname(a) & ": " & size & "  "
    End If
Next a

monrat = moncount / scanlev

txt = txt & vbCrLf & "Estimated Monster Saturation:" & monrat * 200 & "%"

scanmap = txt

makeminimap

End Function

Function getplrfirerate()

Dim minfire As Integer
Dim firerate As Integer

For a = 1 To 12
    zorg = getsys(plrship.weapons(a), "WeaponFire", 3)
    If zorg > 0 And minfire = 0 Then minfire = zorg
    If zorg > 0 And zorg < minfire Then minfire = Val(zorg)
    If zorg > 0 Then numweps = numweps + 1: firerate = firerate + zorg: If zorg < minfire Then zorg = minfire
Next a

'getplrfirerate = firerate / numweps
getplrfirerate = minfire / numweps

End Function

Sub drawwepstats(X, Y)

Dim wepname As Variant
For a = 1 To 12
    wepname = getsys(plrship.weapons(a), "Weapon", 2, "")
    If plrship.weapons(a).name = "" Then wepname = ""
    If Not wepname = "" Then
        If plrship.weapondelay(a) > 0 Then r = 255: g = 0 Else r = 0: g = 255
        drawtext wepname, X, Y + (a * 20), r, g, b
    End If
Next a


End Sub

Sub loadstation(Optional tilet = 20, Optional ovrtile = 28, Optional stationtype = 0)

If stationtype = 0 Then stationtype = roll(1)


'Simple Plus-shaped Station -- Four Rooms
If stationtype = 1 Then
    mapx = 80: mapy = 80
    ReDim map(mapx, mapy)
    loadseg "20cooridor1.seg", 20, 20, 0, tilet, ovrtile
    loadseg "10shop" & roll(6) + 2 & ".seg", 10, 25, 3, tilet, ovrtile, getseller
    loadseg "10shop" & roll(6) + 2 & ".seg", 25, 10, 0, tilet, ovrtile, getseller
    loadseg "20shop" & roll(6) + 2 & ".seg", 40, 20, 1, tilet, ovrtile, getcargoseller
    loadseg "10shop" & roll(6) + 2 & ".seg", 25, 40, 2, tilet, ovrtile, getseller
End If

'Elongated Station -- Four Rooms plus cargo
If stationtype = 2 Then
    
    mapx = 80: mapy = 80
    ReDim map(mapx, mapy)
    
    loadseg "10shop" & roll(6) + 2 & ".seg", 1, 1, 0, tilet, ovrtile, getseller
    loadseg "10TJunc.seg", 1, 11, 3, tilet, ovrtile, getseller
    loadseg "10shop" & roll(6) + 2 & ".seg", 1, 21, 2, tilet, ovrtile, getseller
    loadseg "10IJunc.seg", 11, 11, 1, tilet, ovrtile, getseller
    loadseg "10IJunc.seg", 21, 11, 1, tilet, ovrtile, getseller
    loadseg "10XJunc.seg", 31, 11, 0, tilet, ovrtile, getseller
    loadseg "20shop" & roll(6) + 2 & ".seg", 41, 16, 1, tilet, ovrtile, getcargoseller
    loadseg "10shop" & roll(6) + 2 & ".seg", 31, 11, 0, tilet, ovrtile, getseller
    loadseg "10shop" & roll(6) + 2 & ".seg", 31, 31, 2, tilet, ovrtile, getseller
    
End If

Close #1

End Sub

Function subseg(minrooms, tilet, ovrtile, X, Y, Optional salesobj As String, Optional signobj As String, Optional bij1 As String, Optional bij2 As String)
'Static numsegs

'If numsegs < minrooms Then
'    aroll = roll(6)
'    if aroll=1 then loadseg "10XJunc.seg", x,y,
    
'End If

End Function

Function genstationmap(minrooms)


'zoing = Int(minrooms / 2 + 1)

'mapx = zoing * 10
'mapy = zoing * 10

'ReDim map(mapx, mapy)

'Dim segz(zoing, zoing) As genroomT

'Place rooms
'For a = 1 To minrooms
'5     xr = roll(zoing - 2) + 1
'    yr = roll(zoing - 2) + 1
'    If segz(xr, yr).seg = "" And segz(xr, yr - 1).seg = "" And segz(xr, yr + 1).seg = "" And segz(xr - 1, yr).seg = "" And segz(xr + 1, yr).seg = "" Then segz(xr, yr) = "10shop" & roll(6) + 2 & ".seg": segz(xr, yr).rot = roll(4) - 1 Else GoTo 5
'Next a

'Link rooms

'Find room to link
'roomx = 0: roomy = 0
'For a = 1 To zoing
'    For b = 1 To zoing
'
'    If segz(a, b).seg = "" Then
'    X = a: Y = b
'    xyrot segz(a, b).rot, X, Y
'    If Not segs(X, Y).seg = "" Then GoTo 5 Else roomx = a: roomy = b: Exit For
'    End If
'
'5     Next b
'If roomx > 0 And roomy > 0 Then Exit For
'Next a
'
''Link room to center
'If roomx > 0 And roomy > 0 Then'
'    X = roomx: Y = roomy
'    xyrot segz(roomx, roomy).rot, X, Y
'    curx = X
'    cury = Y
'    cent = Int(zoing / 2)
'
'8    If curx = cent And cury = cent Then GoTo 7
    
    
'    If curx > Int(zoing / 2) Then zx = 1 Else If curx < Int(zoing / 2) Then zx = -1
'    If cury > Int(zoing / 2) Then zy = 1 Else If cury < Int(zoing / 2) Then zy = -1
    
'    If Not segz(roomx + zx).seg = "" Then cury = cury + zy: GoTo 11
'    If Not segz(roomx + zy).seg = "" Then curx = curx + zx: GoTo 11
    
'11  segz(curx, cury).seg = "Connect"
    
'9     GoTo 8
        
'7 End If




End Function

Function xyrot(ByVal rot, ByRef X, ByRef Y)

If rot = 0 Then Y = Y + 1
If rot = 1 Then X = X - 1
If rot = 2 Then Y = Y - 1
If rot = 3 Then X = X + 1

End Function

Function getseller()
    name = getname
    f = createobjtype(name, "girl" & roll(10) & ".bmp", roll(255), roll(255), roll(255), roll(10) / 10, 2)
    addeffect f, getgener("SELLCLOTHES", "SELLARMOR", "SELLWEAPONS", "SELLPOTIONS") 'getstr2(btype, 0), getstr2(btype, 3), getstr2(btype, 4)
    'createobj objtypes(f).name, X + Int(xsize / 2), Y + Int(ysize / 2), objtypes(f).name
    getseller = name
End Function

Function getcargoseller()
    name = getname
    f = createobjtype(name, "girl" & roll(10) & ".bmp", roll(255), roll(255), roll(255), roll(10) / 10, 2)
    addeffect f, "SELLCARGO"
    'addeffect f, getgener("SELLCLOTHES", "SELLARMOR", "SELLWEAPONS", "SELLPOTIONS") 'getstr2(btype, 0), getstr2(btype, 3), getstr2(btype, 4)
    'createobj objtypes(f).name, X + Int(xsize / 2), Y + Int(ysize / 2), objtypes(f).name
    getcargoseller = name
End Function

Function getcargoprice(cargoname)

'Set price here according to cargo type

price = getprice(cargoname)

needs = 0

For a = 1 To UBound(starbase.cargoname())
    If starbase.cargoname(a) = cargoname Then amount = starbase.cargoamt(a)
Next a

For a = 1 To 5
    If starbase.convertsfrom(a, 1) = cargoname Then price = price * 1.5: needs = starbase.convertsfrom(a, 2)
    If starbase.convertsto(a, 1) = cargoname Then price = price * 0.7: needs = -1
    If starbase.produces(a, 1) = cargoname Then price = price * 0.5: needs = -1
    If starbase.consumes(a, 1) = cargoname Then price = price * 2: needs = starbase.consumes(a, 2)
Next a

If needs > 0 Then mult = 1 + ((needs * 5) / greater(amount, 1)) / 50
If needs = -1 Then mult = 1 - greater(0.25, (amount / 500))
If needs = 0 Then mult = 1 - greater(0.25, (amount / 1000))

getcargoprice = Int(price * mult)

End Function


Function getprice(cargoname)

price = 50

filen = getfile("cargotypes.txt", "VRPG.dat")
Open filen For Input As #1

Do While Not EOF(1)
    Input #1, gor
    If gor = "Cargo" Then Input #1, cname, cworth
    If cname = cargoname Then getprice = cworth: Close #1: Exit Function
Loop

Close #1

Select Case cargoname
    Case "Iron Ore": price = 20
    Case "Iron": price = 40
    Case "Steel": price = 60
    Case "Gold Ore": price = 40
    Case "Gold": price = 80
    Case "Raw Xodium Crystals": price = 400
    Case "Xodium Crystals": price = 800
    Case "Raw Emerald": price = 200
    Case "Emeralds": price = 600

End Select

getprice = price

End Function

Function randomcargo(Optional worthlimit = 0) As String

Dim zoink() As String
Dim zoinkworth() As Long

filen = getfile("cargotypes.txt", "VRPG.dat")
Open filen For Input As #1

lastun = 0

Do While Not EOF(1)
    Input #1, gor
    If gor = "Cargo" Then
    lastun = lastun + 1
    ReDim zoink(1 To lastun)
    ReDim zoinkworth(1 To lastun)
    Input #1, zoink(lastun), zoinkworth(lastun)
    End If
Loop

Close #1

5 gurb = roll(lastun)

If worthlimit > 0 Then If zoinkworth(gurb) > worthlimit Then worthlimit = worthlimit + 1: GoTo 5

randomcargo = zoink(gurb)

End Function

Function loadpart(partname, Optional partsfile = "partsdata.txt") As systype

Dim rpart As systype

'Grab parts until end of file
Do While filegrab(partsfile, "#PART", , , gpartname, parttype, graphname, partdescription)
    'when part is found, load all effects then exit
    If gpartname = partname Then
        rpart.name = partname: rpart.slottype = parttype: rpart.graphname = graphname: rpart.desc = partdescription
        zark = 1
        Do While filegrab(partsfile, "#EFFECT", , "#PART", rpart.eff(zark, 1), rpart.eff(zark, 2), rpart.eff(zark, 3), rpart.eff(zark, 4), rpart.eff(zark, 5)) > 0
        zark = zark + 1
        Loop
        Exit Do
    End If
Loop

loadpart = rpart

End Function

Function getpart(Optional partname As String)

Dim part As systype
part = loadpart(partname)

For a = 1 To UBound(partstock())
    If partstock(a).name = "" Then partstock(a) = part: Exit For
Next a

End Function

Function genstarshipmap(shipclassname, ByVal wepbays, ByVal shieldbays, ByVal enginebays, ByVal expbays, tile, ovrtile)

spaceon = 0

Dim shipmap(1 To 11, 1 To 11) As String

curx = 6
cury = 6

wepbays = 0: shieldbays = 0: enginebays = 0: expbays = 0
For a = 1 To UBound(plrshipdat.sections())
    If (plrshipdat.sections(a).partsobj(1).slottype = "" And Not plrshipdat.sections(a).partsobj(1).name = "") Or plrshipdat.sections(a).partsobj(1).slottype = "Expansion" Then expbays = expbays + 1
    If plrshipdat.sections(a).partsobj(1).slottype = "Weapon Actuator" Then wepbays = wepbays + 1
    If plrshipdat.sections(a).partsobj(1).slottype = "Shield Coil" Then shieldbays = shieldbays + 1
    If plrshipdat.sections(a).partsobj(1).slottype = "Engine" Then enginebays = enginebays + 1
    If plrshipdat.sections(a).partsobj(1).slottype = "Thruster" Then enginebays = enginebays + 1
Next a

randword shipclassname

ReDim map(110, 110)
mapx = 110: mapy = 110
fillovr (ovrtile)
'Start at center of ship with bridge, then work your way around at random.
plr.X = curx * 10 + 5: plr.Y = cury * 10 + 5
createbuilding 10, 10, , 1, tile, ovrtile, curx * 10, cury * 10, 1
shipmap(curx, cury) = "Bridge"


createobjtype "WeaponBay", "partsbay.bmp", 255, 0, 0
addeffect "WeaponBay", "SystemBay", 1
createobjtype "ShieldBay", "partsbay.bmp", 0, 150, 0
addeffect "ShieldBay", "SystemBay", 1
createobjtype "EngineBay", "partsbay.bmp", 0, 0, 255
addeffect "EngineBay", "SystemBay", 1
createobjtype "ExpBay", "partsbay.bmp", 250, 250, 250
addeffect "ExpBay", "SystemBay", 1


'Do While wepbays > 0
6    If roll(2) = 1 Then xroll = roll(3) - 2: yroll = 0 Else xroll = 0: yroll = roll(3) - 2
    If shipmap(curx + xroll, cury + yroll) = "" And curx + xroll < 11 And cury + yroll < 11 Then
    curx = curx + xroll: cury = cury + yroll
    createbuilding 10, 10, , 1, tile, ovrtile, curx * 10, cury * 10, 4
    
    If wepbays > 0 Then
        createobj "WeaponBay", curx * 10 + 3, cury * 10 + 5, , lastbay: lastbay = lastbay + 1
        wepbays = wepbays - 1
        If wepbays > 0 Then
            createobj "WeaponBay", curx * 10 + 7, cury * 10 + 5, , lastbay: lastbay = lastbay + 1
            wepbays = wepbays - 1
        End If
        shipmap(curx, cury) = "WeaponBay"
        GoTo 6
    End If

    If shieldbays > 0 Then
        createobj "ShieldBay", curx * 10 + 3, cury * 10 + 5, , lastbay: lastbay = lastbay + 1
        shieldbays = shieldbays - 1
        If shieldbays > 0 Then
            createobj "ShieldBay", curx * 10 + 7, cury * 10 + 5, , lastbay: lastbay = lastbay + 1
            shieldbays = shieldbays - 1
        End If
        shipmap(curx, cury) = "ShieldBay"
        GoTo 6
    End If

    If enginebays > 0 Then
        createobj "EngineBay", curx * 10 + 3, cury * 10 + 5, , lastbay: lastbay = lastbay + 1
        enginebays = enginebays - 1
        If enginebays > 0 Then
            createobj "EngineBay", curx * 10 + 7, cury * 10 + 5, , lastbay: lastbay = lastbay + 1
            enginebays = enginebays - 1
        End If
        shipmap(curx, cury) = "EngineBay"
        GoTo 6
    End If

    If expbays > 0 Then
        createobj "ExpBay", curx * 10 + 3, cury * 10 + 5, , lastbay: lastbay = lastbay + 1
        expbays = expbays - 1
        If expbays > 0 Then
            createobj "ExpBay", curx * 10 + 7, cury * 10 + 5, , lastbay: lastbay = lastbay + 1
            expbays = expbays - 1
        End If
        shipmap(curx, cury) = "ExpBay"
        GoTo 6
    End If
Else: GoTo 6
End If
'Loop



End Function


Function pricepart(part As systype) As Long

regen = getsys(part, "Shield", 2)

End Function

Function createfaction(ByRef faction As factionT, name, fleetfile, color, fleet, size, resources, wealth, diplomacy, Optional homeworldx = 0, Optional homeworldy = 0, Optional alignment = 50, Optional conversname = "")

If conversname = "" Then conversname = name

If homeworldx = 0 Then
homeworldx = roll(400000) + 50000 '100,000-400,000
'If homeworldy = 0 Then
homeworldy = roll(400000) + 50000
End If

With faction
    .color = color
    .name = name
    .fleet = fleet
    .size = size
    .origsize = size
    .resources = resources
    .wealth = wealth 'How quickly they replenish resources
    .diplomacy = diplomacy
    .X = homeworldx
    .origx = homeworldx
    .Y = homeworldy
    .origy = homeworldy
    .alignment = alignment
    .conversname = conversname 'Conversation tree name
    'shiptypes(1 To 5) As shiptype '1 is smallest ship, 5 is largest
End With

getrgb color, red, green, blue

If fleetfile = "" Then fleetfile = "genAship"

If Not faction.name = "The Vortex" Then
    
    createfacship faction.shiptypes(1), name, name & " Fighter", fleetfile & "1s.bmp", 8, 2, , , , , , , , , , , 9, , , , , , , red, green, blue
    
    createfacship faction.shiptypes(2), name, name & " Destroyer", fleetfile & "2s.bmp", , , 1, 60, , , , 70, , , 8, 4, 5, , , , , , , red, green, blue
    
    createfacship faction.shiptypes(3), name, name & " Cruiser", fleetfile & "3s.bmp", 5, , 2, 90, 15, , , 110, , 4, 6, , , , , , , , , red, green, blue
    addweptoship faction.shiptypes(3), 2, "Turret", 1, 6, 2, 9, 14, 100, 1, 0, 15, 2, 1
    
    createfacship faction.shiptypes(4), name, name & " Freighter", fleetfile & "4s.bmp", 4, 12, , 120, 20, , , 300, , , , , , , , , , , , red, green, blue
    addweptoship faction.shiptypes(4), 2, "Turret", 2, 6, 0, 9, 14, 100, 1, 0, 30, 4, 1
    addweptoship faction.shiptypes(4), 3, "Turret", 2, 6, 0, 9, 14, 100, 1, 0, 30, 4, 1
    
    
    createfacship faction.shiptypes(5), name, name & " Capital Ship", fleetfile & "5s.bmp", 3, 10, 3, 300, 12, , 5, 500, , 4, 10, 10, , , , , , , , red, green, blue
    addweptoship faction.shiptypes(5), 2, "Turret", 4, 6, 4, 5, 14, 100, 1, 0, 40, 8, 1
    addweptoship faction.shiptypes(5), 3, "Turret", 4, 6, 4, 5, 14, 100, 1, 0, 40, 8, 1
    addweptoship faction.shiptypes(5), 4, "Turret", 1, 6, 2, 9, 14, 100, 1, 0, 15, 2, 1

Else:
    red = 0
    green = 0
    blue = 0
    
    faction.shiptypes(1).ismonster = 1
    faction.shiptypes(2).ismonster = 1
    faction.shiptypes(3).ismonster = 1
    faction.shiptypes(4).ismonster = 1
    faction.shiptypes(5).ismonster = 1
    
    createfacship faction.shiptypes(1), name, "Monster", "girlspacemonsB.bmp", 8, 2, , 0, 0, , , 400, , , , , 9, , , , , , , red, green, blue

    createfacship faction.shiptypes(2), name, "Monster", "spacemon-amoebas.bmp", 8, 2, , 0, 0, , , 300, , , , , 9, , , , , , , red, green, blue
    
    createfacship faction.shiptypes(3), name, "Monster", "girlspacemons.bmp", 8, 2, , 0, 0, , , 600, , , , , 9, , , , , , , red, green, blue
    createfacship faction.shiptypes(4), name, "Monster", "girlspacemons.bmp", 8, 2, , 0, 0, , , 800, , , , , 9, , , , , , , red, green, blue
    createfacship faction.shiptypes(5), name, "Monster", "girlspacemons.bmp", 8, 2, , 0, 0, , , 1200, , , , , 9, , , , , , , red, green, blue
    
    

End If

End Function

Function CreateUniverse()

'PLANET, x, y, planetfile
'PORTAL, x, y, , destuniverse, tox, toy
'STARBASE

createfaction factions(0), "Neutral", "", RGB(200, 200, 200), 0, 0, 0, 0, 0, 1, 1, 50
createfaction factions(1), "Allied Star Protectorate", "", RGB(145, 145, 185), 400, 200, 100, 10, 50, 210000, 100000, 70
createfaction factions(2), "Clan Ironbelly", "", RGB(20, 20, 145), 280, 180, 90, 10, 30, 390000, 210000, 70
    'createfacship factions(2).shiptypes(1), "Ironbelly Light Cruiser", "ship8s.bmp", , , 3, , 20, , , 60, , , 8, 4, 4
createfaction factions(3), "Cult of Acid", "", RGB(255, 100, 1), 200, 90, 90, 30, 40, 80000, 370000, 60
createfaction factions(17), "Enzyme Sisters", "", RGB(15, 230, 1), 250, 90, 90, 30, 35, 110000, 210000, 65
createfaction factions(4), "Planetary Defense Coalition", "", RGB(30, 160, 0), 320, 220, 80, 8, 60, 250000, 190000, 90
createfaction factions(5), "Interstellar Assault Fleet", "", RGB(185, 10, 2), 700, 275, 200, 10, 30, 210000, 320000, 70
createfaction factions(6), "Star Matriarchy", "", RGB(255, 155, 205), 320, 125, 80, 30, 20, 130000, 300000, 30
createfaction factions(7), "Demon Princesses", "", RGB(120, 0, 0), 180, 90, 80, 10, 30, 290000, 360000, 10
createfaction factions(8), "The Xebbebban Empire", "", RGB(85, 85, 85), 600, 250, 100, 20, 5, 50000, 50000, -100
createfaction factions(9), "The Zadiian Raiders", "", RGB(25, 185, 85), 300, 120, 100, 10, 30, 290000, 440000, -100
createfaction factions(10), "The Zadiian Raiders", "", RGB(25, 185, 85), 300, 120, 100, 10, 30, 380000, 390000, -100
createfaction factions(11), "The Zadiian Raiders", "", RGB(25, 185, 85), 300, 120, 100, 10, 30, 450000, 320000, -100
createfaction factions(12), "The Breasted Regime", "", RGB(125, 85, 75), 350, 160, 100, 8, 10, 180000, 410000, 50
'createfaction factions(13), "Gastro Cartel", "", RGB(120, 170, 0), 180, 100, 50, 20, 5, , , 10
'createfaction factions(14), "Steelbellies", "", RGB(100, 100, 100), 230, 80, 50, 20, 15, , , 20
createfaction factions(15), "Sisters of Sorcery", "", RGB(0, 0, 250), 200, 60, 60, 30, 40, 30000, 360000, 60
createfaction factions(16), "The Vortex", "", RGB(250, 0, 0), 600, 600, 200, 100, 0, 440000, 20000, -200

'Player faction, whatever it's going to be called
createfaction factions(17), "Ladyship of the Stars", "", RGB(250, 100, 200), 0, 0, 0, 0, 0, 1, 1, 90

'Vuchexi
createfaction factions(18), "The Vuchexi", "", RGB(120, 100, 130), 1200, 90, 50, 0, 0, 430000, 80000, 20

    setopinions factions(5), factions(6), -10 'IAF and Matriarchy war
    setopinions factions(5), factions(7), -10 'IAF and Demon Princess war
    setopinions factions(5), factions(12), -10 'IAF/Breasted Regime War
    setopinions factions(3), factions(17), 5 'CoA/Enzyme Sisters aren't fond of each other
    setopinions factions(3), factions(15), 110 'CoA/SoS alliance
    setopinions factions(2), factions(16), -10 'Ironbellies start at war with Vortex
    setopinions factions(2), factions(4), 110 'PDC/Ironbelly alliance
    setopinions factions(2), factions(1), 110 'ASP/Ironbelly alliance
    vortexwar 'Set all factions to war with the Vortex
    setopinions factions(16), factions(18), 90 'Vuchexi and Vortex don't fight
    setopinions factions(9), factions(10), 150
    setopinions factions(9), factions(11), 150
    setopinions factions(10), factions(11), 150 'Zadiians are all allies
    
'totalwar


'createfaction factions(1), "Allied Star Protectorate", RGB(200, 45, 45), 400, 200, 100, 10, 30, roll(100000) + 20000, roll(100000) + 200000
'createfaction factions(2), "Clan Ironbelly", RGB(0, 0, 145), 400, 150, 100, 10, 30, roll(100000) + 200000, roll(100000) + 200000


ReDim planets(1 To 1000)

'planets(a).X = roll(500000): planets(a).Y = roll(500000)
'planets(a).graphname = getgener("starbase3.bmp", "starbase1.bmp", "starbase2.bmp", "planetfire.bmp", "planetfire2.bmp", "planettwilight.bmp")
'planets(a).planetname = getplanetname

planetnum = 1
Do While planetnum < 600
'For a = 1 To UBound(planets()) - 200 'Reserve 200 for starbases, capitals and BS

    medx = roll(450000)
    medy = roll(450000)
    
    If roll(4) = 1 Then
    planets(planetnum).planetname = getplanetname
    planets(planetnum).X = medx + roll(50000): planets(planetnum).Y = medy + roll(50000)
    planets(planetnum).graphname = getgener("planetfire.bmp", "planetfire2.bmp", "planettwilight.bmp")
    planetnum = planetnum + 1
    Else:
    pname = getplanetname
    For a = 1 To roll(8) + 1
    Select Case a
        Case 1: wname = "Alpha "
        Case 2: wname = "Beta "
        Case 3: wname = "Delta "
        Case 4: wname = "Gamma "
        Case 5: wname = "Tau "
        Case 6: wname = "Proxima "
        Case 7: wname = "Eta "
        Case 8: wname = "Nu "
        Case 9: wname = "Epsilon "
    End Select
    planets(planetnum).planetname = wname & pname
    planets(planetnum).X = medx + roll(50000): planets(planetnum).Y = medy + roll(50000)
    planets(planetnum).graphname = getgener("planetfire.bmp", "planetfire2.bmp", "planettwilight.bmp")
    'planets(planetnum).owner = roll(2)
    planetnum = planetnum + 1
    
    'planetnum = planetnum + 1
    Next a
    End If

'Next a
Loop

Do While planetnum < 800
'sx = roll(500000)
'sy = roll(500000)
    planets(planetnum).X = roll(500000): planets(planetnum).Y = roll(500000)
    planets(planetnum).graphname = getgener("starbase3.bmp", "starbase1.bmp", "starbase2.bmp")
    planets(planetnum).data.name = "Starbase"
    zname = getplanetname
    aroll = roll(6)
    Select Case aroll
        Case 1: zname = "Fort " & zname
        Case 2: zname = zname & " Station"
        Case 3: zname = zname & " Station"
        Case 4: zname = zname & " Outpost"
        Case 5: zname = zname & " " & roll(12)
        Case 6: zname = "Port " & zname
    End Select
    planets(planetnum).planetname = zname
    planetnum = planetnum + 1
Loop

For a = 1 To UBound(factions())
6     zdug = factionown((a))
    If zdug < factions(a).size / 90 Then factions(a).size = factions(a).size + 40: GoTo 6

Next a

End Function

Function factionown(facnum As Byte, Optional numplanets = 50) As Integer
'  XXX Returns last planet 'conquered'
'Returns number of planets

'Exit Function

For a = 1 To UBound(planets())
    'If diff(planets(a).X, factions(facnum).X) < factions(facnum).size / 2 _
    'And diff(planets(a).Y, factions(facnum).Y) < factions(facnum).size / 2 Then planets(a).owner = facnum: factionown = a
    If planets(a).owner = 0 Then
        If calcdist(planets(a).X, planets(a).Y, factions(facnum).X, factions(facnum).Y) / 333 < factions(facnum).size Then planets(a).owner = facnum: factionown = a: powned = powned + 1
    Else:
        If calcdist(planets(a).X, planets(a).Y, factions(facnum).X, factions(facnum).Y) / 333 < factions(facnum).size Then
            'Klepto planets if they're closer to you than your enemy
            If calcdist(planets(a).X, planets(a).Y, factions(facnum).X, factions(facnum).Y) < calcdist(planets(a).X, planets(a).Y, factions(planets(a).owner).X, factions(planets(a).owner).Y) Then planets(a).owner = facnum: factionown = a: powned = powned + 1
        End If
        'Lose planets that are yours that are now too far away to rule
        If planets(a).owner = facnum And calcdist(planets(a).X, planets(a).Y, factions(facnum).X, factions(facnum).Y) / 333 > factions(facnum).size Then planets(a).owner = 0
    End If
Next a

factionown = powned

Exit Function

5 pnum = roll(UBound(planets()))

If planets(pnum).owner > 0 Then GoTo 5

planets(pnum).owner = facnum
factions(facnum).X = planets(pnum).X
factions(facnum).Y = planets(pnum).Y

For a = 1 To numplanets
    zark = nearestplanet(planets(pnum))
    planets(zark).owner = facnum
Next a

factionown = zark

End Function

Function nearestplanet(planet As planetT, Optional ptype = "") As Integer

'ptype = planet.data.name

dist2 = 1000000
For a = 1 To UBound(planets())
    dist = greater(diff(planet.X, planets(a).X), diff(planet.Y, planets(a).Y))
    If dist2 > dist And planet.owner <> planets(a).owner Then
    If ptype = "" Or planet.data.name = "" Then dist2 = dist: pnum2 = a
    If ptype = planets(a).data.name Then dist2 = dist: pnum2 = a
    End If
Next a

nearestplanet = pnum2

End Function

Function createstarsystem(numplanets, Optional typestring = "Planets")

'ReDim planets(1 To )



End Function

Function drawstarmap()

nodraw = 0
Dim r4 As RECT
picBuffer.BltColorFill r4, 0 'RGB(0, 10, 145)
picBuffer.SetForeColor RGB(255, 255, 255)
cwhite = RGB(255, 255, 255)
cblue = RGB(10, 20, 255)
'Dim dfont As StdFont
'Set dfont = New StdFont
'dfont.size = 8
'picBuffer.SetFont dfont


'+600 to X rather than 400 is to center the map
For a = 1 To UBound(planets())
    If planets(a).data.name = "" Or planets(a).data.name = "Planet" Then
    picBuffer.SetForeColor factions(planets(a).owner).color
    picBuffer.DrawLine 500 + planets(a).X / 833 - 2, 400 + planets(a).Y / 833, 500 + planets(a).X / 833 + 3, 400 + planets(a).Y / 833
    picBuffer.DrawLine 500 + planets(a).X / 833, 400 + planets(a).Y / 833 - 2, 500 + planets(a).X / 833, 400 + planets(a).Y / 833 + 3
    
    End If
    If planets(a).data.name = "Starbase" Then
    picBuffer.SetForeColor cwhite
    picBuffer.DrawLine 500 + planets(a).X / 833 - 2, 400 + planets(a).Y / 833, 500 + planets(a).X / 833 + 3, 400 + planets(a).Y / 833
    picBuffer.DrawLine 500 + planets(a).X / 833, 400 + planets(a).Y / 833 - 2, 500 + planets(a).X / 833, 400 + planets(a).Y / 833 + 3
    picBuffer.DrawCircle 500 + planets(a).X / 833, 400 + planets(a).Y / 833, 3
    End If
Next a

For a = 1 To UBound(factions())
    If factions(a).size > 0 Then
    picBuffer.SetForeColor factions(a).color
    'picBuffer.DrawCircle factions(a).X, factions(a).Y, factions(a).size
    picBuffer.DrawCircle factions(a).X / 833 + 500, factions(a).Y / 833 + 400, factions(a).size * 0.4
    'picBuffer.DrawCircle factions(a).X, factions(a).Y, factions(a).size / 4
    'picBuffer.DrawCircle factions(a).X, factions(a).Y, factions(a).size / 8
    getrgb factions(a).color, r, g, b
    r = lesser(250, r + 50): b = lesser(250, b + 50): g = lesser(g + 50, 250)
    'drawtext factions(a).name, factions(a).X / 833, factions(a).Y / 833, r, g, b
    drawtext factions(a).name, factions(a).X / 833 - (2 * Len(factions(a).name)) + 100, factions(a).Y / 833 - 5, r, g, b, True
    'drawtext "BADASS", 5, 5, r, g, b
    'picBuffer.drawtext factions(a).X / 640 + 400, factions(a).Y / 480 + 100, factions(a).name, True
    End If
Next a

For a = 1 To UBound(planets())
    If diff(planets(a).X / 833, starmapx - 100) < 5 And diff(planets(a).Y / 833, starmapy) < 5 Then drawtext planets(a).planetname, planets(a).X / 833 + 100, planets(a).Y / 833 - 20, 0, 250, 50 ' picBuffer.SetForeColor RGB(255, 0, 0): picBuffer.drawtext 400 + planets(a).X / 625 - 2, 400 + planets(a).Y / 833, planets(a).planetname, False: Debug.Print planets(a).planetname: picBuffer.SetForeColor RGB(255, 255, 255)
Next a

picBuffer.SetForeColor RGB(255, 255, 0)
picBuffer.DrawCircle plrship.X / 833 + 500, plrship.Y / 833 + 400, 6
picBuffer.DrawLine plrship.X / 833 + 500 - 5, plrship.Y / 833 + 400 - 5, plrship.X / 833 + 500 + 5, plrship.Y / 833 + 400 + 5
picBuffer.DrawLine plrship.X / 833 + 500 + 5, plrship.Y / 833 + 400 - 5, plrship.X / 833 + 500 - 5, plrship.Y / 833 + 400 + 5
drawtext plr.name, plrship.X / 833 + 100, plrship.Y / 833 + 10, 255, 255, 0

drawtext "Stardate " & plr.stardate, 350, 550, 255, 255, 255, True

blt Bridge.Picture1

End Function

Function addship()
curfac = 0
lastdist = 1000000
'Pick closest faction
For a = 1 To UBound(factions())
    
    curdist = calcdist(plrship.X, plrship.Y, factions(a).X, factions(a).Y)
    If calcdist(plrship.X, plrship.Y, factions(a).X, factions(a).Y) > factions(a).size * 833 Then curdist = 1000000
    If curdist < lastdist Then If curdist < factions(a).size * 640 Then curfac = a: lastdist = curdist

Next a
If curfac = 0 Then Exit Function
shipsize = 1
5 If shipsize < 5 And roll(2) = 1 Then shipsize = shipsize + 1: GoTo 5

createshipbytype factions(curfac).shiptypes(shipsize), plrship.X + roll(10000) - 5000, plrship.Y + roll(10000) - 5000

End Function

Function createshipbytype(shipclass As shiptype, X, Y)

For a = 1 To UBound(ships())

    If ships(a).X = 0 Then ships(a) = shipclass: ships(a).X = X: ships(a).Y = Y: ships(a).owner = 2: Exit Function

Next a

ReDim Preserve ships(1 To a)
ships(a) = shipclass: ships(a).X = X: ships(a).Y = Y: ships(a).owner = 2

End Function

Function countships()
scout = 0
For a = 1 To UBound(ships())
    If ships(a).X > 0 And ships(a).owner = 0 Then scount = scount + 1
Next a

countships = scount

End Function


Function genmapfromship(ship1 As shiptype, Optional ovrtile = 26, Optional maptile = 20, Optional sizemult = 1)

Form1.Picture1.Width = shipgraphs(ship1.graphnum).CellWidth * sizemult
Form1.Picture1.Height = shipgraphs(ship1.graphnum).CellHeight * sizemult

shipgraphs(ship1.graphnum).DrawtoDC Form1.Picture1.hDC, 0, 0, 18, , , (sizemult)

mapx = Form1.Picture1.Width
mapy = Form1.Picture1.Height

ReDim map(mapx, mapy)

createdungeon maptile, ovrtile, 6, 6, mapx - 5, mapy - 5

plr.X = mapx / 2: plr.Y = mapy / 2

For a = 1 To mapx
For b = 1 To mapy
    
    If Form1.Picture1.Point(a, b) = 0 Then
        map(a, b).ovrtile = 0
        map(a, b).tile = 0
        For c = -1 To 1
            For d = -1 To 1
        If Not Form1.Picture1.Point(a + c, b + d) = 0 Then map(a, b).ovrtile = ovrtile: Exit For: Exit For
            Next d
        Next c
    Else:
    map(a, b).tile = maptile
    End If
        
Next b
Next a

updatmap
checkaccess2

makeminimap

End Function

Function addshipsystomap()

'Places bays where walls are, at random

createobjtype "WeaponBay", "partsbay.bmp", 255, 0, 0
addeffect "WeaponBay", "SystemBay", 1
createobjtype "ShieldBay", "partsbay.bmp", 0, 150, 0
addeffect "ShieldBay", "SystemBay", 1
createobjtype "EngineBay", "partsbay.bmp", 0, 0, 255
addeffect "EngineBay", "SystemBay", 1
createobjtype "ExpBay", "partsbay.bmp", 250, 250, 250
addeffect "ExpBay", "SystemBay", 1

wepbays = 0: shieldbays = 0: enginebays = 0: expbays = 0
For a = 1 To UBound(plrshipdat.sections())
    If (plrshipdat.sections(a).partsobj(1).slottype = "" And Not plrshipdat.sections(a).partsobj(1).name = "") Or plrshipdat.sections(a).partsobj(1).slottype = "Expansion" Then expbays = expbays + 1
    If plrshipdat.sections(a).partsobj(1).slottype = "Weapon Actuator" Then wepbays = wepbays + 1
    If plrshipdat.sections(a).partsobj(1).slottype = "Shield Coil" Then shieldbays = shieldbays + 1
    If plrshipdat.sections(a).partsobj(1).slottype = "Engine" Then enginebays = enginebays + 1
    If plrshipdat.sections(a).partsobj(1).slottype = "Thruster" Then enginebays = enginebays + 1
Next a

5
x1 = roll(mapx - 4) + 2
y1 = roll(mapy - 4) + 2
tries = tries + 1
If tries > 20000 Then Exit Function
If map(x1, y1).ovrtile = 0 Or map(x1, y1).object > 0 Or map(x1, y1).tile = 0 Then GoTo 5

'Check to make sure it's not buried
If map(x1 + 1, y1).ovrtile > 0 And map(x1, y1 - 1).ovrtile > 0 And map(x1 - 1, y1).ovrtile > 0 And map(x1, y1 + 1).ovrtile > 0 Then GoTo 5

If shieldbays > 0 Then createobj "ShieldBay", x1, y1, , lastbay: shieldbays = shieldbays - 1: lastbay = lastbay + 1: map(x1, y1).ovrtile = 0: GoTo 5
If wepbays > 0 Then createobj "WeaponBay", x1, y1, , lastbay: wepbays = wepbays - 1: lastbay = lastbay + 1: map(x1, y1).ovrtile = 0: GoTo 5
If enginebays > 0 Then createobj "EngineBay", x1, y1, , lastbay: enginebays = enginebays - 1: lastbay = lastbay + 1: map(x1, y1).ovrtile = 0: GoTo 5
If expbays > 0 Then createobj "ExpBay", x1, y1, , lastbay: expbays = expbays - 1: lastbay = lastbay + 1: map(x1, y1).ovrtile = 0: GoTo 5
'If shieldbays > 0 Then createobj "ShieldBay", x1, y1, , lastbay: shieldbays = shieldbays - 1: lastbay = lastbay + 1: map(x1, y1).ovrtile = 0: GoTo 5
        


End Function

Function nearestship(ship1 As shiptype) As shiptype
'returns the nearest ship to you, more or less

lastship = 0
lastdist = 10000
For a = 1 To UBound(ships())
    If ships(a).X = ship1.X Then GoTo 5 'Don't select yourself
    dist = diff(ships(a).X, ship1.X)
    dist = greater(diff(ships(a).Y, ship1.Y), dist)
    
    If dist < lastdist Then lastship = a
5 Next a
If lastship = 0 Then Exit Function
nearestship = ships(lastship)
End Function

Function loadcomms()
'Communications with nearest ship
Dim nearship As shiptype
nearship = nearestship(plrship)
If nearship.faction = "" Then Exit Function
showform10 nearship.faction, 1, , , "Spaceconversations.txt"

End Function

Function spaceturn()
'Space turns--commit wars, change news, etc.

'Public Type factionT
'    name As String
'    color As Long
'    fleet As Single
'    size As Single
'    resources As Single
'    wealth As Single 'How quickly they replenish resources
'    diplomacy As Single
'    alignment As Integer
'    likesplayer As Single 'How much they like the player
'    opinions(1 To 50) As Single 'how much they like/dislike the other factions (-1 or more means at war)
'    x As Single
'    y As Single
'    conversname As String 'Conversation tree name
'    shiptypes(1 To 5) As shiptype '1 is smallest ship, 5 is largest
'End Type

'Exit Function

plr.stardate = plr.stardate + 0.1

For a = 1 To UBound(factions())
    
    If factions(a).X <= 0 Then GoTo 7
    If factions(a).size <= 0 Then GoTo 7 'Else factions(a).x = roll(450000): factions(a).y = roll(450000): factions(a).size = 15: factions(a).fleet = 50
    'Gain 10% of size in fleet each turn
    factions(a).fleet = factions(a).fleet + factions(a).resources / 10
    'Fleet can never be more than quadruple resources
    If factions(a).fleet > factions(a).resources * 4 Then factions(a).fleet = factions(a).resources * 4
    
    'Constant resource growth for the Vortex
    If factions(a).name = "The Vortex" Then factions(a).resources = factions(a).resources + 1

    atwar = 0
    decwar = 0
    For b = 1 To UBound(factions())
        If b = a Then GoTo 5 'Don't do anything with yourself
        If factions(b).size <= 0 Then GoTo 5
        'If you're already at war with faction, go ahead and fight 'em
        If factions(a).opinions(b) < 0 Then
        factionbattle factions(a), factions(b)
        decwar = decwar - 3 '(Factions will tend not to want to declare war when they're already at war)
        If Not factions(b).name = "The Vortex" Then atwar = 1
        factions(a).opinions(b) = factions(a).opinions(b) - 1
        factions(b).opinions(a) = factions(b).opinions(a) - 1
        End If
        
5     Next b
        
        

    'Your territory can never be more than twice the size of your fleet.  You will only grow or shrink when at war.
    
    'If fleet is bigger than size, expand
    If atwar = 0 Then
        If factions(a).origx > factions(a).X Then factions(a).X = factions(a).X + 1000
        If factions(a).origy > factions(a).Y Then factions(a).Y = factions(a).Y + 1000
        If factions(a).origx < factions(a).X Then factions(a).X = factions(a).X - 1000
        If factions(a).origy < factions(a).Y Then factions(a).Y = factions(a).Y - 1000
        If factions(a).size < factions(a).origsize Then factions(a).size = factions(a).size + 1
        changesize = 1
    End If
    
    If factions(a).fleet > factions(a).size * 2 And atwar = 1 Then factions(a).size = factions(a).size + 3: changesize = 1 ': factions(a).fleet = factions(a).fleet - 24
    
    If factions(a).size > factions(a).fleet * 2 Then
        loss = diff(factions(a).size, factions(a).fleet * 2)
        factions(a).size = factions(a).size - loss / 5: factions(a).fleet = factions(a).fleet + loss: changesize = 1
    End If
    
    If factions(a).fleet < factions(a).size / 3 Then factions(a).fleet = factions(a).size / 3
    If factions(a).size < 10 Then factions(a).size = 0: changesize = 1 'Factions are wiped out if their size is 4 or less
    If changesize = 1 Then factionown (a)
    'If factions(a).size <= 0 Then factions(a).resources = 0
7 Next a



End Function

Function factionbattle(faction1 As factionT, faction2 As factionT)
'First, calculate distance

dist = calcdist(faction1.X, faction1.Y, faction2.X, faction2.Y) / 833 '833 is the generic divisor, because the starmap is 450,000 space units wide and high.  .04 is the faction size display modifier

'Severity of damage depends on how much the factions overlap
maxdist = greater(faction1.size, faction2.size) * 0.4
mindist = lesser(faction1.size, faction2.size) * 0.4

dist = dist - mindist
'If faction1.name = "Clan Ironbelly" Then Stop

mult = maxdist / dist
If mult < 1 Then
If faction1.name = "The Vortex" Or faction2.name = "The Vortex" Or faction2.name = "The Zadiian Raiders" Then GoTo 3
If faction2.X > faction1.X Then faction1.X = faction1.X + 1000
If faction2.Y > faction1.Y Then faction1.Y = faction1.Y + 1000
If faction2.X < faction1.X Then faction1.X = faction1.X - 1000
If faction2.Y < faction1.Y Then faction1.Y = faction1.Y - 1000

3 Exit Function 'No battles if they're too far away, but they'll move towards each other
End If

mult = mult / 4
If mult > 3 Then mult = 3

If mult < 0.4 And faction1.size > 40 And Not (faction1.name = "The Vortex" Or faction2.name = "The Vortex" Or faction2.name = "The Zadiian Raiders") Then
If faction2.X > faction1.X Then faction1.X = faction1.X + 1000
If faction2.Y > faction1.Y Then faction1.Y = faction1.Y + 1000
If faction2.X < faction1.X Then faction1.X = faction1.X - 1000
If faction2.Y < faction1.Y Then faction1.Y = faction1.Y - 1000
End If

'Flee if damage is getting heavy
If faction1.size < 40 And Not faction1.name = "Interstellar Assault Fleet" Then
If faction2.X > faction1.X Then faction1.X = faction1.X - 1000
If faction2.Y > faction1.Y Then faction1.Y = faction1.Y - 1000
If faction2.X < faction1.X Then faction1.X = faction1.X + 1000
If faction2.Y < faction1.Y Then faction1.Y = faction1.Y + 1000
End If

'faction1.size = faction1.size = 3
'faction2.size = faction2.size = 3
'Exit Function

droll1 = roll(7) 'faction1.fleet * ((roll(7) + getallies(faction1)) / 100) * mult 'Max damage is 7% of fleet, which comes out to 14% because both sides attack
droll2 = roll(7) 'faction2.fleet * ((roll(7) + getallies(faction2)) / 100) * mult 'during their faction turn, meaning there are two actual battles

'+1 to damage per ally, regardless of whether they border...

If droll1 > droll2 Then dmg1 = roll(20) + 2 + getallies(faction1) * 2 Else dmg1 = roll(20) + getallies(faction1) * 2
dmg1 = dmg1 + Int(greater(faction1.fleet / 200, faction1.size / 200))

If droll2 > droll1 Then dmg2 = roll(20) + 2 + getallies(faction2) * 2 Else dmg2 = roll(20) + getallies(faction2) * 2
dmg2 = dmg2 + Int(greater(faction2.fleet / 200, faction2.size / 200))

dmg1 = dmg1 * mult

'Clan Ironbelly does triple damage to the Vortex, so they tend to hold them off at first
If faction1.name = "Clan Ironbelly" And faction2.name = "The Vortex" Then dmg1 = dmg1 * 3: dmg2 = 0

dmg2 = dmg2 * mult

faction1.fleet = faction1.fleet - dmg2
faction2.fleet = faction2.fleet - dmg1

'1 size changes hands according to who did more damage
'If dmg2 < dmg1 Then faction1.size = faction1.size + 1: faction2.size = faction2.size - 1
'If dmg1 < dmg2 Then faction1.size = faction1.size + 1: faction2.size = faction2.size - 1

'faction1.size = faction1.size - dmg2 / 2
'faction2.size = faction2.size - dmg1 / 2

End Function

Function getfaction(factionname) As factionT

For a = 1 To UBound(factions())
    If factions(a).name = factionname Then getfaction = factions(a): Exit Function
Next a

End Function

Function getfactionnum(factionname)

For a = 1 To UBound(factions())
    If factions(a).name = factionname Then getfactionnum = a: Exit Function
Next a

End Function

Function getallies(faction As factionT)
'Returns the number of allies the faction has

allies = 0
For a = 1 To UBound(factions())
    If faction.opinions(a) >= 100 Then allies = allies + 1
Next a

getallies = allies

End Function

Function setopinions(fac1 As factionT, fac2 As factionT, opinion)
'Shorthand way of setting two faction's opinions of each other to match
fac1.opinions(getfactionnum(fac2.name)) = opinion
fac2.opinions(getfactionnum(fac1.name)) = opinion
End Function

Function totalwar()

For a = 1 To UBound(factions)
    For b = 1 To UBound(factions(a).opinions())
        factions(a).opinions(b) = -50
    Next b
Next a

End Function

Function warwithall(facnum)

a = facnum
    For b = 1 To UBound(factions(a).opinions())
        factions(a).opinions(b) = -50
    Next b


End Function

Function vortexwar()
    a = getfactionnum("The Vortex")
    For b = 1 To UBound(factions())
    If b = a Then GoTo 5
        setopinions factions(a), factions(b), -50
5     Next b

End Function
