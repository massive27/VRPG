Attribute VB_Name = "VRPG"

'Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
' Parameters - handle, ID, MilSecs, address of handler
'SetTimer Me.hwnd, 50, 200, AddressOf TimerHandler
' Parameters - handle, ID
'KillTimer Me.hwnd, 50

'Public Sub TimerHandler(ByVal hwnd As Long, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTime As Long)
' This is our time callback event handler
'battle2.MoveTicker
'End Sub

'm'' CONDITIONAL COMPILATION
'm'' because i'm borded to duplicate the source, here is a
'm'' conditional compilation marker, in order to keep track on
'm'' difference between "genuine" gameplay (and bug...) and the "revamped" gameplay
'm'' I didnt add the 'm'' remark on conditional compilation block because the '#' will do.

#Const USELEGACY = 0

#If USELEGACY = 1 Then
Public Const curversion = "2.14 Legacy" 'm'' modified to know where we are
#Else
Public Const curversion = "2.14 Modded" 'm'' the other one !
#End If

Public Const debugmessageson& = 0 '1
Public Const editon = 0

Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
' Parameters - handle, ID, MilSecs, address of handler
'SetTimer Me.hwnd, 50, 200, AddressOf TimerHandler
' Parameters - handle, ID
'KillTimer Me.hwnd, 50

'Public Sub TimerHandler(ByVal hwnd As Long, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTime As Long)
' This is our time callback event handler
'battle2.MoveTicker
'End Sub

'm'' Declare Sub BASS_Free Lib "bass.dll" () 'm'' useless once DMC removed

Public Type keysettings
    moveN As Integer
    moveNE As Integer
    moveNW As Integer
    moveS As Integer
    moveSE As Integer
    moveSW As Integer
    moveE As Integer
    moveW As Integer
    lifepot As Integer
    manapot As Integer
    eat As Integer
End Type

Public Type charclass
    classname As String
    strmult As Byte
    dexmult As Byte
    intmult As Byte
    endmult As Byte
    school1 As String
    school1mult As Integer
    school2 As String
    school2mult As Integer
    hpmult As Single
    mpmult As Single
End Type


Public Type objecttype
    name As String
    'effect() As String
    effect(1 To 10, 1 To 5) As String
    'graph As cSpriteBitmaps
    graphloaded As Byte '1 means the graphic has been loaded
    graphname As String
    r As Byte
    g As Byte
    b As Byte
    l As Single
    'shove As Byte 'automatically takes it
    cells As Byte
End Type

Public Type clothestypesT
    name As String
    graph As String
    armor As Single
    weight As Single
    wear1 As String
    wear2 As String
    material As String
End Type

Public Type weapontypesT
    name As String
    graph As String
    weight As Single
    dice As Single
    damage As Single
    type As String
    material As String
End Type

Public weapontypes() As weapontypesT
Public clothestypes() As clothestypesT

Public Type materialtypesT
    type As String
    name As String
    r As Integer
    g As Integer
    b As Integer
    armor As Single
    weight As Single
    worth As Single
    goldmult As Single
    effects(1 To 3, 1 To 5) As String
End Type

Public materialtypes() As materialtypesT

Public Type auraT
    loaded As Byte
    graphs As cSpriteBitmaps
    obj As objecttype
    cell As Byte
    type As String
    duration As Byte
End Type

Public auras(1 To 8) As auraT

Public objgraphs() As cSpriteBitmaps

Public Type shootyt
    Active As Byte
    X As Single
    Y As Single
    xplus As Single
    yplus As Single
    frame As Byte
    'angle As Integer
    graphnum As Byte
    is As String  'Dice:Damage:Other
    Time As Integer  'How long it will last
    xhit As Integer
    yhit As Integer 'Last spot they hit a monster in--for piercing
End Type

Public Type bijdangsprites
    cS As cSpriteBitmaps
    sprite As cSprite
    X As Integer
    Y As Integer
End Type

Public Type bijdangsT
    Active As Byte
    graphnum As Integer
    cell As Integer
    X As Integer
    Y As Integer
End Type

Public bijdangs(50) As bijdangsT

Public Type buyarmort
    name As String
    graph As String
    r As Byte
    g As Byte
    b As Byte
    l As Single
    armor As Integer
    gold As Long
    special As String
    wear1 As String 'Or damage, for weapons that don't use D6's
    wear2 As String 'Or weapon type (Sword or bow)
End Type

Public buyarmor(1 To 8) As buyarmort

Public Type wepgraphT
'    graph As cSpriteBitmaps
    graphname As String
    xoff As Integer
    yoff As Integer
    dice As Integer
    damage As Integer
    'bonus As Integer
    type As String
    r As Byte
    g As Byte
    b As Byte
    l As Single
    obj As objecttype
    digged As Integer
    weight As Integer
End Type

Public wepgraph As cSpriteBitmaps
Public wep As wepgraphT

Public Type tiletype
    tile As Byte
    ovrtile As Byte
    blocked As Byte
    monster As Integer
    object As Integer
    used As Byte 'For map generation--means there is already a building there
End Type

Public Type playertype
    name As String
    str As Byte
    strboost As Byte
    dex As Byte
    dexboost As Byte
    int As Byte
    intboost As Byte
    charisma As Byte
    endurance As Byte
    fatigue As Integer
    fatiguemax As Integer
    armorboost As Integer
    hp As Long
    hpmax As Long
    hplost As Long 'Expansion only--HP that cannot be regained via potions
    mpmax As Long
    mp As Long
    X As Integer
    Y As Integer
    exp As Double
    expneeded As Double
    level As Byte
    gp As Long
    xoff As Integer 'offset
    yoff As Integer 'offset
    armor As Integer
    instomach As Integer 'If in monster stomach, and which one
    Class As String
    quest(500) As String
    regen As Byte 'regeneration (in d6's--gained every 10 turns)
    mplost As Long 'MP lost to enchanting (Restored when level is gained)
    curmap As String
    lpotions As Integer
    mpotions As Integer
    lpotionlev As Byte
    mpotionlev As Byte
    diglevel As Byte
    plrdead As Byte
    specials(50, 1) As String
    classskills(1 To 6) As String
    skillpoints As Integer
    foodinbelly As Byte
    monsinbelly As Byte 'foodinbelly cannot fall below monsinbelly
    diggedmons(500) As String 'Monster types you have digested
    alreadysaid(50) As String 'Which conversation things have been said
    difficulty As Byte
    combatskills(1 To 4) As String
    sp As Integer 'Skill points to use on skills
    spmax As Integer
    combatskillpoints As Byte
    classdata As charclass
    'rectum As Integer 'How much the player needs to shit.  When negative, player is stunned because
                      'she is currently shitting
    beento(1 To 150, 1 To 2) As String 'Maps the player has been to and their map filename:X:Y
    hairname As String
    bodyname As String
    haircolor As Long
    timerspeed As Byte
    keys As keysettings
    charpoints As Integer
    stardate As Single
    curquests() As String 'Standard quest is QuestName(make sure it's unique):Map:Type("KILL", "RETRIEVE", "COMPLETED" etc):Extra stuff after this
                                  'KILL:Monstertype:Num (Decremented each time you kill one--when it hits 0, the quest is achieved)
                                  'RETRIEVE:Item:Giveto:Lose Y/N (Have the item in your inventory when you speak to Giveto.  Lose determines if it takes it from your inventory)
                                  'COMPLETED (Quest is completed)
    Swallowtime As Long 'Increases with time and makes player more vulnerable to being eaten
    ExtraInts(1 To 50) As Integer
    ExtraStrings(1 To 50) As String
    ExtraDoubles(1 To 50) As Double
    
End Type

Public Type monstertype
    name As String
    gnum As Integer
    gfile As String
    color As Long
    skill As Byte
    hp As Long
    mp As Long
    exp As Long
    'graph As cSpriteBitmaps
    move As Byte
'    spells(1 To 5) As Integer
'    spellpref(1 To 5) As Byte
'    special(1 To 3) As String
    dice As Byte
    damage As Byte
    eatskill As Byte
    eattype As Byte '0=Standard, 1=Instantaneous/Engulfing (For slimes, vines and such that do not logically have stomachs)
    acid As Byte
    swallow As String
    level As String
    escapediff As Byte
    light As Single
    missileatk As String 'Type:Frequency:Dice:Damage   Note that higher frequency means less shots
    Sound As String
    weaktype As String
    boss As Byte
    trans As Byte '1-polychromatic, 2-metallic, 3-transparent
    colorwhole As Byte
End Type

Public Type carrymonstersT
    montype As monstertype
    numeach As Byte
End Type

Public carrymonsters(1 To 10) As carrymonstersT

Public mongraphs() As cSpriteBitmaps

Public Type clantype
    r As Byte
    g As Byte
    b As Byte
    light As Variant
End Type

Public Type spell
    name As String
    effect(1 To 3, 1 To 3) As String
End Type

Public Type spellT
    name As String
    mp As Variant
    effect As String
    amount As String
    target As String
    school As String
    level As Byte
    spellon As Byte 'If the spell is currently active
    has As Byte 'If the current player has the spell
End Type

Public Type amonsterT
    type As Integer
    hp As Long
    mp As Integer
    X As Integer
    Y As Integer
    xoff As Integer
    yoff As Integer
    cell As Byte
    instomach As Integer 'What other monster's stomach they are in-- -1 means the player
    hasinstomach As Integer 'what monster is in this one's stomach
    stomachlevel As Byte 'Whether the monster is being eaten, has been swallowed etc. works roughly like it does with the
                         'player
    animdelay As Integer 'negative means wait to switch to attack animation,
                         'positive means count down to switch back
    owner As Byte 'if 1, then allied with player. If 2, then can eat other monsters too.
    litx As Integer
    lity As Integer 'Literal X and Y drawing positions, for shooting
    stunned As Byte 'How many turns the monster will be stunned for
End Type

Public Type clothesT
    name As String
    'backgraph As cSpriteBitmaps
    'digbackgraph As cSpriteBitmaps
    loaded As Byte
    wear1 As String
    wear2 As String
    hp As Integer
    armor As Integer
    drawn As Byte
    obj As objecttype
                        'and dropping
    digested As Byte    'if 1, shows digested graphic
    weight As Integer
    raster As Long 'Whether the item is translucent: 1=Translucent
End Type

Public Type clothesgrapht
    graph As cSpriteBitmaps
    diggraph As cSpriteBitmaps
End Type

Public cgraphs(1 To 16) As clothesgrapht

Public Type aobject
    type As Integer
    string As String 'For NPC speech and stuff
    string2 As String 'for NPC graphics
    X As Integer
    Y As Integer
    name As String
    xoff As Integer
    yoff As Integer
    instomach As Integer
End Type

Public Type planetjunk
    moisture  As Byte
    moisturetolerance As Byte
    temperature  As Byte
    temperaturetolerance As Byte
    rockyness As Byte
    rockynesstolerance As Byte
    vegetation As Byte
    vegetationtolerance As Byte
    population As Byte
    populationtolerance As Byte
    name As String
    gfile As String
End Type

Public Type mapjunkT
    maps(1 To 4) As String
    favmonster As String
    favmonsternum As String
    questmonster As String
    terstr As String
    name1 As String
    name2 As String
    name3 As String
    name4 As String
    level As Byte
    planetjunk As planetjunk
End Type

Public Type mapmastertype  'All data needed to save a map
    mapnum As Long 'Gives a number to each map so the correct one can always be loaded
    'mapnums(1 To 4) As Long 'Map numbers to connected maps
    objts As Integer
    objtotal As Long
    totalmonsters As Long
    lastmontype As Integer
    mapx As Integer
    mapy As Integer
    map() As tiletype
    mapjunk As mapjunkT
    mapname As String
    objtypes() As objecttype
    objs() As aobject
    montype() As monstertype
    mon() As amonsterT
End Type

'Public Type roomT
'    eastroom As Long
'    westroom As Long
'    northroom As Long
'    southroom As Long
'    segfile As String
    
'End Type

Public Type textT
    X As Single
    Y As Single
    'size As Byte
    txt As String
    highr As Byte
    highg As Byte
    highb As Byte
    lowr As Byte
    lowg As Byte
    lowb As Byte
    age As Byte
    textid As String
End Type

Public drawtxts(1 To 30) As textT

Public monsterstoput As Byte

Dim cleavage As cSpriteBitmaps
Public plrhair As cSpriteBitmaps
Public plrhairloaded As Byte
Public plrbodyloaded As Byte

Public bijdrawing As Integer

Public Type rainT
    X As Integer
    Y As Integer
    frame As Byte
    stopped As Byte
End Type

'Public rainx(50) As Integer
'Public rainy(50) As Integer
'Public rainxwander(50) As Integer
Public rain(1 To 50) As rainT
Public raincolor As Long
Public rainspeed As Long
Public raindensity As Long
'Public rainwander As Integer
Public raingraph As cSpriteBitmaps

Public mastermap() As mapmastertype 'All maps in the game, stored in one file

Public mapjunk As mapjunkT

Public map() As tiletype
Public mapx As Integer
Public mapy As Integer
Public tilespr As cSpriteBitmaps
Public tilespr2 As cSpriteBitmaps
Public transovrspr As cSpriteBitmaps
Public ovrspr As cSpriteBitmaps
'Public ovrspr2 As cSpriteBitmaps
Public plr As playertype

Public capegraphs As cSpriteBitmaps

Public montype() As monstertype
Public totalmonsters As Integer
Public monstermaps As cSpriteBitmaps
Public mon() As amonsterT
Public lastmon As Long
Public lastmontype As Integer
Public clan(8) As clantype
Public gbody As cSpriteBitmaps

Public clothes(1 To 16) As clothesT

Public Ccom As String 'Current Command
'Public editon As Byte
Public edittile As Byte
Public editmode As Byte
Public waterspr As cSpriteBitmaps
Public swallowcounter As Integer
Public objs() As aobject ' objs(X, 1 to 3) 1 object type, 2 X, 3 Y
'Public npcstrings() As Integer
Public objtypes() As objecttype
Public objts As Integer
Public objtotal As Integer
Public spells() As spellT
Public totalspells As Integer
'Public mapjunk.maps(1 To 4) As String 'maps connected to this one
Public enterfrom As String
Public stilldrawing As Byte
Public turncount As Byte
Public boostcount As Byte
'Public plr.plrdead As Byte
Public plrgraphs As cSpriteBitmaps
Public spellbook(50) As Integer
Public inv(1 To 50) As objecttype
Public bodyimg As String
Public bijdang(1 To 8) As bijdangsprites
Public soundon As Byte
Public digbody(1 To 5) As cSpriteBitmaps
'Public plr.diglevel As Byte 'how digested you are after you're dead
Public turnswitch As Byte 'Timer switch for automatic turns

Public cheaton As Byte
Public winlev As Byte 'how many times player has won game--for access to extra characters
Public allloaded As Byte 'Whether all graphics have been loaded
Public grloading As Byte
Public worlddir As String

Public plrhpgain As Single 'HP being gained slowly through potions and so forth

Public genmaps() As String

Public loadgame As Byte

Public offset As Integer 'Graphics offset (For message box)
Public dummyobj As objecttype
Public dummyobj2 As objecttype
Public talkingto As String 'Object that you're talking to
Public ate(10) As String  'Digestion strings according to what the player has eaten
Public fullbody As cSpriteBitmaps
Public buying As Byte
Public shootygraphs(1 To 8) As cSpriteBitmaps
Public shooties(1 To 100) As shootyt
Public allmove As Byte 'Tell all monsters to move towards the player--for when they're using missile weapons
Public instomachcounter As Integer 'Total turns spent in stomach--it will take half this many turns before you can be swallowed again
Public usingskill As String
Public talkingtonum As Integer
'Public backclothesgraph As cSpriteBitmaps
Public juststarted As Byte

Public extraovrs(1 To 4) As cSpriteBitmaps
Public transextraovrs(1 To 4) As cSpriteBitmaps
Public nodraw As Byte
Public needbodyupdt As Byte

Public nocoloring As Byte

Public dmap() As Byte 'Map for checking accessibility

Public plantprefs() As planetjunk

'Public plr.timerspeed As Byte

Public isexpansion As Byte

Public stomachlevel As Byte 'Whether you are being swallowed, in the stomach or in the intestines

Public updatbonuses As Byte

Const pi = 3.14

Public Const tgrass = 1
Public Const tdesert = 2
Public Const tstonydirt = 3
Public Const twater = 8
Public Const tswamp = 9
Public Const tdirt = 12
Public Const tmud = 13
Public Const tblackstone = 14
Public Const tstone = 15
Public Const tlavarock = 17
Public Const tlava = 18
Public Const tsand = 27
Public Const tarid = 29
Public Const tsnow = 31
Public Const tice = 43
Public Const tsuperarid = 47

Public Const orocks = 5
Public Const oplant = 8
Public Const oflowers = 20
Public Const ograss = 17
Public Const ostonewall = 9
Public Const odirtwall = 10

Public Type statusbarsT
    life As cSpriteBitmaps 'Red Life
    life2 As cSpriteBitmaps 'Green life
    mana As cSpriteBitmaps
    fatigue As cSpriteBitmaps
    empty As cSpriteBitmaps
End Type

Public minimapspr As cSpriteBitmaps
Public minimapspr2 As cSpriteBitmaps

Public xmapoff() As Integer
Public ymapoff() As Integer 'Offsets for moving stomach walls

Public mouthgraph As cSpriteBitmaps

Public bars As statusbarsT

Public mapbufspr As cSpriteBitmaps

Public paused As Byte

Public struggled As Byte 'Whether the player has struggled--increases damage and rate of digestion

'Public plr.swallowtime As Long 'Constantly increases odds of player being eaten.  To make
'sure there's always some vore.  Reset when player enters stomach of monster.

'Public Type starskietype
'    lifespan As Single
'    x As Integer
'    y As Integer
'    xoff As Single
'    yoff As Single
'    xspd As Single
'    yspd As Single
'    lifespan As Integer
'    color As Long
'End Type

'Public starskies()


Sub Main()
dbmsg "Game launch" 'm'' for the log file entry point.
'On Error Resume Next
Randomize
'ChDir "C:\VB\VRPG\"
dbmsg "Initialializing MPQ control"
initdat Form1.MpqControl1
'ChDir "F:\ExVRPG\Pakfile\"
ChDir App.Path '& "\Pakfile\"

'cheaton = 1

dbmsg "Checking for curgame.dat"
If Not Dir("curgame.dat") = "" Then Kill "curgame.dat"
If Not Dir("plrdat.tmp") = "" Then Kill "plrdat.tmp"


'ChDir "C:\VRPG"
dbmsg "Setting keys"
setplrkeys
'redim wep.obj.effect
updatbonuses = 1
isexpansion = 1



Debugger.CharLoad 'm'' initialization of stitched code...
dbmsg "Filesystem and data.pak check" 'm'' debug info
Debugger.TmpFold_Check App.Path 'm'' temporary folder checking (app.path as of yet)
Debugger.DataPak_Check App.Path 'm'' data.pak file supposed to be in game directory

#If USELEGACY <> 1 Then
'm'' mod manager loading
Debugger.PakCount = 0 'm''
ReDim Debugger.PakFiles(1 To 4) 'm''
'm'' command line analysis
If InStr(1, command, "mod", vbTextCompare) > 0 Then 'm''
    'm'' a mod addon is indicated in the commandline
    tmp = Split(command, " ", , vbBinaryCompare) 'm''
    For i = 0 To UBound(tmp) - 1 'm'' analysis parameters
        If tmp(i) Like "-mod" Then 'm''
            Debugger.PakCount = Debugger.PakCount + 1 'm''
            If PakCount = UBound(PakFiles) Then ReDim Preserve PakFiles(1 To PakCount + 4) As String 'm'' extend mod stack
            Debugger.PakFiles(PakCount) = Trim(tmp(i + 1)) 'm''
        End If 'm''
    Next i 'm''
End If 'm''

'm'' adding main pak file
PakCount = PakCount + 1 'm''
Debugger.PakFiles(PakCount) = "Data.pak" 'm''
#End If

'm'' preparing title screen from Data.pak
tspic = getfile_mod("TitleScreen.jpg") 'm''
If (tspic <> "") Then TitleScreen.Picture = LoadPicture(tspic) 'm''



ReDim plr.curquests(1 To 1)

If plr.timerspeed = 0 Or plr.timerspeed > 50 Then plr.timerspeed = 5

plr.bodyname = "body" & roll(6) & ".bmp"
plr.hairname = "hair" & roll(26) & ".bmp"
plr.haircolor = RGB(roll(155), roll(155), roll(155))

dbmsg "Loading body bitmaps"
Set fullbody = New cSpriteBitmaps
dbmsg "Loading Shooties"
loadshooties
dbmsg "Loading Data Files"
loaddata "Spells.txt"
loadclothestypes "Clothesdata.txt"
dbmsg "Loading Preferences"
loadprefs

'Form10.Show: Form10.Hide
'getmonstats

dbmsg "Switching to main screen" 'm'' debug log
TitleScreen.Show 1
Form1.Text5.Visible = True: DoEvents 'm'' shows the "Loading..." panel for better UI

#If USELEGACY <> 1 Then
    Add_UI.RemOldUI 'm'' remove old textboxes from the old UI
#End If

3 If plr.Class = "" Then newchar
If plr.Class = "" Then GoTo 3

Set cStage = New cBitmap
'cStage.CreateFromFile "sky1.bmp"

cStage.CreateAtSize 800, 600
Form1.Show: Form1.Text5.Visible = True: DoEvents 'm'' shows the "Loading..." panel for better UI
dbmsg "bijdang loading" 'm'' debug log
loadbijdang 'm'' note: loadbijdang is pretty slow.

Set backgr = New cSpriteBitmaps

backgr.CreateFromFile "dirtback2.bmp", 1, 1, , GenRGB(255, 255, 0)


'makebacksprite "map.jpg"

lastsprite = 1
lastmap = 1
Form1.Show
DoEvents 'm'' doevents let UI refresh a bit

'ReDim map(1 To 200, 1 To 200) As tiletype

loadalltiles -1

'Set tilespr = New cSpriteBitmaps
'tilespr.CreateFromFile "tiles1.bmp", 5, 5, , RGB(0, 0, 0)
'Set tilespr2 = New cSpriteBitmaps
'tilespr2.CreateFromFile "tiles2.bmp", 5, 5, , RGB(0, 0, 0)
'Set ovrspr = New cSpriteBitmaps
'ovrspr.CreateFromFile "overlays2.bmp", 5, 10, , RGB(0, 0, 0)
'Set transovrspr = New cSpriteBitmaps
'transovrspr.CreateFromFile "transoverlays2.bmp", 5, 10, , RGB(0, 0, 0)
'Set tilespr = New cSprite
'tilespr.SpriteData = tilespr2
clan(1).r = 55: clan(1).g = 55: clan(1).b = 55
clan(1).light = 0.3

'clan(0).r = 255: clan(0).g = 255: clan(0).b = 255
'clan(0).light = 0.5

clan(0).r = 240: clan(0).g = 110: clan(0).b = 240
clan(0).light = 0.3

'makesprite gbody, Form1.Picture1, "body2.bmp"

makesprite digbody(1), Form1.Picture1, "digbody.bmp"
makesprite digbody(2), Form1.Picture1, "digbody2.bmp"
makesprite digbody(3), Form1.Picture1, "digbody3.bmp"
makesprite digbody(4), Form1.Picture1, "digbody4.bmp"
makesprite digbody(5), Form1.Picture1, "digbody5.bmp"

'addclothes "Bra", "bra5.bmp", 1, "Bra", , 256, 155, 155, 0.5, 1
'addclothes "Panties", "panties2.bmp", 1, "Panties", , 256, 155, 155, 0.5, 1
'addclothes "Dress", "dress2.bmp", 4, "Upper", "Lower", 256, 240, 240, 0.5, 1
'addclothes "Shortcape", "cape1.bmp", 1, "Jacket", , 256, 100, 0, 0.2
'addclothes "Shirt", "shirt1.bmp", 1, "Upper", , 230, 230, 230, 0.5
'addclothes "Longshirt", "longshirt1.bmp", 1, "Upper", , 30, 150, 10, 0.3
'addclothes "Loincloth2", "leathertop2.bmp", 1, "Upper"
'addclothes "Blah", "doublet1.bmp", 1, "Upper", , 256, 250, 250, 0.8
'addclothes "BLARG", "swimsuit1.bmp", 1, "Bra", "Panties"
'addclothes "sadg", "leathersuit1.bmp", 1, "Bra", "Panties", 256
'addclothes "Casodu", "chunli1.bmp", 1, "Upper", "Lower", 256, 200, 0, 0.2
'addclothes "Pants", "pants1.bmp", 3, "Lower", , 0, 125, 255
'addclothes "Bracers", "bracers1.bmp", 1, "Arms", , 256, 60, 60, 0.2
'addclothes "Symbiote", "symbiote1.bmp", 1, "Bra", "Panties", , , , , 1

If loadgame = 0 Then plr.gp = 500


'wep.r = 255: wep.g = 255: wep.b = 255: wep.l = 0.5
'lwepgraph "sword1.bmp", Form1.Picture1

Form1.updatbody

'makesprite waterspr, Form1.Picture1, "underwater1.bmp", 0, 0, 255, 1

soundon = 1

'm'' hide the "Loading..." panel for better UI
Form1.Text5.Visible = False: DoEvents 'm''

If loadgame = 0 Then
    'loaddata "VRPGData.txt"
    loaddata "eggcreche.txt"
    'randommap
    'makesprite plrgraphs, Form1.Picture1, "C:\VB\VRPG\testplayer.bmp", 0, 0, 255, 0.1
    
    plr.X = 19: plr.Y = 19
    If Form1.Command2.Enabled = True Then Form1.Command2.SetFocus
    'And isexpansion = 0
    If cheaton = 0 Then showform10 "Birth", 1, "intro1.jpg": MsgBox "Welcome to Duamutef's Glorious Vore RPG! Click on the Help menu if you don't know how to play."
    'MsgBox "You step into your home town, relieved to be back among the people of your own clan. For weeks you have been hearing stories of a horrible dragon. This dragon, Thirsha, lives in a volcano to the far North. Seers prophesy that soon she will sweep through and kill everyone in the region, including you. You must find the three dragon keys and gain entrance to her lair--it is said that if you are able to do so, you can destroy her."
End If

'Form1.Timer1.Enabled = True
'Form1.Timer2.Enabled = True
'If cheaton = 0 Then plr.timerspeed = 100
If cheaton = 1 Then plr.lpotions = 100 ':plr.str = 75: plr.dex = 75

Form1.Timer1.Enabled = True 'Get rid of this to fix timing later
If Form1.Timer1.Enabled = False Then
Do While Not endingprog >= 1
    
    If paused = 0 Then
        If spaceon = 0 Then
            Timer2Z
        Else
            '''spacetimer
        End If
    End If
    DoEvents
Loop
End If

End Sub

Sub randommap()

mapx = 50
mapy = 50
For a = 1 To mapx
For b = 1 To mapy
    map(a, b).tile = roll(3)
    If roll(16) = 1 Then map(a, b).ovrtile = 2: If map(a, b).ovrtile = 1 Or map(a, b).ovrtile = 2 Then map(a, b).blocked = 1
    If roll(36) = 1 Then createmonster roll(lastmontype), a, b
Next b
Next a

End Sub

Function roll(ByVal damage) As Long 'm'' declare...
damage = Int(damage)
roll = Int((damage - 1 + 1) * Rnd + 1)

End Function

Sub drawall()

Static dick As Byte
If dick = 0 Then makeminimap: dick = 1

DoEvents

Dim drawrange As Long 'm'' added declaration
Dim a As Long, b As Long, c As Long  'm'' added declaration
drawrange = 10

If Form1.Visible = False Then Exit Sub

If nodraw = 1 Then Exit Sub

If stilldrawing > 1 Then Exit Sub
'stilldrawing = stilldrawing + 1

'On Error Resume Next

'On Error GoTo 15
'GoTo 10
'15 Resume Next
'10

For a = 1 To totalmonsters
    If mon(a).type = 0 Then GoTo 72
    getlitxy mon(a).X, mon(a).Y, mon(a).litx, mon(a).lity, mongraphs(mon(a).type), 1
72 Next a

If plr.instomach >= 1 Then
    plr.X = mon(plr.instomach).X: plr.Y = mon(plr.instomach).Y
    If mon(plr.instomach).instomach > 0 Then plr.X = mon(mon(plr.instomach).instomach).X: plr.Y = mon(mon(plr.instomach).instomach).Y
    plr.xoff = plr.xoff + mon(plr.instomach).xoff: plr.yoff = plr.yoff + mon(plr.instomach).yoff
    
    If stomachlevel <= 1 Then mon(plr.instomach).cell = 3 Else mon(plr.instomach).cell = 2
End If

Dim r1 As RECT
'backgr.TransparentDraw picBuffer, 0, 0, 1
picBuffer.blt r1, backgr.DXS, r1, 0


'picBuffer.BltColorFill r1, 0


For a = plr.X - drawrange To plr.X + drawrange '+9
'DoEvents
If a < 1 Or a > mapx Then GoTo 6
    For b = plr.Y - drawrange To plr.Y + drawrange
     If b < 1 Or b > mapy Then GoTo 5
     If map(a, b).tile <= 25 Then drawobj tilespr, a, b, map(a, b).tile
     If map(a, b).tile > 25 Then drawobj tilespr2, a, b, map(a, b).tile Mod 25
5   Next b
6 Next a


 
zex = 0
If map(plr.X, plr.Y).blocked = 2 Then zex = 1

'm'' code analysis : this make the attack picture and still picture switch
For c = 1 To totalmonsters
    If mon(c).animdelay < 0 Then
        mon(c).animdelay = mon(c).animdelay + 1
        If mon(c).animdelay = 0 Then mon(c).cell = 3: mon(c).animdelay = 3
    End If
    If mon(c).animdelay > 0 Then
        mon(c).animdelay = mon(c).animdelay - 1
        If mon(c).animdelay = 0 Then mon(c).cell = 1: If plr.instomach = c Then mon(c).cell = 2
    End If
    
    #If USELEGACY = 1 Then
    If mon(c).hasinstomach > 0 Then
    mon(c).cell = 2
        If mon(mon(c).hasinstomach).stomachlevel < 4 Then mon(c).cell = 3
    End If
    #Else
    'm'' due to the fix letting monster boss able to attack even with a full belly,
    'm'' the cell draw must be = 3 (attack) if there is an attack ongoing
    If mon(c).hasinstomach > 0 Then
        If mon(c).animdelay < 0 Then
            mon(c).cell = 3
        Else
            mon(c).cell = 2
        End If
    End If
    'm'' little debug control, player may kill a summon that have been swallowed
    If (mon(c).hasinstomach = 0 And mon(c).cell = 2) Then 'm''
        If (plr.instomach <> c) Then
            mon(c).cell = 1
        Else
            mon(c).cell = 2
        End If
    End If
    #End If
Next c



For a = plr.X - drawrange To plr.X + drawrange
'DoEvents

'If isexpansion = 1 Then maxtile = 50 Else maxtile = 25
maxtile = 50


If a < 1 Or a > mapx Then GoTo 8
    For b = plr.Y - drawrange To plr.Y + drawrange
     If b < 1 Or b > mapy Then GoTo 7
'     drawobj tilespr, a, b, map(a, b).tile
     
     
     If Not map(a, b).ovrtile = 0 Then
        If a - plr.X < 3 And a - plr.X > -1 And b - plr.Y < 3 And b - plr.Y > -1 Then
            If map(a, b).ovrtile < maxtile Then drawobj transovrspr, a, b, map(a, b).ovrtile
            If isexpansion = 1 Then If map(a, b).ovrtile > 50 Then drawobj transextraovrs(map(a, b).ovrtile - 50), a, b
            If isexpansion = 0 Then If map(a, b).ovrtile > 25 Then drawobj transextraovrs(map(a, b).ovrtile - 25), a, b
     
        Else: If map(a, b).ovrtile < maxtile Then drawobj ovrspr, a, b, map(a, b).ovrtile
            If isexpansion = 1 Then If map(a, b).ovrtile > 50 Then drawobj extraovrs(map(a, b).ovrtile - 50), a, b
            If isexpansion = 0 Then If map(a, b).ovrtile > 25 Then drawobj extraovrs(map(a, b).ovrtile - 25), a, b
        End If

     End If
     
     If map(a, b).monster > totalmonsters Then GoTo 7
     If map(a, b).monster > 0 Then
'     Stop
     If mon(map(a, b).monster).type > 0 And mon(map(a, b).monster).stomachlevel < 4 Then drawobj mongraphs(mon(map(a, b).monster).type), a, b, mon(map(a, b).monster).cell, 1, mon(map(a, b).monster).xoff, mon(map(a, b).monster).yoff, mon(map(a, b).monster).litx, mon(map(a, b).monster).lity
     If Not mon(map(a, b).monster).X = a Then map(a, b).monster = 0 'Fix for unattackable monsters
     End If
     
     If map(a, b).object > objtotal Then GoTo 7
     If map(a, b).object > 0 Then
         If objtypes(objs(map(a, b).object).type).graphloaded = 0 And Not objtypes(objs(map(a, b).object).type).graphname = "" Then makesprite objgraphs(objs(map(a, b).object).type), Form1.Picture1, objtypes(objs(map(a, b).object).type).graphname, objtypes(objs(map(a, b).object).type).r, objtypes(objs(map(a, b).object).type).g, objtypes(objs(map(a, b).object).type).b, objtypes(objs(map(a, b).object).type).l, objtypes(objs(map(a, b).object).type).cells: objtypes(objs(map(a, b).object).type).graphloaded = 1
         If objtypes(objs(map(a, b).object).type).graphloaded = 1 Then drawobj objgraphs(objs(map(a, b).object).type), a, b, , 1, objs(map(a, b).object).xoff, objs(map(a, b).object).yoff
     End If
     If a = plr.X And b = plr.Y And (plr.instomach = 0 Or stomachlevel <= 1) Then
     'If zex = 0 Then drawobj plrgraphs, plr.X, plr.Y, , , plr.xoff, plr.yoff, , , 1
     drawauras 1
     If zex = 0 Then plrgraphs.TransparentDraw picBuffer, 400 - (plrgraphs.CellWidth / 2) - 25, 300 - plrgraphs.CellWidth - 24 + zex + 50, 1 Else waterspr.TransparentDraw picBuffer, 400 - waterspr.CellWidth / 2, 300, 1, False

     End If
7   Next b
8 Next a

'Drawing Compensator for Text Box
offset = -30
'If Form1.Text7.Visible = True Then offset = -30

'drawbody cStage.hDC, 700, -offset + 10
'plrgraphs.TransparentDraw picbuffer, 0, 0, 1

fullbody.TransparentDraw picBuffer, 600, -offset + 10, 1
'If Not wep.graphname = "" Then wepgraph.TransparentDraw cStage.hDC, 700 + 24 - wep.xoff, -offset + 10 + 104 - wep.yoff, 1 ', newhdc

minimapspr.TransparentDraw picBuffer, 175, -offset - 10, 1

If editon = 1 Then
    tilespr.TransparentDraw picBuffer, 1, 1, edittile: ovrspr.TransparentDraw picBuffer, 96, 50, edittile, False
    If edittile > 25 Then tilespr2.TransparentDraw picBuffer, 1, 1, edittile Mod 25: ovrspr.TransparentDraw picBuffer, 96, 50, edittile, False
End If

20

'offset = 0
'If Form1.Text7.Visible = True Then offset = -50

'cStage.RenderBitmap Form1.hDC, 0, offset
'drawstatbars 150, 10
drawstatbars 290, 480
drawtexts
blt Form1.Picture7

'Form1.Label1(0).Caption = Form1.Label1(0).Caption + 1
'Form1.Label1(0).Refresh
'stilldrawing = stilldrawing - 1
End Sub

Function drawrain()
Static rloaded As Byte
If rloaded = 0 Then Set raingraph = New cSpriteBitmaps: raingraph.CreateFromFile "rain1.bmp", 4, 1, , 0: rloaded = 1: raingraph.recolor 30, 200, 250
For a = 1 To 50
If rain(a).frame = 0 Then rain(a).frame = 1
If rain(a).frame > 4 Then rain(a).Y = 0: rain(a).X = roll(800): rain(a).stopped = 0: rain(a).frame = 1
If rain(a).stopped = 0 Then rain(a).X = rain(a).X + 5: rain(a).Y = rain(a).Y + 10 Else rain(a).frame = rain(a).frame + 1
If rain(a).stopped = 0 Then If roll(50) = 1 Then rain(a).stopped = 1
raingraph.TransparentDraw picBuffer, rain(a).X, rain(a).Y, rain(a).frame
5 Next a

End Function

Function grv(val1, val2)
If val1 > val2 Then grv = val1 Else grv = val2
'If val1 = val2 Then grv = val1 + 1
End Function

Sub makesprite(sp As cSpriteBitmaps, pb As PictureBox, ByVal filen As String, Optional red = 0, Optional green = 0, Optional blue = 0, Optional light = 0.5, Optional cells = 1, Optional ycells = 1, Optional shiny = 0)
'Dim dummypic As StdPicture
'Set dummypic = New StdPicture
'ChDir App.Path
If filen = "" Then MsgBox "Empty string passed to makesprite routine": Exit Sub

If cells = 0 Then cells = 1
Set sp = Nothing
Set sp = New cSpriteBitmaps

#If USELEGACY = 1 Then 'm'' Duam's leftover debug code.
Open "lastfile.txt" For Output As #69
Write #69, "File Name:" & filen
Close #69
#End If 'm''

sp.CreateFromFile filen, cells, ycells, , 0
If red > 255 Then red = 255
If green > 255 Then green = 255
If blue > 255 Then blue = 255
If red > 0 Or green > 0 Or blue > 0 Then sp.recolor red, green, blue, light, , shiny

'pb.AutoRedraw = True
'pb.Picture = LoadPicture(filen)
'BitBlt dummypic, 0, 0, pb.Width, pb.Height, pb.hDC, 0, 0, SRCCOPY
'rangecolor red, green, blue, pb, light
'sp.CreateFromPicture dummypic, cells, ycells, , RGB(0, 0, 0)
'sp.CreateFromPicture pb.Picture, cells, ycells, , RGB(0, 0, 0)
'If cells = 0 Then Stop

#If USELEGACY = 1 Then 'm'' Duam's leftover debug code.
If Not Dir("lastfile.txt") = "" Then Kill "lastfile.txt"
#End If 'm''

End Sub

Sub makedigsprite(sp As cSpriteBitmaps, pb As PictureBox, ByVal filen As String, Optional red, Optional green, Optional blue, Optional light = 0.5, Optional cells = 1, Optional ycells = 1, Optional amt = 20)
randword filen
Set sp = Nothing
Set sp = New cSpriteBitmaps
If IsMissing(red) Then sp.CreateFromFile filen, cells, 1, , RGB(0, 0, 0): Exit Sub

'pb.AutoRedraw = True
makesprite sp, pb, filen, red, green, blue, light, cells, ycells
'pb.Picture = LoadPicture(filen)
'rangecolor red, green, blue, pb, light
sp.digjunk amt, amt
'digjunk pb, 20, 10
'pb.Picture = pb.image
'selfassign pb
'sp.CreateFromPicture pb, cells, ycells, , RGB(0, 0, 0)
End Sub

Sub slowmakesprite(sp As cSpriteBitmaps, pb As PictureBox, filen As String, Optional red, Optional green, Optional blue, Optional light = 0.5, Optional cells = 1, Optional ycells = 1)

MsgBox "Obsolete function slowmakesprite called"

Stop
grloading = 1
Set sp = Nothing
Set sp = New cSpriteBitmaps
If IsMissing(red) Then sp.CreateFromFile filen, cells, 1, , RGB(0, 0, 0): Exit Sub
'DoEvents
'pb.AutoRedraw = True
pb.Picture = LoadPicture(filen)
slowrangecolor red, green, blue, pb, light
DoEvents
sp.CreateFromPicture pb, cells, ycells, , RGB(0, 0, 0)
grloading = 0
End Sub

Sub drawobj(obj As cSpriteBitmaps, X, Y, Optional cell = 1, Optional midtile = 0, Optional ByRef xoff = 0, Optional ByRef yoff = 0, Optional ByRef litx = 0, Optional ByRef lity = 0, Optional nodecoffs = 0)
'positions something correctly according to it's X and Y coordinates

On Error GoTo 15
GoTo 12
15 MsgBox "Error in Drawobj routine": Exit Sub
12

'If Not roll(8) = 1 Then Exit Sub
If cell = 0 Then Exit Sub
dorkx = (X - plr.X - Y + plr.Y) * 48 + 400 - plr.xoff - (obj.CellWidth / 2) + xoff
If dorkx < -120 Or dorkx > 850 Then Exit Sub
dorky = (X + Y - plr.X - plr.Y) * 24 - plr.yoff + 300 - (obj.CellHeight - 48) - (midtile * 24) + yoff
If xoff = 0 And yoff = 0 Then GoTo 5
If nodecoffs > 0 Then GoTo 5

If yoff > 0 Then yoff = yoff - 12
If yoff < 0 Then yoff = yoff + 12
If xoff > 0 Then xoff = xoff - 12
If xoff < 0 Then xoff = xoff + 12

5
If dorky < -140 Or dorky > 660 Then Exit Sub

litx = dorkx
lity = dorky

obj.TransparentDraw picBuffer, dorkx, dorky, cell, False
'obj.DirectBltSprite cStage.hDC, dorkx, dorky, cell


End Sub

Sub getlitxy(ByVal X, ByVal Y, ByRef litx, ByRef lity, obj As cSpriteBitmaps, Optional midtile = 0)
dorkx = (X - plr.X - Y + plr.Y) * 48 + 400 - plr.xoff - (obj.CellWidth / 2) + xoff
'If dorkx < -40 Or dorkx > 750 Then Exit Sub
dorky = (X + Y - plr.X - plr.Y) * 24 - plr.yoff + 300 - (obj.CellHeight - 48) - (midtile * 24) + yoff
litx = dorkx
lity = dorky

End Sub

Sub getlitxy2(ByVal X, ByVal Y, ByRef litx, ByRef lity, Optional midtile = 0)
dorkx = (X - plr.X - Y + plr.Y) * 48 + 400 - plr.xoff + xoff
'If dorkx < -40 Or dorkx > 750 Then Exit Sub
dorky = (X + Y - plr.X - plr.Y) * 24 - plr.yoff + 300 - (midtile * 24) + yoff
litx = dorkx
lity = dorky

End Sub

Sub loaddata(ByVal filen As String, Optional ByVal datfile = "Data.pak")

ChDir App.Path
'If Not getfile(Left(filen, Len(filen) - 4) & ".dat", , , , 1) = "" Then loadbindata Left(filen, Len(filen) - 4) & ".dat": Exit Sub
If Not getfile(Left(filen, Len(filen) - 4) & ".dat", "curgame.dat", , , 1) = "" Then loadbindata Left(filen, Len(filen) - 4) & ".dat": Exit Sub
If Left(filen, 6) = "VTDATA" Then filen = Right(filen, Len(filen) - 6)
plr.curmap = filen
If Not datfile = "" Then origfile = filen: filen = getfile(filen, datfile) 'Use pak file if provided
If Left(origfile, 6) = "VTDATA" Then MsgBox "Error #1 in Loaddata function": Exit Sub

'Load binary data if no text data is available but binary is (For random pre-generated worlds)
If Dir(filen) = "" Then If Not Dir(plr.name & "\" & Left(filen, Len(filen) - 4) & ".dat") Then loadbindata Left(filen, Len(filen) - 4) & ".dat": Exit Sub

Dim segfilename As String

If Right(filen, Len("VRPGData.txt")) = "VRPGData.txt" Then showform10 "Intro", 1: plr.X = 29: plr.Y = 29: filen = getfile(origfile, datfile)
If filen = "Spells.txt" Then GoTo 113
plr.curmap = filen
revamp
loadtreasuretypes
Form1.Text5.Visible = True
Form1.Text5.Refresh
113
Dim fn As String
Open filen For Input As #1
    Do While Not EOF(1)
        Input #1, durg
        
        'If durg = "#MUSIC" Then Input #1, curmusic: playmusic curmusic
        If durg = "#LEVEL" Then Input #1, mapjunk.level
        
        If durg = "#CURMAP" Then Input #1, plr.curmap
        
        If durg = "#CREATEGEMS" Then Input #1, num: For a = 1 To num: creategem 0, 0, 5: Next a
        
        If durg = "#PLANETJUNK" Then Input #1, mapjunk.planetjunk.temperature, mapjunk.planetjunk.moisture, mapjunk.planetjunk.rockyness, mapjunk.planetjunk.vegetation
        
        If durg = "#MONTYPE" Then
            lastmontype = lastmontype + 1
            ReDim Preserve montype(1 To lastmontype): ReDim Preserve mongraphs(1 To UBound(montype())) As cSpriteBitmaps
            Input #1, montype(lastmontype).name, montype(lastmontype).gfile, montype(lastmontype).swallow, montype(lastmontype).hp, montype(lastmontype).skill, montype(lastmontype).dice, montype(lastmontype).damage, montype(lastmontype).eatskill, montype(lastmontype).exp, montype(lastmontype).move, montype(lastmontype).acid
            
            Input #1, r
            If r = "Clan" Then Input #1, clannum: r = clan(clannum).r: g = clan(clannum).g: b = clan(clannum).b: light = clan(clannum).light Else Input #1, g, b, light
            montype(lastmontype).light = light
            calcexp montype(lastmontype)
            
makesprite mongraphs(lastmontype), Form1.Picture1, montype(lastmontype).gfile, r, g, b, light, 3
5         End If

        If durg = "#MONTYPE2" Then
            lastmontype = lastmontype + 1
            Dim lev As String
            ReDim Preserve montype(1 To lastmontype): ReDim Preserve mongraphs(1 To UBound(montype())) As cSpriteBitmaps
            Input #1, montype(lastmontype).name, montype(lastmontype).gfile, lev
            montype(lastmontype).level = lev 'getfromstring(lev, 1)
            bosslev = Val(getfromstring(lev, 4))
            Input #1, r, g, b, light
            montype(lastmontype).color = RGB(r, g, b)
            montype(lastmontype).light = light
            calcexp montype(lastmontype), bosslev

            makesprite mongraphs(lastmontype), Form1.Picture1, montype(lastmontype).gfile, r, g, b, light, 3

        End If
        
        If durg = "#EXTRAOVERS" Then
            Input #1, name1, name2, name3, name4
            loadextraovrs name1, name2, name3, name4
        End If
        
        If durg = "#MONTYPE3" Then
            lastmontype = lastmontype + 1
            ReDim Preserve montype(1 To lastmontype): ReDim Preserve mongraphs(1 To UBound(montype())) As cSpriteBitmaps
            Input #1, montype(lastmontype).name, montype(lastmontype).gfile, montype(lastmontype).level
            Input #1, r, g, b, light
            montype(lastmontype).color = RGB(r, g, b)
            montype(lastmontype).light = light
            calcexp montype(lastmontype), 4

            makesprite mongraphs(lastmontype), Form1.Picture1, montype(lastmontype).gfile, r, g, b, light, 3

        End If

        If durg = "#PLAYER" And plr.hpmax = 0 Then Input #1, plr.name, plr.str, plr.dex, plr.int, plr.hpmax, plr.mpmax, plr.level, plr.exp, plr.expneeded: plr.hp = plr.hpmax: plr.mp = plr.mpmax
    
        If durg = "#OBJTYPE" Then
        efnum = 0: objts = objts + 1: ReDim Preserve objtypes(1 To objts): ReDim Preserve objgraphs(1 To UBound(objtypes())) As cSpriteBitmaps: Input #1, objtypes(objts).name
        End If
        
        If durg = "#CONVGIRL" Then
        efnum = 1: objts = objts + 1: ReDim Preserve objtypes(1 To objts): ReDim Preserve objgraphs(1 To UBound(objtypes())) As cSpriteBitmaps: Input #1, objtypes(objts).name: name = objtypes(objts).name
            Input #1, X, Y, graphic, frames, portrait, mobile, r, g, b
            objtypes(objts).r = r: objtypes(objts).g = g: objtypes(objts).b = b: objtypes(objts).l = 0.5: objtypes(objts).graphname = graphic: objtypes(objts).cells = frames
            addeffect2 objtypes(objts), "Conversation", name, portrait
            If Val(mobile) > 0 Then addeffect2 objtypes(objts), "Mobile", mobile
            createobj objts, X, Y, name
        End If
        
        If durg = "#EFFECT" Then
            efnum = efnum + 1
            If efnum > 10 Then GoTo 8
            For a = 1 To 5
            Input #1, objtypes(objts).effect(efnum, a)
            Next a
8
        End If
        
        If durg = "#SPELL" Then totalspells = totalspells + 1: ReDim Preserve spells(1 To totalspells): Input #1, spells(totalspells).name, spells(totalspells).mp, spells(totalspells).effect, spells(totalspells).amount, spells(totalspells).target, spells(totalspells).school, spells(totalspells).level
        
        If durg = "#GRAPH" Then Input #1, objtypes(objts).graphname, objtypes(objts).cells, objtypes(objts).r, objtypes(objts).g, objtypes(objts).b, objtypes(objts).l ': makesprite objtypes(objts).graph, Form1.Picture1, objtypes(objts).graphname, objtypes(objts).r, objtypes(objts).g, objtypes(objts).b, objtypes(objts).l, objtypes(objts).cells: objtypes(objts).graphloaded = 1

        If durg = "#CREATEOBJ" Then Input #1, name, X, Y, name2, ostr, ostr2: createobj name, X, Y, name2, ostr, ostr2
        
        If durg = "#MULTCREATEOBJ" Then
        Input #1, num, name, X, Y, name2, ostr, ostr2: swaptxt ostr2, "/", ",": swaptxt ostr2, "$", Chr(34)
        For a = 1 To num
        createobj name, X, Y, name2, ostr, ostr2
        Next a
        End If
        
        If durg = "#RANDOMSEED" Then
        Input #1, bug: Rnd (-5)
        If plr.difficulty = 0 Then Randomize bug Else randword plr.name, Val(bug)
        If Val(bug) = 0 Then Randomize GetTickCount()
        End If
        
        If durg = "#MAPSIZE" Then Input #1, mapx, mapy: ReDim map(1 To mapx, 1 To mapy) As tiletype: fillmap 1: updatmap
        If durg = "#MAPCHUNK" Then Input #1, tt, sizs: randomchunk Val(sizs), , , Val(tt)
        If durg = "#OVRCHUNK" Then Input #1, tt, sizs: ovrchunk Val(sizs), , , Val(tt)
            
        If durg = "#BUILDING" Then Input #1, Xs, Ys, X, Y, tt, ot, TY: createbuilding Xs, Ys, TY, 1, tt, ot, X, Y
        If durg = "#BUILDING2" Then Input #1, Xs, Ys, X, Y, tt, ot, TY: createbuilding Xs, Ys, TY, 0, tt, ot, X, Y
        
        If durg = "#LOADSEG" Then Input #1, segfilename, Xs, Ys, rotat, floortype, walltype, salesobj, signobj, bij1, bij2, doorobj, stealobj: loadseg segfilename, Xs, Ys, rotat, floortype, walltype, (salesobj), (signobj), (bij1), (bij2), (doorobj), (stealobj)

        
        If durg = "#MULTBUILDING" Then
        Input #1, num, Xs, Ys, X, Y, tt, ot, TY
        Do While num > 0
        createbuilding Xs, Ys, TY, 1, tt, ot, X, Y
        num = num - 1
        Loop
        End If
        
        If durg = "#RANDOMMONSTERS" Then
        Input #1, mt, num, xz1, yz1, xz2, yz2: randommonsters mt, num, Val(xz1), Val(yz1), Val(xz2), Val(yz2)
        End If
        If durg = "#LOADMAP" Then Input #1, fn: loadmap fn
        If durg = "#SETMAPS" Then Input #1, mapjunk.maps(1), mapjunk.maps(2), mapjunk.maps(3), mapjunk.maps(4)
        If durg = "#TERSTR" Then Input #1, mapjunk.terstr: gamemsg mapjunk.terstr: Form1.Text7.Refresh
        If durg = "#FAVMONSTER" Then Input #1, mapjunk.favmonster, mapjunk.favmonsternum
        If durg = "#QUESTMONSTER" Then Input #1, mapjunk.questmonster
        If durg = "#SPRINKLEOVR" Then Input #1, num, tt: sprinkleovr num, tt
        If durg = "#SPRINKLE" Then Input #1, num, tt: sprinkle num, tt
        If durg = "#FILLMAP" Then Input #1, tt: fillmap tt
        If durg = "#FILLOVR" Then Input #1, ot: fillovr ot
        If durg = "#DUNGEON" Then Input #1, X, Y, x2, y2, tt, ot, ton: createdungeon tt, ot, X, Y, x2, y2, ton: checkaccess2
        If durg = "#CLANCOLOR" Then Input #1, clan(0).r, clan(0).b, clan(0).g, clan(0).light
        
        
        
    Loop
Close #1


If Not filen = "Spells.txt" Then setupquests
If Not filen = "Spells.txt" And Not mapjunk.terstr = "" Then addhasbeento Right(mapjunk.terstr, Len(mapjunk.terstr) - Len("Now Entering")), filen & ":" & Int(mapx / 2) & ":" & Int(mapy / 2)  ': checkaccess

If Not filen = "Spells.txt" And mapx > 0 Then walloffmap


Form1.updatspells
Form1.Text5.Visible = False

'putcarrymonsters

'loadsummons

'Delete VTDatas
killbinfiles

End Sub

Function createmonster(ByVal mont, ByVal X, ByVal Y) As Integer
wm = 1
If Val(mont) = 0 Then
For b = 1 To lastmontype
    If montype(b).name = mont Then mont = b: Exit For
Next b
If b = lastmontype + 1 Then Exit Function
End If

For a = 1 To totalmonsters
    If mon(a).hp <= 0 Then Exit For
Next a
wm = a
If mont > lastmontype Then MsgBox "Error #1 in Createmonster"
If a = totalmonsters + 1 Then ReDim Preserve mon(1 To a) As amonsterT: totalmonsters = totalmonsters + 1
mon(wm).type = mont
mon(wm).hp = montype(mont).hp
map(X, Y).monster = wm
mon(wm).X = X
mon(wm).Y = Y
mon(wm).cell = 1
createmonster = wm
End Function

Function drawbody(hDC As DirectDrawSurface7, X, Y, Optional newhdc As Boolean = False)

Static cleavageloaded As Byte
On Error Resume Next

'plrhairloaded = 0
If plrhairloaded = 0 Then
    getrgb plr.haircolor, rh, gh, bh, 32
    makesprite plrhair, Form1.Picture2, plr.hairname, rh, gh, bh
    plrhairloaded = 1
End If

If plrbodyloaded = 0 Then gbody.CreateFromFile plr.bodyname, 1, 1, , 0: plrbodyloaded = 1

nocleavage = 0
'If Dir("cleavage.bmp") = "" Then nocleavage = 1: GoTo 215 'In case I forget to include cleavage.bmp in the non-expansion thingy

If cleavageloaded = 0 Then Set cleavage = New cSpriteBitmaps: cleavage.CreateFromFile "cleavage.bmp", 1, 1, , 0: cleavageloaded = 1

215 If plr.diglevel < 3 Then gbody.TransparentDraw hDC, X, Y, 1, newhdc

If plr.diglevel > 0 And plr.diglevel <= 5 Then plr.diglevel = greater(plr.diglevel, 1): digbody(plr.diglevel).TransparentDraw hDC, X, Y, 1, newhdc ', newhdc

If plr.diglevel < 3 Then plrhair.TransparentDraw hDC, X + 27, Y + 5, 1, newhdc ', vbSrcCopy

Dim patchg As cSpriteBitmaps
Set patchg = New cSpriteBitmaps: patchg.CreateFromFile "patch1.bmp", 1, 1: getrgb plr.haircolor, rh, gh, bh, 32: patchg.recolor rh, gh, bh: If plr.diglevel < 3 Then patchg.TranslucentDraw hDC, X + 50, Y + 88, 1, , 2

If plr.plrdead > 0 Then nocleavage = 1

For a = 1 To 16
    clothes(a).drawn = 0
    'Capes have wear1 as Jacket, wear2 as Wings
    If clothes(a).wear2 = "Wings" And clothes(a).loaded = 1 Then capegraphs.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster ': clothes(a).drawn = 1
    If clothes(a).wear1 = "Wings" And clothes(a).loaded = 1 Then cgraphs(a).graph.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster: clothes(a).drawn = 1
Next a

For a = 1 To 16
    'clothes(a).drawn = 0
    If clothes(a).wear1 = "Cyberbody" Then nocleavage = 1
    If clothes(a).loaded = 1 And Left(clothes(a).wear1, 5) = "Cyber" Then cgraphs(a).graph.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster: clothes(a).drawn = 1
Next a

'Draw Cleavage
For a = 1 To 16
    If clothes(a).wear1 = "Upper" Then If Not nocleavage = 1 Then cleavage.TransparentDraw hDC, X + 44, Y + 45, 1, newhdc: Exit For
    If clothes(a).wear2 = "Upper" Then If Not nocleavage = 1 Then cleavage.TransparentDraw hDC, X + 44, Y + 45, 1, newhdc: Exit For
    If clothes(a).wear1 = "Bra" Then If Not nocleavage = 1 Then cleavage.TransparentDraw hDC, X + 44, Y + 45, 1, newhdc: Exit For
    If clothes(a).wear2 = "Bra" Then If Not nocleavage = 1 Then cleavage.TransparentDraw hDC, X + 44, Y + 45, 1, newhdc: Exit For
    If clothes(a).wear1 = "Jacket" Then If Not nocleavage = 1 Then cleavage.TransparentDraw hDC, X + 44, Y + 45, 1, newhdc: Exit For
Next a

For a = 1 To 16
    If clothes(a).drawn = 1 Then GoTo 34
    If clothes(a).loaded = 1 And clothes(a).wear1 = "Panties" Then
    If clothes(a).wear2 = "Bra" Then If Not nocleavage = 1 Then cleavage.TransparentDraw hDC, X + 44, Y + 45, 1, newhdc 'y+47
    cgraphs(a).graph.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster: clothes(a).drawn = 1: Exit For
    End If
34 Next a

For a = 1 To 16
If clothes(a).drawn = 1 Then GoTo 7
    If clothes(a).loaded = 1 And clothes(a).wear1 = "Bra" Then
    If Not nocleavage = 1 Then cleavage.TransparentDraw hDC, X + 44, Y + 45, 1, newhdc 'Draw cleavage if a bra is being worn
    cgraphs(a).graph.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster: clothes(a).drawn = 1: Exit For
    End If
7 Next a

For a = 1 To 16
If clothes(a).drawn = 1 Then GoTo 14
    If clothes(a).loaded = 1 And clothes(a).wear1 = "Legs" Then cgraphs(a).graph.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster: clothes(a).drawn = 1: Exit For
14 Next a

For a = 1 To 16
If clothes(a).drawn = 1 Then GoTo 15
    If clothes(a).loaded = 1 And clothes(a).wear1 = "Arms" Then cgraphs(a).graph.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster: clothes(a).drawn = 1: Exit For
15 Next a

For a = 1 To 16
If clothes(a).drawn = 1 Then GoTo 8
    'If clothes(a).loaded = 1 And clothes(a).wear1 = "Lower" Then cgraphs(a).graph.TransparentDraw hDC, x, y, 1, newhdc: clothes(a).drawn = 1: Exit For
    If clothes(a).loaded = 1 And clothes(a).wear1 = "Lower" Then cgraphs(a).graph.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster: clothes(a).drawn = 1: Exit For

8 Next a

For a = 1 To 16
If clothes(a).drawn = 1 Then GoTo 9
    'If clothes(a).loaded = 1 And clothes(a).wear1 = "Upper" Then cgraphs(a).graph.TransparentDraw hDC, x, y, 1, newhdc: clothes(a).drawn = 1: Exit For
    If clothes(a).wear2 = "Bra" Then If Not nocleavage = 1 Then cleavage.TransparentDraw hDC, X + 44, Y + 45, 1, newhdc
    If clothes(a).loaded = 1 And clothes(a).wear1 = "Upper" Then cgraphs(a).graph.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster: clothes(a).drawn = 1: Exit For
    
    'If clothes(a).loaded = 1 And clothes(a).wear1 = "Upper" Then cgraphs(a).graph.DirectBltSprite hDC, x, y, 1, vbSrcPaint: clothes(a).drawn = 1: Exit For

9 Next a

'Draw everything left but the jackets
For a = 1 To 16
If clothes(a).drawn = 1 Then GoTo 11
    If clothes(a).loaded = 1 And Not clothes(a).wear1 = "Jacket" Then cgraphs(a).graph.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster
11 Next a

'Draw Jackets
For a = 1 To 16
If clothes(a).drawn = 1 Then GoTo 10
    If clothes(a).loaded = 1 And clothes(a).wear1 = "Jacket" Then cgraphs(a).graph.TransparentDraw hDC, X, Y, 1, newhdc, clothes(a).raster: clothes(a).drawn = 1: Exit For
    'If clothes(a).loaded = 1 And clothes(a).wear1 = "Cape" Then cgraphs(a).graph.TransparentDraw hDC, x, y, 1, newhdc, clothes(a).raster: clothes(a).drawn = 1: Exit For
10 Next a

If Not wep.graphname = "" And Not plr.plrdead = 1 Then wepgraph.TransparentDraw hDC, X + 24 - wep.xoff, Y + 104 - wep.yoff, 1, newhdc

End Function

Function digwep()
If wep.graphname = "" Then Exit Function
If geteff(wep.obj, "Digested", 2) = "" Then Exit Function
wep.digged = geteff(wep.obj, "Digested", 2)
'Form1.Picture1.Picture = LoadPicture(wep.graphname)
'lwepgraph wep.graphname, Form1.Picture1, 1
'wepgraph.TransparentDraw Form1.Picture1.hDC, X + 24 - wep.xoff, Y + 104 - wep.yoff, 1, newhdc: Exit Function
randword wep.obj.name

'digjunk Form1.Picture1, wep.digged * 2

'Form1.Picture1.Picture = Form1.Picture1.image
'selfassign Form1.Picture1
wepgraph.digjunk wep.digged * 2
'wepgraph.CreateFromPicture Form1.Picture1, 1, 1, , 0

End Function

Function selfassign(pc As PictureBox)
pc.Picture = pc.image
'BitBlt pc.hDC, 0, 0, pc.Width, pc.Height, pc.hDC, 0, 0, SRCCOPY

End Function

Function ifwearcape() As Boolean
'ifwearcape = False
'For a = 1 To 8
'    If clothes(a).wear1 = "Jacket" Then
'Next a
End Function

Function addclothes(name As String, filen As String, armor As Byte, wear1 As String, Optional wear2 As String = "NONE", Optional r = 256, Optional g = 0, Optional b = 0, Optional l = 0.5, Optional makeobj = 0, Optional raster As RasterOpConstants = vbSrcCopy)

If plr.Class = "Naga" Then If wear1 = "Panties" Or wear1 = "Lower" Or wear2 = "Panties" Or wear2 = "Lower" Then Exit Function 'Nagas can't wear anything below the belt...
'm'' the sub calling addclothes (takeobj) include now the Naga class handling code, to avoid this violent exit


If wear2 = "" Then wear2 = "NONE"
If r = 256 Then r = clan(0).r: g = clan(0).g: b = clan(0).b: l = clan(0).light

'Remove any clothes that take the same slot
For a = 1 To 16
    If clothes(a).loaded = 0 Then GoTo 5
    If wear1 = "ALL" Or wear2 = "ALL" Or clothes(a).wear1 = "ALL" Or clothes(a).wear2 = "ALL" Then GoTo 3
    If clothes(a).wear1 = wear1 Or clothes(a).wear1 = wear2 Or clothes(a).wear2 = wear1 Or clothes(a).wear2 = wear2 Then
    If wear2 = "NONE" And Not clothes(a).wear1 = wear1 And Not clothes(a).wear2 = wear1 Then GoTo 5
3   clothes(a).wear1 = ""
    clothes(a).wear2 = ""
    clothes(a).loaded = 0
    clothes(a).armor = 0
    clothes(a).name = ""
    clothes(a).weight = 0
    'If clothes(a).digested = 0 Then
    clothes(a).digested = 0
    getitem clothes(a).obj
    Set cgraphs(a).graph = Nothing
    End If
5 Next a

'Then add clothes to the first available slot
For a = 1 To 16
    If clothes(a).name = "" Then
    
    'clothes(a).raster = 1
    
    clothes(a).wear1 = wear1
    'clothes(a).raster = raster
    If wear2 = "NONE" Then wear2 = ""
    clothes(a).wear2 = wear2
    clothes(a).loaded = 1
    clothes(a).name = name
    clothes(a).armor = armor
    clothes(a).digested = 0
    Set cgraphs(a).graph = New cSpriteBitmaps
    If clothes(a).wear2 = "Wings" Then makesprite capegraphs, Form1.Picture1, "back" & filen, r, g, b, l 'If clothes(a).wear2 = "Cape" Then makesprite backclothesgraph, Form1.Picture1, filen, r, g, b, l
    makesprite cgraphs(a).graph, Form1.Picture1, filen, r, g, b, l
    If Not Dir("D-" & filen) = "" Then
    makesprite cgraphs(a).diggraph, Form1.Picture1, "D-" & filen, r, g, b, l
    Else:
    amt = 3
    If clothes(a).wear1 = "Belt" Or clothes(a).wear1 = "Panties" Or clothes(a).wear1 = "Bra" Then amt = 1
    If Not clothes(a).wear2 = "" Then amt = 5: If clothes(a).wear1 = "Bra" Then amt = 3
    
    makedigsprite cgraphs(a).diggraph, Form1.Picture1, filen, r, g, b, l
    
    'Set clothes(a).diggraph = clothes(a).graph
    
    End If
    clothes(a).loaded = 1
    addclothes = a
    If makeobj = 1 Then
        F = createobjtype(name, filen, r, g, b, l, 1)
        addeffect F, "Clothes", filen, armor, wear1, wear2
        objtypes(F).r = r: objtypes(F).g = g: objtypes(F).b = b: objtypes(F).l = l
'        g = createobj(name, 0, 0, name)
        clothes(a).obj = objtypes(F)
    End If
'    If Not IsMissing(obj) Then clothes(a).obj = obj
    'checkclothes a
    Exit Function
    End If
Next a

End Function

Sub plrattack(ByVal target, Optional ByVal noturn = 0, Optional ByVal extraattacks = 0, Optional ByVal attackdifferents = 1, Optional damagemult = 1)
If target = 0 Or mon(target).owner > 0 And Not Ccom = "EAT" Then Exit Sub
If mon(target).hp < 1 Then Exit Sub
If mon(target).instomach = -1 Then Exit Sub
If diff(plr.X, mon(target).X) > 1 Or diff(plr.Y, mon(target).Y) > 1 Then If Not wep.type = "Bow" Then Exit Sub
If Ccom = "EAT" And mon(target).owner = 1 Then mon(target).owner = 0

Dim monz As monstertype

If mon(target).type = 0 Then Exit Sub
monz = montype(mon(target).type)

'Experience for attack attempts, regardless of whether player hits
plr.exp = plr.exp + Int(montype(mon(target).type).exp / 10)

makebijdang mon(target).X, mon(target).Y, 1

'player wins
aroll = roll(getdex + 2)

'Sword fatigue modifier--swords use less fatigue
If wep.type = "Sword" Then fatmult = 0.6 Else fatmult = 0.8
If wep.type = "Fast" Then fatmult = 0.7

stdfatigue = (3 + greater((wep.weight - getstr(1)), 0)) * fatmult
addfatigue stdfatigue
'addfatigue (3 + wep.weight / greater(getstr(1) + 3, 4)) * fatmult

'Huntress attack roll bonus
If plr.Class = "Huntress" Then aroll = aroll + (2 * (1 + plr.level / 10))
skillmod aroll, "Accuracy", 2, 1

If Ccom = "EAT" Then
    Ccom = ""
    If skilltotal("Giant Stomach", 1, 1) = 0 Then playsound "full5.wav": Exit Sub
    If skilltotal("Giant Stomach", 1, 1) <= plr.foodinbelly Then gamemsg "You don't have enough room in your stomach to eat that!": randsound "full", 5: GoTo 214
    aroll = succroll(greater(getdex, getstr) + (plr.level / 4) + 2, 6 + monz.boss) + skilltotal("Gluttony", 2, 1)
    mroll = succroll(monz.skill + (mon(target).hp / monz.hp) * 10, 5 - monz.boss)
    If aroll > mroll Then
        gamemsg "You have swallowed the " & monz.name & " whole!"
        mon(target).instomach = -1
        plr.foodinbelly = plr.foodinbelly + 1
        plr.monsinbelly = plr.monsinbelly + 1
        playsound "swallow1.wav"
        Form1.updatbody
        map(mon(target).X, mon(target).Y).monster = 0
        movemon target, plr.X, plr.Y, 1
        
        
        Else:
        playsound "failed1.wav"
        gamemsg "The " & monz.name & " resists your attempts to swallow it!"
    End If
214    If noturn = 0 Then turnthing
    Exit Sub
End If

If usingskill = "Stun" And mon(target).stunned = 0 Then
    If succroll(plr.level, 10 - skilltotal("Stun", 2, 1)) + skilltotal("Stun", 1, 1) > succroll(montype(mon(target).type).level, 5) Then
      If spendsp(5) = True Then
        mon(target).stunned = roll(skilltotal("Stun", 3, 2))
        makebijdang mon(target).X, mon(target).Y, 4
        playsound "thunk.wav"
        End If
    End If
End If

'Mace stun
If wep.type = "Mace" And mon(target).stunned = 0 Then
    If roll(4) = 1 Then
        mon(target).stunned = roll(3 + skilltotal("Stun", 2, 1)) + 2
        makebijdang mon(target).X, mon(target).Y, 4
        playsound "thunk.wav"
    End If
End If


If usingskill = "Frenzy" And Not extraattacks = -1 Then
extraattacks = extraattacks + getplrskill("Frenzy")
'If spendsp(extraattacks * 3) = False Then extraattacks = 0
End If

If wep.type = "Fast" Then extraattacks = extraattacks + 1: freebie = freebie + 1

If extraattacks > 0 Then
    For a = 1 To extraattacks
    If attackdifferents = 1 Then If nearmon(1, target) = 0 Then Exit For
    If freebie = 0 Then If spendsp(3) = False Then Exit For
    If attackdifferents = 0 Then plrattack target, 1, 0, , damagemult * 0.5: freebie = freebie - 1
    If attackdifferents = 1 Then If nearmon(1, target) > 0 Then plrattack nearmon(1, target), 1, -1, , damagemult * 0.5: freebie = freebie - 1: addfatigue stdfatigue
    Next a
End If

'Check after multiple attacks
If mon(target).hp < 1 Then Exit Sub

If aroll >= roll(monz.skill) Then
dmg = Int(getplrdamage * damagemult)

'Damage bonuses for Spear and Axe types
If wep.type = "Spear" Then dmg = dmg * (1 + (getdex / 40))
If wep.type = "Axe" Then dmg = dmg * (1 + (getstr / 40))

'20% bonus for weak against weapons
If wep.dice > 0 Then If wep.type = monz.weaktype Then dmg = Int(dmg * 1.2)

If wep.dice > 0 Then playsound "clang" & roll(5) & ".wav" Else playsound "punch" & roll(2) & ".wav"
'MsgBox "You hit " & monz.name & " for " & dmg & " damage."
If usingskill = "Cripple" Then
If mon(target).hp = montype(mon(target).type).hp Then If spendsp(skilltotal("Cripple", 8, 1)) = True Then skillpercminus mon(target).hp, "Cripple", 25, 4
End If

If usingskill = "Charged Strike" Then
If spendsp(skilltotal("Charged Strike", 6, 3)) = True Then skillpercmod dmg, "Charged Strike", 50, 25
End If

If usingskill = "Power Strike" Then
If spendsp(skilltotal("Power Strike", 2, 1)) = True Then skillmod dmg, "Power Strike", getstr * wep.dice, getstr * wep.dice
End If

If usingskill = "Vital Strike" Then
If spendsp(skilltotal("Vital Strike", 2, 1)) = True Then skillmod dmg, "Vital Strike", getdex * wep.dice, getdex * wep.dice
End If

If usingskill = "Cunning Strike" Then
If spendsp(skilltotal("Cunning Strike", 2, 1)) = True Then skillmod dmg, "Cunning Strike", getint * wep.dice, getint * wep.dice
End If

If HasSkill("Vicious") And Val(monz.level) >= plr.level - 1 Then
If spendsp(skilltotal("Vicious", 4, 2)) = True Then
zark = skilltotal("Vicious", 2, 1) * (Val(monz.level) - plr.level)
dmg = dmg * (1 + zark / 5)
End If
End If

damagemon target, dmg

If mon(target).hp > 0 Then dispatk monz, mon(target).hp Else Form1.Picture5.Visible = False: Form1.Text6.Visible = False

'Life Drain
If plr.Class = "Succubus" Then plrdamage -dmg * 0.02
drainhp = skilltotal2("Drain Life", 1, 0.5)
If drainhp > 0 Then plrdamage -dmg * drainhp / 100, 0

'Magic Drain
drainmp = skilltotal2("Drain Magic", 1, 0.5)
If drainmp > 0 Then losemp -dmg * drainmp / 100

'If mon(target).hp < 1 Then


'enemy wins
Else:
'MsgBox "You attack and miss."
playsound "SWING.wav"
dispatk monz, mon(target).hp
End If


If noturn = 0 Then turnthing

End Sub

Sub damagemon(ByVal target, ByVal damage, Optional ByVal expmult As Single = 1)
mon(target).hp = mon(target).hp - damage

getlitxy mon(target).X, mon(target).Y, damx, damy, mongraphs(mon(target).type)

addtext Int(damage), damx + roll(30) - 15, damy + roll(30) - 15, 255, 10, 3

'Experience per hit -- You get more experience the longer it takes you to kill them.  5 hits will give you full experience value, since attacking alone gives 5% and hitting gives another 5%.  Another incentive to fight tough monsters.
plr.exp = plr.exp + Int(montype(mon(target).type).exp / 10) 'Int((Int(montype(mon(target).type).exp * (Val(montype(mon(target).type).level) + 2 + plr.difficulty) / plr.level) * expmult) / 4)

If mon(target).hp < 1 Then
    If mon(target).type = 0 Then killmon target, 1: Exit Sub
    'plr.exp = plr.exp + Int(montype(mon(target).type).exp * (Val(montype(mon(target).type).level) + 2 + plr.difficulty) / plr.level) * expmult
    ' +/- 10% per level difference, minumum 10% experience
    plr.exp = plr.exp + Int(montype(mon(target).type).exp * greater((1 + (Val(montype(mon(target).type).level) - plr.level) / 10), 0.3))
    addtext Int(montype(mon(target).type).exp * greater((1 + (Val(montype(mon(target).type).level) - plr.level) / 20), 0.3)) & " XP", damx, damy - 50, 100, 100, 250
    
    
    'Swallowtime increases by 1 each time a monster is killed
    plr.Swallowtime = plr.Swallowtime + 1 + plr.difficulty
    
'    MsgBox "You have killed " & montype(mon(target).type).name
    If roll(5) = 1 Then  'Drop Items
        'MsgBox "You gain " & Int(montype(mon(target).type).exp / 2) & " gold!"
        aroll = roll(107) - 10
        If aroll < 90 Then
        broll = roll(3)
        Select Case broll
        Case 1: createobj "Large Gold Pile", mon(target).X, mon(target).Y, "Gold", Int(montype(mon(target).type).exp)
        Case 2: createobj "Medium Gold Pile", mon(target).X, mon(target).Y, "Gold", Int(montype(mon(target).type).exp / 2)
        Case 3: createobj "Small Gold Pile", mon(target).X, mon(target).Y, "Gold", Int(montype(mon(target).type).exp / 4)
        End Select
        End If
        
        'plr.gp = plr.gp + Int(montype(mon(target).type).exp / 2)
        If aroll >= 90 And aroll < 93 Then createclothes Val(montype(mon(target).type).level), mon(target).X, mon(target).Y
        If aroll >= 93 And aroll < 95 Then createclothes Val(montype(mon(target).type).level), mon(target).X, mon(target).Y, "Armor"
        If aroll >= 95 And aroll < 96 Then createclothes Val(montype(mon(target).type).level), mon(target).X, mon(target).Y, "Weapons"
        If aroll >= 96 Then creategem mon(target).X, mon(target).Y, Val(montype(mon(target).type).level)
        
        
        
        End If
    killmon target
End If
End Sub

Sub killmon(ByVal target, Optional ByVal killthemotherfucker = 0)
'movemon target, 0, 0
If target = plr.instomach Then stomachlevel = 0
If target = plr.instomach And plr.plrdead = 1 Then plr.plrdead = 0: plr.hp = 1

If mon(target).instomach > 0 Then mon(mon(target).instomach).hasinstomach = 0
mon(target).instomach = 0
mon(target).stomachlevel = 0

'm'' trying to comment duamutef code :
'm'' if the picture of my monster is a big belly, let's check what he have in its stomach.
'm'' and let's move to his coordinate whatever is in its stomach
If mon(target).cell = 2 Then
For a = 1 To totalmonsters
    If mon(a).instomach = target Then movemon a, mon(target).X, mon(target).Y
Next a
End If

'If killthemotherfucker > 0 Then mon(target).X = 0: mon(target).Y = 0: mon(target).type = 0: mon(target).hp = 0: mon(target).owner = 0: Exit Sub
If mon(target).type = 0 Then GoTo 3
If montype(mon(target).type).name = "Thirsha" Then wingame
If montype(mon(target).type).name = mapjunk.questmonster Then ifquest montype(mon(target).type).name, 1
If plr.instomach = target Then plr.instomach = 0 'Get player out of stomach if that monster is killed
map(mon(target).X, mon(target).Y).monster = 0
3 If killthemotherfucker = 0 Then makebijdang mon(target).X, mon(target).Y, 2
If killthemotherfucker = 0 Then playsound "monsterdie.wav"
mon(target).X = 0
mon(target).Y = 0
mon(target).type = 0
mon(target).owner = 0
End Sub

Function getplrdamage(Optional usedex = 0, Optional missileweapon = 0) As Integer

'Add stuff as objects and things come in
'dmg = rolldice(getstr, 6) + rolldice(Int(plr.level / 3), (plr.level / 6) + 4) + (plr.level * 2)
If usedex = 0 Then dmg = rolldice(getstr, 6) + rolldice(plr.level, 4) _
Else dmg = rolldice(getdex, 6) + rolldice(plr.level, 4)


'If isexpansion = 1 Then dmg = roll(getstr) + 2

dmg = dmg + getbonus("BONUSDAMAGE")

wepdice = wep.dice + getbonus("BONDICE")
wepdamage = wep.damage + getbonus("BONDAMAGE")

'No weapon damage if you're melee attacking with a missile weapon
If wep.type = "Bow" And missileweapon = 0 Then GoTo 6
dmg = dmg + rolldice(wepdice, wepdamage)
6

dmg = dmg * (1 + plr.str / 10)

If usingskill = "Perfect Strike" Then If spendsp(5) Then dmg = ((getstr * 6) + Int((plr.level / 3) * (plr.level / 6 + 4)) + (plr.level * 2) + (wep.dice * wep.damage)) * (1 + getplrskill("Perfect Strike") / 5)

If wep.type = "Sword" Or wep.type = "Small" Then dmg = dmg + wepdice * skilltotal("Sword Mastery", 2, 2)
If wep.type = "Spear" Then dmg = dmg + wepdice * skilltotal("Spear Mastery", 3, 3)
If wep.type = "Axe" Then dmg = dmg + wepdice * skilltotal("Axe Mastery", 3, 3)
'If wep.type = "Bow" Then skillpercmod dmg, "Bow Mastery", 20, 20


If getplrskill("Weapons Mastery") Then dmg = dmg + rolldice(wepdice, wepdamage) * (skilltotal("Weapons Mastery", 20, 6) / 100)

'Amazon damage bonus
If plr.Class = "Amazon" Then dmg = dmg * (1.1)
skillpercmod dmg, "Deathblow", 10, 5
If getplrskill("Critical Strike") > 0 Then If roll(100) < 10 + getplrskill("Critical Strike") * 3 Then dmg = dmg * 2

If wep.dice = 0 Then skillpercmod dmg, "Streetfighting", 30, 15

getplrdamage = dmg
End Function

Function rolldice(ByVal dice, ByVal damage)
If Val(dice) = 0 Then rolldice = 0: Exit Function
If Val(damage) = 0 Then damage = 1
dice = Int(dice)
damage = Int(damage)

total = 0
Do While dice > 0
total = total + roll(damage)
dice = dice - 1
Loop

rolldice = total

End Function

Function movemon(ByVal target, ByVal X, ByVal Y, Optional skipthing = 0)

If mon(target).X = 0 Then GoTo 5
If mon(target).Y = 0 Then GoTo 5
'm''If x > 0 And y > o Then If map(x, y).monster > 0 Then Exit Function 'm'' Duam wrote o instead of 0
If X > 0 And Y > 0 Then If map(X, Y).monster > 0 Then Exit Function 'm''
If skipthing = 0 And map(mon(target).X, mon(target).Y).monster = target Then map(mon(target).X, mon(target).Y).monster = 0
5 If X <= 0 Or X > mapx Or Y <= 0 Or Y > mapy Then Exit Function
'If skipthing = 1 And mon(target).x > 0 And mon(target).y > 0 Then map(mon(target).x, mon(target).y).monster = 0
mon(target).X = X
mon(target).Y = Y
If X > 0 Then map(X, Y).monster = target

End Function

Sub loadmap(filen)

If Dir(filen) = "" Then filen = getfile(filen, "Data.pak")

Open filen For Input As #2

Do While Not EOF(2)
Input #2, durg

    If durg = "#MAPSIZE" Then Input #2, mapx, mapy: ReDim map(1 To mapx, 1 To mapy) As tiletype
    
    If durg = "#TILESONLY" Then
    For a = 1 To mapx
    For b = 1 To mapy
    Input #2, map(a, b).tile
    Input #2, map(a, b).ovrtile
    Input #2, map(a, b).blocked
    Next b
    Next a
    End If
    
    If durg = "#TILESET" Then Input #2, fil2, fil3: tilespr.CreateFromFile fil2, 5, 5, , RGB(0, 0, 0): ovrspr.CreateFromFile fil3, 5, 10, , RGB(0, 0, 0)
    
    If durg = "#MAPINFO" Then
    For a = 1 To mapx
    For b = 1 To mapy
    Input #2, map(a, b).tile
    Input #2, map(a, b).ovrtile
    Input #2, ass, ass
    'Input #2, map(a, b).monster
    'Input #2, map(a, b).object
    Input #2, map(a, b).blocked
    Next b
    Next a
    End If
Loop
Close #2
updatmap

End Sub


Sub savemap(filen)
'm'' declarations...
Dim a As Long, b As Long

Open filen For Output As #1

Write #1, "#MAPSIZE", mapx, mapy
Write #1, "#TILESONLY"
    For a = 1 To mapx
    For b = 1 To mapy
    Write #1, map(a, b).tile
    Write #1, map(a, b).ovrtile
    Write #1, map(a, b).blocked
    Next b
    Next a

Close #1
End Sub

Sub monmove(ByVal target, ByVal xplus, ByVal yplus, Optional skipthing = 0)
If (mon(target).X + xplus) < 1 Or (mon(target).X + xplus) > mapx Then Exit Sub
If (mon(target).Y + yplus) < 1 Or (mon(target).Y + yplus) > mapy Then Exit Sub
If map(mon(target).X + xplus, mon(target).Y + yplus).blocked = 1 Then Exit Sub
If map(mon(target).X + xplus, mon(target).Y + yplus).monster > 0 Then Exit Sub
movemon target, mon(target).X + xplus, mon(target).Y + yplus, skipthing

'plr.xoff = getxoff(xplus, yplus)
'plr.yoff = getyoff(xplus, yplus)

'mon(target).xoff = getxoff(xplus, yplus)
'mon(target).yoff = getyoff(xplus, yplus)
'mon(target).xoff = xplus * -36
'mon(target).yoff = yplus * -24

'mon(target).xoff = (xplus - yplus) * -36
mon(target).yoff = ((xplus + yplus) / 2) * -36
mon(target).xoff = (xplus - ((xplus + yplus) / 2)) * -48

End Sub

Sub turnthing(Optional notmoving = 0)

Static hasntmoved

If map(plr.X, plr.Y).blocked = 1 Then map(plr.X, plr.Y).blocked = 0: map(plr.X, plr.Y).ovrtile = 0: checkaccess2

If monsterstoput > 0 Then putcarrymonsters

If notmoving = 0 Or stomachlevel > 0 Then hasntmoved = 0 Else hasntmoved = hasntmoved + 1

If plr.plrdead = 0 And plrhpgain > 0 Then plog = (lesser(Int(plr.level) + 5, plrhpgain)): plrdamage -plog, 0: plrhpgain = plrhpgain - plog: If plrhpgain < 0 Then plrhpgain = 0
If plr.hp >= gethpmax Then plrhpgain = 0

If editon = 1 Then Exit Sub
turnswitch = 0
turncount = turncount + 1: If turncount > 100 Then turncount = 1

'Swallowcounter is increased every 10 turns
If turncount Mod 10 = 0 Then plr.Swallowtime = plr.Swallowtime + 1
If turncount Mod 4 = 0 And swallowcounter < 0 Then swallowcounter = swallowcounter + 1


monai
movobjs
'If swallowcounter < 0 Then swallowcounter = swallowcounter + 1

'And turncount Mod 3 = 0
If plr.instomach > 0 And plr.plrdead = 0 And roll(3) = 1 And struggled = 1 Then playsound "grunt" & roll(9) + 2 & ".wav"

'If plr.instomach > 0 Then plrdamage montype(mon(plr.instomach).type).acid, 1

'If plr.instomach > 0 And turncount Mod 10 = 0 Then 'swallowcounter = swallowcounter + 1
If plr.instomach And stomachlevel >= 3 Then plrdamage Int(montype(mon(plr.instomach).type).acid / (3 - struggled)), , 1

If plr.instomach > 0 Then swallowcounter = swallowcounter + 1 + struggled * 2: instomachcounter = instomachcounter + 1 + struggled * getplrskill("squirm") + 1 Else instomachcounter = instomachcounter - 2: If instomachcounter < 0 Then instomachcounter = 0
If plr.instomach > 0 And swallowcounter >= 10 Then
'    swallowcounter = 0    Msgbox
    'playsound "gurgle" & roll(2) & ".wav"
    digestplr
    If plr.instomach = 1 Then swallowcounter = 4

End If

struggled = 0

'If turncount Mod 50 = 0 Then
decboosts 'Boosts wear off
If turncount Mod 3 = 0 And plr.regen > 0 Then plrdamage -plr.regen '-rolldice(plr.regen, 6)
'Regeneration
If Not plr.instomach > 0 And Not plr.hp <= 1 Then
    If plr.Class = "Enchantress" And turncount Mod 8 = 0 Then plrdamage -1 * plr.level, 0
End If

If getplrskill("Regeneration") > 0 Then
    If Not plr.instomach > 0 And Not plr.hp <= 1 Then
    If skilltotal("Regeneration", 2, 1) >= 15 Then plrdamage Int(-1 * Sqr(plr.level)), 0 Else If turncount Mod (15 - skilltotal("Regeneration", 2, 1)) = 0 Then plrdamage -1 * Sqr(plr.level), 0
    End If
End If
If plr.Class = "Angel" And turncount Mod 6 = 0 Then plr.mp = plr.mp + 2 * (1 + (plr.level / 10))


zarb = Int(getint / 2) - 9: If zarb < 1 Then zarb = 1
zark = 15 - Int(getint / 2): If zark < 1 Then zark = 1
If turncount Mod zark = 0 Then plr.mp = plr.mp + zarb: If plr.mp > getmpmax Then plr.mp = getmpmax
zarb = skilltotal("Mana Regeneration", 2, 1): If zarb < 1 Then zarb = 1
zark = 10 - skilltotal("Mana Regeneration", 2, 1): If zark < 1 Then zark = 1
If turncount Mod zark = 0 Or hasntmoved > 8 Then plr.mp = plr.mp + zarb: If plr.mp > getmpmax Then plr.mp = getmpmax


If plr.plrdead = 0 And turncount Mod 98 / (3 + skilltotal("Super Acid", 1, 1)) = 0 Then digestfood

If plr.exp >= plr.expneeded Then gainlevel

If plr.plrdead = 0 And HasSkill("Alchemy") And plr.gp >= 2 And plr.mp < getmpmax / 2 Then If spendsp(1) Then plr.mp = plr.mp + skilltotal("Alchemy", 3, 2): plr.gp = plr.gp - 2

'If HasSkill("Poisonous") And turncount Mod 4 = 0 And plr.instomach > 0 Then spendsp 1

If turncount Mod 5 = 0 Then plr.sp = plr.sp + 1: subfatigue Sqr(getend) + 2: If Ccom = "EAT" Then Ccom = ""
If notmoving = 1 And hasntmoved > 8 Then subfatigue (Sqr(getend) + 5) * Int(hasntmoved / 8): losemp -Int(hasntmoved / 8): spendsp -1
If hasntmoved > 25 Then spendsp -2
If hasntmoved > 50 Then spendsp -10
'If notmoving = 1 And hasntmoved > 25 Then subfatigue Sqr(getend) + 5

updathp

updatbonuses = 0

If turncount Mod 3 = 0 Then updatbonuses = 1

'If turncount = 50 Then makeminimap
If turncount Mod 10 = 0 Then updateminimap

End Sub

Sub digestplr()

If montype(mon(plr.instomach).type).eattype = 1 Then stomachlevel = 5: GoTo 4 'Engulfing things act differently
If plr.plrdead = 0 Then stomachlevel = stomachlevel + 1 Else GoTo 4
If stomachlevel = 1 Then addtext "The " & montype(mon(plr.instomach).type).name & " is trying to swallow you whole!", 250
If stomachlevel = 1 Then addtext "You must escape from it!", , 215

If stomachlevel = 2 Then addtext "You are sliding down the " & montype(mon(plr.instomach).type).name & "'s throat!", 275: playsound "swallow6.wav"
If stomachlevel = 2 Then addtext "You must escape quickly, before you reach it's stomach!", 230, 215
If stomachlevel <= 2 Then swallowcounter = 0: Exit Sub
If stomachlevel = 3 Then addtext "You have slithered into the " & montype(mon(plr.instomach).type).name & "'s stomach!", 250: playsound "swallow5.wav"

If stomachlevel > 3 Then plr.Swallowtime = greater(0, plr.Swallowtime - 10)

'Slight randomization in time
If stomachlevel = 5 And roll(2) = 1 Then stomachlevel = stomachlevel + roll(4) - 2
If stomachlevel = 11 And roll(4) = 1 Then stomachlevel = 10
If stomachlevel = 11 And roll(4) = 1 Then stomachlevel = 12


If stomachlevel = 9 Then addtext "You have been forced down into the " & montype(mon(plr.instomach).type).name & "'s intestinal tract!", 250: playsound "swallow4.wav"

If stomachlevel = 12 Then addtext "You are nearing the end of the " & montype(mon(plr.instomach).type).name & "'s digestive tract!", 280, , , , , 60
If stomachlevel = 12 Then addtext "It will have no choice but to shit you out soon!", , 215, , , , 60
If stomachlevel = 13 Then addtext "The " & montype(mon(plr.instomach).type).name & " has shit you out!"
If stomachlevel = 13 Then createobjtype "Shit", "poo1small.bmp": createobj "Shit", plr.X, plr.Y: mon(plr.instomach).cell = 1: plrescape "The " & montype(mon(plr.instomach).type).name & " has shit you out!": playsound "fart57.wav": Exit Sub

4   playsound "stomach" & roll(4) + 1 & ".wav"
    playsound "swallow" & roll(4) + 3 & ".wav"
    'playsound "blurble" & roll(3) & ".wav"
    
    redl = roll(150) + 100
    greenl = redl - roll(redl / 2)
    addtext getgener("*gurgle*", "*blorp*", "*groan*"), 400 + roll(30) - 15, 300 + roll(80) - 40, redl, greenl, 3

    If plr.plrdead = 1 Then
    If roll(2) = 1 Then gamemsg getdig(1)
    If isnaked = True Then plr.diglevel = plr.diglevel + 1: playsound "burp" & roll(5) & ".wav": Form1.updatbody 'If plr.diglevel > 3 Then plr.diglevel = 3: Form1.updatbody
    If isnaked = False Then digestclothes montype(mon(plr.instomach).type).acid, 1
    End If
    
    If plr.plrdead = 0 Then
        dmg = rolldice(montype(mon(plr.instomach).type).acid, 6 + plr.difficulty)
        plrdamage dmg, 0.3
        If stomachlevel > 4 Then gamemsg getdig(1)
        digestclothes (montype(mon(plr.instomach).type).acid)
        If HasSkill("Poisonous") And spendsp(10) Then damagemon plr.instomach, dmg * getplrskill("Poisonous")
        If plr.hp <= 1 Then
        If isnaked = True Then
            plr.diglevel = plr.diglevel + 1: playsound "burp" & roll(5) & ".wav": Form1.updatbody 'If plr.diglevel > 3 Then plr.diglevel = 3: Form1.updatbody
        Else:
            If roll(2) = 1 Or plr.diglevel >= 2 Then digestclothes 255, 1 Else plr.diglevel = plr.diglevel + 1: playsound "burp" & roll(5) & ".wav": Form1.updatbody
        End If
        End If
        #If USELEGACY = 1 Then 'm''
        If plr.diglevel > 3 Then plr.plrdead = 1 'm'' original duam's line
        #Else 'm''
        'm'' if the player is digested with stuff in his belly, it will be released inside the pred
        If plr.diglevel > 3 Then plr.plrdead = 1 'm''
            'TO DO
        #End If 'm''
    End If
    

    swallowcounter = 0
    If plr.diglevel >= 6 Then
        Do While isnaked = False
            digestclothes 255, 1
        Loop
        createobjtype "Shit", "poo1small.bmp": createobj "Shit", plr.X, plr.Y: mon(plr.instomach).cell = 1: mon(plr.instomach).instomach = 0: plr.instomach = 0: stopsounds: playsound "fart57.wav": wep.dice = 0: wep.graphname = "": Set wepgraph = Nothing: plr.plrdead = 1
        gamemsg "You have been utterly digested."
   End If

End Sub

Sub monai()
Dim a As Long 'm'' declare...

For a = 1 To totalmonsters
    If mon(a).type = 0 Then GoTo 5
    If mon(a).stunned > 0 Then mon(a).stunned = mon(a).stunned - 1: GoTo 5
    If mon(a).instomach > 0 Then If mon(mon(a).instomach).hp < 0 Then mon(a).instomach = 0
    
    #If USELEGACY = 1 Then
    #Else
    'm'' monsters eating each others. Duam suppressed this code. I fixed it
    If mon(a).instomach > 0 Then 'm''
        movemon a, 0, 0 'm''
        mondigest a: GoTo 5 'm''
    End If 'm''
    #End If
    
    If mon(a).instomach = -1 Then
        plrdigmon mon(a), a
        GoTo 5
    End If
        

        
    If mon(a).X > 0 Then map(mon(a).X, mon(a).Y).monster = a
    If mon(a).owner > 0 Then friendai mon(a), a: GoTo 5
    
    'm'' this line below prevents monster with a full belly to attack the player.
#If USELEGACY = 1 Then
    If plr.instomach = 0 And plr.plrdead = 0 And Not mon(a).cell = 2 And diff(plr.X, mon(a).X) <= 1 And diff(plr.Y, mon(a).Y) <= 1 And mon(a).instomach = 0 Then monatk a: GoTo 5
#Else
    'm'' because it would be a exploit to "fill" a boss, here is a tweak
    'm'' "if i'm the boss and wherever i've a full belly or not, i'll attack the player"
    If plr.instomach = 0 And plr.plrdead = 0 And diff(plr.X, mon(a).X) <= 1 And diff(plr.Y, mon(a).Y) <= 1 And mon(a).instomach = 0 Then
        If Not mon(a).cell = 2 Then
            monatk a
            GoTo 5
        Else
            If montype(mon(a).type).name = mapjunk.questmonster Then
                monatk a:
                GoTo 5
            End If
        End If
    End If
#End If
    
    
    If roll(6) < montype(mon(a).type).move Or allmove > 0 Then
    
    'm'' movement AI of Duam here : if monster less than 7 cell away from player, move to player.
    'm''
    If montype(mon(a).type).move > 2 Then
        If ((diff(plr.X, mon(a).X) > 7 Or diff(plr.Y, mon(a).Y) > 7) Or (plr.instomach > 0) Or (roll(4) = 1 And allmove = 0)) Then
            monmove a, roll(3) - 2, roll(3) - 2
        Else
            monmove a, posneg(plr.X, mon(a).X), posneg(plr.Y, mon(a).Y)
        End If
    End If
    End If
    If Not montype(mon(a).type).missileatk = "" And plr.instomach = 0 And diff(plr.X, mon(a).X) < 7 And diff(plr.Y, mon(a).Y) < 7 Then
        atk = montype(mon(a).type).missileatk
        If roll(getfromstring(atk, 2)) = 1 Then
        
        'zx = mon(a).x: zy = mon(a).y
        'zx = zx - plr.x + 9: zy = zy - plr.y + 6
        'revgetXY zx, zy
        zx = mon(a).litx + mongraphs(mon(a).type).CellHeight / 2: zy = mon(a).lity + mongraphs(mon(a).type).CellWidth / 2
        'zx = mon(a).litx: zy = mon(a).lity
           shootat 400, 300, getfromstring(atk, 1), "ENEMY:" & getfromstring(atk, 3) & ":" & getfromstring(atk, 4), zx, zy
        mon(a).animdelay = -roll(3)
        End If
    End If
    
'    End If
5 Next a

allmove = 0

End Sub

Sub mondigest(ByVal monnum As Long)  'm'' declare

If mon(mon(monnum).instomach).type = 0 Then Exit Sub

If roll(8) < 7 Then Exit Sub

'monster attempts escape
mon(monnum).stomachlevel = mon(monnum).stomachlevel + 1

If mon(monnum).stomachlevel < 4 Then
    'm'' formula to get out
    If succroll(montype(mon(monnum).type).skill) > succroll(montype(mon(mon(monnum).instomach).type).skill * 2) Then
        mon(monnum).X = mon(mon(monnum).instomach).X: mon(monnum).Y = mon(mon(monnum).instomach).Y
        'm'' loop to seek a place to pop out of the stomach
3         monmove monnum, roll(3) - 2, roll(3) - 2, 1
          If mon(monnum).X = mon(mon(monnum).instomach).X And mon(monnum).Y = mon(mon(monnum).instomach).Y Then GoTo 3
        'm'' mon(mon(monnum).instomach).cell = 1: mon(monnum).instomach = 0 'm'' too early reset
        mon(monnum).stomachlevel = 0
        mon(mon(monnum).instomach).hasinstomach = 0
        mon(mon(monnum).instomach).cell = 1 'm''
        mon(monnum).instomach = 0 'm'' right time to reset
        playsound "burp" & roll(5) & ".wav"
        Exit Sub
        
    End If
End If

#If USELEGACY = 1 Then
If mon(monnum).stomachlevel = 4 Then playsound "swallow6.wav"
'm'' fixed cell rendered depending on stomach level
#Else
If mon(monnum).stomachlevel = 4 Then 'm''
    playsound "swallow6.wav" 'm''
    mon(mon(monnum).instomach).cell = 2 'm''
End If 'm''
#End If

mon(monnum).hp = mon(monnum).hp - rolldice(montype(mon(mon(monnum).instomach).type).acid, 6)

If mon(monnum).hp <= 0 Then

If plr.instomach = monnum Then plr.instomach = mon(monnum).instomach
mon(mon(monnum).instomach).hasinstomach = 0
createobj "Shit", mon(mon(monnum).instomach).X, mon(mon(monnum).instomach).Y: mon(mon(monnum).instomach).cell = 1: mon(monnum).instomach = 0: playsound "fart57.wav"
'map(mon(monnum).x, mon(monnum).y).monster = 0
mon(monnum).stomachlevel = 0
mon(monnum).type = 0
movemon monnum, 0, 0
'mon(monnum).x = 0
'mon(monnum).y = 0

Else:

'squirm if at more than 30% HP, just for effect
If mon(monnum).hp > montype(mon(monnum).type).hp * 0.3 Then
mon(mon(monnum).instomach).xoff = roll(12)
mon(mon(monnum).instomach).yoff = roll(12)
End If

End If


End Sub

Sub friendai(ByRef wmon As amonsterT, ByVal monnum)

t = montype(mon(monnum).type).name

If diff(plr.X, wmon.X) > 15 Or diff(plr.Y, wmon.Y) > 15 Then
'Let friendly monsters teleport to the player, provided they aren't in a stomach and are mobile to begin with
If Not montype(wmon.type).move = 0 Then
    For a = 1 To totalmonsters
        If mon(a).instomach = monnum Then Exit For
    Next a
If a = totalmonsters + 1 Then movemon monnum, plr.X, plr.Y
End If
End If

If diff(plr.X, wmon.X) > 6 Or diff(plr.Y, wmon.Y) > 6 And Not montype(wmon.type).move = 0 Then monmove monnum, posneg(plr.X, mon(monnum).X), posneg(plr.Y, mon(monnum).Y) Else 'Keep within six spaces of the player

'Attack if possible, otherwise move
If wmon.cell = 2 And Not montype(wmon.type).move = 0 Then monmove monnum, roll(3) - 2, roll(3) - 2


For a = mon(monnum).X - 1 To mon(monnum).X + 1
For b = mon(monnum).Y - 1 To mon(monnum).Y + 1

    If a > mapx Or a < 1 Then GoTo 5
    If b > mapy Or b < 1 Then GoTo 5

    If a = mon(monnum).X And b = mon(monnum).Y Then GoTo 5
    If map(a, b).monster > 0 Then If mon(map(a, b).monster).owner = 0 Then targ = map(a, b).monster: Exit For
5 Next b
Next a



'If mon(monnum).Cell = 2 Then monmove monnum, roll(3) - 2, roll(3) - 2: Exit Sub
If targ > 0 Then monatk2 wmon, mon(targ), monnum, targ: Exit Sub

'Else move towards one, or move randomly
If mon(monnum).cell = 2 Then Exit Sub
zerf = findmon(monnum, 0)
If zerf = 0 And Not montype(wmon.type).move = 0 Then monmove monnum, roll(3) - 2, roll(3) - 2

'Use Missile Attack
If zerf >= 1 And Not montype(mon(monnum).type).missileatk = "" Then
        atk = montype(mon(monnum).type).missileatk
        If roll(getfromstring(atk, 2)) = 1 Then
        zx = mon(monnum).litx + mongraphs(mon(monnum).type).CellHeight / 2: zy = mon(monnum).lity + mongraphs(mon(monnum).type).CellWidth / 2
        shootat mon(zerf).litx, mon(zerf).lity, getfromstring(atk, 1), getfromstring(atk, 3) & ":" & getfromstring(atk, 4), zx, zy
        mon(monnum).animdelay = -roll(3)
        End If
End If


End Sub

Function monatk2(ByRef mon1 As amonsterT, ByRef mon2 As amonsterT, ByVal monnum1, ByVal monnum2, Optional ByVal nocounterattack = 0)
'When 2 monsters attack one another
If mon2.type = 0 Or mon1.type = 0 Then Exit Function

#If USELEGACY <> 1 Then
'm'' for some reason, this code lack handling when one monster swallow the other.
'm'' the "prey" is never in the "pred"
'm'' so this is a fix :)
If mon1.instomach > 0 Then Exit Function 'm''
#End If

playsound montype(mon1.type).Sound '  "attack1.wav"
playsound montype(mon2.type).Sound
If Not mon(monnum1).cell = 2 Then mon(monnum1).animdelay = -roll(3)
If Not mon(monnum2).cell = 2 Then mon(monnum2).animdelay = -roll(3)



getlitxy mon1.X, mon1.Y, damx, damy, mongraphs(mon1.type)

If succroll(montype(mon1.type).skill + 20) >= succroll(montype(mon2.type).skill + 20) Then

    #If USELEGACY = 1 Then
    If mon1.cell = 2 Or mon1.hasinstomach > 0 Then GoTo 5 'cannot attack when already full
    #Else
    'm'' make boss able to fight even with a full belly, but unable to swallow more (avoiding party wipe!)
    If mon1.cell = 2 Or mon1.hasinstomach > 0 Then 'm''
        If montype(mon1.type).name = mapjunk.questmonster Then 'm''
            damagemon monnum2, rolldice(montype(mon1.type).dice, montype(mon1.type).damage), 0.5 'm''
            GoTo 5 'm'' party wipe avoid
        Else 'm''
            GoTo 5 'm''
        End If 'm''
    End If 'm''
    #End If 'm''
    
    'mon1.cell = 1
    'mon1.hasinstomach = 0
    'mon2.hp = mon2.hp - rolldice(montype(mon1.type).dice, montype(mon1.type).damage)
    'If mon2.hp <= 0 Then killmon monnum2: GoTo 5
    damagemon monnum2, rolldice(montype(mon1.type).dice, montype(mon1.type).damage), 0.5
    If mon2.type = 0 Then Exit Function
    If Not mon1.owner = 1 And succroll(montype(mon1.type).eatskill) > succroll(montype(mon2.type).skill) Then
        mon2.instomach = monnum1
        mon2.stomachlevel = 1
        mon1.hasinstomach = monnum2
        mon1.cell = 2
        mon(monnum1).animdelay = 0
        playsound "swallow1.wav"
        addtext getgener("*GULP*", "*SLURP*", "*CHOMP*", "*GULP*"), damx, damy - 25, 250, 250, 10
        mon2.xoff = 0: mon2.yoff = 0
        #If USELEGACY <> 1 Then
            'm'' little game message to say one of your pet being swallowed!
            If mon1.owner = 0 Then 'm''
                gamemsg "Your " & montype(mon2.type).name & " have been swallowed by a " & montype(mon1.type).name & " !" 'm''
            End If 'm''
        #End If
    End If
5

Else:
If nocounterattack = 0 Then monatk2 mon2, mon1, monnum2, monnum1, 1
'    If mon2.cell = 2 Or mon1.type = 0 Or mon2.type = 0 Then GoTo 7 'cannot fight back when already full
'    damagemon monnum1, rolldice(montype(mon2.type).dice, montype(mon2.type).damage), 0.5
'    If mon1.type = 0 Or mon2.type = 0 Then Exit Function
'    If Not mon2.owner = 1 And succroll(montype(mon2.type).eatskill) > succroll(montype(mon1.type).skill) And Not mon1.cell = 2 Then
'        mon1.instomach = monnum2
'        mon2.cell = 2
'        mon(monnum2).animdelay = 0
'        playsound "swallow1.wav"
'        addtext getgener("*GULP*", "*SLURP*", "*CHOMP*", "*GULP*"), damx, damy - 25, 250, 250, 10
'        mon1.xoff = 0: mon1.yoff = 0
'    End If
7
End If


End Function

Function findmon(monnum, Optional nomove = 0)

For c = 1 To 5
'Go after monsters close to the monster
'For a = mon(monnum).x - c To mon(monnum).x + c
'For b = mon(monnum).y - c To mon(monnum).y + c

'Go after monsters close to the player
For a = plr.X - c To plr.X + c
For b = plr.Y - c To plr.Y + c

If a > mapx Or a < 1 Then GoTo 5
If b > mapy Or b < 1 Then GoTo 5

If a = mon(monnum).X And b = mon(monnum).Y Then GoTo 5
If map(a, b).monster > 0 Then If mon(map(a, b).monster).owner = 0 Then targ = map(a, b).monster: Exit For

5 Next b
If targ > 0 Then Exit For
Next a
If targ > 0 Then Exit For
Next c

If targ = 0 Then Exit Function
findmon = targ

If nomove = 0 Then monmove monnum, posneg(mon(targ).X, mon(monnum).X), posneg(mon(targ).Y, mon(monnum).Y)

End Function

Function nearmon(ByVal distance, Optional ByVal ignorethismon = 0)
'Returns enemy monster near player
nearmon = 0
For a = plr.X - distance To plr.X + distance
For b = plr.Y - distance To plr.Y + distance
    If a > mapx Or a < 1 Then GoTo 5
    If b > mapy Or b < 1 Then GoTo 5

    If map(a, b).monster > 0 Then If mon(map(a, b).monster).owner = 0 And Not map(a, b).monster = ignorethismon Then nearmon = map(a, b).monster: Exit Function
    

5 Next b
Next a

End Function

Function posneg(v1, v2)
posneg = 0
If v1 - v2 < 0 Then posneg = -1
If v1 - v2 > 0 Then posneg = 1
End Function

Function posit(v1)
If v1 < 1 Then posit = -v1 Else posit = v1
End Function

Sub plrmove(ByVal xplus As Long, ByVal yplus As Long) 'm'' declare
If plr.X + xplus < 1 And Not mapjunk.maps(4) = "" Then
fn = mapjunk.maps(4)
monlev = checklevels(fn)
If monlev > plr.level Then If MsgBox("The lowest level monster in that area is level " & monlev & ".  You are level " & plr.level & ".  Meaning you really don't want to go there.  Do you want to be a 'tard and go there anyway?", vbYesNo) = vbNo Then Exit Sub
gotomap mapjunk.maps(4): plr.X = mapx: plr.Y = Int(mapy / 2): Exit Sub
End If

If plr.Y + yplus < 1 And Not mapjunk.maps(1) = "" Then
fn = mapjunk.maps(1)
monlev = checklevels(fn)
If monlev > plr.level Then If MsgBox("The lowest level monster in that area is level " & monlev & ".  You are level " & plr.level & ".  Meaning you really don't want to go there.  Do you want to be a 'tard and go there anyway?", vbYesNo) = vbNo Then Exit Sub
gotomap mapjunk.maps(1): plr.X = Int(mapx / 2): plr.Y = mapy: Exit Sub
End If

If plr.X + xplus > mapx And Not mapjunk.maps(2) = "" Then
fn = mapjunk.maps(2)
monlev = checklevels(fn)
If monlev > plr.level Then If MsgBox("The lowest level monster in that area is level " & monlev & ".  You are level " & plr.level & ".  Meaning you really don't want to go there.  Do you want to be a 'tard and go there anyway?", vbYesNo) = vbNo Then Exit Sub
gotomap mapjunk.maps(2): plr.X = 1: plr.Y = Int(mapy / 2): Exit Sub
End If

If plr.Y + yplus > mapy And Not mapjunk.maps(3) = "" Then
fn = mapjunk.maps(3)
monlev = checklevels(fn)
If monlev > plr.level Then If MsgBox("The lowest level monster in that area is level " & monlev & ".  You are level " & plr.level & ".  Meaning you really don't want to go there.  Do you want to be a 'tard and go there anyway?", vbYesNo) = vbNo Then Exit Sub
gotomap mapjunk.maps(3): plr.X = Int(mapx / 2): plr.Y = 1: Exit Sub
End If

If plr.X + xplus < 1 Or plr.X + xplus > mapx Or plr.Y + yplus < 1 Or plr.Y + yplus > mapy Then Exit Sub
If map(plr.X + xplus, plr.Y + yplus).blocked = 1 Then playsound "oof.wav": Exit Sub
If map(plr.X + xplus, plr.Y + yplus).monster >= 1 Then If mon(map(plr.X + xplus, plr.Y + yplus).monster).owner = 0 Or Ccom = "EAT" Then plrattack map(plr.X + xplus, plr.Y + yplus).monster: Exit Sub
If takeobj(plr.X + xplus, plr.Y + yplus, 0, dummyobj) Then xplus = 0: yplus = 0

plr.X = plr.X + xplus: plr.Y = plr.Y + yplus

plr.xoff = getxoff(xplus, yplus)
plr.yoff = getyoff(xplus, yplus)
'plr.xoff = xplus * -24
'plr.yoff = yplus * -24


turnthing

End Sub

Function getxoff(movX, movy)
getxoff = ((movX * 48) - (movy * 48)) * -1
End Function

Function getyoff(movX, movy)
getyoff = (movX * 24 + movy * 24) * -1
End Function

Function diff(ByVal v1, ByVal v2)
'If v1 < 0 Then v1 = -v1
'If v2 < 0 Then v2 = -v2
diff = Abs(v1 - v2) 'm'' absolute
'm'' v1 = v1 - v2
'm'' If v1 < 0 Then v1 = -v1
'm'' diff = v1
End Function

Sub monatk(monnum)
Dim monz As monstertype
If mon(monnum).type = 0 Then Exit Sub

'Smack monster for damage if fire aura is on
If skilltotal("Fire Aura", 1, 1) > 0 Then
    stotal = skilltotal("Fire Aura", 5, 3)
    damagemon monnum, stotal ^ 2
    If mon(monnum).hp <= 0 Then Exit Sub
End If

If Not map(mon(monnum).X, mon(monnum).Y).monster = monnum Then map(mon(monnum).X, mon(monnum).Y).monster = monnum

getlitxy mon(monnum).X, mon(monnum).Y, damx, damy, mongraphs(mon(monnum).type)

monz = montype(mon(monnum).type)
'playsound "attack1.wav"
playsound monz.Sound
If mon(monnum).instomach = 0 Then mon(monnum).animdelay = -roll(3)
broll = roll(monz.skill + 12)

'Priestess evasion bonus
If plr.Class = "Priestess" Then broll = broll - 2 * (1 + plr.level / 10)

skillminus broll, "Evasion", 2, 1

'Dodge skill
If roll(100) < skilltotal("Dodge", 5, 2) Then addtext "DODGE", 400, 205, 200, 200, 250: Exit Sub

If broll > roll(getdex + 8) Then

If wep.type = "Staff" Then If roll(4) = 1 Then addtext "BLOCK", 400, 205, 250, 200, 250: Exit Sub ': damage = damage * 0.5

dmg = rolldice(monz.dice, monz.damage)
'MsgBox monz.name & " attacks you and you take " & plrdamage(dmg) & " damage."
plrdamage dmg
'playsound "grunt" & roll(9) + 2 & ".wav"

'Swallowtime increases each time a monster attacks the player
plr.Swallowtime = plr.Swallowtime + 1

If swallowcounter < -3 Then Exit Sub

aroll = succroll(monz.eatskill * (2 + plr.Swallowtime \ 10), 5)

'Angel eating penalty
If plr.Class = "Angel" Then aroll = aroll + 2

If plr.hp < 1 Then GoTo 15
If aroll > succroll(getdex * ((plr.hp / gethpmax) + 1) * 2, 5) Or roll(20) = 1 Or plr.hp < 1 Then
    If instomachcounter <= 0 And roll(4) = 1 And swallowcounter >= 0 Or plr.hp < 1 And plr.diglevel < 1 Then
15      'gamemsg getswallow(monnum)
        plr.instomach = monnum
        If montype(mon(plr.instomach).type).eattype = 1 Then stomachlevel = 5
        plr.xoff = getxoff(mon(monnum).X - plr.X, mon(monnum).Y - plr.Y)
        plr.yoff = getyoff(mon(monnum).X - plr.X, mon(monnum).Y - plr.Y)
        
        mon(monnum).cell = 3
        mon(monnum).animdelay = 4
        playsound "swallow1.wav"
        addtext getgener("*GULP*", "*SLURP*", "*CHOMP*", "*GULP*"), damx, damy - 25, 250, 200, 30
        Form1.updatbody
        swallowcounter = 6
        stomachlevel = 0
    End If
End If

Else:
'MsgBox monz.name & " attacks you and misses"

End If


End Sub

Function plrdamage(ByVal damage, Optional ByVal permmult = 0.25, Optional noarmor = 0)
damage = Int(damage)
If plr.plrdead > 0 Then Exit Function
If damage < 0 Then GoTo 5
'If plr.hp < 1 Then Exit Function

If isexpansion = 0 Then permmult = 0

qdamage = Int(damage / 4)
If noarmor = 1 Then GoTo 3
carmor = clothesarmor + plr.armorboost
skillpercmod carmor, "Defence", 15, 5
damage = damage - Sqr(carmor) '- succroll(carmor, 8)
'damage = damage - plr.armorboost
skillmod damage, "Resilience", -2, -1
3
'Valkyrie armor bonus
If plr.instomach = 0 And plr.Class = "Valkyrie" Then damage = damage - clothesarmor * (0.1 + (plr.level / 100))
'If plr.instomach = 0 Then

If HasSkill("Block") And damage > gethpmax / 10 Then
olddamage = damage
If spendsp(3) Then skillpercminus damage, "Block", 25, 5: gamemsg "You block " & Int(olddamage - damage) & " points of damage!"
End If

If damage < 1 Then damage = 1: GoTo 5

If HasSkill("Mana Shield") And plr.mp > damage / 2 And ((damage > gethpmax / 10) Or plr.hp < gethpmax / 2) Then
    If spendsp(2) = True Then
    'damage = Int(damage / 2)
    mandamage = damage
    For a = 1 To getplrskill("Mana Shield")
    mandamage = mandamage * 0.9
    Next a
    mandamage = Int(mandamage)
    gamemsg "Your mana shield absorbs " & mandamage & " of the damage!"
    plr.mp = plr.mp - Int(mandamage)
    damage = damage - mandamage
    End If
End If

'If damage < qdamage Then damage = qdamage '(Damage will never be more than quartered)

If HasSkill("Damage Energy") Then If spendsp(1) Then plr.mp = plr.mp + (damage * (skilltotal("Damage Energy", 20, 10) / 100))

If damage < 1 Then damage = 1


5 damage = Int(damage)
'If plr.hp <= 0 And damage < 0 Then damage = -1
plr.hplost = plr.hplost + Int(damage * permmult): If plr.hplost < 0 Then plr.hplost = 0
plrdamage = damage
plr.hp = plr.hp - damage

If plr.hp < 0 Then plr.hp = 0

If damage > 0 Then addtext damage, 390 + roll(20), 225, 255, 230, 3 Else addtext -damage, 400, 205, 5, 250, 3

If damage = -10000 Then plr.hplost = 0: plr.fatigue = 0
If plr.hp > gethpmax Then plr.hp = gethpmax
If plr.hp = gethpmax And plr.diglevel > 0 Then plr.diglevel = plr.diglevel - 1: Form1.updatbody
If plr.hp < 0 Then plr.hp = 0
'updathp

'If plr.hp < 1 And plr.instomach = 0 Then MsgBox "You finally succumb to the harsh blows of your opponent." & vbCrLf & "YOU ARE DEAD."
'If plr.hp < 1 And plr.instomach > 0 Then MsgBox "As you mercilessly churn inside " & montype(mon(plr.instomach).type).name & "'s hot depths, your body finally begins to give way under the digestive onslaught." & vbCrLf & "YOU HAVE BEEN UTTERLY DIGESTED."

End Function

Function succroll(ByVal dice, Optional ByVal diff = 6)
tot = 0
dice = Val(dice)
Do While dice > 0
If roll(10) >= diff Then tot = tot + 1
dice = dice - 1
Loop
succroll = tot

End Function

Function randomchunk(ByVal sizeish, Optional ByVal X = 0, Optional ByVal Y = 0, Optional ByVal tiletype = 1)
'Static totalrunning As Long
'Debug.Print "Random Chunk " & x & ", " & y & " Total:" & totalrunning
'totalrunning = totalrunning + 1
sizeish = sizeish * sizeish

2 If X = 0 Then X = roll(mapx)
If Y = 0 Then Y = roll(mapy)
3 If X < 1 Or X > mapx Or Y < 1 Or Y > mapy Then X = 0: Y = 0: GoTo 2
If map(X, Y).tile = tiletype And tries < 100 Then X = X + roll(3) - 2: Y = Y + roll(3) - 2: tries = tries + 1: GoTo 3
If sizeish = 0 Then GoTo 5
If tries >= 100 Then Exit Function
map(X, Y).tile = tiletype: sizeish = sizeish - 1: tries = 0
'Lava and water are impassable
If tiletype = 8 Or tiletype = 18 Then map(X, Y).blocked = 1 Else map(X, Y).blocked = 0
GoTo 3
5

End Function

Function ovrchunk(ByVal sizeish, Optional ByVal X = 0, Optional ByVal Y = 0, Optional ByVal tiletype = 1)
'Static totalrunning As Long
'Debug.Print "Random Chunk " & x & ", " & y & " Total:" & totalrunning
'totalrunning = totalrunning + 1
sizeish = sizeish * sizeish
tries = 0

2 If X = 0 Then X = roll(mapx)
If Y = 0 Then Y = roll(mapy)
3 If X < 1 Or X > mapx Or Y < 1 Or Y > mapy Then X = 0: Y = 0: GoTo 2
If tries > 250 Then Exit Function
If map(X, Y).ovrtile = tiletype Or map(X, Y).blocked = 1 Then tries = tries + 1: X = X + roll(3) - 2: Y = Y + roll(3) - 2: GoTo 3
If sizeish <= 0 Then GoTo 5

map(X, Y).ovrtile = tiletype: sizeish = sizeish - 1: map(X, Y).blocked = 1: If tiletype = 5 Or tiletype = 17 Or tiletype = 20 Then map(X, Y).blocked = 0
GoTo 3
5

End Function

Function createobj(ByVal typename, ByVal X, ByVal Y, Optional ByVal name = "", Optional str = "", Optional str2 = "")
If typename = "" Then Exit Function
objtotal = objtotal + 1
ReDim Preserve objs(1 To objtotal) As aobject

If X > 0 Then GoTo 6
3 X = roll(mapx)
Y = roll(mapy)
If map(X, Y).blocked > 0 Or map(X, Y).object > 0 Then GoTo 3
6 If X = 0 Or Y = 0 Or X > mapx Or Y > mapy Then GoTo 3
If map(X, Y).blocked > 0 Or map(X, Y).object > 0 Then X = X + roll(3) - 2: Y = Y + roll(3) - 2: GoTo 6

If Val(typename) > 0 Then objs(objtotal).type = typename: GoTo 8
For a = 1 To objts
If objtypes(a).name = typename Then objs(objtotal).type = a: Exit For
Next a
8
If a = objts + 1 Then objtotal = objtotal - 1: Exit Function
If name = "" And Val(typename) = 0 Then name = typename
objs(objtotal).name = name
objs(objtotal).string = str
objs(objtotal).string2 = str2
objs(objtotal).X = X
objs(objtotal).Y = Y
map(X, Y).object = objtotal
createobj = objtotal
End Function

Sub updatmap()
'm'' declarations
Dim a As Long, b As Long

'totalmonsters = 0
'ReDim mon(1 To 1) As amonsterT
'objtotal = 0
'ReDim objs(1 To 1) As aobject
For a = 1 To mapx
    For b = 1 To mapy
        If map(a, b).ovrtile > 0 Then map(a, b).blocked = 1
        If map(a, b).ovrtile = 5 Then map(a, b).blocked = 0
        If map(a, b).ovrtile = 17 Then map(a, b).blocked = 0
        If map(a, b).ovrtile = 20 Then map(a, b).blocked = 0
    Next b
Next a


End Sub

Function findclothesmatch(wear1, Optional wear2 = "NONE")
'returns total armor for wear
total = 0
For a = 1 To 16
    If clothes(a).wear1 = wear1 Or clothes(a).wear1 = wear2 Then total = total + clothes(a).armor: GoTo 5
    If Not clothes(a).wear2 = "" Then If clothes(a).wear2 = wear1 Or clothes(a).wear2 = wear2 Then total = total + clothes(a).armor: GoTo 5
5 Next a
findclothesmatch = total

End Function

Function clothesarmor()
For a = 1 To 16
    total = total + clothes(a).armor
Next a
clothesarmor = total + getbonus("BONARMOR")
End Function

Function digestclothes(ByVal acid, Optional ByVal forcedig = 0)

updatbonuses = 1

Dim backgraph As cSpriteBitmaps

'Weapon first
    If roll(6) = 1 And Not wep.graphname = "" Then
    diglev = greater(Int(acid / 4), 1)
    If isexpansion = 1 Then wep.damage = wep.damage - 1 Else wep.dice = wep.dice - diglev
    wep.digged = wep.digged + diglev
    changeeff2 wep.obj, "Digested", wep.digged
    changeeff2 wep.obj, "Weapon", , wep.dice, wep.damage
    If wep.damage > 0 And wep.dice > 0 Then lwepgraph wep.graphname, Form1.Picture1
    If wep.damage <= 0 Or wep.dice <= 0 Then
        Dim snarg As objecttype
        'MsgBox "The " & montype(mon(plr.instomach).type).name & "'s stomach wrenches your " & wep.obj.name & " out of your hands! You grasp for it but it slithers down the beast's intestine..."
        gamemsg "Your " & wep.obj.name & " has been dissolved!"
        wep.obj = snarg
        wep.graphname = ""
        wep.dice = 0
        wep.damage = 0
        wep.weight = 0
        wep.type = ""
        End If
    Form1.updatbody
    End If


a = topclothes(roll(3) - 1)
If a = 17 Then Exit Function
    alreadydigested = 0
    If clothes(a).digested = 1 Then alreadydigested = 1
    'If roll(clothes(a).armor) < roll(acid) Or roll(6) = 1 Then
    If roll(4) = 1 Or forcedig > 0 Then
        If roll(100) < Val(geteff(clothes(a).obj, "BONUNDIG", 2)) Then GoTo 5
        clothes(a).digested = 1: clothes(a).armor = clothes(a).armor - greater(Int(clothes(a).armor / 2), Int(acid / 2))
        If clothes(a).armor < 1 And alreadydigested = 0 Then clothes(a).armor = 1
        For b = 1 To 6
            If clothes(a).obj.effect(b, 1) = "Clothes" Then clothes(a).obj.effect(b, 3) = clothes(a).armor
        Next b
        
        If alreadydigested = 0 Then Set backgraph = cgraphs(a).graph: Set cgraphs(a).graph = cgraphs(a).diggraph: Set cgraphs(a).diggraph = backgraph: changeeff2 clothes(a).obj, "DIGESTED", "1": clothes(a).obj.name = "Partially Digested " & clothes(a).obj.name: clothes(a).digested = 1
        Form1.updatbody
        playsound "burp" & roll(5) & ".wav"
        If clothes(a).armor <= 1 And alreadydigested = 1 Then
        Dim snarg2 As objecttype
        digeas = getgener("digested!", "dissolved!", "absorbed!")
        If plr.plrdead = 0 Then gamemsg "Your " & clothes(a).name & " has been " & digeas
        clothes(a).loaded = 0: clothes(a).armor = 0: clothes(a).name = "": clothes(a).wear1 = "": clothes(a).wear2 = "": clothes(a).obj = snarg2: Form1.updatbody
        End If
5     End If


Set backgraph = Nothing

updatbonuses = 1

End Function

Function topclothes(Optional skipdig = 0)
4 durg = 13
5 durg = durg - 1
If durg = 12 Then smurg = "Jacket"
If durg = 11 Then smurg = "Belt"
If durg = 10 Then smurg = "Arms"
If durg = 9 Then smurg = "Legs"
If durg = 8 Then smurg = "Upper"
If durg = 7 Then smurg = "Lower"
If durg = 6 Then smurg = "Bra"
If durg = 5 Then smurg = "Panties"
If durg = 4 Then smurg = "CyberArms"
If durg = 3 Then smurg = "CyberLegs"
If durg = 2 Then smurg = "CyberBody"

For a = 1 To 16
    If skipdig >= 1 And clothes(a).digested = 1 And durg > 0 Then GoTo 5
    If clothes(a).wear1 = smurg Then topclothes = a: Exit Function
    If durg = 1 And skipdig = 0 Then If clothes(a).loaded = 1 Then topclothes = a: Exit Function
Next a
If durg > 0 Then GoTo 5
If skipdig >= 1 Then skipdig = 0: GoTo 4
topclothes = a

End Function

Function takeobj(X, Y, frominv, dobjtype As objecttype)
'True means kick the player back to the square she was in
'If x is -1, then it will take the object from inventory slot #y
'if dobjtype.name = "Mercenary Girl2" Then Stop

updatbonuses = 1

Dim dobjt As objecttype
takeobj = False
On Error Resume Next
If X = -1 And Y > 0 Then dobjt = inv(Y): dobj = Y: invnum = Y: GoTo 3
If X = 0 And Y = 0 And frominv = 0 Then dobjt = dobjtype

If X = 0 And Y = 0 Then GoTo 2
If map(X, Y).object = 0 Then Exit Function

dobj = map(X, Y).object
dobjt = objtypes(objs(dobj).type)
2 If Not dobjtype.name = "" Then dobjt = dobjtype: If X = 0 Then dobj = dobjt.name
3
If X > 0 And Y > 0 And Ccom = "EAT" Then If eatobj(dobj) = False Then takeobj = True: Exit Function Else Exit Function

For a = 1 To 6
    If dobjt.effect(a, 1) = "GEM" Then Ccom = "GEM:" & dobj: gamemsg "Double-click on the item you want to add the gem to."
    If dobjt.effect(a, 1) = "Pickup" Then If frominv = 0 Then playsound dobjt.effect(a, 2): getitem dobjt: If X > 0 Or Y > 0 Then map(X, Y).object = 0: Exit Function Else Exit Function
    If dobjt.effect(a, 1) = "Backup" Then takeobj = True
    If dobjt.effect(a, 1) = "Givestr" Then plr.str = plr.str + dobjt.effect(a, 2)
    If dobjt.effect(a, 1) = "Givedex" Then plr.dex = plr.dex + dobjt.effect(a, 2)
    If dobjt.effect(a, 1) = "Giveint" Then plr.int = plr.int + dobjt.effect(a, 2)
    If dobjt.effect(a, 1) = "SELL" Then takeobj = True: talkingto = dobj: sellspef dobjt              ': Exit Function ': "ARMOR" 'buy dobj, "NORMAL", dummyobj: takeobj = True
    If Left(dobjt.effect(a, 1), 4) = "SELL" And Len(dobjt.effect(a, 1)) > 4 Then takeobj = True: talkingto = dobj: showform10 Right(dobjt.effect(a, 1), Len(dobjt.effect(a, 1)) - 4), , , , , Val(dobjt.effect(a, 2))
    'If dobjt.effect(a, 1) = "SELLPOTIONS" Then takeobj = True: talkingto = dobj: showform10 "POTIONS" ': Form10.Show 1  ' buy dobj, "Potions", dummyobj: takeobj = True
    'If dobjt.effect(a, 1) = "SELLARMOR" Then takeobj = True: talkingto = dobj: showform10 "ARMOR" ': Form10.Show 1  'buy dobj, "Armor", dummyobj: takeobj = True
    'If dobjt.effect(a, 1) = "SELLCLOTHES" Then takeobj = True: talkingto = dobj: showform10 "" ' Form10.Show 1  'buy dobj, "Clothes", dummyobj: takeobj = True
    'If dobjt.effect(a, 1) = "SELLWEAPONS" Then takeobj = True: talkingto = dobj: showform10 "WEAPON" ': Form10.Show 1  'buy dobj, "Weapons", dummyobj: takeobj = True
    'If dobjt.effect(a, 1) = "SELLCARGO" Then takeobj = True: talkingto = dobj: showform10 "CARGOTRADE"
    'If dobjt.effect(a, 1) = "SELLGEMS" Then takeobj = True: talkingto = dobj: showform10 "GEMS"
    If dobjt.effect(a, 1) = "NPC" Then
'    Form2.msg objs(dobj).name & vbCrLf & objs(dobj).string2, objs(dobj).string, objtypes(objs(dobj).type).r, objtypes(objs(dobj).type).g, objtypes(objs(dobj).type).b, objtypes(objs(dobj).type).l: takeobj = True
    talkingto = dobjt.name: takeobj = True
    Form10.disptext stdfilter(objs(dobj).string2), 1, objs(dobj).string: Form10.Show 1
    'Form10.disptext stdfilter(objs(dobj).string2), 1: Form10.loaddapic objs(dobj).string: Form1.Timer1.Enabled = False: Form10.Show 1: Form1.Timer1.Enabled = True
    'Form2.msg objs(dobj).name & vbCrLf & swaptxt2(swaptxt2(objs(dobj).string2, "/", ","), "$", Chr(34)), objs(dobj).string, objtypes(objs(dobj).type).r, objtypes(objs(dobj).type).g, objtypes(objs(dobj).type).b, objtypes(objs(dobj).type).l: takeobj = True
    End If
    
    If dobjt.effect(a, 1) = "Clothes" Then playsound "clothes1.wav"
    If dobjt.effect(a, 1) = "Clothes" And Not (Ccom = "Gem" And X = -1) Then
        If invnum = Empty Then invnum = 50 'Or MsgBox("This will increase your total armor. Would you like to equip it?", vbYesNo) = vbYes
        'If findclothesmatch(dobjt.effect(a, 4), dobjt.effect(a, 5)) < Val(dobjt.effect(a, 3)) Then
        
        #If USELEGACY <> 1 Then
        'm'' Naga class handler : it cant wear lowerpart clothes
        If plr.Class = "Naga" And X = -1 Then 'm''
            wr1 = dobjt.effect(a, 4) 'm''
            wr2 = dobjt.effect(a, 5) 'm''
            If wr1 = "Panties" Or wr1 = "Lower" Or wr2 = "Panties" Or wr2 = "Lower" Then 'm''
                playsound "failed1.wav" 'm''
                gamemsg "You can't wear that!" 'm''
                GoTo 5 'm''
            End If 'm''
        End If
        #End If
        If X = -1 Then killitem inv(invnum): destroy = 2: F = addclothes(dobjt.name, dobjt.effect(a, 2), Val(dobjt.effect(a, 3)), dobjt.effect(a, 4), dobjt.effect(a, 5), dobjt.r, dobjt.g, dobjt.b, dobjt.l, , getraster(dobjt)): clothes(F).obj = dobjt: checkclothes F: Form1.updatbody: GoTo 5 Else If frominv = 0 Then getitem dobjt: If X > 0 Then map(X, Y).object = 0: GoTo 5 Else GoTo 5
        'If findclothesmatch(dobjt.effect(a, 4), dobjt.effect(a, 5)) >= Val(dobjt.effect(a, 3)) Then If MsgBox("This will weaken your total armor. Would you like to equip it anyway?", vbYesNo) = vbYes Then killitem inv(invnum): destroy = 2: f = addclothes(dobjt.name, dobjt.effect(a, 2), Val(dobjt.effect(a, 3)), dobjt.effect(a, 4), dobjt.effect(a, 5), dobjt.r, dobjt.g, dobjt.b, dobjt.l, , getraster(dobjt)): clothes(f).obj = dobjt: checkclothes f: Form1.updatbody Else If frominv = 0 Then getitem dobjt: If X > 0 Then map(X, Y).object = 0
    End If
    
                            'old weapon check: dobjt.effect(a, 3) * dobjt.effect(a, 4) > wep.dice * wep.damage Or
5   If dobjt.effect(a, 1) = "Weapon" Then If frominv = 1 And Not Ccom = "Gem" Then getitem wep.obj: wep.digged = 0: wep.r = dobjt.r: wep.g = dobjt.g: wep.b = dobjt.b: wep.l = dobjt.l: wep.obj = dobjt: lwepgraph dobjt.effect(a, 2), Form1.Picture1: destroy = 1: wep.dice = dobjt.effect(a, 3): wep.damage = dobjt.effect(a, 4): wep.type = dobjt.effect(a, 5): loadbijdang Else If frominv = 0 And Not Ccom = "Gem" Then getitem dobjt: destroy = 1: loadbijdang
    If dobjt.effect(a, 1) = "IfQuest" Then If Not ifquest(dobjt.effect(a, 2)) Then takeobj = True: Exit Function
    If dobjt.effect(a, 1) = "NotQuest" Then If ifquest(dobjt.effect(a, 2)) Then takeobj = True: Exit Function
    If dobjt.effect(a, 1) = "SetQuest" Then ifquest dobjt.effect(a, 2), 1
    If dobjt.effect(a, 1) = "Destruct" Or destroy = 1 Then If Not X = 0 Then If X > -1 Then map(X, Y).object = 0 Else killitem inv(invnum)
    If destroy = 2 Then If Not X = 0 Then If X > -1 Then map(X, Y).object = 0
    If dobjt.effect(a, 1) = "Give" Then createobj dobjt.effect(a, 2), plr.X, plr.Y: takeobj plr.X, plr.Y, 0, dummyobj
    If dobjt.effect(a, 1) = "Become" Then createobj dobjt.effect(a, 2), X, Y, dobjt.effect(a, 3), dobjt.effect(a, 4), dobjt.effect(a, 5)
    If dobjt.effect(a, 1) = "Heal" Then plrdamage transval(dobjt.effect(a, 2)) * -1
    If dobjt.effect(a, 1) = "Givemp" Then plr.mp = plr.mp + transval(dobjt.effect(a, 2)): If plr.mp > getmpmax Then plr.mp = getmpmax
    If dobjt.effect(a, 1) = "Spell" Then castspell dobjt.effect(a, 2), , 1
    If dobjt.effect(a, 1) = "Teleport" Then plr.X = dobjt.effect(a, 2): plr.Y = dobjt.effect(a, 3)
    If dobjt.effect(a, 1) = "MapTeleport" Then If plr.instomach = 0 Then gotomap dobjt.effect(a, 2): plr.X = dobjt.effect(a, 3): plr.Y = dobjt.effect(a, 4) Else gamemsg "You cannot leave while you're being digested!"
    If dobjt.effect(a, 1) = "Giveexp" Then plr.exp = plr.exp + dobjt.effect(a, 2)
    If dobjt.effect(a, 1) = "GiveGold" Then
        randsound "gold", 3
        plr.gp = plr.gp + dobjt.effect(a, 2)
        addtext dobjt.effect(a, 2) & " Gold", , , 250, 250
        If plr.Class = "Sorceress" Then plr.gp = plr.gp + (dobjt.effect(a, 2) * 0.1) 'Sorceress gold bonus
        goldbon = skilltotal("Greed", 10, 4)
        'If goldbon > 0 Then
        plr.gp = plr.gp + dobjt.effect(a, 2) * (1 + goldbon / 100)
        End If
    If dobjt.effect(a, 1) = "Mongold" Then
        randsound "gold", 3
        gp = Val(objs(dobj).string)
        skillpercmod gp, "Greed", 10, 4
        If plr.Class = "Sorceress" Then gp = gp * 1.3
        plr.gp = plr.gp + gp
        addtext Int(gp) & " Gold", , , 250, 250
        End If
        
    If dobjt.effect(a, 1) = "Cargo" Then
        addtext Val(objs(dobj).string) & " " & dobjt.effect(a, 2)
        addcargo dobjt.effect(a, 2), Val(objs(dobj).string)
    End If
    
    If dobjt.effect(a, 1) = "SystemBay" Then Shipsys.setsysnum (objs(dobj).string): Shipsys.Show: Shipsys.partsupdat: takeobj = True
        
    If dobjt.effect(a, 1) = "GivePart" Then getpart dobjt.effect(a, 2)
        
    If dobjt.effect(a, 1) = "Message" Then addtext dobjt.effect(a, 2), , , , , , lesser(30, Len(dobjt.effect(a, 2))) ':  Form10.disptext dobjt.effect(a, 2), 1, dobjt.effect(a, 3): Form10.Show 1
    If dobjt.effect(a, 1) = "Lifeplus" Then plr.lpotionlev = plr.lpotionlev + dobjt.effect(a, 2)
    If dobjt.effect(a, 1) = "Manaplus" Then plr.mpotionlev = plr.mpotionlev + dobjt.effect(a, 2)
    If dobjt.effect(a, 1) = "Lifepotion" Then plr.lpotions = plr.lpotions + dobjt.effect(a, 2)
    If dobjt.effect(a, 1) = "Manapotion" Then plr.mpotions = plr.mpotions + dobjt.effect(a, 2)
    If dobjt.effect(a, 1) = "Conversation" Then
        startpos = dobjt.effect(a, 4)
        zoink = getquestpending(dobjt.name)
        If Not zoink = "" Then startpos = zoink
        talkingto = dobjt.name: talkingtonum = dobj: showform10 dobjt.effect(a, 2), 1, dobjt.effect(a, 3), (startpos): takeobj = True
    End If


Next a

orginv
Form1.updatinv
updathp


End Function

Function castspell(spellnum, Optional target = 0, Optional nomp = 0, Optional shotx = 400, Optional shoty = 300)

On Error GoTo 99
GoTo 109
99 gamemsg "Invalid Target": Exit Function
109

If Left(Ccom, 6) = "SCROLL" Then Ccom = ""

castspell = False
Dim spelltype As String

If plr.plrdead > 0 Then Exit Function

'Find spell if given by name
If Val(spellnum) = 0 Then
For a = 1 To totalspells
    If spells(a).name = spellnum Then spellnum = a: Exit For
Next a
End If

Dim spell As spellT

spell = spells(spellnum)
If spell.target = "Target" And target = 0 And shotx = 400 And shoty = 300 Then Ccom = "SCROLL" & spellnum: Form1.Command2.caption = "Pick your Target": Exit Function
'If spell.target = "Target" And spelltype = "" And target = 0 Then Exit Function

mpcost = transval(spell.mp)

'Spell Mastery
skillpercminus mpcost, "Spell Mastery", 10, 3

If Left(spell.mp, 4) = "PERM" Then permanent = 1 Else permanent = 0

If target = 0 And spell.effect = "POSSESS" Then Exit Function

'If nomp = 0 Then If plr.mp < mpcost Then gamemsg "Not enough MP": randsound "mana", 2: Exit Function
If nomp = 0 Then If losemp(mpcost, permanent) = False Then gamemsg "Not enough MP": randsound "mana", 2: Exit Function 'plr.mp = plr.mp - spell.mp

If isexpansion = 0 Then dmg = Int(transval(spell.amount) * ((getint * 0.2) + 1)) Else dmg = Int(transval(spell.amount) * ((getint * 0.1) + 1))

If getplrskill(spell.school) > 10 Then skillpercmod dam, spell.school, 10, 10

addfatigue Int(Sqr(mpcost)), 1

'Masteries
If spell.school = "Black" Then skillpercmod dmg, "Black Magic Mastery", 10, 3
If spell.school = "White" Then skillpercmod dmg, "White Magic Mastery", 10, 3
If spell.school = "Grey" Then skillpercmod dmg, "Grey Magic Mastery", 10, 3
If spell.school = "Grey" Then skillpercmod dmg2, "Grey Magic Mastery", 10, 3
target2 = getfromstring(spell.target, 1)
spelltype = getfromstring(spell.target, 2)

origdmg = spell.amount
dmg = Int(dmg)

dmg2 = Int(transval(spell.amount))
If dmg < 0 Then dmg = 1
'If spell.effect = "Damage" Then dmg = dmg * (1 + plr.int / 10): dmg = dmg * (1 + plr.int / 10)

If target2 = "Enchant" Then
    If target = 0 Then Exit Function
    'Enchant target = Enchant:Objreq(Clothes,Weapon,GEM etc.)
    Dim donk2 As objecttype
    donk2 = inv(target)
    If geteff(donk2, getfromstring(spell.target, 2), 2) = "" Then gamemsg "That spell cannot enchant that item.": Ccom = "": Exit Function
    
    If getfromstring(spell.effect, 1) = "Random" Then
    
    makemagicitem donk2, getfromstring(spell.effect, 2)
    Else:
    addeffect2 donk2, getfromstring(spell.effect, 1), getfromstring(spell.effect, 2), getfromstring(spell.effect, 3)
    addeffect2 donk2, getfromstring(spell.effect, 4), getfromstring(spell.effect, 5), getfromstring(spell.effect, 6)
    addeffect2 donk2, getfromstring(spell.effect, 7), getfromstring(spell.effect, 8), getfromstring(spell.effect, 9)
    End If
    inv(target) = donk2
    Ccom = ""
    Form1.updatinv
    Exit Function
End If

If target2 = "Aura" Then
    Dim donk As objecttype
    addeffect2 donk, getfromstring(spell.effect, 1), getfromstring(spell.effect, 3), getfromstring(spell.effect, 2)
    addeffect2 donk, getfromstring(spell.effect, 4), getfromstring(spell.effect, 6), getfromstring(spell.effect, 5)
    addeffect2 donk, getfromstring(spell.effect, 7), getfromstring(spell.effect, 9), getfromstring(spell.effect, 8)
    'Presently, you may have one active aura from each school.  This may be too tough, so watch it.
    addaura getfromstring(spell.target, 2), getfromstring(spell.target, 3), getfromstring(spell.target, 4), getfromstring(spell.target, 5), getfromstring(spell.target, 6), donk, spell.school, getfromstring(spell.target, 7)
    
End If

If Not target = 0 Or target2 = "Target" Then
    shotmult = 1
    If usingskill = "Split Spell" Then If spendsp(spell.mp / 3 * getplrskill("Split Spell")) Then shotmult = shotmult + getplrskill("Split Spell"): losemp mpcost / 2
    If spell.effect = "Damage" Then
    'MsgBox "You cast " & spell.name & " on " & montype(mon(target).type).name & ", inflicting " & dmg & " damage."
    skillpercmod dmg, "Firepower", 10, 5
    dmg = Int(dmg)
    
    If Not spelltype = "" Then
        tt = 2
        angleplus = Val(getfromstring(spell.target, 4))
        shots = Val(getfromstring(spell.target, 3))
        If shots = 0 Then shots = 1
        Radius = Val(getfromstring(spell.target, 5))
        bijdangt = Val(getfromstring(spell.target, 6))
        pierce = Val(getfromstring(spell.target, 7))
        'pierce = Val(getfromstring(spell.target, 5))
        If angleplus = 0 And shotmult > 0 Then angleplus = 7
        If spelltype = "ARROW" Then tt = 1
        If spelltype = "FIRE" Then tt = 2
        If spelltype = "FIRE2" Then tt = 3
        If spelltype = "LIT" Then tt = 4
        If spelltype = "LIT2" Then tt = 5
        If spelltype = "OOZE" Then tt = 6
        If spelltype = "BOLT" Then tt = 7
        'If Not tt = 0 Then
        shootat shotx, shoty, tt, "PLAYER:" & dmg & ":1:" & pierce & ":" & Radius & ":" & bijdangt, , , shots * shotmult, angleplus, Val(getfromstring(spell.target, 7))
    End If
        
    If spelltype = "" And target > 0 Then
    makebijdang mon(target).X, mon(target).Y, 3
    damagemon target, dmg
    End If
    
    If spelltype = "" And target = 0 Then Exit Function
    
    End If
    
    If Not target = 0 Then
    If spell.effect = "POSSESS" Then playsound "spell1.wav": If dmg > mon(target).hp Then mon(target).owner = 2: gamemsg "You have successfully taken control of the " & montype(mon(target).type).name & "!" Else gamemsg "You have failed to control the " & montype(mon(target).type).name & "."
    End If
    
    If spell.effect = "Splitarrow" And wep.type = "Bow" Then
    tt = 1
    shootat shotx, shoty, tt, wep.dice & ":" & wep.damage, , , Val(getfromstring(spell.target, 3)), Val(getfromstring(spell.target, 4)), Val(getfromstring(spell.target, 5))
    End If
    
    If spell.effect = "Teleport" Then
        If InStr(1, plr.curmap, "Thirshasvolcano.txt") > 0 Then gamemsg "You cannot teleport here.": Exit Function
        getXY shotx, shoty
        plr.X = shotx: plr.Y = shoty
        makebijdang plr.X, plr.Y, 3
        If plr.instomach > 0 Then mon(plr.instomach).cell = 1: plr.instomach = 0: plr.mp = Int(plr.mp / 2) 'Uses half of total MP to teleport out of stomachs
        playsound "spell1.wav"
    End If
    
    allmove = 1
End If

If spell.effect = "Summon" Then
    summonmonster getfromstring(spell.target, 1), getfromstring(spell.target, 2), dmg2 + getint & ":" & getfromstring(origdmg, 2) & ":" & getfromstring(origdmg, 3), getfromstring(spell.target, 3), getfromstring(spell.target, 4), getfromstring(spell.target, 5), getfromstring(spell.target, 6), getfromstring(spell.target, 7), getfromstring(spell.target, 8), Val(getfromstring(spell.target, 9))
End If

'm'' "Dispell" summons, asked by some Eka's folk
If spell.effect = "Revoke" Then 'm''
    For a = 1 To totalmonsters 'm''
        If mon(a).owner > 0 Then killmon a, 1 'm''
    Next a 'm''
End If 'm''

If spell.target = "Player" Or spell.target = "" Then

    
    playsound "spell1.wav"
    
    If spell.effect = "Heal" Then gamemsg "You cast " & spell.name & " and regain " & dmg & " HP.": plrdamage -dmg
    If spell.effect = "Givestr" Then gamemsg "You cast " & spell.name & " and your strength increases!": plr.str = plr.str - plr.strboost: plr.strboost = dmg2: plr.str = plr.str + plr.strboost: boostcount = 50
    If spell.effect = "Givedex" Then gamemsg "You cast " & spell.name & " and your dexterity increases!": plr.dex = plr.dex - plr.dexboost: plr.dexboost = dmg2: plr.dex = plr.dex + plr.dexboost: boostcount = 50
    If spell.effect = "Giveint" Then gamemsg "You cast " & spell.name & " and your intelligence increases!": plr.int = plr.int - plr.intboost: plr.intboost = dmg2: plr.int = plr.int + plr.intboost: boostcount = 50
    If spell.effect = "Regen" Then gamemsg "You cast " & spell.name: plr.regen = dmg: boostcount = 50

    If spell.effect = "Create" Then createobj spell.amount, plr.X, plr.Y, spell.amount
    If spell.effect = "Undigest" Then gamemsg "You cast " & spell.name & " to restore one piece of your armor.": undigest
    If spell.effect = "EWeapon" Then gamemsg "You cast a spell to enhance your weapon for as long as you're holding it!": wep.damage = wep.damage + dmg
    If spell.effect = "EArmor" Then gamemsg "You cast a spell to enhance your " & clothes(topclothes()).name & " for as long as you're wearing it!": clothes(topclothes()).armor = clothes(topclothes()).armor + dmg
    
    If spell.effect = "Armboost" Then gamemsg "You cast a spell that temporarily protects you from harm.": plr.armorboost = dmg
    
    If spell.effect = "Recolor" Then
        Form1.FileD.ShowColor
        redd = Form1.FileD.color Mod 256
        green = Int((Form1.FileD.color Mod 65536) / 256)
        blue = Int(Form1.FileD.color / 65536)
        If Form1.List1.ListIndex = -1 Then Exit Function
        inv(Form1.List1.ItemData(Form1.List1.ListIndex)).r = redd
        inv(Form1.List1.ItemData(Form1.List1.ListIndex)).g = green
        inv(Form1.List1.ItemData(Form1.List1.ListIndex)).b = blue
        inv(Form1.List1.ItemData(Form1.List1.ListIndex)).l = 0.5
    End If

    If spell.effect = "Recolorall" Then
        Form1.FileD.ShowColor
        redd = Form1.FileD.color Mod 256
        green = Int((Form1.FileD.color Mod 65536) / 256)
        blue = Int(Form1.FileD.color / 65536)
'        If Form1.List1.ListIndex = -1 Then Exit Function
        
        For a = 1 To 16
        clothes(a).obj.r = redd
        clothes(a).obj.g = green
        clothes(a).obj.b = blue
        clothes(a).obj.l = 0.5
        takeoffclothes a
        Next a
        
'        inv(Form1.List1.ItemData(Form1.List1.ListIndex)).r = redd
'        inv(Form1.List1.ItemData(Form1.List1.ListIndex)).g = green
'        inv(Form1.List1.ItemData(Form1.List1.ListIndex)).b = blue
'        inv(Form1.List1.ItemData(Form1.List1.ListIndex)).l = 0.5
    End If

End If

If spell.target = "All" Then
    If spell.effect = "Damage" Then
    gamemsg "You cast " & spell.name & ", inflicting " & dmg & " damage on all creatures within your range."
    For a = 1 To totalmonsters
        If mon(a).X = 0 Or mon(a).Y = 0 Then GoTo 5
        If diff(plr.X, mon(a).X) < 8 And diff(plr.Y, mon(a).Y) < 8 Then damagemon a, dmg
5    Next a
    End If
End If

If spell.effect = "Damage" And spelltype = "" Then playsound "lightning1.wav"

turnthing
castspell = True

End Function


Function transval(ByVal valstr) As Double
'Value Translator
'By default, transval will use the target's value. Adding an X to the begginning will switch it.

'If it's just a number, use it
If Val(valstr) > 0 Then transval = Val(valstr): Exit Function
transval = getnum(valstr)
If Left(valstr, 1) = "R" Then transval = rolldice(getnum(valstr), getnum(Right(valstr, (InStr(1, valstr, "Die")))))


End Function

Function getnum(ByVal str) As Double
'Chops words off and gets a number out of them for the val function
'Returns a 1 if no numbers are found
mult = 1

5 If Len(str) = 0 Then getnum = 1: Exit Function
If Left(str, 1) = "-" Then mult = -1
If Val(Left(str, 1)) = 0 Then str = Right(str, Len(str) - 1): GoTo 5

getnum = Val(str)

End Function

Function getfromstring(ByVal str, ByVal slot As Byte) As String
'Finds data within a string, separated by :'s

mult = 1
curslot = 1
pos = 1
If InStr(pos, str, ":") = 0 Then If slot = 1 Then getfromstring = str: Exit Function Else Exit Function
str = str & ":" 'Just to make sure the thing reads it right, so you don't have to put a ':' at the end...
5 If pos > Len(str) Or pos = 0 Then getfromstring = "": Exit Function
If curslot = slot Then gstr = Mid(str, pos, InStr(pos, str, ":") - pos) Else curslot = curslot + 1: pos = InStr(pos, str, ":") + 1: GoTo 5

'If Left(str, 1) = "-" Then mult = -1
'If Val(Left(str, 1)) = 0 Then str = Right(str, Len(str) - 1): GoTo 5
'Debug.Print "Getfromstring returning " & gstr
getfromstring = gstr

End Function

Function updathp()
'm'' print the various status. skipped some using the experimental new UI
#If USELEGACY = 1 Then
Form1.Text1.text = "HP:" & plr.hp & "/" & gethpmax
rat = plr.hp / gethpmax
Form1.Text1.ForeColor = RGB(0, 0, 255)

If rat < 0.9 Then Form1.Text1.ForeColor = RGB(0, 255, 0)
If rat < 0.8 Then Form1.Text1.ForeColor = RGB(155, 255, 0)
If rat < 0.7 Then Form1.Text1.ForeColor = RGB(255, 255, 0)
If rat < 0.5 Then Form1.Text1.ForeColor = RGB(255, 155, 0)
If rat < 0.3 Then Form1.Text1.ForeColor = RGB(255, 0, 0)
#End If

If plr.plrdead = 1 Then GoTo 5

If plr.hp < 1 And plr.instomach = 0 Then plr.hp = 1
'If plr.hp < 1 And plr.instomach = 0 Then playsound "die.wav": MsgBox "You finally succumb to the harsh blows of your opponent, falling to the ground and making you very vulnerable.": plr.plrdead = 1
If plr.hp < 1 And plr.instomach > 0 Then
    'playsound "die.wav"
    'gamemsg "As you mercilessly churn inside " & montype(mon(plr.instomach).type).name & "'s hot depths, you gradually begin to lose consciousness. Soon you pass out entirely, leaving the beast to digest you at it's leisure..."
    'plr.plrdead = 1
    End If
5
Form1.Text2.text = "MP:" & plr.mp & "/" & getmpmax

 If plr.sp > plr.spmax Then plr.sp = plr.spmax
Form1.Text8.text = "Combat Points: " & plr.sp & "/" & plr.spmax

zoink = plr.name & " the " & plr.Class & vbCrLf & "Exp: " & plr.exp & "/" & plr.expneeded & vbCrLf & "Level " & plr.level _
    & vbCrLf & vbCrLf & "Strength:" & getstr & "/" & plr.str & vbCrLf & "Endurance:" & plr.endurance & "/" & getend & vbCrLf & "Dexterity:" & getdex & "/" & plr.dex & vbCrLf & "Intelligence:" & getint & "/" & plr.int
If plr.charpoints > 0 Then zoink = zoink & vbCrLf & "(Click here to spend character points)"
Form1.Text3.text = zoink

Form1.Text4.text = "Gold: " & plr.gp

End Function

Function gainlevel(Optional cheat = 0)

If plr.exp < plr.expneeded And cheat = 0 Then Exit Function
plr.level = plr.level + 1
plr.exp = plr.exp - plr.expneeded: plr.expneeded = plr.expneeded * 1.25 + 200
'If expansion = 0 Then
'plr.expneeded = plr.level ^ 3 * 100 + 500 'Else plr.expneeded = plr.level ^ 3 * 100 + 200
plr.expneeded = (plr.level * 1000#) * (1# + plr.level / 20#)

plr.expneeded = round(plr.expneeded)

plr.mpmax = plr.mpmax + plr.mplost: plr.mplost = 0
skillmod plr.hpmax, "Endurance", 2, 1
skillmod plr.mpmax, "Mana Mastery", 2, 1

plr.skillpoints = plr.skillpoints + greater(3, Sqr(plr.level))
plr.charpoints = plr.charpoints + 1
plr.spmax = plr.spmax + 5

'If plr.Class = "Fighter" Then classlevel 2, 0.5, 2, 2, 1, 0
If isexpansion = 1 Then GoTo 7

If plr.Class = "Sorceress" Then classlevel 0.8, 3, 0, 1, 3, 3 ': givespell "Black", "Grey", 3
If plr.Class = "Valkyrie" Then classlevel 2, 1, 2, 2, 1, 1 ': givespell "Grey", "White", 1
If plr.Class = "Amazon" Then classlevel 3, 0, 3, 1, -1, 0
If plr.Class = "Huntress" Then classlevel 2, 1, 1, 3, 1, 1 ': givespell "Huntress", "Black", 2
If plr.Class = "Priestess" Then classlevel 1.3, 1.5, 2, 1, 2, 2 ': givespell "White", "Grey", 2
If plr.Class = "Enchantress" Then classlevel 0.8, 3, 0, 1, 3, 3 ': givespell "Grey", "Black", 3 'Enchantresses use grey magic

If plr.Class = "TombRaider" Then classlevel 2.5, 0, 2, 3, 2, 0
If plr.Class = "Caller" Then classlevel 1, 2, 1, 2, 2, 1 ': givespell "Summon", "Black", 3
If plr.Class = "Streetfighter" Then classlevel 2.5, 0, 2, 2, 1, 0
If plr.Class = "Naga" Then classlevel 3, 1, 3, 3, 2, 1 ': givespell "Black", "NONE", 1

If plr.Class = "Succubus" Then classlevel 3, 1, 3, 3, 2, 1 ': givespell "Black", "NONE", 1
If plr.Class = "Angel" Then classlevel 2, 2, 1, 3, 2, 2 ': givespell "White", "Black", 2

'If plr.Class = "" Then classlevel 1, 1, 1, 1, 1, 1: givespell "None", "None", 3
7 If plr.Class = "" Or isexpansion = 1 Then classlevel plr.classdata.hpmult, plr.classdata.mpmult, plr.classdata.strmult, plr.classdata.dexmult, plr.classdata.intmult, 1, plr.classdata.endmult: givespell "None", "None", 3


5
'If cheat = 1 Then

If plr.level Mod 4 = 0 Then plr.combatskillpoints = plr.combatskillpoints + 1
'Grant combat skills at 1 every 4 levels
'leveldiv = 4
'If Left(plr.combatskills(plr.level / leveldiv), 1) = "#" Then plr.combatskills(level / leveldiv) = Right(plr.combatskills(level / leveldiv), Len(plr.combatskills(level / leveldiv)) - 1)

plr.hp = gethpmax
'If cheat = 0 Then
gamemsg "You have gained a level!"
addtext "You have gained a level!", , , 255, 230, 230
playsound "gainedlevel.wav"
End Function

Function classlevel(Optional hpmult = 1, Optional mpmult = 1, Optional strmult = 1, Optional dexmult = 1, Optional intmult = 1, Optional spellmult = 1, Optional endmult = 1)

'plr.hpmax = plr.hpmax * 1.1 + (6 * hpmult)
plr.hpmax = plr.hpmax + plr.level + (6 * hpmult)
plr.mpmax = plr.mpmax + mpmult * 5 + plr.level '(mpmult * 8) + (level hpmult)
'plr.mpmax = plr.mpmax * (1 + mpmult / 50) + (3 * mpmult)

'Mults:
'5=Every level, 4=Every other level, 3=Every 3 levels, 2=Every 4 levels, 0=Never

moddr = Int(6 - strmult)
If plr.level Mod moddr = 0 Then plr.str = plr.str + 1
moddr = Int(6 - dexmult)
If plr.level Mod moddr = 0 Then plr.dex = plr.dex + 1
moddr = Int(6 - intmult)
If plr.level Mod moddr = 0 Then plr.int = plr.int + 1
moddr = Int(6 - endmult)
If plr.level Mod moddr = 0 Then plr.endurance = plr.endurance + 1

End Function

Sub movobjs()
For a = 1 To objtotal
    If objs(a).type = 0 Then GoTo 5
    If objtypes(objs(a).type).effect(1, 1) = "Mobile" Then If roll(8) < objtypes(objs(a).type).effect(1, 2) Then targmove objs(a).X, objs(a).Y, roll(3) - 2, roll(3) - 2, "Object" & a, objs(a).xoff, objs(a).yoff
5 Next a

End Sub

Sub targmove(ByRef targx, ByRef targy, ByVal xplus, ByVal yplus, ByVal otype As String, ByRef xoff, ByRef yoff)
'Otype should have the object's reference number attached to it

If (targx + xplus) < 1 Or (targx + xplus) > mapx Then Exit Sub
If (targy + yplus) < 1 Or (targy + yplus) > mapy Then Exit Sub
If map(targx + xplus, targy + yplus).blocked = 1 Then Exit Sub
If map(targx + xplus, targy + yplus).monster > 0 Then Exit Sub
If map(targx + xplus, targy + yplus).object > 0 And Left(otype, 6) = "Object" Then Exit Sub
If targx + xplus = plr.X Or targy + yplus = plr.Y Then Exit Sub

If Left(otype, 6) = "Object" Then map(targx, targy).object = 0

xoff = xplus * -24
yoff = yplus * -24

targx = targx + xplus
targy = targy + yplus

If Left(otype, 6) = "Object" Then map(targx, targy).object = getnum(otype)


End Sub

Function swaptxt(txt, swap1, swap2)

'Just a little routine to replace text
'It will look for swap1 and replace it with swap2

5 swappem = InStr(1, txt, swap1)
If swappem > 0 Then txt = Left(txt, swappem - 1) & swap2 & Right(txt, Len(txt) - swappem - Len(swap1) + 1): GoTo 5

swaptxt = txt

End Function

Function swaptxt2(ByVal txt, ByVal swap1, ByVal swap2)

'Just a little routine to replace text
'It will look for swap1 and replace it with swap2

5 swappem = InStr(1, txt, swap1)
If swappem > 0 Then txt = Left(txt, swappem - 1) & swap2 & Right(txt, Len(txt) - swappem - Len(swap1) + 1): GoTo 5

swaptxt2 = txt

End Function

Sub fillmap(tilet)

For a = 1 To mapx
For b = 1 To mapy
    map(a, b).tile = tilet
Next b
Next a
End Sub

Sub fillovr(ot)

For a = 1 To mapx
For b = 1 To mapy
    map(a, b).ovrtile = ot: map(a, b).blocked = 1
Next b
Next a
End Sub

Sub createbuilding(Optional ByVal xsize = 0, Optional ByVal ysize = 0, Optional ByVal btype = "", Optional ByVal overlap = 1, Optional ByVal tile = 3, Optional ByVal ovrtile = 1, Optional ByVal X = 0, Optional ByVal Y = 0, Optional ByVal multdoors = 1)
5 bugger = bugger + 1: If bugger > 10 Then Exit Sub '(Only ten attempts will be made to place any particular building)
If xsize = 0 Then xsize = 3 + roll(9)
If ysize = 0 Then ysize = 3 + roll(9)

2 If X <= 0 Or xsize + X > mapx Then X = roll((mapx - xsize) - 6) + 2: GoTo 2
If Y <= 0 Or ysize + Y > mapy Then Y = roll((mapy - ysize) - 6) + 2: GoTo 2


'checks spot
For a = X To xsize + X
For b = Y To ysize + Y
    If map(a, b).used = 1 Then used = 1: Exit For
Next b
If used = 1 Then Exit For
Next a

If used = 1 And overlap = 0 Then X = 0: Y = 0: GoTo 5

'place building

For a = X To xsize + X
For b = Y To ysize + Y
    If map(a, b).used = 0 Or overlap > 0 Then map(a, b).ovrtile = 0: map(a, b).blocked = 0
    If b = ysize + Y Or b = Y Then If map(a, b).used = 0 Or overlap = 1 And map(a, b).object = 0 Then map(a, b).ovrtile = ovrtile: map(a, b).blocked = 1
    If a = xsize + X Or a = X Then If map(a, b).used = 0 Or overlap = 1 And map(a, b).object = 0 Then map(a, b).ovrtile = ovrtile: map(a, b).blocked = 1
    map(a, b).tile = tile
    
    map(a, b).used = 1
Next b
Next a

'place door(s)
Dim doorx As Integer
Dim doory As Integer

For a = 1 To multdoors
tries = 0
7 doorside = roll(4): tries = tries + 1
If doorside = 1 Then doorx = X + xsize / 2: doory = ysize + Y: xp = 0: yp = 1
If doorside = 2 Then doorx = X + xsize / 2: doory = Y: xp = 0: yp = -1
If doorside = 3 Then doory = Y + ysize / 2: doorx = xsize + X: yp = 0: xp = 1
If doorside = 4 Then doory = Y + ysize / 2: doorx = X: yp = 0: xp = -1
If map(doorx, doory).ovrtile = 0 And tries > 15 Then GoTo 7
map(doorx, doory).ovrtile = 0: map(doorx, doory).blocked = 0

doorx = doorx + xp: doory = doory + yp

'GoTo 3 'Get rid of the tunnelling rooms do for doors; rely on the checkaccess stuff instead
If doorx + xp < 1 Or doorx + xp > mapx Or doory + yp < 1 Or doory + yp > mapy Then GoTo 3

Do Until map(doorx + xp, doory + yp).blocked = 0 And map(doorx, doory).ovrtile = 0
If doorx + xp < 1 Or doorx + xp > mapx Or doory + yp < 1 Or doory + yp > mapy Then Exit Do
    map(doorx, doory).tile = tile
    map(doorx, doory).used = 1
    map(doorx, doory).ovrtile = 0
    map(doorx, doory).blocked = 0
    doorx = doorx + xp: doory = doory + yp
    If doorx + xp < 1 Or doorx + xp > mapx Or doory + yp < 1 Or doory + yp > mapy Then Exit Do
Loop
3 Next a

'roomtypes

If Left(btype, 8) = "Treasure" Then
    worth = getnum(btype)
    If worth < 1 Then worth = mapjunk.level
    aroll = roll(worth)
    If aroll > 10 Then aroll = roll(10)
    If aroll < worth / 10 Then aroll = Int(worth / 10)
    maxtries = 12
8    xr = roll(xsize - 3) + X + 1
    yr = roll(ysize - 3) + Y + 1
    If worth <= 0 Or maxtries <= 0 Then GoTo 12
    If map(xr, yr).object > 0 Then maxtries = maxtries - 1: GoTo 8
    If map(xr, yr).blocked = 1 Then maxtries = maxtries - 1: GoTo 8
    If roll(10) = 1 Then creategem xr, yr Else createtreasure aroll, xr, yr: maxtries = 12
    'worth = worth - aroll
    'GoTo 8
12
End If

If Left(btype, 7) = "Clothes" Or Left(btype, 5) = "Armor" Then
    worth = getnum(btype)
    If worth < 1 Then worth = mapjunk.level
    'worth = greater(worth, lowestlevel)
    worth = mapjunk.level
    maxtries = 12
    aroll = Int(worth)
    If aroll < 1 Then aroll = 1
    'worth = worth Mod 4: If worth = 0 Then worth = 4
15  xr = roll(xsize - 2) + X + 1
    yr = roll(ysize - 2) + Y + 1
    If worth <= 0 Or maxtries <= 0 Then GoTo 16
    If map(xr, yr).object > 0 Then maxtries = maxtries - 1: GoTo 15
    If map(xr, yr).blocked = 1 Then maxtries = maxtries - 1: GoTo 15
If Left(btype, 7) = "Clothes" Then createclothes worth, xr, yr, , 5: maxtries = 16
If Left(btype, 5) = "Armor" Then createclothes worth, xr, yr, "Armor", 4: maxtries = 16
    'worth = worth / 2
    'GoTo 15
End If
16

If Left(btype, 6) = "Potion" Then
    worth = getnum(btype)
    If worth < 1 Then worth = mapjunk.level
    aroll = Int(worth / 3)
    If aroll > 3 Then aroll = roll(3)
    If aroll < worth / 10 Then aroll = Int(worth / 10)
    If aroll <= 0 Then aroll = 1
    maxtries = 12
19    xr = roll(xsize - 3) + X + 1
    yr = roll(ysize - 3) + Y + 1
    If aroll <= 0 Then aroll = 1
    If worth <= 0 Or maxtries <= 0 Then GoTo 20
    If map(xr, yr).object > 0 Then maxtries = maxtries - 1: GoTo 19
    If map(xr, yr).blocked = 1 Then maxtries = maxtries - 1: GoTo 19
    createpotion aroll, xr, yr: maxtries = 12
    worth = worth - aroll
    GoTo 19
20
End If

If Left(btype, 9) = "Objective" Then
    worth = getnum(btype)
    If worth < 1 Then worth = mapjunk.level
    maxtries = 12
    stuffamt = 3
21
    xr = roll(xsize - 3) + X + 1
    yr = roll(ysize - 3) + Y + 1
    If aroll <= 0 Then aroll = 1
    If worth <= 0 Or maxtries <= 0 Then GoTo 22
    If map(xr, yr).object > 0 Then maxtries = maxtries - 1: GoTo 21
    If map(xr, yr).blocked = 1 Then maxtries = maxtries - 1: GoTo 21
    createobjective worth, xr, yr: maxtries = 12
    
    stuffamt = stuffamt - 1
    If stuffamt = 0 Then GoTo 21
22
End If


If Left(btype, 4) = "SELL" Then
    F = createobjtype(getstr2(btype, 1), getstr2(btype, 2), roll(255), roll(255), roll(255), roll(10) / 10, 2)
    addeffect F, getstr2(btype, 0), getstr2(btype, 3), getstr2(btype, 4)
    createobj objtypes(F).name, X + Int(xsize / 2), Y + Int(ysize / 2), objtypes(F).name
End If


End Sub

Function randarmor(Optional worth As Byte = 0, Optional ByRef Cost, Optional ByVal forceworth = 0) As objecttype

Dim buyarmor As clothesT

If worth = 0 Then worth = Int(averagelevel / 4)

aroll = roll(18)
'aroll = 19

wear1 = ""
wear2 = ""

weight = 1

'New thingy
4 aroll = roll(UBound(clothestypes()))
    If Not clothestypes(aroll).material = "Metal" Then GoTo 4
    basarm = clothestypes(aroll).armor: wear1 = clothestypes(aroll).wear1: wear2 = clothestypes(aroll).wear2: graph = clothestypes(aroll).graph: weight = clothestypes(aroll).weight: name = clothestypes(aroll).name
GoTo 7

If aroll = 1 Then name = "Halter": graph = "Halter1.bmp": basarm = 7: wear1 = "Upper": wear2 = "Bra"
If aroll = 2 Then name = "Chain Shirt": graph = "chainshirt1.bmp": basarm = 10: wear1 = "Upper": weight = 3
If aroll = 3 Then name = "Chain Skirt": graph = "chainskirt1.bmp": basarm = 10: wear1 = "Lower": weight = 3
If aroll = 4 Then name = "Heartplate": graph = "Breastplate2.bmp": basarm = 14: wear1 = "Upper": weight = 4
If aroll = 5 Then name = "Breastplate": graph = "Breastplate3.bmp": basarm = 15: wear1 = "Upper": wear2 = "Jacket": weight = 6
If aroll = 6 Then name = "Halfplate": graph = "Breastplate4.bmp": basarm = 11: wear1 = "Upper": weight = 3
If aroll = 7 Then name = "Bracers": graph = "Bracers1.bmp": basarm = 4: wear1 = "Arms": weight = 1
If aroll = 8 Then name = "Armored Skirt": graph = "armorskirt1.bmp": basarm = 7: wear1 = "Lower": weight = 2
If aroll = 9 Then name = "Full Plate": graph = "Breastplate5.bmp": basarm = 20: wear1 = "Upper": wear2 = "Jacket": weight = 8
If aroll = 10 Then name = "Arm Plates": graph = "armplates1.bmp": basarm = 7: wear1 = "Arms": weight = 2
If aroll = 11 Then name = "Leg Plates": graph = "legplates1.bmp": basarm = 12: wear1 = "Lower": weight = 6
If aroll = 12 Then name = "Heavy Plate": graph = "Breastplate6.bmp": basarm = 17: wear1 = "Upper": wear2 = "Jacket": weight = 10
If aroll = 13 Then name = "Chainmail Bodysuit": graph = "chainswimsuit.bmp": basarm = 15: wear1 = "Upper": wear2 = "Lower": weight = 4
If aroll = 14 Then name = "Brace Belt": graph = "belt5.bmp": basarm = 2: wear1 = "Belt" ': wear2 = "Lower"
If aroll = 15 Then name = "Double Brace Belt": graph = "belt1.bmp": basarm = 3: wear1 = "Belt" ': wear2 = "Lower"
If aroll = 16 Then name = "Armored Belt": graph = "belt2.bmp": basarm = 4: wear1 = "Belt": weight = 2 ': wear2 = "Lower"
If aroll = 17 Then name = "Chainmail Bikini": graph = "panties3.bmp": basarm = 4: wear1 = "Panties": wear2 = "Lower"
If aroll = 18 Then name = "Chainmail Swimsuit": graph = "chainswimsuit2.bmp": basarm = 10: wear1 = "Bra": wear2 = "Panties"
'If aroll = 19 Then name = "Ridged Plate": graph = "Breastplate7.bmp": basarm = 15: wear1 = "Upper"

'If aroll = 9 Then name = "Chainmail Dress": graph = "chaindress1.bmp": basarm = 20: wear1 = "Upper": wear2 = "Lower"


'buyarmor(slot).armor = basarm
'buyarmor(slot).name = randarmortype(buyarmor(slot).armor, buyarmor(slot).r, buyarmor(slot).g, buyarmor(slot).b, buyarmor(slot).l, buyarmor(slot).gold, worth) & " " & name

'buyarmor(slot).wear1 = wear1
'buyarmor(slot).wear2 = wear2
'buyarmor(slot).graph = graph
'randarmor = buyarmor(slot).name

'If Not reqwear1 = "" Then If Not wear1 = reqwear1 And Not wear2 = reqwear1 Then GoTo 3

7 buyarmor.armor = basarm
buyarmor.name = randarmortype(basarm, r, g, b, l, gold, worth, forceworth, weight, matnum) & " " & name
'If Not wear2 = "" Then gold = gold * 0.75
If wear1 = "Jacket" Or wear1 = "Belt" Then gold = gold * 2
If Not wear2 = "Jacket" Then If Not wear1 = "" And Not wear2 = "" Then gold = gold / 2
'If Not reqwear1 = "" Then If Not wear1 = reqwear1 And Not wear1 = reqwear2 Then GoTo 3
'If name = "Halter" Then gold = gold / 2
'If name = "Chainmail Bikini" Then gold = gold / 2
If name = "Plate Mail" Then gold = gold * 1.25
If name = "Leg Plates" Then gold = gold * 1.3
If wear1 = "Arms" Then gold = gold * 2

gold = gold * materialtypes(matnum).goldmult

Cost = Int(gold)
'buyarmor.wear1 = wear1
'buyarmor.wear2 = wear2
'buyarmor.graph = graph
Dim cobjt As objecttype
cobjt.graphname = "armorobj1.bmp"
cobjt.name = colorname & buyarmor.name
addeffect2 cobjt, "Clothes", graph, Int(basarm), wear1, wear2
addeffect2 cobjt, "Equipjunk", weight
addeffect2 cobjt, materialtypes(matnum).effects(1, 1), materialtypes(matnum).effects(1, 2), materialtypes(matnum).effects(1, 3)
addeffect2 cobjt, materialtypes(matnum).effects(2, 1), materialtypes(matnum).effects(2, 2), materialtypes(matnum).effects(2, 3)
addeffect2 cobjt, materialtypes(matnum).effects(3, 1), materialtypes(matnum).effects(3, 2), materialtypes(matnum).effects(3, 3)

Cost = getworth(cobjt)

'addeffect2 cobjt, "GRAPH", "clothes.bmp", r, g, b
'addeffect2 cobjt, "Pickup"
cobjt.r = r
cobjt.g = g
cobjt.b = b
cobjt.l = l
randarmor = cobjt

End Function

Function randarmortype(ByRef armor, ByRef r, ByRef g, ByRef b, ByRef l, ByRef gold, worth, Optional ByVal forceworth = 0, Optional ByRef weight = 1, Optional ByRef matnum)

'Find closest worth
acworth = 1
lastdiff = 100
For a = 1 To UBound(materialtypes())
    If diff(materialtypes(a).worth, worth) < lastdiff And materialtypes(a).type = "Metal" Then acworth = a: lastdiff = diff(materialtypes(acworth).worth, worth)
Next a


'New Thingy...again.
4 aroll = roll(UBound(materialtypes()))
    aroll = acworth
    matnum = aroll
    If Not materialtypes(aroll).type = "Metal" Then GoTo 4
    armult = materialtypes(aroll).armor: r = materialtypes(aroll).r: g = materialtypes(aroll).g: b = materialtypes(aroll).b: weight = weight * materialtypes(aroll).weight: armn = materialtypes(aroll).name
GoTo 7


armn = ""
armult = 1

3 aroll = worth + roll(3) - 2 'roll(7)
'If aroll = 11 Then Stop
If aroll > 11 Then aroll = 11
If aroll < 1 Then GoTo 3
'If worth > 11 Then aroll = 1
If forceworth > 0 Then aroll = forceworth
If aroll = 1 Then armn = "Leather": r = 105: g = 75: b = 0: l = 0.4: armult = 1 ': weight = weight * 1.3
If aroll = 2 Then armn = "Bronze": r = 55: g = 35: b = 0: l = 0.2: armult = 1.5: weight = weight * 1.8
If aroll = 3 Then armn = "Iron": r = 105: g = 105: b = 125: l = 0.2: armult = 2: weight = weight * 2
If aroll = 4 Then armn = "Steel": r = 155: g = 155: b = 155: l = 0.7: armult = 2.5: weight = weight * 1.7
If aroll = 5 Then armn = "Opal": r = 255: g = 195: b = 215: l = 0.5: armult = 3.5: weight = weight * 1.6
If aroll = 6 Then armn = "PinkSteel": r = 255: g = 105: b = 240: l = 0.5: armult = 4: weight = weight * 1.6
If aroll = 7 Then armn = "Enchanted Obsidian": r = 155: g = 155: b = 155: l = 0: armult = 5: weight = weight * 1.7
If aroll = 8 Then armn = "BloodSteel": r = 255: g = 0: b = 0: l = 0.3: armult = 6: weight = weight * 1.9
If aroll = 9 Then armn = "Blue Adamant": r = 0: g = 0: b = 255: l = 0.4: armult = 8: weight = weight * 2.2
If aroll = 10 Then armn = "Demonic Iron": r = 105: g = 15: b = 10: l = 0.1: armult = 10: weight = weight * 2.4
If aroll = 11 Then armn = "Angelic Steel": r = 255: g = 225: b = 150: l = 0.6: armult = 12: weight = weight * 1.5


'If isexpansion = 1 Then armult = (1 + 1 + armult) / 3

'If isexpansion = 0 Then
7 armor = armor * (1 + armult / 4): If armor < 1 Then armor = 1 Else armor = armor + armult * 2
If isexpansion = 0 Then gold = (armor * armor) * 10 '* ((roll(10) + 10) / 10)) + (armor * 20)) / 2
If isexpansion = 1 Then gold = ((armor + armult) * armor) * 10 '* ((roll(10) + 10) / 10)) + (armor * 20)) / 2

gold = gold * lesser(0.6, (1 - weight / 200)) '1/2% price reduction per weight, max 60%

randarmortype = armn

End Function

Function randclothes(Optional colorname = "", Optional r, Optional g, Optional b, Optional l, Optional ByVal worth As Byte = 0, Optional ByRef Cost, Optional reqwear1 = "", Optional reqwear2 = "") As objecttype
Dim buyarmor As clothesT
Dim cobjt As objecttype
cobjt.graphname = "clothes.bmp": cobjt.graphloaded = 0
If worth = 0 Then worth = Int(averagelevel / 5)

If colorname = "" Then r = 0: g = 0: b = 0: l = 0.5: colorname = gencolor(colorname, r, g, b, l)
'If colorname = "White" Then l = 0.8
3 aroll = roll(54)

wear1 = ""
wear2 = ""

'New thingy
4 aroll = roll(UBound(clothestypes()))
If Not clothestypes(aroll).material = "Cloth" Then GoTo 4
basarm = clothestypes(aroll).armor: wear1 = clothestypes(aroll).wear1: wear2 = clothestypes(aroll).wear2: graph = clothestypes(aroll).graph: weight = clothestypes(aroll).weight: name = clothestypes(aroll).name
GoTo 7

'This isn't used anymore
If aroll = 1 Then name = "Bikini Top": graph = "Bra1.bmp": basarm = 1: wear1 = "Bra"
If aroll = 2 Then name = "Bra": graph = "Bra2.bmp": basarm = 1: wear1 = "Bra"
If aroll = 3 Then name = "Panties": graph = "panties1.bmp": basarm = 1: wear1 = "Panties"
If aroll = 4 Then name = "Sports Bra": graph = "Bra4.bmp": basarm = 2: wear1 = "Bra"
If aroll = 5 Then name = "Corset Bra": graph = "Bra3.bmp": basarm = 3: wear1 = "Bra"

If aroll = 6 Then name = "Cut-out Swimsuit": graph = "swimsuit1.bmp": basarm = 3: wear1 = "Panties": wear2 = "Bra"
If aroll = 7 Then name = "Tank Swimsuit": graph = "swimsuit2.bmp": basarm = 3: wear1 = "Panties": wear2 = "Bra"

If aroll = 8 Then name = "Combat Dress": graph = "chunli1.bmp": basarm = 10: wear1 = "Upper": wear2 = "Lower"
If aroll = 9 Then name = "Pants": graph = "pants2.bmp": basarm = 4: wear1 = "Lower"
If aroll = 10 Then name = "Doublet Shirt": graph = "doublet1.bmp": basarm = 3: wear1 = "Upper"
If aroll = 11 Then name = "Gloves": graph = "gloves1.bmp": basarm = 4: wear1 = "Arms"
If aroll = 12 Then name = "Jacket": graph = "jacket1.bmp": basarm = 5: wear1 = "Jacket"
If aroll = 13 Then name = "Running Shorts": graph = "leatherbottom1.bmp": basarm = 3: wear1 = "Lower"
If aroll = 14 Then name = "Bodysuit": graph = "leathersuit1.bmp": basarm = 12: wear1 = "Upper": wear2 = "Lower"
If aroll = 15 Then name = "Bodytop": graph = "leathertop1.bmp": basarm = 3: wear1 = "Bra"
If aroll = 16 Then name = "Vestshirt": graph = "leathertop2.bmp": basarm = 4: wear1 = "Upper"
If aroll = 17 Then name = "Loincloth": graph = "loincloth1.bmp": basarm = 1: wear1 = "Lower"
If aroll = 18 Then name = "Loose Pants": graph = "pants1.bmp": basarm = 3: wear1 = "Lower"
If aroll = 19 Then name = "Skirt": graph = "skirt1.bmp": basarm = 3: wear1 = "Lower"
If aroll = 20 Then name = "Undersuit": graph = "leathersuit1.bmp": basarm = 5: wear1 = "Panties": wear2 = "Bra"

If aroll = 21 Then name = "Lace Bra": graph = "Bra5.bmp": basarm = 3: wear1 = "Bra"
If aroll = 22 Then name = "Lace Panties": graph = "panties2.bmp": basarm = 3: wear1 = "Panties"
If aroll = 23 Then name = "Fine Dress": graph = "dress2.bmp": basarm = 9: wear1 = "Upper": wear2 = "Lower"
If aroll = 24 Then name = "Short Dress": graph = "dress1.bmp": basarm = 8: wear1 = "Upper": wear2 = "Lower"
If aroll = 25 Then name = "Shirt": graph = "shirt1.bmp": basarm = 3: wear1 = "Upper"
If aroll = 26 Then name = "Longshirt": graph = "longshirt1.bmp": basarm = 7: wear1 = "Upper": wear2 = "Lower"
If aroll = 27 Then name = "Shortcape": graph = "cape1.bmp": basarm = 4: wear1 = "Jacket"
If aroll = 28 Then name = "Robe Shirt": graph = "robeshirt1.bmp": basarm = 3: wear1 = "Upper"
If aroll = 29 Then name = "Robe Skirt": graph = "robebottom1.bmp": basarm = 3: wear1 = "Lower"
If aroll = 30 Then name = "Shorts": graph = "shorts1.bmp": basarm = 3: wear1 = "Lower"

If aroll = 31 Then name = "Dress": graph = "dress3.bmp": basarm = 11: wear1 = "Upper": wear2 = "Lower"
If aroll = 32 Then name = "Tribal Dress": graph = "dress4.bmp": basarm = 9: wear1 = "Upper": wear2 = "Lower"

If aroll = 33 Then name = "Drape Shirt": graph = "robeshirt2.bmp": basarm = 4: wear1 = "Upper": wear2 = "Jacket"
If aroll = 34 Then name = "Short Skirt": graph = "skirt2.bmp": basarm = 3: wear1 = "Lower"
If aroll = 35 Then name = "Slit Skirt": graph = "skirt3.bmp": basarm = 3: wear1 = "Lower"
If aroll = 36 Then name = "Thigh Boots": graph = "boots1.bmp": basarm = 4: wear1 = "Lower"

If aroll = 37 Then name = "Corset": graph = "corset1.bmp": basarm = 4: wear1 = "Bra" ': wear2 = "Bra"

If aroll = 38 Then name = "Frilly Nightshirt": graph = "Frillyover.bmp": basarm = 3: wear1 = "Upper": wear2 = "Bra"
If aroll = 39 Then name = "Frilly Bra": graph = "Bra6.bmp": basarm = 3: wear1 = "Bra"

'Repeats of more common clothes
If aroll = 40 Then name = "Lace Panties": graph = "panties2.bmp": basarm = 3: wear1 = "Panties"
If aroll = 41 Then name = "Pants": graph = "pants2.bmp": basarm = 4: wear1 = "Lower"
If aroll = 42 Then name = "Bra": graph = "Bra2.bmp": basarm = 1: wear1 = "Bra"
If aroll = 43 Then name = "Panties": graph = "panties1.bmp": basarm = 1: wear1 = "Panties"

If aroll = 44 Then name = "Fishnet Stockings": graph = "fishnets1.bmp": basarm = 1: wear1 = "Legs"
If aroll = 45 Then name = "Panty Hose": graph = "fishnets2.bmp": basarm = 1: wear1 = "Legs"
If aroll = 46 Then name = "Teddy": graph = "teddy1.bmp": basarm = 3: wear1 = "Panties": wear2 = "Bra"
If aroll = 47 Then name = "Lace Corset": graph = "corset2.bmp": basarm = 5: wear1 = "Upper": wear2 = "Bra"
If aroll = 48 Then name = "Shoulder Sash": graph = "sash1.bmp": basarm = 2: wear1 = "Jacket" ': wear2 = "Bra"
'If aroll = 49 Then name = "Cape": graph = "cape2.bmp": basarm = 6: wear1 = "Jacket"
If aroll = 49 Then name = "Light Belt": graph = "belt3.bmp": basarm = 1: wear1 = "Belt" ': wear2 = "Bra"
If aroll = 50 Then name = "Heavy Belt": graph = "belt4.bmp": basarm = 2: wear1 = "Belt" ': wear2 = "Bra"
If aroll = 51 Then name = "Bikini Wrap": graph = "wrap1.bmp": basarm = 2: wear1 = "Panties": wear2 = "Bra"
If aroll = 52 Then name = "Fine Skirt": graph = "skirt4.bmp": basarm = 4: wear1 = "Lower"
If aroll = 53 Then name = "Ribbon Top": graph = "ribbontop1.bmp": basarm = 3: wear1 = "Upper"
If aroll = 54 Then name = "Biker Jacket": graph = "jacket2.bmp": basarm = 8: wear1 = "Jacket"


'If wear1 = "Upper" Then graph = getgener("uniform1.bmp", "uniform2.bmp", "uniform3.bmp", "uniform5.bmp", "uniform6.bmp"): wear2 = "": If graph = "uniform1.bmp" Then wear1 = "Bra": wear2 = "Upper"
'If wear1 = "Upper" Then graph = "breastplate8.bmp"

7 If Not reqwear1 = "" Then If Not wear1 = reqwear1 Then If Not wear1 = reqwear2 Then GoTo 3
If Not reqwear2 = "" Then If Not wear2 = "" Then If Not wear2 = reqwear1 Then If Not wear2 = reqwear2 Then GoTo 3

buyarmor.armor = basarm
buyarmor.name = randclothtype(basarm, r2, g2, b2, l2, gold, worth, , matnum) & " " & name
If Not wear2 = "" And Not wear2 = "Jacket" Then gold = gold / 2 'Int(Sqr(gold) + gold / 3)
'If Not wear2 = "" And Not wear2 = "Jacket" Then gold = (basarm / 2) ^ 2 * 10

If wear1 = "Jacket" Or wear1 = "Belt" Then gold = gold * 2
If Not reqwear1 = "" Then If Not wear1 = reqwear1 And Not wear1 = reqwear2 Then GoTo 3

'NOTE: Gold cost is now determined by Getworth
gold = gold * materialtypes(matnum).goldmult

'If wear1 = "Panties" Then graph = "cyberlegs3.bmp"
'If wear1 = "Bra" Then graph = "cyberarms3.bmp"
'If wear1 = "Panties" And wear2 = "Bra" Then graph = "cyberfull1.bmp"

'buyarmor.wear1 = wear1
'buyarmor.wear2 = wear2
'buyarmor.graph = graph
Cost = Int(gold)
cobjt.name = colorname & buyarmor.name

addeffect2 cobjt, "Clothes", graph, Int(basarm), wear1, wear2
addeffect2 cobjt, "GRAPH", "clothes.bmp", r, g, b

'addeffect2 cobjt, "Pickup"
cobjt.graphname = "clothes.bmp"
cobjt.r = r
cobjt.g = g
cobjt.b = b
cobjt.l = l

Cost = getworth(cobjt)

randclothes = cobjt


End Function

Function randclothtype(ByRef armor, ByRef r, ByRef g, ByRef b, ByRef l, ByRef gold, worth, Optional ByRef weight, Optional ByRef matnum)

'Find closest worth
acworth = materialtypes(1).worth
lastdiff = 100
For a = 1 To UBound(materialtypes())
    If diff(materialtypes(a).worth, worth) < lastdiff And materialtypes(a).type = "Cloth" Then acworth = a: lastdiff = diff(materialtypes(acworth).worth, worth)
Next a


4 aroll = roll(UBound(materialtypes()))
    aroll = acworth
    matnum = aroll
    If Not materialtypes(aroll).type = "Cloth" Then GoTo 4
    armult = materialtypes(aroll).armor: r = materialtypes(aroll).r: g = materialtypes(aroll).g: b = materialtypes(aroll).b:  armn = materialtypes(aroll).name
GoTo 7

armn = ""
armult = 1
colorstr = ""

'aroll = roll(8)
'If aroll = 1 Then colorstr = "Black": r = 55: g = 55: b = 55: l = 0.3
'If aroll = 2 Then colorstr = "Red": r = 255: g = 15: b = 15: l = 0.5
'If aroll = 3 Then colorstr = "Blue": r = 15: g = 15: b = 255: l = 0.5
'If aroll = 4 Then colorstr = "White": r = 255: g = 255: b = 255: l = 0.5
'If aroll = 5 Then colorstr = "Pink": r = 155: g = 15: b = 155: l = 1
'If aroll = 6 Then colorstr = "Purple": r = 105: g = 15: b = 205: l = 0.5
'If aroll = 7 Then colorstr = "Green": r = 25: g = 130: b = 0: l = 0.4
'If aroll = 8 Then colorstr = "Grey": r = 150: g = 150: b = 150: l = 0.4

'3     aroll = worth + roll(5) - 3
'If aroll > worth Then aroll = worth
'If worth > 8 Then aroll = 8 'aroll = Int(worth / 9)
'If aroll < 1 Then aroll = 1
'If aroll > 9 Or aroll < 1 Then GoTo 3
'If aroll = 1 Then armn = "Cloth": r = 255: g = 255: b = 255: l = 0.5: colorstr = "": armult = 1: weight = 1
'If aroll = 2 Then armn = "Hide": r = 105: g = 75: b = 0: l = 0.4: colorstr = "": armult = 0.5: weight = 1
'If aroll = 3 Then armn = "Cotton": armult = 2: weight = 1
'If aroll = 4 Then armn = "Leather": armult = 4: weight = 1
'If aroll = 5 Then armn = "Velvet": armult = 6: weight = 1
'If aroll = 6 Then armn = "Silk": armult = 8: weight = 1
'If aroll = 7 Then armn = "Magicweave": armult = 10: weight = 1
'If aroll = 8 Then armn = "Enchanted Silk": armult = 12: weight = 1


'If isexpansion = 0 Then
7 armor = armor * (1 + armult / 4): If armor < 1 Then armor = 1 Else armor = armor + armult
gold = (armor * armor) * 10 '* ((roll(10) + 10) / 10)) + (armor * 20)) / 2

'gold = gold * lesser(0.6, (1 - weight / 200)) '1/2% price reduction per weight, max 60%

randclothtype = colorstr & " " & armn

End Function


Function getobjtnum(name)
For a = 1 To objts
    If objtypes(a).name = name Then getobjtnum = a: Exit Function
Next a
End Function

'Function getobjnum(name)
'For a = 1 To objtotal
'    if objs(
'Next a
'End Function

Function getobjslot(obj As objecttype, eff As String)

For a = 1 To 6
    If obj.effect(a, 1) = eff Then getobjslot = a: Exit Function
Next a

End Function

Function getmontnum(name)
For a = 1 To objts
    If montype(a).name = name Then getmontnum = a: Exit Function
Next a
End Function

Function createobjtype(ByVal name, Optional graphic As String = "", Optional r = 255, Optional g = 0, Optional b = 0, Optional l = 0.5, Optional cells = 1)

objts = objts + 1: ReDim Preserve objtypes(1 To objts): ReDim Preserve objgraphs(1 To UBound(objtypes())) As cSpriteBitmaps
objtypes(objts).name = name
'If Not graphic = "" Then makesprite objtypes(objts).graph, Form1.Picture1, graphic, r, g, b, l, cells: objtypes(objts).graphloaded = 1
objtypes(objts).r = r
objtypes(objts).g = g
objtypes(objts).b = b
objtypes(objts).cells = cells
objtypes(objts).l = l
objtypes(objts).graphname = graphic
createobjtype = objts

End Function

Function makeobjtype(objtype As objecttype)

objts = objts + 1: ReDim Preserve objtypes(1 To objts): ReDim Preserve objgraphs(1 To UBound(objtypes())) As cSpriteBitmaps
If Not objtype.graphname = "" Then makesprite objgraphs(UBound(objtypes)), Form1.Picture1, objtype.graphname, objtype.r, objtype.g, objtype.b, objtype.l, 1: objtype.graphloaded = 1
objtypes(objts) = objtype
makeobjtype = objts
End Function

Function addeffect(ByVal wobj, ByVal eff1, Optional eff2 = "", Optional eff3 = "", Optional eff4 = "", Optional eff5 = "")

If Val(wobj) <= 1 Then

For a = 1 To objts
    If objtypes(a).name = wobj Then wobj = a: Exit For
Next a

Else: wobj = Val(wobj)

End If
5
For a = 1 To 10
    If objtypes(wobj).effect(a, 1) = "" Then
    objtypes(wobj).effect(a, 1) = eff1
    objtypes(wobj).effect(a, 2) = eff2
    objtypes(wobj).effect(a, 3) = eff3
    objtypes(wobj).effect(a, 4) = eff4
    objtypes(wobj).effect(a, 5) = eff5
    addeffect = a
    Exit For
    End If
Next a

'If a > UBound(objtypes(wobj).effect(), 2) Then ReDim Preserve objtypes(wobj).effect(1 To a, 1 To 5): GoTo 5

End Function

Function addeffect2(ByRef wobj As objecttype, ByVal eff1, Optional eff2 = "", Optional eff3 = "", Optional eff4 = "", Optional eff5 = "")
'Takes direct objects by ref

For a = 1 To 10
    
    'If wobj.effect(a, 1) = eff1 Then
    'wobj.effect(a, 2) = Val(wobj.effect(a, 2)) + Val(eff2)
    'If Not IsMissing(eff3) Then wobj.effect(a, 3) = eff3
    'If Not IsMissing(eff4) Then wobj.effect(a, 4) = eff4
    'If Not IsMissing(eff5) Then wobj.effect(a, 5) = eff5
    'addeffect2 = a
    'Exit For
    'End If
    
    If wobj.effect(a, 1) = "" Then
    wobj.effect(a, 1) = eff1
    If Not IsMissing(eff2) Then If Val(eff2) > 0 Then eff2 = Int(eff2)
    If Not IsMissing(eff2) Then wobj.effect(a, 2) = eff2
    If Not IsMissing(eff3) Then wobj.effect(a, 3) = eff3
    If Not IsMissing(eff4) Then wobj.effect(a, 4) = eff4
    If Not IsMissing(eff5) Then wobj.effect(a, 5) = eff5
    addeffect2 = a
    Exit For
    End If

Next a

'If a > UBound(wobj.effect(), 1) Then ReDim Preserve wobj.effect(1 To a, 1 To 5): GoTo 5


End Function

Function addeffect3(ByRef wobj As objecttype, ByVal eff1, Optional eff2 = "", Optional eff3 = "", Optional eff4 = "", Optional eff5 = "")
'Takes direct objects by ref, combines amounts of effects


For a = 1 To 10
    
    If wobj.effect(a, 1) = eff1 And Val(wobj.effect(a, 2)) > 0 Then
    wobj.effect(a, 2) = Val(wobj.effect(a, 2)) + Val(eff2)
    If Not IsMissing(eff3) Then wobj.effect(a, 3) = eff3
    If Not IsMissing(eff4) Then wobj.effect(a, 4) = eff4
    If Not IsMissing(eff5) Then wobj.effect(a, 5) = eff5
    addeffect3 = a
    Exit For
    End If
    
    If wobj.effect(a, 1) = "" Then
    wobj.effect(a, 1) = eff1
    If Not IsMissing(eff2) Then wobj.effect(a, 2) = eff2
    If Not IsMissing(eff3) Then wobj.effect(a, 3) = eff3
    If Not IsMissing(eff4) Then wobj.effect(a, 4) = eff4
    If Not IsMissing(eff5) Then wobj.effect(a, 5) = eff5
    addeffect3 = a
    Exit For
    End If

Next a

'If a > UBound(wobj.effect(), 2) Then ReDim Preserve wobj.effect(1 To a, 1 To 5): GoTo 5

End Function

Function cleareffs(wobj As objecttype)

'ReDim wobj.effect(1 To 1, 1 To 5)

For a = 1 To 10
    For b = 1 To 5
        wobj.effect(a, b) = ""
    Next b
Next a

End Function

Function geteff(obj As objecttype, name, pos)
'If UBound(obj.effect(), 1) = 0 Then Exit Function
'On Error Resume Next
For a = 1 To 10
    If obj.effect(a, 1) = name Then geteff = obj.effect(a, pos): Exit Function
Next a

End Function


Function randommonsters(ByVal mtype, ByVal num, Optional ByVal xz = 0, Optional ByVal yz = 0, Optional ByVal xz2 = 0, Optional ByVal yz2 = 0)
If xz = 0 Then xz = 6
If yz = 0 Then yz = 6
If xz2 = 0 Then xz2 = mapx - 6
If yz2 = 0 Then yz2 = mapy - 6

'If Not mtype = "ALL" Then mtype = Val(getmontnum(mtype))

For a = 1 To num
    If mtype = "ALL" Then wch = roll(lastmontype) Else wch = mtype
5     X = roll(xz2 - xz) + xz
    Y = roll(yz2 - yz) + yz
    If map(X, Y).blocked = 1 Or map(X, Y).tile = 0 Then GoTo 5
    If X < 8 Or X > mapx - 8 Or Y < 8 Or Y > mapy - 8 Then GoTo 5
    'Or map(x, y).used = 1
    createmonster wch, X, Y
Next a

End Function

Sub revamp()

allloaded = 0
For a = 1 To 4
    mapjunk.maps(a) = ""
Next a

'For a = 1 To lastmontype
'    montype(a).graph.ClearUp
'Next a

'For a = 1 To objts
'    objtypes(a).graph.ClearUp
'Next a

lastmontype = 0
totalmonsters = 0
clearsprites
ReDim montype(1 To 1) As monstertype: ReDim mongraphs(1 To UBound(montype())) As cSpriteBitmaps
ReDim objtypes(1 To 1) As objecttype: ReDim objgraphs(1 To UBound(objtypes())) As cSpriteBitmaps
'Erase montype()
'Erase objtypes()
ReDim mon(1 To 1) As amonsterT
'Erase mon()

Form1.Picture5.Visible = False
Form1.Text6.Visible = False

objtotal = 0
objts = 0
End Sub

Sub getobjdat()

Form4.Text1.text = ""

For a = 1 To objts
    Form4.addtxt "#OBJTYPE, " & objtypes(a).name & vbCrLf
    Form4.addtxt "#GRAPH, " & objtypes(a).graphname & ", " & objtypes(objts).cells & ", " & objtypes(a).r & ", " & objtypes(a).g & ", " & objtypes(a).b & ", " & objtypes(a).l & vbCrLf
    
    For b = 1 To 6
        If objtypes(a).effect(b, 1) = "" Then Exit For
        Form4.addtxt "#EFFECT, "
        For c = 1 To 5
             Form4.addtxt objtypes(a).effect(b, c) & ", "
        Next c
        Form4.addtxt vbCrLf
    Next b
    Form4.addtxt vbCrLf
Next a

Form4.addtxt vbCrLf

For a = 1 To objtotal
    Form4.addtxt "#CREATEOBJ, " & objtypes(objs(a).type).name & ", " & objs(a).X & ", " & objs(a).Y & ", " & objs(a).name & ", " & objs(a).string & ", " & swaptxt(swaptxt(objs(a).string2, Chr(34), "$"), ",", "/") & vbCrLf
Next a

Form4.Show

End Sub

Sub lwepgraph(fn, pb As PictureBox, Optional noupdt = 0)

If fn = "" Then wep.graphname = fn: If noupdt = 0 Then Form1.updatbody: Exit Sub Else Exit Sub
Set wepgraph = New cSpriteBitmaps


wep.graphname = fn
'rangecolor wep.r, wep.g, wep.b, Form1.Picture1, wep.l

'pb.AutoRedraw = True
'pb.Picture = LoadPicture(fn)
'rangecolor wep.r, wep.g, wep.b, pb, wep.l
'wepgraph.CreateFromPicture pb, 1, 1, , RGB(0, 0, 0)
makesprite wepgraph, pb, fn, wep.r, wep.g, wep.g ', wep.l
wasfound = 0
For a = 0 To pb.Width
For b = 0 To pb.Height
    mypoint = pb.Point(a, b)
    getrgb mypoint, r1, g1, b1
    If r1 < 5 And g1 > 250 And b1 > 250 Then wep.xoff = a: wep.yoff = b: pb.PSet (a, b), RGB(0, 0, 0): wasfound = 1: Exit For: Exit For
    If pb.Point(a, b) = RGB(0, 255, 255) Then Stop: wep.xoff = a: wep.yoff = b: pb.PSet (a, b), RGB(0, 0, 0): wasfound = 1: Exit For: Exit For
Next b
If wasfound = 1 Then Exit For
Next a

'wepgraph.CreateFromPicture pb, 1, 1

digwep
If noupdt = 0 Then Form1.updatbody

End Sub


Function randwep(Optional worth As Byte = 0, Optional ByRef Cost, Optional ByVal weptype = "") As objecttype

If worth = 0 Then worth = Int(averagelevel / 5)

Dim buyarmor As clothesT
3 aroll = roll(22)
'If aroll < 16 Then GoTo 3
'If Not isexpansion = 1 And aroll > 18 Then GoTo 3

wear1 = "6"
wear2 = ""
weight = 1

'New thingy
4 aroll = roll(UBound(weapontypes()))
If Not weapontypes(aroll).material = "Metal" Then GoTo 4
If Not weptype = "" Then If Not UCase(weapontypes(aroll).type) = UCase(weptype) Then GoTo 4
If Not weptype = "" Then GoTo 53 'Ignore power limits if this is a shop
If weapontypes(aroll).dice > worth + 8 Then GoTo 4
If weapontypes(aroll).dice * weapontypes(aroll).damage > worth * 15 + 20 Then GoTo 4 'Max damage cannot be more than worth * 15
53 basarm = weapontypes(aroll).dice: wear1 = weapontypes(aroll).damage: wear2 = weapontypes(aroll).type: graph = weapontypes(aroll).graph: weight = weapontypes(aroll).weight: name = weapontypes(aroll).name
GoTo 7

'This part isn't even used anymore, d00d...
If aroll = 1 Then name = "Sword": graph = "sword1.bmp": basarm = 4: weight = 8: wear2 = "Sword"
If aroll = 2 Then name = "Knife": graph = "dagger1.bmp": basarm = 3: wear1 = "3": wear2 = "Fast": weight = 1
If aroll = 3 Then name = "Chaos Blade": graph = "sword2.bmp": basarm = 3: wear1 = "12": wear2 = "Sword":: weight = 9: If worth < 4 Then GoTo 3
If aroll = 4 Then name = "Cleaver": graph = "cleaver1.bmp": basarm = 3: wear1 = "8": wear2 = "Axe": weight = 3: If worth < 3 Then GoTo 3
If aroll = 5 Then name = "Gemsword": graph = "Gemsword1.bmp": basarm = 6: wear1 = "12": weight = 6: If worth < 4 Then GoTo 3
If aroll = 6 Then name = "Mace": graph = "mace1.bmp": basarm = 3: weight = 6: wear2 = "Mace"
If aroll = 7 Then name = "Battleaxe": graph = "battleaxe1.bmp": basarm = 5: weight = 12
If aroll = 8 Then name = "Spear": graph = "Spear1.bmp": basarm = 2: wear1 = 8: wear2 = "Spear": weight = 8
If aroll = 9 Then name = "Fine Bow": graph = "bow1.bmp": basarm = 4: wear1 = 8: wear2 = "Bow": weight = 3
If aroll = 10 Then name = "Bow": graph = "bow2.bmp": basarm = 4: wear1 = 6: wear2 = "Bow": weight = 4
If aroll = 11 Then name = "Long Bow": graph = "bow3.bmp": basarm = 5: wear1 = 10: wear2 = "Bow": weight = 5
If aroll = 12 Then name = "Grand War Bow": graph = "bow4.bmp": basarm = 6: wear1 = 12: wear2 = "Bow": weight = 6: If worth < 4 Then GoTo 3
If aroll = 13 Then name = "Morningstar": graph = "flail1.bmp": basarm = 6: wear1 = "4": wear2 = "Mace": weight = 10
If aroll = 14 Then name = "Flail": graph = "flail2.bmp": basarm = 5: wear1 = "8": wear2 = "Mace": weight = 14
If aroll = 15 Then name = "Gemspear": graph = "Spear2.bmp": basarm = 2: wear1 = 20: wear2 = "Spear": weight = 12: If worth < 2 Then GoTo 3
If aroll = 16 Then name = "Winged Spear": graph = "Spear3.bmp": basarm = 3: wear1 = 24: wear2 = "Spear": weight = 14: If worth < 4 Then GoTo 3
If aroll = 17 Then name = "Katana": graph = "katana1.bmp": basarm = 3: wear1 = "20": wear2 = "Sword": weight = 9: If worth < 4 Then GoTo 3
If aroll = 18 Then name = "Demon Blade": graph = "demonblade1.bmp": basarm = 8: wear1 = "12": wear2 = "Sword": weight = 15: If worth < 6 Then GoTo 3
If aroll = 19 Then name = "Fine Sword": graph = "sword3.bmp": basarm = 6: weight = 6: wear2 = "Sword"
If aroll = 20 Then name = "Short Sword": graph = "dagger3.bmp": basarm = 3: wear1 = "6": wear2 = "Fast": weight = 5
If aroll = 21 Then name = "Dagger": graph = "dagger4.bmp": basarm = 3: wear1 = "4": wear2 = "Fast": weight = 2
If aroll = 22 Then name = "Claws": graph = "claws1.bmp": basarm = 3: wear1 = "5": wear2 = "Fast": weight = 2: If worth < 4 Then GoTo 3


 'If aroll = 3 Then name = "Breastplate": graph = "Breastplate3.bmp": basarm = 14: wear1 = "Upper": wear2 = "Jacket"
'If aroll = 4 Then name = "Halfplate": graph = "Breastplate4.bmp": basarm = 11: wear1 = "Upper"
'If aroll = 5 Then name = "Bracers": graph = "Bracers1.bmp": basarm = 4: wear1 = "Arms"
'If aroll >= 6 Then name = "Armored Skirt": graph = "armorskirt1.bmp": basarm = 7: wear1 = "Lower"

'buyarmor(slot).armor = basarm
'buyarmor(slot).name = randweptype(buyarmor(slot).armor, buyarmor(slot).r, buyarmor(slot).g, buyarmor(slot).b, buyarmor(slot).l, buyarmor(slot).gold, worth, wear1) & " " & name
'If wear2 = "Bow" Then buyarmor(slot).gold = buyarmor(slot).gold * 3
'buyarmor(slot).wear1 = wear1
'buyarmor(slot).wear2 = wear2
'buyarmor(slot).graph = graph
'randwep = buyarmor(slot).name

'weight = 1

7 buyarmor.armor = basarm
buyarmor.name = randweptype(basarm, r, g, b, l, gold, worth, wear1, weight, matnum) & " " & name
'If Not wear2 = "" Then gold = gold * 0.75
'If wear1 = "Jacket" Then gold = gold * 2
'If Not reqwear1 = "" Then If Not wear1 = reqwear1 And Not wear1 = reqwear2 Then GoTo 3
'If name = "Halter" Then gold = gold / 2
'If wear1 = "Arms" Then gold = gold * 2

'gold = (basarm * wear1) * 10 * (1 + (basarm * wear1) / 100)

'It ignores this now, too.  It uses getworth at the end instead.
gold = (basarm * wear1 - weight * (basarm / 6)) * (1 + (basarm * wear1) / 10)
gold = gold * materialtypes(matnum).goldmult
gold = gold * 10
'If gold < 0 Then Stop

'gold = gold * lesser(0.6, (1 - weight / 200)) '1/2% price reduction per weight, max 60%

If wear2 = "Bow" Then gold = gold * 1.5

Cost = Int(gold)
'buyarmor.wear1 = wear1
'buyarmor.wear2 = wear2
'buyarmor.graph = graph
Dim cobjt As objecttype
cobjt.graphname = "sword1.bmp"
cobjt.name = colorname & buyarmor.name
addeffect2 cobjt, "Weapon", graph, Int(basarm), wear1, wear2
addeffect2 cobjt, "Equipjunk", weight
'addeffect2 cobjt, "GRAPH", "clothes.bmp", r, g, b
'addeffect2 cobjt, "Pickup"
cobjt.r = r
cobjt.g = g
cobjt.b = b
cobjt.l = l
Cost = getworth(cobjt)
randwep = cobjt

End Function

Function randweptype(ByRef armor, ByRef r, ByRef g, ByRef b, ByRef l, ByRef gold, worth, ByRef wear1, Optional ByRef weight = 1, Optional ByRef matnum)

'Find closest worth
acworth = materialtypes(1).worth
lastdiff = 100
For a = 1 To UBound(materialtypes())
    If diff(materialtypes(a).worth, worth) < lastdiff And materialtypes(a).type = "Metal" Then acworth = a: lastdiff = diff(materialtypes(acworth).worth, worth)
Next a

4 aroll = roll(UBound(materialtypes()))
    aroll = acworth
    matnum = aroll
    If Not materialtypes(aroll).type = "Metal" Then GoTo 4
    armult = materialtypes(aroll).armor: r = materialtypes(aroll).r: g = materialtypes(aroll).g: b = materialtypes(aroll).b: weight = weight * materialtypes(aroll).weight: armn = materialtypes(aroll).name

GoTo 7

armn = ""
armult = 1

3 aroll = worth + roll(5) - 3
If aroll < 1 Then GoTo 3
If aroll > 11 Then GoTo 3
If worth > 11 Then aroll = Int(worth / 11)
If aroll = 1 Then armn = "Bronze": r = 55: g = 35: b = 0: l = 0.3: armult = 0.5: weight = weight * 1.2
If aroll = 2 Then armn = "Iron": r = 105: g = 105: b = 125: l = 0.2: armult = 1: weight = weight * 1.1
If aroll = 3 Then armn = "Steel": r = 155: g = 155: b = 155: l = 0.7: armult = 2
If aroll = 4 Then armn = "PinkSteel": r = 255: g = 105: b = 240: l = 0.5: armult = 3: weight = weight * 1.1
If aroll = 5 Then armn = "Venomous": r = 20: g = 105: b = 0: l = 0.5: armult = 4
If aroll = 6 Then armn = "Enchanted Obsidian": r = 155: g = 155: b = 155: l = 0: armult = 5: weight = weight * 0.8
If aroll = 7 Then armn = "BloodSteel": r = 255: g = 0: b = 0: l = 0.3: armult = 6: weight = weight * 1.3
If aroll = 8 Then armn = "Blazing": r = 250: g = 150: b = 30: l = 0.6: armult = 8: wear1 = wear1 + 1: weight = weight * 1.1
If aroll = 9 Then armn = "Blue Adamant": r = 0: g = 0: b = 255: l = 0.4: armult = 10: weight = weight * 1.4
If aroll = 10 Then armn = "Angelic": r = 250: g = 250: b = 200: l = 0.5: armult = 12: wear1 = wear1 + 2: weight = weight * 1.2
If aroll = 11 Then armn = "Demonic": r = 120: g = 15: b = 3: l = 0.3: armult = 13: wear1 = wear1 + 4: weight = weight * 1.5
wear1 = wear1 + aroll

'wear1 = wear1 + Int(armult / 2)
7 If r < 1 Then r = 3: If g < 1 Then g = 3: If b < 1 Then b = 3

'If isexpansion = 1 Then
armor = armor * armult
If armor < 1 Then armor = 1
'wear1 = wear1 + armult * 2: armor = armor + (armult / 4) 'Else armor = armor * armult: If armor < 1 Then armor = 1
'gold = (((armor * armor) * 10 * ((roll(10) + 10) / 10)) + (armor * 20)) / 2 * Val(wear1)
'gold = (armor * wear1) * (20 * armult)

randweptype = armn

End Function

Function ifquest(str As String, Optional addq = 0)

ifquest = False

For a = 1 To 500
    If plr.quest(a) = str Then ifquest = True: Exit For
    If plr.quest(a) = "" And addq = 1 Then plr.quest(a) = str: ifquest = True: Exit For
Next a

End Function

Function getdig(dmg)

mname = montype(mon(plr.instomach).type).name

3 aroll = roll(18)

If aroll = 1 Then getdigs = mname & "'s digestive tract squeezes and grinds as it tries to absorb it's meal--you."
If aroll = 2 Then getdigs = "You churn and slop around inside " & mname & "'s hot depths."
If aroll = 3 Then getdigs = "You are being digested alive! " & mname & "'s stomach acids are burning your skin."
If aroll = 4 Then GoTo 3 'getdigs = "You take " & dmg & " points of damage from the churning stomach walls."
If aroll = 5 Then getdigs = "The strong acid inside " & mname & "'s stomach is searing your flesh away. The pain is immense."
If aroll = 6 Then getdigs = "The stinging digestive juices inside " & mname & "'s belly are taking their toll."
If aroll = 7 Then getdigs = "You continue to slime around in " & mname & "'s hot stomach." & vbCrLf & "You squirm as you futilely try to resist it's body's painful attempts to digest you alive."
If aroll = 8 Then getdigs = "Though you continue to struggle, you cannot seem to escape " & mname & "'s stomach." & vbCrLf & "Digestive liquids continue to sting your skin as you writhe in it's belly."
If aroll = 9 Then getdigs = "You are practically swimming in churning digestive juices." & vbCrLf & "You try to escape, but you cannot force your way out."
If aroll = 10 Then getdigs = "You squirm and writhe and try to escape, but " & mname & "'s hot stomach is too tight to get out of." & vbCrLf & "You continue to struggle as the burning digestive juices soak your skin."
If aroll = 11 Then getdigs = "You are squeezed through " & mname & "'s digestive tract, the stomach acid burning your skin in the pitch darkness."
If aroll = 12 Then getdigs = mname & " is digesting you alive!!"
If aroll = 13 Then getdigs = "Ooze and chunks soak your skin as you writhe inside " & mname & "'s stomach. You struggle for air as it's body tries to digest you alive."
If aroll = 14 Then getdigs = "You are being digested!!"
If aroll = 15 Then getdigs = "You are trapped inside " & mname & "'s body. Digestive liquids soak your skin as you slop through it's gastrointestinal tract."
If aroll = 16 Then getdigs = "You are inside " & mname & "'s stomach! You'll be nothing more than a soft lump pushing through it's intestines if you don't escape soon!"
If aroll = 17 Then getdigs = "It's tummy gurgles as it continues to digest. All you can feel is slimy flesh all around you, shifting and oozing as you slither around in it's churning stomach."
If aroll = 18 Then getdigs = "You squirm and shove, knowing that escape is your only hope to avoid digestion, but you cannot push yourself out of it's belly."

'getdigs = getdigs & vbCrLf & "You lose " & dmg & " hit points."
getdig = getdigs


End Function




Function savebindataold(fn)
'killsummons
'fixfuckingstrings

If Dir(plr.name, vbDirectory) = "" Then MkDir plr.name
fn = ".\" & plr.name & "\" & fn
If Not Dir(fn) = "" Then Kill fn
Open fn For Output As #1

Write #1, objts
Write #1, objtotal
Write #1, totalmonsters
Write #1, lastmontype
Write #1, mapx
Write #1, mapy

Write #1, mapjunk.maps(1)
Write #1, mapjunk.maps(2)
Write #1, mapjunk.maps(3)
Write #1, mapjunk.maps(4)
Write #1, mapjunk.terstr
Write #1, mapjunk.favmonster
Write #1, mapjunk.favmonsternum
Write #1, mapjunk.questmonster

'Put array resizes here!!

For a = 1 To mapx
For b = 1 To mapy
    Write #1, map(a, b).monster
    Write #1, map(a, b).object
    Write #1, map(a, b).blocked
    Write #1, map(a, b).tile
    Write #1, map(a, b).ovrtile
Next b
Next a

'write #1, map() '.object

'write #1, objs()
For a = 1 To objts
    Write #1, objtypes(a).r
    Write #1, objtypes(a).g
    Write #1, objtypes(a).b
    Write #1, objtypes(a).cells
    Write #1, objtypes(a).graphname
    Write #1, objtypes(a).l
    Write #1, objtypes(a).name
    For b = 1 To 6
    For c = 1 To 5
        Write #1, objtypes(a).effect(b, c)
    Next c
    Next b
Next a

For a = 1 To objtotal
    Write #1, objs(a).name
    Write #1, objs(a).string
    Write #1, objs(a).string2
    Write #1, objs(a).type
    Write #1, objs(a).X
    Write #1, objs(a).Y
    Write #1, objs(a).xoff
    Write #1, objs(a).yoff
Next a

For a = 1 To lastmontype
    Write #1, montype(a).acid
    Write #1, montype(a).color
    Write #1, montype(a).damage
    Write #1, montype(a).dice
    Write #1, montype(a).eatskill
    Write #1, montype(a).escapediff
    Write #1, montype(a).exp
    Write #1, montype(a).gfile
    Write #1, montype(a).gnum
    Write #1, montype(a).hp
    Write #1, montype(a).level
    Write #1, montype(a).move
    Write #1, montype(a).name
    Write #1, montype(a).skill
    Write #1, montype(a).light
Next a

For a = 1 To totalmonsters
    Write #1, mon(a).cell
    Write #1, mon(a).hp
    Write #1, mon(a).owner
    Write #1, mon(a).type
    Write #1, mon(a).X
    Write #1, mon(a).Y
    Write #1, mon(a).instomach
Next a

'monc = countmonsters
'Stop
'write #1, monc
'For a = 1 To monc
'    If mon(a).hp > 0 Then write #1, mon(a)
'Next a

Close #1

End Function

Function savebindata(fn)
'killsummons
'fixfuckingstrings

ChDir App.Path

If Left(fn, 6) = "VTDATA" Then fn = Right(fn, Len(fn) - 6)

'If Dir(App.Path & "\" & plr.name, vbDirectory) = "" Then MkDir plr.name
'fn = App.Path & "\" & plr.name & "\" & fn
If Not Dir(fn) = "" Then Kill fn
fnum = FreeFile
Open fn For Binary As fnum

Put fnum, , objts
Put fnum, , objtotal
Put fnum, , totalmonsters
Put fnum, , lastmontype
Put fnum, , mapx
Put fnum, , mapy

Put fnum, , mapjunk
Put fnum, , map()

Put fnum, , objtypes()
Put fnum, , objs()
Put fnum, , montype()
Put fnum, , mon()

Close fnum

addfile fn, "curgame.dat", 1
Kill fn

End Function

Function loadbindata(ByVal fn, Optional ByVal datfile As String = "curgame.dat")
'killsummons
'fixfuckingstrings

'Form1.Timer1.Enabled = False
'Form1.Timer2.Enabled = False

ChDir App.Path

If Not datfile = "" Then fn = getfile(fn, datfile, 1) Else fn = getfile(fn, datfile, , , 1)

If fn = "" Then loadbindata = False: Exit Function

Form1.Text5.Visible = True
Form1.Text5.Refresh
'If Dir(App.Path & "\" & plr.name & "\" & fn) = "" And hasbeento(fn) = True And cheaton = 0 Then MsgBox "Map data file not found.": End
'If Dir(App.Path & "\" & plr.name & "\" & fn) = "" Then loadbindata = False: Exit Function
'If Dir(App.Path & "\" & plr.name, vbDirectory) = "" Then MkDir plr.name
'fn = App.Path & "\" & plr.name & "\" & fn
Open fn For Binary As #1

Get #1, , objts
Get #1, , objtotal
Get #1, , totalmonsters
Get #1, , lastmontype
Get #1, , mapx
Get #1, , mapy

Get #1, , mapjunk
ReDim map(1 To mapx, 1 To mapy)
Get #1, , map()

gamemsg mapjunk.terstr

ReDim objtypes(1 To objts)
Get #1, , objtypes()
If objtotal > 0 Then ReDim objs(1 To objtotal)
Get #1, , objs()
ReDim montype(1 To greater(lastmontype, 1))
Get #1, , montype()
ReDim mon(1 To greater(totalmonsters, 1))
Get #1, , mon()

ReDim objgraphs(1 To UBound(objtypes())) As cSpriteBitmaps
For a = 1 To objts
    objtypes(a).graphloaded = 0
    'If Not objtypes(objs(a).type).graphname = "" Then makesprite objgraphs(objs(a).type), Form1.Picture1, objtypes(objs(a).type).graphname, objtypes(objs(a).type).r, objtypes(objs(a).type).g, objtypes(objs(a).type).b, objtypes(objs(a).type).l, objtypes(objs(a).type).cells: objtypes(objs(a).type).graphloaded = 1
Next a

ReDim mongraphs(1 To UBound(montype())) As cSpriteBitmaps
For a = 1 To lastmontype
    getrgb montype(a).color, r, g, b, 32
    If Not montype(a).gfile = "" Then makesprite mongraphs(a), Form1.Picture1, montype(a).gfile, r, g, b, montype(a).light, 3
Next a

Close #1

loadextraovrs mapjunk.name1, mapjunk.name2, mapjunk.name3, mapjunk.name4

'Form1.Timer1.Enabled = True
'Form1.Timer2.Enabled = True

'If Not datfile = "" Then Kill fn

Form1.Text5.Visible = False

loadbindata = True

setupquests
'putcarrymonsters

End Function

Function loadbindataold(fn) As Boolean
'killsummons

If Dir(plr.name, vbDirectory) = "" Then MkDir plr.name
fn = ".\" & plr.name & "\" & fn
If Dir(fn) = "" Then loadbindataold = False: Exit Function

'Form1.Timer1.Enabled = False
'Form1.Timer2.Enabled = False

Form1.Text5.Visible = True
Form1.Text5.Refresh

Open fn For Input As #1

Input #1, objts
Input #1, objtotal
Input #1, totalmonsters
Input #1, lastmontype
Input #1, mapx
Input #1, mapy

Input #1, mapjunk.maps(1)
Input #1, mapjunk.maps(2)
Input #1, mapjunk.maps(3)
Input #1, mapjunk.maps(4)
Input #1, mapjunk.terstr: gamemsg mapjunk.terstr
Input #1, mapjunk.favmonster
Input #1, mapjunk.favmonsternum
Input #1, mapjunk.questmonster

ReDim map(1 To mapx, 1 To mapy) As tiletype
If objts > 0 Then ReDim objtypes(1 To objts) As objecttype: ReDim objgraphs(1 To UBound(objtypes())) As cSpriteBitmaps
If objtotal > 0 Then ReDim objs(1 To objtotal) As aobject
If totalmonsters > 0 Then ReDim mon(1 To totalmonsters) As amonsterT
If lastmontype > 0 Then ReDim montype(1 To lastmontype) As monstertype: ReDim mongraphs(1 To UBound(montype())) As cSpriteBitmaps



For a = 1 To mapx
For b = 1 To mapy

    Input #1, map(a, b).monster
    Input #1, map(a, b).object ': If map(a, b).object > 0
    Input #1, map(a, b).blocked
    Input #1, map(a, b).tile
    Input #1, map(a, b).ovrtile
Next b
Next a

'input #1, map() '.object


'input #1, objs()
For a = 1 To objts
    Input #1, objtypes(a).r
    Input #1, objtypes(a).g
    Input #1, objtypes(a).b
    Input #1, objtypes(a).cells
    Input #1, objtypes(a).graphname
    Input #1, objtypes(a).l
    Input #1, objtypes(a).name
    For b = 1 To 6
    For c = 1 To 5
        Input #1, objtypes(a).effect(b, c)
    Next c
    Next b
    
    'If Not objtypes(a).graphname = "" Then makesprite objtypes(a).graph, Form1.Picture1, objtypes(a).graphname, objtypes(a).r, objtypes(a).g, objtypes(a).b, objtypes(a).l, objtypes(a).cells: objtypes(a).graphloaded = 1
'    makesprite montype(a).graph, Form1.Picture1, montype(a).gfile, r, g, b, montype(a).light, 3
    
    
Next a


For a = 1 To objtotal
    Input #1, objs(a).name
    Input #1, objs(a).string
    Input #1, objs(a).string2
    Input #1, objs(a).type
    Input #1, objs(a).X
    Input #1, objs(a).Y
    Input #1, objs(a).xoff
    Input #1, objs(a).yoff
    
    If objtypes(objs(a).type).graphloaded = 0 Then If Not objtypes(objs(a).type).graphname = "" Then makesprite objgraphs(objs(a).type), Form1.Picture1, objtypes(objs(a).type).graphname, objtypes(objs(a).type).r, objtypes(objs(a).type).g, objtypes(objs(a).type).b, objtypes(objs(a).type).l, objtypes(objs(a).type).cells: objtypes(objs(a).type).graphloaded = 1
    
Next a


For a = 1 To lastmontype
    Input #1, montype(a).acid
    Input #1, montype(a).color
    Input #1, montype(a).damage
    Input #1, montype(a).dice
    Input #1, montype(a).eatskill
    Input #1, montype(a).escapediff
    Input #1, montype(a).exp
    Input #1, montype(a).gfile
    Input #1, montype(a).gnum
    Input #1, montype(a).hp
    Input #1, montype(a).level
    Input #1, montype(a).move
    Input #1, montype(a).name
    Input #1, montype(a).skill
    Input #1, montype(a).light
    
    r = 0: g = 0: b = 0
    getrgb montype(a).color, r, g, b
    If Not montype(a).gfile = "" Then makesprite mongraphs(a), Form1.Picture1, montype(a).gfile, r, g, b, montype(a).light, 3

Next a


For a = 1 To totalmonsters
    Input #1, mon(a).cell
    Input #1, mon(a).hp
    Input #1, mon(a).owner
    Input #1, mon(a).type
    Input #1, mon(a).X
    Input #1, mon(a).Y
    Input #1, mon(a).instomach
Next a


'monc = countmonsters
'Stop
'input #1, monc
'For a = 1 To monc
'    If mon(a).hp > 0 Then input #1, mon(a)
'Next a

Close #1

loadbindataold = True
Form1.Text5.Visible = False

'Form1.Timer1.Enabled = True
'Form1.Timer2.Enabled = True

End Function

Function gotomap(ByVal fn)
'm'' quitting an area and entering a neighbor area
'm'' performance to improve. Looks like it's the recoloring that is slow
Dim a As Long 'm''

'monlev = checklevels(fn)
'If monlev > plr.level Then If MsgBox("The lowest level monster in that area is level " & monlev & ".  You are level " & plr.level & ".  ie you really don't want to go there.  Do you want to be a 'tard and go there anyway?", vbYesNo) = vbNo Then Exit Function
Ccom = ""

Do While plr.foodinbelly > 0
    For a = 1 To UBound(mon())
        If mon(a).instomach = -1 Then plrdigmon mon(a), a, 1
    Next a
    digestfood
Loop

#If USELEGACY = 0 Then 'm'' added cosmetic "loading" screen early
    Form1.Text5.Visible = True: DoEvents
#End If 'm''

ChDir App.Path


#If USELEGACY = 1 Then 'm'' original order of map swtiching ...

If Not Dir("autosave.plr") = "" Then Kill "autosave.plr"

getcarrymonsters
savebindata Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"
fsavechar "plrdat.tmp" 'FileD.FileTitle

addfile "plrdat.tmp", "autosave.plr", 1
addfile "curgame.dat", "autosave.plr", 1

#Else 'm'' new order, fixing mercernaries not saved in autosave game event

If Not Dir("autosave.plr") = "" Then Kill "autosave.plr"

'm'' sadly this is much more tricky than that :
'm'' the autosave saves the whole map, including the merc
'm'' when you enter back the map, it will load the merc, AND adding the merc that
'm'' are following you from previous map, thus duplicating them...
getcarrymonsters
savebindata Left$(plr.curmap, Len(plr.curmap) - 4) & ".dat"
fsavechar "plrdat.tmp" 'FileD.FileTitle

addfile "plrdat.tmp", "autosave.plr", 1
addfile "curgame.dat", "autosave.plr", 1



#End If

clearboosts
clearsprites

'Delete the data file after it's been added to the player's pak file
'Kill Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"
'fsavechar "autosave.plr"
'savebindata Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"
'savegame "autosave.plr"

revamp




loadextraovrs "tree2.bmp", "tree5.bmp", "tree1.bmp", "thorntree1.bmp"

'Will only load text data if binary data unavailable for that map
plr.curmap = fn
If loadbindata(Left(fn, Len(fn) - 4) & ".dat", "curgame.dat") = False Then loaddata fn

'loadbindata Left(fn, Len(plr.curmap) - 4) & ".dat"
clearsides
If totalmonsters > (mapx + mapy / 4) And Not mapjunk.favmonster = "" Then randommonsters mapjunk.favmonster, mapjunk.favmonsternum
plr.sp = plr.spmax


walloffmap
'putcarrymonsters
makeminimap



'Form1.Timer1.Enabled = True
'Form1.Timer2.Enabled = True

End Function

Function getitem(obj As objecttype)
'orginv
addtext obj.name, , , 250, 250
For a = 50 To 1 Step -1
    If a = 1 And inv(a).name = "" Then inv(a) = obj: Exit For
    If a > 1 Then If inv(a).name = "" And Not inv(a - 1).name = "" Then inv(a) = obj: Exit For
Next a
If a = 0 Then addtext "You cannot carry anymore.": dropitem obj
'If Not inv(50).name = "" Then gamemsg "You cannot carry anymore.": dropitem inv(50)
Form1.updatinv

End Function

Function killitem(ByRef obj As objecttype)
obj.name = ""
obj.graphloaded = 0
Erase obj.effect
Form1.updatinv
End Function

Function sprinkleovr(ByVal num, ByVal tt)

For a = 1 To num
3   X = roll(mapx)
    Y = roll(mapy)
    If map(X, Y).used = 1 Then GoTo 3
    map(X, Y).ovrtile = tt: map(X, Y).blocked = 1: If tt = 5 Or tt = 17 Or tt = 20 Then map(X, Y).blocked = 0
Next a

End Function

Function sprinkle(ByVal num, ByVal tt)

For a = 1 To num
3   X = roll(mapx)
    Y = roll(mapy)
    If Not map(X, Y).used = 1 Then map(X, Y).tile = tt Else GoTo 3
Next a

End Function

Function getesc()

mname = montype(mon(plr.instomach).type).name
getesc = getgener("You slither your way out of " & mname & "'s mouth to freedom.", "With a spray of saliva, you escape " & mname & "'s hungry maw.")
Exit Function

mname = montype(mon(plr.instomach).type).name
aroll = roll(6)

If aroll = 1 Then bstr = "You shove and squirm with all your might and force your body out of " & mname & "'s stomach!"
If aroll = 2 Then bstr = "You desperately claw your way back up " & mname & "'s throat and slop to the ground, still covered in it's oozing digestive juices."
If aroll = 3 Then bstr = "With a spray of stomach fluid, you violently force your way out of " & mname & "'s stomach."
If aroll = 4 Then bstr = "Your skin oozes with digestive liquids as you claw your way out of " & mname & "'s belly."
If aroll = 5 Then bstr = "You fight the terrible stinging as you pull yourself to freedom!"
If aroll = 6 Then bstr = "You forcefully pull yourself out of " & mname & "'s body. Stomach liquids drip from your body and soak the dirt beneath you."

getesc = bstr

End Function

Function getswallow(monnum)

swallowtype = montype(mon(monnum).type).swallow

mname = "the " & montype(mon(monnum).type).name
If swallowtype = "Plant" Then GoTo 5
If swallowtype = "Slime" Then GoTo 8
aroll = roll(6)

If aroll = 1 Then bstr = "You reel from " & mname & "'s attack and are disoriented for a moment. Suddenly, it's maw wraps around you and strong muscles pull you down it's throat! You quickly slither into it's hot, churning stomach, and immediately your skin begins to burn."
If aroll = 2 Then bstr = mname & " opens it's maw wide just before it strikes! Within moments, you are being squeezed into it's gastrointestinal tract!"
If aroll = 3 Then bstr = mname & "'s attack drags you painfully towards it's gaping mouth. Though you resist with all your might, it slowly forces you down it's gullet."
If aroll = 4 Then bstr = "While you struggle as much as you can, " & mname & "'s strength gets the best of you. It drags you into it's gaping mouth and swallows hard. A great bulge travelling down it's body is all that marks your passage into it's stomach."
If aroll = 5 Then bstr = "You get a short view of " & mname & "'s throat opening before you until you are abruptly shoved into it. You slide down it easily and find yourself inside " & mname & "'s stomach."
If aroll = 6 Then bstr = mname & " catches you by surprise and picks you up in it's giant jaws. Before you can do anything, your shoulders are being dragged down it's throat. The rest of your body soon follows them into " & mname & "'s digestive tract."

bstr = bstr & vbCrLf & "You have been swallowed whole!"
getswallow = bstr
Exit Function
5 aroll = roll(3)

If aroll = 1 Then bstr = "The plant wraps it's green flesh around your body and clamps down tight. Within moments you feel a painful sting as digestive secretions begin to ooze down your skin."
If aroll = 2 Then bstr = mname & "'s petals wrap tightly around your skin and you suddenly find yourself unable to move! Slimy digestive liquids immediately start oozing out of " & mname & " in preparation to digest it's hearty meal."
If aroll = 3 Then bstr = "The strong appendages of " & mname & " suddenly surge forward and snap shut around you! Your body is suddenly squeezed tightly by the plant, which immediately sets about the process of dissolving your body."
getswallow = bstr
Exit Function

8 aroll = roll(3)

If aroll = 1 Then bstr = "The oozing mass suddenly surges forward and engulfs you! Your whole body is now trapped inside the " & mname & "!"
If aroll = 2 Then bstr = "The gooey " & mname & " suddenly envelops you. The suffocating gook immediately begins to dissolve your skin."
If aroll = 3 Then bstr = "You are suddenly sucked into the slimy mass of the " & mname & "! It is trying to dissolve your body from all sides."
getswallow = bstr
Exit Function

End Function

Sub loadtreasuretypes()

F = createobjtype("Treasure Bag (Orange)", "treasure2.bmp", 200, 140, 0, 0.5)
addeffect F, "GiveGold", "25"
addeffect F, "Destruct"

F = createobjtype("Treasure Bag (Green)", "treasure2.bmp", 0, 180, 0, 0.5)
addeffect F, "GiveGold", "50"
addeffect F, "Destruct"

F = createobjtype("Treasure Bag (Red)", "treasure2.bmp", 230, 0, 0, 0.5)
addeffect F, "GiveGold", "100"
addeffect F, "Destruct"

F = createobjtype("Treasure Bag (Blue)", "treasure2.bmp", 0, 0, 255, 0.5)
addeffect F, "GiveGold", "200"
addeffect F, "Destruct"

F = createobjtype("Treasure Bag (Black)", "treasure2.bmp", 55, 55, 55, 0.3)
addeffect F, "GiveGold", "500"
addeffect F, "Destruct"

F = createobjtype("Treasure Chest (Orange)", "treasure1.bmp", 200, 140, 0, 0.5)
addeffect F, "GiveGold", "100"
addeffect F, "Destruct"

F = createobjtype("Treasure Chest (Green)", "treasure1.bmp", 0, 180, 0, 0.5)
addeffect F, "GiveGold", "250"
addeffect F, "Destruct"

F = createobjtype("Treasure Chest (Red)", "treasure1.bmp", 230, 0, 0, 0.5)
addeffect F, "GiveGold", "500"
addeffect F, "Destruct"

F = createobjtype("Treasure Chest (Blue)", "treasure1.bmp", 0, 0, 255, 0.5)
addeffect F, "GiveGold", "1000"
addeffect F, "Destruct"

F = createobjtype("Treasure Chest (Black)", "treasure1.bmp", 55, 55, 55, 0.3)
addeffect F, "GiveGold", "2000"
addeffect F, "Destruct"

F = createobjtype("Lesser Healing Potion", "potion1.bmp", 255, 0, 255, 0.4)
'addeffect f, "Pickup"
addeffect F, "Lifepotion", "1"
addeffect F, "Destruct"

F = createobjtype("Healing Potion", "potion1.bmp", 255, 30, 255, 0.5)
'addeffect f, "Pickup"
addeffect F, "Lifepotion", "2"
addeffect F, "Destruct"

F = createobjtype("Greater Healing Potion", "potion1.bmp", 255, 30, 255, 0.7)
'addeffect f, "Pickup"
addeffect F, "Lifepotion", "3"
addeffect F, "Destruct"

F = createobjtype("Full Healing Potion", "potion1.bmp", 255, 160, 255, 1)
addeffect F, "Pickup"
addeffect F, "Heal", "10000"
addeffect F, "Destruct"

F = createobjtype("Fountain of Healing", "Fountain1.bmp", 255, 160, 255, 1)
addeffect F, "Heal", "10000"
addeffect F, "Givemp", "10000"
addeffect F, "Backup"
addeffect F, "NoEat"

F = createobjtype("Potion of Strength (Permanent)", "potion1.bmp", 255, 255, 0, 1)
addeffect F, "Pickup"
addeffect F, "Givestr", "1"
addeffect F, "Destruct"

F = createobjtype("Potion of Dexterity (Permanent)", "potion1.bmp", 155, , 155, 0.3)
addeffect F, "Pickup"
addeffect F, "Givedex", "1"
addeffect F, "Destruct"

F = createobjtype("Potion of Intelligence (Permanent)", "potion1.bmp", 0, 0, 255, 0.5)
addeffect F, "Pickup"
addeffect F, "Giveint", "1"
addeffect F, "Destruct"

F = createobjtype("Potion of Experience", "potion1.bmp", 255, 155, 0, 0.3)
addeffect F, "Pickup"
addeffect F, "Giveexp", "1000"
addeffect F, "Destruct"

F = createobjtype("Greater Potion of Experience", "potion1.bmp", 255, 155, 0, 0.6)
addeffect F, "Pickup"
addeffect F, "Giveexp", "10000"
addeffect F, "Destruct"

F = createobjtype("Mega Potion of Experience", "potion1.bmp", 255, 155, 0, 1)
addeffect F, "Pickup"
addeffect F, "Giveexp", "100000"
addeffect F, "Destruct"

F = createobjtype("Lesser Mana Potion", "potion1.bmp", 0, 0, 255, 0.3)
'addeffect f, "Pickup"
addeffect F, "Manapotion", "1"
addeffect F, "Destruct"

F = createobjtype("Mana Potion", "potion1.bmp", 0, 0, 255, 0.5)
'addeffect f, "Pickup"
addeffect F, "Manapotion", "2"
addeffect F, "Destruct"

F = createobjtype("Greater Mana Potion", "potion1.bmp", 0, 0, 255, 0.7)
'addeffect f, "Pickup"
addeffect F, "Manapotion", "3"
addeffect F, "Destruct"

createobjtype "Shit", "poo1small.bmp"

F = createobjtype("Blaze Scroll", "scroll.bmp", 250, 150, 0, 0.5)
addeffect F, "Pickup"
addeffect F, "Spell", "Blaze"
addeffect F, "Destruct"

F = createobjtype("Teleportation Scroll", "scroll.bmp", 150, 150, 150, 0.5)
addeffect F, "Pickup"
addeffect F, "Spell", "Teleport"
addeffect F, "Destruct"

F = createobjtype("Lightning Storm Scroll", "scroll.bmp", 0, 150, 250, 0.5)
addeffect F, "Pickup"
addeffect F, "Spell", "Lightning Storm"
addeffect F, "Destruct"

F = createobjtype("Supernova Scroll", "scroll.bmp", 250, 150, 0, 0.5)
addeffect F, "Pickup"
addeffect F, "Spell", "Supernova"
addeffect F, "Destruct"

F = createobjtype("Greater Strength Scroll", "scroll.bmp", 250, 150, 0, 0.5)
addeffect F, "Pickup"
addeffect F, "Spell", "Greater Strength"
addeffect F, "Destruct"

F = createobjtype("Greater Dexterity Scroll", "scroll.bmp", 150, 0, 250, 0.5)
addeffect F, "Pickup"
addeffect F, "Spell", "Greater Dexterity"
addeffect F, "Destruct"

F = createobjtype("All-Dye Scroll", "scroll.bmp", 150, 0, 250, 0.5)
addeffect F, "Pickup"
addeffect F, "Spell", "All-Dye"
addeffect F, "Destruct"

F = createobjtype("Small Gold Pile", "gold4.bmp")
addeffect F, "Mongold"
addeffect F, "Destruct"

F = createobjtype("Medium Gold Pile", "gold3.bmp")
addeffect F, "Mongold"
addeffect F, "Destruct"

F = createobjtype("Large Gold Pile", "gold2.bmp")
addeffect F, "Mongold"
addeffect F, "Destruct"

F = createobjtype("Huge Gold Pile", "gold1.bmp")
addeffect F, "Mongold"
addeffect F, "Destruct"

End Sub

Function loadsummons(name, level, filen, r, g, b, l, Optional boss = 0)
'Static isloaded As Byte
'If isloaded = 1 Then Exit Function
lastmontype = lastmontype + 1
ReDim Preserve montype(1 To lastmontype): ReDim Preserve mongraphs(1 To UBound(montype())) As cSpriteBitmaps
montype(lastmontype).gfile = filen
montype(lastmontype).name = name
montype(lastmontype).level = level
montype(lastmontype).light = l
montype(lastmontype).color = RGB(r, g, b)
calcexp montype(lastmontype), boss
makesprite mongraphs(lastmontype), Form1.Picture1, montype(lastmontype).gfile, r, g, b, l, 3
loadsummons = lastmontype
'isloaded = 1
End Function

Function createtreasure(ByVal worth, X, Y)

If roll(3) = 1 Then
    broll = roll(3)
    If broll = 3 Then createobj "Huge Gold Pile", X, Y, , rolldice(averagelevel + 3, averagelevel + 8)
    If broll = 2 Then createobj "Medium Gold Pile", X, Y, , rolldice(averagelevel, averagelevel + 6)
    If broll = 1 Then createobj "Small Gold Pile", X, Y, , rolldice(averagelevel, averagelevel + 2)
End If

If worth = 1 Then createobj "Treasure Bag (Orange)", X, Y
If worth = 2 Then createobj "Treasure Bag (Green)", X, Y
If worth = 3 Then createobj "Treasure Bag (Red)", X, Y
If worth = 4 Then createobj "Treasure Chest (Orange)", X, Y
If worth = 5 Then createobj "Treasure Bag (Blue)", X, Y
If worth = 6 Then createobj "Treasure Chest (Green)", X, Y
If worth = 7 Then createobj "Treasure Bag (Black)", X, Y
If worth = 8 Then createobj "Treasure Chest (Red)", X, Y
If worth = 9 Then createobj "Treasure Chest (Blue)", X, Y
If worth = 10 Then createobj "Treasure Chest (Black)", X, Y

End Function

Function createclothes(ByVal worth As Byte, X, Y, Optional buytype = "", Optional magicchance = 5)

mylev = averagelevel
If mylev > worth Then worth = mylev
'If worth = 0 Then worth = averagelevel
If buytype = "Weapons" Or buytype = "Weapon" Then c = makeobjtype(randwep(worth))
If buytype = "Armor" Then c = makeobjtype(randarmor(worth + 1))
If buytype = "" Then c = makeobjtype(randclothes(, , , , , worth + 1))

If roll(magicchance) = 1 Then makemagicitem objtypes(c), , buytype

createobj objtypes(c).name, X, Y

'    randclothes 1, worth
'    If buytype = "Armor" Then randarmor 1, worth
'    If buytype = "" Then f = createobjtype(buyarmor(1).name, "clothes.bmp")
'    If buytype = "Armor" Then f = createobjtype(buyarmor(1).name, "armorobj1.bmp")
'    zerg = "Clothes"
'    If buytype = "Weapons" Then zerg = "Weapon"
'    z = addeffect(f, zerg, buyarmor(1).graph, buyarmor(1).armor, buyarmor(1).wear1, buyarmor(1).wear2)
'    If buytype = "Weapons" Then objtypes(f).effect(z, 5) = buyarmor(1).wear2
'    objtypes(f).r = buyarmor(1).r
'    objtypes(f).g = buyarmor(1).g
'    objtypes(f).b = buyarmor(1).b
'    objtypes(f).l = buyarmor(1).l
'    If buytype = "" Then makesprite objtypes(f).graph, Form1.Picture1, "clothes.bmp", objtypes(f).r, objtypes(f).g, objtypes(f).b, objtypes(f).l, 1
'    If buytype = "Armor" Then makesprite objtypes(f).graph, Form1.Picture1, "armorobj1.bmp", objtypes(f).r, objtypes(f).g, objtypes(f).b, objtypes(f).l, 1
'    createobj buyarmor(1).name, x, y, buyarmor(1).name

End Function

Function createdungeon(ByVal tt, ByVal ot, ByVal X, ByVal Y, ByVal x2, ByVal y2, Optional treasureon = 0)

If X = 0 Then X = 1: If x2 = 0 Then x2 = mapx
If Y = 0 Then Y = 1: If y2 = 0 Then y2 = mapy

For a = X To x2
For b = Y To y2
    map(a, b).tile = tt
    map(a, b).ovrtile = ot
    map(a, b).blocked = 1
Next b
Next a

'If eside = 0 Then eside = roll(4)
For a = 1 To ((x2 - X) / 9) * ((y2 - Y) / 9)
    trstr = ""
    troll = roll(60)
    If troll <= 4 Then trstr = "Treasure" & wildroll(12) * treasureon
    If troll = 5 Or troll = 6 Then trstr = "Potion" & wildroll(8) * treasureon
    If troll = 7 Then trstr = "Clothes" & wildroll(2) * treasureon
    If troll = 8 Then trstr = "Armor" & wildroll(1) * treasureon
    If treasureon = 0 Then trstr = ""
    createbuilding , , trstr, , tt, ot, roll(x2) + X - 10, roll(y2) + Y - 10, roll(2) + 2
Next a

'Create objective room
'troll = roll(5)
'    If troll <= 3 Then trstr = "Treasure" & wildroll(12) * (treasureon * 2) + 6
'    If troll = 4 Then trstr = "Clothes" & wildroll(8) * (treasureon * 2) + 6
'    If troll = 5 Then trstr = "Armor" & wildroll(4) * (treasureon * 2) + 6
If Not treasureon = 0 Then createbuilding , , "Objective" & treasureon * 2, , tt, ot, x2 / 2, y2 / 2, 1

End Function

Function wildroll(damage)
wildroll = lowestlevel
'5 aroll = roll(damage)
'total = total + aroll
'If aroll = damage Then GoTo 5
'wildroll = total
End Function

Function createpotion(ByVal worth, X, Y)

If worth = 1 Then createobj "Lesser Healing Potion", X, Y
If worth = 2 Then createobj "Healing Potion", X, Y
If worth = 3 Then createobj "Greater Healing Potion", X, Y

End Function

Function createobjective(ByRef worth, X, Y)

aroll = roll(8)
'If aroll = 1 Then
'If worth <= 1 Then createobj "Potion of Experience", X, Y: worth = worth - 1
'If worth = 2 Then createobj "Greater Potion of Experience", X, Y: worth = worth - 2
'If worth >= 3 Then createobj "Mega Potion of Experience", X, Y: worth = worth - 3
'End If

If aroll = 2 Then createobj "Potion of Strength (Permanent)", X, Y: worth = worth - 1
If aroll = 3 Then createobj "Potion of Dexterity (Permanent)", X, Y: worth = worth - 1
If aroll = 4 Then createobj "Potion of Intelligence (Permanent)", X, Y: worth = worth - 1
If aroll = 5 Or aroll = 1 Then createclothes worth * 2, X, Y, , 1
If aroll = 6 Then createclothes worth * 1, X, Y, "Armor", 1
If aroll = 7 Or aroll = 8 Then createclothes worth * 2, X, Y, "Weapon", 1


End Function

Function getstr2(ByVal str, Optional ByVal pos = 0)

tpos = 1
If pos = 0 Then getstr2 = Mid(str, 1, InStr(tpos, str, ":") - 1): Exit Function
For a = 1 To pos
    tpos = InStr(tpos + 1, str, ":")
Next a

getstr2 = Mid(str, InStr(tpos, str, ":") + 1)
If Not InStr(tpos + 1, str, ":") = 0 Then If Not InStr(InStr(tpos + 1, str, ":"), str, ":") = 0 Then getstr2 = Mid(str, InStr(tpos, str, ":") + 1, InStr(InStr(tpos + 1, str, ":"), str, ":") - tpos - 1)

End Function

Function clearsides()
For a = 1 To 4
    If Not mapjunk.maps(a) = "" Then clearside a, map(1, 1).tile
Next a
End Function

Function clearside(side, tile)
doorx = X
doory = Y
If side = 1 Then xp = 0: yp = 1: doorx = Int(mapx / 2): doory = 1
If side = 2 Then xp = -1: yp = 0: doorx = mapx: doory = Int(mapy / 2)
If side = 3 Then xp = 0: yp = -1: doorx = Int(mapx / 2): doory = mapy
If side = 4 Then xp = 1: yp = 0: doorx = 1: doory = Int(mapy / 2)

Do Until map(doorx, doory).blocked = 0 And map(doorx, doory).ovrtile = 0
    map(doorx, doory).tile = tile
    map(doorx, doory).used = 1
    map(doorx, doory).ovrtile = 0
    map(doorx, doory).blocked = 0
    doorx = doorx + xp: doory = doory + yp
    If doorx < 1 Or doorx > mapx Or doory < 1 Or doory > mapy Then Exit Do
Loop

End Function

Function calcexp(ByRef mn As monstertype, Optional ByVal boss = 0)
'Set up monsters by level and type

mn.boss = boss

If Val(mn.level) > 0 Then
setmon mn, 1, 1, 3, 6, 4, 4, 4, 1
mn.Sound = "MonSTANDARD.wav"
If mn.gfile = "thirsha1.bmp" Then setmon mn, 2, 2, 8, 6, 6, 8, 4, 2
If mn.gfile = "giantess.bmp" Then setmon mn, 2, 1, 4, 6, 6, 8, 4, 1
If mn.gfile = "kari1.bmp" Then setmon mn, 1.5, 1.2, 3, 4, 4, 6, 4, 1: mn.weaktype = "Spear"
If mn.gfile = "snake1.bmp" Then setmon mn, 1, 1, 2, 6, 5, 6, 4, 1.2: mn.weaktype = "Sword"
If mn.gfile = "snakewoman1.bmp" Then setmon mn, 1.3, 1.4, 3, 8, 5, 5, 4, 1.5: mn.Sound = "Monwierd.wav": mn.weaktype = "Axe"
If mn.gfile = "snakewoman2.bmp" Then setmon mn, 1.8, 1.5, 2, 10, 5, 6, 4, 1: mn.Sound = "Monwierd.wav": mn.weaktype = "Axe"
If mn.gfile = "worm1.bmp" Then setmon mn, 1.2, 1, 2, 4, 6, 4, 4, 2: mn.Sound = "Monbeast.wav": mn.weaktype = "Axe"
If mn.gfile = "worm2.bmp" Then setmon mn, 1, 1.2, 2, 4, 7, 6, 2, 2: mn.Sound = "Monbeast.wav": mn.weaktype = "Axe"
If mn.gfile = "tendrils1.bmp" Then setmon mn, 0.5, 1, 1, 4, 7, 3, 2, 0.8: mn.Sound = "Monfrog.wav": mn.weaktype = "Sword": mn.eattype = 1
If mn.gfile = "venus1.bmp" Then setmon mn, 2, 0.8, 2, 6, 8, 9, 1, 2: mn.Sound = "Monfrog.wav": mn.eattype = 1
If mn.gfile = "groundmaw2.bmp" Then setmon mn, 8, 2, 1, 3, 15, 9, 0, 0.8: mn.Sound = "Monfrog.wav": mn.eattype = 1
If mn.gfile = "frog1.bmp" Then setmon mn, 0.8, 1, 1, 4, 7, 4, 4, 0.6: mn.Sound = "Monfrog.wav": mn.weaktype = "Mace"
If mn.gfile = "slime1.bmp" Then setmon mn, 2, 1, 4, 8, 7, 3, 4, 4: mn.Sound = "Monslime.wav": mn.eattype = 1: mn.weaktype = "Mace"
If mn.gfile = "enzyme2.bmp" Then setmon mn, 2, 1, 4, 8, 7, 3, 4, 4: mn.Sound = "Monslime.wav": mn.eattype = 1: mn.weaktype = "Mace"
If mn.gfile = "centauress1.bmp" Then setmon mn, 1.4, 1, 3, 6, 3, 8, 4, 1: mn.weaktype = "Spear"
If mn.gfile = "harpy1.bmp" Then setmon mn, 1.1, 2, 4, 6, 3, 4, 6, 1.5: mn.Sound = "Monshriek.wav": mn.weaktype = "Bow"
If mn.gfile = "sprite1.bmp" Then setmon mn, 0.6, 2, 1, 8, 4, 5, 7, 1.5: mn.Sound = "Magic2.wav": mn.weaktype = "Fast"
If mn.gfile = "succubus1.bmp" Then setmon mn, 2, 1.5, 3, 12, 3, 6, 4, 1.3: mn.weaktype = "Spear"
If mn.gfile = "demoness1.bmp" Then setmon mn, 2.5, 1.3, 4, 8, 2, 6, 5, 3: mn.weaktype = "Staff"
If mn.gfile = "demoness2.bmp" Then setmon mn, 1.5, 1.2, 3, 6, 3, 4, 5, 2, , 2, 6, 6, 2: mn.weaktype = "Staff"
If mn.gfile = "demoness4f.bmp" Then setmon mn, 2.5, 1.3, 4, 8, 2, 6, 5, 3: mn.weaktype = "Staff"
If mn.gfile = "lizardwoman1.bmp" Then setmon mn, 1.2, 2, 3, 6, 3, 3, 5, 1: mn.weaktype = "Fast"
If mn.gfile = "flower1.bmp" Then setmon mn, 1.3, 2, 1, 2, 8, 7, 0, 1: mn.Sound = "Monwierd.wav": mn.weaktype = "Axe"
If mn.gfile = "mage1.bmp" Then setmon mn, 0.8, 0.8, 2, 6, 2, 5, 2, 1, , 2, 2, 4, 2: mn.Sound = "Magic2.wav": mn.weaktype = "Sword"
If mn.gfile = "merc1.bmp" Then setmon mn, 2, 2, 3, 20, 1, 6, 5, 1: mn.Sound = "swing.wav"
If mn.gfile = "warrior1.bmp" Then setmon mn, 2, 3, 3, 20, 1, 6, 5, 1: mn.Sound = "swing.wav"
End If

'Calculate experience before boss HP bonuses (To keep from gaining a level every time you hit a boss...)
'Experience: 1/10th of HP plus weighted skill totals * 5 (HP makes less difference now, since more HP is reflected by XP per hit--more HP doesn't matter if they die in one hit anyway

mn.exp = (mn.hp * 0.1) + (mn.skill * 1.5 + mn.dice * 2 + Val(mn.missileatk) + mn.eatskill) * 5

If boss >= 0 And Not boss = 1 And Not boss = 2 Then mn.hp = mn.hp * ((boss + 1) * (boss + 1))
'Boss 1.5: x6.25 3: x16, 4: x25, 5: x36, 6: x42, 7: x64, 8: x81, 9: x100
If boss = 2 Then mn.hp = mn.hp * 10
If boss = 1 Then mn.hp = mn.hp * 3



'mn.exp = round(mn.hp * avgmult((mn.skill / 10) * (mn.damage / 3) * (mn.dice / 3) * (mn.eatskill / 10) * ((mn.acid + mn.eatskill) / 10) * avgmult(mn.move / 10)))
mn.name = mn.name 'Get rid of--simply for quick debug viewing
'mn.hp = mn.hp / 2 'Cuts HP in half because I seem to have gotten a little excessive
End Function

Function setmon(mont As monstertype, ByVal hpmult, ByVal skillmult, ByVal dice, ByVal damage, ByVal eatskill, ByVal escapediff, ByVal move, ByVal acidmult, Optional ByVal level, Optional ByVal missiledice = 0, Optional ByVal missiledamage = 0, Optional ByVal missilefreq, Optional ByVal missiletype)

If plr.difficulty > 0 Then If IsMissing(level) Then level = greater(Val(mont.level), Int(plr.level * 0.7))
If IsMissing(level) Then level = Val(mont.level)

If Not getfromstring(mont.level, 2) = "" Then
    missiledice = 2 + missiledice
    missiledamage = 6 + missiledamage
    missilefreq = getfromstring(mont.level, 3)
    missiletype = getfromstring(mont.level, 2)
End If

'Commented values are origional values

'mont.hp = (level * 2) * (1 + level / 20) * 10 * hpmult
'mont.skill = ((level / 2) + 3) * skillmult
'mont.dice = dice + (dice * (level / 3))
'mont.damage = damage + (level / 4)
'mont.eatskill = (level / 3) + eatskill + 2
'mont.escapediff = (level / 3) + escapediff + 2
'mont.move = move
'mont.acid = (level * 0.7 + 1) * acidmult

'If isexpansion = 0 Then

'    mont.hp = (1 + level * 2) * (1 + level / 20) * 10 * hpmult * (1 + plr.difficulty / 2)
'    mont.skill = ((level / 3) + 2) * skillmult
'    mont.dice = dice + (dice * (level / 6))
'    mont.damage = damage + (level / 10) + plr.difficulty * 2
'    mont.eatskill = (level / 4) + eatskill + plr.difficulty
'    mont.escapediff = (level / 4) + escapediff + plr.difficulty
'    mont.move = move
'    mont.acid = lesser((level * 0.8 + 2) * acidmult + plr.difficulty * 2, 255)

'Else:

    'HP: 50 + 10 per level, +10% per level, +25% per difficulty level
    mont.hp = (50 + level * 10) * (1 + level / 10) * hpmult * (1 + plr.difficulty / 4)
    mont.skill = ((level / 2) + 2) * skillmult
    mont.dice = dice + (level / 5)
    mont.damage = damage + (level / 10) + plr.difficulty * 2
    mont.eatskill = (level / 3) + eatskill + plr.difficulty
    mont.escapediff = (level / 3) + escapediff + plr.difficulty
    mont.move = move
    mont.acid = lesser((level * 0.5 + 2) * acidmult + plr.difficulty * 2, 255)

'End If

If missiledice > 0 Then
    missiledice = missiledice + (level / 4)
    missiledamage = missiledamage + level
    mont.missileatk = Int(missiletype) & ":" & Int(missilefreq) & ":" & Int(missiledice) & ":" & Int(missiledamage)
End If

End Function

Function avgmult(ByVal snarg As Double) As Double

snarg = (snarg + snarg + 1) / 3
avgmult = snarg
End Function

Function round(ByVal num As Double) As Long
If num > 10000000 Then num = 10000000
num = Int(num)
num = num - (num Mod 25)

'num = num + 5
round = num
End Function

Sub clearboosts()
plr.str = plr.str - plr.strboost: plr.strboost = 0
plr.int = plr.int - plr.intboost: plr.intboost = 0
plr.dex = plr.dex - plr.dexboost: plr.dexboost = 0
plr.armorboost = 0
plr.regen = 0
plr.monsinbelly = 0
End Sub

Sub decboosts()
If boostcount > 0 Then boostcount = boostcount - 1
If boostcount > 0 Then Exit Sub
If plr.strboost > 0 Then plr.str = plr.str - 1: plr.strboost = plr.strboost - 1
If plr.dexboost > 0 Then plr.dex = plr.dex - 1: plr.dexboost = plr.dexboost - 1
If plr.intboost > 0 Then plr.int = plr.int - 1: plr.intboost = plr.intboost - 1
If plr.regen > 0 Then plr.regen = plr.regen - 1
If plr.armorboost > 0 Then plr.armorboost = plr.armorboost / 2
boostcount = 40
End Sub

Function losemp(ByVal mp, Optional permanent = 0) As Boolean

losemp = True
If mp > 0 Then If plr.mp >= mp Then plr.mp = Int(plr.mp) - Int(mp) Else losemp = False 'If permanent = 0 Then mp = mp - plr.mp: plr.mp = 0: losemp = False ': If plr.fatigue < plr.fatiguemax Then plr.fatigue = plr.fatigue + mp * 5: losemp = True Else losemp = False
If mp < 0 Then plr.mp = plr.mp - mp
If permanent = 1 Then plr.mpmax = plr.mpmax - mp: plr.mplost = plr.mplost + mp
If plr.mp < 0 Then plr.mp = 0
If plr.mp > getmpmax Then plr.mp = getmpmax

End Function

Sub undigest()

a = topclothes()
If clothes(a).digested = 1 Then GoTo 5
For a = 1 To 16
    If clothes(a).digested = 1 Then Exit For
Next a
5 If a = 9 Then Exit Sub
clothes(a).armor = clothes(a).armor * 2
clothes(a).digested = 0
swapgraph cgraphs(a).graph, cgraphs(a).diggraph
Form1.updatbody

End Sub

Sub swapgraph(gr1 As cSpriteBitmaps, gr2 As cSpriteBitmaps)
Dim backgraph As cSpriteBitmaps
Set backgraph = gr1
Set gr1 = gr2
Set gr2 = backgraph

End Sub

Sub takeoffclothes(wch)
If wch < 1 Then Exit Sub
    If clothes(wch).loaded = 0 Then Exit Sub
    clothes(wch).wear1 = ""
    clothes(wch).wear2 = ""
    clothes(wch).loaded = 0
    clothes(wch).armor = 0
    clothes(wch).name = ""
    'If clothes(wch).digested = 0 Then
    getitem clothes(wch).obj
    killitem clothes(wch).obj
    'clothes(wch).digested = 0
    Set cgraphs(wch).graph = Nothing
    
Form1.updatbody
End Sub

Sub makeriver()

Dim xdirec As Integer
Dim ydirec As Integer
wdir = roll(8)
wid = roll(5) + 1
xdirec = Sin(wdir)
ydirec = Cos(wdir)
If xdirec = 1 Then wx1 = 1
If xdirec = -1 Then wx1 = mapx
If xdirec = 0 Then wx1 = roll(mapx)

If ydirec = 1 Then yx1 = 1
If ydirec = -1 Then yx1 = mapy
If ydirec = 0 Then yx1 = roll(mapy)

For a = 1 To wid
    'puttile
Next a

End Sub

'Sub puttile(x, y, tt)
'If x > mapx Or x < 1 Then puttile = 0: Exit Sub
'If y > mapx Or y < 1 Then puttile = 0: Exit Sub
'map(x, y).tile = tt
'End Sub

Function getdigging()

aroll = roll(12)
mname = montype(mon(plr.instomach).type).name

If aroll = 1 Then txt = "Your unconcious body floats in the " & mname & "'s stomach, slowly being dissolved into nothing but bodily waste."
If aroll = 2 Then txt = "The " & mname & "'s body groans loudly as it slowly digests your unconcious body."
If aroll = 3 Then txt = "You have lost. You are now nothing but a large lump of food being digested inside the " & mname & "'s body."
If aroll = 4 Then txt = "The " & mname & " happily goes about it's business as you slowly digest in it's belly..."
If aroll = 5 Then txt = "You have lost all of your HP. You are now doomed to be digested by the " & mname & " until it excretes your undigestable remains."
If aroll = 6 Then txt = "You are now a meal for the " & mname & ". Even now, your unconscious body is being gradually digested."
If aroll = 7 Then txt = "The acid in the " & mname & "'s belly is slowly stripping the meat from your bones as it grinds you into mush to be absorbed by it's intestines."
If aroll = 8 Then txt = "Your limp body sits in the soft depths of the " & mname & ", slowly being churned away as the " & mname & " digests you..."
If aroll = 9 Then txt = "You have failed in your quest, having become dinner for a " & mname & "."
If aroll = 10 Then txt = "You lie unconscious as the " & mname & " contentedly digests it's dinner. Unfortunately, that dinner is you."
If aroll = 11 Then txt = "You have lost all of your HP. There is now no hope of escape from the " & mname & "'s stomach."
If aroll = 12 Then txt = "The " & mname & " continues to digest your body, it's stomach considering you nothing more than a big meal to be digested like any other."

'txt = txt & vbCrLf & "GAME OVER."
getdigging = txt
End Function

Sub fsavechar(filen)

If Not Dir(filen) = "" Then Kill filen

Open filen For Binary As #1 'Len = 30000
Put #1, , plr
    
    Put #1, , wep
    'Put #1, , wep.graphname
    'Put #1, , wep.xoff
    'Put #1, , wep.yoff
    'Put #1, , wep.dice
    'Put #1, , wep.damage
    'Put #1, , wep.type
    'Put #1, , wep.r
    'Put #1, , wep.g
    'Put #1, , wep.b
    'Put #1, , wep.l
    
    
    '    Put #1, , wep.obj.name
    '    For a = 1 To 6
    '        For b = 1 To 5
    '            Put #1, , wep.obj.effect(a, b): Next b: Next a
    '    Put #1, , wep.obj.graphname
    '    Put #1, , wep.obj.r
    '    Put #1, , wep.obj.g
    '    Put #1, , wep.obj.b
    '    Put #1, , wep.obj.l
    '    Put #1, , wep.obj.cells
    

Put #1, , spells()
'Put #1, , clothes()

'Open filen For Random As #1 Len = 500

Put #1, , clothes()

'For a = 1 To 8
'    putclothes clothes(a)
'Next a

Put #1, , inv()

'For a = 1 To 50
'    putobj inv(a)
'Next a

Close #1

End Sub

Sub floadchar(filen)

If Dir(filen) = "" Or filen = "" Then MsgBox "File not found!": Exit Sub

Form1.Text5.Visible = True
Form1.Text5.Refresh

strip

Open filen For Binary As #1 'Len = 30000
Get #1, , plr
    
Get #1, , wep

Get #1, , spells()

Get #1, , clothes()
For a = 1 To 16
    If Not clothes(a).name = "" Then
    cfilen = geteff(clothes(a).obj, "Clothes", 2)
    makesprite cgraphs(a).graph, Form1.Picture1, cfilen, clothes(a).obj.r, clothes(a).obj.g, clothes(a).obj.b, clothes(a).obj.l
    If Not Dir("D-" & cfilen) = "" Then
    makesprite cgraphs(a).diggraph, Form1.Picture1, "D-" & cfilen, clothes(a).obj.r, clothes(a).obj.g, clothes(a).obj.b, clothes(a).obj.l
    Else:
    makedigsprite cgraphs(a).diggraph, Form1.Picture1, cfilen, clothes(a).obj.r, clothes(a).obj.g, clothes(a).obj.b, clothes(a).obj.l
    'Set clothes(a).diggraph = clothes(a).graph
    
    End If
    
    If clothes(a).digested > 0 Then swapgraph cgraphs(a).graph, cgraphs(a).diggraph
    
    End If
Next a

Get #1, , inv()


Close #1

bodyimg = plr.bodyname
If plr.Class = "" Then bodyimg = "body2.bmp"
makesprite gbody, Form1.Picture1, bodyimg


lwepgraph wep.graphname, Form1.Picture1
'bodyimg = plr.Class & ".bmp"

Form1.updatbody
Form1.updatinv
Form1.updatspells
fn = plr.curmap

revamp
clearboosts
'savebindata Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"
'loaddata fn
If loadbindata(Left(fn, Len(fn) - 4) & ".dat") = False Then loaddata fn 'loadbindata Left(fn, Len(plr.curmap) - 4) & ".dat"

clearsides

Form1.Text5.Visible = False

If Not Dir("autosave.plr") = "" Then Kill "autosave.plr"

End Sub

Sub floadcharold(filen)

If Dir(filen) = "" Then MsgBox "File not found!": Exit Sub

Form1.Text5.Visible = True
Form1.Text5.Refresh

strip

Open filen For Binary As #1 'Len = 30000
Get #1, , plr

    Get #1, , wep.graphname
    Get #1, , wep.xoff
    Get #1, , wep.yoff
    Get #1, , wep.dice
    Get #1, , wep.damage
    Get #1, , wep.type
    Get #1, , wep.r
    Get #1, , wep.g
    Get #1, , wep.b
    Get #1, , wep.l

        Get #1, , wep.obj.name
        For a = 1 To 6
            For b = 1 To 5
                Get #1, , wep.obj.effect(a, b): Next b: Next a
        Get #1, , wep.obj.graphname
        Get #1, , wep.obj.r
        Get #1, , wep.obj.g
        Get #1, , wep.obj.b
        Get #1, , wep.obj.l
        Get #1, , wep.obj.cells


Get #1, , spells()
'Get #1, , clothes

'Open filen For Random As #1 Len = 500
lwepgraph wep.graphname, Form1.Picture1
bodyimg = plr.Class & ".bmp"
If plr.Class = "" Then bodyimg = "body2.bmp"
makesprite gbody, Form1.Picture1, bodyimg

inv(1).name = ""
For c = 1 To 16
    getclothes clothes(c)
    takeoffclothes c
'    takeobj -1, 1
    a = getobjslot(inv(1), "Clothes")
    If a = Empty Then GoTo 17
    F = addclothes(inv(1).name, inv(1).effect(a, 2), Val(inv(1).effect(a, 3)), inv(1).effect(a, 4), inv(1).effect(a, 5), inv(1).r, inv(1).g, inv(1).b, inv(1).l): clothes(F).obj = inv(1): checkclothes F ': destroy = 1: Form1.updatbody
    inv(1).name = ""

17 Next c

For a = 1 To 50
    getobj inv(a)
Next a


Close #1



Form1.updatbody
Form1.updatinv
Form1.updatspells
fn = plr.curmap

revamp
clearboosts
'savebindata Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"
'loaddata fn
If loadbindata(Left(fn, Len(fn) - 4) & ".dat") = False Then loaddata fn 'loadbindata Left(fn, Len(plr.curmap) - 4) & ".dat"
clearsides

Form1.Text5.Visible = False

End Sub

Sub newchar(Optional conf As Byte = 0)
If conf = 1 Then If MsgBox("Are you sure you want to make a new character?", vbOKCancel) = vbCancel Then Exit Sub
'loaddata "Spells.txt"


loadgame = 0
15 Form6.Show '1
DoEvents 'm'' an infinite loop to avoid
16
If plr.Class = "" Then If loadgame = 0 Then DoEvents: GoTo 16 Else GoTo 15
If loadgame > 0 Then Exit Sub

If plr.name = "" Then plr.name = getname
strip

If Not plr.classskills(1) = "" Then GoTo 5

    plr.classskills(1) = "Deathblow"
    plr.classskills(2) = "Critical Strike"
    plr.classskills(3) = "Dodge"
    plr.classskills(4) = "Endurance"
    plr.classskills(5) = "Resilience"
    plr.classskills(6) = "Weapons Mastery"


If plr.Class = "Naga" Then
    
    plr.classskills(1) = "Giant Stomach"
    plr.classskills(2) = "Gluttony"
    plr.classskills(3) = "Water Magic"
    plr.classskills(4) = "Endurance"
    plr.classskills(5) = "Resilience"
    plr.classskills(6) = "Mana Mastery"

    addskill "Defence", 0
    addskill "Sorcery", 0
    addskill "Spear Mastery", 0
    addskill "Bow Mastery", 0
    addskill "Greed", 0
    addskill "Regeneration", 0
    addskill "Mana Regeneration", 0
    addskill "Super Acid", 0
    
    #If USELEGACY = 1 Then
    'm'' nothing
    #Else
    'm'' Following Naga class description, Giant Stomach should be at lvl 1
    addskill "Giant Stomach" 'm''
    #End If


End If

If plr.Class = "Angel" Then
    plr.classskills(1) = "Mana Regeneration"
    plr.classskills(2) = "Spell Mastery"
    plr.classskills(3) = "Air Magic"
    plr.classskills(4) = "Water Magic"
    plr.classskills(5) = "Regeneration"
    plr.classskills(6) = "Defence"

    addskill "Dodge", 0
    addskill "Grey Magic", 0
    addskill "Mana Mastery", 0
    addskill "Bow Mastery", 0
    addskill "Streetfighting", 0
    addskill "Giant Stomach", 0
    addskill "Super Acid", 0
    addskill "Gluttony", 0


End If

If plr.Class = "Succubus" Then
    plr.classskills(1) = "Drain Life"
    plr.classskills(2) = "Drain Magic"
    plr.classskills(3) = "Demon Summoning"
    plr.classskills(4) = "Firepower"
    plr.classskills(5) = "Endurance"
    plr.classskills(6) = "Fire Magic"

    addskill "Sorcery", 0
    addskill "Mana Regeneration", 0
    addskill "Mana Mastery", 0
    addskill "Critical Strike", 0
    addskill "Deathblow", 0
    #If USELEGACY = 1 Then
    addskill "Giant Stomach", 1
    #Else
    addskill "Giant Stomach", 0 'm'' Duam have set it to 1, but description says it should be 0
    #End If
    addskill "Super Acid", 0
    addskill "Gluttony", 0

End If

If plr.Class = "TombRaider" Then
    plr.classskills(1) = "Dodge"
    plr.classskills(2) = "Endurance"
    plr.classskills(3) = "Critical Strike"
    plr.classskills(4) = "Deathblow"
    plr.classskills(5) = "Accuracy"
    plr.classskills(6) = "Greed"

    addskill "Resilience", 0
    addskill "Regeneration", 0
    addskill "Squirm", 0
    addskill "Evasion", 0
    addskill "Defence", 0
    addskill "Giant Stomach", 0
    addskill "Super Acid", 0
    addskill "Gluttony", 0


End If

If plr.Class = "Streetfighter" Then
    plr.classskills(1) = "Dodge"
    plr.classskills(2) = "Endurance"
    plr.classskills(3) = "Resilience"
    plr.classskills(4) = "Deathblow"
    plr.classskills(5) = "Streetfighting"
    plr.classskills(6) = "Accuracy"

    addskill "Earth Magic", 0
    addskill "Regeneration", 0
    addskill "Squirm", 0
    addskill "Critical Strike", 0
    addskill "Drain Life", 0
    addskill "Giant Stomach", 0
    addskill "Super Acid", 0
    addskill "Gluttony", 0


End If

If plr.Class = "Enchantress" Then
    plr.classskills(1) = "Sorcery"
    plr.classskills(2) = "Grey Magic"
    plr.classskills(3) = "Spell Mastery"
    plr.classskills(4) = "Regeneration"
    plr.classskills(5) = "Mana Mastery"
    plr.classskills(6) = "Squirm"

    addskill "Water Magic", 0
    addskill "Air Magic", 0
    addskill "Fire Magic", 0
    addskill "Earth Magic", 0
    addskill "Dodge", 0
    addskill "Giant Stomach", 0
    addskill "Super Acid", 0
    addskill "Gluttony", 0


End If

If plr.Class = "Huntress" Then
    plr.classskills(1) = "Accuracy"
    plr.classskills(2) = "Critical Strike"
    plr.classskills(3) = "Weapons Mastery"
    plr.classskills(4) = "Deathblow"
    plr.classskills(5) = "Fire Magic"
    plr.classskills(6) = "Dodge"
    
    addskill "Spear Mastery", 0
    addskill "Axe Mastery", 0
    addskill "Sword Mastery", 0
    addskill "Bow Mastery", 0
    addskill "Drain Life", 0
    addskill "Giant Stomach", 0
    addskill "Super Acid", 0
    addskill "Gluttony", 0

    
End If

If plr.Class = "Valkyrie" Then
    plr.classskills(1) = "Critical Strike"
    plr.classskills(2) = "Defence"
    plr.classskills(3) = "Resilience"
    plr.classskills(4) = "Earth Magic"
    plr.classskills(5) = "Mana Regeneration"
    plr.classskills(6) = "Air Magic"

    addskill "Streetfighting", 0
    addskill "Bow Mastery", 0
    addskill "Sword Mastery", 1
    addskill "Earth Magic", 0
    addskill "Water Magic", 0
    addskill "Giant Stomach", 0
    addskill "Super Acid", 0
    addskill "Gluttony", 0

    'wep.r = 150
    'wep.g = 150
    'wep.b = 150
    'wep.l = 150
    'wep.dice = 4
    'wep.damage = 6
    'wep.weight = 4
    'wep.type="

End If

If plr.Class = "Amazon" Then
    plr.classskills(1) = "Deathblow"
    plr.classskills(2) = "Critical Strike"
    plr.classskills(3) = "Dodge"
    plr.classskills(4) = "Endurance"
    plr.classskills(5) = "Resilience"
    plr.classskills(6) = "Weapons Mastery"
    
    addskill "Streetfighting", 0
    addskill "Axe Mastery", 0
    addskill "Sword Mastery", 0
    addskill "Spear Mastery", 0
    addskill "Bow Mastery", 0
    addskill "Giant Stomach", 0
    addskill "Super Acid", 0
    addskill "Gluttony", 0

End If

If plr.Class = "Priestess" Then
    plr.classskills(1) = "Defence"
    plr.classskills(2) = "Dodge"
    plr.classskills(3) = "Evasion"
    plr.classskills(4) = "Endurance"
    plr.classskills(5) = "Earth Magic"
    plr.classskills(6) = "Air Magic"
    
    addskill "Water Magic", 0
    addskill "Axe Mastery", 0
    addskill "Resilience", 0
    addskill "Accuracy", 0
    addskill "Regeneration", 0
    addskill "Giant Stomach", 0
    addskill "Super Acid", 0
    addskill "Gluttony", 0

    
End If

If plr.Class = "Sorceress" Then
    plr.classskills(1) = "Mana Mastery"
    plr.classskills(2) = "Mana Regeneration"
    plr.classskills(3) = "Fire Magic"
    plr.classskills(4) = "Sorcery"
    plr.classskills(5) = "Firepower"
    plr.classskills(6) = "Spell Mastery"
    
    addskill "Earth Magic", 0
    addskill "Air Magic", 0
    addskill "Water Magic", 0
    addskill "Grey Magic", 0
    addskill "Greed", 0
    addskill "Giant Stomach", 0
    addskill "Super Acid", 0
    addskill "Gluttony", 0
    
    
End If

If plr.Class = "Caller" Then
    plr.classskills(1) = "Mana Mastery"
    plr.classskills(2) = "Critical Strike"
    plr.classskills(3) = "Endurance"
    plr.classskills(4) = "Demon Summoning"
    plr.classskills(5) = "Sorcery"
    plr.classskills(6) = "Spell Mastery"

    addskill "Water Magic", 0
    addskill "Air Magic", 0
    addskill "Fire Magic", 0
    addskill "Earth Magic", 0
    addskill "Dodge", 0
    addskill "Giant Stomach", 0
    addskill "Super Acid", 0
    addskill "Gluttony", 0


End If

5 'updatbuttnz


Select Case plr.Class
    Case "Amazon"
        plr.hpmax = 60: plr.str = 5: plr.dex = 3: plr.int = 2: plr.mpmax = 0: plr.endurance = 5
        'redd = 130: green = 80: blue = 0: lit = 0.4
        redd = 230: green = 80: blue = 0: lit = 0.4
            ''addclothes "Halter", "halter1.bmp", 4, "Bra", "Upper", redd, green, blue, lit, 1
            ''addclothes "Loincloth", "loincloth1.bmp", 3, "Lower", , redd, green, blue, lit, 1
            addskill "Sword Mastery"
            addskill "Axe Mastery"
            'addclothes "Chain Bodysuit", "chainswimsuit.bmp", 6, "Upper", "Lower", 150, 150, 150, 0.6, 1
            'addclothes "Crimson Sash", "sash1.bmp", 2, "Jacket", , 230, 0, 0, 0.4, 1
'    plr.combatskills(1) = "Power Strike"
'    plr.combatskills(2) = "Charged Strike"
'    plr.combatskills(3) = "Frenzy"
'    plr.combatskills(4) = "Cripple"
    
    
    Case "Sorceress"
        plr.hpmax = 30: plr.str = 2: plr.dex = 3: plr.int = 5: plr.mpmax = 60: plr.endurance = 2
        redd = 40: green = 40: blue = 40: lit = 0.3
            'addclothes "Fine Dress", "dress2.bmp", 4, "Upper", "Lower", redd, green, blue, lit, 1
            'addclothes "Lace Bra", "bra5.bmp", 2, "Bra", , redd, green, blue, lit, 1
            'addclothes "Lace Panties", "panties2.bmp", 2, "Panties", , redd, green, blue, lit, 1
            'givespefspell "Firebolt"
            addskill "Fire Magic"
            
            
            
  '  plr.combatskills(1) = "Alchemy"
  '  plr.combatskills(2) = "Damage Energy"
  '  plr.combatskills(3) = "Mana Shield"
  '  plr.combatskills(4) = "Split Spell"
    
    Case "Priestess"
        plr.hpmax = 45: plr.str = 4: plr.dex = 2: plr.int = 3: plr.mpmax = 30: plr.endurance = 4
        redd = 80: green = 120: blue = 10: lit = 0.4
            'addclothes "Shortcape", "cape1.bmp", 2, "Jacket", , redd, green, blue, lit, 1
        redd = 250: green = 250: blue = 250: lit = 0.4
            'addclothes "Gloves", "gloves1.bmp", 2, "Arms", , redd, green, blue, lit, 1
            'addclothes "Lace Panties", "panties2.bmp", 1, "Panties", , redd, green, blue, lit, 1
            'addclothes "Lace Bra", "bra5.bmp", 1, "Bra", , redd, green, blue, lit, 1
            addskill "Axe Mastery"
   ' plr.combatskills(1) = "Stun"
   ' plr.combatskills(2) = "Block"
   ' plr.combatskills(3) = "Power Strike"
   ' plr.combatskills(4) = "Charged Strike"
                
                
    Case "Enchantress"
        plr.hpmax = 35: plr.str = 2: plr.dex = 3: plr.int = 4: plr.mpmax = 45: plr.endurance = 3
        redd = 40: green = 40: blue = 120: lit = 0.4
            'addclothes "Robe Skirt", "robebottom1.bmp", 3, "Lower", , redd, green, blue, lit, 1
            'addclothes "Panties", "panties2.bmp", 1, "Panties", , 250, 250, 250, 0.5, 1
            'addclothes "Corset Bra", "bra3.bmp", 2, "Bra", , redd, green, blue, lit, 1
            'givespefspell "Speed"
            addskill "Grey Magic"
  '  plr.combatskills(1) = "Alchemy"
  '  plr.combatskills(2) = "Damage Energy"
  '  plr.combatskills(3) = "Mana Shield"
  '  plr.combatskills(4) = "Cunning Strike"
        
    Case "Huntress"
        plr.hpmax = 40: plr.str = 3: plr.dex = 4: plr.int = 3: plr.mpmax = 25: plr.endurance = 5
        redd = 110: green = 60: blue = 0: lit = 0.3
            'addclothes "Shirt", "doublet1.bmp", 2, "Upper", , redd, green, blue, lit, 1
            'addclothes "Armored Skirt", "armorskirt1.bmp", 2, "Lower", , redd, green, blue, lit, 1
            'addclothes "Bra", "bra1.bmp", 1, "Bra", , redd, green, blue, lit, 1
            'addclothes "Panties", "panties1.bmp", 1, "Panties", , redd, green, blue, lit, 1
        addskill "Bow Mastery"
        addskill "Spear Mastery"

    Case "Valkyrie"
        plr.hpmax = 35: plr.str = 3: plr.dex = 3: plr.int = 3: plr.mpmax = 35: plr.endurance = 4
        redd = 50: green = 120: blue = 220: lit = 0.5
            'addclothes "Plate Mail", "breastplate5.bmp", 8, "Upper", , redd, green, blue, lit, 1
            'addclothes "Bra", "bra1.bmp", 1, "Bra", , redd, green, blue, lit, 1
            'addclothes "Panties", "panties1.bmp", 1, "Panties", , redd, green, blue, lit, 1
            'addclothes "Armplates", "armplates1.bmp", 2, "Arms", , redd, green, blue, lit, 1
        addskill "Spear Mastery"

    Case "Angel"
        plr.hpmax = 50: plr.str = 2: plr.dex = 5: plr.int = 4: plr.mpmax = 35: plr.endurance = 3
        redd = 200: green = 200: blue = 200: lit = 0.2
            'addclothes "Robe", "robe1.bmp", 4, "Lower", "Upper", redd, green, blue, lit, 1

    Case "Succubus"
        plr.hpmax = 75: plr.str = 5: plr.dex = 3: plr.int = 4: plr.mpmax = 15: plr.endurance = 5
        redd = 20: green = 20: blue = 20: lit = 0.4
            ''addclothes "Lace Corset", "corset1.bmp", 3, "Upper", "Bra", redd, green, blue, lit, 1
            ''addclothes "Lace Panties", "panties2.bmp", 2, "Panties", , redd, green, blue, lit, 1
            'addclothes "Lace Teddy", "teddy1.bmp", 3, "Panties", "Bra", redd, green, blue, lit, 1
            'addclothes "Fishnet Stockings", "fishnets1.bmp", 2, "Legs", , redd, green, blue, lit, 1
     
    Case "TombRaider"
        plr.hpmax = 60: plr.str = 3: plr.dex = 4: plr.int = 4: plr.mpmax = 0: plr.endurance = 5
        'redd = 130: green = 80: blue = 0: lit = 0.4
        redd = 230: green = 230: blue = 230: lit = 0.6
            'addclothes "Bra", "bra2.bmp", 1, "Bra", , redd, green, blue, lit, 1
            'addclothes "Panties", "panties1.bmp", 1, "Panties", , redd, green, blue, lit, 1
            'addclothes "Shirt", "shirt1.bmp", 4, "Upper", , 70, 200, 210, 0.5, 1
            'addclothes "Shorts", "shorts1.bmp", 3, "Lower", , 120, 50, 8, 0.6, 1
            'addclothes "Belt", "belt.bmp", 2, "Gloves", , , , , , 1
    
    Case "Streetfighter"
        plr.hpmax = 60: plr.str = 5: plr.dex = 4: plr.int = 3: plr.mpmax = 0: plr.endurance = 5
        redd = 250: green = 250: blue = 250: lit = 0.5
            'addclothes "Combat Dress", "chunli1.bmp", 4, "Upper", "Lower", 0, 140, 255, lit, 1
            'addclothes "Lace Bra", "bra5.bmp", 2, "Bra", , redd, green, blue, lit, 1
            'addclothes "Lace Panties", "panties2.bmp", 2, "Panties", , redd, green, blue, lit, 1
    
    
    Case "Caller"
        plr.hpmax = 20: plr.str = 2: plr.dex = 3: plr.int = 4: plr.mpmax = 45: plr.endurance = 3
        redd = 20: green = 150: blue = 0: lit = 0.5
            'addclothes "Bra", "bra2.bmp", 1, "Bra", , redd, green, blue, lit, 1
            'addclothes "Panties", "panties1.bmp", 1, "Panties", , redd, green, blue, lit, 1
            redd = 20: green = 150: blue = 0: lit = 0.5
            'addclothes "Shirt", "doublet1.bmp", 4, "Upper", , redd, green, blue, 0.5, 1
            'addclothes "Robe Skirt", "robebottom1.bmp", 3, "Lower", , redd, green, blue, lit, 1
            addskill "Magic Summoning"
            'givespefspell "Firebolt"
            'givespefspell "Summon Faerie"
            
    Case "Naga"
        plr.hpmax = 40: plr.str = 3: plr.dex = 3: plr.int = 2: plr.mpmax = 15: plr.endurance = 4
            
End Select
If plr.Class = "" Then GoTo 15
plr.hpmax = plr.hpmax + 30
plr.gp = 500
plr.spmax = 80: plr.sp = plr.spmax
'bodyimg = plr.Class & ".bmp"
'If isexpansion = 0 Then plr.bodyname = plr.Class & ".bmp"
plr.bodyname = "Body" & roll(6) & ".bmp"
If plr.Class = "Succubus" Then plr.bodyname = "body12.bmp"
If plr.Class = "Angel" Then plr.bodyname = "body13.bmp"
If plr.Class = "Naga" Then plr.bodyname = "body11.bmp"
If plr.Class = "TombRaider" Then plr.bodyname = "laracroft.bmp": plr.hairname = "hair7.bmp": plr.haircolor = RGB(120, 100, 0)
bodyimg = plr.bodyname
Pickcombatskills.Show 1
makesprite gbody, Form1.Picture1, bodyimg
'gbody.CreateFromFile bodyimg, 1, 1, 1, RGB(0, 0, 0)
'If cheaton = 1 Then plr.name = "Trixie": GoTo 1245
If plr.Class = "TombRaider" Then plr.name = InputBox("Character name?", , "Lara Croft"): GoTo 1245
If plr.Class = "StreetFighter" Then plr.name = InputBox("Character name?", , "Chun Li"): GoTo 1245
If plr.Class = "Caller" Then plr.name = InputBox("Character name?", , "Rydia"): GoTo 1245

plr.name = InputBox("Character name?", , getname)
1245 If plr.name = "" Then plr.name = getname
plr.level = 1: plr.exp = 0: plr.expneeded = 600
If plr.Class = "Succubus" Or plr.Class = "Angel" Or plr.Class = "TombRaider" Then plr.expneeded = 900
plr.hp = plr.hpmax: plr.mp = plr.mpmax
Form1.updatbody
plr.lpotionlev = 1
plr.mpotionlev = 1
plr.fatigue = 1
plr.fatiguemax = greater(50, getend * 30 + 50)
Form1.updatinv

End Sub

Function strip()

For a = 1 To 16
    If clothes(a).loaded = 0 Then GoTo 5
    clothes(a).wear1 = ""
    clothes(a).wear2 = ""
    clothes(a).loaded = 0
    clothes(a).armor = 0
    clothes(a).name = ""
    clothes(a).digested = 0
    Set cgraphs(a).graph = Nothing
5 Next a

End Function

Sub loadbijdang()
Static isloaded As Long 'm'' added type

makesprite bijdang(1).cS, Form1.Picture1, "slash.bmp", wep.r, wep.g, wep.b, wep.l, 5
If wep.r = 0 And wep.g = 0 And wep.b = 0 Then makesprite bijdang(1).cS, Form1.Picture1, "slash.bmp", 255, 250, 245, 0.1, 5
Set bijdang(1).sprite = New cSprite
bijdang(1).sprite.SpriteData = bijdang(1).cS
bijdang(1).sprite.Create cStage.hDC
bijdang(1).sprite.cell = 5

If isloaded = 1 Then Exit Sub 'do not load death animations more than once

Set bijdang(2).cS = New cSpriteBitmaps
bijdang(2).cS.CreateFromFile "kablooie.bmp", 5, 2, , RGB(0, 0, 0)
'makesprite bijdang(1).cs, Form1.Picture1, "slash.bmp", 255, 250, 245, 0.1, 5
Set bijdang(2).sprite = New cSprite
bijdang(2).sprite.SpriteData = bijdang(2).cS
bijdang(2).sprite.Create cStage.hDC
bijdang(2).sprite.cell = 10

makesprite bijdang(3).cS, Form1.Picture1, "kablooie2.bmp", 50, 150, 250, , 5, 2
Set bijdang(3).sprite = New cSprite
bijdang(3).sprite.SpriteData = bijdang(3).cS
bijdang(3).sprite.Create cStage.hDC
bijdang(3).sprite.cell = 10

makesprite bijdang(4).cS, Form1.Picture1, "kablooieyellow.bmp", 50, 150, 250, , 5, 2
Set bijdang(4).sprite = New cSprite
bijdang(4).sprite.SpriteData = bijdang(4).cS
bijdang(4).sprite.Create cStage.hDC
bijdang(4).sprite.cell = 10

makesprite bijdang(5).cS, Form1.Picture1, "RedExp.bmp", 50, 150, 250, , 6, 1
Set bijdang(5).sprite = New cSprite
bijdang(5).sprite.SpriteData = bijdang(5).cS
bijdang(5).sprite.Create cStage.hDC
bijdang(5).sprite.cell = 6

makesprite bijdang(6).cS, Form1.Picture1, "PurpleExp.bmp", 50, 150, 250, , 8, 1
Set bijdang(6).sprite = New cSprite
bijdang(6).sprite.SpriteData = bijdang(6).cS
bijdang(6).sprite.Create cStage.hDC
bijdang(6).sprite.cell = 8

makesprite bijdang(7).cS, Form1.Picture1, "LitExp.bmp", 50, 150, 250, , 6, 1
Set bijdang(7).sprite = New cSprite
bijdang(7).sprite.SpriteData = bijdang(7).cS
bijdang(7).sprite.Create cStage.hDC
bijdang(7).sprite.cell = 6

makesprite bijdang(8).cS, Form1.Picture1, "RedExp2.bmp", 50, 150, 250, , 6, 1
Set bijdang(8).sprite = New cSprite
bijdang(8).sprite.SpriteData = bijdang(8).cS
bijdang(8).sprite.Create cStage.hDC
bijdang(8).sprite.cell = 6

isloaded = 1
End Sub


Sub drawbijdangold()

If bijdang(1).sprite.cell < 6 Then
X = bijdang(1).X
Y = bijdang(1).Y
dorkx = (X - plr.X - Y + plr.Y) * 48 + 400 - plr.xoff - (bijdang(1).cS.CellWidth / 2) + xoff
'If dorkx < -40 Or dorkx > 750 Then Exit Sub
dorky = (X + Y - plr.X - plr.Y) * 24 - plr.yoff + 300 - (bijdang(1).cS.CellHeight - 48) - (midtile * 24) + yoff
dorkx = Int(dorkx)

dorky = Int(dorky) '+ smurg
bijdang(1).sprite.X = dorkx
bijdang(1).sprite.Y = dorky '- smurg * 10
'If Not bijdang(1).sprite.Cell = 1 Then bijdang(1).sprite.RestoreBackground cStage.hdc
'bijdang(1).sprite.StoreBackground cStage.hdc, dorkx, dorky + smurg
'bijdang(1).sprite.TransparentDraw Form1.hDC, dorkx, dorky + offset, bijdang(1).sprite.cell, True
'bijdang(1).sprite.StageToScreen Form1.hdc, cStage.hdc
bijdang(1).sprite.cell = bijdang(1).sprite.cell + 1
End If

For a = 2 To UBound(bijdang)
If bijdang(a).sprite.cell < 11 Then
X = bijdang(a).X
Y = bijdang(a).Y
dorkx = (X - plr.X - Y + plr.Y) * 48 + 400 - plr.xoff - (bijdang(a).cS.CellWidth / 2) + xoff
'If dorkx < -40 Or dorkx > 750 Then Exit Sub
dorky = (X + Y - plr.X - plr.Y) * 24 - plr.yoff + 300 - (bijdang(a).cS.CellHeight - 48) - (midtile * 24) + yoff

dorkx = Int(dorkx)
dorky = Int(dorky) '+ smurg
bijdang(a).sprite.X = dorkx
bijdang(a).sprite.Y = dorky
If Not bijdang(a).sprite.cell = 1 Then bijdang(a).sprite.RestoreBackground Form1.hDC
bijdang(a).sprite.StoreBackground cStage.hDC, dorkx, dorky
'bijdang(a).sprite.TransparentDraw Form1.hDC, dorkx, dorky + offset, bijdang(a).sprite.cell, True
'bijdang(a).sprite.StageToScreen Form1.hdc, cStage.hdc
bijdang(a).sprite.cell = bijdang(a).sprite.cell + 1
End If
12 Next a

drawshooties

End Sub

Sub drawbijdang()
Dim a As Long 'm'' declare

'bijdrawing = bijdrawing + 1
'DoEvents

'If bijdrawing > 2 Then Exit Sub

If Form1.Visible = False Then bijdrawing = 0: Exit Sub

For a = 1 To 50
    If bijdangs(a).Active = 1 Then
    X = bijdangs(a).X
    Y = bijdangs(a).Y
    dorkx = (X - plr.X - Y + plr.Y) * 48 + 400 - plr.xoff - (bijdang(bijdangs(a).graphnum).cS.CellWidth / 2) + xoff
    'If dorkx < -40 Or dorkx > 750 Then Exit Sub
    dorky = (X + Y - plr.X - plr.Y) * 24 - plr.yoff + 300 - (bijdang(bijdangs(a).graphnum).cS.CellHeight - 48) - (midtile * 24) + yoff
    dorkx = Int(dorkx)

    dorky = Int(dorky) '+ smurg
    bijdang(bijdangs(a).graphnum).sprite.X = dorkx
    bijdang(bijdangs(a).graphnum).sprite.Y = dorky '- smurg * 10
    'If Not bijdang(1).sprite.Cell = 1 Then bijdang(1).sprite.RestoreBackground cStage.hdc
    'bijdang(1).sprite.StoreBackground cStage.hdc, dorkx, dorky + smurg
    
    'm'' i dont get the point of this line...
    bijdang(bijdangs(a).graphnum).sprite.TransparentDraw picBuffer, dorkx, dorky + offset, bijdangs(a).cell, True
    'bijdang(bijdangs(a).graphnum).sprite.TransparentDraw picBuffer, dorkx * 2, dorky * 2 + offset, bijdangs(a).cell, True
    
    bijdangs(a).cell = bijdangs(a).cell + 1
    If bijdangs(a).cell > bijdang(bijdangs(a).graphnum).sprite.cell Then bijdangs(a).Active = 0
    
    'drawobj bijdang(bijdangs(a).graphnum).cS, bijdangs(a).x + roll(6), bijdangs(a).y + roll(6), bijdangs(a).cell
    
    'bijdang(1).sprite.StageToScreen Form1.hDC, cStage.hDC
    'bijdang(1).sprite.cell = bijdang(1).sprite.cell + 1
    End If
    
   
Next a

drawshooties
'bijdrawing = bijdrawing - 1
blt Form1.Picture7

End Sub

Sub makebijdang(ByVal X As Integer, ByVal Y As Integer, ByVal wch As Integer)
'm'' declarations. bijdang are the animated sprites
Dim a As Long 'm''

If wch = 5 Then playsound "1cannon1.wav"
If wch = 6 Then playsound "1laser4.wav"
If wch = 7 Then playsound "1lightninggun.wav"

For a = 1 To 50
    If bijdangs(a).Active = 0 Then
    bijdangs(a).X = X
    bijdangs(a).Y = Y
    bijdangs(a).graphnum = wch
    bijdangs(a).cell = 1
    bijdangs(a).Active = 1
    Exit Sub
    End If
Next a

End Sub

Function isnaked() As Boolean 'm'' declare...
For a = 1 To 16
    If clothes(a).loaded = 1 Then isnaked = False: Exit Function
Next a
isnaked = True
End Function

Function greater(num1, num2)
If num1 > num2 Then greater = num1 Else greater = num2
End Function

Function lesser(num1, num2)
If num1 < num2 Then lesser = num1 Else lesser = num2
End Function

Function givespell(school, school2, spellmult, Optional abslevel = 0)
Dim a As Long 'm'' declare

If abslevel > 0 Then

    For a = 1 To UBound(spells)
        
        If spells(a).has = 0 Then If spells(a).school = school And spells(a).level <= abslevel Then spells(a).has = 1: gamemsg "You have learned " & spells(a).name: Form1.updatspells ': Exit Function
        'start with spells of your school and exits if you get one
        
    Next a
    
    GoTo 12
End If

For a = 1 To UBound(spells)
    
    If spells(a).has = 0 Then If spells(a).school = school And spells(a).level <= plr.level - (6 - spellmult * 2) Then spells(a).has = 1: gamemsg "You have learned " & spells(a).name: Form1.updatspells ': Exit Function
    'start with spells of your school and exits if you get one
    
Next a

For a = 1 To UBound(spells)

    If spells(a).has = 0 Then If spells(a).school = school2 And spells(a).level * 2 <= plr.level - (6 - spellmult) Then spells(a).has = 1: gamemsg "You have learned " & spells(a).name: Form1.updatspells ': Exit Function
    'Then goes to your secondary school

Next a

'Delete lower-level duplicates from the same school (For spell upgrades)
12
For a = 1 To UBound(spells)
    For b = 1 To UBound(spells)
        If a = b Then GoTo 5
        If spells(a).has And spells(b).has And spells(a).school = spells(b).school Then
            If getfromstring(spells(a).name, 1) = getfromstring(spells(b).name, 1) Then If Val(spells(a).level) > Val(spells(b).level) Then spells(b).has = 0 Else spells(a).has = 0
        End If
5     Next b
Next a


If Form1.Visible = True Then Form1.updatspells

'For a = 1 To UBound(spells)

'    If spells(a).has = 0 Then If spells(a).level < plr.level - (15 - spellmult) Then spells(a).has = 1: MsgBox "You have learned " & spells(a).name: Form1.updatspells: Exit Function
    'Then goes to whatever's left

'Next a

End Function

Function givespefspell(name)

For a = 1 To UBound(spells)
    If getfromstring(spells(a).name, 1) = name Then spells(a).has = 1: Form1.updatspells: Exit Function
Next a



End Function

Sub putobj(objc As objecttype)

    Put #1, , objc.name
    
    For a = 1 To 6
    For b = 1 To 5
    Put #1, , objc.effect(a, b): Next b: Next a
 '   graph As cSpriteBitmaps
 '   graphloaded As Byte '1 means the graphic has been loaded
    Put #1, , objc.graphname
    Put #1, , objc.r
    Put #1, , objc.g
    Put #1, , objc.b
    Put #1, , objc.l
    'Put #1, , objc.shove
    Put #1, , objc.cells

End Sub

Sub getobj(ByRef objc As objecttype)

    Get #1, , objc.name
    For a = 1 To 6
    For b = 1 To 5
    Get #1, , objc.effect(a, b): Next b: Next a
 '   graph As cSpriteBitmaps
 '   graphloaded As Byte '1 means the graphic has been loaded
    Get #1, , objc.graphname
    Get #1, , objc.r
    Get #1, , objc.g
    Get #1, , objc.b
    Get #1, , objc.l
    'Get #1, , objc.shove
    Get #1, , objc.cells

End Sub

Sub putclothes(ByRef clo As clothesT)

    Put #1, , clo.name
    'clo.loaded
    Put #1, , clo.wear1
    Put #1, , clo.wear2
    Put #1, , clo.hp
    Put #1, , clo.armor
    Put #1, , clo.drawn
    Put #1, , clo.loaded
    putobj clo.obj
    Put #1, , clo.digested

End Sub

Sub getclothes(ByRef clo As clothesT)

    Get #1, , clo.name
    'clo.loaded
    Get #1, , clo.wear1
    Get #1, , clo.wear2
    Get #1, , clo.hp
    Get #1, , clo.armor
    Get #1, , clo.drawn
    Get #1, , clo.loaded
    getobj clo.obj
    Get #1, , clo.digested

End Sub

Function updatmonz()

For a = 1 To mapx
For b = 1 To mapy
    map(a, b).monster = 0
Next b
Next a

For a = 1 To totalmonsters
'    On Error GoTo 15
 '   killmon a, 1
    If mon(a).type > lastmontype Or mon(a).X < 1 Or mon(a).X > mapx Or mon(a).Y < 1 Or mon(a).Y > mapy Then killmon a, 1: GoTo 3
    If mon(a).hp > 0 Then
    If mon(a).X > mapx Then killmon a: GoTo 3
    map(mon(a).X, mon(a).Y).monster = a
    Else:
    'map(mon(a).x, mon(a).y).monster = 0
    mon(a).type = 0
    End If
'15 killmon a, 1
3
Next a

End Function

Function countmonsters() As Long

For a = 1 To totalmonsters
    If mon(a).hp > 0 Then zab = zab + 1
Next a

countmonsters = zab

End Function

Function loadprefs()
'Load cheat levels, preferences etc.
If Dir("Settings.dat") = "" Then Exit Function
Open "Settings.dat" For Input As #1
    Do While Not EOF(1)
        Input #1, durg
        
        If durg = "#WURGBAG" Then
            Input #1, durg
            If durg = "I LIKE CHEESE" And winlev < 2 Then winlev = 1
            If durg = "I REALLY REALLY LIKE CHEESE" Then winlev = 2
        End If
        
        If durg = "#DUAM WUZ HEER" Then cheaton = 1
Loop
Close #1
End Function

Sub wingame()

'MsgBox "You strike Thirsha down, finally destroying her once and for all. As a reward for your ultimate studliness, you now have access to two new character classes. That's all the ending you get, until I think of a better one."
winlev = greater(winlev, plr.difficulty + 1)


If winlev > 0 Then
Open "Settings.dat" For Output As #1
    If winlev = 1 Then Write #1, "#WURGBAG", "I LIKE CHEESE", , ,
    If winlev >= 2 Then Write #1, "#WURGBAG", "I REALLY REALLY LIKE CHEESE", , ,
    Write #1,

Close #1
End If

showform10 "Win", 1, "Win1.jpg"

#If USELEGACY = 1 Then
End
#Else
'm'' let's continue the game if we won
rep = MsgBox("You've reached the original End of the game. Do you want to continue playing?", vbYesNo + vbInformation, "Duamutef's VRPG Modded edition") 'm''
If rep = vbNo Then Debugger.Quitting 'm''
#End If
End Sub

Sub dispatk(mont As monstertype, hp)

'Picture5 381
'Label2
'Dim mont As monstertype
'mont = montype(mon(monnum).type)
Form1.Picture5.Visible = True
Form1.Text6.Visible = True
Form1.Picture5.Cls
Form1.Picture5.Line (0, 0)-Step((hp / mont.hp) * 381, 25), RGB(180, 20, 0), BF
Form1.Text6.text = mont.name & " " & hp & "/" & mont.hp
Form1.Picture1.Refresh
End Sub

Sub checkclothes(num)
If clothes(num).digested = 0 Then
Dim backgraph As cSpriteBitmaps
If geteff(clothes(num).obj, "DIGESTED", 2) = "1" Then clothes(num).digested = 1
If clothes(num).digested = 1 Then Set backgraph = cgraphs(num).graph: Set cgraphs(num).graph = cgraphs(num).diggraph: Set cgraphs(num).diggraph = backgraph
Set backgraph = Nothing
End If

End Sub

Function summonmonster(monname, filen, level, r, g, b, l, boss, caneat, Optional maxmonsters = 0)

F = getsummons(monname)
If F = 0 Then
'    If monname = "Kari" Then filen = "kari1.bmp": r = 200: g = 50: b = 0: l = 0.6
    F = loadsummons(monname, level, filen, r, g, b, l, boss)
End If

5 X = plr.X + roll(7) - 4
Y = plr.Y + roll(7) - 4
If X > mapx Or X < 1 Then GoTo 5
If Y > mapx Or Y < 1 Then GoTo 5
If map(X, Y).blocked = 1 Or map(X, Y).monster > 0 Then GoTo 5
z = createmonster(F, X, Y)
mon(z).owner = caneat


If maxmonsters > 0 Then killextrasummons monname, maxmonsters

End Function

Function killextrasummons(monname, maxmonsters)

If maxmonsters = 0 Then Exit Function
curmons = 0
For a = totalmonsters To 1 Step -1
    If mon(a).type = 0 Then GoTo 5
    If montype(mon(a).type).name = monname Then curmons = curmons + 1: If curmons > maxmonsters Then killmon a
5 Next a

End Function

Function getsummons(monname) As Integer

For a = 1 To lastmontype
    If montype(a).name = monname Then getsummons = a: Exit Function
Next a

End Function

Function killsummons()

For a = 1 To totalmonsters
    If mon(a).owner > 0 Then killmon a
Next a

End Function


Function getrgb(ByVal color, ByRef r, ByRef g, ByRef b, Optional forcebitdepth = 0)

If display.ddpfPixelFormat.lRGBBitCount = 0 Then DD.GetDisplayMode display

If color = 0 Then getrgb = 0: Exit Function

If (display.ddpfPixelFormat.lRGBBitCount = 32 Or display.ddpfPixelFormat.lRGBBitCount = 24 Or display.ddpfPixelFormat.lRGBBitCount = 0 Or forcebitdepth = 32) And Not forcebitdepth = 16 Then
        
        r = color And 255
        g = (color And &HFF00&) \ 256
        b = (color And &HFF0000) \ 65536
        
        Else:
        
        b = color And display.ddpfPixelFormat.lBBitMask
        g = (color And display.ddpfPixelFormat.lGBitMask) \ ((display.ddpfPixelFormat.lBBitMask + 1) * 2) '256
        r = (color And display.ddpfPixelFormat.lRBitMask) \ display.ddpfPixelFormat.lGBitMask
        
        'Convert to 1-255 range
        mult = 255 \ display.ddpfPixelFormat.lBBitMask
        r = r * mult
        g = g * mult
        b = b * mult
                
End If

End Function
Function fixfuckingstrings()

For a = 1 To lastobj
    swaptxt objs(a).string, ",", "/"
Next a

End Function

Function genworld(worldname, maxx, maxy, Optional maxlevel = 50) As String

'ReDim genmaps(1 To maxx, 1 To maxy) As String
'For a = 1 To maxx / 2
'    For b = 1 To maxy / 2
'        genmaps(a + maxx / 2, b + maxy / 2) = genmap(worldname, a + b * 3 + roll(3), a + maxx / 2, b + maxy / 2, maxlevel)
'        genmaps(maxx / 2 - x, maxy / 2 - b) = genmap(worldname, a + b * 3 + roll(3), a + maxx / 2, b + maxy / 2, maxlevel)
'    Next b
'Next a


End Function

Function genmap(ByVal worldname As String, ByVal level As Byte, ByVal X, ByVal Y, Optional mapnorth = "", Optional mapeast = "", Optional mapsouth = "", Optional mapwest = "", Optional same = 0, Optional maxlevel = 50) As String
'Generate a random map--returns the map name
worlddir = worldname

Static prefix, suffix, suffix2 As String
Static basetile, tile1, tile2, tile3, ovr1, ovr2, sprinkle1, sprinkle2 As Byte
Dim r, g, b As Integer
Dim mons(12) As monstertype
Static lastfile

'If same > 0 Then GoTo 5

tile1 = 0: tile2 = 0: tile3 = 0: ovr1 = 0: ovr2 = 0: sprinkle1 = 0: sprinkle2 = 0

aroll = roll(8)
basetile = gettilenum
Select Case aroll
    Case 1: prefix = "Muddy ": basetile = 12
    Case 2: prefix = "Dark ": basetile = 14
    Case 3: prefix = "Grassy ": basetile = 1
    Case 4: prefix = "Sandy ": basetile = 2
End Select

aroll = roll(6)
Select Case aroll
    Case 1: suffix = "swamp": tile1 = 9: tile2 = 12: sprinkle1 = 8: sprinkle2 = 17: ovr1 = 8: If roll(3) = 1 Then ovr2 = 11
    Case 2: suffix = "desert": tile1 = 2: tile2 = 2: sprinkle1 = 5: If roll(2) = 1 Then sprinkle2 = 8
    Case 3: suffix = "mountain": tile1 = 3: tile2 = 15: sprinkle1 = 5: sprinkle2 = 14: ovr1 = 9: ovr2 = 10
    Case 4: suffix = "plains": tile1 = 3: tile2 = 12: sprinkle1 = 5: sprinkle2 = 17
    Case 5: suffix = "land": tile1 = gettilenum: tile2 = gettilenum: sprinkle1 = roll(20): sprinkle2 = roll(20): ovr1 = roll(20): ovr2 = roll(20)
    Case 6: suffix = "forest": tile1 = 1: tile2 = 12: sprinkle1 = getplantnum: sprinkle2 = getplantnum: ovr1 = getplantnum: ovr2 = getplantnum
    Case 6: suffix = "bog": tile1 = 9: tile2 = 13
End Select

aroll = roll(6)

Select Case aroll
    Case 1: suffix2 = " of beasts"
    Case 2: suffix2 = " of sorrow"
    Case 3: suffix2 = " of pain"
    Case 4: suffix2 = " of stomachs"
    Case 5: suffix2 = " of goo"
    Case 6: suffix2 = " of terror"
    Case 7: suffix2 = " of acid"
    Case 8: suffix2 = " of monsters"
    Case 9: suffix2 = " of death"
    Case 10: suffix2 = " of digestion"
    Case 10: suffix2 = " of chaos": tile1 = roll(19): tile2 = roll(19)
End Select


aroll = roll(6)

Select Case aroll
    Case 1:
End Select





tile3 = gettilenum

5 If roll(6) > 4 Then tile3 = gettilenum
If Dir(App.Path & "\" & worldname, vbDirectory) = "" Then MkDir App.Path & "\" & worldname
ChDir App.Path & "\" & worldname

lname = prefix & suffix & suffix2

lastfile = lastfile + 1
mnum = 1
7
If Not Dir(lname & " " & mnum & ".txt") = "" Then mnum = mnum + 1: GoTo 7
lname = lname & " " & mnum

genmaps(X, Y) = lname

Open lname & ".txt" For Output As #lastfile

Write #lastfile, "Random map file " & lname
Write #lastfile, "---"
Write #lastfile, "#TERSTR", "Now entering " & lname

Write #lastfile, "#RANDOMSEED", roll(100000)
If roll(3) = 1 Then Write #lastfile, "#MAPSIZE", roll(8) * 25, roll(8) * 25 Else Write #lastfile, "#MAPSIZE", 100, 100

Write #lastfile, "#FILLMAP", basetile

For a = 1 To roll(5)
    If tile1 > 0 Then Write #lastfile, "#MAPCHUNK", tile1, roll(15) + 10
Next a

For a = 1 To roll(5)
    If tile2 > 0 Then Write #lastfile, "#MAPCHUNK", tile2, roll(15) + 10
Next a

For a = 1 To roll(5)
    If ovr1 > 0 Then Write #lastfile, "#OVRCHUNK", ovr1, roll(15) + 10
Next a

For a = 1 To roll(5)
    If ovr2 > 0 Then Write #lastfile, "#OVRCHUNK", ovr2, roll(15) + 10
Next a

For a = 1 To roll(5)
    If sprinkle1 > 0 Then Write #lastfile, "#SPRINKLEOVR", roll(15) * 20, sprinkle1
Next a

For a = 1 To roll(5)
    If sprinkle2 > 0 Then Write #lastfile, "#SPRINKLEOVR", roll(15) * 20, sprinkle2
Next a

Write #lastfile,
   
For a = 1 To 12
    mons(a).name = ""
Next a

For a = 1 To 12
    If a > 3 Then If roll(6) < 3 Then Exit For
    
    aroll = roll(32)
    Select Case aroll
        Case 1: mons(a).name = "Kari": mons(a).gfile = "kari1.bmp"
        Case 2: mons(a).name = "Kari": mons(a).gfile = "kari1.bmp"
        Case 3: mons(a).name = "Giantess": mons(a).gfile = "giantess.bmp"
        Case 4: mons(a).name = "Naga": mons(a).gfile = "snakewoman1.bmp"
        Case 4: mons(a).name = "Girlserpent": mons(a).gfile = "snakewoman1.bmp"
        Case 5: mons(a).name = "Snakewoman": mons(a).gfile = "snakewoman2.bmp"
        Case 6: mons(a).name = "Snake": mons(a).gfile = "snake1.bmp"
        Case 7: mons(a).name = "Serpent": mons(a).gfile = "snake1.bmp"
        Case 7: mons(a).name = "Python": mons(a).gfile = "snake1.bmp"
        Case 8: mons(a).name = "Worm": mons(a).gfile = "worm1.bmp"
        Case 9: mons(a).name = "Vines": mons(a).gfile = "tendrils1.bmp"
        Case 10: mons(a).name = "Tentacles": mons(a).gfile = "tendrils1.bmp"
        Case 11: mons(a).name = "Plant": mons(a).gfile = "venus1.bmp"
        Case 12: mons(a).name = "Trap": mons(a).gfile = "venus1.bmp"
        Case 13: mons(a).name = "Frog": mons(a).gfile = "frog1.bmp"
        Case 14: mons(a).name = "Toad": mons(a).gfile = "frog1.bmp"
        Case 15: mons(a).name = "Slime": mons(a).gfile = "slime1.bmp"
        Case 16: mons(a).name = "Goo": mons(a).gfile = "slime1.bmp"
        Case 17: mons(a).name = "Centauress": mons(a).gfile = "centauress1.bmp"
        Case 18: mons(a).name = "Mare": mons(a).gfile = "centauress1.bmp"
        Case 19: mons(a).name = "Harpy": mons(a).gfile = "harpy1.bmp"
        Case 20: mons(a).name = "Raven": mons(a).gfile = "harpy1.bmp"
        Case 21: mons(a).name = "Sprite": mons(a).gfile = "sprite1.bmp"
        Case 22: mons(a).name = "Faerie": mons(a).gfile = "sprite1.bmp"
        Case 23: mons(a).name = "Succubus": mons(a).gfile = "succubus1.bmp"
        Case 24: mons(a).name = "Demoness": mons(a).gfile = "succubus1.bmp"
        Case 25: mons(a).name = "Dragon": mons(a).gfile = "thirsha1.bmp"
        Case 27: mons(a).name = "Winged Serpent": mons(a).gfile = "wingedserpent1.bmp"
        Case 28: mons(a).name = "Snakebat": mons(a).gfile = "wingedserpent1.bmp"
        Case 29: mons(a).name = "Alligator": mons(a).gfile = "croc2.bmp"
        Case 30: mons(a).name = "Crocodile": mons(a).gfile = "croc2.bmp"
        Case 31: mons(a).name = "Spiderwoman": mons(a).gfile = "kari1.bmp"
        Case 32: mons(a).name = "Grub": mons(a).gfile = "worm1.bmp"
    End Select
    
    If roll(4) = 1 And a > 1 Then mons(a).name = mons(a - 1).name: mons(a).gfile = mons(a - 1).gfile
    
    aroll = roll(23)
    Select Case aroll
        Case 1: mpref = "Sea ": r = 10: g = 80: b = 200
        Case 2: mpref = "Blue ": r = 10: g = 10: b = 200
        Case 3: mpref = "Grey ": r = 100: g = 100: b = 100
        Case 4: mpref = "Black ": r = 20: g = 20: b = 20
        Case 5: mpref = "White ": r = 210: g = 210: b = 210
        Case 6: mpref = "Fire ": r = 210: g = roll(150): b = 10
        Case 7: mpref = "Red ": r = 210: g = 50: b = 10
        Case 8: mpref = "Death ": r = 40: g = 40: b = 40
        Case 9: mpref = "Dark ": r = 30: g = 30: b = 400
        Case 10: mpref = "Jade ": r = 40: g = 140: b = 20
        Case 11: mpref = "Forest ": r = 10: g = 140: b = 10
        Case 12: mpref = "Golden ": r = 180: g = 140: b = 10
        Case 13: mpref = "Stone ": r = 140: g = 130: b = 120
        Case 14: mpref = "Mud ": r = 100: g = 70: b = 10
        Case 15: mpref = "Fae ": r = 210: g = 40: b = 210
        Case 16: mpref = "Arcane ": r = 120: g = 10: b = 160
        Case 17: mpref = "Death ": r = 40: g = 40: b = 40
        Case 18: mpref = "Marsh ": r = 20: g = 140: b = 100
        Case 19: mpref = "Belly ": r = 200: g = 140: b = 10
        Case 20: mpref = "Whore ": r = roll(200): g = roll(200): b = roll(200)
        Case 21: mpref = "Bitch ": r = roll(200): g = roll(200): b = roll(200)
        Case 22: mpref = "Magic ": r = roll(200): g = roll(200): b = roll(200)
        Case 23: mpref = "Sword ": r = roll(200): g = roll(200): b = roll(200)
    End Select
    
    'mons(a).name = mpref & mons(a).name
    
    r = r + (roll(60) - 30): If r < 0 Then r = 0
    g = g + (roll(60) - 30): If g < 0 Then g = 0
    b = b + (roll(60) - 30): If b < 0 Then b = 0
    
    mons(a).color = RGB(r, g, b)
    mons(a).light = 0.5 + ((roll(7) - 4) / 10)
    lev = level + roll(7) - 4: If lev < 1 Then lev = 1
    mons(a).level = lev

    If Not mons(a).name = "" Then Write #lastfile, "#MONTYPE2", mpref & mons(a).name, mons(a).gfile, mons(a).level, r, g, b, mons(a).light
Next a

For a = 1 To roll(20) + 4
    If roll(4) = 1 Then
    Write #lastfile, "#BUILDING", 3 + roll(5), 3 + roll(5), 0, 0, basetile, getwallnum, "ARMOR" & Int(roll(3) + level / 2)
    Else:
    Write #lastfile, "#BUILDING", 3 + roll(5), 3 + roll(5), 0, 0, basetile, getwallnum, "TREASURE" & Int(roll(3) + level / 2)
    End If
Next a

Write #lastfile, "#RANDOMMONSTERS", "ALL", 500 + roll(500), 0, 0, 0, 0

If mapy < UBound(genmaps(), 2) Then mapnorth = genmaps(mapx, mapy + 1)
If mapx < UBound(genmaps(), 1) Then mapeast = genmaps(mapx + 1, mapy)
If mapx > 1 Then mapwest = genmaps(X - 1, Y)
If mapy > 1 Then mapsouth = genmaps(X, Y - 1)

If level < maxlevel And X < UBound(genmaps(), 1) And Y < UBound(genmaps(), 2) And X > 1 And Y > 1 Then
    If mapnorth = "" Then mapnorth = genmap(worldname, level + 4 + roll(6), X, Y + 1, , , lname & ".txt", , roll(1), maxlevel)
    If mapsouth = "" Then mapsouth = genmap(worldname, level + 4 + roll(6), X, Y - 1, lname & ".txt", , , , roll(1), maxlevel)
    If mapeast = "" Then mapeast = genmap(worldname, level + 4 + roll(6), X + 1, Y, , , , lname & ".txt", roll(1), maxlevel)
    If mapwest = "" Then mapwest = genmap(worldname, level + 4 + roll(6), X - 1, Y, , lname & ".txt", , , roll(1), maxlevel)
End If

Write #lastfile, "#SETMAPS", mapnorth, mapeast, mapsouth, mapwest

Write #lastfile,
Write #lastfile,

Close #lastfile
lastfile = lastfile - 1
genmap = worldname & "\" & lname & ".txt"
ChDir App.Path
End Function

Function getplantnum()

gurp = roll(9) + 11
If gurp = 13 Then gurp = 2
If gurp = 14 Then gurp = 8
If gurp = 15 Then gurp = 11
getplantnum = gurp
End Function

Function gettilenum()

5 gurp = roll(17)
If gurp = 4 Or gurp = 5 Or gurp = 6 Or gurp = 7 Or gurp = 10 Or gurp = 11 Then GoTo 5
'4 5 6 7 10 11
gettilenum = gurp

End Function

Function getwallnum()

5 gurp = roll(13)
If gurp = 8 Or gurp = 5 Or gurp = 11 Or gurp = 2 Then GoTo 5
'4 5 6 7 10 11
getwallnum = gurp

End Function

Function getplrskill(skillname, Optional addbonus = 1) As Integer

For a = 0 To 50
    If plr.specials(a, 0) = skillname Then getplrskill = Val(plr.specials(a, 1)) + (getbonusskill(skillname) * addbonus): Exit Function
    If plr.specials(a, 0) = "" Then
    getplrskill = getplrskill + (getbonusskill(skillname) * addbonus)
    Exit Function
    End If
Next a

End Function

Function skillmod(ByRef var, skillname, firstinc, levelinc)

If getplrskill(skillname) > 0 Then var = var + firstinc + (levelinc * (getplrskill(skillname) - 1))

End Function

Function skillpercmod(ByRef var, skillname, firstinc, levelinc, Optional skillreq = 0)
'Skillreq subtracts from the effective skill level

If getplrskill(skillname) > 0 Then var = Int(var * (1 + (firstinc + (levelinc * (getplrskill(skillname) - 1 - skillreq))) / 100))

End Function

Function skillminus(ByRef var, skillname, firstinc, levelinc)

If getplrskill(skillname) > 0 Then var = var - firstinc - (levelinc * (getplrskill(skillname) - 1))

End Function

Function skillpercminus(ByRef var, skillname, firstinc, levelinc)
If Val(var) = 0 Then Exit Function
If getplrskill(skillname) > 0 Then var = Int(var * (1 - ((firstinc + (levelinc * (getplrskill(skillname) - 1))) / 100)))

End Function

Function skilltotal(skillname, firstinc, levelinc) As Integer
If getplrskill(skillname) > 0 Then skilltotal = firstinc + (levelinc * getplrskill(skillname) - 1)
End Function

Function skilltotal2(skillname, firstinc, levelinc) As Single
If getplrskill(skillname) > 0 Then skilltotal2 = firstinc + (levelinc * getplrskill(skillname) - 1)
End Function

Function addskill(skillname, Optional amt = 1)

'If a player has a skill at 0, it will show up in the skill screen.

For a = 0 To 30
    If plr.specials(a, 0) = skillname Then plr.specials(a, 1) = Val(plr.specials(a, 1)) + amt: Exit For
    If plr.specials(a, 0) = "" Then plr.specials(a, 1) = amt: plr.specials(a, 0) = skillname: Exit For
Next a

givespell skillname, "None", 0, Val(plr.specials(a, 1))

End Function

Function totalskills() As Integer
Dim a As Long, b As Long, tot As Long  'm'' declares
tot = 0
For a = 0 To 30
    tot = tot + Val(plr.specials(a, 1)) * 2
    For b = 1 To 6
        If Val(plr.specials(a, 1)) > 0 And plr.classskills(b) = plr.specials(a, 0) Then tot = tot - Val(plr.specials(a, 1))
    Next b
Next a
totalskills = tot
End Function

Function ifclassskill(skillname) As Boolean
ifclasskill = False
For a = 1 To 6
If plr.classskills(a) = skillname Then ifclassskill = True: Exit Function
Next a

End Function

Function checklevels(ByVal filen) As Integer
'Returns the level of the lowest level monster in a map file
'If filen = "" Then Exit Function
Exit Function
lev = 10000

If Left(filen, 6) = "VTDATA" Then filen = Right(filen, Len(filen) - 6)
origfile = filen
filen = getfile(filen, "Data.pak")
If filen = "" Then MsgBox "Error #1 in checklevels function": Exit Function
Open filen For Input As #1
    Do While Not EOF(1)
        Input #1, durg
                    
1            If durg = "#MONTYPE2" Then
            If lastmontype > 20 Then Close #1: GoTo 15
            'lastmontype = lastmontype + 1
            'ReDim Preserve montype(1 To lastmontype): ReDim Preserve mongraphs(1 To UBound(montype())) As cSpriteBitmaps
            Input #1, a, b, monlev
            If monlev < lev Then lev = monlev
            End If

    Loop

Close #1

'Default to level 0
15 If lev = 10000 Then lev = 0
checklevels = lev
End Function

Function gamemsg(ByVal txt As String)  'm'' declares
Dim tmp As String 'm'' local buffer (faster than the form object)
tmp = Left$(Form1.Text7.text, 5000) 'm'' added $ to use Left$() which is faster than Left()
Form1.Text7.text = txt & vbCrLf & tmp 'm'' use buffer
End Function

Function stdfilter(txt) As String

txt = swaptxt(txt, "/", ",")
txt = swaptxt(txt, "$NAME", plr.name)
txt = swaptxt(txt, "$VERSION", curversion)

'Don't put any other filters that use $'s after this one
txt = swaptxt(txt, "$", Chr(34))
stdfilter = txt

End Function

Function getobjtype(name) As objecttype
Dim a As Long 'm'' declare
For a = 1 To UBound(objtypes())
    If objtypes(a).name = name Then getobjtype = objtypes(a): Exit Function
Next a

End Function


'Sub buy(objnum, buyt, dobjtype As objecttype) 'As objtype

'wobj = objnum
'buytype = buyt

'Dim bobjt As objecttype


'lasobj = 1
'bobjt = objtypes(objs(objnum).type)
'If objnum > 0 Then Form2.msg objs(objnum).name, objs(objnum).string
'If objnum = 0 Then bobjt = dobjtype: Form2.msg dobjtype.name

'Form2.Command2(1).Visible = True
'Form2.Height = Form2.Height + (Form2.Command2(1).Height + 40) * 8



'If buyt = "Armor" Then worth = Val(geteff(bobjt, "SELLARMOR", 2)): sellnum = Val(geteff(bobjt, "SELLARMOR", 3)): Form2.rarmor: GoTo 5
'If buyt = "Clothes" Then worth = Val(geteff(bobjt, "SELLCLOTHES", 2)): sellnum = Val(geteff(bobjt, "SELLCLOTHES", 3)): Form2.rclothes: GoTo 5
'If buyt = "Weapons" Then worth = Val(geteff(bobjt, "SELLWEAPONS", 2)): sellnum = Val(geteff(bobjt, "SELLWEAPONS", 3)): Form2.rweps: GoTo 5

'If buyt = "Potions" Then
'    goldam(1) = (plr.lpotionlev + 1 * plr.lpotionlev + 1) * (10 * plr.lpotionlev): Form2.Command2(1).Visible = True: Form2.Command2(1).caption = "Life Potions " & goldam(1)
'    goldam(2) = (plr.mpotionlev + 1 * plr.mpotionlev + 1) * (10 * plr.mpotionlev): Form2.Command2(2).Visible = True: Form2.Command2(2).caption = "Mana Potions " & goldam(2)
'    For a = 3 To 8: Command2(a).Visible = False: Next a
'End If

'For a = 1 To 6
'    If bobjt.effect(a, 1) = "SELL" Then Command2(lasobj).caption = bobjt.effect(a, 2) & ", " & bobjt.effect(a, 3) & " Gold": Command2(lasobj).Visible = True: sellobjs(lasobj) = bobjt.effect(a, 2): goldam(lasobj) = bobjt.effect(a, 3): lasobj = lasobj + 1
'Next a
'5
'For a = 1 To 8
'Form2.Command2(a).Top = Label1.Top + Label1.Height + (a * Form2.Command2(a).Height)
'Form2.Command2(a).Left = Form2.Width / 2 - Form2.Command2(a).Width / 2
'Next a

'Form2.Top = Screen.Height / 2 - Form2.Height / 2
'Form2.Command1.caption = "EXIT"
'End Sub

Sub randword(ByVal word As String, Optional ByVal add = 0)
'Alters the random seed according to a word
'Will fuck up with two and possibly three letter words, I'm guessing
If word = "" Then word = "RANDOM"
Rnd (-5)
If Val(word) = 0 Then Randomize Asc(word) + Asc(Right(word, 1)) + Asc(Mid(word, Len(word) / 2, 1)) + add Else Randomize word + Val(add)

End Sub

Function gencolor(Optional ByVal colorname As String = "", Optional ByRef r, Optional ByRef g, Optional ByRef b, Optional ByRef l = 0.5) As String
If colorname = "0" Then colorname = ""
If Not colorname = "" Then If r > 0 Or b > 0 Or g > 0 And l > 0 Then gencolor = colorname: Exit Function
If colorname = "" Then
aroll = roll(17)

Select Case aroll
    Case 1: colorname = "Red" ': r = 230: g = 0: b = 0
    Case 2: colorname = "Blue" ': r = 0: g = 0: b = 210
    Case 3: colorname = "Green" ': r = 0: g = 150: b = 0
    Case 4: colorname = "Black" ': r = 20: g = 20: b = 20
    Case 5: colorname = "Orange" ': r = 250: g = 180: b = 0
    Case 6: colorname = "Yellow" ': r = 250: g = 240: b = 0
    Case 7: colorname = "Grey" ': r = 100: g = 100: b = 100
    Case 8: colorname = "Jade" ': r = 30: g = 100: b = 0
    Case 9: colorname = "Brown" ': r = 100: g = 40: b = 0
    Case 10: colorname = "Pink" ': r = 250: g = 0: b = 200: l = 1
    Case 11: colorname = "Purple" ': r = 130: g = 0: b = 180
    Case 12: colorname = "Crimson" ': r = 210: g = 10: b = 5
    Case 13: colorname = "White" ': r = 240: g = 240: b = 240: l = 0.8
    Case 14: colorname = "Sky Blue" ': r = 10: g = 120: b = 250
    Case 15: colorname = "Hot Pink" ': r = 250: g = 0: b = 200
    Case 16: colorname = "Levi Blue" ': r = 0: g = 133: b = 210
    Case 17: colorname = "Black" ': r = 20: g = 20: b = 20
End Select
End If
l = 0.5
Select Case colorname
    Case "Red": r = 230: g = 0: b = 0
    Case "Blue": r = 0: g = 0: b = 210
    Case "Green": r = 0: g = 150: b = 0
    Case "Black": r = 40: g = 40: b = 40
    Case "Orange": r = 250: g = 180: b = 0
    Case "Yellow": r = 250: g = 240: b = 0
    Case "Grey": r = 100: g = 100: b = 100
    Case "Jade": r = 30: g = 100: b = 0
    Case "Brown": r = 100: g = 40: b = 0
    Case "Pink": r = 250: g = 0: b = 200: l = 1
    Case "Purple": r = 130: g = 0: b = 180
    Case "Crimson": r = 210: g = 10: b = 5
    Case "White": r = 240: g = 240: b = 240: l = 0.8
    Case "Sky Blue": r = 10: g = 120: b = 250: l = 0.8
    Case "Hot Pink": r = 250: g = 0: b = 200: l = 0.7
    Case "Levi Blue": r = 0: g = 133: b = 210
End Select

gencolor = colorname


End Function

Function lowestlevel() As Integer
'Actually returns the average level of surrounding maps...so sue me.

'If mapjunk.level > 0 Then lowestlevel = mapjunk.level
lowestlevel = mapjunk.level: Exit Function
murg = 1
tot = 0
zap1 = checklevels(mapjunk.maps(1)): If zap1 > 0 Then murg = murg + zap1: tot = tot + 1
zap2 = checklevels(mapjunk.maps(2)): If zap2 > 0 Then murg = murg + zap2: tot = tot + 1
zap3 = checklevels(mapjunk.maps(3)): If zap3 > 0 Then murg = murg + zap3: tot = tot + 1
zap4 = checklevels(mapjunk.maps(4)): If zap4 > 0 Then murg = murg + zap4: tot = tot + 1
If tot = 0 Then lowestlevel = mapjunk.level: Exit Function
mapjunk.level = greater(Int(murg / tot), averagelevel)
lowestlevel = Int(murg / tot)

End Function

Function averagelevel()
'Returns the average level of monsters in the current map
zap1 = 0
For a = 1 To UBound(montype())
    zap1 = zap1 + Val(montype(a).level)
Next a

If a > 0 Then zap1 = zap1 / (a - 1)
If zap1 = 0 Then zap1 = lowestlevel
averagelevel = zap1

If averagelevel = 0 Then averagelevel = lowestlevel

End Function

Function potioncost(lev) As Integer

'If isexpansion = 1 Then potioncost = plr.level * 10: Exit Function
potioncost = plr.level * 5
'potioncost = lev * 10 * 3 * lev

End Function

Function distortchar(destsurf As cSpriteBitmaps, Optional Width As Integer = 10, Optional noupdt As Integer = 0, Optional centery As Integer = 50)
'Dim centery As Integer
Dim centerx As Integer
centery = 45 '50
centerx = 158
If Width < 30 Then centery = centery + ((30 - Width) / 2)
If noupdt = 0 Then Form1.updatbody
If plr.Class = "Naga" Then centery = centery + 80: centerx = centerx - 20: Width = Width + 10

destsurf.distort centerx, centery, 15, Width + 30

'distort centerx, centery, 15, 30 + Width, Form1.Picture1

'Form1.Picture1.Picture = Form1.Picture1.image
'selfassign Form1.Picture1

'If noupdt < 2 Then Form1.Picture2.Width = 212
'If noupdt < 2 Then Form1.Picture2.PaintPicture Form1.Picture1.Picture, 0, 0, 56, 109

'Form1.Picture2.Picture = Form1.Picture2.image
'selfassign Form1.Picture2
'Set plrgraphs = New cSpriteBitmaps

'plrgraphs.CreateFromPicture Form1.Picture2.Picture, 1, 1, , RGB(0, 0, 0)
'If noupdt = 0 Then Form1.updatbody

End Function

Function distortboobs(Optional Width As Integer = 10, Optional noupdt As Integer = 0)
Dim centery As Integer
Dim centerx As Integer
centery = 50
centerx = 158
'If Width < 30 Then centery = centery + ((30 - Width) / 2)
If noupdt = 0 Then Form1.updatbody
If plr.Class = "Naga" Then centery = centery + 80: centerx = centerx - 20

distort centerx, centery, Width + 15, 15, Form1.Picture1

'Form1.Picture1.Picture = Form1.Picture1.image
selfassign Form1.Picture1

Form1.Picture2.Width = 212
Form1.Picture2.PaintPicture Form1.Picture1.Picture, 0, 0, 56, 109

'Form1.Picture2.Picture = Form1.Picture2.image
selfassign Form1.Picture2

Set plrgraphs = New cSpriteBitmaps
'plrgraphs.CreateFromPicture Form1.Picture2.Picture, 1, 1, , RGB(0, 0, 0)
If noupdt = 0 Then Form1.updatbody

End Function

Function digestfood()

#If USELEGACY <> 1 Then 'm''
    If Form1.menu_mod(7).Checked = True Then Exit Function 'm'' do not digest
#End If 'm''

If plr.foodinbelly > 0 And plr.foodinbelly > plr.monsinbelly Then
plr.foodinbelly = plr.foodinbelly - 1: plr.hp = plr.hp + (plr.hpmax * roll(20) / 100)
Form1.updatbody
'distortchar fullbody, plr.foodinbelly * 5
playsound "burp" & roll(5) & ".wav"
playsound "stomach" & roll(4) & ".wav"
displasteatstr
If plr.foodinbelly = 0 Then Form1.updatbody: playsound "fart57.wav": createobj "Shit", plr.X, plr.Y: gamemsg "Your " & getbelly("") & " has diligently " & getabsorb & "ed everything you have put in it."
End If


End Function

Function eatobj(dobj, Optional noneedstomach = 0)
If Not geteff(objtypes(objs(dobj).type), "NoEat", 1) = "" Then gamemsg "You cannot eat that.": eatobj = False: Exit Function
If noneedstomach = 0 Then If skilltotal("Giant Stomach", 1, 1) = 0 Then playsound "full5.wav": Ccom = "": Exit Function
If noneedstomach = 0 Then If plr.foodinbelly >= skilltotal("Giant Stomach", 1, 1) Then randsound "full", 5: Ccom = "": gamemsg "Your stomach is too full to eat that!": eatobj = False: Exit Function
If Not geteff(objtypes(objs(dobj).type), "NPC", 1) = "" Then eatchar = 1
If Not geteff(objtypes(objs(dobj).type), "Conversation", 1) = "" Then eatchar = 1
name = objs(dobj).name
If eatchar = 0 Then gamemsg "You swallow " & objs(dobj).name & " and feel it slithering down into your belly."
If eatchar = 1 Then
If noneedstomach = 0 Then If MsgBox("Are you sure you want to swallow " & name & "!?", vbYesNo) = vbNo Then Ccom = "": Exit Function
gamemsg "You swallow poor " & objs(dobj).name & " alive and whole.  You feel her squirming in surprise in your stretched belly.": plr.skillpoints = plr.skillpoints + 1
atenpc = 1
Select Case roll(1)
Case 1: addeatstr "You can feel " & name & " " & getwrithe("ing") & " in your " & getbelly(".")
End Select

Select Case roll(6)
Case 1: addeatstr "Judging by the intestinal tremblor you just felt, " & name & " was just digested."
Case 2: addeatstr "Oh, feels like " & name & " is finishing up her inside tour of your " & getbelly("") & "..."
Case 3: addeatstr "Feels like you just " & getabsorb & "ed " & name & "."
Case 4: addeatstr "Apparently your " & getbelly("") & " is finished " & getdigest("ing ") & name & "."
Case 5: addeatstr "Whoopsie!  So much for " & name & "..."
Case 6: addeatstr "Oops!  Was that " & name & "?"
End Select

playsound "swallowed" & roll(3) & ".wav"
plr.foodinbelly = plr.foodinbelly + 1
'Form1.updatbody
End If

playsound "swallow1.wav"
map(objs(dobj).X, objs(dobj).Y).object = 0
objs(dobj).name = ""
objs(dobj).X = 0
objs(dobj).Y = 0
objs(dobj).type = 0
plr.foodinbelly = plr.foodinbelly + 1

If atenpc = 1 Then ifnpcsaw
Form1.updatbody
'distortchar plr.foodinbelly * 5
Ccom = ""
eatobj = True
End Function

Function addeatstr(str)

For a = 0 To 10
    If ate(a) = "" Then ate(a) = str: Exit Function
Next a

End Function

Function displasteatstr()

If ate(0) = "" Then ate(0) = "You just " & getabsorb & "ed something.  You've eaten so much you're not even sure what it is."
gamemsg ate(0)
ate(0) = ""

For a = 0 To 9
    If ate(a + 1) = "" Then Exit For
    swap ate(a), ate(a + 1)
Next a

End Function

Sub swap(ByRef a, ByRef b)
c = b
b = a
a = c
End Sub

Function plrdigmon(ByRef mon As amonsterT, ByVal monnum As Integer, Optional force = 0)

Dim mont As monstertype
mont = montype(mon.type)
'damagemon a, rolldice(plr.level, 6)
'mon.x = plr.x: mon.y = plr.y
movemon monnum, 0, 0

#If USELEGACY <> 1 Then 'm''
    If force = 0 And Form1.menu_mod(7).Checked = True Then Exit Function 'm'' do not digest
#End If 'm''

If turncount Mod 15 = 0 Or force > 0 Then
    playsound "stomach" & roll(4) & ".wav"
     dmg = rolldice(plr.level, 8 + skilltotal("Super Acid", 2, 1))
     mon.hp = mon.hp - dmg
     plrdamage -(Int(dmg / 5)), 0
     gamemsg "You have " & getabsorb & "ed " & dmg & " hit points from the " & mont.name & " and gained " & Int(dmg / 5) & " from it."
    If mon.hp > 0 Then playsound "stomach" & roll(4) & ".wav"
    If mon.hp <= 0 Then
    gamemsg "You have utterly digested the " & mont.name & "!" '  It's strength is now yours."
    #If USELEGACY = 1 Then
        createobj "Shit", plr.X, plr.Y
    #Else
        'm'' I always think it is kinky to be able to poo outside while inside a monster
        If plr.instomach = 0 Then 'm''
            createobj "Shit", plr.X, plr.Y 'm''
        Else 'm''
            'm'' the digested monster will be killed, but the bijdang animation must be rightly located
            mon.X = VRPG.mon(plr.instomach).X 'm'' force position to be current pred's
            mon.Y = VRPG.mon(plr.instomach).Y 'm''
        End If 'm''
    #End If
    If ifmonstereaten(montype(mon.type).name) = False Then gamemsg "You have gained a skill point from digesting a new monster!"
    playsound "burp" & roll(5) & ".wav"
    plrdamage -(mont.hp / 10)
    plr.exp = plr.exp + mont.exp
    If Val(mont.level) > plr.level Then plr.exp = plr.exp + mont.exp * 2
    movemon monnum, plr.X, plr.Y
    killmon monnum
    If plr.monsinbelly > 0 Then plr.monsinbelly = plr.monsinbelly - 1
    If plr.foodinbelly > 0 Then plr.foodinbelly = plr.foodinbelly - 1
    playsound "fart22.wav"
    Form1.updatbody
    Exit Function
    End If

    mroll = succroll(Val(mont.level) + ((mon.hp / mont.hp) * 10), 7): aroll = succroll(plr.level + getstr, 6)
    If mroll > aroll Then
        gamemsg "The " & mont.name & " has forced it's way out of your stomach!": playsound "grunt" & roll(9) + 2 & ".wav": playsound "burp" & roll(5) & ".wav": mon.instomach = 0
        'playsound "escape" & roll(4) & ".wav"
        randsound "escape", 4
        plr.monsinbelly = plr.monsinbelly - 1
        plr.foodinbelly = plr.foodinbelly - 1
        mon.X = plr.X: mon.Y = plr.Y
        'm'' very rare : you swallowed a monster, get swallowed, and then the monster escape your stomach, ending in the other monster stomach
        If plr.instomach > 0 And plr.Swallowtime > 2 Then 'm''
            mon.instomach = plr.instomach 'm''
            VRPG.mon(plr.instomach).hasinstomach = monnum 'm''
            gamemsg mont.name & " is squeezed against you inside " & montype(VRPG.mon(plr.instomach).type).name & "'s innards!"
        Else 'm''
            monmove monnum, roll(3) - 2, roll(3) - 2
        End If 'm''
        'movemon monnum, roll(3) - 2 + plr.x, roll(3) - 2 + plr.y
        Form1.updatbody
    End If
    If aroll >= mroll And diff(mroll, aroll) <= 3 Then plr.xoff = roll(12): plr.yoff = roll(12): gamemsg "The " & mont.name & " is thrashing around inside your " & getbelly("") & ".  It is difficult to hold it inside."

End If
If plr.foodinbelly = 0 Then Form1.updatbody: gamemsg "Your " & getbelly("") & " has diligently " & getabsorb & "ed everything you have put in it."

End Function

Function ifnpcsaw() As Boolean
Dim a As Long, b As Long, c As Long  'm'' declare
On Error Resume Next
Dim tobj As objecttype
For a = plr.X - 6 To plr.X + 6
    If a < 1 Or a > mapx Then GoTo 6
    For b = plr.Y - 6 To plr.Y + 6
        If b > mapy Or b < 1 Then GoTo 5
        If map(a, b).object > 0 Then
            objn = map(a, b).object
            If Not geteff(objtypes(objs(objn).type), "NPC", 1) = "" Or Not geteff(objtypes(objs(objn).type), "Conversation", 1) = "" Then
            tobj = objtypes(objs(objn).type)
            If createmonster("City Guard", a + roll(3), b + roll(3)) = 0 Then createmontype "City Guard", "merc1.bmp", lowestlevel + 2, tobj.r, tobj.g, tobj.b, tobj.l
            For c = 1 To roll(3)
            createmonster "City Guard", a + roll(3), b + roll(3)
            Next c
            gamemsg "One of the townspeople saw you swallow her and summoned the guards!"
            End If
        End If
5    Next b
6 Next a

End Function

Function createmontype(name, gfile, level, r, g, b, light)
            lastmontype = lastmontype + 1
            ReDim Preserve montype(1 To lastmontype): ReDim Preserve mongraphs(1 To UBound(montype())) As cSpriteBitmaps
            montype(lastmontype).name = name
            montype(lastmontype).gfile = gfile
            montype(lastmontype).level = level
'            Input #1, montype(lastmontype).name, montype(lastmontype).gfile, montype(lastmontype).level
'            Input #1, r, g, b, light
            montype(lastmontype).color = RGB(r, g, b)
            montype(lastmontype).light = light
            calcexp montype(lastmontype)

            makesprite mongraphs(lastmontype), Form1.Picture1, montype(lastmontype).gfile, r, g, b, light, 3
createmontype = lastmontype

End Function

Function randsound(name As String, num As Byte)

playsound name & roll(num) & ".wav"

End Function

Function sellspef(obj As objecttype)
'Sells according to what an object has listed as selling...
buying = 1
Form10.clearchoices
Form10.disptext "What can I interest you in?"
For a = 1 To 6
    If obj.effect(a, 1) = "SELL" Then
    Form10.addchoice "##BUY2:" & obj.effect(a, 2) & ":" & obj.effect(a, 3), "I would like to buy a " & obj.effect(a, 2) & " (" & obj.effect(a, 3) & " gold)"
    End If
'    Form10.Label2(a).Visible = True
Next a
Form10.addchoice "##EXIT##", "(Leave)"
Form10.Show
'Form10.Label2(a).Visible = True
'Form10.Label2(a).caption = "Fuck."
End Function

Function ifsaid(ByVal str As String) As Boolean  'm'' declare
Dim a As Long 'm'' declare
hassaid = False
For a = 0 To 50
    If plr.alreadysaid(a) = "" Then Exit Function
    If plr.alreadysaid(a) = str Then ifsaid = True: Exit Function
Next a

End Function

Function addsaid(ByVal str As String) 'm'' declare
Dim a As Long 'm'' declare
For a = 0 To 50
    If plr.alreadysaid(a) = str Then Exit Function
    If plr.alreadysaid(a) = "" Then plr.alreadysaid(a) = str: Exit Function
Next a

End Function

Function ifmonstereaten(monname) As Boolean
Dim a As Long 'm'' declare
ifmonstereaten = True
For a = 0 To 500
    If plr.diggedmons(a) = monname Then Exit Function
    If plr.diggedmons(a) = "" Then plr.diggedmons(a) = monname: plr.skillpoints = plr.skillpoints + 1: ifmonstereaten = False: Exit Function
Next a

End Function

Function changeeff(ByVal objtnum As Integer, eff1, Optional ByVal eff2, Optional ByVal eff3, Optional ByVal eff4, Optional ByVal eff5)

For a = 1 To 6
    If objtypes(objtnum).effect(a, 1) = eff1 Then
        If Not IsMissing(eff2) Then objtypes(objtnum).effect(a, 2) = eff2
        If Not IsMissing(eff3) Then objtypes(objtnum).effect(a, 3) = eff3
        If Not IsMissing(eff4) Then objtypes(objtnum).effect(a, 4) = eff4
        If Not IsMissing(eff5) Then objtypes(objtnum).effect(a, 5) = eff5

        Exit Function
    End If
Next a

End Function

Function changeeff2(objt As objecttype, eff1, Optional ByVal eff2, Optional ByVal eff3, Optional ByVal eff4, Optional ByVal eff5)

For a = 1 To 6
    If objt.effect(a, 1) = eff1 Then
        If Not IsMissing(eff2) Then objt.effect(a, 2) = eff2
        If Not IsMissing(eff3) Then objt.effect(a, 3) = eff3
        If Not IsMissing(eff4) Then objt.effect(a, 4) = eff4
        If Not IsMissing(eff5) Then objt.effect(a, 5) = eff5

        Exit Function
    End If
Next a

If a = 7 Then addeffect2 objt, eff1, eff2, eff3, eff4, eff5

End Function

Sub loadshooties()

'INCREASE THIS IF YOU MAKE MORE THAN 5
For a = 1 To 8
Set shootygraphs(a) = New cSpriteBitmaps
Next a

shootygraphs(1).CreateFromFile "arrows.bmp", 18, 1, , 0
shootygraphs(2).CreateFromFile "fireballs.bmp", 18, 1, , RGB(255, 0, 255)
shootygraphs(3).CreateFromFile "fireball2s.bmp", 18, 1, , RGB(255, 0, 255)
shootygraphs(4).CreateFromFile "lightnings.bmp", 18, 1, , RGB(255, 0, 255)
shootygraphs(5).CreateFromFile "lightning2s.bmp", 18, 1, , RGB(255, 0, 255)
shootygraphs(6).CreateFromFile "goops.bmp", 18, 1, , 0
shootygraphs(7).CreateFromFile "bolts.bmp", 18, 1, , 0

End Sub

Function createshooty(xplus, yplus, stype, shootyis As String, Optional startx = 400, Optional starty = 300) As Long 'm'' declare
'Shootyis = Owner:dice:damage:pierce:radius:bijdang
Dim a As Long 'm'' declare
4
    For a = 1 To 100
        If shooties(a).Active = 0 Then
            shooties(a).frame = Int(((yplus + 1) / 2) * 9) '+ 1 '9 / ((yplus + 1) / 2 + 1)
            If xplus < 0 Then shooties(a).frame = 18 - shooties(a).frame '+ 1 '+ 9
            If shooties(a).frame <= 0 Then shooties(a).frame = 1
            shooties(a).graphnum = Val(stype)
            shooties(a).is = shootyis
            shooties(a).X = startx
            shooties(a).Y = starty
            shooties(a).Active = 1
            shooties(a).Time = 60
            shooties(a).xplus = xplus * 10
            shooties(a).yplus = yplus * 10
            shooties(a).xhit = 0
            shooties(a).yhit = 0
            createshooty = a
            Exit Function
        End If
    Next a
If a > 100 Then shooties(1).Active = False: GoTo 4


End Function

Function shootat(X, Y, stype, shootyis As String, Optional ByVal shootfromx = 0, Optional ByVal shootfromy = 0, Optional ByVal multiple = 1, Optional ByVal multangle = 10, Optional ByVal pierce = 0, Optional ByVal tiletargeting = 0)
'Shootyis = Owner:dice:damage:pierce:radius:bijdang
'If Not getfromstring(shootyis, 1) = "ENEMY" Then x = x - 400: y = y - 300 Else x = x - shootfromx: y = y - shootfromy
'x = x - 399: y = y - 300

If tiletargeting = 1 Then getlitxy2 X, Y, X, Y
If tiletargeting = 1 Then getlitxy2 shootfromx, shootfromy, shootfromx, shootfromy

Dim zod As Double
If Not pierce = 0 Then pierce = ":" & pierce Else pierce = ""
'If x = Empty Then x = 0
'If y = Empty Then y = 0
zim = greater(posit(X), posit(Y))
'zurg = lesser(posit(x), posit(y))
If zim = 0 Then zim = 1
zod = 1# / zim
'zod = zurg / zim



speed = 12
If stype = 1 Then playsound "bow.wav": speed = 18
If stype = 5 Or stype = 4 Then playsound "lightning2.wav": speed = 20
If stype = 3 Or stype = 2 Then playsound "burn.wav" ': speed = 10
If stype = 6 Then playsound "burpatk.wav"
If stype = 7 Then playsound "spellzap2.wav": speed = 16

If shootfromx = 0 Then shootfromx = 400: shootfromy = 300
If Not Val(pierce) = 0 Then replaceinstr shootyis, 4, pierce
a = createshooty(zod * X, zod * Y, stype, shootyis, shootfromx, shootfromy)


'If Not IsMissing(shootfromx) Then
calcangle shootfromx, shootfromy, X, Y, shooties(a).xplus, shooties(a).yplus, speed, shooties(a).frame, 18

'If IsMissing(shootfromx) Then createshooty zod * x, zod * y, stype, shootyis 'Else
If multiple > 1 Then
    If multangle = 0 Then multangle = 10
    If multiple > 36 Then multiple = 36
    For c = 2 To multiple
        a = createshooty(zod * X, zod * Y, stype, shootyis, shootfromx, shootfromy)
        If c Mod 2 = 0 Then
        calcangle shootfromx, shootfromy, X, Y, shooties(a).xplus, shooties(a).yplus, speed, shooties(a).frame, 18, multangle * Int(c / 2)
        Else:
        calcangle shootfromx, shootfromy, X, Y, shooties(a).xplus, shooties(a).yplus, speed, shooties(a).frame, 18, -multangle * Int(c / 2)
        End If
    Next c

End If


End Function

Function drawshooties()
'Shootyis = Owner:dice:damage:pierce:radius:bijdang
Dim a As Long 'm'' declare

For a = 1 To 100
    If shooties(a).Active = 1 Then
        shooties(a).X = shooties(a).X + shooties(a).xplus
        shooties(a).Y = shooties(a).Y + shooties(a).yplus
        If shooties(a).graphnum >= 1 Then
        If shooties(a).graphnum = 1 Or shooties(a).graphnum = 6 Then
        shootygraphs(shooties(a).graphnum).TransparentDraw picBuffer, shooties(a).X - shootygraphs(shooties(a).graphnum).CellWidth / 2, shooties(a).Y - 30 + offset - shootygraphs(shooties(a).graphnum).CellHeight / 2, shooties(a).frame, True
        Else:
        shootygraphs(shooties(a).graphnum).TransparentDraw picBuffer, shooties(a).X - shootygraphs(shooties(a).graphnum).CellWidth / 2, shooties(a).Y - 30 + offset - shootygraphs(shooties(a).graphnum).CellHeight / 2, shooties(a).frame
        End If
        End If
        
        shooties(a).Time = shooties(a).Time - 1
        If shooties(a).Time <= 0 Then shooties(a).Active = False
        
        x2 = shooties(a).X: y2 = shooties(a).Y
        getXY x2, y2
        If x2 > mapx Or x2 < 1 Or y2 > mapy Or y2 < 1 Then shooties(a).Active = False: GoTo 6
        If map(x2, y2).blocked > 0 And map(x2, y2).ovrtile > 0 Then shooties(a).Active = False: GoTo 6
        If getfromstring(shooties(a).is, 1) = "ENEMY" Then
            If plr.X = x2 And plr.Y = y2 And plr.instomach = 0 Then
            If HasSkill("Reflection") And spendsp(3) Then
                gamemsg "You reflect the projectile!"
                shooties(a).xplus = -shooties(a).xplus
                shooties(a).yplus = -shooties(a).yplus
                shooties(a).frame = shooties(a).frame + 9
                If shooties(a).frame > 18 Then shooties(a).frame = shooties(a).frame - 18
                replaceinstr shooties(a).is, 2, getfromstring(shooties(a).is, 2) * getplrskill("Reflection")
                replaceinstr shooties(a).is, 1, "PLAYER"
                'shooties(a).is = getfromstring(shooties(a).is, 2) * getplrskill("Reflection") & ":" & getfromstring(shooties(a).is, 3)
                GoTo 6
            End If
            plrdamage rolldice(getfromstring(shooties(a).is, 2), getfromstring(shooties(a).is, 3)): shooties(a).Active = False
            If shooties(a).graphnum = 6 Then If roll(8) = 1 Then digestclothes rolldice(getfromstring(shooties(a).is, 2), getfromstring(shooties(a).is, 3)) / 2
            End If
        Else:
            'Shootyis = Owner:dice:damage:pierce:radius:bijdang
            If map(x2, y2).monster > 0 And Not shooties(a).xhit = x2 And Not shooties(a).yhit = y2 Then
                monn = map(x2, y2).monster
                If Not mon(monn).owner > 0 And Not plr.instomach = monn Then
                If shooties(a).graphnum = 1 Then playsound "thunk.wav"
                If Val(getfromstring(shooties(a).is, 6)) > 0 Then makebijdang x2, y2, getfromstring(shooties(a).is, 6)
                If Val(getfromstring(shooties(a).is, 5)) > 0 Then
                    radiusdamage getfromstring(shooties(a).is, 5), x2, y2, rolldice(getfromstring(shooties(a).is, 2), getfromstring(shooties(a).is, 3))
                Else:
                    damagemon monn, rolldice(getfromstring(shooties(a).is, 2), getfromstring(shooties(a).is, 3))
                End If
                If mon(monn).hp > 0 Then dispatk montype(mon(monn).type), mon(monn).hp
                
                'Multiple shots
                nurg = Val(getfromstring(shooties(a).is, 4))
                
                'Negative 'pierce' means it bounces like chain lightning
                If nurg < 0 Then
                    nurg = nurg + 1
                    replaceinstr shooties(a).is, 4, nurg
                    shooties(a).xhit = x2: shooties(a).yhit = y2
                    targ = closemon(x2, y2, 8)
                    'If roll(8) = 1 Or targ = 0 Then x3 = plr.x: y3 = plr.y Else
                    If targ > 0 Then x3 = mon(targ).X: y3 = mon(targ).Y
                    If x3 = 0 Or y3 = 0 Or targ = 0 Then shooties(a).Active = False
                    calcangle x2, y2, x3, y3, shooties(a).xplus, shooties(a).yplus, 14, shooties(a).frame, 18, , 1
                End If
                
                'Positive Pierce
                If nurg >= 0 Then
                    nurg = nurg - 1
                    If nurg < 0 Then shooties(a).Active = False Else shooties(a).xhit = x2: shooties(a).yhit = y2: replaceinstr shooties(a).is, 4, nurg
                
                End If
                End If
            End If
        End If
6
        
    End If
Next a

'raindensity = 1
If raindensity > 0 Then drawrain

End Function

Function outofbounds(ByVal X, ByVal Y)
outofbounds = False
If X > mapx Or X < 1 Then outofbounds = True
If Y > mapy Or Y < 1 Then outofbounds = True
End Function

Function radiusdamage(Radius, X, Y, damage, Optional bijdangnum = 0)

For a = -Radius To Radius
    For b = -Radius To Radius
    If outofbounds(X + a, Y + b) Then GoTo 5
    If map(X + a, Y + b).monster > 0 Then
        divis = greater(diff(0, a), diff(0, b))
        damagemon map(X + a, Y + b).monster, damage
    End If
    
    
5     Next b
Next a

If bijdangnum > 0 Then makebijdang X, Y, bijdangnum

End Function

Function getXY(ByRef X, ByRef Y)
'converts pixels to tile positions

Y = Y + -offset

orgX = X
X = Int((X - 12) / 96 + (Y - 12) / 48) + plr.X - 10
Y = Int(Y / 48 - orgX / 96) + plr.Y - 2

End Function

Function revgetXY2(ByRef X, ByRef Y)
'converts tile positions to pixels
'I'm not sure if this function works yet

'y = y + -offset

dorkx = (X - plr.X - Y + plr.Y) * 48 + 400 - plr.xoff - (50) + xoff
'If dorkx < -40 Or dorkx > 750 Then Exit Sub
dorky = (X + Y - plr.X - plr.Y) * 24 - plr.yoff + 300 - (50 - 48) - (midtile * 24) + yoff

X = dorkx
Y = dorky + offset

'orgX = x
'x = Int((x - 12) / 96 + (y - 12) / 48) + plr.x - 10
'y = Int(y / 48 - orgX / 96) + plr.y - 2

End Function

Function revgetXY(ByRef X, ByRef Y)
'converts tile positions to pixels
'I'm not sure if this function works yet

'y = y + -offset

dorkx = (X * 48) - (Y * 24)

'dorkx = (x - y) * 48 + 400 - plr.xoff - 50
'dorkx = (x - plr.x - y + plr.y) * 48 + 400 - plr.xoff - (50) + xoff
'If dorkx < -40 Or dorkx > 750 Then Exit Sub

dorky = (X + Y) * 24

'dorky = (x + y) * 24 - plr.yoff + 300 - 2 - midtile * 24
'dorky = (x + y - plr.x - plr.y) * 24 - plr.yoff + 300 - (50 - 48) - (midtile * 24) + yoff

X = dorkx
Y = dorky + offset

'orgX = x
'x = Int((x - 12) / 96 + (y - 12) / 48) + plr.x - 10
'y = Int(y / 48 - orgX / 96) + plr.y - 2

End Function

Function calcdist(x1, y1, x2, y2)

xlen = diff(x1, x2)
ylen = diff(y1, y2)
calcdist = Sqr((xlen * xlen) + (ylen * ylen))

End Function

Function calcangle(ByVal fromx, ByVal fromy, ByVal tox, ByVal toy, ByRef movex, ByRef movey, ByVal speed, ByRef cell, ByVal totalcells, Optional ByVal addang = 0, Optional ByVal tiletargeting = 0)

If tiletargeting = 1 Then getlitxy2 fromx, fromy, fromx, fromy: getlitxy2 tox, toy, tox, toy

litx = Abs(fromx) - Abs(tox)
lity = Abs(fromy) - Abs(toy)
If lity = 0 Then lity = 1
'litx = diff(fromx, tox)
'lity = diff(fromy, toy)

hyp = Sqr((litx ^ 2) + (lity ^ 2))

sine = lity / hyp
litang = Atn(litx / lity)
ang = Atn(litx / lity) * (180 / pi) + addang

'ang = sine * (180 / pi)

    If Sgn(lity) = 1 Then ang = ang + 90
    If Sgn(lity) = 0 And Sgn(temp_x#) = 1 Then ang = 180
    If Sgn(lity) = 0 And Sgn(temp_x#) = -1 Then ang = 0
    If Sgn(lity) = -1 Then ang = ang + 270

'Debug.Print "Angle:" & ang
3 If ang > 360 Then ang = ang - 360: GoTo 3
If ang < 0 Then ang = ang + 360: GoTo 3

If fromx < 0 Then fromx = fromx * -2 ' + 400
If fromy < 0 Then fromy = fromy * -2 '+ 300
movex = ((tox - fromx) / hyp) * speed
movey = ((toy - fromy) / hyp) * speed

ang2 = ang + 90
5 If ang2 > 360 Then ang2 = ang2 - 360: GoTo 5
If ang2 < 0 Then ang2 = ang2 + 360: GoTo 5
movex = Sin(ang2 * pi / 180) * speed
movey = Cos(ang2 * pi / 180) * speed

'movey = Sin(litang) * speed

cell = (ang / 360) * totalcells
If cell < 1 Or cell > totalcells Then cell = 1

End Function

Function calcaccel(ByVal angle, ByVal accel, ByVal topspeed, ByRef movex, ByRef movey)

'litx = Abs(fromx) - Abs(tox)
'lity = Abs(fromy) - Abs(toy)
'If lity = 0 Then lity = 1
'litx = diff(fromx, tox)
'lity = diff(fromy, toy)

'hyp = Sqr((litx ^ 2) + (lity ^ 2))

'sine = lity / hyp
'litang = Atn(litx / lity)
'ang = Atn(litx / lity) * (180 / pi) + addang

ang = angle
'ang = sine * (180 / pi)

'    If Sgn(lity) = 1 Then ang = ang + 90
'    If Sgn(lity) = 0 And Sgn(temp_x#) = 1 Then ang = 180
'    If Sgn(lity) = 0 And Sgn(temp_x#) = -1 Then ang = 0
'    If Sgn(lity) = -1 Then ang = ang + 270

'Debug.Print "Angle:" & ang
3 If ang > 360 Then ang = ang - 360: GoTo 3
If ang < 0 Then ang = ang + 360: GoTo 3

'If fromx < 0 Then fromx = fromx * -2 ' + 400
'If fromy < 0 Then fromy = fromy * -2 '+ 300
'movex = ((tox - fromx) / hyp) * speed
'movey = ((toy - fromy) / hyp) * speed

ang2 = ang + 90
5 If ang2 > 360 Then ang2 = ang2 - 360: GoTo 5
If ang2 < 0 Then ang2 = ang2 + 360: GoTo 5
xacc = Sin(ang2 * pi / 180) * accel
movex = movex + xacc
topx = Sin(ang2 * pi / 180) * topspeed: If topx < 0 Then topx = -topx
If diff(movex, 0) > topx Then reduce movex, accel

yacc = Cos(ang2 * pi / 180) * accel
movey = movey + yacc
topy = Cos(ang2 * pi / 180) * topspeed: If topy < 0 Then topy = -topy
If diff(movey, 0) > topy Then reduce movey, accel
'movex = topx: movey = topy
'movey = Sin(litang) * speed

'cell = (ang / 360) * totalcells
'If cell < 1 Or cell > totalcells Then cell = 1

End Function

Function sgndiff(ByVal val1, ByVal val2)
If val1 < 0 And val2 > 0 Then val2 = -val2
If val1 > 0 And val2 < 0 Then val2 = -val2
sgndiff = diff(val1, val2)
End Function

Function reduce(val1, amt)
If amt < 0 Then amt = -amt
If val1 > 0 Then If amt > val1 Then val1 = 0 Else val1 = val1 - amt Else If -amt < val1 Then val1 = 0 Else val1 = val1 + amt

End Function

Function showform10(Optional buything As String, Optional loadconv = 0, Optional picfile As String = "", Optional startpos As String = "MAIN", Optional convfile As String = "Conversations.txt", Optional worth = 0)
Form10.Cls
If worth = 0 Then worth = mapjunk.level
If loadconv = 0 Then Form10.buyclothes buything, (worth)
If loadconv = 1 Then Form10.loadconv buything, picfile, startpos, convfile
If loadconv = -1 Then Form10.disptext buything, 1, picfile
'Form1.Timer1.Enabled = False
Form10.Show 1
'Form1.Timer1.Enabled = True
End Function

Function ClearSprites2()

Erase objgraphs

Erase mongraphs

Erase auras()
'objgraphs()
Set wepgraph = Nothing
'mongraphs()
Erase bijdang()
Erase digbody()
Set plrgraphs = Nothing
Erase cgraphs()
Set cleavage = Nothing
Set plrhair = Nothing
Set raingraph = Nothing
Set tilespr = Nothing
Set tilespr2 = Nothing
Set transoverspr = Nothing
Set ovrspr = Nothing
Set capegraphs = Nothing
Set monstermaps = Nothing
Set gbody = Nothing
Set waterspr = Nothing
Set fullbody = Nothing
Erase shootygraphs()
Erase extraovrs()
Erase transextraovrs()
Set bars.life = Nothing
Set bars.life2 = Nothing
Set bars.mana = Nothing
Set bars.fatigue = Nothing
Set bars.empty = Nothing
'bars.life/life2/mana/fatigue/empty
Set minimapspr = Nothing
Set minimapspr2 = Nothing
Set mouthgraph = Nothing
Set mapbufspr = Nothing
Erase spritemaps() '.cmap
Set sprt = Nothing
Set sprt2 = Nothing
Set backgr = Nothing

End Function

Function clearsprites()

Erase objgraphs

Erase mongraphs


'Unload Form1.DMC1
'Set ovrspr = Nothing
'Set tilespr = Nothing
'Set transovrspr = Nothing
End Function

Function combatskill(skillname, amt, mpcost)
combatskill = False
If mpcost > plr.mp Then gamemsg "Not enough mana": Exit Function

'If skillname = "Piercing Arrow" Then
    
'End If

End Function

Function spendsp(amt) As Boolean

Static lastmsgamt

If plr.sp > 20 Then lastmsgamt = 0

spendsp = False
If amt > plr.sp Then
    If lastmsgamt = 0 Then
        gamemsg "Not enough skill points"
        lastmsgamt = 1
        randsound "skill", 2
    End If
    usingskill = ""
    Form1.Picture6.Visible = False
    Exit Function
End If
plr.sp = plr.sp - amt: spendsp = True

End Function

Function remspaces(ByVal str) As String
remspaces = swaptxt(str, " ", "")
End Function

Function iscombatskill(skillname)
iscombatskill = 0
For a = 1 To 4
    If plr.combatskills(a) = skillname Then iscombatskill = a: Exit Function
Next a

End Function

Function baseequip()

Select Case plr.Class
    Case "Amazon"
        'plr.hpmax = 60: plr.str = 5: plr.dex = 3: plr.int = 2: plr.mpmax = 0
        'redd = 130: green = 80: blue = 0: lit = 0.4
        redd = 230: green = 80: blue = 0: lit = 0.4
            'addclothes "Halter", "halter1.bmp", 4, "Bra", "Upper", redd, green, blue, lit, 1
            'addclothes "Loincloth", "loincloth1.bmp", 3, "Lower", , redd, green, blue, lit, 1
            'addskill "Sword Mastery"
            'addskill "Axe Mastery"
            addclothes "Chain Bodysuit", "chainswimsuit.bmp", 6, "Upper", "Lower", 150, 150, 150, 0.6, 1
            addclothes "Crimson Sash", "sash1.bmp", 2, "Jacket", , 230, 0, 0, 0.4, 1
'    plr.combatskills(1) = "Power Strike"
'    plr.combatskills(2) = "Charged Strike"
'    plr.combatskills(3) = "Frenzy"
'    plr.combatskills(4) = "Cripple"
    
    
    Case "Sorceress"
        'plr.hpmax = 30: plr.str = 2: plr.dex = 3: plr.int = 5: plr.mpmax = 40
        redd = 40: green = 40: blue = 40: lit = 0.3
            addclothes "Fine Dress", "dress2.bmp", 4, "Upper", "Lower", redd, green, blue, lit, 1
            addclothes "Lace Bra", "bra5.bmp", 2, "Bra", , redd, green, blue, lit, 1
            addclothes "Lace Panties", "panties2.bmp", 2, "Panties", , redd, green, blue, lit, 1
            'givespefspell "Firebolt"
  '  plr.combatskills(1) = "Alchemy"
  '  plr.combatskills(2) = "Damage Energy"
  '  plr.combatskills(3) = "Mana Shield"
  '  plr.combatskills(4) = "Split Spell"
    
    Case "Priestess"
        'plr.hpmax = 45: plr.str = 4: plr.dex = 2: plr.int = 3: plr.mpmax = 30
        redd = 80: green = 120: blue = 10: lit = 0.4
            addclothes "Shortcape", "cape1.bmp", 2, "Jacket", , redd, green, blue, lit, 1
        redd = 250: green = 250: blue = 250: lit = 0.4
            addclothes "Gloves", "gloves1.bmp", 2, "Arms", , redd, green, blue, lit, 1
            addclothes "Lace Panties", "panties2.bmp", 1, "Panties", , redd, green, blue, lit, 1
            addclothes "Lace Bra", "bra5.bmp", 1, "Bra", , redd, green, blue, lit, 1
            'addskill "Axe Mastery"
   ' plr.combatskills(1) = "Stun"
   ' plr.combatskills(2) = "Block"
   ' plr.combatskills(3) = "Power Strike"
   ' plr.combatskills(4) = "Charged Strike"
                
                
    Case "Enchantress"
        'plr.hpmax = 35: plr.str = 2: plr.dex = 3: plr.int = 4: plr.mpmax = 35
        redd = 40: green = 40: blue = 120: lit = 0.4
            addclothes "Robe Skirt", "robebottom1.bmp", 3, "Lower", , redd, green, blue, lit, 1
            addclothes "Panties", "panties2.bmp", 1, "Panties", , 250, 250, 250, 0.5, 1
            addclothes "Corset Bra", "bra3.bmp", 2, "Bra", , redd, green, blue, lit, 1
            'givespefspell "Speed"
  '  plr.combatskills(1) = "Alchemy"
  '  plr.combatskills(2) = "Damage Energy"
  '  plr.combatskills(3) = "Mana Shield"
  '  plr.combatskills(4) = "Cunning Strike"
        
    Case "Huntress"
        'plr.hpmax = 40: plr.str = 3: plr.dex = 4: plr.int = 3: plr.mpmax = 25
        redd = 110: green = 60: blue = 0: lit = 0.3
            addclothes "Shirt", "doublet1.bmp", 2, "Upper", , redd, green, blue, lit, 1
            addclothes "Armored Skirt", "armorskirt1.bmp", 2, "Lower", , redd, green, blue, lit, 1
            addclothes "Bra", "bra1.bmp", 1, "Bra", , redd, green, blue, lit, 1
            addclothes "Panties", "panties1.bmp", 1, "Panties", , redd, green, blue, lit, 1
        'addskill "Bow Mastery"
        'addskill "Spear Mastery"

    Case "Valkyrie"
        'plr.hpmax = 35: plr.str = 3: plr.dex = 3: plr.int = 3: plr.mpmax = 35
        redd = 50: green = 120: blue = 220: lit = 0.5
            addclothes "Plate Mail", "breastplate5.bmp", 8, "Upper", , redd, green, blue, lit, 1
            addclothes "Bra", "bra1.bmp", 1, "Bra", , redd, green, blue, lit, 1
            addclothes "Panties", "panties1.bmp", 1, "Panties", , redd, green, blue, lit, 1
            addclothes "Armplates", "armplates1.bmp", 2, "Arms", , redd, green, blue, lit, 1
        'addskill "Spear Mastery"

    Case "Angel"
        'plr.hpmax = 50: plr.str = 2: plr.dex = 5: plr.int = 4: plr.mpmax = 35
        redd = 200: green = 200: blue = 200: lit = 0.2
            addclothes "Robe", "robe1.bmp", 4, "Lower", "Upper", redd, green, blue, lit, 1

    Case "Succubus"
        'plr.hpmax = 75: plr.str = 5: plr.dex = 3: plr.int = 4: plr.mpmax = 15
        redd = 20: green = 20: blue = 20: lit = 0.4
            'addclothes "Lace Corset", "corset1.bmp", 3, "Upper", "Bra", redd, green, blue, lit, 1
            'addclothes "Lace Panties", "panties2.bmp", 2, "Panties", , redd, green, blue, lit, 1
            addclothes "Lace Teddy", "teddy1.bmp", 3, "Panties", "Bra", redd, green, blue, lit, 1
            addclothes "Fishnet Stockings", "fishnets1.bmp", 2, "Legs", , redd, green, blue, lit, 1
     
    Case "TombRaider"
        'plr.hpmax = 60: plr.str = 3: plr.dex = 4: plr.int = 4: plr.mpmax = 0
        'redd = 130: green = 80: blue = 0: lit = 0.4
        redd = 230: green = 230: blue = 230: lit = 0.6
            addclothes "Bra", "bra2.bmp", 1, "Bra", , redd, green, blue, lit, 1
            addclothes "Panties", "panties1.bmp", 1, "Panties", , redd, green, blue, lit, 1
            addclothes "Shirt", "shirt1.bmp", 4, "Upper", , 70, 200, 210, 0.5, 1
            addclothes "Shorts", "shorts1.bmp", 3, "Lower", , 120, 50, 8, 0.6, 1
            addclothes "Belt", "belt.bmp", 2, "Belt", , , , , , 1
    
    Case "Streetfighter"
        'plr.hpmax = 60: plr.str = 5: plr.dex = 4: plr.int = 3: plr.mpmax = 0
        redd = 250: green = 250: blue = 250: lit = 0.5
            addclothes "Combat Dress", "chunli1.bmp", 4, "Upper", "Lower", 0, 140, 255, lit, 1
            addclothes "Lace Bra", "bra5.bmp", 2, "Bra", , redd, green, blue, lit, 1
            addclothes "Lace Panties", "panties2.bmp", 2, "Panties", , redd, green, blue, lit, 1
    
    
    Case "Caller"
        'plr.hpmax = 20: plr.str = 2: plr.dex = 3: plr.int = 4: plr.mpmax = 45
        redd = 20: green = 150: blue = 0: lit = 0.5
            addclothes "Bra", "bra2.bmp", 1, "Bra", , redd, green, blue, lit, 1
            addclothes "Panties", "panties1.bmp", 1, "Panties", , redd, green, blue, lit, 1
            redd = 20: green = 150: blue = 0: lit = 0.5
            addclothes "Shirt", "doublet1.bmp", 4, "Upper", , redd, green, blue, 0.5, 1
            addclothes "Robe Skirt", "robebottom1.bmp", 3, "Lower", , redd, green, blue, lit, 1
            'givespefspell "Firebolt"
            'givespefspell "Summon Faerie"
            
End Select

Form1.updatbody

End Function

Function getmonstats()

Dim mn(1 To 6) As monstertype

mn(1).level = 1
mn(2).level = 5
mn(3).level = 10
mn(4).level = 20
mn(5).level = 30
mn(6).level = 40

MonStats.prnt "GIANTESS" & vbCrLf
For a = 1 To 6: mn(a).gfile = "giantess.bmp": setmon mn(a), 2, 1, 4, 6, 6, 8, 4, 1: monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "KARI" & vbCrLf
For a = 1 To 6: mn(a).gfile = "kari1.bmp": setmon mn(a), 1.5, 1.2, 3, 4, 4, 6, 4, 1: mn(a).weaktype = "Spear": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "GIANT SNAKE" & vbCrLf
For a = 1 To 6: mn(a).gfile = "snake1.bmp": setmon mn(a), 1, 1, 2, 6, 5, 6, 4, 1.2: mn(a).weaktype = "Sword": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "NAGA" & vbCrLf
For a = 1 To 6: mn(a).gfile = "snakewoman1.bmp": setmon mn(a), 1.3, 1.4, 3, 8, 4, 5, 4, 1.5: mn(a).Sound = "Monwierd.wav": mn(a).weaktype = "Axe": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "ROYAL NAGA" & vbCrLf
For a = 1 To 6: mn(a).gfile = "snakewoman2.bmp": setmon mn(a), 1.8, 1.5, 2, 10, 3, 6, 4, 1: mn(a).Sound = "Monwierd.wav": mn(a).weaktype = "Axe": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "GIANT WORM" & vbCrLf
For a = 1 To 6: mn(a).gfile = "worm1.bmp": setmon mn(a), 1.2, 1, 2, 4, 6, 4, 4, 2: mn(a).Sound = "Monbeast.wav": mn(a).weaktype = "Axe": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "ACID VINES" & vbCrLf
For a = 1 To 6: mn(a).gfile = "tendrils1.bmp": setmon mn(a), 0.5, 1, 1, 4, 7, 3, 2, 0.8: mn(a).Sound = "Monfrog.wav": mn(a).weaktype = "Sword": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "GIRLCRUNCH PLANT" & vbCrLf
For a = 1 To 6: mn(a).gfile = "venus1.bmp": setmon mn(a), 2, 0.8, 2, 6, 8, 9, 1, 2: mn(a).Sound = "Monfrog.wav": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "GIANT FROG" & vbCrLf
For a = 1 To 6: mn(a).gfile = "frog1.bmp": setmon mn(a), 0.8, 1, 1, 4, 7, 4, 4, 0.6: mn(a).Sound = "Monfrog.wav": mn(a).weaktype = "Bow": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "SLIME" & vbCrLf
For a = 1 To 6: mn(a).gfile = "slime1.bmp": setmon mn(a), 2, 1, 4, 8, 7, 3, 4, 2: mn(a).Sound = "Monslime.wav": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "CENTAURESS" & vbCrLf
For a = 1 To 6: mn(a).gfile = "centauress1.bmp": setmon mn(a), 1.4, 1, 3, 6, 3, 8, 4, 1: mn(a).weaktype = "Spear": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "HARPY" & vbCrLf
For a = 1 To 6: mn(a).gfile = "harpy1.bmp": setmon mn(a), 1.1, 2, 4, 6, 3, 4, 6, 1.5: mn(a).Sound = "Monshriek.wav": mn(a).weaktype = "Bow": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "SPRITE" & vbCrLf
For a = 1 To 6: mn(a).gfile = "sprite1.bmp": setmon mn(a), 0.6, 2, 1, 8, 4, 5, 7, 1.5: mn(a).Sound = "Magic2.wav": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "SUCCUBUS" & vbCrLf
For a = 1 To 6: mn(a).gfile = "succubus1.bmp": setmon mn(a), 2, 1.5, 3, 12, 3, 6, 4, 1.3: mn(a).weaktype = "Spear": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "DEMONESS" & vbCrLf
For a = 1 To 6: mn(a).gfile = "demoness1.bmp": setmon mn(a), 2.5, 1.3, 4, 8, 2, 6, 5, 3: monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "DEMONETTE" & vbCrLf
For a = 1 To 6: mn(a).gfile = "demoness2.bmp": setmon mn(a), 1.5, 1.2, 3, 6, 3, 4, 5, 2, , 2, 6, 6, 2: monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "LIZARD WOMAN" & vbCrLf
For a = 1 To 6: mn(a).gfile = "lizardwoman1.bmp": setmon mn(a), 1.2, 2, 3, 6, 3, 3, 5, 1: monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "CARNIVOROUS FLOWER" & vbCrLf
For a = 1 To 6: mn(a).gfile = "flower1.bmp": setmon mn(a), 1.3, 2, 1, 2, 8, 7, 0, 1: mn(a).Sound = "Monwierd.wav": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

MonStats.prnt vbCrLf & "WITCH" & vbCrLf
For a = 1 To 6: mn(a).gfile = "mage1.bmp": setmon mn(a), 0.8, 0.8, 1, 6, 2, 5, 2, 1: mn(a).Sound = "Magic2.wav": monstats2 mn(a): Next a
If Not mn(1).weaktype = "" Then MonStats.prnt "Weak against " & mn(1).weaktype & "s": mn(1).weaktype = ""

End Function

Function monstats2(mn As monstertype)

MonStats.prnt "Level " & mn.level & "-- HP: " & mn.hp & ", " & " Attack Skill: " & mn.skill & ", Damage: " & mn.dice & "D" & mn.damage & ", Swallowing Skill: " & mn.eatskill & ", Stomach Acid: " & mn.acid

End Function

Function loadextraovrs(ByVal name1, ByVal name2, ByVal name3, ByVal name4)

mapjunk.name1 = name1
mapjunk.name2 = name2
mapjunk.name3 = name3
mapjunk.name4 = name4

If Not name1 = "" Then
Set extraovrs(1) = New cSpriteBitmaps
extraovrs(1).CreateFromFile name1, 1, 1, , 0
End If

'If Not Dir("trans" & name1) = "" Then
    If Not name1 = "" Then
    Set transextraovrs(1) = New cSpriteBitmaps
    transextraovrs(1).CreateFromFile "trans" & name1, 1, 1, , 0
    End If
'End If

'If Not Dir("trans" & name2) = "" Then
    If Not name2 = "" Then
    Set transextraovrs(2) = New cSpriteBitmaps
    transextraovrs(2).CreateFromFile "trans" & name2, 1, 1, , 0
    End If
'End If

'If Not Dir("trans" & name3) = "" Then
    If Not name3 = "" Then
    Set transextraovrs(3) = New cSpriteBitmaps
    transextraovrs(3).CreateFromFile "trans" & name3, 1, 1, , 0
    End If
'End If

'If Not Dir("trans" & name4) = "" Then
    If Not name4 = "" Then
    Set transextraovrs(4) = New cSpriteBitmaps
    transextraovrs(4).CreateFromFile "trans" & name4, 1, 1, , 0
    End If
'End If

If name2 = "" Then GoTo 15
Set extraovrs(2) = New cSpriteBitmaps
extraovrs(2).CreateFromFile name2, 1, 1, , 0
'transextraovrs(2).CreateFromFile name2, 1, 1, , 0

If name3 = "" Then GoTo 15
Set extraovrs(3) = New cSpriteBitmaps
extraovrs(3).CreateFromFile name3, 1, 1, , 0

If name4 = "" Then GoTo 15
Set extraovrs(4) = New cSpriteBitmaps
extraovrs(4).CreateFromFile name4, 1, 1, , 0

15
End Function

Function getbonus(bonusname)

zark = 0
zark = zark + geteff(wep.obj, bonusname, 2)

For a = 1 To 16
zark = zark + geteff(clothes(a).obj, bonusname, 2)
Next a

For a = 1 To UBound(auras())
If auras(a).loaded > 0 Then zark = zark + Val(geteff(auras(a).obj, bonusname, 2))
Next a

getbonus = zark

End Function

Function getbonusskill(skillname) As Integer
'Static zod
'If updatbonuses = 0 Then getbonusskill = zod: Exit Function
zod = 0

zark = geteff(wep.obj, "BONSKILL", 3)
If zark = skillname Then zod = zod + geteff(wep.obj, "BONSKILL", 2)

    For b = 1 To UBound(wep.obj.effect(), 1)
    If Not wep.obj.effect(b, 1) = "BONSKILL" Then GoTo 3
    If wep.obj.effect(b, 3) = skillname Then zod = zod + wep.obj.effect(b, 2)
3    Next b

For a = 1 To 16
    If clothes(a).loaded = 0 Then GoTo 12
    For b = 1 To UBound(clothes(a).obj.effect(), 1)
    If Not clothes(a).obj.effect(b, 1) = "BONSKILL" Then GoTo 5
    If clothes(a).obj.effect(b, 3) = skillname Then zod = zod + clothes(a).obj.effect(b, 2)
5    Next b
    
12 Next a

For a = 1 To UBound(auras())
    If auras(a).loaded = 0 Then GoTo 9
    For b = 1 To UBound(auras(a).obj.effect(), 1)
    If Not auras(a).obj.effect(b, 1) = "BONSKILL" Then GoTo 7
    If auras(a).obj.effect(b, 3) = skillname Then zod = zod + auras(a).obj.effect(b, 2)
7    Next b
    
9 Next a

getbonusskill = zod

End Function

Function getstr(Optional nofatigue = 0) As Integer

Static lastret
If updatbonuses = 1 Or lastret = 0 Then lastret = plr.str + getbonus("BONSTR")
If nofatigue = 0 Then getstr = greater(lastret * (getfatiguemult), 1)
If nofatigue = 1 Then getstr = greater(1, lastret)
'If nofatigue = 0 Then getstr = greater(1, (plr.str + getbonus("BONSTR")) * (getfatiguemult * 0.7))
'If Not nofatigue = 0 Then getstr = greater(1, plr.str + getbonus("BONSTR"))

End Function

Function getdex(Optional nofatigue = 0) As Integer

Static lastret
If updatbonuses = 1 Then lastret = plr.dex + getbonus("BONDEX") - dexloss
dexweight = dexloss 'Int(greater((getplrweight - plr.str) / 4, 0))
If nofatigue = 0 Then getdex = greater(((lastret - plr.foodinbelly) * getfatiguemult) - dexweight, 1)
If nofatigue = 1 Then getdex = greater(1, lastret - plr.foodinbelly - dexweight)

End Function

Function dexloss()
'Lose 1 dex for every 5 weight points by which you exceed your strength * 5

effstr = getstr * 5
dexloss = greater(0, Int((getplrweight - effstr) / 5))

End Function

Function getfatiguemult()
getfatiguemult = 1
If plr.fatigue = 0 Or plr.fatigue < plr.fatiguemax / 3 Then Exit Function
getfatiguemult = 1 - (plr.fatigue / plr.fatiguemax * 0.7) * 0.7
End Function

Function getint()
Static lastret
If updatbonuses = 1 Then lastret = greater(1, plr.int + getbonus("BONINT"))
getint = lastret
End Function

Function getend()
Static lastret
If updatbonuses = 1 Then lastret = greater(1, plr.endurance + getbonus("BONEND"))
getend = lastret
End Function

Function gethpmax()
Static lastret
If updatbonuses = 1 Then lastret = greater(1, plr.hpmax + getbonus("BONHP"))
gethpmax = lastret
End Function

Function getmpmax()
Static lastret
If updatbonuses = 1 Then lastret = plr.mpmax + getbonus("BONMP")
getmpmax = lastret
End Function

Function makegem(obj As objecttype, Optional ByVal worth As Byte = 0)

red = 150: green = 0: blue = 150: l = 1

addeffect2 obj, "Pickup", "gem1.wav"

If worth = 0 Then worth = averagelevel

size = 1
3 size = Int(Sqr(worth) - 1 + roll(3) - 1): If size > 5 Then size = 5 Else If size < 1 Then GoTo 3
'If size <= 4 Then If roll(50) <= worth Then size = size + 1: GoTo 3

If size = 1 Then sizestr = "Tiny "
If size = 2 Then sizestr = "Small "
If size = 3 Then sizestr = "Medium "
If size = 4 Then sizestr = "Large "
If size = 5 Then sizestr = "Priceless "

obj.graphname = "gem" & size & ".bmp"

If roll(6) = 1 Then uncut = "Uncut ": size = size / 2: GoTo 32 Else uncut = ""
If roll(6) = 1 Then em = "Emmaculate ": uncut = "" Else em = ""
32

aroll = roll(9)

Select Case aroll
    Case 1:
        gemname = "Diamond"
        addeffect2 obj, "GEM", gemname
        red = 250: green = 250: blue = 250
        If em = "" Then
        addeffect2 obj, "BONSTR", 1 * size
        addeffect2 obj, "BONDEX", 1 * size
        addeffect2 obj, "BONINT", 1 * size
        Else:
        addeffect2 obj, "BONSTR", 3 * size
        addeffect2 obj, "BONDEX", 3 * size
        addeffect2 obj, "BONINT", 3 * size
        End If
    Case 2:
        gemname = "Ruby"
        red = 200: green = 10: blue = 0
        addeffect2 obj, "GEM", gemname
        addeffect2 obj, "BONSTR", 2 * size
        If Not em = "" Then addeffect2 obj, "BONHP", 15 * size
    Case 3:
        gemname = "Emerald"
        red = 0: green = 150: blue = 10
        addeffect2 obj, "GEM", gemname
        addeffect2 obj, "BONDEX", 2 * size
        If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Regeneration"
    Case 4:
        gemname = "Sapphire"
        red = 0: green = 0: blue = 200
        addeffect2 obj, "GEM", gemname
        addeffect2 obj, "BONINT", 2 * size
        If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Mana Mastery"
    
    Case 5: 'Rares
    
        aroll2 = roll(8)
    
        Select Case aroll2
    
        Case 1:
            gemname = "Heartstone"
            red = 100: green = 0: blue = 0
            addeffect2 obj, "GEM", gemname
            addeffect2 obj, "BONDEX", 1 * size
            addeffect2 obj, "BONSTR", 2 * size
            If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Drain Life"
        Case 2:
            gemname = "Black Diamond"
            red = 20: green = 20: blue = 20: l = 0.5
            addeffect2 obj, "GEM", gemname
            addeffect2 obj, "BONINT", 4 * size
            addeffect2 obj, "BONDEX", 2 * size
            addeffect2 obj, "BONSKILL", Int(size / 2), "Drain Mana"
            If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Mana Mastery"
        Case 3:
            gemname = "Glowstone"
            red = 150: green = 200: blue = 250
            addeffect2 obj, "GEM", gemname
            addeffect2 obj, "BONMP", 10 * size
            addeffect2 obj, "BONSKILL", Int(size / 2), "Mana Regeneration"
            If Not em = "" Then addeffect2 obj, "BONSKILL", Int(size / 2), "Firepower"
        Case 4:
            gemname = "Vortex Stone"
            red = 40: green = 20: blue = 20
            addeffect2 obj, "GEM", gemname
            addeffect2 obj, "BONSKILL", size, "Firepower"
            If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Black Magic Mastery"
        
        Case 5:
            red = 50: green = 250: blue = 80
            gemname = "Life Rune"
            addeffect2 obj, "GEM", gemname
            addeffect2 obj, "BONSKILL", size * 2, "Regeneration"
            If Not em = "" Then addeffect2 obj, "BONHP", size * 10
            
        Case 6:
            red = 80: green = 250: blue = 150
            gemname = "Antacid Rune"
            addeffect2 obj, "GEM", gemname
            addeffect2 obj, "BONUNDIG", size * 15
            If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Defence"
            
        Case 7:
        
            croll = roll(7)
            'Ultra Rares
            Select Case croll
                Case 1:
                    red = 80: green = 250: blue = 150
                    gemname = "Gastro Rune"
                    addeffect2 obj, "GEM", gemname
                    'addeffect2 obj, "BONUNDIG", size * 15
                    addeffect2 obj, "BONSKILL", size + 1, "Giant Stomach"
                    addeffect2 obj, "BONSKILL", Int(size / 2), "Gluttony"
                    If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Super Acid"
                Case 2:
                    red = 30: green = 30: blue = 30
                    gemname = "Death Rune"
                    addeffect2 obj, "GEM", gemname
                    'addeffect2 obj, "BONUNDIG", size * 15
                    addeffect2 obj, "BONSKILL", size + 2, "Critical Strike"
                    If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Deathblow"
                
                Case 3:
                    red = 130: green = 0: blue = 150
                    gemname = "Blurring Rune"
                    addeffect2 obj, "GEM", gemname
                    'addeffect2 obj, "BONUNDIG", size * 15
                    addeffect2 obj, "BONSKILL", size + 1, "Dodge"
                    addeffect2 obj, "BONSKILL", Int(size / 2), "Evasion"
                    If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Deathblow"
                
                Case 4:
                    red = 230: green = 100: blue = 0
                    gemname = "Mana Stone"
                    addeffect2 obj, "GEM", gemname
                    'addeffect2 obj, "BONUNDIG", size * 15
                    addeffect2 obj, "BONMP", size * 40
                    If Not em = "" Then addeffect2 obj, "BONSKILL", size + 1, "Mana Regeneration"
                Case 5:
                    red = 80: green = 250: blue = 150
                    gemname = "Tummy Rune"
                    addeffect2 obj, "GEM", gemname
                    'addeffect2 obj, "BONUNDIG", size * 15
                    addeffect2 obj, "BONHP", size * 3
                    addeffect2 obj, "BONSKILL", size, "Giant Stomach"
                    If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Gluttony"
                
                
            End Select
            
        Case 8:
            gemname = "Starstone"
            red = 80: green = 200: blue = 250
            addeffect2 obj, "GEM", gemname
            addeffect2 obj, "BONMP", 15 * size
            addeffect2 obj, "BONHP", 15 * size
            If Not em = "" Then addeffect2 obj, "BONSKILL", 1 + Int(size / 3), "Accuracy"
            
            
    End Select

    Case 6:
            gemname = "Fire Rune"
            red = 250: green = 150: blue = 0
            addeffect2 obj, "GEM", gemname
            'addeffect2 obj, "BONDICE", size * 2
            If em = "" Then addeffect2 obj, "BONDICE", size Else addeffect2 obj, "BONDICE", size * 2

    Case 7:
            gemname = "Lightning Rune"
            red = 250: green = 250: blue = 50
            addeffect2 obj, "GEM", gemname
            If em = "" Then addeffect2 obj, "BONDAMAGE", size Else addeffect2 obj, "BONDAMAGE", size * 2

    Case 8:
            gemname = "Damage Rune"
            red = 250: green = 150: blue = 150
            addeffect2 obj, "GEM", gemname
            If em = "" Then addeffect2 obj, "BONUSDAMAGE", size * 10 Else addeffect2 obj, "BONUSDAMAGE", size * 20
    
    Case 9:
            gemname = "Topaz"
            red = 200: green = 150: blue = 0
            addeffect2 obj, "GEM", gemname
            addeffect2 obj, "BONHP", size * 5
            If Not em = "" Then addeffect2 obj, "BONSKILL", size, "Resilience"




End Select

obj.l = l
If Right(gemname, 4) = "Rune" Then obj.l = 0.5: If size < 3 Then obj.graphname = "rune1.bmp" Else obj.graphname = "rune2.bmp"

obj.name = sizestr & uncut & em & gemname
obj.r = red: obj.g = green: obj.b = blue

End Function



Function makemagicitem(obj As objecttype, Optional worth As Byte = 0, Optional buytype = "")

If worth = 0 Then worth = Int(averagelevel / 5)
If worth < 1 Then worth = 1
suf = ""

genworth = Int((roll(worth * 1.5) + (worth)) / 2)

defmag = 0

1

If roll(2) = 1 Then  'Prefix Power
5   aroll = roll(16)
    Select Case aroll
    Case 1:
        prname = "Demonic"
        If Not buytype = "Weapons" Or buytype = "" Then GoTo 5
        addeffect3 obj, "BONDICE", genworth * 3
        addeffect3 obj, "BONDAMAGE", genworth * 4
    Case 2:
        prname = "Jagged"
        If Not buytype = "Weapons" Or buytype = "" Then GoTo 5
        addeffect3 obj, "BONDICE", genworth
        addeffect3 obj, "BONDAMAGE", genworth + 2
    Case 3:
        prname = "Undigestable"
        addeffect3 obj, "BONUNDIG", genworth * 5 + 10
    Case 4:
        prname = "Slippery"
        addeffect3 obj, "BONSKILL", genworth, "Squirm"
                
    Case 5:
        prname = "Magical"
        addeffect3 obj, "BONMP", genworth * genworth * 3 + 3
                
    Case 6:
        prname = "Comfortable"
        If buytype = "Weapon" Then GoTo 5
        addeffect3 obj, "BONDEX", 1
                
    Case 7:
        prname = "Protective"
        addeffect3 obj, "BONSKILL", Int(genworth / 3) + 1, "Defence"
                
    Case 8:
        prname = "Maneuverable"
        addeffect3 obj, "BONSKILL", Int(genworth / 3) + 1, "Dodge"
        
    Case 9:
        prname = "Bright"
        addeffect3 obj, "BONSKILL", Int(genworth / 3) + 1, "White Magic Mastery"
        addeffect3 obj, "BONSKILL", Int(genworth / 3) + 1, "Defence"
        
    Case 10:
        prname = "Rugged"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Resilience"
        
    Case 11:
        prname = "Ablative"
        addeffect3 obj, "BONHP", genworth * genworth * 2
                
    Case 12:
        prname = "Enfeebling"
        addeffect3 obj, "BONSTR", -Int(genworth / 2) - 1: genworth = genworth * 2
        defmag = 1
                
    Case 13:
        prname = "Crippling"
        addeffect3 obj, "BONSTR", -genworth
        addeffect3 obj, "BONDEX", -Int(genworth / 2) - 1
        genworth = genworth * 3
        defmag = 1
                
    Case 14:
        prname = "Stiff"
        addeffect3 obj, "BONDEX", -Int(genworth / 2) - 1: genworth = genworth * 2
        defmag = 1
                
    Case 15:
        prname = "Enraging"
        addeffect3 obj, "BONSTR", genworth + 2
        addeffect3 obj, "BONINT", -Int(genworth / 2) - 1
    Case 16:
        prname = "Acid-resistant"
        addeffect3 obj, "BONUNDIG", genworth * 3 + 5
                
    End Select
End If


btype = roll(4) '1=Standard, 2=Specific Skill, 3+=None

If btype = 1 Or roll(2) = defmag Then
3   aroll = roll(30)
    Select Case aroll
    Case 1:
        skname = "Dexterity"
        addeffect3 obj, "BONDEX", genworth
    Case 2:
        skname = "Murder"
        If Not buytype = "Weapons" Or buytype = "" Then GoTo 3
        addeffect3 obj, "BONDICE", genworth
        addeffect3 obj, "BONDAMAGE", genworth + 2
    Case 3:
        skname = "Antacid"
        addeffect3 obj, "BONUNDIG", genworth * 5
    Case 4:
        skname = "Hunger"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Giant Stomach"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Gluttony"
    Case 5:
        skname = "Magic"
        addeffect3 obj, "BONMP", genworth * genworth * 4
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Spell Mastery"
    Case 6:
        skname = "Mana"
        addeffect3 obj, "BONMP", genworth * genworth * 4
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Mana Regeneration"
    Case 7:
        skname = "Energy"
        addeffect3 obj, "BONMP", genworth * genworth * 5
    Case 8:
        skname = "Strength"
        addeffect3 obj, "BONSTR", genworth
    Case 9:
        skname = "Intellect"
        addeffect3 obj, "BONINT", genworth
    Case 10:
        skname = "Zaera"
        If Not roll(3) = 1 Then GoTo 3 '(Magical trait is rare)
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 2, "Giant Stomach"
        addeffect3 obj, "BONSKILL", Int(genworth * 1.5) + 2, "Super Acid"
        addeffect3 obj, "BONSKILL", roll(2), "Gluttony"
    Case 11:
        skname = "Jitiiess"
        If Not roll(3) = 1 Then GoTo 3 '(Magical trait is rare)
        addeffect3 obj, "BONSKILL", genworth, "Giant Stomach"
        addeffect3 obj, "BONHP", genworth * genworth * 2 * 10 + 50
        
    Case 12:
        skname = "Yttme"
        If Not roll(3) = 1 Then GoTo 3 '(Magical trait is rare)
        addeffect3 obj, "BONSKILL", genworth, "Squirm"
        addeffect3 obj, "BONDEX", -genworth
    
    Case 13:
        skname = "Lerania"
        If Not roll(3) = 1 Then GoTo 3 '(Magical trait is rare)
        addeffect3 obj, "BONSKILL", genworth * 2, "Giant Stomach"
    
    Case 14:
        skname = "Phoebe"
        If Not roll(3) = 1 Then GoTo 3 '(Magical trait is rare)
        addeffect3 obj, "BONSKILL", genworth, "Giant Stomach"
        addeffect3 obj, "BONSKILL", -1, "Super Acid"
        addeffect3 obj, "BONSKILL", 2, "Regeneration"
    Case 15:
        skname = "The Traveller"
        addeffect3 obj, "BONHP", genworth * genworth * 3 + 10
    Case 16:
        skname = "The Warrior"
        addeffect3 obj, "BONSTR", Int(genworth / 2) + 1
        addeffect3 obj, "BONHP", genworth * genworth * 4
    Case 17:
        If Not roll(2) = 1 Then GoTo 3
        skname = "The Queen"
        addeffect3 obj, "BONSTR", Int(genworth / 2) + 1
        addeffect3 obj, "BONINT", Int(genworth / 2) + 1
        addeffect3 obj, "BONDEX", Int(genworth / 2) + 1
        addeffect3 obj, "BONHP", genworth * genworth * 8 + 20
    Case 18:
        skname = "Swallowing"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Giant Stomach"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Gluttony"
    Case 19:
        skname = "The Huntress"
        addeffect3 obj, "BONDEX", Int(genworth / 2) + 1
        addeffect3 obj, "BONDAMAGE", genworth
    Case 20:
        skname = "Protection"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Defence"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Resilience"
    Case 21:
        skname = "Mastery"
        addeffect3 obj, "BONSKILL", genworth, "Weapons Mastery"
    Case 22:
        skname = "Skill"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Weapons Mastery"
    Case 23:
        skname = "Athletics"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Dodge"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Evasion"
    
    Case 24:
        skname = "Precision"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Accuracy"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Critical Strike"
    Case 25:
        skname = "War"
        addeffect3 obj, "BONHP", genworth * genworth * 5 + 10
        addeffect3 obj, "BONDICE", Int(genworth / 2) + 1
        addeffect3 obj, "BONDAMAGE", Int(genworth / 2) + 1
    Case 26:
        skname = "Survival"
        addeffect3 obj, "BONUNDIG", genworth * 3 + 5
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Squirm"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Resilience"
    Case 27:
        skname = "Destruction"
        addeffect3 obj, "BONDICE", genworth
        addeffect3 obj, "BONDAMAGE", genworth * 2
        addeffect3 obj, "BONDEX", -genworth
    Case 28:
        skname = "Balance"
        addeffect3 obj, "BONSTR", Int(genworth / 2) + 1
        addeffect3 obj, "BONINT", Int(genworth / 2) + 1
        addeffect3 obj, "BONDEX", Int(genworth / 2) + 1
    Case 29:
        skname = "Theft"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Drain Life"
        addeffect3 obj, "BONSKILL", Int(genworth / 2) + 1, "Drain Magic"
    Case 30:
        skname = "Rage"
        addeffect3 obj, "BONSTR", Int(genworth) + 2
        addeffect3 obj, "BONINT", -Int(genworth / 2) - 1
    
    
    End Select
End If

If btype = 2 Then

    aroll = roll(29)
    Select Case aroll
        Case 1: skname = "Deathblow": suf = "s"
        Case 2: skname = "Critical Strike": suf = "ing"
        Case 3: skname = "Resilience"
        Case 4: skname = "Defence"
        Case 5: skname = "Evasion"
        Case 6: skname = "Dodge": suf = "ing"
        Case 7: skname = "Endurance"
        Case 8: skname = "Squirm": suf = "ing"
        Case 9: skname = "Drain Life"
        Case 10: skname = "Drain Magic"
        Case 11: skname = "Regeneration"
        Case 12: skname = "Firepower"
        Case 13: skname = "Accuracy"
        Case 14: skname = "Greed"
        Case 15: skname = "Spell Mastery"
        Case 16: skname = "Mana Mastery"
        Case 17: skname = "Mana Regeneration"
        Case 18: skname = "White Magic Mastery"
        Case 19: skname = "Grey Magic Mastery"
        Case 20: skname = "Black Magic Mastery"
        Case 21: skname = "Streetfighting"
        Case 22: skname = "Weapons Mastery"
        Case 23: skname = "Giant Stomach"
        Case 24: skname = "Super Acid"
        Case 25: skname = "Gluttony"
        Case 26: skname = "Sword Mastery"
        Case 27: skname = "Spear Mastery"
        Case 28: skname = "Axe Mastery"
        Case 29: skname = "Bow Mastery"
    End Select
    addeffect3 obj, "BONSKILL", genworth, skname
    skname = skname & suf

End If

If skname = "" And prname = "" Then GoTo 1

If Not skname = "" Then skname = " of " & skname

obj.name = prname & " " & obj.name & skname

End Function

Function addgem(objnum, gemnum)

If geteff(inv(gemnum), "GEM", 1) = "" Then addgem = False: Exit Function Else addgem = True

For a = 1 To 6
    If Left(inv(gemnum).effect(a, 1), 3) = "BON" Then addeffect2 inv(objnum), inv(gemnum).effect(a, 1), inv(gemnum).effect(a, 2), inv(gemnum).effect(a, 3)
Next a
zimr = inv(objnum).r
zimg = inv(objnum).g
zimb = inv(objnum).b
inv(objnum).r = (zimr + zimr + zimr + inv(gemnum).r) / 4
inv(objnum).g = (zimg + zimg + zimg + inv(gemnum).g) / 4
inv(objnum).b = (zimb + zimb + zimb + inv(gemnum).b) / 4

End Function


Function creategem(Optional ByVal X = 0, Optional ByVal Y = 0, Optional ByVal worth = 0)

If worth <= 0 Then worth = averagelevel + 1
If X = 0 Then X = roll(mapx): Y = roll(mapy)

F = createobjtype("Bug")
makegem objtypes(F), worth
createobj objtypes(F).name, X, Y

End Function

Function displayspell(spell As spellT)

If getfromstring(spell.target, 1) = "Aura" Then
    Dim donk As objecttype
    donk.name = spell.name & " (Aura)"
    addeffect2 donk, getfromstring(spell.effect, 1), getfromstring(spell.effect, 3), getfromstring(spell.effect, 2)
    addeffect2 donk, getfromstring(spell.effect, 4), getfromstring(spell.effect, 6), getfromstring(spell.effect, 5)
    addeffect2 donk, getfromstring(spell.effect, 7), getfromstring(spell.effect, 9), getfromstring(spell.effect, 8)
    displayobj donk, 2
End If

If getfromstring(spell.effect, 1) = "Damage" Then
    gamemsg spell.name & " (Attack spell, " & getfromstring(spell.amount, 1) & ")"
End If

If getfromstring(spell.effect, 1) = "Summon" Then
    gamemsg spell.name & " (Level " & getfromstring(spell.amount, 1) & " Summon)"
End If


End Function

Function displayobj(obj As objecttype, Optional disp = 1)

txt = obj.name
If txt = "" And disp < 2 Then displayobj = "Nothing": Exit Function

For a = 1 To 6
    If obj.effect(a, 1) = "Equipjunk" Then
        If Not obj.effect(a, 2) = "" Then txt = txt & ", Weight " & obj.effect(a, 2)
        If Not Val(obj.effect(a, 3)) = 0 Then txt = txt & ", " & obj.effect(a, 2) & " sockets"
        End If
    If obj.effect(a, 1) = "GEM" Then txt = txt & " (Insertable)"
    If obj.effect(a, 1) = "BONDEX" Then txt = txt & ", +" & obj.effect(a, 2) & " to dexterity"
    If obj.effect(a, 1) = "BONSTR" Then txt = txt & ", +" & obj.effect(a, 2) & " to strength"
    If obj.effect(a, 1) = "BONINT" Then txt = txt & ", +" & obj.effect(a, 2) & " to intelligence"
    If obj.effect(a, 1) = "BONARMOR" Then txt = txt & ", +" & obj.effect(a, 2) & " to armor"
    If obj.effect(a, 1) = "BONHP" Then txt = txt & ", +" & obj.effect(a, 2) & " to max hit points"
    If obj.effect(a, 1) = "BONMP" Then txt = txt & ", +" & obj.effect(a, 2) & " to max magic points"
    If obj.effect(a, 1) = "BONSKILL" Then txt = txt & ", +" & obj.effect(a, 2) & " to " & obj.effect(a, 3) & " skill"
    If obj.effect(a, 1) = "BONDICE" Then txt = txt & ", +" & obj.effect(a, 2) & " to number of weapon dice"
    If obj.effect(a, 1) = "BONDAMAGE" Then txt = txt & ", +" & obj.effect(a, 2) & " to maximum weapon damage per die"
    If obj.effect(a, 1) = "BONUSDAMAGE" Then txt = txt & ", +" & obj.effect(a, 2) & " to attack damage"
    If obj.effect(a, 1) = "BONUNDIG" Then txt = txt & ", " & obj.effect(a, 2) & "% chance of avoiding digestion"
    If obj.effect(a, 1) = "Weapon" Then txt = txt & ", " & obj.effect(a, 3) & "-" & Val(obj.effect(a, 3)) * Val(obj.effect(a, 4)) & " Damage" '& Val(obj.effect(a, 5)) & " Type"
    If obj.effect(a, 1) = "Clothes" Then txt = txt & ", AR " & obj.effect(a, 3)
Next a
If disp < 2 Then txt = txt & ", " & getworth(obj) & " Gold"
If disp >= 1 Then gamemsg txt
displayobj = txt
End Function

Function getworth(obj As objecttype) As Double

Cost = 0
totalmult = 1

weight = 1
For a = 1 To 6
    If obj.effect(a, 1) = "EQUIPJUNK" Then weight = obj.effect(a, 2)
Next a

For a = 1 To 6
    sign = Sgn(Val(obj.effect(a, 2)))
    genjunk = (Val(obj.effect(a, 2))) * Val(obj.effect(a, 2)) * sign
    If obj.effect(a, 1) = "BONDEX" Then Cost = Cost + genjunk * 100
    If obj.effect(a, 1) = "BONSTR" Then Cost = Cost + genjunk * 100
    If obj.effect(a, 1) = "BONINT" Then Cost = Cost + genjunk * 75
    If obj.effect(a, 1) = "BONHP" Then Cost = Cost + obj.effect(a, 2) * 25
    If obj.effect(a, 1) = "BONMP" Then Cost = Cost + obj.effect(a, 2) * 20
    If obj.effect(a, 1) = "BONSKILL" Then Cost = Cost + genjunk * 125
    If obj.effect(a, 1) = "BONDICE" Then Cost = Cost + obj.effect(a, 2) * 500
    If obj.effect(a, 1) = "BONDAMAGE" Then Cost = Cost + obj.effect(a, 2) * 500
    If obj.effect(a, 1) = "BONUSDAMAGE" Then Cost = Cost + obj.effect(a, 2) * 50
    If obj.effect(a, 1) = "BONUNDIG" Then Cost = Cost * (1 + obj.effect(a, 2) / 200)
    'Weapons cost dice * damage * 10, +1% per damage max
    If obj.effect(a, 1) = "Weapon" Then
        wepcost = (obj.effect(a, 3) * obj.effect(a, 4)) * 10 * (1 + (obj.effect(a, 3) * obj.effect(a, 4)) / 10)
        If obj.effect(a, 5) = "Bow" Then wepcost = wepcost * 3
        Cost = Cost + wepcost
    End If
    If obj.effect(a, 1) = "Clothes" Then
        If obj.effect(a, 5) = "" Then
        Cost = Cost + greater(1, (obj.effect(a, 3) + 1 - weight)) ^ 3
        Else:
        Cost = Cost + greater(1, ((obj.effect(a, 3) + 1 - weight) / 2.2)) ^ 3
        End If
        'cost = cost + obj.effect(a, 3) ^ 3 * (1 - (weight / 100)) ' (obj.effect(a, 3) * obj.effect(a, 3)) * 10 * ((obj.effect(a, 3) * obj.effect(a, 3)) / weight)
    End If
    If obj.effect(a, 1) = "GEM" Then totalmult = 2
Next a

Cost = Cost * totalmult
getworth = Int(greater(Cost, 1))

End Function

Function addhasbeento(name, filenandXY)

For a = 1 To 150
    If plr.beento(a, 1) = name Then Exit Function
    If plr.beento(a, 1) = "" Then plr.beento(a, 1) = name: plr.beento(a, 2) = filenandXY: Exit Function
Next a

End Function

Function hasbeento(filen)
hasbeento = False
For a = 1 To 150
    'If plr.beento(a, 1) = filen Then hasbeento = True: Exit Function
Next a
End Function

Function swapobj(ByRef obj1 As objecttype, ByRef obj2 As objecttype)

Dim obj3 As objecttype
obj3 = obj1
obj1 = obj2
obj2 = obj3

End Function

Function orginv()
'Organizes inventory, taking out empty spaces

3 For a = 1 To 49
    If inv(a).name = "" And Not inv(a + 1).name = "" Then swapobj inv(a), inv(a + 1): a = 0
Next a

Form1.updatinv

End Function

Sub checkaccess()
'Make sure all parts of map are accessible

'On Error Resume Next
4 ReDim dmap(1 To mapx, 1 To mapy) As Byte
blarg = 0
For a = mapx / 4 To mapx * 0.75
For b = mapy / 4 To mapy * 0.75
If map(a, b).blocked = 0 Then subcheck mapx / 2, mapy / 2: GoTo 5
Next b
Next a
5
'subcheck 1, 1
For X = 1 To mapx
For Y = 1 To mapy
    If dmap(X, Y) = 0 And map(X, Y).blocked = 0 Then
    ay2 = 0: ax2 = 0
    If X - 1 > 0 Then If map(X - 1, Y).blocked = 1 Then ax2 = -1: GoTo 6
    If X + 1 < mapx Then If map(X + 1, Y).blocked = 1 Then ax2 = 1: GoTo 6
    If Y - 1 > 0 Then If map(X, Y - 1).blocked = 1 Then ay2 = -1: GoTo 6
    If Y + 1 < mapy Then If map(X, Y + 1).blocked = 1 Then ay2 = 1: GoTo 6

6             ax = X: ay = Y
            Do While (1)
                ax = ax + ax2: ay = ay + ay2
                If ax2 = 0 And ay2 = 0 Then Exit Do
                If ax > mapx Or ay > mapy Or ax < 1 Or ay < 1 Then Exit Do
                If map(ax, ay).blocked = 0 Then Exit Do
                map(ax, ay).blocked = 0: map(ax, ay).ovrtile = 0
            Loop
        If ax2 <> 0 Or ay2 <> 0 Then blarg = 1
    End If
Next Y
Next X
If blarg = 1 Then GoTo 4

End Sub

Sub subcheck(ByVal X As Integer, ByVal Y As Integer, Optional ByVal ctile)

If map(X, Y).blocked = 1 Then dmap(X, Y) = 2: Exit Sub
dmap(X, Y) = 1
If Not IsMissing(ctile) Then map(X, Y).tile = ctile
If X + 1 < mapx Then If dmap(X + 1, Y) = 0 Then subcheck X + 1, Y, ctile
If X - 1 > 0 Then If dmap(X - 1, Y) = 0 Then subcheck X - 1, Y, ctile
If Y + 1 < mapy Then If dmap(X, Y + 1) = 0 Then subcheck X, Y + 1, ctile
If Y - 1 > 0 Then If dmap(X, Y - 1) = 0 Then subcheck X, Y - 1, ctile


End Sub

Sub dropitem(obj As objecttype)

If obj.graphname = "" Then obj.graphname = "clothes.bmp"
obj.graphloaded = 0
makeobjtype obj
createobj obj.name, plr.X, plr.Y
obj.name = ""
'updatinv
End Sub

Sub giverandomclothes()

Randomize

pantycolor = gencolor(, r, g, b, l)

If roll(3) > 1 Then takeobj 0, 0, 0, randclothes(pantycolor, r, g, b, l, , , "Bra")
takeobj 0, 0, 0, randclothes(pantycolor, r, g, b, l, , , "Panties")
If roll(2) = 1 Then pantycolor = gencolor(, r, g, b, l)
If roll(2) > 1 Then takeobj 0, 0, 0, randclothes(pantycolor, r, g, b, l, , , "Lower")
If roll(2) = 1 Then pantycolor = gencolor(, r, g, b, l)
If roll(2) > 1 Then takeobj 0, 0, 0, randclothes(pantycolor, r, g, b, l, , , "Upper")
If roll(4) = 1 Then takeobj 0, 0, 0, randclothes(, , , , , , , "Belt")
If roll(4) = 1 Then takeobj 0, 0, 0, randclothes(, , , , , , , "Jacket")

takeobj 0, 0, 0, randarmor
takeobj 0, 0, 0, randarmor

'addclothes "Big Mouth", "bigmouth.bmp", 0, "Mouth" ', , , , , , 1

'addclothes "Cape", "cape3.bmp", 1, "Jacket", "Wings", 255, 0, 0, , 1

End Sub

Sub checkallconvs()
'WARNING:  Only checks #CHOICE commands with standard gotos.

Form1.Text7.text = ""

Dim convs(1 To 100, 1 To 8) As String
Dim cbranches(100) As String 'Branches pointed to
Dim fbranches(100) As String 'Branches found

Open "conversations.txt" For Input As #1

Do While Not EOF(1)
    
fbranch = 1: branch = 1: choice = 1
    
5    Input #1, zerf
    If zerf = "#CHOICE" Then Input #1, filler, gotoname: cbranches(branch) = gotoname: branch = branch + 1: If gotoname = "#CHOICE" Or gotoname = "#BRANCH" Or gotoname = "" Then gamemsg "Warning: No goto for '" & filler & "' in conversation " & convn
    If zerf = "#BRANCH" Then Input #1, fbranches(fbranch): fbranch = fbranch + 1
    If zerf = "#CONVERSATION" Then
    
    For a = 1 To branch
    If Left(cbranches(a), 1) = "#" Then GoTo 6
        For b = 1 To fbranch
        If fbranches(b) = cbranches(a) Then b = 0: Exit For
        Next b
        If b > fbranch Then gamemsg "Referenced goto not found: Conversation " & convn & ", goto '" & cbranches(a) & "'"
6    Next a
    fbranch = 1
    branch = 1
    choice = 1
    Erase fbranches()
    Erase cbranches()
    Input #1, convn
    
    End If
    If Not EOF(1) Then If Not zerf = "#CONVERSATION" Then GoTo 5

Loop

Close #1

End Sub

Public Sub TimerHandler1(ByVal hwnd As Long, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTime As Long)
'If Form10.Visible = False Then Call Form1.TimerH1 Else Form10.continue
'Call Form1.TimerH1
MsgBox "TimerHandler 1 called.  This shouldn't happen."
Stop
If nodraw = 1 Then Exit Sub

'Form1.Label2.caption = Val(Form1.Label2.caption) + 1: Form1.Label2.Refresh ': Exit Sub

If Form10.Visible = True Then Form10.continue: Exit Sub

If endingprog > 0 Then endingprog = endingprog + 1: Exit Sub

If needbodyupdt = 1 Then Form1.updatbody: needbodyupdt = 0
'Timer1.Enabled = False
turnswitch = turnswitch + 1
If turnswitch >= 7 Then turnswitch = 0: turnthing 1

If stilldrawing = 0 And bijdrawing = 0 Then drawall

If plr.xoff <> 0 Or plr.yoff <> 0 Then
For a = 1 To 100
    If diff(plr.xoff, 0) >= 24 Then shooties(a).X = shooties(a).X + plr.xoff
    If diff(plr.yoff, 0) >= 24 Then shooties(a).Y = shooties(a).Y + plr.yoff
Next a
End If

If plr.xoff > 0 Then plr.xoff = plr.xoff - 24: If diff(plr.xoff, 0) < 24 Then plr.xoff = 0
If plr.xoff < 0 Then plr.xoff = plr.xoff + 24: If diff(plr.xoff, 0) < 24 Then plr.xoff = 0
If plr.yoff > 0 Then plr.yoff = plr.yoff - 12: If diff(plr.yoff, 0) < 12 Then plr.yoff = 0
If plr.yoff < 0 Then plr.yoff = plr.yoff + 12: If diff(plr.yoff, 0) < 12 Then plr.yoff = 0

Form1.updatpotions

End Sub

Public Sub Timer1Z()
'If Form10.Visible = False Then Call Form1.TimerH1 Else Form10.continue
'Call Form1.TimerH1

If nodraw = 1 Then Exit Sub

'Form1.Label2.caption = Val(Form1.Label2.caption) + 1: Form1.Label2.Refresh ': Exit Sub

If Form10.Visible = True Then Form10.continue: Exit Sub

If endingprog > 0 Then endingprog = endingprog + 1: Exit Sub

If needbodyupdt = 1 Then Form1.updatbody: needbodyupdt = 0
'Timer1.Enabled = False
turnswitch = turnswitch + 1
If turnswitch >= 7 Then turnswitch = 0: turnthing 1

If stilldrawing <= 0 And bijdrawing < 2 Then stilldrawing = stilldrawing + 1: drawall: stilldrawing = stilldrawing - 1

If plr.xoff <> 0 Or plr.yoff <> 0 Then
For a = 1 To 100
    If diff(plr.xoff, 0) >= 24 Then shooties(a).X = shooties(a).X + plr.xoff
    If diff(plr.yoff, 0) >= 24 Then shooties(a).Y = shooties(a).Y + plr.yoff
Next a
End If

If plr.xoff > 0 Then plr.xoff = plr.xoff - 24: If diff(plr.xoff, 0) < 24 Then plr.xoff = 0
If plr.xoff < 0 Then plr.xoff = plr.xoff + 24: If diff(plr.xoff, 0) < 24 Then plr.xoff = 0
If plr.yoff > 0 Then plr.yoff = plr.yoff - 12: If diff(plr.yoff, 0) < 12 Then plr.yoff = 0
If plr.yoff < 0 Then plr.yoff = plr.yoff + 12: If diff(plr.yoff, 0) < 12 Then plr.yoff = 0

Form1.updatpotions

End Sub


Public Sub TimerHandler2(ByVal hwnd As Long, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTime As Long)
'If Form10.Visible = False Then Call Form1.TimerH2 Else Form10.continue
'Call Form1.TimerH2
'DoEvents
Exit Sub
If bijdrawing > 1 Or nodraw = 1 Then Exit Sub
Static cyclez
cyclez = cyclez + 1
If cyclez >= 4 Then cyclez = 0: Timer1Z

If bijdrawing < 1 Then bijdrawing = bijdrawing + 1: drawbijdang: bijdrawing = bijdrawing - 1
End Sub

Public Sub Timer2Z()
'If Form10.Visible = False Then Call Form1.TimerH2 Else Form10.continue
'Call Form1.TimerH2
'DoEvents
Static lasttick As Long

#If USELEGACY <> 1 Then 'm''
    If lasttick = 0 Then lasttick = GetTickCount() 'm'' fix to avoid unloaded dx game
#End If 'm''

If Not GetTickCount() > lasttick + plr.timerspeed Then Exit Sub
lasttick = GetTickCount()

If Form10.Visible = True Then Form10.continue: Exit Sub
If bijdrawing > 1 Or nodraw = 1 Then Exit Sub
Static cyclez
cyclez = cyclez + 1
If cyclez >= 4 Then cyclez = 0: Timer1Z '4

If bijdrawing < 1 Then bijdrawing = bijdrawing + 1: drawbijdang: bijdrawing = bijdrawing - 1
End Sub

' This is our time callback event handler
'battle2.MoveTicker
'End Sub

Function genmap2(Optional temperature = -1, Optional moisture = -1, Optional vegetation = -1, Optional rockyness = -1, Optional population = -1, Optional filen As String = "", Optional level = 0, Optional autoload = 0, Optional northmap = "", Optional eastmap = "", Optional southmap = "", Optional westmap = "") As String

Static lastfile As Long
Dim wprefs As planetjunk

If lastfile = 0 Then lastfile = roll(300)

createplantprefs

If temperature = -1 Then temperature = roll(100)
If moisture = -1 Then moisture = roll(100)
If vegetation = -1 Then vegetation = roll(100)
If rockyness = -1 Then rockyness = roll(100)
If population = -1 Then population = roll(100)

'If temperature > 60 And moisture > 30 Then moisture = moisture - (temperature - 60)
If moisture > 20 Then If vegetation > moisture * 2 Then vegetation = moisture * 2

If moisture > 60 Then vegetation = vegetation + 10
If moisture > 80 Then vegetation = vegetation + 10

'If rockyness > 60 Then If vegetation + rockyness > 100 Then vegetation = 100 - rockyness

If moisture < 0 Then moisture = 0 'Do this for all things that can be reduced
If vegetation < 0 Then vegetation = 0

If vegetation < 10 And rockyness < 10 Then If roll(2) = 1 Then vegetation = 40 Else rockyness = 65

wprefs = createpref(0, 0, moisture, 0, temperature, 0, rockyness, 0, vegetation, 0, population, 0)

btile = tdirt 'Base tile type

'moddo = temperature / 10 - moisture / 10

btile = gettilebypref(temperature, moisture)

'If roll(1) = 1 Then btile = roll(50)

'If moddo >= 2 Then btile = tdesert
'If moddo >= 4 Then btile = 14
'If moddo <= -1 Then btile = tgrass
'If moddo <= -3 Then btile = tswamp
'If moddo <= -5 Then btile = tmud

'If rockyness > 70 Then btile = tstonydirt
'If rockyness > 85 Then btile = tstone

lastfile = lastfile + 1
If filen = "" Then filen = str(lastfile)

'If Dir(App.Path & "\" & plr.name, vbDirectory) = "" Then MkDir App.Path & "\" & plr.name

'ChDir App.Path & "\" & plr.name

Debug.Print "Saving as " & filen & ".txt"
Debug.Print "Temperature: " & temperature
Debug.Print "Moisture: " & moisture
Debug.Print "Rockyness: " & rockyness
Debug.Print "Vegetation: " & vegetation

Open filen & ".txt" For Output As #lastfile

'Fill the map with the base tile type
Write #lastfile, "#CURMAP", filen
Write #lastfile, "#PLANETJUNK", temperature, moisture, rockyness, vegetation
Write #lastfile, "#MAPSIZE", 125, 125
Write #lastfile, "#FILLMAP", btile

Write #lastfile, "#SETMAPS", northmap, eastmap, southmap, westmap

ntile = tdirt
'Add random tiles by average terrain moisture
3 tries = tries + 1
ntile1 = gettilebypref(temperature, moisture, 1)
     ntile2 = gettilebypref(temperature, moisture, 1)
    If ntile1 = btile Or ntile2 = btile Then If tries < 300 Then GoTo 3
For a = 1 To (2 + roll(3))
    Write #lastfile, "#MAPCHUNK", ntile1, roll(30) + 30
    Write #lastfile, "#MAPCHUNK", ntile2, roll(30) + 30
Next a

'Add random rocks
Write #lastfile, "#SPRINKLEOVR", rockyness * 3 + 100, orocks

'Add rivers and junk
If moisture > 40 And roll(2) = 1 Or moisture > 60 Then
    For a = 1 To roll(3)
        If temperature < moisture * 1.5 Then Write #lastfile, "#MAPCHUNK", twater, moisture / 2
    Next a
End If

'Add walls to high rockyness areas
If rockyness > 50 Then
    For a = 1 To roll(rockyness / 10) + 2
        rockyroll = rockyness + (roll(30) - 15)
        If rockyroll > 75 Then wallt = ostonewall Else wallt = odirtwall
        Write #lastfile, "#OVRCHUNK", wallt, rockyness / 2 + 10,
    Next a
End If

'Add rocky junk
If rockyness > 30 Then
For a = 1 To (1 + roll(3))
    If roll(3) = 1 Then Write #lastfile, "#MAPCHUNK", tstone, 30 Else Write #lastfile, "#MAPCHUNK", tstonydirt, roll(20) + 10
Next a
End If

'Add vegetation
Dim p1 As planetjunk
Dim p2 As planetjunk
Dim p3 As planetjunk
pickveggies wprefs, p1, p2, p3

Write #lastfile, "#EXTRAOVERS", p1.gfile, p2.gfile, p3.gfile, ""

Write #lastfile, "#SPRINKLEOVR", vegetation * 3, 51
Write #lastfile, "#SPRINKLEOVR", vegetation * 2, 52
Write #lastfile, "#SPRINKLEOVR", vegetation * 1, 53

If vegetation > 50 Then
    For a = 1 To (vegetation - 40) / 10
        Write #lastfile, "#OVRCHUNK", 51, Int(vegetation / 2)
    Next a
End If

'//EXTRAOVERS, Tree2.bmp, Tree5.bmp, Tree1.bmp, Thornytree1.bmp
Dim mon As monstertype

If level = 0 Then level = plr.level

Write #lastfile, "#LEVEL", level

For a = 1 To roll(3) + 3
    aroll = roll(4)
    t1 = btile
    If aroll = 3 Then t1 = ntile1
    If aroll = 4 Then t1 = ntile2
    mon = randmonstertype(level, , t1)
    'mon = randenzymetype(level, , t1)
    addmonsters mon, lastfile
Next a

Write #lastfile, "#RANDOMMONSTERS", "ALL", 400, 1, 1, 125, 125


addcargotomap lastfile, "Xodium Crystals", roll(300), "crystals", 20, 20, 150
addcargotomap lastfile, "Iron Ore", 600 + roll(500)
addcargotomap lastfile, "Gold Ore", 200 + roll(300), , 255, 180, 10

Write #lastfile,

Close #lastfile

genmap2 = filen & ".txt"
If autoload = 0 Then Exit Function

ChDir App.Path
loaddata filen & ".txt"


End Function

Function addcargotomap(ByVal filenum, ByVal typename, ByVal amt, Optional ByVal filen = "ore", Optional r = 150, Optional g = 150, Optional b = 150)

For a = 1 To 4
Write #filenum, "#OBJTYPE", typename & " size " & a
Write #filenum, "#EFFECT", "Cargo", typename, "", "", ""
Write #filenum, "#EFFECT", "Destruct", "", "", "", ""
Write #filenum, "#GRAPH", filen & a & ".bmp", 1, r, g, b, ".5"
Next a

Do While amt > 0
    zark = roll(Sqr(amt))
    If roll(3) = 1 Then zark = roll(3)
    If roll(5) = 1 Then zark = zark * 3
    amt = amt - zark

    size = 1
    If zark > 3 Then size = 2
    If zark > 8 Then size = 3
    If zark > 15 Then size = 4

    Write #filenum, "#CREATEOBJ", typename & " size " & size, 0, 0, typename, zark, ""
Loop

End Function

Function pickveggies(prefs As planetjunk, ByRef p1 As planetjunk, ByRef p2 As planetjunk, ByRef p3 As planetjunk)

Dim cplant(1 To 3) As planetjunk
Dim plantlist() As planetjunk
ReDim plantlist(1 To UBound(plantprefs())) As planetjunk
'cplant = 0 'Current selected plant
fit = 0 'How good the current fit is
For a = 1 To UBound(plantprefs())
    plantlist(a) = plantprefs(a)
Next a

For a = 2 To UBound(plantprefs())
    If compareprefs(plantlist(a), prefs) > compareprefs(plantlist(a - 1), prefs) Then swappref plantlist(a), plantlist(a - 1): a = 1
Next a

p1 = plantlist(1)
p2 = plantlist(2)
p3 = plantlist(3)

End Function

Function swappref(ByRef a As planetjunk, ByRef b As planetjunk)
Dim c As planetjunk
c = a
a = b
b = c
End Function

Function within(checknum, num1, num2) As Integer

End Function

Function pickbetter(val1, val2, valB1, valB2, num1, num2, tol1, tol2)
'Picks the better between two based on two criteria


End Function

Function compareprefs(o1 As planetjunk, o2 As planetjunk)

zork = zork + within2(o1.moisture, o2.moisture, o1.moisturetolerance)
zork = zork + within2(o1.temperature, o2.temperature, o1.temperaturetolerance)
zork = zork + within2(o1.rockyness, o2.rockyness, o1.rockynesstolerance)
zork = zork + within2(o1.population, o2.population, o1.populationtolerance)
zork = zork + within2(o1.vegetation, o2.vegetation, o1.vegetationtolerance)
compareprefs = zork

End Function

Function createpref(name, gfile, moisture, moisturetolerance, temperature, temperaturetolerance, rockyness, rockynesstolerance, vegetation, vegetationtolerance, population, populationtolerance) As planetjunk

Dim o1 As planetjunk
o1.name = name
o1.gfile = gfile
o1.moisture = moisture
o1.moisturetolerance = moisturetolerance
o1.temperature = temperature
o1.temperaturetolerance = temperaturetolerance
o1.rockyness = rockyness
o1.rockynesstolerance = rockynesstolerance
o1.vegetation = vegetation
o1.vegetationtolerance = vegetationtolerance
o1.population = population
o1.populationtolerance = populationtolerance

createpref = o1

End Function

Function within2(checknum, num1, tolerance) As Integer
'Returns 1-10--the higher the number, the closer the number is to it's optimum level
'If the value is outside the range, it will return a negative number

'wok = 0

'If num1 - tolerance < checknum Then wok = checknum - (num1 - tolerance)
'If num1 + tolerance > checknum Then wok2 = (num1 + tolerance) - checknum
If tolerance = 100 Then within2 = 1: Exit Function
num = diff(checknum, num1)
'If num > tolerance Then within2 = 0: Exit Function
If num = 0 Then num = 1
dork = 10 - (10 * (num / tolerance))
If dork < 10 - num Then dork = 10 - num 'Small ranges won't mean small returns
'If dork < 0 Then dork = dork / 2

within2 = dork

End Function

Function savemastermap(filen)


Open filen For Binary As #1
Put #1, , mastermap()
Close #1


End Function

Function createplantprefs()

ReDim plantprefs(1 To 10) As planetjunk

plantprefs(1) = createpref("Dead Tree", "Tree1.bmp", 20, 40, 50, 60, 50, 100, 30, 30, 50, 100)
plantprefs(2) = createpref("Sparse Tree", "Tree2.bmp", 40, 20, 40, 10, 50, 100, 30, 10, 50, 100)
plantprefs(3) = createpref("Swamp Tree", "Tree3.bmp", 70, 20, 40, 30, 40, 20, 60, 20, 50, 100)
plantprefs(4) = createpref("Big Swamp Tree", "Tree4.bmp", 80, 20, 40, 30, 40, 20, 70, 30, 50, 100)
plantprefs(5) = createpref("Green Tree", "Tree5.bmp", 50, 20, 45, 15, 50, 100, 30, 30, 50, 100)
plantprefs(6) = createpref("Orange Tree", "Tree6.bmp", 45, 15, 55, 15, 50, 100, 30, 30, 50, 100)
plantprefs(7) = createpref("Orange Sparse Tree", "Tree7.bmp", 35, 15, 55, 25, 50, 100, 30, 30, 50, 100)
plantprefs(8) = createpref("Red Crag", "redcrag1.bmp", 10, 10, 90, 15, 50, 100, 15, 15, 50, 100)
plantprefs(9) = createpref("Red Crag 2", "redcrag2.bmp", 10, 10, 90, 15, 70, 100, 15, 15, 50, 100)
plantprefs(10) = createpref("Thorn Tree", "thorntree1.bmp", 15, 10, 60, 25, 50, 100, 15, 15, 50, 100)


End Function

Function createmonsterprefs()

ReDim plantprefs(1 To 50) As planetjunk

'plantprefs(1)=createpref("Kari","kari1.bmp"

End Function

Function genmap3(mapxs, mapys, tile1, tile2, tile3, ovr1, ovr2, ovr3, wall1, wall2, Optional tileamt = 40, Optional ovramt = 30, Optional wallamt = 40, Optional rocks = 5)

mapx = mapxs
mapy = mapys
ReDim map(1 To mapx, 1 To mapy) As tiletype: updatmap

    fillmap tile1

For a = 1 To (mapx / 50) * (mapy / 50)
    randomchunk tileamt, , , tile2
    randomchunk tileamt, , , tile2
    randomchunk tileamt, , , tile3
    
    sprinkleovr ovramt, ovr1
    sprinkleovr ovramt, ovr1
    sprinkleovr ovramt, ovr2
    
    If rocks > 0 Then sprinkleovr ovramt, rocks
    
    ovrchunk wallamt, , , wall1
    ovrchunk wallamt, , , wall1
    ovrchunk wallamt, , , wall2
Next a

End Function

Function gettilebypref(temperature, moisture, Optional variance = 0)

5 t = Int(temperature / 20)
m = Int(moisture / 20)
tt = 0 'Tile Type

If variance = 1 Then
8 t = t + roll(3) - 2
m = m + roll(3) - 2
If t < 0 Or m < 0 Then GoTo 5
End If

If t = 0 Then
    If m = 0 Then tt = tice
    If m = 1 Then tt = tsnow
    If m = 2 Then tt = tswamp
    If m = 3 Then tt = tswamp
    If m >= 4 Then tt = tmud
End If

If t = 1 Then
    If m = 0 Then tt = tice
    If m = 1 Then tt = tsnow
    If m = 2 Then tt = tgrass
    If m = 3 Then tt = tdirt
    If m >= 4 Then tt = tmud
End If

If t = 2 Then
    If m = 0 Then tt = tsnow
    If m = 1 Then tt = tgrass
    If m = 2 Then tt = tdirt
    If m = 3 Then tt = tsand
    If m >= 4 Then tt = tdesert
End If

If t = 3 Then
    If m >= 4 Then tt = tarid
    If m = 0 Then tt = tgrass
    If m = 1 Then tt = tdirt
    If m = 2 Then tt = tsand
    If m = 3 Then tt = tdesert
End If

If t >= 4 Then
    If m = 3 Then tt = tarid
    If m >= 4 Then tt = tsuperarid
    If m = 0 Then tt = tdirt
    If m = 1 Then tt = tsand
    If m = 2 Then tt = tdesert
End If

gettilebypref = tt

End Function

Function savemapseg(filen)

Open filen For Binary As #1

Put #1, , mapx
Put #1, , mapy

Put #1, , map()

Close #1

End Function

Function loadmapseg(filen)

oldmapx = mapx
oldmapy = mapy

Open filen For Binary As #1

Get #1, , mapx
Get #1, , mapy

If mapx < 1 Or mapx > 500 Or mapy < 1 Or mapy > 500 Then Debug.Print "Segment loading error": Close #1: mapx = oldmapx: mapy = oldmapy: Exit Function

ReDim map(1 To mapx, 1 To mapy)

Get #1, , map()

Close #1

End Function

Function getraster(obj1 As objecttype) As RasterOpConstants

getraster = vbSrcCopy
If geteff(obj1, "Translucent", 2) > 0 Then getraster = vbSrcAnd

End Function

Function loadseg(ByVal segfilen As String, X, Y, rotation, tilet, ovrtile, Optional salesobj As String = "", Optional signobj As String = "", Optional bij1 As String = "", Optional bij2 As String = "", Optional doorobj As String = "", Optional stealobj As String = "")

If Dir(segfilen) = "" Then ChDir App.Path: If Dir(segfilen) = "" Then gamemsg "Segment file '" & segfilen & "' not found": Exit Function
filenum = FreeFile

Open segfilen For Binary As filenum
Dim segx As Integer
Dim segy As Integer
Dim munkee() As tiletype
Dim munkee2() As tiletype

Get #filenum, , segx
Get #filenum, , segy

If segx < 1 Or segy < 1 Or segx > 300 Or segy > 300 Then gamemsg "Error in " & segfilen: Close #1: Exit Function

ReDim munkee(1 To segx, 1 To segy) As tiletype

Get #filenum, , munkee()

Close filenum

If rotation > 0 Then

    ReDim munkee2(1 To segx, 1 To segy) As tiletype



    For rot = 1 To rotation
        
        'Assign current segment to buffer segment
        For a = 1 To segx: For b = 1 To segy: munkee2(a, b) = munkee(a, b): Next b: Next a
        
        'rotate real segment via buffer segment
        For a = 1 To segx: For b = 1 To segy
            
            'Rotation formula for 90 degree turn:
            'Y=X, X=Max+1-Y
                    
            munkee(segx + 1 - b, a) = munkee2(a, b)
                    
        Next b: Next a
        
    Next rot

End If

'Put segment on map
For a = 1 To segx: For b = 1 To segy
    map(a + X, b + Y) = munkee(a, b)
Next b: Next a

'Replace objects on actual map
For a = X To X + segx: For b = Y To Y + segy
    map(a, b).tile = tilet
    If map(a, b).ovrtile = 25 Then map(a, b).ovrtile = 0: map(a, b).blocked = 0: If Not signobj = "" Then createobj signobj, a, b, signobj & roll(3000)
    If map(a, b).ovrtile = 24 Then map(a, b).ovrtile = ovrtile
    If map(a, b).ovrtile = 23 Then map(a, b).ovrtile = 0: map(a, b).blocked = 0: If Not salesobj = "" Then createobj salesobj, a, b, salesobj & roll(3000)
    If map(a, b).ovrtile = 22 Then map(a, b).ovrtile = 0: map(a, b).blocked = 0: If Not stealobj = "" Then createobj stealobj, a, b, stealobj & roll(3000)
    If map(a, b).ovrtile = 21 Then map(a, b).ovrtile = 0: map(a, b).blocked = 0: If Not doorobj = "" Then createobj doorobj, a, b, doorobj & roll(3000)
    If map(a, b).ovrtile = 18 Then map(a, b).ovrtile = 0: map(a, b).blocked = 0: If Not bij1 = "" Then createobj bij1, a, b, bij1 & roll(3000)
    If map(a, b).ovrtile = 19 Then map(a, b).ovrtile = 0: map(a, b).blocked = 0: If Not bij2 = "" Then createobj bij2, a, b, bij2 & roll(3000)

Next b: Next a


End Function

Sub loadrain(r, g, b, speed, density, wander)

'Public rainx(50) As Integer
'Public rainy(50) As Integer
'Public raincolor As Long
'Public rainspeed As Long
'Public raindensity As Long
'Public raingraph As cSpriteBitmaps

End Sub

Sub loadstationobjs()

F = createobjtype("Treasure Bag (Green)", "treasure2.bmp", 0, 180, 0, 0.5)
addeffect F, "GiveGold", "50"
addeffect F, "Destruct"


End Sub

Function loadalltiles(Optional r = -1, Optional g = 0, Optional b = 0, Optional translucency = 1.5)
Set tilespr = New cSpriteBitmaps
tilespr.CreateFromFile "tiles1.bmp", 5, 5, , RGB(0, 0, 0)

'tilespr.recolor 0, 0, 0, , , , 1
If r > -1 Then tilespr.recolorall r, g, b, , , , translucency

Set tilespr2 = New cSpriteBitmaps
If isexpansion = 1 Then tilespr2.CreateFromFile "tiles2.bmp", 5, 5, , RGB(0, 0, 0)

'tilespr2.recolor 0, 0, 0, , , , 1
If r > -1 Then tilespr2.recolorall r, g, b, , , , translucency

If isexpansion = 1 Then
Set ovrspr = New cSpriteBitmaps
ovrspr.CreateFromFile "overlays2.bmp", 5, 10, , RGB(0, 0, 0)
Else:
Set ovrspr = New cSpriteBitmaps
ovrspr.CreateFromFile "overlays1.bmp", 5, 5, , RGB(0, 0, 0)
End If

'ovrspr.recolor 0, 0, 0, , , , 1
If r > -1 Then ovrspr.recolorall r, g, b, , , , translucency

If isexpansion = 1 Then
Set transovrspr = New cSpriteBitmaps
transovrspr.CreateFromFile "transoverlays2.bmp", 5, 10, , RGB(0, 0, 0)
Else:
Set transovrspr = New cSpriteBitmaps
transovrspr.CreateFromFile "transoverlays.bmp", 5, 5, , RGB(0, 0, 0)
End If

'transovrspr.recolor 0, 0, 0, , , , 1
If r > -1 Then transovrspr.recolorall r, g, b, , , , translucency

End Function

Sub drawtexts()

'm'' added some declaration to avoid automation errors
Dim lX_CT As Long 'm'' X coordinate of text to print
Dim lY_CT As Long 'm'' Y coordinate of text to print
Dim a As Long 'm''

#If USELEGACY <> 1 Then
'm'' experimental UI
Add_UI.drawui 'm''
#End If

With picBuffer 'm'' avoid object long jump load

For a = 1 To UBound(drawtxts()) 'To 1 Step -1 'Draw in reverse order
    If drawtxts(a).age = 0 Or drawtxts(a).txt = "" Then GoTo 5
    'age = 30 - drawtxts(a).age * 5: If age < 0 Then age = 0
    age = 25 - drawtxts(a).age: If age < 0 Then age = 0
    age2 = drawtxts(a).age
    'Shadow
    r = drawtxts(a).highr
    g = drawtxts(a).highg
    b = drawtxts(a).highb
    
    r2 = drawtxts(a).lowr
    g2 = drawtxts(a).lowg
    b2 = drawtxts(a).lowb
    
    
    'If age < 3 Then GoTo 7 'Do not fade for first five frames
    'r = ((r * age2) + r2) / age2 + 1
    'g = ((g * age2) + g2) / age2 + 1
    'b = ((b * age2) + b2) / age2 + 1
    
    rd3 = diff(r, r2)
    rd3 = rd3 * (1 - age / 30)
    r = r2 + rd3
    
    g3 = diff(g, g2)
    g3 = g3 * (1 - age / 30)
    g = g2 + g3
    
    b3 = diff(b, b2)
    b3 = b3 * (1 - age / 30)
    b = b2 + b3
    
7
    'If spaceon = 1 Then
    xadd = 400: yadd = 400
    
    'Dim sfont As IFont
    'sfont.size = 12
    'sfont.name = "Arial"
    'sfont.bold = True
    'picBuffer.SetFont Font
    
    .SetForeColor RGB(r / 2, g / 2, b / 2)
    lX_CT = drawtxts(a).X + xadd + 1 'm'' VB will correctly convert the result
    lY_CT = drawtxts(a).Y + yadd + 1 'm'' VB will correctly convert the result
    .drawtext lX_CT, lY_CT, drawtxts(a).txt, False
    'Main Text
    .SetForeColor RGB(r, g, b)
    .drawtext drawtxts(a).X + xadd, drawtxts(a).Y + yadd, drawtxts(a).txt, False

    
    
    
    drawtxts(a).Y = drawtxts(a).Y - 2
    drawtxts(a).age = drawtxts(a).age - 1
    If drawtxts(a).age <= 0 Then drawtxts(a).txt = ""
5 Next a

End With 'm''

End Sub

Sub drawtext(text, X, Y, r, g, b, Optional bold = False)
    X = Int(X)
    
    xadd = 400: yadd = 400
    picBuffer.SetForeColor RGB(r / 2, g / 2, b / 2)
    picBuffer.drawtext X + xadd + 2, Y + yadd + 2, text, False
    'Main Text
    picBuffer.SetForeColor RGB(r, g, b)
    picBuffer.drawtext X + xadd, Y + yadd, text, False

End Sub

Sub addtext(ByVal txt, Optional ByVal X = 300, Optional ByVal Y = 200, Optional ByVal r = 250, Optional ByVal g = 250, Optional ByVal b = 250, Optional ByVal age = 30, Optional ByVal textid As String = "", Optional ByVal r2 = -1, Optional ByVal g2 = -1, Optional ByVal b2 = -1)

If spaceon = 1 Then X = X + 400 - plrship.X + sxoff
If spaceon = 1 Then Y = Y + 300 - plrship.Y + syoff

'If textid = "plrshields" Then Stop



For a = 1 To UBound(drawtxts())
    If Not textid = "" And Not textid = "0" And drawtxts(a).textid = textid Then
    If Val(txt) > -1 And Val(drawtxts(a).txt) > 0 Then txt = Int(Val(drawtxts(a).txt)) + Int(Val(txt))
    GoTo 3
    End If
    
    If drawtxts(a).X = X Then If diff(drawtxts(a).Y, Y) < 5 Then drawtxts(a).Y = drawtxts(a).Y - 10
    If drawtxts(a).age = 0 Then
3   drawtxts(a).X = X
    drawtxts(a).Y = Y
    drawtxts(a).highr = r
    drawtxts(a).highg = g
    drawtxts(a).highb = b
    
    If r2 = -1 Then r2 = r / 2
    If g2 = -1 Then g2 = g / 2
    If b2 = -1 Then b2 = b / 2
    
    drawtxts(a).lowr = r2
    drawtxts(a).lowb = b2
    drawtxts(a).lowg = g2
    
    drawtxts(a).age = age
    drawtxts(a).txt = txt
    drawtxts(a).textid = textid
    Exit Sub
    End If
Next a

End Sub

Function calcforshot(ByVal realxfrom, ByVal realyfrom, ByVal realxto, ByVal realyto, ByVal targxmove, ByVal targymove, ByVal wepspeed, ByRef xmove, ByRef ymove, ByRef angle, Optional totalcells = 18)
'Takes the target's speed into account and aims appropriately

'Figure out time it will take to hit the target

'First, figure out weapon speed
calcangle realxfrom, realyfrom, realxto, realyto, xmove2, ymove2, wepspeed, angle2, totalcells

'Find flight time by comparing X values
realxto2 = diff(realxto, realxfrom)
realxto2 = posit(realxto2)
xmove2 = posit(xmove2)

xtime = realxto2 / xmove2
If xtime < 1 Then xtime = 1

realyto2 = diff(realyto, realyfrom)
realyto2 = posit(realyto2)
ymove2 = posit(ymove2)

ytime = realyto2 / ymove2
If ytime < 1 Then ytime = 1

'If xtime > 20 Or ytime > 20 Then Stop

'If ftime < 15 Then GoTo 5 'For short periods, just fire straight

'Find target's new position
realxto = realxto + (targxmove * xtime)
realyto = realyto + (targymove * ytime)

'Then fire at that position

5 calcangle realxfrom, realyfrom, realxto, realyto, xmove, ymove, wepspeed, angle, totalcells

End Function

Function turntowards(ByVal angle, ByRef angletochange)

If angle = angletochange Then Exit Function

angle = Int(angle)
ang3 = angle
ang4 = angle

For a = 1 To 18


    If ang3 > 18 Then ang3 = 1
    If ang4 = 0 Then ang4 = 18
    If ang3 = angletochange Then angletochange = angletochange + 1: If angletochange > 18 Then angletochange = 1: Exit Function
    If ang4 = angletochange Then angletochange = angletochange - 1: If angletochange < 1 Then angletochange = 18: Exit Function
    ang3 = ang3 - 1
    ang4 = ang4 + 1

Next a

End Function

Function randarmor2(ByRef gold, worth, Optional ByVal forcetype = "") As objecttype

armn = ""
armult = 1

3 aroll = worth + roll(3) - 2 'roll(7)
'If aroll = 11 Then Stop
If aroll > 11 Then aroll = 11
If aroll < 1 Then GoTo 3
'If worth > 11 Then aroll = 1
If forceworth > 0 Then aroll = forceworth
If aroll = 1 Then armn = "Leather": r = 105: g = 75: b = 0: l = 0.4: armult = 1: weight = weight * 3
If aroll = 2 Then armn = "Bronze": r = 55: g = 35: b = 0: l = 0.2: armult = 1.5: weight = weight * 8
If aroll = 3 Then armn = "Iron": r = 105: g = 105: b = 125: l = 0.2: armult = 2: weight = weight * 10
If aroll = 4 Then armn = "Steel": r = 155: g = 155: b = 155: l = 0.7: armult = 2.5: weight = weight * 7
If aroll = 5 Then armn = "Opal": r = 255: g = 195: b = 215: l = 0.5: armult = 3.5: weight = weight * 6
If aroll = 6 Then armn = "PinkSteel": r = 255: g = 105: b = 240: l = 0.5: armult = 4: weight = weight * 6
If aroll = 7 Then armn = "Enchanted Obsidian": r = 155: g = 155: b = 155: l = 0: armult = 5: weight = weight * 7
If aroll = 8 Then armn = "BloodSteel": r = 255: g = 0: b = 0: l = 0.3: armult = 6: weight = weight * 9
If aroll = 9 Then armn = "Blue Adamant": r = 0: g = 0: b = 255: l = 0.4: armult = 8: weight = weight * 12
If aroll = 10 Then armn = "Demonic Iron": r = 105: g = 15: b = 10: l = 0.1: armult = 10: weight = weight * 14
If aroll = 11 Then armn = "Angelic Steel": r = 255: g = 225: b = 150: l = 0.6: armult = 12: weight = weight * 5

'Feathersteel (Ultra-light)
'Magicite (Light, mana bonus)
'

armor = armor * armult: If armor < 1 Then armor = 1
gold = (armor * armor) * 10 '* ((roll(10) + 10) / 10)) + (armor * 20)) / 2

gold = gold * lesser(0.6, (1 - weight / 200)) '1/2% price reduction per weight, max 60%

'randarmortype = armn

End Function

Function getmoncolorbyname(ByRef name As String, ByRef r, ByRef g, ByRef b, Optional ByRef levelmod)

If name = "" Or name = "Generic" Then name = getgener("Snow", "Frost", "Ice", "Soot", "Ash", "Dark")

If name = "Snow" Then r = 230: g = 230: b = 250: levelmod = 1
If name = "Frost" Then r = 130: g = 180: b = 250: levelmod = 1
If name = "Ice" Then r = 30: g = 130: b = 250: levelmod = 1

If name = "Fire" Then r = 250: g = 80: b = 0: levelmod = 1
If name = "Flame" Then r = 200: g = 30: b = 0: levelmod = 0.8
If name = "Magma" Then r = 250: g = 100: b = 0: levelmod = 1.2

If name = "Forest" Then r = 30: g = 80: b = 0: levelmod = 1
If name = "Jungle" Then r = 150: g = 30: b = 0: levelmod = 0.8
If name = "Glade" Then r = 130: g = 80: b = 120: levelmod = 1.2

If name = "Desert" Then r = 215: g = 180: b = 90: levelmod = 0.8
If name = "Sand" Then r = 215: g = 180: b = 60: levelmod = 1.2
If name = "Arid" Then r = 250: g = 150: b = 90: levelmod = 1.2

'Tundra

If name = "Mountain" Then r = 90: g = 90: b = 70: levelmod = 1.1
If name = "Stone" Then r = 170: g = 155: b = 150: levelmod = 1

If name = "Soot" Then r = 60: g = 60: b = 60: levelmod = 0.9
If name = "Ash" Then r = 40: g = 40: b = 40: levelmod = 1
If name = "Dark" Then r = 20: g = 20: b = 20: levelmod = 1.1

If name = "Foul" Then r = 140: g = 120: b = 60: levelmod = 1
If name = "Swamp" Then r = 95: g = 115: b = 55: levelmod = 1

If name = "Sea" Then r = 0: g = 115: b = 115: levelmod = 1
If name = "Water" Then r = 0: g = 80: b = 115: levelmod = 1



End Function

Function randmonstertype(level, Optional typename As String = "", Optional ByVal tt = 0) As monstertype

If typename = "" Then typename = getmontype(tt)
Dim mon As monstertype

getmoncolorbyname typename, r, g, b, levelmod
mon.color = RGB(r, g, b)
mon.light = 0.5
mon.level = Int(level * levelmod)

aroll = roll(41)
    Select Case aroll
        Case 1: mon.name = "Kari": mon.gfile = "kari1.bmp"
        Case 2: mon.name = "Kari": mon.gfile = "kari1.bmp"
        Case 3: mon.name = "Giantess": mon.gfile = "giantess.bmp"
        Case 4: mon.name = "Naga": mon.gfile = "snakewoman1.bmp"
        Case 4: mon.name = "Girlserpent": mon.gfile = "snakewoman1.bmp"
        Case 5: mon.name = "Snakewoman": mon.gfile = "snakewoman2.bmp"
        Case 6: mon.name = "Snake": mon.gfile = "snake1.bmp"
        Case 7: mon.name = "Serpent": mon.gfile = "snake1.bmp"
        Case 7: mon.name = "Python": mon.gfile = "snake1.bmp"
        Case 8: mon.name = "Worm": mon.gfile = "worm1.bmp"
        Case 9: mon.name = "Vines": mon.gfile = "tendrils1.bmp"
        Case 10: mon.name = "Tentacles": mon.gfile = "tendrils1.bmp"
        Case 11: mon.name = "Plant": mon.gfile = "venus1.bmp"
        Case 12: mon.name = "Trap": mon.gfile = "venus1.bmp"
        Case 13: mon.name = "Frog": mon.gfile = "frog1.bmp"
        Case 14: mon.name = "Toad": mon.gfile = "frog1.bmp"
        Case 15: mon.name = "Slime": mon.gfile = "slime1.bmp"
        Case 16: mon.name = "Goo": mon.gfile = "slime1.bmp"
        Case 17: mon.name = "Centauress": mon.gfile = "centauress1.bmp"
        Case 18: mon.name = "Mare": mon.gfile = "centauress1.bmp"
        Case 19: mon.name = "Harpy": mon.gfile = "harpy1.bmp"
        Case 20: mon.name = "Raven": mon.gfile = "harpy1.bmp"
        Case 21: mon.name = "Sprite": mon.gfile = "sprite1.bmp"
        Case 22: mon.name = "Faerie": mon.gfile = "sprite1.bmp"
        Case 23: mon.name = "Succubus": mon.gfile = "succubus1.bmp"
        Case 24: mon.name = "Batwoman": mon.gfile = "succubus1.bmp"
        Case 25: mon.name = "Dragon": mon.gfile = "thirsha1.bmp"
        Case 27: mon.name = "Winged Serpent": mon.gfile = "wingedserpent1.bmp"
        Case 28: mon.name = "Snakebat": mon.gfile = "wingedserpent1.bmp"
        Case 29: mon.name = "Alligator": mon.gfile = "croc2.bmp"
        Case 30: mon.name = "Crocodile": mon.gfile = "croc2.bmp"
        Case 31: mon.name = "Spiderwoman": mon.gfile = "kari1.bmp"
        Case 32: mon.name = "Grub": mon.gfile = "worm1.bmp"
        Case 33: mon.name = "Elemental": mon.gfile = "elemental1.bmp": mon.colorwhole = 3
        Case 34: mon.name = "Demonette": mon.gfile = "demoness2.bmp"
        Case 35: mon.name = "Demongirl": mon.gfile = "demoness2.bmp"
        Case 36: mon.name = "Demoness": mon.gfile = "demoness1.bmp"
        Case 37: mon.name = "Flower": mon.gfile = "flower1.bmp"
        Case 38: mon.name = "Spitweed": mon.gfile = "flower1.bmp"
        Case 39: mon.name = "Xebbebba": mon.gfile = "xebebba.bmp"
        Case 40: mon.name = "Wyvern": mon.gfile = "wyvern1.bmp"
        Case 41: mon.name = "Drake": mon.gfile = "wyvern1.bmp"
        'Case 42: mon.name = "Xebbebba": mon.gfile = "xebebba.bmp"
        
    End Select

'mon.gfile = "xebebba.bmp"

mon.name = typename & " " & mon.name

randmonstertype = mon

End Function

Function getmontype(Optional ByVal tt = 0)

' tgrass = 1
' tdesert = 2
' tstonydirt = 3
' twater = 8
' tswamp = 9
' tdirt = 12
' tmud = 13
' tblackstone = 14
' tstone = 15
' tlavarock = 17
' tlava = 18
' tsand = 27
' tarid = 29
' tsnow = 31
' tice = 43
' tsuperarid = 47

    If tt = 0 Then tt = map(roll(mapx), roll(mapy)).tile
    gmontype = "Generic"
    'if tt=t
    If tt < 50 Then gmontype = getgener("Forest", "Glade", "Jungle")
    If tt = tgrass Then gmontype = getgener("Forest", "Glade", "Jungle")
    If tt = tmud Then gmontype = getgener("Mud", "Swamp", "Foul")
    If tt = tsand Then gmontype = getgener("Sand", "Stone")
    If tt = tdesert Or tt = tarid Then gmontype = getgener("Desert", "Sand", "Fire", "Arid")
    If tt = tdirt Then gmontype = getgener("Mud", "Swamp", "Foul")

getmontype = gmontype

End Function

Function addmonsters(mon As monstertype, filenum)

getrgb mon.color, r, g, b
If Val(mon.level) < 1 Then mon.level = 1
If Not mon.name = "" Then Write #filenum, "#MONTYPE2", mon.name, mon.gfile, mon.level, r, g, b, 0.5


End Function

Function genmap4(filenum, mapxs, mapys, tile1, tile2, tile3, ovr1, ovr2, ovr3, wall1, wall2, Optional tileamt = 40, Optional ovramt = 30, Optional wallamt = 40, Optional rocks = 5)

mapx = mapxs
mapy = mapys
ReDim map(1 To mapx, 1 To mapy) As tiletype: updatmap

    fillmap tile1

For a = 1 To (mapx / 50) * (mapy / 50)
    If filenum = 0 Then GoTo 3
    Write #filenum, "#MAPCHUNK", tile1, tileamt
    Write #filenum, "#MAPCHUNK", tile2, tileamt
    Write #filenum, "#MAPCHUNK", tile3, tileamt
    Write #filenum, "#OVRCHUNK", tile3, tileamt
    'randomchunk tileamt, , , tile2
    'randomchunk tileamt, , , tile3
    
3   sprinkleovr ovramt, ovr1
    sprinkleovr ovramt, ovr1
    sprinkleovr ovramt, ovr2
    
    If rocks > 0 Then sprinkleovr ovramt, rocks
    
    ovrchunk wallamt, , , wall1
    ovrchunk wallamt, , , wall1
    ovrchunk wallamt, , , wall2
Next a

End Function

'Function wordgenmap(word)

'If word = "Swamp" Then
    't1 = tswamp: t2 = tmud: t3 = tdirt
'    genmap4 100, 100, tswamp, tdirt, tmud, 5, 5, 5, 0, 0
    
'End If

'End Function

Function addfatigue(amt, Optional noweight = 0)
'Adds fatigue, affected by armor vs. strength

'If isexpansion = 0 Then Exit Function

'plrweight = getplrweight

amt = amt + plr.monsinbelly

If plr.endurance = 0 Then plr.endurance = plr.str + plr.dex
plr.fatiguemax = greater(50, getend * 30 + 50)

'Weight increases by 15% per point of weight above str+endurance
If noweight = 0 Then weightmult = (getplrweight - getstr(1) + getend) / 7
'If noweight = 0 Then weightmult = 1 + (plr.monsinbelly / 2) + (getplrweight / greater(getstr(1) + 3, 4)) / 20 Else weightmult = 1
If weightmult < 1 Then weightmult = 1

plr.fatigue = plr.fatigue + amt * weightmult

If plr.fatigue > plr.fatiguemax Then plr.fatigue = plr.fatiguemax
End Function

Function getplrweight()
Static weight
If updatbonuses = 0 Then getplrweight = weight: Exit Function

weight = 0

For a = 1 To UBound(clothes())
    If clothes(a).loaded = 1 And clothes(a).weight = 0 Then clothes(a).weight = geteff(clothes(a).obj, "Equipjunk", 2)
    If clothes(a).loaded = 1 Then weight = weight + clothes(a).weight
Next a

'If wep.weight = 0 Then
wep.weight = geteff(wep.obj, "Equipjunk", 2)
weight = weight + wep.weight

getplrweight = weight

End Function

Function subfatigue(amt)

If amt < 1 Then amt = 1
plr.fatigue = plr.fatigue - amt: If plr.fatigue < 0 Then plr.fatigue = 0

End Function

Function drawstatbars(X, Y)
'm'' part of integrated GUI draw
Static statbarsloaded As Byte

If statbarsloaded = 0 Then loadstatbars: statbarsloaded = 1

bars.empty.TransparentDraw picBuffer, X, Y, 1
If plr.hplost < gethpmax Then bars.life.TransparentDraw picBuffer, X, Y, 1, , , , ((gethpmax - plr.hplost) / gethpmax) * 240
If plrhpgain > 0 Then bars.fatigue.TransparentDraw picBuffer, X, Y, 1, , , , ((plr.hp + plrhpgain) / gethpmax) * 240
If plr.hp > 1 Then bars.life2.TransparentDraw picBuffer, X, Y, 1, , , , (plr.hp / gethpmax) * 240

If plr.mpmax < 1 Then GoTo 5
bars.empty.TransparentDraw picBuffer, X, Y + 10, 1
If plr.mp > 0 Then bars.mana.TransparentDraw picBuffer, X, Y + 10, 1, , , , (plr.mp / plr.mpmax) * 240

5
bars.empty.TransparentDraw picBuffer, X, Y + 20, 1
If plr.fatigue < plr.fatiguemax Then bars.fatigue.TransparentDraw picBuffer, X, Y + 20, 1, , , , 240 - lesser((plr.fatigue / plr.fatiguemax) * 240, 240)


End Function

Function loadstatbars()

makesprite bars.life, Form1.Picture1, "statusbar.bmp", 250, 0, 0
makesprite bars.life2, Form1.Picture1, "statusbar.bmp", 20, 250, 0
makesprite bars.mana, Form1.Picture1, "statusbar.bmp", 10, 50, 250
makesprite bars.fatigue, Form1.Picture1, "statusbar.bmp", 250, 150, 0
makesprite bars.empty, Form1.Picture1, "statusbar.bmp", 50, 50, 50

End Function

Function setplrkeys()

plr.keys.moveN = vbKeyNumpad9
plr.keys.moveNE = vbKeyNumpad6
plr.keys.moveE = vbKeyNumpad3
plr.keys.moveSE = vbKeyNumpad2
plr.keys.moveS = vbKeyNumpad1
plr.keys.moveSW = vbKeyNumpad4
plr.keys.moveW = vbKeyNumpad7
plr.keys.moveNW = vbKeyNumpad8
plr.keys.eat = vbKeyE
plr.keys.lifepot = vbKeyDecimal
plr.keys.manapot = vbKeyNumpad0

End Function

Function addaura(filen, r, g, b, cells, obj As objecttype, Optional auratype As String = "", Optional duration = 50)

 '   aroll = roll(6)
 '   Dim runk As objecttype
 '   runk.effect(1, 1) = getgener("BONSTR", "BONDEX", "BONINT")
 '   runk.effect(1, 2) = roll(3) + 2
'    If aroll = 6 Then addaura "swordaura.bmp", roll(255), roll(255), roll(255), 15, runk
 '   If aroll = 1 Then addaura "aura6.bmp", roll(255), roll(255), roll(255), 7, runk
 '   If aroll = 2 Then addaura "aura1.bmp", roll(255), roll(255), roll(255), 3, runk
'    If aroll = 3 Then addaura "aura5.bmp", roll(255), roll(255), roll(255), 6, runk
'    If aroll = 4 Then addaura "aura2.bmp", roll(255), roll(255), roll(255), 4, runk
'    If aroll = 5 Then addaura "aura3.bmp", roll(255), roll(255), roll(255), 4, runk

updatbonuses = 1

For a = 1 To UBound(auras())
    If auras(a).type = auratype Then auras(a).loaded = 0
Next a

For a = 1 To UBound(auras())
    If auras(a).loaded = 0 Or a = UBound(auras()) Then
        makesprite auras(a).graphs, Form1.Picture1, filen, r, g, b, , cells
        auras(a).loaded = 1
        auras(a).obj = obj
        auras(a).graphs.recolor r, g, b, , 1
        auras(a).duration = duration
        auras(a).type = auratype
        Exit Function
    End If
Next a

End Function

Function drawauras(Optional addcell As Byte = 0)



For a = 1 To UBound(auras())
    If auras(a).loaded = 1 Then
        If addcell = 1 Then auras(a).cell = auras(a).cell + 1
        If auras(a).cell > auras(a).graphs.cells Then auras(a).cell = 1
        drawobj auras(a).graphs, plr.X, plr.Y, auras(a).cell, , (plr.xoff), (plr.yoff) + -((a - 1) * 3)
        If turncount Mod 10 = 0 Then auras(a).duration = auras(a).duration - 1: If auras(a).duration < 1 Then auras(a).loaded = 0: updatbonuses = 1
        End If
Next a

End Function

Function clearauras()

For a = 1 To UBound(auras())
    auras(a).loaded = 0
Next a

End Function

Function churnoffsets()

If UBound(xmapoff()) <> UBound(map(), 1) Then ReDim xmapoff(UBound(map(), 1))
If UBound(ymapoff()) <> UBound(map(), 1) Then ReDim ymapoff(UBound(map(), 1))

For a = 1 To UBound(xmapoff())
    xmapoff(a) = roll(3) - 2
Next a

End Function

Function randenzymetype(level, Optional typename As String = "", Optional ByVal tt = 0) As monstertype

If typename = "" Then typename = getgener("Digestive", "Stomach", "Gastric")  'getmontype(tt)
Dim mon As monstertype

'getmoncolorbyname typename, r, g, b, levelmod
mon.color = RGB(roll(100) + 150, roll(150) + 100, 0) 'RGB(r, g, b)
mon.light = 0.5
mon.level = Int(level * greater(levelmod, 1))

aroll = roll(11)
    Select Case aroll
        Case 1: mon.name = "Slime": mon.gfile = "slime1.bmp"
        Case 2: mon.name = "Goo": mon.gfile = "slime1.bmp"
        Case 3: mon.name = "Enzyme": mon.gfile = "slime1.bmp"
        Case 4: mon.name = "Juice": mon.gfile = "slime1.bmp"
        Case 5: mon.name = "Acid": mon.gfile = "slime1.bmp"
        Case 6: mon.name = "Liquid": mon.gfile = "slime1.bmp"
        Case 7: mon.name = "Bile": mon.gfile = "slime1.bmp"
        Case 8: mon.name = "Secretion": mon.gfile = "slime1.bmp"
        Case 9: mon.name = "Sludge": mon.gfile = "enzyme2.bmp"
        Case 10: mon.name = "Chunks": mon.gfile = "enzyme2.bmp"
        Case 11: mon.name = "Juices": mon.gfile = "enzyme2.bmp"
        
        Case 17: mon.name = "Centauress": mon.gfile = "centauress1.bmp"
        Case 18: mon.name = "Mare": mon.gfile = "centauress1.bmp"
        Case 19: mon.name = "Harpy": mon.gfile = "harpy1.bmp"
        Case 20: mon.name = "Raven": mon.gfile = "harpy1.bmp"
        Case 21: mon.name = "Sprite": mon.gfile = "sprite1.bmp"
        Case 22: mon.name = "Faerie": mon.gfile = "sprite1.bmp"
        Case 23: mon.name = "Succubus": mon.gfile = "succubus1.bmp"
        Case 24: mon.name = "Batwoman": mon.gfile = "succubus1.bmp"
        Case 25: mon.name = "Dragon": mon.gfile = "thirsha1.bmp"
        Case 27: mon.name = "Winged Serpent": mon.gfile = "wingedserpent1.bmp"
        Case 28: mon.name = "Snakebat": mon.gfile = "wingedserpent1.bmp"
        Case 29: mon.name = "Alligator": mon.gfile = "croc2.bmp"
        Case 30: mon.name = "Crocodile": mon.gfile = "croc2.bmp"
        Case 31: mon.name = "Spiderwoman": mon.gfile = "kari1.bmp"
        Case 32: mon.name = "Grub": mon.gfile = "worm1.bmp"
        Case 33: mon.name = "Elemental": mon.gfile = "elemental1.bmp": mon.colorwhole = 3
        Case 34: mon.name = "Demonette": mon.gfile = "demoness2.bmp"
        Case 35: mon.name = "Demongirl": mon.gfile = "demoness2.bmp"
        Case 36: mon.name = "Demoness": mon.gfile = "demoness1.bmp"
        Case 37: mon.name = "Flower": mon.gfile = "flower1.bmp"
        Case 38: mon.name = "Spitweed": mon.gfile = "flower1.bmp"
        Case 38: mon.name = "Xebbebba": mon.gfile = "xebebba.bmp"
    End Select

'mon.gfile = "xebebba.bmp"

mon.name = typename & " " & mon.name

randenzymetype = mon

End Function

Function makeminimap()

Form1.Picture1.Cls
Form1.Picture1.Width = mapx + 1
Form1.Picture1.Height = mapy + 1
Set minimapspr2 = New cSpriteBitmaps
minimapspr2.CreateFromPicture Form1.Picture1, 1, 1, , RGB(255, 255, 0)

minimapspr2.LockMe
tilespr.LockMe

For a = 1 To mapx
    For b = 1 To mapy
        tilet = map(a, b).tile - 1
        X = (tilet Mod 5) * 96 + 48
        Y = Int(tilet / 5) * 52 + 27
        col = tilespr.DXS.GetLockedPixel(X, Y)
        If map(a, b).blocked > 0 Then col = RGB(100, 100, 100)
        'If map(a, b).object > 0 Then col = RGB(objtypes(objs(map(a, b).object).type).b, objtypes(objs(map(a, b).object).type).g, objtypes(objs(map(a, b).object).type).r)
        'If a = plr.x Or b = plr.y Then col = RGB(250, 250, 10)
        minimapspr2.DXS.SetLockedPixel a, b, col
    Next b
Next a

tilespr.UnlockMe
minimapspr2.UnlockMe

updateminimap

End Function

Function updateminimap()


'Form1.Picture1.Cls
'Form1.Picture1.Width = mapx
'Form1.Picture1.Height = mapy
Set minimapspr = New cSpriteBitmaps
'minimapspr.CreateFromPicture Form1.Picture1, 1, 1, , RGB(255, 255, 0)
If minimapspr2.CellWidth <> mapx + 1 Then makeminimap: Exit Function 'Exit, because makeminimap will call a new instance of updateminimap
If minimapspr2.CellHeight <> mapy + 1 Then makeminimap: Exit Function
minimapspr.CreateFromSurface minimapspr2.DXS, 1, 1


minimapspr.LockMe
'tilespr.LockMe

For a = 1 To mapx
    For b = 1 To mapy
        drawex = 0 'Draws a big X for plot stuff
        drawyes = 0
        tilet = map(a, b).tile - 1
        X = (tilet Mod 5) * 96 + 48
        Y = Int(tilet / 5) * 52 + 27
        'col = tilespr.DXS.GetLockedPixel(x, y)
        'If map(a, b).blocked > 0 Then col = RGB(100, 100, 100)
        If map(a, b).object > 0 Then col = RGB(objtypes(objs(map(a, b).object).type).b, objtypes(objs(map(a, b).object).type).g, objtypes(objs(map(a, b).object).type).r): drawyes = 1
        If map(a, b).object > 0 Then If Not geteff(objtypes(objs(map(a, b).object).type), "NoEat", 1) = "" Then drawex = 1
        
        If drawyes = 1 And drawex = 1 Then
            For c = -3 To 3
                For d = 1 To 3
                drx = a + c + d: drx2 = a + -c + d
                dry = b + c
                If drx < mapx And drx > 0 And dry < mapy - 2 And dry > 2 Then minimapspr.DXS.SetLockedPixel drx, dry, col
                If drx2 < mapx And drx2 > 0 And dry < mapy - 2 And dry > 2 Then minimapspr.DXS.SetLockedPixel drx2, dry, col
                Next d
            Next c
            drawex = 0
        End If
                
        If a = plr.X Or b = plr.Y Then col = RGB(250, 250, 10): drawyes = 1
        If map(a, b).monster > 0 Then
        'm'' line below let avoid drawing friendly monster on minimap, but it makes flicker them on main map
        'm'' If mon(map(a, b).monster).owner > 0 Then map(a, b).monster = 0: GoTo 72 'm''
        If mon(map(a, b).monster).type = 0 Then GoTo 72
        col = RGB(0, 0, 250): drawyes = 1
        If mon(map(a, b).monster).owner > 0 Then col = RGB(0, 250, 0)
        If mapjunk.questmonster = montype(mon(map(a, b).monster).type).name Then drawex = 1
72         End If
        
        If drawyes = 1 Then minimapspr.DXS.SetLockedPixel a, b, col
        
        'Draw objectives as a big "X"
        If drawyes = 1 And drawex = 1 Then
            For c = -3 To 3
                For d = 1 To 3
                drx = a + c + d: drx2 = a + -c + d
                dry = b + c
                If drx < mapx And drx > 0 And dry < mapy - 2 And dry > 2 Then minimapspr.DXS.SetLockedPixel drx, dry, col
                If drx2 < mapx And drx2 > 0 And dry < mapy - 2 And dry > 2 Then minimapspr.DXS.SetLockedPixel drx2, dry, col
                Next d
            Next c
            drawex = 0
        End If
        
    Next b
Next a

'tilespr.UnlockMe
minimapspr.UnlockMe

End Function


Function loadclothestypes(filen)
Static djerk
If djerk = 1 Then Exit Function
djerk = 1
filen = getfile(filen, "Data.pak")

fnum = FreeFile
Open filen For Input As fnum

Do While Not EOF(fnum)
    
    Input #fnum, snarg
    
    If snarg = "#CLOTHES" Then
        zim = zim + 1
        ReDim Preserve clothestypes(1 To zim)
        Input #fnum, clothestypes(zim).name, clothestypes(zim).graph, clothestypes(zim).armor, clothestypes(zim).weight, clothestypes(zim).wear1, clothestypes(zim).wear2, clothestypes(zim).material
    End If
    
    If snarg = "#EFFECT" Then
        zorg = zorg + 1
        Input #fnum, materialtypes(zorg).effects(zorg, 1), materialtypes(zorg).effects(zorg, 2), materialtypes(zorg).effects(zorg, 3), materialtypes(zorg).effects(zorg, 4), materialtypes(zorg).effects(zorg, 5)
    End If
    
    If snarg = "#TRANSLUCENT" Then
        materialtypes(zor).effects(3, 1) = "Translucent": materialtypes(zor).effects(3, 2) = 1
    End If
    
    If snarg = "#MATERIAL" Then
        zor = zor + 1
        ReDim Preserve materialtypes(1 To zor)
        Input #fnum, materialtypes(zor).type, materialtypes(zor).name, materialtypes(zor).r, materialtypes(zor).g, materialtypes(zor).b, materialtypes(zor).armor, materialtypes(zor).weight, materialtypes(zor).worth, materialtypes(zor).goldmult
        zorg = 0
    End If
    
    If snarg = "#WEAPON" Then
        zung = zung + 1
        ReDim Preserve weapontypes(1 To zung)
        Input #fnum, weapontypes(zung).name, weapontypes(zung).graph, weapontypes(zung).weight, weapontypes(zung).dice, weapontypes(zung).damage, weapontypes(zung).type, weapontypes(zung).material
    End If

Loop

Close fnum

End Function

Function filegrab(ByVal filen, ByVal StringID, Optional ByVal datafile = "VRPGData.dat", Optional quitstring = "#KUBLARGY", Optional ByRef val1, Optional ByRef val2 _
, Optional ByRef val3, Optional ByRef val4 _
, Optional ByRef val5, Optional ByRef val6, Optional ByRef val7)
'Universal text datafile reader

'-2: Quitstring encountered
'-1: File not found
'0: End if file
'1: Continue (Data found)

Static filenum As Integer

If Dir(filen) = "" Then filen = getfile(filen, datafile)
If Dir(filen) = "" Then gamemsg ("File not found:" & filen): filegrab = -1

If filenum = 0 Then filenum = FreeFile: Open filen For Input As filenum
Do While Not gur = StringID
If EOF(filenum) = True Then GoTo 5
Input #filenum, gur
If quitstring = gur Then filegrab = -2: Close filenum: filenum = 0: Exit Function
Loop

If gur = StringID Then
If Not IsMissing(val7) Then Input #filenum, val1, val2, val3, val4, val5, val6, val7: GoTo 5
If Not IsMissing(val6) Then Input #filenum, val1, val2, val3, val4, val5, val6: GoTo 5
If Not IsMissing(val5) Then Input #filenum, val1, val2, val3, val4, val5: GoTo 5
If Not IsMissing(val4) Then Input #filenum, val1, val2, val3, val4: GoTo 5
If Not IsMissing(val3) Then Input #filenum, val1, val2, val3: GoTo 5
If Not IsMissing(val2) Then Input #filenum, val1, val2: GoTo 5
If Not IsMissing(val1) Then Input #filenum, val1: GoTo 5
End If
5
filegrab = 1
If EOF(filenum) = True Then filegrab = 0: Close filenum: filenum = 0: Exit Function

End Function


Function filestr(Optional strcom = "#$STRNAME", Optional strname = "$STRNAME", Optional convfile = "Spaceconvs.txt") As String

'If Strname is omitted, it reads every string until it hits a blank line.
'If Strcom is ommitted, it will assume it is the same as the strname, only with a # at the beginning.

Dim gstr As String
Dim gstr2 As String
If strcom = "#$STRNAME" Then strcom = "#" & strname
If strname = "$STRNAME" Then strname = ""

filenum = FreeFile

Open convfile For Input As #filenum
Dim newzes(1 To 50) As String
Do While Not EOF(filenum)
    Input #filenum, gstr
    
    'Switch news adding on or off
    'If gstr = strcom Then Input #1, gstr2
    If gstr = strcom Then
    'newson = 1 Else newson = 0
        newsnum = 1
        Do While (1)
        Input #1, gstr
        If gstr = "" Then Exit Do
        If Not strname = "" Then If gstr = strname Then Input #filenum, newzes(newsnum): newsnum = newsnum + 1 Else Exit Do
        If strname = "" Then newzes(newsnum) = gstr: newsnum = newsnum + 1
        Loop

    End If
    
Loop

3       aroll = roll(newsnum)
        If newzes(aroll) = "" Then GoTo 3
        newsstr = newsstr & newzes(aroll)

filestr = newsstr

Close #filenum

End Function

Function savegame(filen)

savebindata Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"
fsavechar "plrdat.tmp" 'FileD.FileTitle

addfile "plrdat.tmp", filen, 1
addfile "curgame.dat", filen, 1

'Kill "plrdat.tmp"
'If Not Dir(Left(plr.curmap, Len(plr.curmap) - 4) & ".dat") = "" Then Kill Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"

End Function

Function loadallmongraphs()

lastmontype = UBound(montype())

ReDim Preserve montype(1 To lastmontype): ReDim Preserve mongraphs(1 To UBound(montype())) As cSpriteBitmaps

montype(lastmontype).light = light
calcexp montype(lastmontype)
            
For a = 1 To lastmontype
getrgb montype(a).color, r, g, b
makesprite mongraphs(a), Form1.Picture1, montype(a).gfile, r, g, b, 0.5, 3
Next a

End Function

Function checkaccess2(Optional ByVal startx = 0, Optional ByVal starty = 0)

If startx = 0 And starty = 0 Then startx = Int(mapx / 2): starty = Int(mapy / 2) 'startx = plr.x: starty = plr.y

map(startx, starty).ovrtile = 0: map(startx, starty).blocked = 0

maxdist = 2

zarg = 0
origx = startx
origy = starty

5 ReDim dmap(1 To mapx, 1 To mapy)

map(startx, starty).blocked = 0
map(startx, starty).ovrtile = 0

'makeminimap
'updateminimap
'drawall

'With Form1.Picture1
'    .Visible = True
'    .Width = mapx
'    .Height = mapy
'
'For a = 1 To mapx
'    For b = 1 To mapy
'        If map(a, b).blocked = 1 Then Form1.Picture1.PSet (a, b), RGB(150, 150, 150)
'    Next b
'Next a
'End With


subdmap startx, starty

zarg = zarg + 1

Select Case zarg
    Case 2: startx = Int(mapx * 0.25): starty = Int(mapy * 0.25)
    Case 3: startx = Int(mapx * 0.75): starty = Int(mapy * 0.25)
    Case 4: startx = Int(mapx * 0.75): starty = Int(mapy * 0.75)
    Case 5: startx = Int(mapx * 0.25): starty = Int(mapy * 0.75)
    Case 6: startx = Int(mapx * 0.5): starty = Int(mapy * 0.5)
    Case 7: startx = origx: starty = origy ': zarg = 0
    Case 8: Exit Function
End Select

'With Form1.Picture1
'
'For a = 1 To mapx
'    For b = 1 To mapy
'        If dmap(a, b) = 1 Then Form1.Picture1.PSet (a, b), RGB(250, 0, 250)
'    Next b
'Next a
'
'.Refresh
'
'End With



'For a = 1 To mapx
'For b = 1 To mapy
'    If dmap(a, b) > 0 Then map(a, b).tile = 1 ': zork = zork + 1
'Next b
'Next a

'Debug.Print zork

'Exit Function

35 For a = 1 To mapx
For b = 1 To mapy
    'On Error Resume Next
    'Ignore walls and 'Non' tiles
    If map(a, b).tile = 0 Or map(a, b).blocked > 0 Then GoTo 7
    'If dmap(a, b) > 0 Then map(a, b).tile = 1
    tunnel = 0
    If dmap(a, b) = 0 Then
    
    'Tunnel didn't work--attempting Jump algorithm.  It checks to see if it can hop over any walls to become accessible.
    'maxdist = 2
61  If a + maxdist > mapx Then GoTo 62
    If dmap(a + maxdist, b) = 1 Then
        For c = a To a + maxdist: map(c, b).ovrtile = 0: map(c, b).blocked = 0: dmap(c, b) = 1: Next c: tunnel = 1: GoTo 5
    End If
    
62  If a - maxdist < 1 Then GoTo 63
    If dmap(a - maxdist, b) = 1 Then
        For c = a To a - maxdist Step -1: map(c, b).ovrtile = 0: map(c, b).blocked = 0: dmap(c, b) = 1: Next c: tunnel = 1: GoTo 5
    End If
    
63  If b - maxdist < 1 Then GoTo 64
    If dmap(a, b - maxdist) = 1 Then
        For c = b To b - maxdist Step -1: map(a, c).ovrtile = 0: map(a, c).blocked = 0: dmap(a, c) = 1: Next c: tunnel = 1: GoTo 5
    End If

64  If b + maxdist > mapy Then GoTo 65
    If dmap(a, b + maxdist) = 1 Then
        For c = b To b + maxdist: map(a, c).ovrtile = 0: map(a, c).blocked = 0: dmap(a, c) = 1: Next c: tunnel = 1: GoTo 5
    End If
65
    
    
    'If a tile is inaccessible, 'Tunnel' towards the starting point until you hit an accessible tile
'    tunnel = tunnel + 1
'    zx = a
'    zy = b
'12  If zx = startx And zy = starty Then GoTo 9
'    If dmap(zx, zy) = 1 Then GoTo 9
'    If zx > startx Then zx = zx - 1: GoTo 15
'    If zx < startx Then zx = zx + 1: GoTo 15
'    If zy > starty Then zy = zy - 1: GoTo 15
'    If zy < starty Then zy = zy + 1: GoTo 15
'15  map(zx, zy).ovrtile = 0: map(zx, zy).blocked = 0: dmap(zx, zy) = 1
'    GoTo 12
    
   End If
    'After making the area accessible, restart
   If tunnel = 1 Then Stop: GoTo 5
    'End If
7 Next b
Next a



If maxdist > 10 Then Exit Function

For a = 1 To mapx
For b = 1 To mapy
    If dmap(a, b) = 0 And map(a, b).ovrtile = 0 And map(a, b).tile > 0 Then maxdist = maxdist + 1: GoTo 5
Next b
Next a



End Function

Function subdmap2(ByVal X, ByVal Y)

For a = 1 To mapx
    For b = 1 To mapy
        dmap(X, Y) = 1
    Next b
Next a

End Function


Function subdmap3(ByVal X, ByVal Y)
'Uses a fill algorithm of sorts
'probably isn't fast, but will work.

5 stillfill = 0
For a = 1 To mapx
    For b = 1 To mapy
        If dmap(a, b) = 2 Then
            If a <= mapx Then If dmap(a + 1, b) = 0 Then dmap(a + 1, b) = 3
            If a > 1 Then If dmap(a - 1, b) = 0 Then dmap(a - 1, b) = 3
            If b <= mapy Then If dmap(a, b + 1) = 0 Then dmap(a, b + 1) = 3
            If b > 1 Then If dmap(a, b - 1) = 0 Then dmap(a, b - 1) = 3
            dmap(a, b) = 1
            stillfill = 1
        End If
    Next b
Next a

If stillfill = 1 Then
For a = 1 To mapx
    For b = 1 To mapy
        If dmap(a, b) = 3 Then dmap(a, b) = 2
    Next b
Next a
GoTo 5
End If


End Function

Function subdmap(ByVal X, ByVal Y)

Static iterations

'Exit Function
iterations = iterations + 1
'If iterations > 15 Then Stop
If iterations > 1000 Then iterations = iterations - 1:  Exit Function

If map(X, Y).tile = 0 Then iterations = iterations - 1: Exit Function

'map(x, y).tile = 1
If dmap(X, Y) = 1 Then iterations = iterations - 1: Exit Function



'On Error GoTo 5
'GoTo 10
'5 Stop: End
'10

If map(X, Y).blocked > 0 Then subdmap = -1: iterations = iterations - 1: Exit Function ':Stop
dmap(X, Y) = 1: subdmap = 1


GoTo 7

'x2 = X: y2 = Y
'
''Continue in all straight lines until a wall is hit
'Do While Not map(x2, y2).blocked > 0
'If x2 > mapx Then Exit Do
'If x2 < 1 Then Exit Do
'If y2 < 1 Then Exit Do
'If y2 > mapy Then Exit Do
'dmap(x2, y2) = 1
'x2 = x2 + 1
'Loop
''When a wall IS hit, subd again on the previous tile (Wall reflect, basically)
'x2 = x2 - 1:  dmap(x2, y2) = 0: subdmap x2, y2
'
'x2 = X: y2 = Y
'Do While Not map(x2, y2).blocked > 0
'If x2 > mapx Then Exit Do
'If x2 < 1 Then Exit Do
'If y2 < 1 Then Exit Do
'If y2 > mapy Then Exit Do
'dmap(x2, y2) = 1
'x2 = x2 - 1
'Loop
'x2 = x2 + 1: dmap(x2, y2) = 0: subdmap x2, y2
'
'
'x2 = X: y2 = Y
'Do While Not map(x2, y2).blocked > 0
'If x2 > mapx Then Exit Do
'If x2 < 1 Then Exit Do
'If y2 < 1 Then Exit Do
'If y2 > mapy Then Exit Do
'dmap(x2, y2) = 1
'y2 = y2 + 1
'Loop
'y2 = y2 - 1: dmap(x2, y2) = 0: subdmap x2, y2
'
'
'
'x2 = X: y2 = Y
'Do While Not map(x2, y2).blocked > 0
'If x2 > mapx Then Exit Do
'If x2 < 1 Then Exit Do
'If y2 < 1 Then Exit Do
'If y2 > mapy Then Exit Do
'dmap(x2, y2) = 1
'y2 = y2 - 1
'Loop
'y2 = y2 + 1:  dmap(x2, y2) = 0: subdmap x2, y2
'
'
''Then, Subd all diagonals
''If map(x + 1, y + 1).blocked = 0 And dmap(x + 1, y + 1) = 0 Then subdmap x + 1, y + 1
''If map(x - 1, y + 1).blocked = 0 And dmap(x - 1, y + 1) = 0 Then subdmap x - 1, y + 1
''If map(x + 1, y - 1).blocked = 0 And dmap(x + 1, y - 1) = 0 Then subdmap x + 1, y - 1
''If map(x - 1, y - 1).blocked = 0 And dmap(x - 1, y - 1) = 0 Then subdmap x - 1, y - 1

7


If X + 1 < mapx Then If map(X + 1, Y).blocked = 0 And dmap(X + 1, Y) = 0 Then subdmap X + 1, Y
If X - 1 > 1 Then If map(X - 1, Y).blocked = 0 And dmap(X - 1, Y) = 0 Then subdmap X - 1, Y
If Y - 1 > 1 Then If map(X, Y - 1).blocked = 0 And dmap(X, Y - 1) = 0 Then subdmap X, Y - 1
If Y + 1 < mapy Then If map(X, Y + 1).blocked = 0 And dmap(X, Y + 1) = 0 Then subdmap X, Y + 1


iterations = iterations - 1

End Function

Function walloffmap(Optional ovrtile = 0)

If map(1, 1).tile = 0 Then Exit Function

If ovrtile = 0 Then  'Grab a random tile from map
    Do While ovrtile = 0
    tries = tries + 1
    ovrtile = map(roll(mapx), roll(mapy)).ovrtile
    If tries > 1000 Then ovrtile = 1: Exit Do
    Loop
End If

If mapjunk.maps(1) = "" Then For a = 1 To mapx: map(a, 1).blocked = 1: map(a, 1).ovrtile = ovrtile: Next a
If mapjunk.maps(2) = "" Then For a = 1 To mapy: map(mapx, a).blocked = 1: map(mapx, a).ovrtile = ovrtile: Next a
If mapjunk.maps(3) = "" Then For a = 1 To mapx: map(mapy, 1).blocked = 1: map(mapy, 1).ovrtile = ovrtile: Next a
If mapjunk.maps(4) = "" Then For a = 1 To mapy: map(1, a).blocked = 1: map(1, a).ovrtile = ovrtile: Next a

End Function

Function getcarrymonsters()

If monsterstoput = 1 Then Exit Function

'Clear Carrymonsters
For b = 1 To 10
carrymonsters(b).montype.name = ""
carrymonsters(b).numeach = 0
Next b

5 For a = 1 To UBound(mon())
    If mon(a).owner > 0 And mon(a).type > 0 And mon(a).hp > 0 And mon(a).X > 0 And mon(a).instomach = 0 Then 'm'' added handling of in-stomach monsters
        For b = 1 To 10
        If montype(mon(a).type).name = carrymonsters(b).montype.name Then carrymonsters(b).numeach = carrymonsters(b).numeach + 1: killmon a, 1: Exit For
        If carrymonsters(b).montype.name = "" Then carrymonsters(b).montype = montype(mon(a).type): carrymonsters(b).numeach = carrymonsters(b).numeach + 1: killmon a, 1: Exit For
        Next b
    End If
Next a

monsterstoput = 1

End Function

Function putcarrymonsters()

'If monsterstoput = 0 Then Exit Function


For b = 1 To 10
    If carrymonsters(b).numeach < 1 Then Exit For
    getrgb carrymonsters(b).montype.color, r, g, bl
    mtnum = createmontype(carrymonsters(b).montype.name, carrymonsters(b).montype.gfile, carrymonsters(b).montype.level, r, g, bl, carrymonsters(b).montype.light)
    For c = 1 To carrymonsters(b).numeach
    mnum = createmonster(mtnum, plr.X, plr.Y)
    mon(mnum).owner = 1
    Next c
    
Next b

monsterstoput = 0

End Function

Function plrescape(msg)
'Escape from stomach
    
    gamemsg msg
    stopsounds
    playsound "swallow3.wav"
    playsound "burp" & roll(5) & ".wav"
    If plr.diglevel < 4 Then playsound "grunt" & roll(9) + 2 & ".wav"
    mon(plr.instomach).cell = 1: plr.instomach = 0: swallowcounter = -6: If plr.hp < 1 Then plr.hp = 1: plr.plrdead = 0

stomachlevel = 0
End Function

Function tryescape(ByVal adddifficulty)
'Returns successes

If plr.Swallowtime > 200 Then adddifficulty = adddifficulty + 1
If plr.Swallowtime > 500 Then adddifficulty = adddifficulty + 1


If montype(mon(plr.instomach).type).eattype = 1 Then adddifficulty = -1 'Engulfing beasts are always the same difficulty to escape from
'If stomachlevel > 0 Then adddifficulty = adddifficulty + 1
If stomachlevel > 2 And plr.plrdead = 0 Then struggled = 1

If plr.diglevel < 4 And plr.fatigue < plr.fatiguemax * 0.8 Then mon(plr.instomach).xoff = 12 * (roll(3) - 2): mon(plr.instomach).yoff = 12 * (roll(3) - 2) 'plr.yoff = 12 * (roll(3) - 2): plr.xoff = 12 * (roll(3) - 2)
'No more than one attempt per turn
If lastmove = turncount Then Exit Function Else lastmove = turncount + 1
If plr.diglevel > 3 Then GoTo 3
If plr.fatigue > plr.fatiguemax * 0.8 Then gamemsg "You are too exhausted to struggle.": GoTo 3
If roll(2) = 1 Then GoTo 3 'Any attempt will fail 50% of the time

'If KeyCode = vbKeyNumpad5 Or roll(2) = 1 Or plr.plrdead = 1 Then GoTo 3

'Generic opposed strength/level roll
'More HP increases your skill, being deep in the stomach reduces it

If montype(mon(plr.instomach).type).eattype = 0 Then addmskill = stomachlevel

'Adddifficulty adds to the succroll target number (So anything higher than 4 makes it nearly impossible--any negative number makes it much easier)
If succroll(getstr + (plr.hp / gethpmax) * 5, 7 - lesser(plr.diglevel, 1) + adddifficulty) > succroll(montype(mon(plr.instomach).type).escapediff + addmskill, 5) Then escaped = 1

'If (succroll((plr.str / 3 + 2) + (plr.level / 4) + (((plr.hp / plr.hpmax) * 10) - 4) + (Int(instomachcounter / 4) - 4), 6 - lesser(plr.diglevel, 1) + adddifficulty) + roll(skilltotal("Squirm", 2, 1)) > succroll(montype(mon(plr.instomach).type).escapediff * 2)) Then escaped = 1

'If you have high HP, you have a chance
'If (roll(20 - (((plr.hp / plr.hpmax) * 10) - 4)) = 1) Then escaped = 1

'If you have no fatigue and no difficulty adds you escape automatically
If plr.fatigue = 0 And montype(mon(plr.instomach).type).eattype = 1 And adddifficulty <= 0 Then escaped = 1

'If your HP is already low, you have better odds
'If (roll(roll(plr.hpmax / 5)) > plr.hp And plr.hp > 1) Then escaped = 1

3 If escaped = 1 Then
    'if in the throat, escape.  Otherwise, claw up to the throat
    If stomachlevel < 2 Or montype(mon(plr.instomach).type).eattype = 1 Then plrescape getesc Else stomachlevel = 0: addtext "You have miraculously clawed your way back up the " & montype(mon(plr.instomach).type).name & "'s throat!"
End If

'Struggling costs fatigue, whether you succeed or not, but doesn't penalize you as much if it's really hard to do
addfatigue greater(8 - (adddifficulty * 1.5), 0) 'm'' alternate formula from other source code
Exit Function 'm''

'm'' the original formula may go wrong, so i added a little correction
Dim Mlevel As Long 'm''
If plr.instomach = 0 Then Exit Function 'm'' plants dont have stomach ...
dp = InStrRev(montype(mon(plr.instomach).type).level, ":", , vbBinaryCompare) 'm''
If dp > 0 Then 'm''
    Mlevel = Val(Mid(montype(mon(plr.instomach).type).level, dp + 1)) 'm''
Else 'm''
    Mlevel = Val(montype(mon(plr.instomach).type).level) 'm''
End If 'm''

If Val(Mlevel) = 0 Then Stop 'm''
'm''If Not escaped = 1 Then addfatigue greater(montype(mon(plr.instomach).type).level \ 2 - (adddifficulty * 1.5), 1) Else addfatigue greater(plr.level - (adddifficulty * 1.5), 1)

If Not escaped = 1 Then addfatigue greater(Mlevel \ 2 - (adddifficulty * 1.5), 1) Else addfatigue greater(plr.level - (adddifficulty * 1.5), 1)

End Function

Function addquest(quest As String)

'Standard quest is QuestName(make sure it's unique):Map:Description:Givenby:Experience:Gold:Type("KILL", "RETRIEVE", "COMPLETED" etc):Extra stuff after this
'KILL:Monstertype:Num:Color:Level:R:G:B (Decremented each time you kill one--when it hits 0, the quest is achieved)
'RETRIEVE:Item:Giveto:Lose Y/N (Have the item in your inventory when you speak to Giveto.  Lose determines if it takes it from your inventory)
'COMPLETED (Quest is completed)
'If you have accomplished a quest but it isn't marked 'Completed', the conversations with any character named 'Givenby'
'switch to the conversation named '(givenby)(questname)'
'So if a girl named Rasella gives a quest called 'DestroyKaris' it will go to the conversation 'RasellaDestroyKaris'

'It will not add quests with non-unique names
qname = getfromstring(quest, 1)
For a = 1 To UBound(plr.curquests())
    If getfromstring(plr.curquests(a), 1) = qname Then Exit Function
Next a

ReDim Preserve plr.curquests(1 To UBound(plr.curquests()) + 1)
qnum = UBound(plr.curquests())

plr.curquests(qnum) = quest

End Function

Function getquestpending(ByVal charname) As String
'Returns the conversation thingy name if qualifications for quest have been met but quest not marked completed

For a = 1 To UBound(plr.curquests())
    If Not LCase(getfromstring(plr.curquests(a), 4)) = LCase(charname) Then GoTo 3
    If getfromstring(plr.curquests(a), 7) = "COMPLETED" Then GoTo 3
    
    If getfromstring(plr.curquests(a), 7) = "RETRIEVE" Then
        b = getfromstring(plr.curquests(a), 13)
        g = getfromstring(plr.curquests(a), 12)
        r = getfromstring(plr.curquests(a), 11)
        itype = getfromstring(plr.curquests(a), 8)
        For c = 1 To 50
            If inv(c).name = itype Then
                completequest getfromstring(plr.curquests(a), 1)
                getquestpending = charname & getfromstring(plr.curquests(a), 1)
                If getfromstring(plr.curquests(a), 10) = "Y" Then killitem inv(c): orginv
                Exit Function
            End If
        Next c
        GoTo 3
    End If
    
    If getfromstring(plr.curquests(a), 7) = "KILL" Then
        b = getfromstring(plr.curquests(a), 14)
        g = getfromstring(plr.curquests(a), 13)
        r = getfromstring(plr.curquests(a), 12)
        graph = getfromstring(plr.curquests(a), 11)
        level = getfromstring(plr.curquests(a), 10)
        num = getfromstring(plr.curquests(a), 9)
        mtype = getfromstring(plr.curquests(a), 8)
        If num = 0 Then completequest getfromstring(plr.curquests(a), 1): getquestpending = charname & getfromstring(plr.curquests(a), 1): Exit Function
        GoTo 3
    End If

3 Next a


End Function

Function getquestnum(ByVal questname As String)

For a = 1 To UBound(plr.curquests())
    If getfromstring(plr.curquests(a), 1) = questname Then qnum = a: Exit For
Next a

getquestnum = qnum

End Function

Function completequest(ByVal questname)
qnum = getquestnum(questname)

swaptxt plr.curquests(qnum), getfromstring(plr.curquests(qnum), 7), "COMPLETED"
plr.gp = plr.gp + getfromstring(plr.curquests(qnum), 6)
plr.exp = plr.exp + getfromstring(plr.curquests(qnum), 5)

End Function

Function setupquests()
'Standard quest is QuestName(make sure it's unique):Map:Description:Givenby:Experience:Gold:Type("KILL", "RETRIEVE", "COMPLETED" etc):Extra stuff after this
'KILL:Monstertype:Num:Color:Level:R:G:B (Decremented each time you kill one--when it hits 0, the quest is achieved)
'RETRIEVE:Item:Giveto:Lose Y/N:R:G:B:Graph (Have the item in your inventory when you speak to Giveto.  Lose determines if it takes it from your inventory)
'COMPLETED (Quest is completed)
'If you have accomplished a quest but it isn't marked 'Completed', the conversations with any character named 'Givenby'
'switch to the conversation named '(givenby)(questname)'
'So if a girl named Rasella gives a quest called 'DestroyKaris' it will go to the conversation 'RasellaDestroyKaris'
mapname = plr.curmap
If Left(mapname, 6) = "VTDATA" Then mapname = Right(mapname, Len(mapname) - 6)
For a = 1 To UBound(plr.curquests())
    If Not LCase(getfromstring(plr.curquests(a), 2)) = LCase(mapname) Then GoTo 3
    If getfromstring(plr.curquests(a), 7) = "COMPLETED" Then GoTo 3
    
    If getfromstring(plr.curquests(a), 7) = "RETRIEVE" Then
        b = getfromstring(plr.curquests(a), 13)
        g = getfromstring(plr.curquests(a), 12)
        r = getfromstring(plr.curquests(a), 11)
        itype = getfromstring(plr.curquests(a), 8)
        graph = getfromstring(plr.curquests(a), 14)
        num = countobjtype((itype))
        If num = 0 Then
        objtn = createobjtype(itype, (graph), r, g, b)
        createobj itype, roll(mapx), roll(mapy)
        addeffect objtn, "Pickup"
        addeffect objtn, "NoEat"
        End If
        GoTo 3
    End If
    
    If getfromstring(plr.curquests(a), 7) = "KILL" Then
        b = getfromstring(plr.curquests(a), 14)
        g = getfromstring(plr.curquests(a), 13)
        r = getfromstring(plr.curquests(a), 12)
        graph = getfromstring(plr.curquests(a), 11)
        level = getfromstring(plr.curquests(a), 10)
        num = getfromstring(plr.curquests(a), 9)
        mtype = getfromstring(plr.curquests(a), 8)
        addmons = num - countmonstertype((mtype))
        If countmonstertype((mtype)) = 0 Then mtype = createmontype(mtype, (graph), level, r, g, b, 0.5)
        createmonster mtype, roll(mapx), roll(mapy)
        GoTo 3
    End If

3 Next a

End Function

Function countmonstertype(montype As String)

zork = 0

montypenum = getmontnum(montype)

    For a = 1 To UBound(mon())
        If mon(a).type = montypenum Then zork = zork + 1
    Next a

countmonstertype = zork

End Function

Function countobjtype(objtypen As String)

zork = 0
Dim objtypenum As Integer
objtypenum = getobjtnum(objtypen)
If objtypenum = 0 Then countobjtype = 0: Exit Function
objtypename = objtypes(objtypenum).name


    For a = 1 To UBound(objs())
        If objs(a).type = objtypenum Then zork = zork + 1
    Next a

    For a = 1 To 50
        If inv(a).name = objtypename Then zork = zork + 1
    Next a

countobjtype = zork

End Function

Function textcycle(ByVal curtxt, Optional txt1 As String = "", Optional txt2 As String = "", Optional txt3 As String = "", Optional txt4 As String = "", Optional txt5 As String = "", Optional txt6 As String = "", Optional txt7 As String = "", Optional txt8 As String = "", Optional txt9 As String = "", Optional txt10 As String = "")
'Note:  Cycles in reverse order

If curtxt = txt1 Then
    If Not txt10 = "" Then curtxt = txt10: GoTo 5
    If Not txt9 = "" Then curtxt = txt9: GoTo 5
    If Not txt8 = "" Then curtxt = txt8: GoTo 5
    If Not txt7 = "" Then curtxt = txt7: GoTo 5
    If Not txt6 = "" Then curtxt = txt6: GoTo 5
    If Not txt5 = "" Then curtxt = txt5: GoTo 5
    If Not txt4 = "" Then curtxt = txt4: GoTo 5
    If Not txt3 = "" Then curtxt = txt3: GoTo 5
    If Not txt2 = "" Then curtxt = txt2: GoTo 5
End If

If curtxt = txt10 Then curtxt = txt9: GoTo 5
If curtxt = txt9 Then curtxt = txt8: GoTo 5
If curtxt = txt8 Then curtxt = txt7: GoTo 5
If curtxt = txt7 Then curtxt = txt6: GoTo 5
If curtxt = txt6 Then curtxt = txt5: GoTo 5
If curtxt = txt5 Then curtxt = txt4: GoTo 5
If curtxt = txt4 Then curtxt = txt3: GoTo 5
If curtxt = txt3 Then curtxt = txt2: GoTo 5
If curtxt = txt2 Then curtxt = txt1: GoTo 5

5 textcycle = curtxt

End Function

Function replaceinstr(ByRef txt, ByVal pos, ByVal newval)

'Current limit is 50 total values in a DS, though you can change this just by adding to this array
Dim dunk(1 To 50) As String

'Get all current values and assign them to an array
zork = 1
Do While Not zork > 50
If getfromstring(txt, zork) = "" Then tlast = tlast + 1 Else tlast = 0
dunk(zork) = getfromstring(txt, zork): zork = zork + 1
Loop
zork = zork - 1 - tlast

'Assign the new value to the appropriate position in the array
dunk(pos) = newval

If pos > zork Then Debug.Print "ReplaceDS Error: Position outside DS range.  No replacement will be made."

'Reconstruct the string with the new value
putstr = dunk(1)
For a = 2 To zork
    putstr = putstr & ":" & dunk(a)
Next a

'Finally, assign the new string to the flag DS
txt = putstr

End Function

Function closemon(X, Y, maxdist)

closemon = 0
xdist = maxdist
ydist = maxdist
For a = 1 To UBound(mon())
    If mon(a).type = 0 Then GoTo 5
    If mon(a).X = X Or mon(a).Y = Y Then GoTo 5
    If diff(X, mon(a).X) > xdist Then GoTo 5
    If diff(Y, mon(a).Y) > ydist Then GoTo 5
    xdist = diff(X, mon(a).X)
    ydist = diff(Y, mon(a).Y)
    closemon = a
5 Next a



End Function

Function getgener(txt1 As String, Optional txt2 As String, Optional txt3 As String, Optional txt4 As String, Optional txt5 _
                  As String, Optional txt6 As String, Optional txt7 As String, Optional txt8 As String, Optional txt9 _
                  As String, Optional txt10 As String, Optional txt11 As String, Optional txt12 As String, Optional txt13 _
                  As String, Optional txt14 As String, Optional txt15 As String, Optional txt16 As String, Optional txt17 _
                  As String, Optional txt18 As String, Optional txt19 As String, Optional txt20 As String, Optional txt21 As String) As String

5 arollgen = roll(21)

Select Case arollgen
    Case 1: gstr = txt1
    Case 2: gstr = txt2
    Case 3: gstr = txt3
    Case 4: gstr = txt4
    Case 5: gstr = txt5
    Case 6: gstr = txt6
    Case 7: gstr = txt7
    Case 8: gstr = txt8
    Case 9: gstr = txt9
    Case 10: gstr = txt10
    Case 11: gstr = txt11
    Case 12: gstr = txt12
    Case 13: gstr = txt13
    Case 14: gstr = txt14
    Case 15: gstr = txt15
    Case 16: gstr = txt16
    Case 17: gstr = txt17
    Case 18: gstr = txt18
    Case 19: gstr = txt19
    Case 20: gstr = txt20
    Case 21: gstr = txt21
End Select

If gstr = "" Then GoTo 5
getgener = gstr
End Function

Private Function GetRGBColor(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer)

    Dim ddsurf As DDSURFACEDESC2
    Dim rv As Long
    Dim gv As Long
    Dim bv As Long
    
    DD.GetDisplayMode ddsurf
    With ddsurf.ddpfPixelFormat
        If (.lFlags And DDPF_RGB) > 0 Then
            rv = CLng(CSng(.lRBitMask) * (r / 255)) And .lRBitMask
            gv = CLng(CSng(.lGBitMask) * (g / 255)) And .lGBitMask
            bv = CLng(CSng(.lBBitMask) * (b / 255)) And .lBBitMask
            
            GetRGBColor = rv Or gv Or bv
        End If
    End With

End Function

'Dim iColor As Integer
'  Dim iBlue As Integer
'  Dim iGreen As Integer
'  Dim iRed As Integer
  
'  iColor = &H7FFF ' 111 1111 1111 1111 b
  
'  iBlue = iColor And 31
'  iGreen = (iColor \ 32) And 31
'  iRed = (iColor \ 1024) And 31

Public Function GenRGB(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As Long

    'Dim ddsurf As DDSURFACEDESC2
    Dim rv As Long
    Dim gv As Long
    Dim bv As Long
    
    With display.ddpfPixelFormat
        If (.lFlags And DDPF_RGB) > 0 Then
            rv = CLng(CSng(.lRBitMask) * (r / 255)) And .lRBitMask
            gv = CLng(CSng(.lGBitMask) * (g / 255)) And .lGBitMask
            bv = CLng(CSng(.lBBitMask) * (b / 255)) And .lBBitMask
            
            'If r > 10 And r < 255 Then Stop
            'rv = 20 'r * display.ddpfPixelFormat.lBBitMask
            'gv = 20 'g * display.ddpfPixelFormat.lGBitMask
            'bv = 20 'b * display.ddpfPixelFormat.lRBitMask
            
            GenRGB = rv Or gv Or bv
            
        'r = color And display.ddpfPixelFormat.lBBitMask
        'g = color And (display.ddpfPixelFormat.lGBitMask) \ ((display.ddpfPixelFormat.lBBitMask + 1) * 2) '256
        'b = color And (display.ddpfPixelFormat.lRBitMask) \ display.ddpfPixelFormat.lGBitMask '65536
            
        End If
    End With

End Function

Function loadclick(FilD As CommonDialog) As Boolean

loadclick = False

On Error GoTo 5
GoTo 10
5 Exit Function
10

loadgame = 1

ChDir App.Path

FilD.FileName = App.Path & "\" & "*.plr"
FilD.DefaultExt = "plr"
FilD.ShowOpen

On Error GoTo 0

'If FilD.FileName = "" Then Exit Function

If Not Dir("curgame.dat") = "" Then Kill "curgame.dat"
If Not Dir("plrdat.tmp") = "" Then Kill "plrdat.tmp"


fname = getfile("plrdat.tmp", FilD.FileName)

fname2 = getfile("curgame.dat", FilD.FileName, , 1)

'm'' error 75 avoid
If fname = "" Or fname2 = "" Then Exit Function
Name fname2 As "curgame.dat"

floadchar fname

loadclick = True



End Function

Function killobj(wobj As aobject)

If wobj.X > 0 And wobj.Y > 0 Then map(wobj.X, wobj.Y).object = 0
wobj.X = 0
wobj.Y = 0

End Function

Function killallmonsters()

For a = 1 To UBound(mon())
killmon a, 1
Next a

End Function

Function dbmsg(txt)
'Debug Messager

#If USELEGACY = 1 Then
    If debugmessageson = 1 Then MsgBox txt 'm'' original behavior
#Else
    'm'' new behavior to output debug messages
    Dim fp As Integer
    fp = FreeFile
    Open "debug.log" For Append As #fp
    Print #fp, Date & " " & Timer, txt
    Close #fp
#End If


End Function

Function HasSkill(ByRef skillname As String) As Boolean  'm'' declare

HasSkill = False

For a = 1 To 4
    If UCase(plr.combatskills(a)) = UCase(skillname) Then If getplrskill(skillname) > 0 Then HasSkill = True: Exit Function
Next a

End Function
