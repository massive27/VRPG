Attribute VB_Name = "Debugger"
'm'' DEBUGGER.BAS
'm'' ============
'm'' module created to make here the
'm'' various sub in order to improve the game
'm'' or fix bugs.
'm''
'm'' debugger's name : Massive27 from Eka's portal
'm'' http://aryion.com/forum/viewforum.php?f=52
'm''
'm'' MANY THANKS to : Aleph-Null and wsensor
'm''  who find back the source code :D
'm''
'm'' <= I use that kind of comment in order to get track
'm''    of all my modifications
'm''
'm'' historic
'm'' 2012-03-23 :
'm'' - fixed two call to legacy .TransparentDraw calls
'm'' - fixed legacy TransparentDraw of cSprite.cls
'm'' - rewrote SetTimer call to get the handler for KillTimer
'm'' - adding another control on getfile() in order to make the game clear itself better when quitting
'm'' - wrote a quitting function to avoid crash or memory leakage
'm'' - fixed skilltree.frm and VRmaped.frm from stitching. now, must be integrated

Public API_Timer_Handle As Long

'm'' following declaration to make stitched sources working ...
Public TileColors(1 To 50) As Long 'Minimap colors for tiles
Public chars() As playertype
Public CurChar As Byte 'Current Character

Sub CharLoad()
'm'' from another VRPG source, chars type loading
'm'' for some reason, Duam started to code an handler of 4 players.
'm'' code must be downgraded, each Chars(CurChar). replaced by plr.
'm'' else skill attribution wont work.

ReDim chars(1 To 4) As playertype
'Character initialization
CurChar = 1
For a = 1 To 4
    'chars(a).CharNum = a
    'chars(a).CurActionSlot = 1
    'chars(a).CurActionSet = 1
    'chars(a).TimePoints = 10
    chars(a).bodyname = "body" & roll(6) & ".bmp"
    chars(a).hairname = "hair" & roll(26) & ".bmp"
    chars(a).haircolor = RGB(roll(155), roll(155), roll(155))
    chars(a).Class = getgener("Sorceress", "Amazon", "Valkyrie")
    If chars(a).name = "" Then chars(a).name = getname
    'ApplyClass chars(a), chars(a).Class
    chars(a).level = 1: chars(a).exp = 0: chars(a).expneeded = 600
    chars(a).fatiguemax = greater(50, getend * 30 + 50)
    'setplrskilldescs chars(a)
    
    'chars(a).FaceFile = "face1.bmp"
    'chars(a).NoseFile = "nose1.bmp"
    'chars(a).EyesFile = "eyes1.bmp"
    'chars(a).LipsFile = "lips1.bmp"
    'chars(a).BackHairFile = "backhair1.bmp"
    'chars(a).FrontHairFile = "fronthair1.bmp"
    
    'chars(a).SkinColorLight = RGB(180, 140, 120)
    'chars(a).SkinColorDark = RGB(130, 70, 30)
    
    'chars(a).LipsColor = RGB(235, 10, 15)
    
    'chars(a).HairColorLight = RGB(255, 122, 36)
    'chars(a).HairColorDark = RGB(111, 60, 0)
    
Next a

End Sub

Sub Quitting()
'm'' from Form1_QueryUnload()
'improved to avoid many errors

Form1.Timer1.Enabled = False
KillTimer Form1.hwnd, API_Timer_Handle

nodraw = 1
endingprog = 1
ClearSprites2
BASS_Free
Form1.DMC1.TerminateBASS


For a = 1 To 100
DoEvents
Next a

If Not Dir("curgame.dat") = "" Then Kill "curgame.dat"
If Not Dir("plrdat.tmp") = "" Then Kill "plrdat.tmp"
If Not Dir(App.Path & "\" & plr.name & "\VTDATA*.*") = "" Then Kill App.Path & "\" & plr.name & "\VTDATA*.*"
If Not Dir(App.Path & "\VTDATA*.*") = "" Then Kill App.Path & "\VTDATA*.*"

'm'' Duam forgot to freed the directx stuff.
Set DXLib.picBuffer = Nothing
Set DXLib.Primary = Nothing
Set DXLib.picBuffer2 = Nothing

Dim frm As Form
For Each frm In Forms
     Unload frm
Next frm

End Sub
