Attribute VB_Name = "Add_UI"
'm'' ADD_UI
'm'' ============
'm'' module created to handle the new UI
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


Sub drawui()
Dim HPS As String 'health point string
Dim MPS As String 'mana point string
Dim CPS As String 'combat point string
Dim GS As String 'gold amount string
Dim SST As String 'status string

Dim HPcol As Long 'colors

Dim lHPm As Long
Dim lcalc As Long

'get datas
lHPm = gethpmax 'call for max hp

'some calculation that shouldnt be here
If plr.sp > plr.spmax Then plr.sp = plr.spmax
If plr.hp < 1 And plr.instomach = 0 Then plr.hp = 1
If plr.plrdead = 1 Then plr.hp = 0


'set strings
HPS = "HP:" & plr.hp & "/" & lHPm
MPS = "MP:" & plr.mp & "/" & getmpmax
CPS = "Combat Points: " & plr.sp & "/" & plr.spmax
GS = "Gold: " & plr.gp

'SST = plr.name & " the " & plr.Class & vbCrLf & "Exp: " & plr.exp & "/" & plr.expneeded & vbCrLf & "Level " & plr.level & vbCrLf & vbCrLf & "Strength:" & getstr & "/" & plr.str & vbCrLf & "Endurance:" & plr.endurance & "/" & getend & vbCrLf & "Dexterity:" & getdex & "/" & plr.dex & vbCrLf & "Intelligence:" & getint & "/" & plr.int
If plr.charpoints > 0 Then SST = SST & vbCrLf & "(Click here to spend character points)"



'calc colors
HPcol = vbBlue ' RGB(0, 0, 255)
lcalc = plr.hp * 10 \ lHPm
Select Case lcalc
    Case 9
        HPcol = vbGreen 'RGB(0, 255, 0)
    Case 8
        HPcol = RGB(155, 255, 0)
    Case 7, 6
        HPcol = vbYellow 'RGB(255, 255, 0)
    Case 5, 4
        HPcol = RGB(255, 155, 0)
    Case 3, 2, 1, 0
        HPcol = vbRed 'RGB(255, 0, 0)
End Select

'draw
Dim UIF As IFont
'Set UIF = IFont
'UIF.put_Name "MS Sans Serif"
'UIF.put_Bold 1
'UIF.put_Size 10

With DXLib.picBuffer
    .SetForeColor 0& 'black
    '.SetFont UIF
    .SetFillColor 0&
    .DrawBox 404, 404, 510, 420
    .DrawBox 404, 424, 510, 441
    .drawtext 405, 405, HPS, False
    .drawtext 405, 425, MPS, False
    .drawtext 405, 450, CPS, False
    .drawtext 405, 480, GS, False
    '.drawtext 405, 500, SST, False
    .SetForeColor HPcol
    .drawtext 404, 404, HPS, False
    .SetForeColor vbBlue
    .drawtext 404, 424, MPS, False
    .SetForeColor &HC0C0C0
    .drawtext 404, 449, CPS, False
    .SetForeColor &H9F9FF
    .drawtext 404, 479, GS, False
End With


End Sub

Sub RemOldUI()

#If USELEGACY <> 1 Then
    'm'' hide old textboxes
    Form1.Text1.Visible = False 'm''
    Form1.Text2.Visible = False 'm''
    Form1.Text8.Visible = False 'm''
    Form1.Text4.Visible = False 'm''
#End If
End Sub
