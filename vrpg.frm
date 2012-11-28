VERSION 5.00
Object = "{2C1EC115-F1BA-11D3-BF43-00A0CC32BE58}#9.0#0"; "DMC2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DA729162-C84F-11D4-A9EA-00A0C9199875}#1.60#0"; "MpqCtl.ocx"
Begin VB.Form Form1 
   Caption         =   "Duamutef's Glorious Vore RPG V2.0"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   915
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin DMC2.DMC DMC1 
      Left            =   9480
      Top             =   4320
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9360
      Top             =   2760
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SELL ITEM"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   840
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      Height          =   1500
      Left            =   10440
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Hide"
      Height          =   255
      Left            =   9720
      TabIndex        =   25
      Top             =   8760
      Width           =   615
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   7800
      Width           =   7575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pause"
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Skills"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   22
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DROP ITEM"
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   8280
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text6"
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   3120
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   381
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.PictureBox Picture4 
      Height          =   1215
      Left            =   10680
      Picture         =   "vrpg.frx":0000
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   18
      Top             =   7680
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   3600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   7200
      Width           =   495
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   7200
      Width           =   495
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Sound"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   8640
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cast"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Blocked"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10200
      TabIndex        =   10
      Top             =   720
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4680
      TabIndex        =   9
      Text            =   "Loading..."
      Top             =   6840
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0009F9FF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Gold:"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.ListBox Combo1 
      Height          =   1620
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog FileD 
      Left            =   11160
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FileName        =   "*.map"
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Water"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10200
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   9600
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   7680
      Picture         =   "vrpg.frx":2162
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9000
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture7 
      Height          =   6855
      Left            =   0
      ScaleHeight     =   453
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   597
      TabIndex        =   30
      Top             =   0
      Width           =   9015
   End
   Begin MPQCONTROLLib.MpqControl MpqControl1 
      Left            =   11280
      Top             =   5040
      _Version        =   65542
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
      TitleHidden     =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Note to self:  You need to make the menus visible in the menu editor again"
      Height          =   975
      Left            =   3720
      TabIndex        =   29
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Menu filemen 
      Caption         =   "File"
      Begin VB.Menu savchar 
         Caption         =   "Save Character"
      End
      Begin VB.Menu loadchar 
         Caption         =   "Load Character"
      End
      Begin VB.Menu randomworldmen 
         Caption         =   "Random World"
         Visible         =   0   'False
      End
      Begin VB.Menu reasskeys 
         Caption         =   "Reassign Keys"
      End
      Begin VB.Menu inctimer 
         Caption         =   "Increase Game Speed"
      End
      Begin VB.Menu dectimer 
         Caption         =   "Decrease Game Speed"
      End
      Begin VB.Menu exitgame 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu savm 
         Caption         =   "Save Map"
      End
      Begin VB.Menu loadm 
         Caption         =   "Load Map"
      End
      Begin VB.Menu nmap 
         Caption         =   "New Map"
      End
      Begin VB.Menu rchunk 
         Caption         =   "Random Chunk"
      End
      Begin VB.Menu fl 
         Caption         =   "Fill"
      End
      Begin VB.Menu opmaped 
         Caption         =   "MAP EDITOR"
      End
      Begin VB.Menu getobd 
         Caption         =   "Get Object Data"
      End
      Begin VB.Menu saveseg 
         Caption         =   "Save Segment"
      End
      Begin VB.Menu loadsegmen 
         Caption         =   "Load Segment"
      End
   End
   Begin VB.Menu vclothes 
      Caption         =   "View Equipment"
   End
   Begin VB.Menu takoff 
      Caption         =   "Take Clothes Off"
   End
   Begin VB.Menu uneqwep 
      Caption         =   "Unequip Weapon"
   End
   Begin VB.Menu helpmen 
      Caption         =   "Help"
      Begin VB.Menu helpbas 
         Caption         =   "How To Play"
      End
      Begin VB.Menu aboutmen 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Use the Newsprite command to create a sprite.
'After that, just changing that particular sprite's x, y, and cell (frame) values
'will automatically change where and how it is drawn.

'The drawmain command draws to whatever object hdc you specify via hdc.

Dim curcheatword As String

Dim horktype As Integer
Dim bugger As Integer
Dim endingprog As Byte
'Dim buggersprite(500)


Private Sub aboutmen_Click()
playsound "girmyself.wav"
showform10 "ABOUT", 1, "Duamutef.jpg"
End Sub

Private Sub Check3_Click()
'soundon = Check3.Value
If Check3.Value = 1 Then soundoff = 0 Else soundoff = 1
End Sub

Private Sub Check4_Click()
nodraw = Check4.Value
End Sub

Private Sub Combo1_Click()
mp = spells(Combo1.ItemData(Combo1.ListIndex)).mp
skillpercminus mp, "Spell Mastery", 10, 3
displayspell spells(Combo1.ItemData(Combo1.ListIndex))
Command2.caption = "CAST (" & mp & " MP)"
End Sub

Private Sub Command1_Click()

checkallconvs

'If Not Dir("test.dat") = "" Then Kill "test.dat"

'Open "fork.txt" For Output As #1
'Write #1, "XXXXXXX"
'Close #1

'Open "fork2.txt" For Output As #1
'Write #1, "YYY"
'Close #1

'Open "fork3.txt" For Output As #1
'Write #1, "ZZZZZZZZZ"
'Close #1


'addfile "fork.txt", "test.dat", 1
'addfile "fork2.txt", "test.dat", 1
'addfile "fork3.txt", "test.dat", 1

'Stop

'addfile "fork2.txt", "test.dat", 1

End Sub

Private Sub Command2_Click()
If Combo1.ListIndex = -1 Or plr.instomach > 0 Then Exit Sub
If spells(Combo1.ItemData(Combo1.ListIndex)).target = "Target" Or getfromstring(spells(Combo1.ItemData(Combo1.ListIndex)).target, 1) = "Enchant" Then Ccom = "SPELL" & Combo1.ItemData(Combo1.ListIndex): Command2.caption = "PICK YOUR TARGET" Else castspell Combo1.ItemData(Combo1.ListIndex)
End Sub

Private Sub Command3_Click()
If List1.ListIndex = -1 Then Exit Sub
Ccom = ""
If inv(List1.ItemData(List1.ListIndex)).graphname = "" Then inv(List1.ItemData(List1.ListIndex)).graphname = "clothes.bmp"
inv(List1.ItemData(List1.ListIndex)).graphloaded = 0
makeobjtype inv(List1.ItemData(List1.ListIndex))
createobj inv(List1.ItemData(List1.ListIndex)).name, plr.x, plr.y
inv(List1.ItemData(List1.ListIndex)).name = ""
updatinv
End Sub

Private Sub Command4_Click()
'm'' load player status into first Chars() struct
chars(CurChar) = plr 'm''

Form9.Show 1
End Sub

Private Sub Command5_Click()
If paused = 1 Then paused = 0 Else paused = 1
'If Timer1.Enabled = True Then Timer1.Enabled = False: Timer2.Enabled = False: gamemsg "Game Paused--Click Again to Resume": Exit Sub
'If Timer1.Enabled = False Then Timer1.Enabled = True: Timer2.Enabled = True
End Sub

Private Sub Command6_Click()
If Command6.caption = "Hide" Then
Text7.Visible = False
Command6.caption = "Show"
Else:
Text7.Visible = True
Command6.caption = "Hide"
End If

End Sub

Private Sub Command7_Click()
If List1.ListIndex = -1 Then Exit Sub
Ccom = ""
If Not geteff(inv(List1.ItemData(List1.ListIndex)), "NoEat", 1) = "" Then gamemsg "You cannot sell that.": Exit Sub
dough = greater(Int(getworth(inv(List1.ItemData(List1.ListIndex))) / 4), 1)
randsound "gold", 3
plr.gp = plr.gp + dough
gamemsg "You exchange the " & inv(List1.ItemData(List1.ListIndex)).name & " for " & dough & " gold."
killitem inv(List1.ItemData(List1.ListIndex))
orginv
End Sub

Private Sub dectimer_Click()
If plr.timerspeed = 5 Then plr.timerspeed = 10: Exit Sub
If plr.timerspeed <= 200 Then plr.timerspeed = plr.timerspeed + 20
End Sub

Private Sub exitgame_Click()

    Call Debugger.Quitting 'm''

End Sub

Private Sub fl_Click()
fillmap edittile

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

keyt = Chr(KeyAscii)

'hipsforme = Gain 10,000 HP
'leylines = Gain 10,000 MP
'tittedtitan = Gain 50 Strength, Dexterity, Intelligence and Endurance
'oneup = Gain 1 level
'levelmebaby = Gain 20 levels
'striptease = Gain a random set of clothing and/or armor
'beingpoorsucks = Gain 10,000 gold
'workfromhome = Gain 1,000,000 gold
'laxative = Escape from current monster's belly

'allaboutallia = Teleport to Allia
'chillinwithfirebellies = Teleport to Skinbane
'wormygurgle = Teleport to wormgut
'sexyinblack = Teleport to blackscourge
'ihaveadeathwish = Teleport to Thirsha's Lair
'iwannagohome = Teleport to home city
'angelsaretasty = Teleport to City of Angels

If Val(keyt) > 0 Then KeyAscii = 0: Exit Sub
'cheat codes
curcheatword = Right(curcheatword, 20) & keyt
If InStr(1, curcheatword, "hipsforme") Then Ccom = "": curcheatword = "": plr.hpmax = plr.hpmax + 10000
If InStr(1, curcheatword, "leylines") Then Ccom = "": curcheatword = "": plr.mpmax = plr.mpmax + 10000
If InStr(1, curcheatword, "tittedtitan") Then Ccom = "": curcheatword = "": plr.str = plr.str + 50: plr.dex = plr.dex + 50: plr.int = plr.int + 50: plr.endurance = plr.endurance + 50
If InStr(1, curcheatword, "levelmebaby") Then Ccom = "": curcheatword = "": For a = 1 To 20: gainlevel 1: Next a: plr.exp = 0: plr.plrdead = 0
If InStr(1, curcheatword, "oneup") Then Ccom = "": curcheatword = "": gainlevel 1: plr.exp = 0: plr.plrdead = 0
If InStr(1, curcheatword, "striptease") Then Ccom = "": curcheatword = "": giverandomclothes
If InStr(1, curcheatword, "workfromhome") Then Ccom = "": curcheatword = "": plr.gp = plr.gp + 1000000
If InStr(1, curcheatword, "beingpoorsucks") Then Ccom = "": curcheatword = "": plr.gp = plr.gp + 10000
If InStr(1, curcheatword, "specialrecipe") Then Ccom = "": curcheatword = "": plr.lpotions = 100
'If InStr(1, curcheatword, "editoron") Then Ccom = "": curcheatword = "": editon = 1: edit.Visible = True

If InStr(1, curcheatword, "oopsie") Then Ccom = "": curcheatword = "": killallmonsters

If InStr(1, curcheatword, "laxative") Then
    Ccom = "": curcheatword = ""
    gamemsg getesc
    stopsounds
    playsound "swallow3.wav"
    playsound "burp" & roll(5) & ".wav"
    If plr.diglevel < 4 Then playsound "grunt" & roll(9) + 2 & ".wav"
    mon(plr.instomach).cell = 1: plr.instomach = 0: swallowcounter = -6: If plr.hp < 1 Then plr.hp = 1: plr.plrdead = 0
    stomachlevel = 0
End If

If InStr(1, curcheatword, "allaboutallia") Then Ccom = "": curcheatword = "": plr.x = 5: plr.y = 5: gotomap "Allia.txt"
If InStr(1, curcheatword, "firebaby") Then Ccom = "": curcheatword = "": plr.x = 5: plr.y = 5: gotomap "skinbane.txt"
If InStr(1, curcheatword, "wormygurgle") Then Ccom = "": curcheatword = "": plr.x = 5: plr.y = 5: gotomap "cityofwormgut.txt"
If InStr(1, curcheatword, "sexyinblack") Then Ccom = "": curcheatword = "": plr.x = 5: plr.y = 5: gotomap "blackscourge.txt"
If InStr(1, curcheatword, "ihaveadeathwish") Then Ccom = "": curcheatword = "": plr.x = 5: plr.y = 5: gotomap "thirshaslair.txt"
If InStr(1, curcheatword, "iwannagohome") Then Ccom = "": curcheatword = "": plr.x = 5: plr.y = 5: gotomap "VRPGData.txt"
If InStr(1, curcheatword, "angelsaretasty") Then Ccom = "": curcheatword = "": plr.x = 5: plr.y = 5: gotomap "cityofangels.txt"

KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If paused = 1 Then MsgBox "The game is paused.  Unpause the game to continue.": Exit Sub

Static lastmove

'If KeyCode = vbKeyNumpad6 Then plr.x = plr.x + 1: plr.y = plr.y - 1: plr.xoff = -96
'If KeyCode = vbKeyNumpad2 Then plr.x = plr.x + 1: plr.y = plr.y + 1: plr.yoff = -48
'If KeyCode = vbKeyNumpad4 Then plr.x = plr.x - 1: plr.y = plr.y + 1: plr.xoff = 96
'If KeyCode = vbKeyNumpad8 Then plr.x = plr.x - 1: plr.y = plr.y - 1: plr.yoff = 48
'If KeyCode = vbKeyNumpad3 Then plr.x = plr.x + 1: plr.yoff = -24: plr.xoff = -48
'If KeyCode = vbKeyNumpad1 Then plr.y = plr.y + 1: plr.yoff = -24: plr.xoff = 48
'If KeyCode = vbKeyNumpad7 Then plr.x = plr.x - 1: plr.yoff = 24: plr.xoff = 48
'If KeyCode = vbKeyNumpad9 Then plr.y = plr.y - 1: plr.yoff = 24: plr.xoff = -48

If Left(Ccom, 7) = "BindKey" Then
    zork = Right(Ccom, Len(Ccom) - 7)
    If zork = "NW" Then plr.keys.moveNW = KeyCode: gamemsg "Press desired North (Up and Right) Key": Ccom = "BindKeyN"
    If zork = "N" Then plr.keys.moveN = KeyCode: gamemsg "Press desired NorthEast (Directly Right) Key": Ccom = "BindKeyNE"
    If zork = "NE" Then plr.keys.moveNE = KeyCode: gamemsg "Press desired East (Down and Right) Key": Ccom = "BindKeyE"
    If zork = "E" Then plr.keys.moveE = KeyCode: gamemsg "Press desired SouthEast (Directly Down) Key": Ccom = "BindKeySE"
    If zork = "SE" Then plr.keys.moveSE = KeyCode: gamemsg "Press desired South (Down and Left) Key": Ccom = "BindKeyS"
    If zork = "S" Then plr.keys.moveS = KeyCode: gamemsg "Press desired SouthWest (Directly Left) Key": Ccom = "BindKeySW"
    If zork = "SW" Then plr.keys.moveSW = KeyCode: gamemsg "Press desired West (Up and Left) Key": Ccom = "BindKeyW"
    If zork = "W" Then plr.keys.moveW = KeyCode: gamemsg "Press desired 'Eat' Key": Ccom = "BindKeyEat"
    If zork = "Eat" Then plr.keys.eat = KeyCode: gamemsg "Press desired Life Potion Key": Ccom = "BindKeyLifePot"
    If zork = "LifePot" Then plr.keys.lifepot = KeyCode: gamemsg "Press desired Mana Potion Key": Ccom = "BindKeyManaPot"
    If zork = "ManaPot" Then plr.keys.manapot = KeyCode: gamemsg "Reassignment Complete.": Ccom = "BindKeyDone": zork = "Done"
    If zork = "Done" Then gamemsg "Reassignment complete.": Ccom = ""
    KeyCode = 0: Exit Sub
End If

'Form1.SetFocus
If KeyCode = vbKey3 Then If getplrskill(plr.combatskills(1)) > 0 Then Picture6.Picture = LoadPicture("skillicon-" & remspaces(plr.combatskills(1)) & ".jpg"): usingskill = plr.combatskills(1): Picture6.Visible = True
If KeyCode = vbKey4 Then If getplrskill(plr.combatskills(2)) > 0 Then Picture6.Picture = LoadPicture("skillicon-" & remspaces(plr.combatskills(2)) & ".jpg"): usingskill = plr.combatskills(2): Picture6.Visible = True
If KeyCode = vbKey5 Then If getplrskill(plr.combatskills(3)) > 0 Then Picture6.Picture = LoadPicture("skillicon-" & remspaces(plr.combatskills(3)) & ".jpg"): usingskill = plr.combatskills(3): Picture6.Visible = True
If KeyCode = vbKey6 Then If getplrskill(plr.combatskills(4)) > 0 Then Picture6.Picture = LoadPicture("skillicon-" & remspaces(plr.combatskills(4)) & ".jpg"): usingskill = plr.combatskills(4): Picture6.Visible = True
If KeyCode = vbKey7 Then Picture6.Visible = False: usingskill = ""

If KeyCode = vbKeyA And cheaton = 1 And Shift = 2 Then
    aroll = roll(6)
    Dim runk As objecttype
    runk.effect(1, 1) = getgener("BONSTR", "BONDEX", "BONINT")
    runk.effect(1, 2) = roll(3) + 2
    If aroll = 6 Then addaura "swordaura.bmp", roll(255), roll(255), roll(255), 15, runk
    If aroll = 1 Then addaura "aura6.bmp", roll(255), roll(255), roll(255), 7, runk
    If aroll = 2 Then addaura "aura1.bmp", roll(255), roll(255), roll(255), 3, runk
    If aroll = 3 Then addaura "aura5.bmp", roll(255), roll(255), roll(255), 6, runk
    If aroll = 4 Then addaura "aura2.bmp", roll(255), roll(255), roll(255), 4, runk
    If aroll = 5 Then addaura "aura3.bmp", roll(255), roll(255), roll(255), 4, runk
End If

If KeyCode = vbKeyT And cheaton = 1 And Shift = 2 Then Cargotrade.Show
If KeyCode = vbKeyP And cheaton = 1 And Shift = 2 Then showform10 "PARTS:#$SHIELDPARTS": Shipsys.Show

If KeyCode = vbKeyC And cheaton = 1 And Shift = 2 Then giverandomclothes
If KeyCode = vbKeyC And cheaton = 1 And Shift = 3 Then Clothespicker.Show
'''If KeyCode = vbKeyZ And cheaton = 1 And Shift = 2 Then savebindata plr.curmap & ".dat": loadbridge ' addfile plr.curmap & ".dat", "Universe.dat", 1: loadbridge
'''If KeyCode = vbKeyS And Shift = 2 And cheaton = 1 Then loadstation 'zark = genmap2: loaddata zark
'''If KeyCode = vbKeyA And Shift = 3 And cheaton = 1 Then genstarshipmap "Random", 4, 4, 4, 3, 1, 26
If KeyCode = vbKeyN And cheaton = 1 And Shift = 2 Then
'''    addnews "#MONSTEREAT", "The Allied Star Protectorate", "Beastfodder", "Cruiser", getgener("space amoeba", "space dragon", "space squid", "space octopus")
'''    addnews "#WAR", "The Xebbebban Empire", "The Zadiian Consulate", "Xebbebba", "Zadiian", "Giedi Prime"
    'Dim zarkstr As String
    'If roll(2) = 1 Then suf = "Rumors" Else suf = ""
    'zarkstr = createnews(suf)
    'showform10 zarkstr, -1
    newsname = getname
    showform10 "#RANDOM", 1
End If

If KeyCode = 107 Then
    zerk = 0
    For a = 1 To 4
    If plr.combatskills(a) = usingskill Then zerk = a: Exit For
    Next a
    
    Do While (1)
    zerk = zerk + 1
    If zerk > 4 Then zerk = 0: Picture6.Visible = False: usingskill = "": Exit Do
    If getplrskill(plr.combatskills(zerk)) > 0 Then Picture6.Picture = LoadPicture(getfile("skillicon-" & remspaces(plr.combatskills(zerk)) & ".jpg")): usingskill = plr.combatskills(zerk): Picture6.Visible = True: Exit Do
    Loop

End If



If cheaton = 1 Then Command1.Visible = True

'If cheaton = 1 And KeyCode = vbKeyB Then distortboobs 35, 1
'If KeyCode = vbKeyB Then makebijdang plr.x + 1, plr.y + 1, 2
If cheaton = 1 And KeyCode = vbKeyN And Shift = 2 Then gamemsg getname ': getname3

If KeyCode = vbKeyG And Shift = 2 And cheaton = 1 Then plr.gp = 5000000

If KeyCode = vbKeyD And Shift = 2 And cheaton = 1 Then If isnaked = False Then digestclothes 2500, 1 Else plr.diglevel = plr.diglevel + 1: playsound "grunt" & roll(9) + 2 & ".wav": playsound "burp" & roll(5) & ".wav": If plr.diglevel > 7 Then plr.diglevel = 0: Form1.updatbody Else Form1.updatbody
If KeyCode = vbKeyL And Shift = 2 And cheaton = 1 Then gainlevel 1: plr.exp = 0: plr.plrdead = 0
'If KeyCode = vbKeyC And Shift = 2 And cheaton = 1 Then plr.Class = InputBox("Class?")
If KeyCode = vbKeyL And Shift = 3 And cheaton = 1 Then For a = 1 To 20: gainlevel 1: Next a: plr.exp = 0: plr.plrdead = 0

If KeyCode = vbKeyA And Shift = 3 And cheaton = 1 Then loadseg "10shop8.seg", 2, 2, 3, 22, 29, "Cyberbody", "Cyberlegs", "Cyberarms", "Cyberarms", "Black Cotton Bra", "Black Cotton Panties"
If KeyCode = vbKeyS And Shift = 3 And cheaton = 1 Then updatbody 1

'If plr.plrdead = 1 And cheaton = 1 Then turnthing: Exit Sub
'If plr.plrdead = 1 Then Exit Sub

'If KeyCode = vbKey1 or  And plr.lpotions > 0 Then
'If plr.hp = 0 Then If plr.lpotions > 1 Then plr.lpotions = Int(plr.lpotions * 0.75) 'Take 25% of potions if HP at 0
'plrdamage -rolldice((plr.lpotionlev + 2) ^ 3, 6): plr.lpotions = plr.lpotions - 1
'End If

'Life Potions
If KeyCode = plr.keys.lifepot Or KeyCode = vbKey1 Then
If plr.lpotions > 0 Then
'If plr.hp = 0 Then If plr.lpotions > 1 Then plr.lpotions = Int(plr.lpotions * 0.75)
    'If isexpansion = 0 Then plrdamage -rolldice((plr.lpotionlev + 2) ^ 3, 6), 0.5: plr.lpotions = plr.lpotions - 1
    'If isexpansion = 1 Then plrdamage -rolldice((plr.level + 2), 10), 0.5: plr.lpotions = plr.lpotions - 1
    plrhpgain = plrhpgain + plr.level * 10 * (plr.lpotionlev / 5 + 1): plr.lpotions = plr.lpotions - 1: plr.hplost = lesser(plr.hplost, plr.hpmax - plr.hp - plrhpgain)
KeyCode = 0
Exit Sub
End If

End If
'rolldice((plr.lpotionlev + 2) ^ 3, 6)
If (KeyCode = plr.keys.manapot Or KeyCode = vbKey2) And plr.mpotions > 0 Then plr.mp = plr.mp + plr.mpmax * (0.5 + plr.mpotionlev * 0.25): plr.mpotions = plr.mpotions - 1: If plr.mp > plr.mpmax Then plr.mp = plr.mpmax

    'PLAYER IN STOMACH

If plr.instomach > 0 Then
    'If KeyCode = vbKeyE And Shift = 1 And cheaton = 1 Then gamemsg "MEGAESCAPE" & vbCrLf & "You force your way out of " & montype(mon(plr.instomach).type).Name & "'s belly!": mon(plr.instomach).cell = 1: plr.instomach = 0: swallowcounter = -9: GoTo 5
    'If plr.diglevel < 4 And plr.fatigue < plr.fatiguemax * 0.8 Then mon(plr.instomach).xoff = 12 * (roll(3) - 2): mon(plr.instomach).yoff = 12 * (roll(3) - 2) 'plr.yoff = 12 * (roll(3) - 2): plr.xoff = 12 * (roll(3) - 2)
    'If lastmove = turncount Then Exit Sub Else lastmove = turncount + 1
    'If plr.diglevel > 3 Or stomachlevel > 2 Then GoTo 3
    'If plr.fatigue > plr.fatiguemax * 0.8 Then gamemsg "You are too exhausted to struggle.": GoTo 3
    'If KeyCode = vbKeyNumpad5 Then GoTo 3
    'If plr.hp / gethpmax < 0.2 Then addiff = 1
    If stomachlevel = 1 Then addiff = 1
    If stomachlevel > 1 Then addiff = 2
    If stomachlevel > 2 Then addiff = 3
    If plr.hp / gethpmax <= 0.1 Then addiff = addiff + 1 '10% hp or less makes it much harder to escape
    If plr.hp / gethpmax >= 0.8 Then addiff = addiff - 1 '80% hp or more makes it much easier
    'If stomachlevel > 2 Then playsound "blurble" & roll(3) & ".wav"
    tryescape addiff
    
    'If (succroll((plr.str / 3 + 2) + (plr.level / 4) + (((plr.hp / plr.hpmax) * 10) - 4) + (Int(instomachcounter / 4) - 4), 6 - lesser(plr.diglevel, 1)) + roll(skilltotal("Squirm", 2, 1)) > succroll(montype(mon(plr.instomach).type).escapediff * 2)) Or (roll(20 - (((plr.hp / plr.hpmax) * 10) - 4)) = 1) Or plr.fatigue = 0 Or (roll(roll(plr.hpmax / 5)) > plr.hp And plr.hp > 1) Then
    '    plrescape getesc
        'stopsounds
        'playsound "swallow3.wav"
        'playsound "burp" & roll(5) & ".wav"
        'If plr.diglevel < 4 Then playsound "grunt" & roll(9) + 2 & ".wav"
        'mon(plr.instomach).cell = 1: plr.instomach = 0: swallowcounter = -6: If plr.hp < 1 Then plr.hp = 1: plr.plrdead = 0: GoTo 5
    'End If
    'addfatigue 5
    '3 turnthing
    
    Exit Sub
End If
5
If plr.plrdead = 1 Then Exit Sub
'If KeyCode = vbKeyNumpad6 Then plrmove 1, -1
'If KeyCode = vbKeyNumpad4 Then plrmove -1, 1
'If KeyCode = vbKeyNumpad8 Then plrmove -1, -1
'If KeyCode = vbKeyNumpad2 Then plrmove 1, 1
'If KeyCode = vbKeyNumpad3 Or KeyCode = vbKeyRight Then plrmove 1, 0
'If KeyCode = vbKeyNumpad7 Or KeyCode = vbKeyLeft Then plrmove -1, 0
'If KeyCode = vbKeyNumpad9 Or KeyCode = vbKeyUp Then plrmove 0, -1
'If KeyCode = vbKeyNumpad1 Or KeyCode = vbKeyDown Then plrmove 0, 1

If KeyCode = plr.keys.moveNE Then plrmove 1, -1
If KeyCode = plr.keys.moveSW Then plrmove -1, 1
If KeyCode = plr.keys.moveNW Then plrmove -1, -1
If KeyCode = plr.keys.moveSE Then plrmove 1, 1
If KeyCode = plr.keys.moveE Then plrmove 1, 0
If KeyCode = plr.keys.moveW Then plrmove -1, 0
If KeyCode = plr.keys.moveN Then plrmove 0, -1
If KeyCode = plr.keys.moveS Then plrmove 0, 1


If KeyCode = vbKeyA Then Ccom = "Attack"
'If KeyCode = vbKeyC Then If spells(Combo1.ItemData(Combo1.ListIndex)).target = "Target" Then Ccom = "SPELL" & Combo1.ItemData(Combo1.ListIndex) Else castspell Combo1.ItemData(Combo1.ListIndex)

'If KeyCode = vbKeyE And cheaton = 1 And Shift = 3 Then editon = 1: gamemsg "***Edit Mode On***": edittile = 2: Text1.Visible = False: Text2.Visible = False: Text3.Visible = False: Text4.Visible = False: Combo1.Visible = False: Form5.Show

If KeyCode = vbKeyAdd Then edittile = edittile + 1
If KeyCode = vbKeySubtract And edittile > 0 Then edittile = edittile - 1

If KeyCode = plr.keys.eat And Shift = 0 Then Ccom = "EAT": gamemsg "Choose who you want to eat."

If KeyCode = vbKeyS And editon = 1 And Shift = 3 Then
FileD.DefaultExt = "map"
FileD.ShowSave
savemap FileD.FileName
End If

KeyCode = 0

'drawall
End Sub

Private Sub Form_Load()
Me.caption = "Duamutef's Glorious Vore RPG, V" & curversion
'For a = 0 To 50
'newsprite Me.hDC, "explos.bmp", a * 2, a Mod 20, 5
'Next a
'bugger = 1
'recolor RGB(0, 0, 0), RGB(100, 100, 0), Picture1
'Picture1.Picture = Picture1.Image
'spritemaps(1).cmap.CreateFromPicture Picture1.Picture, 1, 1, , RGB(0, 0, 0)

'SetTimer Me.hwnd, 51, 75, AddressOf TimerHandler1
'm''SetTimer Me.hwnd, 52, 15, AddressOf TimerHandler2
Debugger.API_Timer_Handle = SetTimer(Me.hwnd, 52, 15, AddressOf TimerHandler2)

If editon = 0 Then edit.Visible = False
' Parameters - handle, ID
'KillTimer Me.hwnd, 50

' Parameters - handle, ID
'KillTimer Me.hwnd, 50

End Sub

Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

dorkx = Int((x - 12) / 96 + (y - 12) / 48) + plr.x - 10
dorky = Int(y / 48 - x / 96) + plr.y - 2

getXY x, y


'dorkx = (X - plr.X - Y + plr.Y) / 48 + 400 - plr.xoff
'dorkY = (X + Y - plr.X - plr.Y) * 24 - plr.yoff + 300 - (obj.CellHeight - 48) - (midtile * 24)

If Command1.Visible = True Then Command1.caption = "X:" & x & " Y:" & y

If x < 1 Or x > mapx Then Exit Sub
If y < 1 Or y > mapy Then Exit Sub

If map(x, y).monster > 0 Then dispatk montype(mon(map(x, y).monster).type), mon(map(x, y).monster).hp


If editon = 1 And Button = 1 And x > 0 And y > 0 Then
If Shift = 0 Then map(x, y).tile = edittile: map(x, y).blocked = Check1.Value * 2 Else map(x, y).ovrtile = edittile: map(x, y).blocked = Check2.Value
If Shift = 1 And edittile = 0 Then map(x, y).blocked = 0

End If

If editon = 1 And Button = 2 And x > 0 And y > 0 Then
If Shift = 1 Then map(x, y).blocked = 0: map(x, y).ovrtile = 0 Else map(x, y).tile = 0
End If

End Sub



Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
y = y - offset
x2 = x: y2 = y
getXY x, y
If x < 1 Or x > mapx Then Exit Sub
If y < 1 Or y > mapy Then Exit Sub
If plr.plrdead = 1 Then Exit Sub
'Shift-Rightclick copies an object, CTRL-rightclick makes a new one, ALT-rightclick
'pastes an object or edits it's strings

'If Button = 1 And map(x, y).monster > 0 Then MsgBox "You charm" & montype(mon(map(x, y).monster).type).name: mon(map(x, y).monster).owner = 2

'If Button = 1 And map(x, y).monster = 0 Then f = createmonster("Succubus", x, y): mon(f).owner = 1

'If Shift = 2 And Button = 2 Then
'    If map(x, y).object > 0 Then Form3.start objs(map(x, y).object).type Else horktype = createobjtype("New Object", InputBox("Graphic?"), 115, 115, 115, 0.5): Form3.start horktype
'End If

'If Shift = 1 And Button = 2 And map(x, y).object > 0 Then
'    horktype = objs(map(x, y).object).type
'End If

'If Shift = 4 And Button = 2 Then
'    If map(x, y).object > 0 Then
'    objs(map(x, y).object).name = InputBox("Name? (Currently " & objs(map(x, y).object).name & ")")
'    objs(map(x, y).object).string = InputBox("String 1? (Currently " & objs(map(x, y).object).string & ")")
'    objs(map(x, y).object).string2 = InputBox("String 2? (Currently " & objs(map(x, y).object).string2 & ")")
'    Else:
'    createobj objtypes(horktype).name, x, y, InputBox("Name?", "COPY OBJECT"), InputBox("String 1?"), InputBox("String 2?")
'    End If
'End If

'If Shift = 3 And Button = 2 Then map(x, y).object = 0

'PUT ANYTHING THAT SHOULD BE POSSIBLE WHILE IN A MONSTER STOMACH BEFORE THIS LINE
If plr.instomach > 0 And Not Left(Ccom, 6) = "SCROLL" Then Exit Sub

'If Button = 1 And Shift = 0 Then
'    zarkx = 0: zarky = 0
'    If x > plr.x Then zarkx = 1
'    If x < plr.x Then zarkx = -1
'    If y > plr.y Then zarky = 1
'    If y < plr.y Then zarky = -1
'    plrmove zarkx, zarky
'End If

If Button = 1 And wep.type = "Bow" And plr.plrdead = 0 And plr.instomach = 0 Then
    'dicedice = wep.dice
    dicedam = getplrdamage(1, 1)
    'dicedam = wep.damage + (plr.level / 4)
    skillpercmod dicedam, "Bow Mastery", 20, 20
    skillpercmod dicedam, "Weapons Mastery", 20, 6
    addfatigue greater(wep.weight - greater(getstr, getdex), 0)
    If usingskill = "Split Arrow" Then If spendsp(skilltotal("Split Arrow", 2, 1)) Then shootat x2, y2, 1, wep.dice & ":" & dicedam, , , skilltotal("Split Arrow", 2, 1): GoTo 5
    If usingskill = "Piercing Arrow" Then If spendsp(skilltotal("Piercing Arrow", 2, 1)) Then shootat x2, y2, 1, wep.dice & ":" & dicedam * (1 + skilltotal("Piercing Arrow", 1, 1) / 10), , , , , skilltotal("Piercing Arrow", 2, 1): GoTo 5
    shootat x2, y2, 1, wep.dice + Int(plr.dex / 4) & ":" & dicedam
5 allmove = 1: turnthing
End If
'If Button = 1 Then shootat x2 - 400, y2 - 300, roll(4) + 1, 12 & ":" & 12: turnthing


'If Ccom = "Attack" Then plrattack map(X, Y).monster
If Left(Ccom, 5) = "SPELL" Then Call castspell(getnum(Ccom), , , x2, y2) ' Then turnthing
If Left(Ccom, 6) = "SCROLL" Then Call castspell(getnum(Ccom), , 1, x2, y2) ' Then turnthing
                                'Spells were costing two turns (castspell calls turnthing), so turnthing is commented out
'If Button = 2 Then Stop
If Button = 2 And Not Combo1.ListIndex < 0 Then Call castspell(Combo1.ItemData(Combo1.ListIndex), map(x, y).monster, , x2, y2) ' Then turnthing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Debugger.Quitting
Exit Sub

'On Error Resume Next
nodraw = 1
endingprog = 1
ClearSprites2
BASS_Free
Form1.DMC1.TerminateBASS
'SetTimer Me.hwnd, 50, 200, AddressOf TimerHandler
' Parameters - handle, ID
'KillTimer Me.hwnd, 51
KillTimer Me.hwnd, 52

For a = 1 To 100
DoEvents
Next a

If Not Dir("curgame.dat") = "" Then Kill "curgame.dat"
If Not Dir("plrdat.tmp") = "" Then Kill "plrdat.tmp"
If Not Dir(App.Path & "\" & plr.name & "\VTDATA*.*") = "" Then Kill App.Path & "\" & plr.name & "\VTDATA*.*"
If Not Dir(App.Path & "\VTDATA*.*") = "" Then Kill App.Path & "\VTDATA*.*"

Unload Form10
Unload Clothespicker
Unload Form3
Unload Form4
Unload Form5
Unload Form6
'Unload Form7
Unload Form8
Unload Form9
Unload MonStats
Unload Pickcombatskills
Unload TitleScreen
'Unload All

Dim frm As Form

For Each frm In Forms
     Unload frm
Next frm


'End

'Do While (1)
'End
'Loop
End Sub

Private Sub Form_Resize()
Picture7.Width = Form1.ScaleWidth
Picture7.Height = Form1.ScaleHeight

Text7.Top = Form1.ScaleHeight - Text7.Height
List1.Top = Form1.ScaleHeight - 200

Command7.Top = Form1.ScaleHeight - 50
Command3.Top = Form1.ScaleHeight - 50
Command5.Top = Form1.ScaleHeight - Command5.Height
Command6.Top = Form1.ScaleHeight - Command6.Height

Picture3(0).Top = Form1.ScaleHeight - 120
Picture3(1).Top = Form1.ScaleHeight - 120

Picture4.Top = Form1.ScaleHeight - Picture4.Height
Picture4.Left = Form1.ScaleWidth - Picture4.Width

Picture6.Top = Form1.ScaleHeight - 216
Picture6.Left = Form1.ScaleWidth - Picture6.Width

Command4.Top = Form1.ScaleHeight - 112
Command4.Left = Form1.ScaleWidth - Command4.Width

End Sub

Private Sub Form_Terminate()

'Clean up data files

'End
'End
'End

End Sub

Private Sub Form_Unload(Cancel As Integer)
'DMC1.TerminateBASS
End Sub

Private Sub getobd_Click()
End
End Sub

Private Sub helpbas_Click()
showform10 "HELP", 1, , "MAIN2"
'MsgBox "Control your character using the number pad (Make sure Num Lock is on). Simply move into items to pick them up, move into characters to speak to them, and move into monsters to attack them.  To cast a spell, select it in the list on the left and either hit the 'Cast' button or right-click on your target for the spell. Note that the spell controls will not appear until you learn your first spell--only the Sorceress starts with any.  Double-click on any item in your inventory to equip it or use it. The two potions on the bottom of your screen are life and mana potions--hit the '1' and '2' keys (On the top of the keyboard, not on the number pad) to use them.  The 'New Skill' button in the lower right allows you to pick new skills for your character--you gain skill points each time you gain a level which can be spent on these skills.  The button brings up a selection screen--the skills set apart on the far left are your class skills, which can be purchased for half the skill point cost."
'MsgBox "To use the new combat skills, first you must purchase them (In the skills screen) then press 1-4 to activate them.  They will be set up in the order that you selected them when you made your character--ie the first skill you selected will be assigned to the '1' key.  To stop using a combat skill, hit '5'."
End Sub

Private Sub inctimer_Click()
If plr.timerspeed = 10 Then plr.timerspeed = 5
If plr.timerspeed = 20 Then plr.timerspeed = 10
If plr.timerspeed >= 30 Then plr.timerspeed = plr.timerspeed - 20
'plr.timerspeed = greater(1, Int(plr.timerspeed / 2))
End Sub

Private Sub List1_Click()
If List1.ListIndex > -1 Then displayobj inv(List1.ItemData(List1.ListIndex))
End Sub

Private Sub List1_DblClick()
If List1.ListIndex = -1 Then Exit Sub
If plr.plrdead > 0 Then Exit Sub
If Left(Ccom, 5) = "SPELL" Then castspell Right(Ccom, Len(Ccom) - 5), List1.ListIndex + 1: Exit Sub
If getfromstring(Ccom, 1) = "GEM" Then If addgem(List1.ItemData(List1.ListIndex), getfromstring(Ccom, 2)) Then gamemsg "You add the " & inv(getfromstring(Ccom, 2)).name & " to the " & inv(List1.ItemData(List1.ListIndex)).name: killitem inv(getfromstring(Ccom, 2)): Ccom = "": playsound "gem1.wav": Exit Sub

takeobj -1, List1.ItemData(List1.ListIndex), 1, dummyobj
End Sub

Private Sub loadchar_Click()

loadclick FileD
Exit Sub

'On Error GoTo 5
'GoTo 10
'5 Exit Sub
'10
'
'ChDir App.Path
'
'FileD.FileName = "*.plr"
'FileD.DefaultExt = "plr"
'FileD.ShowOpen
'
'fname = getfile("plrdat.tmp", FileD.FileName)
'
'fname2 = getfile("curgame.dat", FileD.FileName, , 1)
'If Not Dir("curgame.dat") = "" Then Kill "curgame.dat"
'Name fname2 As "curgame.dat"
'floadchar fname ' FileD.FileName
'
''Kill "plrdat.tmp"


End Sub

Private Sub loadm_Click()
FileD.FileName = "*.seg;*.map"
FileD.ShowOpen
FileD.DefaultExt = "seg"
loadmap FileD.FileName
End Sub

Private Sub loadsegmen_Click()
On Error GoTo 10
GoTo 5
10 Exit Sub
5
FileD.DefaultExt = "seg"
FileD.FileName = "*.seg"
FileD.ShowOpen
If FileD.FileName = "*.seg" Or FileD.FileName = "" Then Exit Sub
Call loadmapseg(FileD.FileTitle)
'loadseg FileD.FileTitle, 1, 1, 1, 1, 1
plr.x = Int(mapx / 2)
plr.y = Int(mapy / 2)
End Sub

Private Sub MpqControl1_GotMessage(ByVal text As String)
Debug.Print text
End Sub

Private Sub newcharmen_Click()
3 newchar 1

'newchar
If plr.Class = "" Then GoTo 3
Set cStage = New cBitmap

plr.plrdead = 0
plr.diglevel = 0
plr.instomach = 0

cStage.CreateAtSize 800, 600
loadbijdang

lastsprite = 1
lastmap = 1
Form1.Show
'Form1.Show

ReDim map(1 To 200, 1 To 200) As tiletype

Set tilespr = New cSpriteBitmaps
tilespr.CreateFromFile "tiles1.bmp", 5, 5, , RGB(0, 0, 0)
Set tilespr2 = New cSpriteBitmaps
tilespr2.CreateFromFile "tiles2.bmp", 5, 5, , RGB(0, 0, 0)
Set ovrspr = New cSpriteBitmaps
ovrspr.CreateFromFile "overlays2.bmp", 5, 10, , RGB(0, 0, 0)
Set transovrspr = New cSpriteBitmaps
tilespr.CreateFromFile "transoverlays2.bmp", 5, 10, , RGB(0, 0, 0)

makesprite digbody(1), Form1.Picture1, "digbody.bmp"
makesprite digbody(2), Form1.Picture1, "digbody2.bmp"
makesprite digbody(3), Form1.Picture1, "digbody3.bmp"
makesprite digbody(4), Form1.Picture1, "digbody4.bmp"
makesprite digbody(5), Form1.Picture1, "digbody5.bmp"

plr.gp = 500

Form1.updatbody

makesprite waterspr, Form1.Picture1, "underwater1.bmp", 0, 0, 255, 1

loaddata "VRPGData.txt"

plr.x = 25: plr.y = 25
'soundon = 1
'MsgBox "Welcome to Duamutef's Glorious Vore RPG! Click on the Help menu if you don't know how to play."
'MsgBox "You step into your home town, relieved to be back among the people of your own clan. For weeks you have been hearing stories of a horrible dragon. This dragon, Thirsha, lives in a volcano to the far North. Seers prophesy that soon she will sweep through and kill everyone in the region, including you. You must find the three dragon keys and gain entrance to her lair--it is said that if you are able to do so, you can destroy her."

'Form1.Timer1.Enabled = True
'Form1.Timer2.Enabled = True
If Form1.Command2.Enabled = True Then Form1.Command2.SetFocus

End Sub

Private Sub nmap_Click()
mapx = Val(InputBox("Width?"))
mapy = Val(InputBox("Height?"))
ReDim map(1 To mapx, 1 To mapy)

For a = 1 To mapx
For b = 1 To mapy
    map(a, b).tile = 1
Next b
Next a

ReDim mon(1 To 1)
totalmonsters = 0
plr.x = 5
plr.y = 5
End Sub

Private Sub opmaped_Click()
Form5.Show
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Randomize
'        col = Picture1.Point(X, Y)
'        redd = col Mod 256
'        green = Int((col Mod 65536) / 256)
'        blue = Int(col / 65536)
'        Command1.Caption = "R" & redd & "G" & green & "B" & blue
'Picture1.Picture = LoadPicture("greybonedragon.bmp")
'rangecolor Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255), Picture1, Rnd
'spritemaps(1).cmap.CreateFromPicture Picture1.Picture, 1, 1, , RGB(0, 0, 0)
End Sub

Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
x = x * 800 / Picture7.ScaleWidth + 1
y = y * 600 / Picture7.ScaleHeight + 1
Form_MouseMove Button, Shift, x, y
End Sub

Private Sub Picture7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
x = x * 800 / Picture7.ScaleWidth + 1
y = y * 600 / Picture7.ScaleHeight + 1
Form_MouseUp Button, Shift, x, y
End Sub

Private Sub randomworldmen_Click()
ReDim genmaps(12, 12) As String
stonk = genmap("TestRandom", 1, 6, 6, , , , , , 10)
gotomap stonk

End Sub

Private Sub rchunk_Click()
randomchunk Val(InputBox("Size-ish?")), , , edittile
End Sub

Private Sub reasskeys_Click()
Ccom = "BindKeyNW": gamemsg "Key Reassignment: Press the keys you wish to assign.  First, Press your desired NorthWest (Up) Key."
End Sub

Private Sub savchar_Click()

On Error GoTo 5
GoTo 10
5 Exit Sub
10

ChDir App.Path


FileD.FileName = App.Path & "\" & plr.name & ".plr"
FileD.DefaultExt = "plr"
FileD.ShowSave

savebindata Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"
fsavechar "plrdat.tmp" 'FileD.FileTitle

addfile "plrdat.tmp", FileD.FileName, 1
addfile "curgame.dat", FileD.FileName, 1

Kill Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"

'Kill "plrdat.tmp"
'Kill Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"

'fsavechar plr.name & ".plr"

End Sub

Private Sub saveseg_Click()
On Error GoTo 10
GoTo 5
10 Exit Sub
5
FileD.DefaultExt = "seg"
FileD.FileName = "*.seg"
FileD.ShowSave
If FileD.FileName = "*.seg" Or FileD.FileName = "" Then Exit Sub
Call savemapseg(FileD.FileTitle)
End Sub

Private Sub savm_Click()
FileD.DefaultExt = "map"
FileD.ShowSave
savemap FileD.FileName
End Sub

Private Sub takoff_Click()

If plr.instomach > 0 Then gamemsg "You're too busy being digested to change clothes now.": Exit Sub


If topclothes < 9 Then takeoffclothes topclothes
End Sub

Private Sub Text1_GotFocus()
If Command2.Enabled = True Then Command2.SetFocus
End Sub

Public Sub TimerH1()
'If spritesloaded = 0 Then Exit Sub
'bugger = bugger + 1
'If bugger = 6 Then bugger = 1
'For a = 1 To lastsprite - 1
'sprite(a).Cell = bugger
'sprite(a).X = bugger * 30
'Next a
'drawmain Me.hDC

MsgBox "Obsolete function TimerH1 called"
Stop
If endingprog > 0 Then endingprog = endingprog + 1: Exit Sub

If needbodyupdt = 1 Then updatbody: needbodyupdt = 0
'Timer1.Enabled = False
turnswitch = turnswitch + 1
If turnswitch >= 7 Then turnswitch = 0: turnthing

If stilldrawing = 0 Then drawall
If plr.xoff <> 0 Or plr.yoff <> 0 Then
For a = 1 To 100
    If diff(plr.xoff, 0) >= 24 Then shooties(a).x = shooties(a).x + plr.xoff
    If diff(plr.yoff, 0) >= 24 Then shooties(a).y = shooties(a).y + plr.yoff
Next a
End If

If plr.xoff > 0 Then plr.xoff = plr.xoff - 24: If diff(plr.xoff, 0) < 24 Then plr.xoff = 0
If plr.xoff < 0 Then plr.xoff = plr.xoff + 24: If diff(plr.xoff, 0) < 24 Then plr.xoff = 0
If plr.yoff > 0 Then plr.yoff = plr.yoff - 12: If diff(plr.yoff, 0) < 12 Then plr.yoff = 0
If plr.yoff < 0 Then plr.yoff = plr.yoff + 12: If diff(plr.yoff, 0) < 12 Then plr.yoff = 0

updatpotions

'DoEvents
'Timer1.Interval = 100
'Timer1.Enabled = True

'If turncount Mod 15 = 0 Then updatbody
'If plr.instomach > 0 Then updatbody
End Sub


Function updatbody(Optional sav = 0)

'If Timer1.Enabled = False Then needbodyupdt = 1: Exit Function

Picture2.Width = 56
Picture2.Height = 109

Picture1.Width = 212 'formerly 112
Picture1.Height = 219

Picture1.Picture = LoadPicture("")
Picture2.Picture = LoadPicture("")
'picture1.Width=


Set picDD1 = DD.CreateSurface(makesurfdesc(212, 219))
clssurface picDD1
drawbody picDD1, 100, 0, True

'gbody.TransparentDraw Picture1.hDC, 0, 0, 1
'gpanties.TransparentDraw Picture1.hDC, 0, 0, 1
'gbra.TransparentDraw Picture1.hDC, 0, 0, 1
'glower.TransparentDraw Picture1.hDC, 0, 0, 1
'gupper.TransparentDraw Picture1.hDC, 0, 0, 1


'Restore this line for stomach gook drippings...
'If plr.instomach > 0 Then gookup Picture1, 30, plr.instomach


'digjunk Picture1, 30, 10
'Picture1.Picture = Picture1.image
'selfassign Picture1

fullbody.CreateFromSurface picDD1, 1, 1

    'Girl graphics maker thingy
    If sav = 1 Then
        'Picture1.Visible = True
        'Big: 112x219
        'Small: 56x109
        Picture1.Width = 212 'formerly 112
        Picture1.Height = 219
        
        Picture2.Cls
        Picture2.AutoRedraw = True
        Picture2.Width = 168 '318 + 53
        fullbody.DrawtoDC Picture1.hDC, 1, 1, 1
        'distortchar 30, 1
        Picture1.Visible = True: Picture2.Visible = True
        
        'Set mouthgraph = New cSpriteBitmaps
        'mouthgraph.CreateFromFile "mouth.bmp", 1, 1
        'gbody.LockMe
        'skincolor = gbody.DXS.GetLockedPixel(58, 70)
        'gbody.UnlockMe
        'getrgb skincolor, r1, g1, b1
        'mouthgraph.recolor r1, g1, b1
        'mouthgraph.DrawtoDC Picture1.hDC, 1, 1, 1
        
        'Normal Pic
        fullbody.DrawtoDC Picture1.hDC, 1, 1, 1
        Picture2.PaintPicture Picture1.image, 0, 0, 56, 109, 112, , 112

        'Mouth pic
        Set mouthgraph = New cSpriteBitmaps
        mouthgraph.CreateFromFile "mouth.bmp", 1, 1
        Picture1.Width = 212 'formerly 112
        Picture1.Height = 219
        gbody.LockMe
        skincolor = gbody.DXS.GetLockedPixel(180, 70)
        gbody.UnlockMe
        getrgb skincolor, b1, g1, r1
        mouthgraph.recolor r1, g1, b1 ', 0.2
        mouthgraph.TransparentDraw fullbody.DXS, 99, 0, 1
        fullbody.DrawtoDC Picture1.hDC, 1, 1, 1
        Picture2.PaintPicture Picture1.image, 112, 0, 56, 109, 112, , 112 '213, 0, 106, 109, 112, , 112
        'Exit Function
        
        'Stuffed Pic
        fullbody.CreateFromSurface picDD1, 1, 1
        distortchar fullbody, 30, 2
        fullbody.DrawtoDC Picture1.hDC, 1, 1, 1
        Picture2.PaintPicture Picture1.image, 56, 0, 56, 109, 112, , 112
        
        'Save Image
        Picture2.Picture = Picture2.image
        For a = 1 To 500
            If Dir("girl" & a & ".bmp") = "" Then SavePicture Picture2.Picture, "girl" & a & ".bmp": Exit For
        Next a
    End If

If plr.foodinbelly > 0 Then distortchar fullbody, plr.foodinbelly * 5, 1

'fullbody.CreateFromPicture Picture1, 1, 1, , 0
'Picture2.Width = 106

'halffy Picture1, Picture2
'Picture2.AutoRedraw = True
'Picture2.PaintPicture Picture1.Picture, 0, 0, 106, 109
'BitBlt Picture2.hDC, 0, 0, Picture1.Width, Picture1.Height, Picture1.hDC, 0, 0, SRCCOPY

'Picture2.Picture = Picture2.image
'selfassign Picture2
Set plrgraphs = New cSpriteBitmaps
'plrgraphs.CreateFromPicture Picture2, 1, 1, , RGB(0, 0, 0)

plrgraphs.CreateFromSurface fullbody.DXS, 1, 1, , , 0.5


End Function

Sub updatspells()
Combo1.clear
spelz = 0
For a = 1 To totalspells
    If spells(a).has = 0 Then GoTo 5
    spelz = spelz + 1
    Combo1.AddItem getfromstring(spells(a).name, 1)
    Combo1.ItemData(spelz - 1) = a
5 Next a
If spelz = 0 Then Combo1.Enabled = False: Combo1.Visible = False: Command2.caption = "You don't know any spells": Exit Sub Else Combo1.Enabled = True: Combo1.Visible = True: Command2.caption = "Select a Spell"
If Combo1.ListIndex = -1 Then Combo1.ListIndex = 0
End Sub

Public Sub TimerH2()
'Timer2.Enabled = False
drawbijdang
'If Form10.Visible = False And Form1.Timer1.Enabled = False Then Form1.Timer1.Enabled = True
'Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
If paused = 1 Then GoTo 5
    Timer1.Interval = plr.timerspeed
    If spaceon = 0 Then Timer2Z Else spacetimer
5    DoEvents
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
'If grloading = 1 Then Stop
'If allloaded = 0 Then
'    For a = 1 To objts
'        If objtypes(a).graphloaded = 0 And objtypes(a).cells > 0 Then slowmakesprite objtypes(a).graph, Picture1, objtypes(a).graphname, objtypes(a).r, objtypes(a).g, objtypes(a).b, objtypes(a).l, objtypes(a).cells: objtypes(a).graphloaded = 1: Exit For
'    Next a
'    If a = objts + 1 Then alloaded = 1
'End If
End Sub

Private Sub Text3_Click()
Clothespicker.Show
End Sub

Private Sub uneqwep_Click()
If plr.instomach > 0 Then gamemsg "If you let go of your weapon, it will likely slip down into the beast's intestines.": Exit Sub
getitem wep.obj: wep.digged = 0: killitem wep.obj: wep.dice = 0: wep.damage = 0: lwepgraph "", Picture1: loadbijdang
wep.type = ""
End Sub

'Private Sub Timer3_Timer()
'If plr.plrdead = 1 Then
'turnthing
'End Sub

Private Sub vclothes_Click()
bongstr = ""

For a = 1 To 8
    buckstr = ""
    If clothes(a).loaded = 1 Then buckstr = displayobj(clothes(a).obj, 0) & vbCrLf
    'If clothes(a).digested = 1 Then buckstr = "Partially digested " & buckstr
    If Not clothes(a).name = "" Then bongstr = bongstr & buckstr
Next a
Dim dong As Long
dong = wep.dice * wep.damage

MsgBox "You are wearing:" & vbCrLf & bongstr & vbCrLf & vbCrLf & "Total Armor: " & clothesarmor & vbCrLf & "Current Weapon: " & displayobj(wep.obj, 0) & " (" & wep.type & " class)" & vbCrLf & "Total Weight:" & getplrweight

End Sub

Sub updatinv()

List1.clear

For a = 1 To 50
    If Not inv(a).name = "" Then List1.AddItem inv(a).name: List1.ItemData(List1.ListCount - 1) = a
Next a

lpl = plr.lpotionlev: If lpl < 1 Then lpl = 1 Else If lpl > 3 Then lpl = 3
mpl = plr.mpotionlev: If mpl < 1 Then mpl = 1 Else If mpl > 3 Then mpl = 3

lfn = getfile("life" & lpl & ".jpg", "Data.pak")
mfn = getfile("magic" & mpl & ".jpg", "Data.pak")
Picture3(0).Picture = LoadPicture(lfn)
Picture3(1).Picture = LoadPicture(mfn)

End Sub

Sub updatpotions()

Label1(0).caption = plr.lpotions
Label1(1).caption = plr.mpotions

End Sub


Private Sub DMC1_ErrorOccurred(ByVal where As String, ByVal info As String)

    If DMC1.Error = True Then
        Debug.Print "DMC ERROR"
        Debug.Print "Where: " & where & "  Why: " & info & "  Details: " & DMC1.LastError
    End If

End Sub
 

