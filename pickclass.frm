VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   Caption         =   "Choose Your Class"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form6"
   ScaleHeight     =   8010
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Gastronomical"
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   19
      Top             =   6840
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Gluttonous"
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   18
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Difficulty"
      Height          =   1095
      Left            =   7080
      TabIndex        =   16
      Top             =   6120
      Width           =   3015
      Begin VB.OptionButton Option1 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Naga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Caller"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Streetfighter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LOAD SAVED GAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   7440
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TombRaider"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Succubus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3550
      Index           =   7
      Left            =   4920
      Picture         =   "pickclass.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Angel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3550
      Index           =   6
      Left            =   4920
      Picture         =   "pickclass.frx":4636
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Valkyrie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3550
      Index           =   5
      Left            =   3360
      Picture         =   "pickclass.frx":8401
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enchantress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3550
      Index           =   4
      Left            =   1800
      Picture         =   "pickclass.frx":BE75
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Priestess"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3550
      Index           =   3
      Left            =   240
      Picture         =   "pickclass.frx":F5AF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Huntress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3550
      Index           =   2
      Left            =   1800
      Picture         =   "pickclass.frx":12820
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sorceress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3550
      Index           =   1
      Left            =   3360
      Picture         =   "pickclass.frx":15D7A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Amazon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3550
      Index           =   0
      Left            =   240
      Picture         =   "pickclass.frx":18985
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog FileD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FileName        =   "*.map"
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "CHOOSE YOUR CHARACTER CLASS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   7560
      Width           =   5175
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   6480
      TabIndex        =   9
      Top             =   840
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   8
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
plr.Class = Command1(Index).caption
Unload Me
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Label1.caption = Command1(Index).caption

If Label1.caption = "Amazon" Then Label2.caption = "Dedicated exclusively to the role of warrior, the Amazons are capable of facing monsters on their own terms. Their stubborn bravery has caused many of them to feed the monsters that were just a little too strong, but the damage the amazons can dish out is substantial." & vbCrLf & vbCrLf & "- +10% to damage, +1% per level" & vbCrLf & "- High HP and Strength" & vbCrLf & "- Pathetic spellcasting ability, moderate dexterity" & vbCrLf & "- Starts with Sword and Axe Masteries" & vbCrLf & "- Excellent selection of combat skills"
If Label1.caption = "Huntress" Then Label2.caption = "The Huntresses are charged with keeping the cities safe for girlkind. They specialize in hunting down monsters and are well known for their precision and talent." & vbCrLf & vbCrLf & "- +2 to hit, +1 per 5 levels" & vbCrLf & "- High Dexterity" & vbCrLf & "- Primary Fire Magic" & vbCrLf & "- Starts with Bow and Spear Masteries" & vbCrLf & "- Broad selection of weapons masteries"
If Label1.caption = "Valkyrie" Then Label2.caption = "The enigmatic Valkyries combine potent defensive abilities with well-rounded combat skills. Their magic helps boost their abilities in combat, and they possess the unusual talent of being able to deflect attacks with any armor they are wearing." & vbCrLf & vbCrLf & "- +10% armor, +1% per level" & vbCrLf & "- Well rounded Attributes" & vbCrLf & "- Primary Air and Earth Magic" & vbCrLf & "- Secondary Water Magic" & vbCrLf & "- Starts with Sword Mastery" & vbCrLf & "- Excellent Defensive Skills"
If Label1.caption = "Priestess" Then Label2.caption = "The Priestess focuses on defensive and lightning magic, but still maintains the power to smite her enemies physically. She is strong but still vulnerable, so to compensate all priestesses are taught to evade attacks." & vbCrLf & vbCrLf & "- -2 to enemy attack rolls" & vbCrLf & "- High Strength, Low Dexterity" & vbCrLf & "- Primary Earth and Air Magic" & vbCrLf & "- Secondary Water Magic" & vbCrLf & "- Starts with Axe Mastery"
If Label1.caption = "Enchantress" Then Label2.caption = "Enchantresses specialize in spells that enhance items and provide other unique abilities. They are weak and fragile on their own, but their magics can make them faster, stronger and smarter than even the mighty Amazons when their powers grow strong enough. Though they cannot take much damage, the magics that suffuse the Enchantresses are such that they gain an almost supernatural ability to heal wounds naturally." & vbCrLf & vbCrLf & "- Slow Natural Regeneration" & vbCrLf & "- High Intelligence, very low Strength" & vbCrLf & "- Primary Grey Magic and Sorcery" & vbCrLf & "- Secondary Magic in all elements" & vbCrLf & "- Has regeneration as a class skill"
If Label1.caption = "Sorceress" Then Label2.caption = "The Sorceress abandons all subtlety in her magics and uses them to destroy her enemies. She gains highly dangerous spells very quickly, but must keep her distance from her opponents, for she is not strong enough to avoid becoming dinner if she isn't careful. Sorceresses have keen eyes for jewels and valuables, and often profit greatly from their ability to appraise such items." & vbCrLf & vbCrLf & "- +30% Gold, +3% per level" & vbCrLf & "- High Intelligence, Low Strength and Dexterity" & vbCrLf & "- Primary Fire Magic and Sorcery" & vbCrLf & "- Secondary Grey and Elemental Magics" & vbCrLf & "- Variety of magic-enhancing skills"
If Label1.caption = "Succubus" Then Label2.caption = "The Succubus' dark powers allow her to steal the life from creatures she destroys. She is strong and her black magic is potent, but she knows no other magical fields." & vbCrLf & vbCrLf & "- Steals 2% Life from Enemies" & vbCrLf & "- High Strength, High HP" & vbCrLf & "- Primary Fire Magic" & vbCrLf & "- Demon summoning magic" & vbCrLf & "Life and Mana draining class skills"
If Label1.caption = "Angel" Then Label2.caption = "The Angel is a potent warrior and magician, capable of great feats of magic, but unfortunately monsters know just how tasty she is and it shows as she finds every creature within ten miles attempting to swallow her down. She regenerates magic more quickly than normal, but is unusually susceptible to being eaten." & vbCrLf & vbCrLf & "- Regenerates Magic Faster" & vbCrLf & "- Enemies more likely to swallow her" & vbCrLf & "- Primary Air and Water Magic" & vbCrLf & "- Secondary Grey Magic"
If Label1.caption = "TombRaider" Then Label2.caption = "Who doesn't want to see her devoured, eh? Eh?" & vbCrLf & vbCrLf & "- High Strength, Dexterity and Intelligence" & vbCrLf & "- Excellent combat skills" & vbCrLf & "- No innate special abilities or magic"
If Label1.caption = "Caller" Then Label2.caption = "A magician who specializes in summoning powerful monsters to do her fighting for her." & vbCrLf & vbCrLf & "- Summoning Spells" & vbCrLf & "- Primary Sorcery" & vbCrLf & "- Secondary Elemental Magics"
If Label1.caption = "Streetfighter" Then Label2.caption = "The Streetfighter forgoes weapons and instead just pummels her enemies with her bare hands." & vbCrLf & vbCrLf & "- High Physical Stats" & vbCrLf & "- Streetfighting Class Skill" & vbCrLf & "- Secondary Earth Magic"
If Label1.caption = "Naga" Then Label2.caption = "A snake woman with an unsurpassed stomach capacity." & vbCrLf & vbCrLf & "- Giant Stomach Class Skill" & vbCrLf & "- Starts with Giant Stomach skill" & vbCrLf & "- Can only wear clothes that go on her upper half, since she doesn't have legs" & vbCrLf & "- Primary Water Magic" & vbCrLf & "- Secondary Sorcery"


End Sub

Private Sub Command2_Click()

If loadclick(FileD) = True Then Unload Me

Exit Sub
'loadgame = 1
''On Error GoTo 5
''GoTo 10
''5 Exit Sub
''10
'
''FileD.FileName = App.Path & "\" & "*.plr"
''FileD.DefaultExt = "plr"
''FileD.ShowOpen
''Unload Me
''Form1.Timer1.Enabled = False
'
'ChDir App.Path
'
'FileD.FileName = "*.plr"
'FileD.DefaultExt = "plr"
'FileD.ShowOpen
'
'If Dir(FileD.FileName) = "" Then Exit Sub
'
'fname = getfile("plrdat.tmp", FileD.FileName)
'
'getfile "curgame.dat", FileD.FileName, , 1
'
'floadchar fname ' FileD.FileName
'
''Kill "plrdat.tmp"
'
''If FileD.FileTitle = "" Or Dir(FileD.FileTitle) = "" Then Exit Sub
'Me.Hide
'Form1.Show
'
''floadchar FileD.FileTitle
'
Unload Me

Exit Sub

loadgame = 1
'On Error GoTo 5
'GoTo 10
'5 Exit Sub
10

FileD.FileName = "*.plr"
FileD.DefaultExt = "plr"
FileD.ShowOpen
'Unload Me
'Form1.Timer1.Enabled = False
Form1.Show

floadchar FileD.FileName

Unload Me

Exit Sub

'If plr.Class = "" Then GoTo 3
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

plr.X = 25: plr.Y = 25
soundon = 1
'MsgBox "Welcome to Duamutef's Glorious Vore RPG! Click on the Help menu if you don't know how to play."
'MsgBox "You step into your home town, relieved to be back among the people of your own clan. For weeks you have been hearing stories of a horrible dragon. This dragon, Thirsha, lives in a volcano to the far North. Seers prophesy that soon she will sweep through and kill everyone in the region, including you. You must find the three dragon keys and gain entrance to her lair--it is said that if you are able to do so, you can destroy her."

'Form1.Timer1.Enabled = True
'Form1.Timer2.Enabled = True
If Form1.Command2.Enabled = True Then Form1.Command2.SetFocus

Unload Me

End Sub

Private Sub Form_Load()
If winlev = 0 Then Command1(6).Visible = False: Command1(7).Visible = False: Command1(8).Visible = False: Command1(9).Visible = False: Command1(10).Visible = False: Command1(11).Visible = False
Option1(0).Value = True
If winlev = 1 Then Command1(8).Visible = False: Command1(9).Visible = False: Command1(10).Visible = False: Command1(11).Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.caption = ""
Label2.caption = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
Do While (1)
End
Loop
End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.caption = "DIFFICULTY"
Label2.caption = "This selects the difficulty.  Higher difficulties mean monsters do more damage and have higher hit points, but also give more experience and gold.  Shops and maps are randomized when on any difficulty other than Normal.  When you win the game, you unlock extra character classes depending on what difficulty you won the game on."
End Sub

Private Sub Option1_Click(Index As Integer)
plr.difficulty = Index
For a = 0 To 2
If Not Index = a Then Option1(a).Value = False
Next a
End Sub
