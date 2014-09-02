VERSION 5.00
Begin VB.Form Pickcombatskills 
   Caption         =   "Pick Combat Skills"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form2"
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Vicious"
      Height          =   1800
      Index           =   17
      Left            =   8640
      Picture         =   "Pickcombatskills.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reflection"
      Height          =   1800
      Index           =   16
      Left            =   6960
      Picture         =   "Pickcombatskills.frx":480A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Perfect Strike"
      Height          =   1800
      Index           =   15
      Left            =   6960
      Picture         =   "Pickcombatskills.frx":8171
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2760
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Poisonous"
      Height          =   1800
      Index           =   14
      Left            =   8640
      Picture         =   "Pickcombatskills.frx":CD68
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Damage Energy"
      Height          =   1800
      Index           =   13
      Left            =   5280
      Picture         =   "Pickcombatskills.frx":1161B
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alchemy"
      Height          =   1800
      Index           =   12
      Left            =   3600
      Picture         =   "Pickcombatskills.frx":14F07
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mana Shield"
      Height          =   1800
      Index           =   11
      Left            =   1920
      Picture         =   "Pickcombatskills.frx":19126
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Split Spell"
      Height          =   1800
      Index           =   10
      Left            =   240
      Picture         =   "Pickcombatskills.frx":1CB4C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Piercing Arrow"
      Height          =   1800
      Index           =   9
      Left            =   5280
      Picture         =   "Pickcombatskills.frx":20A05
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Split Arrow"
      Height          =   1800
      Index           =   8
      Left            =   3600
      Picture         =   "Pickcombatskills.frx":25435
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cripple"
      Height          =   1800
      Index           =   7
      Left            =   1920
      Picture         =   "Pickcombatskills.frx":2A051
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Block"
      Height          =   1800
      Index           =   6
      Left            =   240
      Picture         =   "Pickcombatskills.frx":2E62E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stun"
      Height          =   1800
      Index           =   5
      Left            =   8640
      Picture         =   "Pickcombatskills.frx":3285C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cunning Strike"
      Height          =   1800
      Index           =   4
      Left            =   6960
      Picture         =   "Pickcombatskills.frx":36E8F
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vital Strike"
      Height          =   1800
      Index           =   3
      Left            =   5280
      Picture         =   "Pickcombatskills.frx":3B4B9
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Power Strike"
      Height          =   1800
      Index           =   2
      Left            =   3600
      Picture         =   "Pickcombatskills.frx":3FBA0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Charged Strike"
      Height          =   1800
      Index           =   1
      Left            =   1920
      Picture         =   "Pickcombatskills.frx":44262
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Frenzy"
      Height          =   1800
      Index           =   0
      Left            =   240
      Picture         =   "Pickcombatskills.frx":48836
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   16
      Top             =   6600
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PICK COMBAT SKILLS (4)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Pickcombatskills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastpicked As Byte

Private Sub Command1_Click(Index As Integer)
Command1(Index).Visible = False
plr.combatskills(lastpicked + 1) = Command1(Index).caption
lastpicked = lastpicked + 1
Label1.caption = "PICK COMBAT SKILLS (" & 4 - lastpicked & ")"
If lastpicked = 4 Then Unload Me


End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Command1(Index).caption
        Case "Cripple": Label2.caption = "Inflicts 25% damage on the first hit, +4% per level"
        Case "Frenzy": Label2.caption = "Attacks 1 extra nearby enemy per level"
        Case "Split Spell": Label2.caption = "Adds 50% to the mana cost of a spell and adds one per skill level to the number of bolts it fires.  Mana increase is 50% regardless of level."
        Case "Stun": Label2.caption = "Has a chance to stun enemies.  Increased skill increases chances and duration of stun."
        Case "Charged Strike": Label2.caption = "Adds 50% damage, +25% per level"
        Case "Power Strike": Label2.caption = "Adds (Strength times Weapon Dice times skill level) to damage"
        Case "Vital Strike": Label2.caption = "Adds (Dexterity times Weapon Dice times skill level) to damage"
        Case "Cunning Strike": Label2.caption = "Adds (Intelligence times Weapon Dice times skill level) to damage"
        Case "Mana Shield": Label2.caption = "Half damage is absorbed by mana, with a 10% reduction in damage taken to mana per level"
        Case "Block": Label2.caption = "Damage reduced by 25%, +5% per level"
        Case "Damage Energy": Label2.caption = "You get an amount of mana equal to 20% of the damage you take, +10% per level"
        Case "Poisonous": Label2.caption = "Monsters take damage when they try to digest you"
        Case "Alchemy": Label2.caption = "Gain three points of mana plus two per level every turn at the flat cost of 2 goldpieces per turn"
        Case "Split Arrow": Label2.caption = "Fire one extra arrow per level"
        Case "Piercing Arrow": Label2.caption = "Arrows pierce up to 3 enemies, +1 per dice to damage per level"
        Case "Perfect Strike": Label2.caption = "Attacks automatically inflict maximum damage, +5% per level."
        Case "Reflection": Label2.caption = "Reflects weapons back at their firer, multiplying their damage by your level in the reflection skill."
        Case "Vicious": Label2.caption = "Significantly increases damage against more dangerous enemies, based on your relative levels and the level of Vicious you have.  (Roughly +20%-30% per level they have over you, +15% per vicious level)"
End Select
End Sub

