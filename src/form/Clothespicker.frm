VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Clothespicker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Screen"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   5
      Left            =   6360
      TabIndex        =   54
      Text            =   "Combo2"
      Top             =   5880
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   4
      Left            =   6360
      TabIndex        =   53
      Text            =   "Combo2"
      Top             =   5640
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   3
      Left            =   6360
      TabIndex        =   52
      Text            =   "Combo2"
      Top             =   5400
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   2
      Left            =   6360
      TabIndex        =   51
      Text            =   "Combo2"
      Top             =   5160
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      Left            =   6360
      TabIndex        =   50
      Text            =   "Combo2"
      Top             =   4920
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      Left            =   6360
      TabIndex        =   48
      Text            =   "Combo2"
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Dexterity"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   44
      Top             =   5880
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Hit Points"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   43
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Strength"
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   42
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Reset"
      Height          =   255
      Left            =   7080
      TabIndex        =   38
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Load Char Class"
      Height          =   375
      Left            =   6720
      TabIndex        =   37
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save Char Class"
      Height          =   375
      Left            =   6720
      TabIndex        =   36
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   34
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   33
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   31
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   30
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   28
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   27
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   25
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   22
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   19
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   15
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   3345
      Left            =   3840
      Picture         =   "Clothespicker.frx":0000
      ScaleHeight     =   219
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   112
      TabIndex        =   8
      Top             =   240
      Width           =   1740
   End
   Begin MSComDlg.CommonDialog Colorpick 
      Left            =   6840
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1005
      Left            =   2640
      Picture         =   "Clothespicker.frx":11FB4
      ScaleHeight     =   945
      ScaleWidth      =   975
      TabIndex        =   4
      Top             =   480
      Width           =   1035
   End
   Begin VB.ComboBox Colorcombo 
      Height          =   315
      Left            =   6360
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.ComboBox List2 
      Height          =   315
      Left            =   6360
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.ComboBox List1 
      Height          =   315
      Left            =   6360
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD CLOTHES"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Class Skills"
      Height          =   255
      Left            =   6720
      TabIndex        =   49
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tertiary Attribute:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   480
      TabIndex        =   47
      Top             =   5880
      Width           =   1905
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Secondary Attribute:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   120
      TabIndex        =   46
      Top             =   5520
      Width           =   2220
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Primary Attribute:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   480
      TabIndex        =   45
      Top             =   5160
      Width           =   1890
   End
   Begin VB.Label Label6 
      Caption         =   "CHARACTER POINTS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   4320
      Width           =   4335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MAX SP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   32
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MAX MP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   29
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MAX HP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "INTELLIGENCE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   1785
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ENDURANCE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   1605
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DEXTERITY:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "STRENGTH:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Label Label3 
      Caption         =   "Amazon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "BODY (Click to change)"
      Height          =   435
      Index           =   1
      Left            =   3960
      TabIndex        =   10
      Top             =   3600
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "HAIR (Click to change)"
      Height          =   435
      Index           =   0
      Left            =   2640
      TabIndex        =   9
      Top             =   0
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "COLOR"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2280
      Width           =   735
   End
End
Attribute VB_Name = "Clothespicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hairnum
'plrhairloaded, plrbodyloaded

Private Sub Combo2_Click(Index As Integer)
plr.classskills(Index + 1) = Combo2(Index).List(Combo2(Index).ListIndex) ' . ItemData
End Sub

Private Sub Command1_Click()

'croll = roll(UBound(clothestypes()))

'mroll = roll(UBound(materialtypes()))
'Do While Not materialtypes(mroll).type = clothestypes(croll).MATERIAL
'    mroll = roll(UBound(materialtypes()))
'Loop

croll = List1.ListIndex + 1
mroll = List2.ListIndex + 1
colorname = Colorcombo.List(Colorcombo.ListIndex)
gencolor colorname, r, g, b
'r = materialtypes(mroll).r
'g = materialtypes(mroll).g
'b = materialtypes(mroll).b

'If r + g + b = 0 Then colorname = gencolor(, r, g, b)

addclothes colorname & " " & clothestypes(croll).name, clothestypes(croll).graph, clothestypes(croll).armor * materialtypes(mroll).armor, clothestypes(croll).wear1, clothestypes(croll).wear2, r, g, b, , 1
Form1.updatinv
Form1.updatbody

updatall

End Sub

Private Sub Command2_Click(Index As Integer)
If Index = 0 Then hairnum = hairnum - 1 Else hairnum = hairnum + 1
If getfile("hair" & hairnum & ".gif", , , , 1) = "" Then If hairnum > 1 Then hairnum = 1 Else hairnum = 1
fname = getfile("hair" & hairnum & ".gif", "Data.pak")
Picture1.Picture = LoadPicture(fname)
plr.hairname = "hair" & hairnum & ".bmp"
plrhairloaded = 0
Form1.updatbody

updatall

End Sub

Private Sub Command3_Click(Index As Integer)

Select Case Index
    Case 0: If plr.str > 1 Then plr.str = plr.str - 1: plr.charpoints = plr.charpoints + 1
    Case 1: If plr.dex > 1 Then plr.dex = plr.dex - 1: plr.charpoints = plr.charpoints + 1
    Case 2: If plr.endurance > 1 Then plr.endurance = plr.endurance - 1: plr.charpoints = plr.charpoints + 1
    Case 3: If plr.int > 1 Then plr.int = plr.int - 1: plr.charpoints = plr.charpoints + 1
    Case 4: If plr.hpmax > 15 Then plr.hpmax = plr.hpmax - 15: plr.charpoints = plr.charpoints + 1
    Case 5: If plr.mpmax > 0 Then plr.mpmax = plr.mpmax - 15: plr.charpoints = plr.charpoints + 1
    Case 6: If plr.spmax > 80 Then plr.spmax = plr.spmax - 4: plr.charpoints = plr.charpoints + 1
End Select

updatall

End Sub

Private Sub Command4_Click(Index As Integer)

If plr.charpoints < 1 Then Exit Sub
#If USELEGACY = 1 Then
Select Case Index
    Case 0: plr.str = plr.str + 1: plr.hpmax = plr.hpmax + 2
    Case 1: plr.dex = plr.dex + 1: plr.spmax = plr.spmax + 2
    Case 2: plr.endurance = plr.endurance + 1: plr.hpmax = plr.hpmax + 2
    Case 3: plr.int = plr.int + 1: plr.mpmax = plr.mpmax + 2: plr.spmax = plr.spmax + 1
    Case 4: plr.hpmax = plr.hpmax + 15
    Case 5: plr.mpmax = plr.mpmax + 15
    Case 6: plr.spmax = plr.spmax + 6
End Select
#Else 'm'' added overflow checks
Select Case Index 'm''
    Case 0 'm''
        If plr.str = 255 Then Exit Sub 'm''
        plr.str = plr.str + 1: plr.hpmax = plr.hpmax + 2 'm''
    Case 1 'm''
        If plr.dex = 255 Then Exit Sub 'm''
        plr.dex = plr.dex + 1: plr.spmax = plr.spmax + 2 'm''
    Case 2 'm''
        If plr.endurance = 255 Then Exit Sub 'm''
        plr.endurance = plr.endurance + 1: plr.hpmax = plr.hpmax + 2 'm''
    Case 3 'm''
        If plr.int = 255 Then Exit Sub 'm''
        plr.int = plr.int + 1: plr.mpmax = plr.mpmax + 2: plr.spmax = plr.spmax + 1 'm''
    Case 4: plr.hpmax = plr.hpmax + 15 'm''
    Case 5: plr.mpmax = plr.mpmax + 15 'm''
    Case 6: plr.spmax = plr.spmax + 6 'm''
End Select 'm''
#End If

plr.charpoints = plr.charpoints - 1

updatall

End Sub

Private Sub Command5_Click()
Colorpick.FileName = "*.cls"
Colorpick.DefaultExt = "*.cls"
Colorpick.ShowOpen

filen = Colorpick.FileName

Open filen For Binary As #1

Put #1, , plr

Close #1

End Sub

Private Sub Command6_Click()
Colorpick.FileName = "*.cls"
Colorpick.DefaultExt = "*.cls"
Colorpick.ShowOpen

filen = Colorpick.FileName

Open filen For Binary As #1

Put #1, , plr

Close #1

End Sub

Private Sub Command7_Click()
'plr.dex = 1
'plr.str = 1
'plr.endurance = 1
'plr.int = 1
'plr.hpmax = 20
'plr.mpmax = 0
'plr.spmax = 80
'plr.charpoints = 100
updatall
End Sub

Private Sub Command8_Click(Index As Integer)
Command8(Index).caption = textcycle(Command8(Index).caption, "Mana", "Hit Points", "Intelligence", "Endurance", "Dexterity", "Strength")

'If Command8(Index).caption = "Dexterity" Then Command8(Index).caption = "Endurance"
'If Command8(Index).caption = "Strength" Then Command8(Index).caption = "Dexterity"

plr.classdata.strmult = 0
plr.classdata.dexmult = 0
plr.classdata.intmult = 0
plr.classdata.hpmult = 0
plr.classdata.mpmult = 0
plr.classdata.endmult = 0

For a = 2 To 0 Step -1
    If Command8(a).caption = "Strength" Then plr.classdata.strmult = 3 - Index
    If Command8(a).caption = "Dexterity" Then plr.classdata.dexmult = 3 - Index
    If Command8(a).caption = "Strength" Then plr.classdata.strmult = 3 - Index
    If Command8(a).caption = "Hit Points" Then plr.classdata.hpmult = 3 - Index
    If Command8(a).caption = "Mana" Then plr.classdata.mpmult = 3 - Index
    If Command8(a).caption = "Endurance" Then plr.classdata.endmult = 3 - Index
Next a


End Sub

Private Sub Form_Load()
loadclothestypes "clothesdata.txt"
createskillcombos

'If cheaton = 1 Then plr.charpoints = 100

For a = 1 To UBound(clothestypes())
    List1.AddItem clothestypes(a).name
Next a

For a = 1 To UBound(materialtypes())
    List2.AddItem materialtypes(a).name
Next a

For a = 1 To 17

Select Case a
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
    Colorcombo.AddItem colorname
Next a

List1.ListIndex = 0
List2.ListIndex = 0
Colorcombo.ListIndex = 0

updatall

End Sub

Private Sub Label1_Click()
Colorpick.ShowColor

Label1.BackColor = Colorpick.color
plr.haircolor = Colorpick.color

plrhairloaded = 0
Form1.updatbody
updatall
End Sub



Private Sub Label3_Click()
plr.Class = InputBox("Class Name?", plr.Class)
plr.classdata.classname = plr.Class
End Sub

Private Sub Label5_Click()

plr.name = InputBox("Name?", , plr.name)
If plr.name = "" Then plr.name = getname
updatall
End Sub

Private Sub Picture1_Click()
updatall
End Sub

Private Sub Picture2_Click()
Static lastbody

If plr.Class = "Succubus" Then Exit Sub
If plr.Class = "Angel" Then Exit Sub
If plr.Class = "Naga" Then Exit Sub
If plr.Class = "TombRaider" Then Exit Sub

lastbody = lastbody + 1
If getfile("body" & lastbody & ".gif", , , , 1) = "" Or lastbody > 10 Then lastbody = 1

Picture2.Picture = LoadPicture(getfile("body" & lastbody & ".gif"))
plr.bodyname = "body" & lastbody & ".gif"

plrbodyloaded = 0
Form1.updatbody
updatall
End Sub

Sub updatall()

Label4(0).caption = "STRENGTH: " & plr.str
Label4(1).caption = "DEXTERITY: " & plr.dex
Label4(2).caption = "ENDURANCE: " & plr.endurance
Label4(3).caption = "INTELLIGENCE: " & plr.int
Label4(4).caption = "HP MAX: " & plr.hpmax
Label4(5).caption = "MP MAX: " & plr.mpmax
Label4(6).caption = "SP MAX: " & plr.spmax

Text1(0).text = plr.hpmax
Text1(1).text = plr.mpmax
Text1(2).text = plr.spmax


Label6.caption = "CHARACTER POINTS: " & plr.charpoints

Label5.caption = plr.name

Label3.caption = plr.Class

Picture2.AutoRedraw = True
'Picture2.ScaleMode = pixel
Picture2.Width = gbody.CellWidth
Picture2.Height = gbody.CellHeight

fullbody.DrawtoDC Picture2.hDC, -100, 1, 1

Picture2.Refresh
End Sub

Private Sub Text1_Change(Index As Integer)

'If Index = 0 Then plr.hpmax = Val(Text1(Index).text)
'If Index = 1 Then plr.mpmax = Val(Text1(Index).text)
'If Index = 2 Then plr.spmax = Val(Text1(Index).text)

End Sub

Function createskillcombos()

'Combo2(a).AddItem
For a = 0 To 5
Combo2(a).clear
Combo2(a).AddItem "Deathblow" ', "sword", "Adds 10% damage, +5% per skill level"
Combo2(a).AddItem "Critical Strike" ', "skull", "10% chance to inflict double damage. Chance increases by 3% per level."
Combo2(a).AddItem "Resilience" ', "circleplus", "-3 points of damage from incoming attacks, plus -2 per level"
Combo2(a).AddItem "Defence" ', "circleplusblue", "+15% to armor, +5% per level"
Combo2(a).AddItem "Dodge" ', "splatterblue" ', "5% chance to dodge incoming attacks, +2% per level"
Combo2(a).AddItem "Endurance" ', "heartorange" ', "+2 HP for every character level, +1 per skill level"
Combo2(a).AddItem "Drain Life" ', "heartdrip" ', "Gain 1% of damage when hitting enemies, +0.5% per level"
Combo2(a).AddItem "Drain Magic" ', "heartdripblue" ', "Convert 1% of the damage you inflict on enemies into mana, +0.5% per level"
Combo2(a).AddItem "Regeneration" ', "plus" ', "Regenerate HP over time. Rate increases with skill level."
Combo2(a).AddItem "Firepower" ', "firepower" ', "Increases damage done with direct attack spells. +10% base, +5% per level."
Combo2(a).AddItem "Accuracy" ', "swordblue" ', "Increases chances to hit with melee weapons. +2 base, +1 per level."
Combo2(a).AddItem "Greed" ', "cash" ', "+10% to gold collected, +4% per level"
Combo2(a).AddItem "Spell Mastery" ', "trianglefire" ', "Spells cost 10% less mana to cast, +3% per level."
Combo2(a).AddItem "Mana Mastery" ', "manafire" ', "+2 mana per character level, +1 per skill level."
Combo2(a).AddItem "Mana Regeneration" ', "heartblue" ', "Increases mana regeneration rate."
Combo2(a).AddItem "Streetfighting" ', "fists" ', "Increases unarmed damage by 30%, plus 15% per level."
Combo2(a).AddItem "Weapons Mastery" ', "swordgrey" ', "Increases all weapon damage by 20%, plus 6% per level."
Combo2(a).AddItem "Giant Stomach" ', "stomach" ', "Allows you to eat monsters!  Higher levels allow you to hold more monsters in your stomach at a time.  1 monster base, +1 monster per skill level." & vbCrLf & "(Press 'E' before attacking something to attempt to eat it)"
Combo2(a).AddItem "Super Acid" ', "stomachacid" ', "Increases the speed at which you digest monsters and the damage they take while in your stomach."
Combo2(a).AddItem "Gluttony" ', "mouth" ', "Increases your chance of successfully swallowing monsters."
Combo2(a).AddItem "Sword Mastery" ', "swordmastery" ', "Increases damage inflicted by sword and knife type weapons.  +2 per die per level."
Combo2(a).AddItem "Spear Mastery" ', "spearmastery" ', "Increases damage inflicted by spear type weapons.  +3 per die per level."
Combo2(a).AddItem "Axe Mastery" ', "axemastery" ', "Increases damage inflicted by axe and mace type weapons.  +2 per die per level."
Combo2(a).AddItem "Bow Mastery" ', "bowmastery" ', "Increases damage inflicted by bow type weapons.  +20% per level."
Combo2(a).AddItem "White Magic" ', "whitemagic" ', "Increases power of white spells and gives access to new white spells."
Combo2(a).AddItem "Grey Magic" ', "greymagic" ', "Increases power of grey spells and gives access to new grey spells."
Combo2(a).AddItem "Fire Magic" ', "firemagic" ', "Increases power of fire spells and gives access to new fire spells."
Combo2(a).AddItem "Air Magic" ', "lightningmagic" ', "Increases power of lightning spells and gives access to new lightning spells."
Combo2(a).AddItem "Demon Summoning" ', "demonsummoning" ', "Increases the power of summoned demons and gives you the ability to summon more."
Combo2(a).AddItem "Sorcery" ', "sorcery" ', "'Basic' magic.  Gives a broad range of spells and abilities that increase your magical powers."
Combo2(a).AddItem "Magic Summoning" ', "magicsummoning" ', "Gives you the ability to summon a magical creature to aid you."
Next a

End Function
