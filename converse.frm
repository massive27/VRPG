VERSION 5.00
Begin VB.Form Form10 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Dialogue"
   ClientHeight    =   8025
   ClientLeft      =   2280
   ClientTop       =   2070
   ClientWidth     =   10725
   LinkTopic       =   "Form10"
   Picture         =   "converse.frx":0000
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   715
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4935
      Left            =   3600
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   7920
      Top             =   2280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hello"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice or whatever"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice or whatever"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   7680
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice or whatever"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   7440
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice or whatever"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice or whatever"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice or whatever"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice or whatever"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   6480
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice or whatever"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hello"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   135
      TabIndex        =   10
      Top             =   3495
      Width           =   10575
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim displev As Double
Dim curmsg As String
Dim curchoice As Byte

Private Type choicet
    caption As String
    choicename As String '0=Caption, 1=Choicename
    command As String
End Type

Dim lastchoice
Dim lastcom
Dim buying

Dim choice(7) As choicet 'Conversation ID, choice#, caption/choiceto
Dim commands(100, 15, 3) As String
Dim commandname(100) As String
Dim choiceobjs(6) As objecttype
Dim obr, obg, obb As Integer 'Object R, G and B

Sub disptext(txt, Optional leave = 0, Optional pic = "")
Form10.caption = "Dialogue"
'Form10.Cls
displev = 0

If txt = "'#RUMOR'" Then txt = createnews("Rumors")
If txt = "'#NEWS'" Then txt = createnews
If txt = "'#WANT'" Then txt = createwant

stdfilter txt
curmsg = txt
If Not pic = "" Then loaddapic pic
'curmsg = swaptxt(swaptxt(txt, "/", ","), "$NAME", plr.name)
Label1.FontSize = 12
Label3.FontSize = 12
If Len(curmsg) > 680 Then Label1.FontSize = 10: Label3.FontSize = 10
If Len(curmsg) > 880 Then Label1.FontSize = 8: Label3.FontSize = 8
If leave = 1 Then clearchoices: addchoice "##EXIT##", "(Leave)"
'Timer1.Enabled = True
End Sub

Sub loaddapic(filen)
Form10.Cls
origfile = filen
If filen = "(NONE)" Then Exit Sub
If Dir(filen) = "" Then filen = getfile(filen, "Data.pak", , , 1)
If filen = "" Then filen = getfile(Left(origfile, Len(origfile) - 4) & ".jpg", "Data.pak", , , 1)
If filen = "(NONE)" Or filen = "" Or Dir(filen) = "" Then Exit Sub
Picture1.Picture = LoadPicture(filen)

Form10.PaintPicture Picture1.Picture, Form10.ScaleWidth / 2 - (Picture1.ScaleWidth / 2), 1 ', Picture1.ScaleWidth, Picture1.ScaleHeight
End Sub

Private Sub Form_GotFocus()
DoEvents

On Error GoTo 5
GoTo 10
5 Unload Me
10

For a = 0 To 7
Label2(a).ForeColor = RGB(150, 150, 150)
Next a
'Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Form10.Cls
Label1.ForeColor = RGB(255, 255, 255)
Erase commands()
For a = 0 To 7
Label2(a).ForeColor = RGB(150, 150, 150)
Next a
'Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then Cancel = 1 Else Form10.Cls
End Sub

Private Sub Label1_Click()
'Label1.caption = Len(curmsg)
End Sub

Private Sub Label2_Click(Index As Integer)
curchoice = findchoice(choice(Index).choicename, Index)
If buying = 0 Then updatchoice
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


For a = 0 To 7
Label2(a).ForeColor = RGB(150, 150, 150)
Next a

Label2(Index).ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Timer1_Timer()

Form10.continue

'Timer1.Enabled = False
'Exit Sub
'If Form10.Visible = False Then Form1.Timer1.Enabled = True: Timer1.Enabled = False
'If displev > Len(curmsg) Then Exit Sub
'displev = displev + 6
'Label1.caption = Left(curmsg, displev)
'Label3.caption = Label1.caption
'Label1.Refresh
End Sub

Sub loadconv(convname As String, Optional picfile As String = "", Optional startpos As String = "MAIN", Optional ByVal convfile As String = "Conversations.txt")
If Not Left(convname, 1) = "#" Then Form10.caption = convname

If Dir(convfile) = "" Then convfile = getfile(convfile, "Data.pak")

lastchoice = 0
lastcom = 0
buying = 0
Form10.Cls
Erase commands()
Erase commandname()
If convname = "#WARP" Then loadwarps: Exit Sub
If convname = "#RANDOM" Then createconv: gotochoice "#RANDCONV": Exit Sub

Open convfile For Input As #1

Do While Not EOF(1)
    Input #1, durg
    
    If durg = "#CONVERSATION" Then
        Input #1, durg
        If durg = convname Then
            Do While Not durg = "#CONVERSATION" And Not EOF(1) 'Continue picking up choices until the next conversation is hit
                Input #1, durg2
                durg = getfromstring(durg2, 1)
                If Not durg = "" Then commands(lastchoice, lastcom, 0) = durg2
                If durg = "#BRANCH" Then lastchoice = lastchoice + 1: lastcom = 1: Input #1, commandname(lastchoice)
                If durg = "#SAY" Then commands(lastchoice, lastcom, 0) = durg2: Input #1, commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                If durg = "#TEXT" Then commands(lastchoice, lastcom, 0) = durg2: Input #1, commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                If durg = "#CHOICE" Then commands(lastchoice, lastcom, 0) = durg2: Input #1, commands(lastchoice, lastcom, 1), commands(lastchoice, lastcom, 2): lastcom = lastcom + 1
                If durg = "#EXIT" Then commands(lastchoice, lastcom, 0) = durg2: Input #1, commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                If durg = "#ADDCHOICES" Then commands(lastchoice, lastcom, 0) = durg2: Input #1, commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                If durg = "#BUY" Then Input #1, commands(lastchoice, lastcom, 1), commands(lastchoice, lastcom, 2): stdfilter commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                If durg = "#ITEM" Then Input #1, commands(lastchoice, lastcom, 1), commands(lastchoice, lastcom, 2): stdfilter commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                If lastcom > 0 Then stdfilter commands(lastchoice, lastcom - 1, 1)
                If durg = "#EFFECT" Then commands(lastchoice, lastcom, 0) = durg: Input #1, commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                If durg = "#IMAGE" Then commands(lastchoice, lastcom, 0) = durg2: Input #1, commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                If durg = "#QUEST" Then commands(lastchoice, lastcom, 0) = durg2: Input #1, commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                'If durg = "#CHOICE" Then Input #1, choice(lastchoice).choices(lastc2, 0), choice(lastchoice).choices(lastc2, 1): lastc2 = lastc2 + 1
                
            Loop
        End If
    End If
Loop

Close #1

If Not picfile = "" Then loaddapic picfile
'Me.Show
If startpos = "" Then startpos = "MAIN"
gotochoice startpos
'Show Me

End Sub

Function createconv()

addbranch "#RANDCONV", "#RUMOR", , , "#BLAB", "#RANDCONV", "#BLAB", "#RANDCONV", "#BLAB", "#RANDCONV", "#BLAB", "#RANDCONV"
addbranch "#GOODBYE", "Anyway/ I'm really/ really sick of talking to you.", , , "(Leave)", "#EXIT"
addbranch "#WANT", "#WANT", , , "Yes.", "#OFFER", "Not really.", "#GOODBYE"
addbranch "#OFFER", "(Put offer stuff here...)", , , "Yes.", "#OFFER", "Not really.", "#GOODBYE"

addbranch "#OFFICEROFFER", "#OFFICEROFFER"

End Function

Function createbranch(num)

End Function

Function addchoice2(wbranch As String, caption As String, dest As String)

For a = 0 To UBound(commandname())
    If commandname(a) = wbranch Then wbranchn = a: Exit For
Next a

For a = 1 To 15
    If commands(wbranchn, a, 0) = "" Then Exit For
Next a

If a > 15 Then Exit Function
commands(wbranchn, a, 0) = "#CHOICE"
commands(wbranchn, a, 1) = caption
commands(wbranchn, a, 2) = dest

End Function

Function addbranch(ByVal branchname As String, ByVal txt As String, Optional ByVal clear As Byte = 0, Optional ByVal wchoice = -1 _
, Optional ByVal choice1 As String, Optional ByVal choicedest1 As String _
, Optional ByVal choice2 As String, Optional ByVal choicedest2 As String _
, Optional ByVal choice3 As String, Optional ByVal choicedest3 As String _
, Optional ByVal choice4 As String, Optional ByVal choicedest4 As String _
, Optional ByVal choice5 As String, Optional ByVal choicedest5 As String _
, Optional ByVal choice6 As String, Optional ByVal choicedest6 As String)

Static lastchoice As Integer
If wchoice = -1 Then lastchoice = lastchoice + 1: wchoice = lastchoice
'lastchoice = lastchoice + 1: lastcom = 1: Input #1,

commandname(wchoice) = branchname

commands(wchoice, 15, 0) = "#SAY"
commands(wchoice, 15, 1) = txt
'commands(wchoice, 15, 2) = choicedest1

If clear = 1 Then
    For a = 0 To 14
    commands(wchoice, a, 0) = ""
    commands(wchoice, a, 1) = ""
    commands(wchoice, a, 2) = ""
    Next a
End If

If choice1 = "" Then Exit Function
commands(wchoice, 1, 0) = "#CHOICE"
commands(wchoice, 1, 1) = choice1
commands(wchoice, 1, 2) = choicedest1

If choice2 = "" Then Exit Function
commands(wchoice, 2, 0) = "#CHOICE"
commands(wchoice, 2, 1) = choice2
commands(wchoice, 2, 2) = choicedest2

If choice3 = "" Then Exit Function
commands(wchoice, 3, 0) = "#CHOICE"
commands(wchoice, 3, 1) = choice3
commands(wchoice, 3, 2) = choicedest3

If choice4 = "" Then Exit Function
commands(wchoice, 4, 0) = "#CHOICE"
commands(wchoice, 4, 1) = choice4
commands(wchoice, 4, 2) = choicedest4

If choice5 = "" Then Exit Function
commands(wchoice, 5, 0) = "#CHOICE"
commands(wchoice, 5, 1) = choice5
commands(wchoice, 5, 2) = choicedest5

If choice6 = "" Then Exit Function
commands(wchoice, 6, 0) = "#CHOICE"
commands(wchoice, 6, 1) = choice6
commands(wchoice, 6, 2) = choicedest6

End Function

Function findchoice(choicename, Optional choicenum = -1) As Byte
'Processes all the choice stuff

choicename2 = getfromstring(choicename, 1)
If choicename = "##EXIT##" Then Me.Hide ': Form1.Timer1.Enabled = True
If choicename = "#EXIT" Then Me.Hide ': Form1.Timer1.Enabled = True
If choicename2 = "#SELLCLOTHES" Then buyclothes , , getfromstring(choicename, 2), getfromstring(choicename, 3), getfromstring(choicename, 4): Exit Function
If choicename2 = "#SELLMAGICCLOTHES" Then buyclothes , , getfromstring(choicename, 2), getfromstring(choicename, 3), getfromstring(choicename, 4), 1: Exit Function
If choicename2 = "#SELLPOTIONS" Then buyclothes "POTIONS": Exit Function
If choicename2 = "#SELLGEMS" Then buyclothes "GEMS": Exit Function
If choicename2 = "#SELLARMOR" Then buyclothes "ARMOR", Val(getfromstring(choicename, 2)), getfromstring(choicename, 3): Exit Function
If choicename2 = "#SELLMAGICARMOR" Then buyclothes "ARMOR", Val(getfromstring(choicename, 2)), getfromstring(choicename, 3), , , 1: Exit Function
If choicename2 = "#SELLWEAPONS" Then buyclothes "WEAPON", Val(getfromstring(choicename, 2)), getfromstring(choicename, 3): Exit Function
If choicename2 = "#SELLMAGICWEAPONS" Then buyclothes "WEAPON", Val(getfromstring(choicename, 2)), getfromstring(choicename, 3), , , 1: Exit Function

If choicename2 = "##ITEM" Then takeobj 0, 0, 0, getobjtype(getfromstring(choicename, 2)): choicename = getfromstring(choicename, 3) ': Exit Function
If choicename2 = "#ITEM" Then takeobj 0, 0, 0, getobjtype(getfromstring(choicename, 2)): choicename = getfromstring(choicename, 3) ': Exit Function

If choicename2 = "#RANDCONV" Then
    If roll(8) = 1 Then choicename = "#GOODBYE"
    'If roll(5) = 1 Then choicename = "#OFFER"
    'If roll(5) = 1 Then choicename = "#WANT"
End If

If getfromstring(choicename, 1) = "##BUY" Then
    If plr.gp < getfromstring(choicename, 3) Then
    disptext "You do not have enough gold."
    Else:
    takeobj 0, 0, 0, choiceobjs(getfromstring(choicename, 2))
    plr.gp = plr.gp - getfromstring(choicename, 3)
    disptext "Thank you."
    End If
    Exit Function
End If

If getfromstring(choicename, 1) = "##BUY2" Then
    If plr.gp < getfromstring(choicename, 3) Then
    disptext "You do not have enough gold."
    Else:
    takeobj 0, 0, 0, getobjtype(getfromstring(choicename, 2))
    plr.gp = plr.gp - getfromstring(choicename, 3)
    disptext "Thank you."
    End If
End If

If choicename = "#BUYLIFE" Then
    num = InputBox("How many?")
    If isexpansion = 0 Then
        If plr.gp < num * potioncost(plr.lpotionlev) Then
        disptext "You do not have enough gold for that many."
        Else:
        plr.lpotions = plr.lpotions + num
        plr.gp = plr.gp - num * potioncost(plr.lpotionlev)
        disptext "Thank you."
        End If
    End If
    
    If isexpansion = 1 And Not num = "" Then
        If plr.gp < num * potioncost(plr.level * 10) Then
        disptext "You do not have enough gold for that many."
        Else:
        plr.lpotions = plr.lpotions + num
        plr.gp = plr.gp - num * potioncost(plr.level * 10)
        disptext "Thank you."
        End If
    End If
    
    
    Exit Function
End If

If choicename = "#BUYMANA" Then
    num = InputBox("How many?")
    If plr.gp < num * potioncost(plr.mpotionlev) Then
    disptext "You do not have enough gold for that many."
    Else:
    plr.mpotions = plr.mpotions + num
    plr.gp = plr.gp - num * potioncost(plr.mpotionlev)
    disptext "Thank you."
    Exit Function
    End If
End If

If choicename2 = "#DESTROYME" Then
    killobj objs(talkingtonum)
    choicename = getfromstring(choicename, 2)
End If

If choicename2 = "#CHANGESTART" Then
    changeeff getobjtnum(talkingto), "Conversation", , , getfromstring(choicename, 2)
    choicename = getfromstring(choicename, 3)
End If

If choicename2 = "#EFFECT" Then
    cleareffs dummyobj2
    addeffect2 dummyobj2, getfromstring(choicename, 3), getfromstring(choicename, 4), getfromstring(choicename, 5), getfromstring(choicename, 6), getfromstring(choicename, 7)
    takeobj 0, 0, 0, dummyobj2
    choicename = getfromstring(choicename, 2)
End If

If choicename2 = "#TAKEOBJ" Then
    takeobj 0, 0, 0, dummyobj2
    choicename = getfromstring(choicename, 2)
End If

If choicename2 = "#EATME" Then
    eatobj talkingtonum, 1: Unload Me
End If

If choicename2 = "#DAMAGE" Then
    plrdamage getfromstring(choicename, 3)
    choicename = getfromstring(choicename, 2)
End If

If choicename2 = "#BASEEQUIP" Then
    baseequip
    choicename = getfromstring(choicename, 2)
End If

If choicename2 = "#DIGESTTO" Then
    plr.diglevel = getfromstring(choicename, 3)
    choicename = getfromstring(choicename, 2)
    If plr.diglevel >= 4 Then
        plr.diglevel = 6
        Do While isnaked = False
        digestclothes 255
        Loop
        createobjtype "Shit", "poo1small.bmp": createobj "Shit", plr.X, plr.Y: plr.instomach = 0: stopsounds: playsound "fart57.wav": wep.dice = 0: wep.graphname = "": plr.plrdead = 1
        gamemsg "You have been utterly digested."
        Unload Me
    End If
    Form1.updatbody
End If

If choicename2 = "#GIVESKILL" Then
   addskill (getfromstring(choicename, 2))
   choicename = getfromstring(choicename, 3)
End If

If choicename2 = "#ADDQUEST" Then
    ifquest getfromstring(choicename, 2), 1
    choicename = getfromstring(choicename, 3)
End If



For a = 0 To 100
    If commandname(a) = choicename And choicenum > -1 Then
    If ifreq(choicenum) = "SAYONCE" Then addsaid choice(choicenum).caption & talkingto
    End If
    If commandname(a) = choicename Then findchoice = a: Exit Function
    'If choice(a).choicename = choicename Then findchoice = a: Exit Function
Next a

If a = 101 And Not choicename = "MAIN" Then findchoice = findchoice("MAIN")

End Function

Function gotochoice(choicename)
curchoice = findchoice(choicename): updatchoice
End Function

Function updatchoice()

For a = 0 To 7
    Label2(a).Visible = False
Next a

cleareffs dummyobj2

For a = 1 To 15
zerf = getfromstring(commands(curchoice, a, 0), 1)

'Check for conversation requirements, if any
req = getfromstring(commands(curchoice, a, 0), 2)
If req = "" Then GoTo 3
If req = "SAYONCE" Then If ifsaid(commands(curchoice, a, 1) & talkingto) Then GoTo 8
If req = "IFSAID" Then If Not ifsaid(getfromstring(commands(curchoice, a, 0), 3)) Then GoTo 8
If req = "IFQUEST" Then If Not ifquest(getfromstring(commands(curchoice, a, 0), 3)) = True Then GoTo 8
If req = "IFNOTQUEST" Then If ifquest(getfromstring(commands(curchoice, a, 0), 3)) = True Then GoTo 8
If req = "DEX" Then If plr.dex < getfromstring(commands(curchoice, a, 0), 3) Then GoTo 8
If req = "STR" Then If plr.str < getfromstring(commands(curchoice, a, 0), 3) Then GoTo 8
If req = "INT" Then If plr.int < getfromstring(commands(curchoice, a, 0), 3) Then GoTo 8
If req = "NOTDEX" Then If plr.dex >= getfromstring(commands(curchoice, a, 0), 3) Then GoTo 8
If req = "NOTSTR" Then If plr.str >= getfromstring(commands(curchoice, a, 0), 3) Then GoTo 8
If req = "NOTINT" Then If plr.int >= getfromstring(commands(curchoice, a, 0), 3) Then GoTo 8


3
If zerf = "#IMAGE" Then loaddapic commands(curchoice, a, 1)
If zerf = "#SAY" Then disptext "'" & commands(curchoice, a, 1) & "'": Label1.ForeColor = RGB(255, 255, 255)
If zerf = "#TEXT" Then disptext commands(curchoice, a, 1): Label1.ForeColor = RGB(150, 150, 250)
If zerf = "#ADDCHOICES" Then addchoices commands(curchoice, a, 1)
If zerf = "#CHOICE" Then addchoice commands(curchoice, a, 2), commands(curchoice, a, 1), commands(curchoice, a, 0)
If zerf = "#QUEST" Then addquest commands(curchoice, a, 1)

If zerf = "#EFFECT" Then addeffect2 dummyobj2, getfromstring(commands(curchoice, a, 1), 1), getfromstring(commands(curchoice, a, 1), 2), getfromstring(commands(curchoice, a, 1), 3), getfromstring(commands(curchoice, a, 1), 4), getfromstring(commands(curchoice, a, 1), 5)

If zerf = "#EXIT" Then addchoice "##EXIT##", commands(curchoice, a, 1), commands(curchoice, a, 0)
If zerf = "#ITEMCHOICE" Then addchoice "##ITEM:" & commands(curchoice, a, 2), commands(curchoice, a, 1), commands(curchoice, a, 0)
If zerf = "#BUY" Then addchoice "##BUY2:" & commands(curchoice, a, 2), commands(curchoice, a, 1), commands(curchoice, a, 0): buying = 1
If zerf = "#ITEM" Then addchoice "##ITEM:" & commands(curchoice, a, 2), commands(curchoice, a, 1), commands(curchoice, a, 0)


8 Next a

If Label2(0).Visible = False Then addchoice "##EXIT##", "(Exit)", ""

End Function

Function addchoice(choicen, ByVal caption, Optional command = "")

If caption = "#BLAB" Then caption = getblab

For a = 0 To 7
    If Label2(a).Visible = False Then
    If Left(choicen, 6) = "##BUY2" Then buying = 1
    choice(a).command = command
    choice(a).caption = caption
    choice(a).choicename = choicen
    Label2(a).caption = caption
    If Len(caption) > 90 Then Label2(a).FontSize = 8 Else Label2(a).FontSize = 12
    Label2(a).Visible = True
    Exit Function
    End If
Next a

End Function

Function addchoices(choicen)
For a = 0 To 100
    If commandname(a) = choicen Then
        For b = 0 To 15
            If commands(a, b, 0) = "#CHOICE" Then addchoice commands(a, b, 2), commands(a, b, 1)
            If commands(a, b, 0) = "#EXIT" Then addchoice commands(a, b, 2), commands(a, b, 1), "#EXIT"
        Next b
        Exit Function
    End If
Next a

End Function

Function buyclothes(Optional wtype As String = "", Optional worth As Byte = 0, Optional colorname As String = "", Optional wear1 As String = "", Optional wear2 As String = "", Optional magical = 0)

'obr = roll(255)
'obg = roll(255)
'obb = roll(255)

Dim dong As objecttype

If wtype = "CLOTHES" Then wtype = "" 'Clothes are default anyway

If plr.difficulty = 0 Then randword talkingto Else randword plr.name & talkingto, plr.difficulty

If worth = 0 Then worth = Int(lowestlevel) ': If worth > 1 Then worth = worth + 1

'worth = 4

buying = 1
Label1.ForeColor = RGB(255, 255, 255)

disptext "What can I interest you in?"

clearchoices

cost = 2
colorname = gencolor(colorname, obr, obg, obb, obl)

If wtype = "CARGOTRADE" Then Cargotrade.Show: Unload Me
wleft = getfromstring(wtype, 1)
If wleft = "PARTS" Then buyparts wtype: Exit Function

For a = 0 To 6
4
'Clear Object
choiceobjs(a) = dong
If wtype = "POTIONS" Then
    addchoice "#BUYLIFE", "Life Potions (" & potioncost(plr.lpotionlev) & " GP Each)"
    addchoice "#BUYMANA", "Mana Potions (" & potioncost(plr.mpotionlev) & " GP Each)"
    Exit For
    End If
If wtype = "" Then choiceobjs(a) = randclothes(colorname, obr, obg, obb, obl, worth, cost, wear1, wear2)
If wtype = "ARMOR" Then choiceobjs(a) = randarmor(worth, cost, 1) ', Val(colorname))
If wtype = "WEAPON" Or wtype = "WEAPONS" Then
    Select Case a
        Case 0: weapontype = "SWORD"
        Case 1: weapontype = "SPEAR"
        Case 2: weapontype = "AXE"
        Case 3: weapontype = "MACE"
        Case 4: weapontype = "STAFF"
        Case 5: weapontype = "BOW"
        Case 5: weapontype = "FAST"
        Case Else:
        weapontype = ""
    End Select
    choiceobjs(a) = randwep(worth, cost, weapontype)
    
End If
If wtype = "GEMS" Then makegem choiceobjs(a), worth: cost = getworth(choiceobjs(a))
If wleft = "CARGO" Then
    cargtype = getfromstring(wtype, 2)
    amt = ((a + 1) * (a + 1)) * 5
    choiceobjs(a).name = a
    addeffect2 choiceobjs(a), "Cargo", cargtype, amt
    cost = getcargoprice(cargtype) * amt / 2
End If

If wleft = "PARTS" Then
    Dim partstype As String
    partstype = getfromstring(wtype, 2)
    partstype = filestr(partstype)
    Dim zonk As systype
    zonk = loadpart(partstype)
    choiceobjs(a).name = partstype
    addeffect2 choiceobjs(a), "GivePart", choiceobjs(a).name
    cost = 10
End If

    'Prevent Repeats
    For b = 0 To a
        retries = retries + 1
        If retries > 100 Then Exit For
        If b = a Then GoTo 6
        If choiceobjs(b).name = choiceobjs(a).name Then GoTo 4
6     Next b
If magical > 0 Then makemagicitem choiceobjs(a), worth - 1 + magical: cost = getworth(choiceobjs(a)) * 2
If wtype = "ARMOR" Or wtype = "" Then addchoice "##BUY:" & a & ":" & cost, "I would like the " & choiceobjs(a).name & "  (AR " & choiceobjs(a).effect(1, 3) & ", Weight " & Val(geteff(choiceobjs(a), "Equipjunk", 2)) & ", " & cost & " Gold Pieces)": GoTo 8
If wtype = "WEAPON" Or wtype = "WEAPONS" Then addchoice "##BUY:" & a & ":" & cost, "I would like the " & choiceobjs(a).name & "  (" & choiceobjs(a).effect(1, 3) & "-" & Val(choiceobjs(a).effect(1, 3)) * Val(choiceobjs(a).effect(1, 4)) & " Damage, Weight " & geteff(choiceobjs(a), "Equipjunk", 2) & ", " & cost & " Gold Pieces)": GoTo 8
If Not choiceobjs(a).name = "" Then addchoice "##BUY:" & a & ":" & cost, "I would like the " & choiceobjs(a).name & " (" & displayobj(choiceobjs(a), 0) & ")": GoTo 8
'addchoice "Howdy.", "Howdy."
8 Next a

addchoice "##EXIT##", "(Leave)"

End Function


Function clearchoices()
For a = 0 To 7
    Label2(a).Visible = False
Next a
End Function

Function ifreq(choicenum)

ifreq = getfromstring(choice(choicenum).command, 2)
'Stop
'For a = 0 To 100
'    For b = 0 To 15
'        zark = commands(a, b, 0)
'        zark = getfromstring(zark, 2)
'        If Not zark = "" Then
'    Next b
'Next a

End Function

Sub loadwarps()

For a = 1 To 8
If Not plr.beento(a, 1) = "" Then addchoice "#CHOICE", "Warp to " & plr.beento(a, 1), "#WARP:" & plr.beento(a, 2)
Next a

End Sub

Sub continue()
'If Form10.Visible = False Then Form1.Timer1.Enabled = True: Timer1.Enabled = False
If displev > Len(curmsg) Then Exit Sub
displev = displev + 6
Label1.caption = Left(curmsg, displev)
Label3.caption = Label1.caption
Label1.Refresh
End Sub

Function buyparts(wtype, Optional picfile = "")

buying = 0

wleft = getfromstring(wtype, 1)
Dim partstype As String
partstype = getfromstring(wtype, 2)
'Example: PARTS:SHIELDPARTS, or whatever
If wleft = "PARTS" Then
    addbranch "MAIN", "Here are the parts I sell.  I'd be more than willing to tell you about them.", 1, 0
    If Not Left(partstype, 2) = "#$" Then partstype = "#$" & partstype
    For a = 0 To 6
    partstype = getfromstring(wtype, 2)
    partstype = filestr(partstype)
    Dim zonk As systype
    zonk = loadpart(partstype)
    choiceobjs(a).name = partstype
    addeffect2 choiceobjs(a), "GivePart", choiceobjs(a).name
    addbranch "BUY" & zonk.name, zonk.desc, , , "(Buy this)", "##BUY:" & a & "10", "(Back)", "MAIN"
    addchoice2 "MAIN", zonk.name, "BUY" & zonk.name
    cost = 10
    Next a
End If

addchoice2 "MAIN", "(Leave)", "#EXIT"
gotochoice "MAIN"
addchoice 8, "EXIT", "#EXIT"

End Function


Sub addtoconv(Optional command = "#CHOICE", Optional val1 = "Choice Caption", Optional val2 = "BRANCHTO")
Form10.Cls
'Erase commands()
'Erase commandname()

                durg2 = command
                durg = getfromstring(durg2, 1)
                If Not durg = "" Then commands(lastchoice, lastcom, 0) = durg2
                If durg = "#BRANCH" Then lastchoice = lastchoice + 1: lastcom = 1: commandname(lastchoice) = val1
                If durg = "#SAY" Then commands(lastchoice, lastcom, 0) = durg2: commands(lastchoice, lastcom, 1) = val1: lastcom = lastcom + 1
                If durg = "#TEXT" Then commands(lastchoice, lastcom, 0) = durg2: commands(lastchoice, lastcom, 1) = val1: lastcom = lastcom + 1
                If durg = "#CHOICE" Then commands(lastchoice, lastcom, 0) = durg2: commands(lastchoice, lastcom, 1) = val1: commands(lastchoice, lastcom, 2) = val2: lastcom = lastcom + 1
                If durg = "#EXIT" Then commands(lastchoice, lastcom, 0) = durg2: commands(lastchoice, lastcom, 1) = val1: lastcom = lastcom + 1
                If durg = "#ADDCHOICES" Then commands(lastchoice, lastcom, 0) = durg2: commands(lastchoice, lastcom, 1) = val1: lastcom = lastcom + 1
                If durg = "#BUY" Then commands(lastchoice, lastcom, 1) = val1: commands(lastchoice, lastcom, 2) = val2: stdfilter commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                If durg = "#ITEM" Then commands(lastchoice, lastcom, 1) = val1: commands(lastchoice, lastcom, 2) = val2: stdfilter commands(lastchoice, lastcom, 1): lastcom = lastcom + 1
                If lastcom > 0 Then stdfilter commands(lastchoice, lastcom - 1, 1)
                If durg = "#EFFECT" Then commands(lastchoice, lastcom, 0) = durg: commands(lastchoice, lastcom, 1) = val1: lastcom = lastcom + 1
                If durg = "#IMAGE" Then commands(lastchoice, lastcom, 0) = durg2: commands(lastchoice, lastcom, 1) = val1: lastcom = lastcom + 1
                'If durg = "#CHOICE" Then Input #1, choice(lastchoice).choices(lastc2, 0), choice(lastchoice).choices(lastc2, 1): lastc2 = lastc2 + 1


'If Not picfile = "" Then loaddapic picfile
'Me.Show
'If startpos = "" Then startpos = "MAIN"
'gotochoice startpos
'Show Me

End Sub



