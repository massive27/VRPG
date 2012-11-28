VERSION 5.00
Begin VB.Form Shipsys 
   Caption         =   "Starship System"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   ">>>"
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   27
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<<"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   26
      Top             =   480
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4695
      Left            =   600
      Max             =   10
      TabIndex        =   22
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   17
      Left            =   8760
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   21
      Top             =   6840
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   16
      Left            =   7200
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   20
      Top             =   6840
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   15
      Left            =   5640
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   19
      Top             =   6840
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   14
      Left            =   4080
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   18
      Top             =   6840
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   13
      Left            =   2520
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   17
      Top             =   6840
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   12
      Left            =   960
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   16
      Top             =   6840
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   11
      Left            =   8760
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   15
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   10
      Left            =   7200
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   14
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   9
      Left            =   5640
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   13
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   8
      Left            =   4080
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   12
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   7
      Left            =   2520
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   11
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   6
      Left            =   960
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   10
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   5
      Left            =   8760
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   4
      Left            =   7200
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   3
      Left            =   5640
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   2
      Left            =   4080
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   1
      Left            =   2520
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   0
      Left            =   960
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   3
      Left            =   5640
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   2
      Left            =   4080
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   1
      Left            =   2520
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   0
      Left            =   960
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   7200
      TabIndex        =   25
      Top             =   240
      Width           =   4815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SHIP'S STOCK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4725
      TabIndex        =   24
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3330
      TabIndex        =   23
      Top             =   360
      Width           =   1515
   End
End
Attribute VB_Name = "Shipsys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sysnum As Byte
Private grabnum As Byte

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then setsysnum sysnum - 1 Else setsysnum sysnum + 1
End Sub

Private Sub Form_Load()
'partsupdat
If sysnum = 0 Then sysnum = 1
partsupdat
End Sub

Sub setsysnum(systemnum As Byte)
sysnum = systemnum
If sysnum < 1 Then sysnum = 1
If sysnum > UBound(plrshipdat.sections()) Then sysnum = UBound(plrshipdat.sections())
partsupdat
End Sub

Sub partsupdat()

zarp = VScroll1.Value * 6

Label1.caption = plrshipdat.sections(sysnum).type & " (Bay " & sysnum & ")"

For a = 1 To 4
    If Not plrshipdat.sections(sysnum).partsobj(a).graphname = "" Then Picture1(a - 1).Picture = LoadPicture(plrshipdat.sections(sysnum).partsobj(a).graphname) Else Picture1(a - 1).Picture = LoadPicture("")
Next a

For a = 1 To 18 'UBound(partstock())
    Picture2(a - 1).Visible = False
    If a + zarp > UBound(partstock()) Then GoTo 5
    If Not partstock(a + zarp).graphname = "" Then Picture2(a - 1).Picture = LoadPicture(partstock(a + zarp).graphname): Picture2(a - 1).Visible = True
5 Next a

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
setplrship
pause = 0
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Label3.caption = plrshipdat.sections(sysnum).partsobj(Index + 1).name & vbCrLf & "(" & plrshipdat.sections(sysnum).partsobj(Index + 1).slottype & ")" & vbCrLf & plrshipdat.sections(sysnum).partsobj(Index + 1).desc

End Sub

Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If partstock(grabnum).slottype = plrshipdat.sections(sysnum).partsobj(Index + 1).slottype Then
If Button = 1 Then
If grabnum = 0 Then grabnum = 1
    partswap plrshipdat.sections(sysnum).partsobj(Index + 1), partstock(grabnum)
'End If
End If

If Button = 2 Then
    For a = 1 To UBound(partstock())
    If partstock(a).name = "" Then Exit For
    Next a
    grabnum = a
    partswap plrshipdat.sections(sysnum).partsobj(Index + 1), partstock(grabnum)
End If

partsupdat
End Sub

Private Sub Picture2_Click(Index As Integer)
grabnum = Index + (VScroll1.Value * 6) + 1
End Sub

Private Sub Picture2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
zork = Index + (VScroll1.Value * 6) + 1

Label3.caption = partstock(zork).name & vbCrLf & "(" & partstock(zork).slottype & ")" & vbCrLf & partstock(zork).desc

End Sub

Private Sub VScroll1_Change()
partsupdat
End Sub
