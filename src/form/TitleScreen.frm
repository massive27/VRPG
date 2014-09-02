VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TitleScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duamutef's Glorious Vore RPG"
   ClientHeight    =   8940
   ClientLeft      =   1665
   ClientTop       =   1050
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog FileD 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   415
      Left            =   3600
      TabIndex        =   2
      Top             =   8520
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   5655
      Left            =   6840
      TabIndex        =   1
      Top             =   3120
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   4935
   End
End
Attribute VB_Name = "TitleScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.caption = "Duamutef's Glorious Vore RPG, V" & curversion
Label3.caption = "Version " & curversion
'If isexpansion = 1 Then Label3.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
Do While (1)
'm'' redirect for cleaner ending
Debugger.Quitting 'm''
End
Loop
End If
End Sub

Private Sub Label1_Click()
If Not Dir("autosave.plr") = "" Then Kill "autosave.plr"
Unload Me
End Sub

Private Sub Label2_Click()

loadgame = 1
Load Form1 'm'' let form1 preloads here
loadclick FileD

DoEvents 'm'' refresh the screen to remove the ugly cmdialog ghosted hdc

Me.Hide
Form1.Show

'floadchar FileD.FileTitle

Unload Me
Exit Sub
'On Error GoTo 5
'GoTo 10
'5 Exit Sub
10

'FileD.FileName = App.Path & "\" & "*.plr"
'FileD.DefaultExt = "plr"
'FileD.ShowOpen
'Unload Me
'Form1.Timer1.Enabled = False

ChDir App.Path

FileD.FileName = "*.plr"
FileD.DefaultExt = "plr"
FileD.ShowOpen

If Dir(FileD.FileName) = "" Then Exit Sub

fname = getfile("plrdat.tmp", FileD.FileName)

getfile "curgame.dat", FileD.FileName, , 1

floadchar fname ' FileD.FileName

'Kill "plrdat.tmp"

'If FileD.FileTitle = "" Or Dir(FileD.FileTitle) = "" Then Exit Sub
Me.Hide
Form1.Show

Add_UI.RemOldUI

'floadchar FileD.FileTitle

Unload Me

End Sub
