VERSION 5.00
Begin VB.Form Bridge 
   BackColor       =   &H00000000&
   Caption         =   "Bridge"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9480
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   596
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   796
      TabIndex        =   1
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command2 
         Caption         =   "Exit Bridge"
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Bridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Command2_Click()
genstarshipmap "Random", 4, 4, 4, 3, 1, 26
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
spaceon = 0
nodraw = 0
End Sub

Private Sub Timer1_Timer()
If spaceon = 1 Then drawships
End Sub

Private Sub Form_Resize()
Picture1.Width = Bridge.ScaleWidth
Picture1.Height = Bridge.ScaleHeight
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyUp Then calcaccel plrship.facing * 20, 1, 5, plrship.xspeed, plrship.yspeed
'If KeyCode = vbKeyLeft Then plrship.facing = plrship.facing + 1
'If KeyCode = vbKeyRight Then plrship.facing = plrship.facing - 1
'If plrship.facing = 0 Then plrship.facing = 18
'If plrship.facing = 19 Then plrship.facing = 1
KeyCode = 0
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
starmapx = X * (800 / Picture1.ScaleWidth)
starmapy = Y * (600 / Picture1.ScaleHeight)
If starmapon = 1 Then drawstarmap
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cheaton = 1 Then plrship.X = X * (800 / Picture1.ScaleWidth) * 833: plrship.Y = Y * (600 / Picture1.ScaleHeight) * 833
End Sub
