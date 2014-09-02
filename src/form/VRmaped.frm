VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   Caption         =   "Map Editor"
   ClientHeight    =   10440
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   16395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   ScaleHeight     =   696
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1093
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "<<<<<"
      Height          =   1695
      Index           =   3
      Left            =   120
      TabIndex        =   32
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "\/ \/ \/ \/ \/"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   31
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>>>>"
      Height          =   1695
      Index           =   1
      Left            =   10680
      TabIndex        =   30
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "/\ /\ /\ /\ /\"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   29
      Top             =   0
      Width           =   2055
   End
   Begin VB.PictureBox Picture5 
      Height          =   1335
      Left            =   13560
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   28
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog FileD 
      Left            =   10080
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   8160
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6615
      Left            =   120
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   27
      Top             =   9120
      Visible         =   0   'False
      Width           =   10095
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   12000
      Max             =   50
      Min             =   1
      TabIndex        =   25
      Top             =   2640
      Value           =   1
      Width           =   2295
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   9480
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   24
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   12720
      TabIndex        =   22
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   12720
      TabIndex        =   20
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   12720
      TabIndex        =   18
      Top             =   2880
      Width           =   2895
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   12000
      Max             =   30
      TabIndex        =   15
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-->"
      Height          =   255
      Index           =   1
      Left            =   12000
      TabIndex        =   13
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<--"
      Height          =   255
      Index           =   0
      Left            =   11280
      TabIndex        =   12
      Top             =   4320
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   11280
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   11
      Top             =   4560
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Monsters (Click to add, right-click to delete)"
      Height          =   375
      Index           =   2
      Left            =   11280
      TabIndex        =   10
      Top             =   1320
      Width           =   4335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Objects (Click to select, shift-click to create, right-click to edit"
      Height          =   375
      Index           =   1
      Left            =   11280
      TabIndex        =   9
      Top             =   840
      Width           =   4455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Map Tiles (Shift for overlays, right-click to erase)"
      Height          =   255
      Index           =   0
      Left            =   11280
      TabIndex        =   8
      Top             =   480
      Value           =   -1  'True
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7920
      TabIndex        =   6
      Text            =   "0"
      Top             =   7560
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   255
      Left            =   9600
      TabIndex        =   5
      Top             =   7200
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   255
      Left            =   9840
      TabIndex        =   4
      Top             =   7200
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6615
      Left            =   10440
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   6840
      Width           =   10095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6615
      Left            =   360
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   0
      Top             =   240
      Width           =   10095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Tile Brush Width"
      Height          =   255
      Left            =   12480
      TabIndex        =   26
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "(Click on picture to edit object or monster types)"
      Height          =   255
      Left            =   11760
      TabIndex        =   23
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Unique Value 2:"
      Height          =   255
      Index           =   2
      Left            =   11280
      TabIndex        =   21
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Unique Value 1:"
      Height          =   255
      Index           =   1
      Left            =   11280
      TabIndex        =   19
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Unique Name:"
      Height          =   255
      Index           =   0
      Left            =   11280
      TabIndex        =   17
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Brush Noise: 0"
      Height          =   255
      Left            =   12240
      TabIndex        =   16
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Monster Name"
      Height          =   195
      Left            =   11280
      TabIndex        =   14
      Top             =   4080
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Tile Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Menu filemen 
      Caption         =   "File"
      Begin VB.Menu savemen 
         Caption         =   "Save Map"
         Shortcut        =   ^S
      End
      Begin VB.Menu loadmen 
         Caption         =   "Load Map"
         Shortcut        =   ^O
      End
      Begin VB.Menu newmen 
         Caption         =   "New Map"
         Shortcut        =   ^N
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zoom
Dim blockx
Dim blocky
Private SelectedObject

Function updatmap()

Picture4.Cls

For a = 1 To mapx
For b = 1 To mapy
    tilet = map(a, b).tile '- 1
    If tilet = 0 Then tilet = 1
    'getrgb TileColors(tilet), r1, g1, b1
    Picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), TileColors(tilet), BF
    'If map(a, b).tile = 1 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(0, 150, 0), BF
    'If map(a, b).tile = 2 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(250, 250, 0), BF
    'If map(a, b).tile = 3 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(150, 150, 150), BF
    'If map(a, b).tile = 4 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(175, 150, 0), BF
    'If map(a, b).tile = 5 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(150, 100, 0), BF
    'If map(a, b).tile = 6 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(150, 0, 150), BF
    'If map(a, b).tile = 7 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(250, 220, 0), BF
    'If map(a, b).tile = 8 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(0, 0, 250), BF
    'If map(a, b).tile = 9 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(30, 150, 150), BF
    'If map(a, b).tile = 10 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(150, 120, 0), BF
    'If map(a, b).tile = 11 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(150, 80, 0), BF
    'If map(a, b).tile = 12 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(80, 60, 0), BF
    'If map(a, b).tile = 13 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(120, 110, 20), BF
    'If map(a, b).tile = 14 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(40, 40, 40), BF
    'If map(a, b).tile = 15 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(140, 140, 110), BF
    'If map(a, b).tile = 16 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(110, 220, 190), BF
    'If map(a, b).tile = 17 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(200, 200, 0), BF
    'If map(a, b).tile = 18 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(255, 150, 0), BF
    'If map(a, b).tile > 18 Then
    '    tilenum = map(a, b).tile
    '
    '    picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), TileColors(map(a, b).tile), BF
    'End If
        
    'If map(a, b).ovrtile = 0 Then map(a, b).blocked = 0
    If map(a, b).ovrtile > 0 Then Picture4.Line (a * zoom, b * zoom)-Step(zoom - 1, zoom - 1), RGB(150, 150, 150), B: map(a, b).blocked = 1
    'If map(a, b).ovrtile = 1 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(150, 150, 150), B: map(a, b).blocked = 1
    'If map(a, b).ovrtile = 2 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(150, 150, 0), B: map(a, b).blocked = 1
    'If map(a, b).ovrtile = 3 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(250, 50, 0), B: map(a, b).blocked = 1
    'If map(a, b).ovrtile = 4 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(250, 150, 0), B: map(a, b).blocked = 1
    'If map(a, b).ovrtile = 5 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(80, 80, 80), B: map(a, b).blocked = 0
'    If map(a, b).ovrtile = 1 Then picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(150, 150, 150), B: map(a, b).blocked = 1
    
    If map(a, b).monster > 0 Then Picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(255, 0, 0), BF
    
    If map(a, b).object > 0 Then Picture4.Line (a * zoom, b * zoom)-Step(zoom, zoom), RGB(255, 255, 0), BF
    
Next b
Next a

Picture1.PaintPicture Picture4.Image, 1, 1

Label1.caption = "Tile:" & edittile

End Function

Private Sub Command1_Click()
zoom = zoom + 1
End Sub

Private Sub Command2_Click()
zoom = zoom - 1
End Sub

Sub updatedittile()
        'Select Case edittile:
        If edittile < 26 Then Picture3.Picture = LoadPicture(getfile("tiles1.bmp"))
        If edittile > 25 And edittile <= 50 Then Picture3.Picture = LoadPicture(getfile("tiles2.bmp"))
        If edittile = 51 Then Picture2.Picture = LoadPicture(getfile(mapjunk.name1)): Exit Sub
        If edittile = 52 Then Picture2.Picture = LoadPicture(getfile(mapjunk.name2)): Exit Sub
        If edittile = 53 Then Picture2.Picture = LoadPicture(getfile(mapjunk.name3)): Exit Sub
        If edittile = 54 Then Picture2.Picture = LoadPicture(getfile(mapjunk.name4)): Exit Sub
            'Case 25: Picture3.Picture = LoadPicture(getfile("tiles1.bmp"))
            'Case 2: Picture3.Picture = LoadPicture(getfile("tiles1.bmp"))
            'Case 25: Picture3.Picture = LoadPicture(getfile("tiles1.bmp"))
            'Case 26: Picture3.Picture = LoadPicture(getfile("tiles2.bmp"))
            'Case 50: Picture3.Picture = LoadPicture(getfile("tiles2.bmp"))
        'End Select
        x = ((edittile Mod 25 - 1) Mod 5) * 96 + 1
        y = Int((edittile Mod 25 - 1) / 5) * 52 + 1
        Picture2.Width = 96
        Picture2.Height = 52
        
        Picture2.PaintPicture Picture3.Picture, 0, 0, 96, 52, x, y, 96, 52
        
        Picture3.Picture = LoadPicture(getfile("overlays2.bmp"))
        x = ((edittile Mod 25 - 1) Mod 5) * 96 + 1
        y = Int((edittile Mod 25 - 1) / 5) * 144 + 1
        Picture5.Width = 96
        Picture5.Height = 144
        
        Picture5.Visible = True
        Picture5.PaintPicture Picture3.Picture, 0, 0, 96, 144, x, y, 96, 144
        
    

End Sub

Private Sub Command3_Click(Index As Integer)
    
If CurEdLayer = EdTiles Then
    Select Case Index
        Case 0: If edittile <= 1 Then edittile = 54 Else edittile = edittile - 1
        Case 1: If edittile >= 54 Then edittile = 1 Else edittile = edittile + 1
    End Select
    updatedittile
    
End If
    
If CurEdLayer = EdMonsters Then
    
    Select Case Index
        Case 0: If CurEdMonster <= 1 Then CurEdMonster = UBound(montype()) Else CurEdMonster = CurEdMonster - 1
        Case 1: If CurEdMonster >= UBound(montype()) Then CurEdMonster = 1 Else CurEdMonster = CurEdMonster + 1
    End Select
    Label3.caption = "Monster Name: " & montype(CurEdMonster).name
    Picture2.Picture = LoadPicture(montype(CurEdMonster).gfile)

End If

If CurEdLayer = EdObjects Then
    
    Select Case Index
        Case 0: If CurEdObject <= 1 Then CurEdObject = UBound(objtypes()) Else CurEdObject = CurEdObject - 1
        Case 1: If CurEdObject >= UBound(objtypes()) Then CurEdObject = 1 Else CurEdObject = CurEdObject + 1
    End Select
    Label3.caption = "Object Name: " & objtypes(CurEdObject).name
    Picture2.Picture = LoadPicture(getfile(objtypes(CurEdObject).graphname))

End If




End Sub

Private Sub Command4_Click(Index As Integer)
    
    FileD.FileName = "*.txt;*.map"
    FileD.ShowOpen
    mapjunk.maps(Index + 1) = FileD.FileTitle

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyAdd Then edittile = edittile + 1
If KeyCode = vbKeySubtract Then edittile = edittile - 1

End Sub

Private Sub Form_Load()
edittile = 1
zoom = 3
CurEdMonster = 1
CurEdObject = 1
updatedittile
End Sub

Private Sub HScroll2_Change()
Label4.caption = "Brush Noise: " & HScroll2.Value
End Sub

Private Sub HScroll3_Change()
Label7.caption = "Tile Brush Width: " & HScroll3.Value
End Sub

Private Sub loadmen_Click()
FileD.FileName = "*.seg;*.map"
FileD.ShowOpen
FileD.DefaultExt = "seg"
'loadmap FileD.FileName
loadbindata FileD.FileName, ""
End Sub

Private Sub newmen_Click()
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
chars(CurChar).x = 5
chars(CurChar).y = 5
End Sub

Private Sub Option1_Click(Index As Integer)

CurEdLayer = Index

Select Case Index
    Case 0: Picture5.Visible = True
    Case 1: Picture2.Picture = LoadPicture(getfile(objtypes(CurEdMonster).graphname))
    Label3.caption = "Object Name: " & objtypes(CurEdObject).name
    Picture5.Visible = False
    Case 2: Picture2.Picture = LoadPicture(montype(CurEdMonster).gfile)
    Label3.caption = "Monster Name: " & montype(CurEdMonster).name
    Picture5.Visible = False
End Select


End Sub



Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
x = Int(x / zoom)
y = Int(y / zoom)
If x > mapx Or y > mapy Or x < 1 Or y < 1 Then Exit Sub
If Shift = 2 Then blockx = x: blocky = y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
x = Int(x / zoom)
y = Int(y / zoom)
Label2.caption = "X" & x & ", Y" & y

'Add brush noise
If HScroll2.Value > 0 Then x = x + roll(HScroll2.Value * 2) - Int(HScroll2.Value / 2) - 1: y = y + roll(HScroll2.Value * 2) - Int(HScroll2.Value / 2) - 1

If x > mapx Or y > mapy Or x < 1 Or y < 1 Then Exit Sub

If CurEdLayer = EdTiles Then
    For a = x - (HScroll3.Value - 1) To x + (HScroll3.Value - 1)
        For b = y - (HScroll3.Value - 1) To y + (HScroll3.Value - 1)
            If a > mapx Or a < 1 Or b > mapy Or b < 1 Then GoTo 12
            If Button = 1 And Shift = 0 And edittile <= 50 Then map(a, b).tile = edittile: map(a, b).blocked = Text1.text
            If Button = 2 Then map(a, b).ovrtile = 0: map(a, b).blocked = 0
            If Button = 1 And (Shift = 1 Or edittile > 50) Then map(a, b).ovrtile = edittile
            'If Not Shift = 2 Or Shift = 6 Then updatmap
12      Next b
    Next a
End If

If CurEdLayer = EdMonsters Then
    If Button = 1 And Shift = 0 Then If map(x, y).monster = 0 And map(x, y).blocked = 0 Then createmonster CurEdMonster, x, y
    If Button = 2 Then If map(x, y).monster > 0 Then killmon map(x, y).monster, 1
End If

'updatmap

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
x = Int(x / zoom)
y = Int(y / zoom)
If x > mapx Or y > mapy Or x < 1 Or y < 1 Then Exit Sub

If CurEdLayer = EdMonsters Then
    If Button = 1 And Shift = 0 Then If map(x, y).monster = 0 And map(x, y).blocked = 0 Then createmonster CurEdMonster, x, y
    If Button = 1 And Shift = 1 Then If map(x, y).monster > 0 Then killmon map(x, y).monster, 1
End If

If CurEdLayer = EdTiles Then
    If blockx > 0 Then
        For a = blockx To x
        For b = blocky To y
            If Int(Shift / 4) = 1 Then map(a, b).ovrtile = edittile Else map(a, b).tile = edittile
        Next b
        Next a
        
        blockx = 0: blocky = 0
    End If
End If

If CurEdLayer = EdObjects Then
    If Button = 1 And Shift = 0 Then
        SelectedObject = map(x, y).object
        If SelectedObject = 0 Then Exit Sub
        CurEdObject = objs(map(x, y).object).type
        Label3.caption = "Object Name: " & objtypes(CurEdObject).name
        Picture2.Picture = LoadPicture(getfile(objtypes(CurEdObject).graphname))
        
        Text2(0).text = objs(map(x, y).object).name
        Text2(1).text = objs(map(x, y).object).string
        Text2(2).text = objs(map(x, y).object).string2
        
    End If
    
    If Button = 1 And Shift = 1 Then
        If map(x, y).object = 0 Then createobj objtypes(CurEdObject).name, x, y, Text2(0).text, Text2(1).text, Text2(2).text
    End If
    
    If Button = 2 Then
        map(x, y).object = 0
    End If
    
End If

End Sub

Private Sub Picture2_Click()

    Select Case CurEdLayer
        
        Case EdObjects: Form3.Show 1
            Picture2.Picture = LoadPicture(getfile(objtypes(CurEdMonster).graphname))
            Label3.caption = "Object Name: " & objtypes(CurEdObject).name
            objtypes(CurEdObject).graphloaded = 0
        
        Case EdMonsters: MonsterEd.Show 1
            Picture2.Picture = LoadPicture(montype(CurEdMonster).gfile)
            Label3.caption = "Monster Name: " & montype(CurEdMonster).name
            getrgb montype(CurEdMonster).color, r, g, b
            makesprite mongraphs(CurEdMonster), Form1.Picture1, montype(CurEdMonster).gfile, r, g, b, montype(CurEdMonster).light, 3
        
        Case EdTiles: FileD.FileName = "*.bmp;*.gif"
            Select Case edittile
                Case 51: FileD.ShowOpen: mapjunk.name1 = FileD.FileTitle
                Case 52: FileD.ShowOpen: mapjunk.name2 = FileD.FileTitle
                Case 53: FileD.ShowOpen: mapjunk.name3 = FileD.FileTitle
                Case 54: FileD.ShowOpen: mapjunk.name4 = FileD.FileTitle
            End Select
            loadextraovrs mapjunk.name1, mapjunk.name2, mapjunk.name3, mapjunk.name4
            
    End Select
        
End Sub

Private Sub savemen_Click()
FileD.DefaultExt = "map"
FileD.ShowSave
'm'' savebindata FileD.FileName, 0
savebindata FileD.FileName 'm'' removed one args for compatibility
'savemap FileD.FileName
End Sub

Private Sub Text2_Change(Index As Integer)
If SelectedObject = 0 Then Exit Sub

Select Case Index
    Case 0: objs(SelectedObject).name = Text2(Index).text
    Case 1: objs(SelectedObject).string = Text2(Index).text
    Case 2: objs(SelectedObject).string2 = Text2(Index).text
End Select

End Sub

Private Sub Timer1_Timer()
updatmap
DoEvents
End Sub
