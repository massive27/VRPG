VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Class:    cTile
' Filename: cTile.cls
' Author:   SP McMahon
' Date:     14 July 1998
'
' Tiles a sprite bitmap into a specified area
' ======================================================================

' Filename of bitmap to use in tiling:
Private m_sFileName As String
' m_cSprite & m_cSpriteBitmap Objects:
Private m_cBitmap As cBitmap
' Object to render on:
Private m_objRender As Object

Public Sub TileDC(ByVal lHDC&, ByVal lWidth&, ByVal lHeight&, _
        ByVal rDebutX!, ByVal rDebutY!)
    
    Dim lX&, lY&, dX&, dY&

    dX = m_cBitmap.Width
    dY = m_cBitmap.Height
    If dX < 2 Or dY < 2 Then Exit Sub
    
    ' Pour éviter de commencer tjrs au même endroit
    lY = -rDebutY * dY
    Do While lY <= lHeight
        lX = -rDebutX * dX
        Do While lX <= lWidth
            m_cBitmap.RenderBitmap lHDC, lX, lY
            lX = lX + dX
        Loop
        lY = lY + dY
    Loop
    
End Sub

Public Sub AgrandissementImgDC(ByVal lHDC&, ByVal lWScreen&, ByVal lHScreen&, _
        ByVal rStartX!, ByVal rStartY!, _
        ByVal bDecAGauche As Boolean, ByVal bDecEnHaut As Boolean)
    
    Dim lDebX&, lDebY&, lWImg&, lHImg&
    lWImg = m_cBitmap.Width
    lHImg = m_cBitmap.Height
    If lWImg < 2 Or lHImg < 2 Then Exit Sub
    
    ' Pour éviter de commencer tjrs au même endroit
    '  c'est un économiseur d'écran au fait !
    '  (et pas un testeur de résistance d'écran...)
    If lWImg > lWScreen Then
        If bDecAGauche Then
            lDebX = -rStartX * (lWImg - lWScreen)
            lWScreen = lWScreen - lDebX
        Else
            lDebX = 0
            lWScreen = lWScreen + rStartX * (lWImg - lWScreen)
        End If
    Else
        If bDecAGauche Then
            ' On diminue quand même la taille de l'image pour
            '  la décaller de façon à ne pas tjrs allumer les mêmes pixels
            lDebX = -rStartX * lWImg / 4
            ' On agrandit aussi l'écran de projection car on a rogné l'image source
            lWScreen = lWScreen - lDebX
        Else
            lDebX = 0
            lWScreen = lWScreen + rStartX * lWImg / 4
        End If
    End If
    If lHImg > lHScreen Then
        If bDecEnHaut Then
            lDebY = -rStartY * (lHImg - lHScreen)
            lHScreen = lHScreen - lDebY
        Else
            lDebY = 0
            lHScreen = lHScreen + rStartY * (lHImg - lHScreen)
        End If
        lDebY = -rStartY * (lHImg - lHScreen)
    Else
        If bDecEnHaut Then
            lDebY = -rStartY * lHImg / 4
            lHScreen = lHScreen - lDebY
        Else
            lDebY = 0
            lHScreen = lHScreen + rStartY * lHImg / 4
        End If
    End If
    m_cBitmap.AgrandissementImage lHDC, lDebX, lDebY, lWImg, lHImg, _
        lWScreen, lHScreen
    
End Sub

Property Let FileName(ByVal sFileName$)
    Dim hDCRender&
    If (sFileName <> m_sFileName) Then
        If (sFileName = "") Then
            ClearUp
        Else
            If (pbValid()) Then
                ' We create the sprite bitmap from the file if possible:
                If (m_cBitmap.CreateFromFile(sFileName)) Then
                    m_sFileName = sFileName
                End If
            End If
        End If
    End If
End Property

Property Let Picture(ByRef oPic As StdPicture)
    ClearUp
    ' We create the sprite bitmap from the file if possible:
    Set m_cBitmap = New cBitmap
    If (m_cBitmap.CreateFromPicture(oPic)) Then
    End If
End Property
Private Function pbValid() As Boolean
    If (Not (m_objRender Is Nothing)) Then
        pbValid = True
    Else
        pErr 1, "Object not initialised"
    End If
End Function
Public Sub Initialise(ByVal objRender As Object)
    Set m_objRender = objRender
End Sub
Private Sub pErr(ByVal lErrNum&, ByVal sErrText$)
    ' err.raise vbObjectError + 1048 + lErrNum,App.ExeName & ".m_cTileile",sErrText
    Debug.Print "Error: " & sErrText
End Sub
Private Sub ClearUp()
    Set m_cBitmap = Nothing
    Set m_objRender = Nothing
    m_sFileName = ""
End Sub
Private Sub Class_Initialize()
    ' Create valid instances of the sprite and
    ' sprite bitmap Objects:
    Debug.Print "m_cTileile:Initialise"
    Set m_cBitmap = New cBitmap
End Sub

Private Sub Class_Terminate()
    ' clear up all the Object references:
    Debug.Print "m_cTileile:Terminate"
    ClearUp
End Sub
