VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Class:    cBitmap
' Filename: cBitmap.cls
' Author:   SP McMahon
' Date:     14 July 1998
'
' A class to store a bitmap.  Similar to cSpriteBitmap but doesn't
' use a mask.
' ======================================================================

' Memory DC & Bitmap to hold the bitmap
Private m_hdcBitmap As Long
Private m_hBmpBitmap As Long
Private m_hBmpBitmapOld As Long
' Width & Height of the bitmap
Private m_lDX As Long
Private m_lDY As Long

#Const HYDEBUG = 1 'm''

Public Function CreateAtSize(ByVal lDX&, ByVal lDY&) As Boolean

    ' First ensure that we've cleared up anything we already have:
    ClearUp
    
    ' Create a DC at size to hold the bitmap:
    If (GDIMakeDCAndBitmap(False, m_hdcBitmap, m_hBmpBitmap, m_hBmpBitmapOld, _
        lDX, lDY)) Then
        m_lDX = lDX
        m_lDY = lDY
        CreateAtSize = True
    End If
    
End Function

Public Function CreateFromFile(ByVal sBitmapBitmap$) As Boolean
    
#If HYDEBUG = 1 Then 'm''
dbmsg ("cBitmaps.CreateFromFile, " & sBitmapBitmap$) 'm''
#End If 'm''

    Dim tBM As Bitmap

    ' First ensure that we've cleared up anything we already have:
#If HYDEBUG = 1 Then 'm''
dbmsg ("cBitmaps.CreateFromFile call ClearUp") 'm''
#End If 'm''

    ClearUp
    
#If HYDEBUG = 1 Then 'm''
dbmsg ("cBitmaps.CreateFromFile API GDILoadBitmapIntoDC") 'm''
#End If 'm''
    
    ' Load the Bitmap bitmap:
    If (GDILoadBitmapIntoDC(False, sBitmapBitmap, m_hdcBitmap, m_hBmpBitmap, _
        m_hBmpBitmapOld)) Then
        ' Get the dimensions of the bitmap (which is a DIB):
        GetObjectAPI m_hBmpBitmap, Len(tBM), tBM
        ' Store dimensions:
        m_lDX = tBM.bmWidth
        m_lDY = tBM.bmHeight
        CreateFromFile = True
    End If
    
#If HYDEBUG = 1 Then 'm''
dbmsg ("cBitmaps.CreateFromFile End") 'm''
#End If 'm''

End Function

Public Function CreateFromPicture(ByRef picThis As StdPicture)
    
    Dim tBM As Bitmap
    ClearUp
    If (GDILoadPictureIntoDC(False, picThis, m_hdcBitmap, _
        m_hBmpBitmap, m_hBmpBitmapOld)) Then
        ' Get the dimensions of the bitmap (which is a DIB):
        GetObjectAPI m_hBmpBitmap, Len(tBM), tBM
        ' Store dimensions:
        m_lDX = tBM.bmWidth
        m_lDY = tBM.bmHeight
        CreateFromPicture = True
    End If

End Function

Property Get Width() As Long
    Width = m_lDX
End Property

Property Get Height() As Long
    Height = m_lDY
End Property

Property Get hDC() As Long
    hDC = m_hdcBitmap
End Property

Public Sub RenderBitmap(ByVal lHDC&, ByVal lX&, ByVal lY&, _
        Optional ByVal OpCode As RasterOpConstants = vbSrcCopy)
    BitBlt lHDC, lX, lY, m_lDX, m_lDY, m_hdcBitmap, 0, 0, OpCode
End Sub

Public Sub AgrandissementImage(ByVal lHDC&, ByVal lX&, ByVal lY&, _
        ByVal lWSrc&, ByVal lHSrc&, ByVal lWScreen&, ByVal lHScreen&, _
        Optional ByVal OpCode As RasterOpConstants = vbSrcCopy)
        
    Dim lRet&
    lRet = SetStretchBltMode(hDC, 3)
    lRet = StretchBlt(lHDC, lX, lY, lWScreen, lHScreen, m_hdcBitmap, _
        0, 0, lWSrc, lHSrc, OpCode)
        
End Sub

Private Sub ClearUp()
    ' Here we ensure all our Objects are destroyed:
    GDIClearDCBitmap m_hdcBitmap, m_hBmpBitmap, m_hBmpBitmapOld
End Sub

Private Sub Class_Terminate()
    ClearUp
End Sub
