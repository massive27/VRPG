VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CRES.CLS
'==========
'
'object to handle resource
'for vrpg
'to get rid of mpq

Option Base 1
Option Explicit

Private Type RESFILE
    offset As Long
    length As Long
    extension As String * 4
    Name As String
End Type

Private Type RESFILEHEAD
    version As Long
    resarraylen As Long
    resdataoffset As Long
End Type
    
Private Type RESMEMORY
    FullName As String
    extension As String * 3
    Loaded As Boolean
End Type

Private Declare Sub CpMem Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Sub Class_Initialize()

End Sub
'add a folder as a resource
Public Sub AddFolder(DirName As String, Optional ByVal priority As Long = 1)

End Sub

Public Sub AddPack(PackName As String, Optional ByVal priority As Long = 1)

End Sub
'extract a file from a pack
Private Function ExtractMe(TargetPack As String, Destination As String, ByRef ResFilePtr As RESFILE)
Dim buf() As Byte
Dim fp As Integer
Dim tp As Integer

    ReDim buf(1 To ResFilePtr.length) As Byte
    
    On Local Error Resume Next
    fp = FreeFile
    Open TargetPack For Binary Access Read As #fp
    tp = FreeFile
    Open Destination For Binary As #tp
        Get #fp, ResFilePtr.offset, buf()
        Put #tp, 1, buf()
    Close #fp
    Close #tp
    
    If Err.Number <> 0 Then
        ExtractMe = False
        Err.clear
    Else
        ExtractMe = True
    End If

End Function

'read a pack file structure
Private Function LoadPack(TargetPack As String, ByRef tResFile() As RESFILE) As Boolean
Dim fp As Integer
Dim tmpHead As RESFILEHEAD

    fp = FreeFile
    Open TargetPack For Binary Access Read As #fp
        Get #fp, 1, tmpHead
        If tmpHead.version = 1 Then
        
            ReDim tResFile(1 To tmpHead.resarraylen) As RESFILE
            Get #fp, 13, tResFile()
            
            If tResFile(1).offset <> tmpHead.resdataoffset Then
                'corrupted
            End If
            
        Else
            'future version here
        End If
    Close #fp

End Function

'create a Pack file from a directory
Function BuildPack(ByVal DirSource As String, TargetPack As String, Optional ByVal Verbose As Boolean = False) As Boolean
Dim sfile As String
Dim fp As Integer
Dim rp As Integer, rlen As Long, buf() As Byte
Dim i As Long
Dim mb As VbMsgBoxResult
Dim miniRes() As RESFILE, miniHead As RESFILEHEAD
Dim ResCnt As Long
Dim snm As String, sext As String
Dim p As Long
Dim nextOffset As Long

    DirSource = CleanDirName(DirSource)
    If Dir$(DirSource, vbDirectory) = "" Then
        BuildPack = False
        If Verbose Then MsgBox "Directory " & DirSource & " not found.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    fp = FreeFile
    On Local Error GoTo notarget
    Open TargetPack For Binary As #fp
    
    
    sfile = Dir$(DirSource & "*", vbNormal)
    Do Until sfile = ""
        'array of resource
        ResCnt = ResCnt + 1
        If ResCnt Mod 50 = 0 Then
            ReDim Preserve miniRes(1 To ((ResCnt \ 50) + 1) * 50) As RESFILE
        End If
        
        'name cleanup
        p = InStrRev(sfile, ".", , vbBinaryCompare)
        If p = Len(sfile) - 4 Then
            sext = Right$(sfile, 4)
            snm = Left$(sfile, Len(sfile) - 5)
        Else
            sext = "    "
            snm = sfile
        End If
        
        'storing result
        miniRes(ResCnt).extension = sext
        miniRes(ResCnt).Name = snm
        miniRes(ResCnt).length = FileLen(DirSource & sfile)
        
    Loop
    
    'resizing resource table
    ReDim Preserve miniRes(1 To ResCnt) As RESFILE
    
    'preparing header
    nextOffset = Len(miniRes) + 16&
    miniHead.version = 1&
    miniHead.resarraylen = ResCnt
    miniHead.resdataoffset = nextOffset
    
    For i = 1 To ResCnt
        rp = FreeFile
        Open DirSource & miniRes(i).Name & miniRes(i).extension For Binary Access Read As #rp
            ReDim buf(1 To miniRes(i).length) As Byte
            Get #rp, 1, buf()
        Close #rp
        Put #fp, nextOffset, buf()
        miniRes(i).offset = nextOffset
        nextOffset = nextOffset + miniRes(i).length
    Next i
    
    Put #fp, 1, miniHead
    Put #fp, 13, miniRes()
    Close #fp
    
    BuildPack = True
    If Verbose Then MsgBox "Building " & TargetPack & " done.", vbInformation + vbOKOnly
    Exit Function

notarget:
    If Verbose Then
        mb = MsgBox(Err.Number & " : " & Err.Description, vbCritical + vbAbortRetryIgnore)
        Err.clear
        If mb = vbRetry Then
            Resume
        ElseIf mb = vbIgnore Then
            Resume Next
        Else
            BuildPack = False
            Exit Function
        End If
    Else
        Err.clear
        BuildPack = False
    End If
End Function
'quick function to cleanup dirname
Private Function CleanDirName(ThisDir As String) As String
    If Right$(ThisDir, 1) <> "\" Then
        CleanDirName = ThisDir & "\"
    Else
        CleanDirName = ThisDir
    End If
End Function
