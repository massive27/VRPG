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
'object to handle resourcepack file
'for vrpg
'to get rid of mpq

Option Base 1
Option Explicit

Private Type RESFILE '12
    offset As Long
    length As Long
    vindex As Long
End Type
Private Res_Names() As String
Private Res_Ext() As String

Private Type RESFILEHEAD '20
    version As Long
    resarraylen As Long
    resdataoffset As Long
    resnamesoffset As Long
    resextoffset As Long
End Type


Private MyPack As String
Private MyTempFolder As String

'tracker of already-extracted files
'not doing Type for faster lookup
Private MyExtracted() As String
Private MyExtPath() As String
Private MyExtCnt As Long

Private aResource() As RESFILE
Private ResLen As Long

Public NumResource As Long

Private Declare Sub CpMem Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Sub Class_Initialize()
    MyTempFolder = CleanDirName(App.Path)
End Sub

'sets the packfile to use. Will load it if exists
Public Function SetPackFile(PackFileName As String) As Boolean

    If Dir$(PackFileName, vbNormal) <> "" Then
        MyPack = PackFileName
        SetPackFile = LoadPack(MyPack, aResource)
        If SetPackFile Then
            ResLen = UBound(aResource)
        End If
    Else
        SetPackFile = False
    End If

End Function
'sets the temp folder for extraction
Public Function SetTempFolder(DestFolder As String) As Boolean
Dim sBuf As String

    'test if exists
    sBuf = CleanDirName(DestFolder)
    If Dir$(sBuf, vbDirectory) = "" Then
        SetTempFolder = False
    Else
        MyTempFolder = sBuf
        SetTempFolder = True
    End If

End Function
'return a path to a file, or null if the file is not existing in the pack
Public Function GetResToFile(ThisFile As String) As String
Dim i As Long, p As Long
Dim bareName As String
Dim sExt As String
Dim sDest As String

    Call CleanFileExt(ThisFile, bareName, sExt)
    
    'is already extracted ?
    For i = 1 To MyExtCnt
        If MyExtracted(i) Like bareName Then
            GetResToFile = MyExtPath(i)
            Exit Function
        End If
    Next i
    
    'not extracted, look in pack
    For i = 1 To ResLen
        If Res_Names(i) Like bareName Then
            'found. prepare output
            sDest = MyTempFolder & Res_Names(i) & sExt
            'extracts
            If ExtractMe(sDest, aResource(i)) Then
                GetResToFile = sDest
                'update tracker
                MyExtCnt = MyExtCnt + 1
                MyExtracted(MyExtCnt) = bareName
                MyExtPath(MyExtCnt) = sDest
            Else
                GetResToFile = vbNullString
            End If
            Exit Function
        End If
    Next i
    GetResToFile = vbNullString

End Function

'extract all file from current pack
Function ExtractAll() As Boolean
Dim i As Long
Dim b As Boolean
Dim sDest As String
    ExtractAll = True
    For i = 1 To ResLen
        sDest = MyTempFolder & Res_Names(i) & Res_Ext(i)
        b = ExtractMe(sDest, aResource(i))
        ExtractAll = ExtractAll And b
    Next i
    
End Function

'extract a file from a pack
Private Function ExtractMe(Destination As String, ByRef ResFilePtr As RESFILE)
Dim buf() As Byte
Dim fp As Integer
Dim tp As Integer

    ReDim buf(1 To ResFilePtr.length) As Byte
    
    'On Local Error Resume Next
    fp = FreeFile
    Open MyPack For Binary Access Read As #fp
    tp = FreeFile
    Open Destination For Binary As #tp
        Get #fp, ResFilePtr.offset, buf()
        Put #tp, 1, buf()
    Close #fp
    Close #tp
    
    'possible error : unable to write to dest, io error
    If Err.Number <> 0 Then
        ExtractMe = False
        Err.Clear
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
            ReDim Res_Names(1 To tmpHead.resarraylen) As String
            ReDim Res_Ext(1 To tmpHead.resarraylen) As String
            Me.NumResource = tmpHead.resarraylen
            Get #fp, 21, tResFile()
            Get #fp, tmpHead.resextoffset, Res_Ext()
            Get #fp, tmpHead.resnamesoffset, Res_Names()
            
            If tResFile(1).offset <> tmpHead.resdataoffset Then
                'corrupted. to be handled. later.
            Else
                'prepare the extract tracker
                ReDim MyExtracted(1 To tmpHead.resarraylen) As String
                ReDim MyExtPath(1 To tmpHead.resarraylen) As String
                LoadPack = True
            End If
            
        Else
            'future version here
        End If
    Close #fp

End Function

'create a Pack file from a directory
Function BuildPack(ByVal DirSource As String, TargetPack As String, Optional ByVal Verbose As Boolean = False) As Boolean
Dim sFile As String
Dim fp As Integer
Dim rp As Integer, rlen As Long, buf() As Byte
Dim i As Long
Dim mb As VbMsgBoxResult
Dim miniRes() As RESFILE, miniHead As RESFILEHEAD
Dim rNames() As String, rExt() As String
Dim ResCnt As Long
Dim snm As String, sExt As String
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
    
    ReDim miniRes(1 To 50) As RESFILE
    ReDim rNames(1 To 50) As String
    ReDim rExt(1 To 50) As String
    ResCnt = 0
    
    sFile = Dir$(DirSource & "*", vbNormal)
    Do Until sFile = vbNullString
        'array of resource
        ResCnt = ResCnt + 1
        If ResCnt Mod 50 = 0 Then
            ReDim Preserve miniRes(1 To ((ResCnt \ 50) + 1) * 50) As RESFILE
            ReDim Preserve rNames(1 To ((ResCnt \ 50) + 1) * 50) As String
            ReDim Preserve rExt(1 To ((ResCnt \ 50) + 1) * 50) As String
        End If
        
        'name cleanup
        Call CleanFileExt(sFile, snm, sExt)
        
        'storing result
        rExt(ResCnt) = sExt
        rNames(ResCnt) = snm
        miniRes(ResCnt).length = FileLen(DirSource & sFile)
        
        sFile = Dir$()
    Loop
    
    'resizing resource table
    ReDim Preserve miniRes(1 To ResCnt) As RESFILE
    ReDim Preserve rExt(1 To ResCnt) As String
    ReDim Preserve rNames(1 To ResCnt) As String
    
    'preparing header
    miniHead.version = 1&
    'total number of entries
    miniHead.resarraylen = ResCnt
    'offset to extention string array
    nextOffset = ResCnt * 20& + 17&
    miniHead.resextoffset = nextOffset
    'offset to name string array
    nextOffset = nextOffset + LenAr(rExt, ResCnt)
    miniHead.resnamesoffset = nextOffset
    'offset to files data
    nextOffset = nextOffset + LenAr(rNames, ResCnt)
    miniHead.resdataoffset = nextOffset

    'adding header
    Put #fp, 1, miniHead
    'adding extensions array
    Put #fp, miniHead.resextoffset, rExt()
    'adding names array
    Put #fp, miniHead.resnamesoffset, rNames()
    
    'copying resources file into pack
    For i = 1 To ResCnt
        rp = FreeFile
        Open Trim$(DirSource & rNames(i) & rExt(i)) For Binary Access Read As #rp
            ReDim buf(1 To miniRes(i).length) As Byte
            Get #rp, 1, buf()
        Close #rp
        Put #fp, nextOffset, buf()
        miniRes(i).offset = nextOffset
        nextOffset = nextOffset + miniRes(i).length
    Next i
    
    Put #fp, 1, miniHead
    Put #fp, 21, miniRes()
    Close #fp
    
    BuildPack = True
    If Verbose Then MsgBox "Building " & TargetPack & " done.", vbInformation + vbOKOnly
    Exit Function

notarget:
    If Verbose Then
        mb = MsgBox(Err.Number & " : " & Err.Description, vbCritical + vbAbortRetryIgnore)
        Err.Clear
        If mb = vbRetry Then
            Resume
        ElseIf mb = vbIgnore Then
            Resume Next
        Else
            BuildPack = False
            Exit Function
        End If
    Else
        Err.Clear
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
'quick function to strip extension
Private Sub CleanFileExt(ThisFile As String, ByRef OutName As String, ByRef OutExt As String)
Dim p As Long

    'name cleanup : separate extension
    p = InStrRev(ThisFile, ".", , vbBinaryCompare)
    If p = Len(ThisFile) - 3 Then
        OutExt = Right$(ThisFile, 4)
        OutName = Left$(ThisFile, Len(ThisFile) - 4)
    Else
        OutExt = "    "
        OutName = ThisFile
    End If
    
End Sub

'quick function for array len of string
Private Function LenAr(ByRef MyArray() As String, ByVal MyAL As Long)
Dim i As Long
Dim c As Long

    For i = 1 To MyAL
        c = c + Len(MyArray(i)) + 2
    Next i
    LenAr = c
    
End Function
