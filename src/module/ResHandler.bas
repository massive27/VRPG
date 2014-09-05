Attribute VB_Name = "ResHandler"
'resource data handler module
'
' created september 2014
'replace getfile() legacy function. Adds multiple pack searchthrough
'require cRes class, "Resourcepack handler class"

Dim ResPack() As cRes
Dim ResPackCount As Long

Dim ResPackPriority() As Long


Sub AddPack(FilePackName As String)
Dim Unable As Boolean

    If Dir$(FilePackName, vbNormal) = FilePackName Then
        ResPackCount = ResPackCount + 1
        ReDim Preserve ResPack(1 To ResPackCount)
        Set ResPack(ResPackCount) = New cRes
        If ResPack(ResPackCount).SetPackFile(FilePackName) Then
            'good
            Unable = False
        End If
    ElseIf Dir$(FilePackName, vbDirectory) = "." Then
        ResPackCount = ResPackCount + 1
        ReDim Preserve ResPack(1 To ResPackCount)
        Set ResPack(ResPackCount) = New cRes
        If ResPack(ResPackCount).SetFolderAsPackFile(FilePackName) Then
            'good
            Unable = False
        End If
    Else
        'bad target
        MsgBox "Cannot find " & FilePackName & " as resource for VRPG! Please check for misspelling.", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    If Unable Then
        'target found, but cannot be handled. corrupted ?
        ResPackCount = ResPackCount - 1
    Else
        'set a priority
        ReDim Preserve ResPackPriority(1 To ResPackCount) As Long
        ResPackPriority(ResPackCount) = ResPackCount
        'set temp folder. SHOULD BE SMARTER THAN THIS
        ResPack(ResPackCount).SetTempFolder App.Path
    End If
    
End Sub

Function GetResFile(ByVal WantedFile As String, Optional ByVal PackSrc As String = vbNullString) As String
Dim i As Long
'all checks and temp extraction are handled by the class, so the search
'is as simple as below.
'i use do loop to avoid initializing a for next when there is most of the time
'only 1 pack loaded

    i = 1
    Do
        GetResFile = ResPack(ResPackPriority(i)).GetResToFile(WantedFile)
        If GetResFile <> vbNullString Then Exit Function
        i = i + 1
    Loop Until i > ResPackCount

End Function
