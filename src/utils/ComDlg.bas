Attribute VB_Name = "ComDlg"

' COMDLG
'========
'
'
'sources : allapi.net

Option Compare Binary

'commondialogbox folder
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'commondialogbox
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type



'folder browser popup
Function SelectADir(FormhWnd As Long, ByVal LeTitre As String) As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    
    If LeTitre = vbNullString Then LeTitre = "Select a dir"
    With udtBI
        .hwndOwner = FormFrom
        .lpszTitle = lstrcat(LeTitre, vbNullString)
        .ulFlags = 1
    End With

    'modal window
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(260, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
        SelectADir = sPath
    Else
        'cancel
        SelectADir = vbNullString
    End If
    
End Function

'open a file box
Function SelectAfile(FormFrom As Long, Titre As String, Filtre As String) As String
Dim OFName As OPENFILENAME

    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = FormFrom
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Replace$(Filtre, "|", Chr$(0), 1, , vbBinaryCompare)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = App.Path
    OFName.lpstrTitle = Titre
    OFName.flags = 0

    'pops the window
    If GetOpenFileName(OFName) Then
        SelectAfile = Trim$(OFName.lpstrFile)
    Else
        SelectAfile = vbNullString
    End If
End Function

'save as dialog
Function SaveAfile(FormFrom As Long, Titre As String, Filtre As String) As String
Dim OFName As OPENFILENAME

    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = FormFrom
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Replace(Filtre, "|", Chr$(0), 1)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = App.Path
    OFName.lpstrTitle = Titre
    OFName.flags = 0

    'pop the box
    If GetSaveFileName(OFName) Then
        SaveAfile = Left$(OFName.lpstrFile, InStr(1, OFName.lpstrFile, Chr$(0)) - 1)
    Else
        SaveAfile = vbNullString
    End If
End Function

Function GetFileName(Fichier As String) As String
    Fichier = Mid$(Fichier, InStrRev(Fichier, "\") + 1)
End Function
Function GetDossier(Fichier As String) As String
    GetDossier = Left$(Fichier, InStrRev(Fichier, "\"))
End Function
