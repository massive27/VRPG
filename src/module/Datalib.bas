Attribute VB_Name = "Datalib"
'This structure will describe our binary file's
'size and number of contained files
Private Type FILEHEADER
    NumFiles As Integer      'How many files are inside?
    FileSize As Long         'How big is this file? (Used to check integrity)
End Type

'This structure will describe each file contained
'in our binary file
Private Type INFOHEADER
    FileSize As Long         'How big is this chunk of stored data?
    FileStart As Long        'Where does the chunk start?
    FileName As String * 32  'What's the name of the file this data came from?
End Type

Public datinited

Public curaddstatus
Public curaddmax

Public mpq As MpqControl

Dim fileinstance As Integer

Function initdat(mpqc As MpqControl)

Set mpq = mpqc
datinited = 1

End Function

Function addfile(ByVal filen As String, ByVal PakFile As String, Optional replace = 0)

If datinited = 0 Then MsgBox "MPQ control not initialized.  Call Initdat to provide an mpq control reference.": End

'If Dir(pakfile) = "" Then
'mpf = mpq.mOpenMpq(pakfile) ': opened = 1

'mpq.mAddFile mpf, filen, "", 1
mpq.addfile PakFile, filen, filen, 1

'mpq.mCloseMpq mpf
'Stop
End Function

Function getfile(ByVal filen As String, Optional ByVal PakFile As String = "Data.pak", Optional ByVal add As Byte = 0, Optional extract As Byte = 0, Optional noerr = 0, Optional pakfileonly = 0) As String

ChDir App.Path

getfile = vbNullString

#If USELEGACY <> 1 Then
    'm'' alternate getfile procedure
    If PakFile = "Data.pak" Or PakFile = "" Then 'm'' may be a gamesave
        getfile = ResHandler.GetResFile(filen)  'm''
        Exit Function 'm''
    Else 'm''
        'm'' hotfix to prevent searching in modfolder when it should be in app.path
        If Left$(filen, 1) = "." Then
            filen = Mid$(filen, InStrRev(filen, "\") + 1)
        End If
    End If 'm''
#End If

On Error GoTo 3
GoTo 10
3 zerk = "Error in Getfile function while attempting to unpack " & filen & " from " & PakFile & "." & vbCrLf & "Application path is " & App.Path & "."
If Dir(PakFile) = "" Then zerk = zerk & vbCrLf & "Pak file was not found." Else zerk = zerk & vbCrLf & "Pak file was found, but internal file could not be accessed."
MsgBox zerk
Exit Function
10

'If filen = "patch1.gif" Then Stop
If pakfileonly = 0 And add = 0 And Not Dir(filen) = "" Then getfile = filen: Exit Function

If Not Dir("VTDATA" & filen) = "" Then GoTo 5
If Left(filen, 6) = "VTDATA" Then filen = Right(filen, Len(filen) - 6)

If datinited = 0 Then MsgBox "MPQ control not initialized.  Call Initdat to provide an mpq control reference.": End

If mpq.FileExists(PakFile, filen) = False Then
    If Right(filen, 4) = ".bmp" Then filen = Left(filen, Len(filen) - 4) & ".gif"
    If mpq.FileExists(PakFile, filen) = False Then getfile = "": Exit Function ' MsgBox "File not found:" & filen: Stop: Exit Function
End If

mpq.getfile PakFile, filen, App.Path, False
'mpq.getfile pakfile, filen, "", True
If Not Dir("VTDATA" & filen) = "" Then GoTo 5
ChDir App.Path
If extract = 0 Then Name filen As "VTDATA" & filen

5 If extract = 0 Then getfile = "VTDATA" & filen Else getfile = filen
'm'' making SURE there is VTDATA added to the filename
    If Not (Dir(filen)) = "" And Dir("VTDATA" & filen) = "" Then 'm''
        Name filen As "VTDATA" & filen 'm''
        getfile = "VTDATA" & filen 'm''
    ElseIf Not Dir("VTDATA" & filen) = "" Then 'm''
        getfile = "VTDATA" & filen 'm''
    End If 'm''
If Dir("VTDATA" & filen) = "" And Dir(filen) = "" Then MsgBox "File not found:" & filen: getfile = ""

End Function



Function killbinfiles()
'm'' modified clearing of files to benefits from cache (2012-11-20 release)
#If USELEGACY = 1 Then 'm''

On Local Error Resume Next
If Not Dir("VTDATA*.*") = vbNullString Then Kill "VTDATA*.*" 'm''

#Else
    On Local Error GoTo ErC 'm''
    Kill "VTDATA*.txt" 'm''
    Kill "VTDATA*.dat" 'm''
    Exit Function 'm''
ErC: 'm''
    Err.clear 'm'' it's faster to clear error than doing the Dir() stuff
    Resume Next 'm''
#End If

End Function

Function findfile(ByVal filen As String, ByVal PakFile As String) As Boolean

'Sees if the requested file is actually in the datafile

findfile = False

fileinstance = fileinstance + 1
If fileinstance > 100 Then fileinstance = 1

If Dir(PakFile) = "" Then Debug.Print "Pak file " & PakFile & " not found.": findfile = False: Exit Function

destext = Right(filen, 4) 'Get file extension
dfilen = "VTDATA" & fileinstance & destext
paknum = FreeFile
Open PakFile For Binary As paknum
Dim header As FILEHEADER
'paknum = FreeFile
Get paknum, 1, header

'Find file
Dim djerk As INFOHEADER
For a = 1 To header.NumFiles
    Get paknum, , djerk
    If Trim(djerk.FileName) = filen Then
        findfile = True: Close paknum: Exit Function
        Exit For
    End If
    Seek paknum, Seek(paknum) + djerk.FileSize
Next a

Close paknum 'Close both files, because we're done

'If Not found = 1 Then If add = 1 And Not Dir(filen) = "" Then addfile filen, pakfile 'Add to data file if not found
'If Not found = 1 Then addfile filen, pakfile: dfilen = filen  'If it wasn't found in the datafile, use the actual file if possible

End Function

