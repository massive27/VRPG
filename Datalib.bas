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

Function addallfiles(pakfile As String, Optional ByRef addstatus, Optional ByRef totalfiles, Optional ByRef endwhendone = 0)

On Error GoTo 3
Dim dursnaw(1 To 5000) As String

'ChDir "C:\VB\ExpansionVRPG\MainDataFile" 'App.Path
dursnaw(1) = Dir("*.*")
lastfile = 1
For a = 2 To 5000
    dursnaw(a) = Dir
    If Not dursnaw(a) = "" Then totalfiles = a
Next a

    'will add any file except .exes and .dat files, so it won't add the database to itself

3 filen = dursnaw(lastfile)
If filen = "" Then GoTo 5
If Not filen = pakfile And Not Right(filen, 4) = ".exe" Then
addfile filen, pakfile, 1
End If
lastfile = lastfile + 1
If Not IsMissing(addstatus) Then addstatus = lastfile: DoEvents
GoTo 3

5 If endwhendone = 1 Then End

End Function

Function getfileold(ByVal filen As String, Optional ByVal pakfile As String = "Data.pak", Optional ByVal add As Byte = 0, Optional extract As Byte = 0, Optional noerr = 0) As String
'Creates a file in the directory from a file in a datafile; returns the filename created
'(So that we don't overwrite existing files with data file stuff--that would be bad, especially
'since it probably won't work right at first)
'Static instance As Integer

fileinstance = fileinstance + 1
If fileinstance > 100 Then fileinstance = 1

If Not Dir(filen) = "" Then getfileold = filen: Exit Function 'Use file if it's already in the directory, to make sure
                                                           'that new files in patches will work

If Left(filen, 6) = "VTDATA" Then filen = Right(filen, Len(filen) - 6): If Not Dir(filen) = "" Then getfileold = filen: Exit Function 'Check both with and without VTDATA at front

If Dir(pakfile) = "" And add = 0 Then getfileold = filen: ''debug.print "Pak file " & pakfile & " not found.": Exit Function
If Dir(pakfile) = "" Then createpak filen, pakfile

destext = Right(filen, 4) 'Get file extension
dfilen = "VTDATA" & filen 'Destination filename = "VTDATA" & origional filename
'dfilen = "VTDATA" & fileinstance & destext
If extract = 1 Then dfilen = filen
paknum = FreeFile
Open pakfile For Binary As paknum
Dim header As FILEHEADER
'paknum = FreeFile
Get paknum, 1, header

'Find file
checkletter = "ZZZ"
Dim djerk As INFOHEADER
For a = 1 To header.NumFiles
    Get paknum, , djerk
    If LCase(Left(djerk.FileName, Len(checkletter))) = LCase(checkletter) Then MsgBox "Error #1 in getfileold function."
    If LCase(Trim(djerk.FileName)) = LCase(filen) Or LCase(Trim(djerk.FileName)) = LCase("VTDATA" & filen) Then
        found = 1
        filenum = FreeFile
        Dim snuh() As Byte
        ReDim snuh(djerk.FileSize - 1) 'The -1 is because it starts at 0, not 1
        Get paknum, , snuh 'Read file data
        Open dfilen For Binary As filenum
        Put filenum, , snuh 'Write data to separate file
        Close filenum
        Exit For
    End If
    Seek paknum, Seek(paknum) + djerk.FileSize
    If Trim(djerk.FileName) = "##END" Then Exit For
Next a

Close paknum 'Close both files, because we're done

If Not found = 1 Then If add = 1 And Not Dir(filen) = "" Then addfile filen, pakfile Else getfileold = "" 'Add to data file if not found
If Not found = 1 And Not Dir(filen) = "" Then addfile filen, pakfile: dfilen = filen  'If it wasn't found in the datafile, use the actual file if possible
If Not found = 1 And add = 0 And noerr = 0 And Dir(filen) = "" Then MsgBox "File not found in " & pakfile & ": " & filen 'If not, return an error

If found = 1 Then getfileold = dfilen

'Wait for HD to catch up with writing file
Do While Dir(dfilen) = ""
If Not found = 1 Then Exit Do
Loop

'instance = instance - 1
End Function

Function addfile(ByVal filen As String, ByVal pakfile As String, Optional replace = 0)

If datinited = 0 Then MsgBox "MPQ control not initialized.  Call Initdat to provide an mpq control reference.": End

'If Dir(pakfile) = "" Then
'mpf = mpq.mOpenMpq(pakfile) ': opened = 1

'mpq.mAddFile mpf, filen, "", 1
mpq.addfile pakfile, filen, filen, 1

'mpq.mCloseMpq mpf
'Stop
End Function

Function getfile(ByVal filen As String, Optional ByVal pakfile As String = "Data.pak", Optional ByVal add As Byte = 0, Optional extract As Byte = 0, Optional noerr = 0, Optional pakfileonly = 0) As String

ChDir App.Path

getfile = ""

On Error GoTo 3
GoTo 10
3 zerk = "Error in Getfile function while attempting to unpack " & filen & " from " & pakfile & "." & vbCrLf & "Application path is " & App.Path & "."
If Dir(pakfile) = "" Then zerk = zerk & vbCrLf & "Pak file was not found." Else zerk = zerk & vbCrLf & "Pak file was found, but internal file could not be accessed."
MsgBox zerk
Exit Function
10

'If filen = "patch1.gif" Then Stop
If pakfileonly = 0 And add = 0 And Not Dir(filen) = "" Then getfile = filen: Exit Function

If Not Dir("VTDATA" & filen) = "" Then GoTo 5
If Left(filen, 6) = "VTDATA" Then filen = Right(filen, Len(filen) - 6)

If datinited = 0 Then MsgBox "MPQ control not initialized.  Call Initdat to provide an mpq control reference.": End

If mpq.FileExists(pakfile, filen) = False Then
    If Right(filen, 4) = ".bmp" Then filen = Left(filen, Len(filen) - 4) & ".gif"
    If mpq.FileExists(pakfile, filen) = False Then getfile = "": Exit Function ' MsgBox "File not found:" & filen: Stop: Exit Function
End If

mpq.getfile pakfile, filen, App.Path, False
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

Function addfileold(ByVal filen As String, ByVal pakfile As String, Optional replace = 0)

If Left(filen, 6) = "VTDATA" Then MsgBox "Error #1 in addfileold function.": filen = Right(filen, Len(filen) - 6)
paknum = FreeFile
If Dir(filen) = "" Then Debug.Print "Addfile:  File not found: " & filen: Exit Function
If Dir(pakfile) = "" Then createpak filen, pakfile: Exit Function


Open pakfile For Binary As paknum

Dim header As FILEHEADER

Get paknum, 1, header

'If header.FileSize <> LOF(paknum) Then ''debug.print "Addfile failed: File size failed check": Close paknum: Exit Function

Dim djerk As INFOHEADER

'Exit if the file already exists
totalsize = 0
For a = 1 To header.NumFiles '+1 is for ##END record
    Get paknum, , djerk
    '''debug.print "Found file:" & djerk.FileName
    If Trim(djerk.FileName) = "##END" Then Seek paknum, Seek(paknum) - Len(djerk): GoTo 5
    totalsize = totalsize + djerk.FileSize + Len(djerk)
    If Trim(djerk.FileName) = Trim(filen) Then
        If replace = 0 Then Close paknum: Exit Function
        If replace = 1 Then
            bytepos = Seek(paknum) - Len(djerk) 'Get position of start of current file
            readpos = bytepos + djerk.FileSize + Len(djerk)
            'readpos = Seek(paknum) 'Get position of next file in line
            Seek paknum, bytepos
            
            Dim doing As Byte
            doing = 0
            
            Flen = LOF(paknum)
            
            'cursp = 1
            'zoing = LOF(paknum) - bytepos - Len(djerk) 'This is here so it will leave the #END where it is!
            'Do While Not EOF(paknum) 'Overwrite data, by deleting the old file from the pak
            For b = readpos To Flen - Len(djerk)
                Get #paknum, readpos + cursp, doing
                Put #paknum, bytepos + cursp, doing
                cursp = cursp + 1
            Next b
            'Loop
            replaced = 1
            'Seek paknum, bytepos + cursp - 1
            'Seek paknum, LOF(paknum) - Len(djerk)
            GoTo 5  'Then go ahead and put the new file at the end of the data file as normal

        End If
    End If
    'If djerk.FileSize < 1 Then Exit Function
    Seek paknum, Seek(paknum) + djerk.FileSize
    
Next a

'Go to end of file
'Seek paknum, LOF(paknum) + 1 'Get rid of +1 if this doesn't work
'Stop:
Seek paknum, LOF(paknum) - Len(djerk)
5
'Add the file to the end of the datafile
Dim putfile() As Byte
filenum = FreeFile
Open filen For Binary As filenum

If LOF(filenum) = 0 Then MsgBox "Error #2 in addfileold function."

ReDim putfile(LOF(filenum) - 1)
Get filenum, , putfile
djerk.FileName = filen
djerk.FileSize = LOF(filenum)
Put paknum, , djerk 'Put file info
Put paknum, , putfile 'Put actual file data

Dim DUMM As INFOHEADER
DUMM.FileName = "##END"
Put paknum, , DUMM 'Mark the end of the file (To make sure replace stuff works

'Overwrite old header with updated header
If Not replaced = 1 Then header.NumFiles = header.NumFiles + 1
header.FileSize = LOF(paknum)
Put paknum, 1, header

Close filenum
Close paknum

End Function

Function createpak(ByVal filen As String, ByVal pakfile As String)

If Dir(filen) = "" Then Exit Function

If Not Dir(pakfile) = "" Then Exit Function
paknum = FreeFile
Open pakfile For Binary As paknum
filenum = FreeFile
Open filen For Binary As filenum

Dim doink As FILEHEADER
Put paknum, 1, doink

Dim djerk As INFOHEADER
djerk.FileName = filen
djerk.FileSize = LOF(filenum)

Put paknum, , djerk
Dim fnord() As Byte
ReDim fnord(LOF(filenum) - 1) As Byte

Get filenum, , fnord
Put paknum, , fnord

Dim DUMM As INFOHEADER
DUMM.FileName = "##END"
Put paknum, , DUMM 'Mark the end of the file (To make sure replace stuff works

doink.FileSize = LOF(paknum)
doink.NumFiles = 2
Put paknum, 1, doink

Close filenum
Close paknum

End Function

Function loadbinpic(picfile As String, pakfile As String) As StdPicture

'zark = getfile(picfile, pakfile, 1)
'loadbinpic = LoadPicture(zark)
MsgBox "Obsolete function Loadbinpic called"
Stop

End Function

Function openbinfile(ByVal filen As String, ByVal pakfile As String, Optional ByVal filenum) As Integer
'Directly replaces 'open' command by doing everything necessary to extract and open a file

If IsMissing(filenum) Then filenum = FreeFile

zarf = getfile(filen, pakfile, 1)
Open zarf For Binary As filenum

openbinfile = filenum

End Function

Function killbinfiles()
On Local Error Resume Next
If Not Dir("VTDATA*.*") = "" Then Kill "VTDATA*.*"

End Function

Function findfile(ByVal filen As String, ByVal pakfile As String) As Boolean

'Sees if the requested file is actually in the datafile

findfile = False

fileinstance = fileinstance + 1
If fileinstance > 100 Then fileinstance = 1

If Dir(pakfile) = "" Then Debug.Print "Pak file " & pakfile & " not found.": findfile = False: Exit Function

destext = Right(filen, 4) 'Get file extension
dfilen = "VTDATA" & fileinstance & destext
paknum = FreeFile
Open pakfile For Binary As paknum
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

