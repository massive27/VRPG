VERSION 5.00
Object = "{DA729162-C84F-11D4-A9EA-00A0C9199875}#1.60#0"; "MpqCtl.ocx"
Begin VB.Form frmConverter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pack converter"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      Caption         =   "new savegame to old"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Old savegame to new"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "pkf to pak"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
      Begin VB.CommandButton Command2 
         Caption         =   "Convert"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pak to pkf"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "Convert"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin MPQCONTROLLib.MpqControl MpqControl1 
      Left            =   360
      Top             =   3960
      _Version        =   65542
      _ExtentX        =   1085
      _ExtentY        =   1296
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Base 1

Const Slash$ = "\"

'legacy savegame format
'<player>.plr (mopak)
' ==> plrdat.tmp (vb data)
' ==> curgame.dat (mopak)
' ====> <level1>.dat (vb data)
' ====> <level2>.dat (vb data)

Sub AddLog(LogEntry As String)
    List1.AddItem LogEntry
    List1.ListIndex = List1.ListCount - 1
End Sub

Sub AddLogSelf(LogEntry As String)
Dim i As Long
    With List1
        i = .ListCount
        .RemoveItem (i - 1)
        .AddItem LogEntry
        .ListIndex = .ListCount - 1
    End With
End Sub

Sub SaveGame_Old()
savebindata Left(plr.curmap, Len(plr.curmap) - 4) & ".dat"
fsavechar "plrdat.tmp" 'FileD.FileTitle

AddFile "plrdat.tmp", filen, 1
AddFile "curgame.dat", filen, 1

End Sub

Private Sub SaveGame_bindata()

ChDir App.Path

If Left(fn, 6) = "VTDATA" Then fn = Right(fn, Len(fn) - 6)

'If Dir(App.Path & "\" & plr.name, vbDirectory) = "" Then MkDir plr.name
'fn = App.Path & "\" & plr.name & "\" & fn
If Not Dir(fn) = "" Then Kill fn
fnum = FreeFile
Open fn For Binary As fnum

Put fnum, , objts
Put fnum, , objtotal
Put fnum, , totalmonsters
Put fnum, , lastmontype
Put fnum, , mapx
Put fnum, , mapy

Put fnum, , mapjunk
Put fnum, , map()

Put fnum, , objtypes()
Put fnum, , objs()
Put fnum, , montype()
Put fnum, , mon()

Close fnum

AddFile fn, "curgame.dat", 1
Kill fn
End Sub

Sub DoPack()
If datinited = 0 Then MsgBox "MPQ control not initialized.  Call Initdat to provide an mpq control reference.": End

'If Dir(pakfile) = "" Then
'mpf = mpq.mOpenMpq(pakfile) ': opened = 1

'mpq.mAddFile mpf, filen, "", 1
mpq.AddFile PakFile, filen, filen, 1

'mpq.mCloseMp
End Sub

Private Sub Command1_Click()
'create pkf from pak
Dim Src As String, Dest As String
Dim TmpFold As String
Dim hMPQ As Long
Dim sBuf As String, i As Long, c As Long
Dim FileList() As String
Dim Pkf As cRes

AddLog "Selecting source .pak file..."
Src = SelectAfile(Me.hWnd, "Convert this pak", "pak file (*.pak)|*.pak|all files (*.*)|*.*|")
If Len(Src) > 0 Then
    AddLog "Selecting destination .pkf file..."
    Dest = SaveAfile(Me.hWnd, "Target pack file", "pkf file (*.pkf)|*.*|all files (*.*)|*.*|")
    If Len(Dest) > 0 Then
        If Mid$(Dest, Len(Dest) - 4, 1) <> "." Then Dest = Dest & ".pkf"
        AddLog "Try opening source .pak file..."
        hMPQ = MpqControl1.sOpenMpq(Src)
        If hMPQ > 0 Then
            sBuf = MpqControl1.sListFiles(hMPQ, vbNullString)
            FileList = Split(sBuf, vbCrLf, , vbBinaryCompare)
            c = UBound(FileList)
            AddLog "Found " & c - 1 & " files in source .pak "
            TmpFold = App.Path & Slash & Timer
            AddLog "Extracting to " & TmpFold
            MkDir TmpFold
            DoEvents
            AddLog "Extraction..."
            For i = 2 To c
                AddLogSelf "Extract: " & Int(i / c * 100) & "% (" & FileList(i) & ")"
                DoEvents
                Call MpqControl1.sGetFile(hMPQ, FileList(i), TmpFold, False)
            Next i
            AddLog "Closing source .pak file..."
            Call MpqControl1.sCloseMpq(hMPQ)
            
            AddLog "Starting new .pkf file..."
            Set Pkf = New cRes
            AddLog "Building new .pkf file...": DoEvents
            If Pkf.BuildPack(TmpFold, Dest) Then
                AddLog "Successfully built " & Dest
            Else
                AddLog "Failed to build " & Dest
            End If
            AddLog "Cleaning out temporary folder..."
            For i = 2 To c - 1
                Kill TmpFold & Slash & FileList(i)
            Next i
            RmDir TmpFold
            Set Pkf = Nothing
            AddLog "Job finished."
        Else
            AddLog "Failed. Aborting conversion."
        End If
    Else
        AddLog "Cancelled."
    End If
Else
    AddLog "Cancelled."
End If
        

End Sub

Private Sub Command2_Click()
'create pkf from pak
Dim Src As String, Dest As String
Dim sExt As String, TmpFold As String
Dim hMPQ As Long
Dim sBuf As String, i As Long, c As Long
Dim FileList() As String
Dim Pkf As cRes

AddLog "Selecting source .pkf file..."
Src = SelectAfile(Me.hWnd, "Convert this pkf", "pkf file (*.pkf)|*.pkf|all files (*.*)|*.*|")
If Len(Src) > 0 Then
    AddLog "Selecting destination .pak file..."
    Dest = SaveAfile(Me.hWnd, "Target pack file", "pak file (*.pak)|*.*|all files (*.*)|*.*|")
    If Len(Dest) > 0 Then
        If Mid$(Dest, Len(Dest) - 4, 1) <> "." Then Dest = Dest & ".pak"
        AddLog "Try opening source .pkf file..."
        Set Pkf = New cRes
        If Pkf.SetPackFile(Src) Then
            Call Pkf.GetAllFiles(FileList)
            c = UBound(FileList)
            AddLog "Found " & c & " files in source .pkf"
            TmpFold = App.Path & Slash & Timer
            AddLog "Extracting to " & TmpFold
            MkDir TmpFold
            DoEvents
            If Pkf.SetTempFolder(TmpFold) Then
                If Pkf.ExtractAll() Then
                    AddLog "Extraction successful."
                    
                     'definying the mpq table index
                    MpqControl1.DefaultMaxFiles = c + 1
                    
                    AddLog "Creating the target .pak file..."
                    hMPQ = MpqControl1.mOpenMpq(Dest)
                    Call MpqControl1.mCloseMpq(hMPQ)
                    DoEvents
                    hMPQ = MpqControl1.mOpenMpq(Dest)
                    If hMPQ > 0 Then
                        AddLog "Adding..."
                        For i = 1 To c
                            AddLogSelf "Add: " & Int(i / c * 100) & "% (" & FileList(i) & ")"
                            DoEvents
                            Call MpqControl1.mAddFile(hMPQ, TmpFold & Slash & FileList(i), FileList(i), 1)
                        Next i
                        Call MpqControl1.mCloseMpq(hMPQ)
                        
                        AddLog "Cleaning out temporary folder..."
                        Set Pkf = Nothing
                        For i = 1 To c
                            Kill TmpFold & Slash & FileList(i)
                        Next i
                        DoEvents
                        RmDir TmpFold
                        AddLog "Job finished."

                    Else
                        AddLog "Couldnt create target .pak file."
                    End If
                Else
                    AddLog "Some files may havent been extracted."
                End If
            Else
                AddLog "Couldnt reach temp folder. Aborting."
            End If
        Else
            AddLog "Failed. Aborting conversion."
        End If
    Else
        AddLog "Cancelled."
    End If
Else
    AddLog "Cancelled."
End If
        
End Sub

Private Sub Form_Load()
AddLog "Ready"
End Sub

Private Sub MpqControl1_GotMessage(ByVal Text As String)
    Debug.Print test
End Sub
