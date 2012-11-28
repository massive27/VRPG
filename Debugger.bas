Attribute VB_Name = "Debugger"
'm'' DEBUGGER.BAS
'm'' ============
'm'' module created to make here the
'm'' various sub in order to improve the game
'm'' or fix bugs.
'm''
'm'' debugger's name : Massive27 from Eka's portal
'm'' http://aryion.com/forum/viewforum.php?f=52
'm''
'm'' MANY THANKS to : Aleph-Null and wsensor
'm''  who find back the source code :D
'm''
'm'' <= I use that kind of comment in order to get track
'm''    of all my modifications
'm''
'm'' historic
'm'' 2012-03-23 :
'm'' - fixed two call to legacy .TransparentDraw calls
'm'' - fixed legacy TransparentDraw of cSprite.cls
'm'' - rewrote SetTimer call to get the handler for KillTimer
'm'' - adding another control on getfile() in order to make the game clear itself better when quitting
'm'' - wrote a quitting function to avoid crash or memory leakage
'm'' - fixed skilltree.frm and VRmaped.frm from stitching. now, must be integrated
'm'' 2012-04-05
'm'' - skilltree.frm code downgraded to be compatible with main code
'm'' - modified TitleScreen property to avoid resize
'm'' - renamed curversion to 2.14 Legacy
'm'' - added File I/O check and data.pak available check for compatibility troubleshoot
'm'' - fixed attack animation and alike (bijdang) to render properly instead of first frame
'm'' - improved cSpriteBitmaps.CreateFromFile function for faster exec.
'm'' - gotomap function viewed
'm'' 2012-04-06
'm'' - fixed bad format wave file by on-the-fly rewritting header before loading
'm'' - fixed error when temporary files couldnt be deleted
'm'' - using alternative formula for tryescape
'm'' 2012-04-07
'm'' - UI tweak : at first loading, refresh properly with the "Loading" wait message
'm'' - removed the SetTimer() call that is actually unused and may crash the game
'm'' - fixed the "once every 9" not drawn friendly monsters
'm'' - removed "arrays bound checks" when compiling for smaller exe.
'm'' - cSpriteBitmaps loading improvements in progress
'm'' - fixed a swap between succubus class and naga class (non genuine)
'm'' 2012-04-08
'm'' - Reverted the swap between succubus and Naga. removed Giant Stomach 1 from succubus and added to Naga instead.
'm'' 2012-04-10
'm'' - added a "slayer" cheat code to kill minions but boss, to give an alternative to "oopsie". see Debugger.killallmonster_butboss
'm'' - in takeobj, added a check for Naga class in order to avoid destruction (and crash) of non wearable clothes
'm'' 2012-04-11
'm'' - code format : added conditional compilation in order to keep non-genuine changes from original gameplay
'm'' - fixed monsters fighting each other never eating each other (non genuine) see monai and monatk2
'm'' - fixed bosses that swallowed a monster unable to attack the player (non genuine) (exploit)
'm'' 2012-04-14
'm'' - fixed still checkbox of sound now moving with other objects when resizing
'm'' - fixed crash when eaten monster managed to escape another monster belly
'm'' - fixed boss eating everything that attacks him, whenever it has already a full belly or not
'm'' - removed titlescreen from the game internal resource, it is now loaded from external data.pak.
'm'' - added Mod support through command line parsing
'm'' - rewritten getfile() procedure to handle Mod resource as prioritized
'm'' 2012-04-15
'm'' - fixed crash when cancelling choosing a saved game
'm'' - fixed an infinite loop hang when choosing and cancelling saved game twice
'm'' - fixed crash when inputting a non-numerical value for buying pot
'm'' - added the option to continue game after killing Trisha
'm'' 2012-05-24
'm'' - fixed the directx not loading properly following the choose/cancel save game fix
'm'' 2012-11-05
'm'' - added an output-to-file debugger for debug purpose
'm'' - Astar algorithm started but unused


Public API_Timer_Handle As Long

Public BackBufHDC As Long

'm'' following declaration to make stitched sources working ...
Public TileColors(1 To 50) As Long 'Minimap colors for tiles
Public chars() As playertype
Public CurChar As Byte 'Current Character

'declaration for directsound
Public DSND As DirectSound
Public DSD_PlayBuf As DirectSoundBuffer
Public dsD_Init As Boolean

Type DSDBUFF
    SoundName As String
    DSBuffer As DirectSoundBuffer
End Type

Dim soundbank() As DSDBUFF
Dim soundbanklen As Long
Dim soundbufferDsc As DSBUFFERDESC

'declaration for temporary files management
Public d_TmpFolder As String

'declaration for multipack (mod) management
Public PakFiles() As String
Public PakCount As Long

'declaration for A-star pathfinding algorithm ====
Private Type ASTAR_POINT
    X As Long
    Y As Long
    F As Long 'G + H
    g As Long 'G score : how "much" to move here?
    H As Long 'H score : manhattan distance
    Parent As Long 'index to parent point
    Closed As Long 'flag to say : do not use
End Type
Private ASTAR_MATRIX(1 To 9) As ASTAR_POINT 'pseudo-matrix of x,y point for linear calculation of pathfinding

Private Type tCell
    X As Long               'Coordinates of the listed cell
    Y As Long
    Parent As Long          'Parent Index within the list (-1 for start point)
    Cost As Long          'Cost to get til here
    Heuristic As Long     'Estimated cost til target
    Closed As Boolean       'Not considered anymore
End Type
Private Type tGrid
    ListStat As eListStat   'Status of the list element
    Index As Long           'Index into the open list.
End Type
Private Enum eListStat
    Unprocessed = 0&
    IsOpen = 1&
    IsClosed = 2&
End Enum
Private Type tPoint
    X As Long
    Y As Long
End Type
Dim asMaxX As Long 'map maximum x coordinate
Dim asMaxY As Long 'map maximum y coordinate

'a-star for monster "memory of route" declaration
Type AS_TRACK
    IndexRoute As Long
    Route() As tPoint
End Type
Dim monroute() As AS_TRACK

Sub GD_BulgeIt(picFrom As PictureBox, picTo As PictureBox, Xcenter, Ycenter, Radius, Factor)
Dim pcd As Double

'experimental Bulge generator to have cleaner big belly.
'to fine tune

picTo.Picture = picFrom.Picture
    
'direct calc, no conv matrix
minx = Xcenter - Radius
miny = Ycenter - Radius
maxx = Xcenter + Radius
maxy = Ycenter + Radius
For si = minx To maxx
    pcd = (Xcenter - si) * (Xcenter - si)
    For sj = miny To maxy
        'a = Atan2(Ycenter - sj, Xcenter - si)
        r = Sqr(pcd + (Ycenter - sj) ^ 2)
        If r <= Radius Then
            rn = r + Sin((r / Radius) * 3.141592 + 3.141592) * Factor
        Else
            rn = r
        End If
        ox = Xcenter - (rn * Cos(a))
        oy = Ycenter - (rn * Sin(a))
        pxl = picFrom.Point(ox, oy)
         picTo.PSet (si, sj), pxl
    Next sj
Next si
End Sub

Function getfile_mod(ByVal filen As String, Optional ByVal PakFile As String = "Data.pak", Optional ByVal add As Byte = 0, Optional extract As Byte = 0, Optional noerr = 0, Optional pakfileonly = 0) As String
Dim i As Long

'm'' modified getfile function to handle multiple pack

getfile_mod = ""

'm'' file already extracted
If Not Dir(filen) = "" Then getfile_mod = filen: Exit Function
If Not Dir("VTDATA" & filen) = "" Then getfile_mod = "VTDATA" & filen: Exit Function

'm'' handling unextracted files
If Left(filen, 6) = "VTDATA" Then filen = Right(filen, Len(filen) - 6)
If datinited = 0 Then MsgBox "MPQ control not initialized. Restart the game.": Exit Function

'm'' seeking if file exists
For i = 1 To PakCount

If Right$(PakFiles(i), 1) = "\" Then
    'not a pak, but a folder
    cfile$ = ".\" & PakFiles(i) & filen
    If Not (Dir(cfile$) = "") Then
        'file exists in the folder. we simply said that "it is here"
        getfile_mod = cfile$
        Exit Function
    Else
        'file not exists. Duam stuff : not a bitmap, maybe a gif ?
        If Right(filen, 4) = ".bmp" Then filen = Left(filen, Len(filen) - 4) & ".gif"
        cfile$ = ".\" & PakFiles(i) & filen
        If Not (Dir(cfile$) = "") Then
            'file exists in the folder. we simply said that "it is here"
            getfile_mod = cfile$
            Exit Function
        End If
        'trick end, nothing.
    End If
    
Else
    'a pak file
    ChDir App.Path
    If mpq.FileExists(PakFiles(i), filen) = False Then
        'Duam extension swap
        If Right(filen, 4) = ".bmp" Then
            filen = Left(filen, Len(filen) - 4) & ".gif"
        Else
            filen = Left(filen, Len(filen) - 4) & ".bmp"
        End If
        If mpq.FileExists(PakFiles(i), filen) = False Then Exit Function  ' MsgBox "File not found:" & filen: Stop: Exit Function
    End If
    
    mpq.getfile PakFiles(i), filen, App.Path, False
    
    'm'' now the file is extracted, let's rename it
    If Not (Dir(filen) = "") Then
        If Dir("VTDATA" & filen) = "" Then
            'm'' rename
            Name filen As "VTDATA" & filen
            getfile_mod = "VTDATA" & filen
            Exit Function
        End If
    End If
    
End If
Next i

End Function

Sub killallmonster_butboss()
'a little code to have an alternative of the "oopsie" cheatcode.
'this will check on name of creature in order to avoid killing the boss.
Dim a As Long

    'the name of boss is currently the questmonster.
    For a = 1 To UBound(mon())
        If mon(a).type > 0 Then
            If montype(mon(a).type).name = mapjunk.questmonster Then
                'skip
            Else
                killmon a, 1 'parameter 1 avoid the "death" sound
            End If
        Else
            killmon a, 1
        End If
    Next a

End Sub


Sub DataPak_Check(InThisDir As String)
'will test for data file presence

Dim FD As String
Dim fp As Integer
Dim buf As Long
Dim MsgR

    On Local Error GoTo errmgr

    If InStrRev(InThisDir, "\", , vbBinaryCompare) = Len(InThisDir) Then
        FD = InThisDir & "Data.pak"
    Else
        FD = InThisDir & "\Data.pak"
    End If
    
    'attempt to locate the file
    If Dir(FD) = "" Then
        MsgR = MsgBox("Warning : data.pak not found in " & InThisDir & ". Game wont be able to load.", vbCritical + vbOKOnly)
        Exit Sub
    End If
    'attempt to read the file
    fp = FreeFile
    Open FD For Binary As #fp
        Get #fp, , buf
    Close #fp
    Exit Sub
    
errmgr:
    MsgBox "DataPak_Check failed! error code : " & Err.Number & ", error description : " & Err.Description & ". Game may fail to load.", vbExclamation + vbOKOnly
    Err.clear
    Resume Next
        
    
End Sub

Sub TmpFold_Check(ThisDir As String)
'will test for file access read and delete,
'this is for compatibility with NTFS security inside winvista and win7

Dim MsgR
Dim FooFile As String
Dim fp As Integer, a As Long, dstep As Long

    FooFile = "_tmp.tmp"
    a = &H12345678
    
10:
    If Dir(ThisDir, vbDirectory) = "" Then
    'directory doesnt exists (very rare)
        MsgR = MsgBox("Critical error : temporary directory " & ThisDir & " is unavailable. Please input another folder name, else the game will quit.", vbCritical + vbOKCancel)
        If MsgR = vbCancel Then
            Debugger.Quitting
            End
        Else
            MsgR = InputBox("Please write another folder", "Temporary folder", ThisDir)
            If MsgR = vbCancel Then
                Debugger.Quitting
                End
            Else
                GoTo 10
            End If
        End If
    
    Else
        If InStrRev(ThisDir, "\", , vbBinaryCompare) < Len(ThisDir) Then
            ThisDir = ThisDir & "\"
        End If
            
        fp = FreeFile
        On Local Error GoTo errmanager
        'vb error will be raised
        dstep = 1
        Open ThisDir & FooFile For Binary As #fp
            dstep = 2
            Put #fp, , a
        Close fp
        dstep = 3
        Open ThisDir & FooFile For Binary As #fp
            dstep = 4
            Get #fp, , a
            dstep = 5
        Close fp
        dstep = 6
        Kill ThisDir & FooFile
        
    End If
    Exit Sub
    
errmanager:
    'this part will be enhanced with all the error value encoutered later
    '75, 53, etc
    a = Err.Number
    
    MsgR = MsgBox("Error : in TmpFold_Check(), vbcode = " & a & ", vbdesc = " & Err.Description & ", dstep = " & dstep & ". Please report!", vbExclamation + vbAbortRetryIgnore, "VRPG Debug")
        
    If MsgR = vbIgnore Then
        Err.clear
        Resume Next
    ElseIf MsgR = vbRetry Then
        Err.clear
        Resume
    Else
        Err.clear
        Exit Sub
    End If

End Sub

Sub Snd_Init(ByVal SomeHWND As Long)
'will initialize sound management

    If ds_Init = False Then
        
        Set DSND = dX.DirectSoundCreate("")
        DSND.SetCooperativeLevel SomeHWND, DSSCL_NORMAL
        
        soundbufferDsc.lFlags = DSBCAPS_STATIC Or DSBCAPS_CTRLVOLUME Or DSBCAPS_GLOBALFOCUS Or DSBCAPS_LOCSOFTWARE
        
        ds_Init = True
    End If

End Sub

Sub Snd_Play(ByVal filen As String)
'm'' play a sound. Also check the sound bank
Dim t As Long

    t = Snd_IsHere(filen)
    If t > 0 Then
        soundbank(t).DSBuffer.Play DSBPLAY_DEFAULT
    Else
        soundbanklen = soundbanklen + 1
        ReDim Preserve soundbank(1 To soundbanklen) As DSDBUFF
        soundbank(soundbanklen).SoundName = filen
        Set soundbank(soundbanklen).DSBuffer = Snd_Loadme(filen)
        
        soundbank(soundbanklen).DSBuffer.Play DSBPLAY_DEFAULT
    End If

End Sub

Sub Snd_Stop()
'm'' force stopping of all sounds.
Dim i As Long

    If ds_Init = False Then Exit Sub 'm'' who knows... let's avoid error
    
    For i = 1 To soundbanklen
        soundbank(soundbanklen).DSBuffer.Stop
    Next i

End Sub

Private Function Snd_Loadme(filen As String) As DirectSoundBuffer
'm'' quick file checker and buffer creator for DX7
Dim fp As Integer
Dim WvFrEx As WAVEFORMATEX

    On Local Error GoTo WaveFail

    fp = FreeFile
    Open filen For Binary Access Read As #fp
        Get #fp, 21, WvFrEx
    Close #fp
    WvFrEx.lExtra = 0
    Set Snd_Loadme = DSND.CreateSoundBufferFromFile(filen, soundbufferDsc, WvFrEx)
    
    
    Exit Function
WaveFail:
    If Err.Number = 5 Then
        'm'' if err.number = 5 it's because the wave file format isnt standard
        'm'' will load the file in memory and try to fix it.
        Err.clear
        Dim bbuf() As Byte, ibuf As Integer, i As Long, lbuf As Long
        fp = FreeFile
        Open filen For Binary As #fp
            For i = 21 To 500 'scans the 500 first bytes for the "data" header
                Get #fp, i, ibuf
                If ibuf = &H6164 Then
                    Get #fp, i + 2, ibuf
                    If ibuf = &H6174 Then
                        'located the "data"
                        Get #fp, , lbuf
                        ReDim bbuf(1 To lbuf)
                        'load the PCM WAVE data
                        Get #fp, , bbuf
                        'rewrite the header
                        i = lbuf + 36
                        Put #fp, 1, &H46464952 '"RIFF"
                        Put #fp, , i
                        Put #fp, 37, &H61746164 '"data"
                        Put #fp, , lbuf
                        Put #fp, , bbuf
                        Exit For
                    End If
                End If
            Next i
        Close #fp
        
        Resume
        'm'' create an ampty buffer in case of failure ...
        'm''Set Snd_Loadme = DSND.CreateSoundBuffer(soundbufferDsc, WvFrEx)
    End If
    
End Function

Private Function Snd_IsHere(sname As String) As Long
'm'' find or not a loaded sound and return its index
Dim i As Long
Dim m As Long
Dim dc As DSCURSORS

    m = 0
    For i = 1 To soundbanklen
        If soundbank(i).SoundName = sname Then
            'm'' trick : to have multiple instance of the same sound,
            'm'' we have to duplicate the sndbuff, unless another is ready
            soundbank(i).DSBuffer.GetCurrentPosition dc
            If dc.lPlay > 0 Then 'm'' buffer being read.
                If m = 0 Then m = i 'm'' we store the first buffer index for duplicating
            Else
                Snd_IsHere = i
                Exit Function
            End If
        End If
    Next i
    
    If m > 0 Then 'm'' no instance of the wanted sound available.
        'm'' duplicating...
        soundbanklen = soundbanklen + 1
        ReDim Preserve soundbank(1 To soundbanklen) As DSDBUFF
        soundbank(soundbanklen).SoundName = soundbank(m).SoundName
        Set soundbank(soundbanklen).DSBuffer = DSND.DuplicateSoundBuffer(soundbank(m).DSBuffer)
        Snd_IsHere = soundbanklen
        Exit Function
    End If
    
    Snd_IsHere = -1

End Function

Sub Snd_Clear()
'm'' to cleanly unload the sound bank
Dim i As Long

    If ds_Init = True Then
        PrimaryDNSB.Stop
        For i = 1 To soundbanklen
            soundbank(i).DSBuffer.Stop 'if playing...
            Set soundbank(i).DSBuffer = Nothing
        Next i
        Erase soundbank
        soundbanklen = 0
        Set DSND = Nothing
        ds_Init = False
    End If

End Sub

Sub APlusCalc(monster As amonsterT, ByVal monindex As Long)
'm'' base call for A-star pathfinding
Dim ThePath() As tPoint

    'm'' this is always monster to player pathfinding
    t = APlus(monster.X, monster.Y, plr.X, plr.Y, ThePath)
    
    If t Then
        If UBound(monroute) < monindex Then ReDim Preserve monroute(monindex)
        monroute(monindex).Route = ThePath
        monroute(monindex).IndexRoute = 0
    End If
    
End Sub

Private Function APlus(ByVal SX As Long, ByVal SY As Long, ByVal TX As Long, ByVal TY As Long, Path() As tPoint) As Boolean
    'A+ Pathfinding Algorithm:
    'Implementation by Herbert Glarner (herbert.glarner@bluewin.ch)
    'tweaked for VRPG by Massive27
    'Unlimited use for whatever purpose allowed provided that above credits are given.
    'Suggestions and bug reports welcome.
    Dim lMaxList As Long
    Dim lActList As Long
    Dim sCheapCost As Long, lCheapIndex As Long
    Dim sTotalCost As Long
    Dim lCheapX As Long, lCheapY As Long
    Dim lOffX As Long, lOffY As Long
    Dim lTestX As Long, lTestY As Long
    Dim sAdditCost As Long
    Dim lPathPtr As Long
    Dim abGridCopy() As tGrid
    
    'The test program wants to access this grid. For this reason it is defined
    'and initialized globally. Usually one would define and initialize it only
    'in this procedure.
    'The two fields of tGrid can also be merged into the source matrix.
    '   Dim abGridCopy() As tGrid

    
    'For each cell of the grid a bit is defined to hold it's "closed" status
    'and the index to the Open-List.
    'The test program wants to access this grid. For this reason it is defined
    'and initialized globally. Usually one would define and initialize it only
    'in this procedure. (Don't omit here: we need an empty matrix.)
    ReDim abGridCopy(0 To mapx, 0 To mapy) As tGrid
    
    'The starting point is added to the working list. It has no parent (-1).
    'The cost to get here is 0 (we start here). The direct distance enters
    'the Heuristic.
    ReDim grList(0 To 0) As tCell
    With grList(0)
        .X = SX: .Y = SY: .Parent = -1: .Cost = 0
        .Heuristic = Abs(TX - SX) + Abs(TY - SY)
    End With
    
    'Start the algorithm
    Do
        'Get the cell with the lowest Cost+Heuristic. Initialize the cheapest cost
        'with an impossible high value (change as needed). The best found index
        'is set to -1 to indicate "none found".
        sCheapCost = 10000000
        lCheapIndex = -1
        'Check all cells of the list. Initially, there is only the start point,
        'but more will be added soon.
        For lActList = 0 To lMaxList
            'Only check if not closed already.
            If Not grList(lActList).Closed Then
                'If this cells total cost (Cost+Heuristic) is lower than the so
                'far lowest cost, then store this total cost and the cell's index
                'as the so far best found.
                sTotalCost = grList(lActList).Cost + grList(lActList).Heuristic
                If sTotalCost < sCheapCost Then
                    'New cheapest cost found.
                    sCheapCost = sTotalCost: lCheapIndex = lActList
                End If
            End If
        Next lActList
        
        'lCheapIndex contains the cell with the lowest total cost now.
        'If no such cell could be found, all cells were already closed and there
        'is no path at all to the target.
        If lCheapIndex = -1 Then
            'There is no path.
            APlus = False: Exit Function
        End If
        
        'Get the cheapest cell's coordinates
        lCheapX = grList(lCheapIndex).X
        lCheapY = grList(lCheapIndex).Y
        
        'If the best field is the target field, we have found our path.
        If lCheapX = TX And lCheapY = TY Then
            'Path found.
            Exit Do
        End If
       
        
        'Check all immediate neighbors
        For lOffY = -1 To 1
            For lOffX = -1 To 1
                'Ignore the actual field, process all others (8 neighbors).
                If lOffX <> 0 Or lOffY <> 0 Then
                    'Get the neighbor's coordinates.
                    lTestX = lCheapX + lOffX: lTestY = lCheapY + lOffY
                    'Don't test beyond the grid's boundaries.
                    If lTestX >= 0 And lTestX <= mapx And lTestY >= 0 And lTestY <= mapy Then
                        'The cell is within the grid's boundaries.
                        'Make sure the field is accessible. To be accessible,
                        'the cell must have the value as per the function
                        'argument FreeCell (change as needed). Of course, the
                        'target is allowed as well.
                        If map(lTestX, lTestY).blocked = 0 Or map(lTestX, lTestY).monster = 0 Then
                            'The cell is accessible.f
                            'For this we created the "bitmatrix" abGridCopy().
                            If abGridCopy(lTestX, lTestY).ListStat = Unprocessed Then
                                'Register the new cell in the list.
                                lMaxList = lMaxList + 1
                                ReDim Preserve grList(0 To lMaxList) As tCell
                                With grList(lMaxList)
                                    'The parent is where we come from (the cheapest field);
                                    'it's index is registered.
                                    .X = lTestX: .Y = lTestY: .Parent = lCheapIndex
                                    'Additional cost is 1 for othogonal movement, cSqr2 for
                                    'diagonal movement (change if diagonal steps should have
                                    'a different cost).
                                    If Abs(lOffX) + Abs(lOffY) = 1 Then sAdditCost = 10& Else sAdditCost = 14&
                                    'Store cost to get there by summing the actual cell's cost
                                    'and the additional cost.
                                    .Cost = grList(lCheapIndex).Cost + sAdditCost
                                    'Calculate distance to target as the heuristical part
                                    .Heuristic = Abs(TX - lTestX) + Abs(TY - lTestY)
                                End With
                                'Register in the Grid copy as open.
                                abGridCopy(lTestX, lTestY).ListStat = IsOpen
                                'Also register the index to quickly find the element in the
                                '"closed" list.
                                abGridCopy(lTestX, lTestY).Index = lMaxList

                            ElseIf abGridCopy(lTestX, lTestY).ListStat = IsOpen Then
                                'Is the cost to get to this already open field cheaper when using
                                'this path via lTestX/lTestY ?
                                lActList = abGridCopy(lTestX, lTestY).Index
                                sAdditCost = IIf(Abs(lOffX) + Abs(lOffY) = 1, 10&, 14&)
                                If grList(lCheapIndex).Cost + sAdditCost < grList(lActList).Cost Then
                                    'The cost to reach the already open field is lower via the
                                    'actual field.
                                    
                                    'Store new cost
                                    grList(lActList).Cost = grList(lCheapIndex).Cost + sAdditCost
                                    'Store new parent
                                    grList(lActList).Parent = lCheapIndex

                                End If
                            'ElseIf abGridCopy(lTestX, lTestY) = IsClosed Then
                            '   'This cell can be ignored
                            End If
                        End If
                    End If
                End If
            Next lOffX
        Next lOffY
        'Close the just checked cheapest cell.
        grList(lCheapIndex).Closed = True
        abGridCopy(lCheapX, lCheapY).ListStat = IsClosed

    Loop
        
    'We arrive here only when a path was found.
    
    'The path can be found by backtracing from the field TX/TY until SX/SY.
    'The path is traversed in backwards order and stored reversely (!) in
    'the "argument" Path().
    ReDim Path(0 To 0) As tPoint
    lPathPtr = -1
    'lCheapIndex (lCheapX/Y) initially contains the target TX/TY
    Do
        'Store the coordinates of the current cell
        lPathPtr = lPathPtr + 1
        ReDim Preserve Path(0 To lPathPtr) As tPoint
        Path(lPathPtr).X = grList(lCheapIndex).X
        Path(lPathPtr).Y = grList(lCheapIndex).Y
        
        'Follow the parent
        lCheapIndex = grList(lCheapIndex).Parent
    Loop While lCheapIndex <> -1
    
    APlus = True: Exit Function
End Function


Sub AStar(monster As amonsterT)
'm'' A Star algorithm, tuned here to find the player.

Dim Pxp As Long, Pyp As Long 'player X and Y position, cached for speedup calculation
Dim OpenList() As ASTAR_POINT
Dim ClosedList() As ASTAR_POINT
Dim AP As ASTAR_POINT, APcalc As ASTAR_POINT
Dim n As Long, m As Long 'openlist and closedlist index pointer
Dim Cpx As Long, Cpy As Long 'buffer for "calculated position x" and "... y"
Dim MinF As Long, MinF_n As Long

Pxp = plr.X
Pyp = plr.Y

n = 1
m = 0
ReDim OpenList(1) As ASTAR_POINT
OpenList(1).X = monster.X
OpenList(1).Y = monster.Y
OpenList(1).Parent = -1
OpenList(1).H = Abs(Pxp - monster.X) + Abs(Pyp - monster.Y)
OpenList(1).F = OpenList(1).H

Do While n > 0

    'search for the lower F score. tweak : starts from end of openlist
    MinF = 1000000
    MinF_n = -1
    For j = n To 1 Step -1
        If OpenList(n).F < MinF And OpenList(n).Closed = 0 Then
            MinF = OpenList(n).F
            MinF_n = j
        End If
    Next j
    
    If (MinF_n = -1) Then Exit Do 'no path solutions !
    
    AP = OpenList(MinF_n)
    
    'is current point the goal ?
    If AP.X = Pxp And AP.Y = Pyp Then Exit Do 'goal !
    
    
    'add current to closedlist
    m = m + 1
    ReDim Preserve ClosedList(m)
    ClosedList(m) = AP
    

    'calc an star
    For i = 1 To 8
        APcalc.X = monster.X + ASTAR_MATRIX(i).X
        APcalc.Y = monster.Y + ASTAR_MATRIX(i).Y
        
        'is in closedlist ?
        
        'should test if in map bound here
        
        If (map(APcalc.X, APcalc.Y).blocked = 0) Then
            n = n + 1
            APcalc.H = Abs(Pxp - APcalc.X) + Abs(Pyp - APcalc.Y)
            APcalc.g = ASTAR_MATRIX(i).g + AP.g
            APcalc.F = APcalc.H + APcalc.g
            APcalc.Parent = MinF_n
            'register IF NOT ALREADY IN
            ReDim Preserve OpenList(n)
            OpenList(n) = APcalc
        End If
    Next i
        

Loop

'm'' Manhattan distance of each 8 surrounding possible place
Cpx = monster.X - 1
Cpy = monster.Y - 1
If (map(Cpx, Cpy).blocked = 0) Then
End If



End Sub


Sub MoreInit()
'm'' for any initialization required with this updated source of VRPG

'm'' some drawing function seek the hDC of the backbuffer
'm'' let's set it here, once for all

    BackBufHDC = Bridge.Picture1.hDC
    
'm'' A-star pathfinding

    ASTAR_MATRIX(1).X = -1
    ASTAR_MATRIX(1).Y = -1
    ASTAR_MATRIX(1).g = 14
    ASTAR_MATRIX(2).X = 0
    ASTAR_MATRIX(2).Y = -1
    ASTAR_MATRIX(2).g = 10
    ASTAR_MATRIX(3).X = 1
    ASTAR_MATRIX(3).Y = -1
    ASTAR_MATRIX(3).g = 14
    ASTAR_MATRIX(4).X = -1
    ASTAR_MATRIX(4).Y = 0
    ASTAR_MATRIX(4).g = 10
    ASTAR_MATRIX(5).X = 1
    ASTAR_MATRIX(5).Y = 0
    ASTAR_MATRIX(5).g = 10
    ASTAR_MATRIX(6).X = -1
    ASTAR_MATRIX(6).Y = 1
    ASTAR_MATRIX(6).g = 14
    ASTAR_MATRIX(7).X = 0
    ASTAR_MATRIX(7).Y = 1
    ASTAR_MATRIX(7).g = 10
    ASTAR_MATRIX(8).X = 1
    ASTAR_MATRIX(8).Y = 1
    ASTAR_MATRIX(8).g = 14



End Sub

Sub CharLoad()
'm'' from another VRPG source, chars type loading
'm'' for some reason, Duam started to code an handler of 4 players.
'm'' code must be downgraded, each Chars(CurChar). replaced by plr.
'm'' else skill attribution wont work.

ReDim chars(1 To 4) As playertype
'Character initialization
CurChar = 1
For a = 1 To 4
    'chars(a).CharNum = a
    'chars(a).CurActionSlot = 1
    'chars(a).CurActionSet = 1
    'chars(a).TimePoints = 10
    chars(a).bodyname = "body" & roll(6) & ".bmp"
    chars(a).hairname = "hair" & roll(26) & ".bmp"
    chars(a).haircolor = RGB(roll(155), roll(155), roll(155))
    chars(a).Class = getgener("Sorceress", "Amazon", "Valkyrie")
    If chars(a).name = "" Then chars(a).name = getname
    'ApplyClass chars(a), chars(a).Class
    chars(a).level = 1: chars(a).exp = 0: chars(a).expneeded = 600
    chars(a).fatiguemax = greater(50, getend * 30 + 50)
    'setplrskilldescs chars(a)
    
    'chars(a).FaceFile = "face1.bmp"
    'chars(a).NoseFile = "nose1.bmp"
    'chars(a).EyesFile = "eyes1.bmp"
    'chars(a).LipsFile = "lips1.bmp"
    'chars(a).BackHairFile = "backhair1.bmp"
    'chars(a).FrontHairFile = "fronthair1.bmp"
    
    'chars(a).SkinColorLight = RGB(180, 140, 120)
    'chars(a).SkinColorDark = RGB(130, 70, 30)
    
    'chars(a).LipsColor = RGB(235, 10, 15)
    
    'chars(a).HairColorLight = RGB(255, 122, 36)
    'chars(a).HairColorDark = RGB(111, 60, 0)
    
Next a

End Sub

Sub Quitting()
'm'' from Form1_QueryUnload()
'improved to avoid many errors

Form1.Timer1.Enabled = False
'KillTimer Form1.hwnd, API_Timer_Handle 'm'' SetTimer() isnt called
Debugger.Snd_Clear 'm'' DX7 sound stopping

nodraw = 1
endingprog = 1
ClearSprites2
'm''BASS_Free
'm''Form1.DMC1.TerminateBASS


For a = 1 To 100
DoEvents
Next a

On Local Error GoTo errmgr 'handle any exception
If Not Dir("curgame.dat") = "" Then Kill "curgame.dat"
If Not Dir("plrdat.tmp") = "" Then Kill "plrdat.tmp"
If Not Dir(App.Path & "\" & plr.name & "\VTDATA*.*") = "" Then Kill App.Path & "\" & plr.name & "\VTDATA*.*"
If Not Dir(App.Path & "\VTDATA*.*") = "" Then Kill App.Path & "\VTDATA*.*"


'm'' Duam forgot to freed the directx stuff.
Set DXLib.picBuffer = Nothing
Set DXLib.Primary = Nothing
Set DXLib.picBuffer2 = Nothing

Dim frm As Form
For Each frm In Forms
     Unload frm
Next frm

Exit Sub
errmgr:
If Err.Number = 70 Then
    MsgBox "Some temporary files couldn't be deleted. error 70. Please remove them manually.", vbInformation + vbOKOnly
    Err.clear
Else
    MsgBox "Unhandled error code " & Err.Number & ", description : " & Err.Description & " in Quittin(). Please report!", vbInformation + vbOKOnly
    Err.clear
    Resume Next
End If
End Sub
