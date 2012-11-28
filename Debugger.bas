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
'm'' 2012-05-04
'm'' - skilltree.frm code downgraded to be compatible with main code
'm'' - modified TitleScreen property to avoid resize
'm'' - renamed curversion to 2.14 Legacy
'm'' - added File I/O check and data.pak available check for compatibility troubleshoot
'm'' - fixed attack animation and alike (bijdang) to render properly instead of first frame
'm'' - improved cSpriteBitmaps.CreateFromFile function for faster exec.
'm'' - gotomap function viewed


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
        Err.clear
        'm'' create an ampty buffer instead...
        Set Snd_Loadme = DSND.CreateSoundBuffer(soundbufferDsc, WvFrEx)
        'm'' TODO : loading wave file in memory and copying in this buffer
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



Sub MoreInit()
'm'' some drawing function seek the hDC of the backbuffer
'm'' let's set it here, once for all

    BackBufHDC = Bridge.Picture1.hDC

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
KillTimer Form1.hwnd, API_Timer_Handle
Debugger.Snd_Clear 'm'' DX7 sound stopping

nodraw = 1
endingprog = 1
ClearSprites2
'm''BASS_Free
'm''Form1.DMC1.TerminateBASS


For a = 1 To 100
DoEvents
Next a

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

End Sub
