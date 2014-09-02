Attribute VB_Name = "Sound"
Private isloaded As Byte
'Private DMC1 As DMC
Public soundoff As Byte

Sub playsound(ByVal filen)

If Dir(filen) = "" Then filen = getfile(filen, "Data.pak", , 1)

If Not isloaded = 1 Then loaddmc: isloaded = 1
If soundoff = 1 Then Exit Sub

'm'' warps to DX7

Debugger.Snd_Play filen
Exit Sub 'm''

Static chanl As Byte
chanl = chanl + 1
'Debug.Print "Playing Sample:" & App.Path & "\" & filen
If chanl > 8 Then chanl = 1
'DoEvents
'Form1.DMC1.StopSample
'Form1.DMC1.SampleChanToModify = chanl

'Debug.Print "Playing Sample:" & App.Path & "\" & filen

'If Not Form1.DMC1.SampleIsActive = True Then
'    Form1.DMC1.OpenSample App.Path & "\" & filen
'    Form1.DMC1.PlaySample False
'End If

'If Form1.DMC1.Error = True Then Form1.DMC1.StopSample
'If Form1.DMC1.Info_HWSampSlotsFree <= 0 Then Exit Sub
'm''Form1.DMC1.AutoPlaySample App.Path & "\" & filen, False
'If Form1.DMC1.SampleIsActive = False Then Form1.DMC1.CloseSample
'If Form1.DMC1.Error = True Then Form1.DMC1.CloseSample
'DoEvents
End Sub

Sub playmusic(filen)

'm'' playmusic function is inhibited. can safely be desactivated

If Not isloaded Then loaddmc
If soundon = 0 Then Exit Sub
''On Error GoTo 5
Debug.Print "Playing " & filen
'm'' Form1.DMC1.OpenModule filen, True
'm'' Form1.DMC1.PlayModule


'Music.WaveIndex = 5
'Music.FilenameRead = filen
'Music.Play = 5
5
End Sub

Sub loaddmc()

'm'' wraps to DX7
Debugger.Snd_Init Form1.hwnd
Exit Sub 'm''

'm'' Form1.DMC1.InitBASS Form1.hwnd, 22050, False, False
'Form1.DMC1.SampleVol = 10

End Sub

Sub stopsounds()

'm'' stop all DX7 sounds.
Debugger.Snd_Stop
Exit Sub

For a = 1 To 16
'm''Form1.DMC1.SampleChanToModify = a
'Form1.DMC1.SampleVol = 10
'm''Form1.DMC1.PauseSample
'm''Form1.DMC1.StopSample
Next a

End Sub

Function txtnum(num As Integer, Optional numtxt As String = "0") As String

txt = ""

For a = 1 To num
    txt = txt & numtxt
Next a
txtnum = txt
End Function
