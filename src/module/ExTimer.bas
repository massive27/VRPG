Attribute VB_Name = "ModTimer"

Global Const gbytMSPerFrame = 12     'This should give us approx. 40fps
Dim lngLastFrame As Long    'What was the TickCount last frame?
Dim lngLastSecond As Long   'What was the TickCount last second?
Dim bytFPSCurrent As Byte   'How many frames have been displayed since lngLastSecond
Dim bytFPSElapsed As Byte   'How many frames were displayed in the second that just elapsed


Dim mlngTimer As Long       'Holds system time since last frame was displayed
Public mlngElapsed As Long     'MS of elapsed time since last frame
Dim mlngFrameTimer As Long  'Stores the time since the last FPS display
Public mintFPS As Long         'An FPS storage variable
Dim mintFPSCounter As Long  'An FPS counter


Declare Function GetTickCount Lib "kernel32" () As Long

Public Function Ticker(Optional Reset As Boolean) As Long

Static oldtime As Long

    If Reset Then
        'reset timer and return zero
        oldtime = GetTickCount()
        Ticker = 0
    Else
        'return difference between oldtime and current time
        Ticker = GetTickCount() - oldtime
    End If

End Function

Public Function Time() As Long

    'Simply returns the system tickcount
    Time = GetTickCount()

End Function




'*****************************************************
' Purpose:  Store the time elapsed since last call and
'           only exit the subroutine when it is time
'           for the next frame to be displayed
'           (dependent on the value of gbytMSPerFrame
'           global constant)
'*****************************************************

Public Sub FrameRate()

    'Ensure that we do not exceed our frame rate
    Do While GetTickCount - lngLastFrame < gbytMSPerFrame
        DoEvents
    Loop

    'Reset the frame tickcount
    lngLastFrame = GetTickCount

    'Add one to the FPS count
    bytFPSCurrent = bytFPSCurrent + 1

    'Check if it is time to update the FPS
    If GetTickCount - lngLastSecond > 1000 Then
        bytFPSElapsed = bytFPSCurrent   'Update the static FPS variable
        bytFPSCurrent = 0               'Reset the incrementing FPS variable
        lngLastSecond = GetTickCount    'Reset the "LastSecond" tickcount
    End If

End Sub

'*****************************************************
' Purpose:  Return the number of frames that were
'           displayed during the last second
'*****************************************************

Public Function FPS() As Integer

    'Return the FPS
    FPS = bytFPSElapsed

End Function





Public Sub PPSTimer()  'PxelsPerSecondTimer

    'Determine the time that has elapsed since the last frame was displayed
    mlngElapsed = GetTickCount() - mlngTimer
    'Reset the general timer
    mlngTimer = GetTickCount()
    'Check if one second has elapsed
    If GetTickCount() - mlngFrameTimer >= 1000 Then
        'Set the FPS storage var, and reset the FPS counter/timer
        mintFPS = mintFPSCounter + 1
        mintFPSCounter = 0
        mlngFrameTimer = GetTickCount()
    Else
        'If a second hasn't elapsed, add to the FPS counter
        mintFPSCounter = mintFPSCounter + 1
    End If
    'Show the FPS
    'frmShip.Caption = "FPS: " & mintFPS

End Sub
