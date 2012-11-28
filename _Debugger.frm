VERSION 5.00
Begin VB.Form LegacyDebugger 
   Caption         =   "Legacy debugger for VRPG"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Copy to clipboard"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "LegacyDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Dim buf As String

    For i = 0 To List1.ListCount
        buf = buf & vbCrLf & List1.List(i)
    Next i
    
    Clipboard.Clear
    Clipboard.SetText buf
    
    List1.AddItem "Copied to clipboard " & i & " lines."
    
End Sub

Private Sub Form_Load()
Dim CPath As String
List1.Clear

CPath = App.Path & "\VRPG2.exe"

'Are we in VRPG folder ?
If Dir(CPath) = "" Then
    List1.AddItem "VRPG2.exe not found. Result of the debugger may be unusable.", vbInformation + vbOKCancel
Else
    List1.AddItem "VRPG2.exe found in " & CPath
    'choose this dir as active directory (well... it should be done already)
    ChDir App.Path
End If


'VRPG use its own application dir as temporary fold
'there may be issue, so here we test file I/O
List1.AddItem "Testing file I/O for temporary data..."
TmpFold_Check App.Path
List1.AddItem "TmpFold_Check(" & App.Path & ") done."
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
            'Debugger.Quitting 'only in vrpg2.exe
            End
        Else
            MsgR = InputBox("Please write another folder", "Temporary folder", ThisDir)
            If MsgR = vbCancel Then
                'Debugger.Quitting 'only in vrpg2.exe
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
    
    List1.AddItem "TmpFold_Check(), vbcode = " & a & ", vbdesc = " & Err.Description & ", dstep = " & dstep
    
    MsgR = MsgBox("Error : in TmpFold_Check(), vbcode = " & a & ", vbdesc = " & Err.Description & ", dstep = " & dstep & ". Please report!", vbExclamation + vbAbortRetryIgnore, "VRPG Debug")
    
    
        
    If MsgR = vbIgnore Then
        List1.AddItem "Action : ignore..."
        Err.Clear
        Resume Next
    ElseIf MsgR = vbRetry Then
        List1.AddItem "Action : retry..."
        Err.Clear
        Resume
    Else
        List1.AddItem "Action : cancel."
        Err.Clear
        Exit Sub
    End If

End Sub

Private Sub Form_Resize()
'moving controls
Command1.Left = Me.Width - 1575
Command1.Top = Me.Height - 930
Command2.Top = Me.Height - 930
List1.Width = Me.Width - 360
List1.Height = Me.Height - 1140
End Sub
