VERSION 5.00
Begin VB.Form frmPack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VRPG Pack utility rev.a"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data"
      Height          =   2175
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "# of files :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "size :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "pack name :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Extract"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Attempts to extract the file selected on the left, else extacts all files"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Create a genuine pack from all files found in a folder"
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Opens a pack file"
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim MyPackFile As cRes

Private Const PFilt = "Pack file (*.pkf)|*.pkf|all files (*.*)|*.*|"

Private Sub Command1_Click()
Dim Target As String
Dim FileList() As String
Dim i As Long

    Target = ComDlg.SelectAfile(Me.hWnd, "Open a pack file", PFilt)
    If LenB(Target) > 0 Then
        Label4.Caption = ComDlg.GetFileName(Target)
        Set MyPackFile = New cRes
        
        If MyPackFile.SetPackFile(Target) Then
            Label6.Caption = MyPackFile.NumResource
            Call MyPackFile.GetAllFiles(FileList)
            List1.Clear
            For i = 1 To UBound(FileList)
                List1.AddItem FileList(i)
            Next i
        End If
        
    End If
        
End Sub

Private Sub Command2_Click()
Dim source As String
Dim Target As String
    source = ComDlg.SelectADir(Me.hWnd, "Select dir to pack up")
    If LenB(source) > 0 Then
        Target = ComDlg.SaveAfile(Me.hWnd, "Sets target pack file", PFilt)
        If LenB(Target) > 0 Then
            Set MyPackFile = New cRes
            If MyPackFile.BuildPack(source, Target, True) Then
                MsgBox "Done!"
            End If
        End If
    End If
    
End Sub

Private Sub Command3_Click()
Dim TDir As String
Dim i As Long

    If MyPackFile Is Nothing Then Exit Sub
    
    TDir = ComDlg.SelectADir(Me.hWnd, "Select dir to unpack to")
    If LenB(TDir) > 0 Then
        If MyPackFile.SetTempFolder(TDir) Then
            'if no file were selected in the listbox, extracts all
            If List1.SelCount = 0 Then
                If MyPackFile.ExtractAll() Then
                    MsgBox "Files extracted"
                Else
                    MsgBox "Fail to extract files"
                End If
            Else
            'extract selected files
                For i = 0 To List1.ListCount - 1
                    If List1.Selected(i) Then
                        If MyPackFile.GetResToFile(List1.List(i)) <> "" Then
                            MsgBox "Extraction item " & i & " successful"
                        Else
                            MsgBox "Extraction failed"
                        End If
                    End If
                Next i
            End If
        Else
            MsgBox "Target folder unreachable"
        End If
    End If
        
    
End Sub

Private Sub Command4_Click()
If MyPackFile Is Nothing Then
Else
    Set MyPackFile = Nothing
End If
Unload Me
End
End Sub
