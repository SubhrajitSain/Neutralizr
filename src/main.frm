VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Neutralizr"
   ClientHeight    =   2610
   ClientLeft      =   5760
   ClientTop       =   3405
   ClientWidth     =   5250
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRunTasks 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Optimize!"
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox chkShutdown 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check4"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox chkKillApps 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check3"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox chkPrefetch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check2"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chkTemp 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Visit my site at: anw.is-a.dev"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(c) Subhrajit Sain (ANW), 2025."
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"main.frx":08CA
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "main.frx":097A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shutdown computer after optimizing"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kill major known apps"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete prefetch files (apps may load slower for the first time)"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear temporary files"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   4815
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Main.Hide
    splash.Show
    DoEvents
End Sub
' Clean Temporary Files (both %TEMP% and %TMP%)
Sub CleanTempFiles()
    CleanTempPath Environ("TEMP")
    CleanTempPath Environ("TMP")
End Sub

Sub CleanTempPath(TempPath As String)
    Dim FileSpec As String
    Dim FileName As String

    TempPath = TempPath & "\"
    FileSpec = TempPath & "*.*"
    FileName = Dir(FileSpec)

    Do While FileName <> ""
        If (GetAttr(TempPath & FileName) And vbDirectory) = 0 Then
            On Error Resume Next
            Kill TempPath & FileName
            On Error GoTo 0
        End If
        FileName = Dir
    Loop
End Sub

' Clear Prefetch Files
Sub ClearPrefetch()
    Dim PrefetchPath As String
    Dim FileSpec As String
    Dim FileName As String

    PrefetchPath = Environ("WINDIR") & "\Prefetch\"
    FileSpec = PrefetchPath & "*.*"
    FileName = Dir(FileSpec)

    Do While FileName <> ""
        If (GetAttr(PrefetchPath & FileName) And vbDirectory) = 0 Then
            On Error Resume Next
            Kill PrefetchPath & FileName
            On Error GoTo 0
        End If
        FileName = Dir
    Loop
End Sub

' Kill Major Applications
Sub KillApplications()
    Dim AppList() As String
    Dim i As Integer

    AppList = Split("chrome.exe,firefox.exe,iexplore.exe,msedge.exe,wmplayer.exe,code.exe,opera.exe,launcher.exe", ",")

    For i = LBound(AppList) To UBound(AppList)
        On Error Resume Next
        Shell "taskkill /F /IM " & AppList(i), vbHide
        On Error GoTo 0
    Next i
End Sub

Private Sub cmdRunTasks_Click()
    show_done = False
    If chkTemp.Value = vbChecked Then
        CleanTempFiles
        show_done = True
    End If

    If chkPrefetch.Value = vbChecked Then
        ClearPrefetch
        show_done = True
    End If

    If chkKillApps.Value = vbChecked Then
        KillApplications
        show_done = True
    End If
    
    ' Add Shutdown Logic
    If chkShutdown.Value = vbChecked Then
        Shell "shutdown -s -t 0", vbHide ' Shutdown immediately
        show_done = False
    End If
    
    If show_done = True Then
        done_msg = MsgBox("Optimization completed!", vbOKOnly + vbInformation, "Neutralizr")
    End If
End Sub
