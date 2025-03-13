VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Neutralizr"
   ClientHeight    =   3000
   ClientLeft      =   5745
   ClientTop       =   3390
   ClientWidth     =   5505
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRunTasks 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Optimize!"
      Height          =   375
      Left            =   4200
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
      Caption         =   "Hello"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tasks:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "   This project is open-source: https://github.com/SubhrajitSain/Neutralizr"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   2760
      Width           =   5535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "   Visit my site at: anw.is-a.dev"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "   (c) Subhrajit Sain (ANW), 2025."
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"main.frx":08CA
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   4695
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
      Caption         =   "Shutdown computer after optimization has been completed"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kill major known apps"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete prefetch files"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear temporary files"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   5055
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
    Dim FullPath As String

    TempPath = TempPath & "\"
    FileSpec = TempPath & "*.*"
    FileName = Dir(FileSpec, vbDirectory) ' Include directories

    Do While FileName <> ""
        If FileName <> "." And FileName <> ".." Then ' Avoid current and parent directory
            FullPath = TempPath & FileName
            If (GetAttr(FullPath) And vbDirectory) = vbDirectory Then ' Check if it's a directory
                DeleteDirectory FullPath ' Recursively delete directory
            Else
                On Error Resume Next
                Kill FullPath ' Delete file
                On Error GoTo 0
            End If
        End If
        On Error Resume Next
        If Dir = "" Then Exit Do 'prevent error 5.
        If Err.Number <> 0 Then Exit Do
        FileName = Dir ' Get next file/directory
        If FileName = "" Then Exit Do
    Loop
End Sub
Sub KillApplications()
    Dim AppList() As String
    Dim i As Integer

    AppList = Split("chrome.exe,firefox.exe,iexplore.exe,msedge.exe,notepad.exe,wmplayer.exe", ",") ' Add more apps

    For i = LBound(AppList) To UBound(AppList)
        On Error Resume Next
        Shell "taskkill /F /IM " & AppList(i), vbHide
        On Error GoTo 0
    Next i
End Sub

Sub ClearPrefetch()
    Dim PrefetchPath As String
    Dim FileSpec As String
    Dim FileName As String
    Dim FullPath As String

    PrefetchPath = Environ("WINDIR") & "\Prefetch\"
    FileSpec = PrefetchPath & "*.*"
    FileName = Dir(FileSpec, vbDirectory) ' Include directories

    Do While FileName <> ""
        If FileName <> "." And FileName <> ".." Then ' Avoid current and parent directory
            FullPath = PrefetchPath & FileName
            If (GetAttr(FullPath) And vbDirectory) = vbDirectory Then ' Check if it's a directory
                DeleteDirectory FullPath ' Recursively delete directory
            Else
                On Error Resume Next
                Kill FullPath ' Delete file
                On Error GoTo 0
            End If
        End If
        On Error Resume Next
        If Dir = "" Then Exit Do 'prevent error 5.
        If Err.Number <> 0 Then Exit Do
        FileName = Dir ' Get next file/directory
        If FileName = "" Then Exit Do
    Loop
End Sub
Sub DeleteDirectory(DirPath As String)
    Dim FileSpec As String
    Dim FileName As String
    Dim FullPath As String

    FileSpec = DirPath & "\*.*"
    FileName = Dir(FileSpec, vbDirectory)

    Do While FileName <> ""
        If FileName <> "." And FileName <> ".." Then
            FullPath = DirPath & "\" & FileName
            If (GetAttr(FullPath) And vbDirectory) = vbDirectory Then
                DeleteDirectory FullPath ' Recursive call
            Else
                On Error Resume Next
                Kill FullPath ' Delete file
                On Error GoTo 0
            End If
        End If
        FileName = Dir ' Get next file/directory
        If FileName = "" Then Exit Do 'prevent error 5.
    Loop
    On Error Resume Next
    RmDir DirPath ' Remove the empty directory
    On Error GoTo 0
End Sub
Private Sub cmdRunTasks_Click()
    show_done = False
    If chkTemp.Value = vbChecked Then
        ' runs 5 times to make sure everything is clean. same for prefetch.
        CleanTempFiles
        CleanTempFiles
        CleanTempFiles
        CleanTempFiles
        CleanTempFiles
        show_done = True
    End If

    If chkPrefetch.Value = vbChecked Then
        ClearPrefetch
        ClearPrefetch
        ClearPrefetch
        ClearPrefetch
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
        done.Show vbModal
    End If
End Sub
