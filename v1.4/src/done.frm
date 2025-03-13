VERSION 5.00
Begin VB.Form done 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All tasks finished!"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Continue"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What do you want to do?"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Optimization complete!"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "done.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "done"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdExit_Click()
    Unload Main
    Unload splash
    Unload Me
End Sub

