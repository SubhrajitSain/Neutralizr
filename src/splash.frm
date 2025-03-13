VERSION 5.00
Begin VB.Form splash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame splash_frame 
      BackColor       =   &H00FFFFFF&
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7305
      Begin VB.Timer tmrSplash 
         Interval        =   2000
         Left            =   6720
         Top             =   240
      End
      Begin VB.Label lblWarning 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Image imgLogo 
         Height          =   1785
         Left            =   240
         Picture         =   "splash.frx":000C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(C) Subhrajit Sain (ANW), 2025"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All rights reserved."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version 1.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5580
         TabIndex        =   3
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "For Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4875
         TabIndex        =   4
         Top             =   1920
         Width           =   1980
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         Caption         =   " Neutralizr "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   2400
         TabIndex        =   6
         Top             =   1200
         Width           =   2250
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ANW's"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   360
         Left            =   2400
         TabIndex        =   5
         Top             =   840
         Width           =   1125
      End
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Version 1.3"
    lblProductName.Caption = " Neutralizr "
End Sub

Private Sub tmrSplash_Timer()
    tmrSplash.Enabled = False
    splash.Hide
    Main.Show
    Unload Me
End Sub
