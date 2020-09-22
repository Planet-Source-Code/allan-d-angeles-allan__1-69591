VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmhelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Help"
   ClientHeight    =   11010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser br 
      Height          =   11055
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   12255
      ExtentX         =   21616
      ExtentY         =   19500
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   11055
      Left            =   0
      ScaleHeight     =   11025
      ScaleWidth      =   2985
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin Project1.chameleonButton chameleonButton1 
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   10320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12648447
         FCOL            =   0
      End
      Begin VB.Label lblr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reports Module"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label lblt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Module"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label lblf 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "File Maintenance Module"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Shortcut Key"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000010&
         Caption         =   "Help Topic"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
Unload Me
mdimain.Show
End Sub

Private Sub Form_Load()
br.Navigate "C:\POSforbayshore\mp.html"
End Sub

Private Sub lblf_Click()
br.Navigate "C:\POSforbayshore\file.html"
End Sub

Private Sub lblf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblf.BackStyle = 1
lblf.BorderStyle = 1
End Sub

Private Sub lblp_Click()
br.Navigate "C:\POSforbayshore\pm.html"
End Sub

Private Sub lblp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblp.BackStyle = 1
lblp.BorderStyle = 1
End Sub



Private Sub lblr_Click()
br.Navigate "C:\POSforbayshore\rm.html"
End Sub

Private Sub lblr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblr.BackStyle = 1
lblr.BorderStyle = 1
End Sub

Private Sub lbls_Click()
br.Navigate "C:\POSforbayshore\shortcut.html"
End Sub

Private Sub lbls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbls.BackStyle = 1
lbls.BorderStyle = 1
End Sub



Private Sub lblt_Click()
br.Navigate "C:\POSforbayshore\tm.html"
End Sub

Private Sub lblt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblt.BackStyle = 1
lblt.BorderStyle = 1
End Sub



Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblt.BackStyle = 0
lblt.BorderStyle = 0
lblf.BackStyle = 0
lblf.BorderStyle = 0
lbls.BackStyle = 0
lbls.BorderStyle = 0
lblr.BackStyle = 0
lblr.BorderStyle = 0
End Sub
