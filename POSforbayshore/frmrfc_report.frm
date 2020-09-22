VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrfc_report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return from customer report form"
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
   Begin VB.PictureBox picprint 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1065
      ScaleWidth      =   3105
      TabIndex        =   20
      Top             =   9120
      Visible         =   0   'False
      Width           =   3135
      Begin Project1.chameleonButton cmdprint 
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "PRINT"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12648447
         FCOL            =   16711680
      End
      Begin Project1.chameleonButton chameleonButton2 
         Height          =   495
         Left            =   1800
         TabIndex        =   22
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "CANCEL"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12648447
         FCOL            =   16711680
      End
   End
   Begin VB.PictureBox picinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   240
      ScaleHeight     =   9225
      ScaleWidth      =   14745
      TabIndex        =   0
      Top             =   240
      Width           =   14775
      Begin VB.PictureBox picmonthly 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   5400
         ScaleHeight     =   1185
         ScaleWidth      =   9105
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   9135
         Begin VB.ComboBox mm 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "frmrfc_report.frx":0000
            Left            =   1200
            List            =   "frmrfc_report.frx":0028
            TabIndex        =   12
            Text            =   "mm"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox yyyy 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2880
            MaxLength       =   4
            TabIndex        =   11
            Text            =   "yyyy"
            Top             =   360
            Width           =   2535
         End
         Begin Project1.chameleonButton cmdenter 
            Height          =   375
            Left            =   5520
            TabIndex        =   13
            Top             =   360
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "ENTER"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            FCOL            =   0
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   14
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.PictureBox picdaily 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   5400
         ScaleHeight     =   1185
         ScaleWidth      =   9105
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   9135
         Begin VB.TextBox txtyyyy 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "yyyy"
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cbodd 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "frmrfc_report.frx":005C
            Left            =   2880
            List            =   "frmrfc_report.frx":00BD
            TabIndex        =   6
            Text            =   "dd"
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox cbomm 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "frmrfc_report.frx":013D
            Left            =   1200
            List            =   "frmrfc_report.frx":0165
            TabIndex        =   5
            Text            =   "mm"
            Top             =   360
            Width           =   1695
         End
         Begin Project1.chameleonButton chameleonButton1 
            Height          =   375
            Left            =   6840
            TabIndex        =   8
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "ENTER"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            FCOL            =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   9
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   240
         ScaleHeight     =   1185
         ScaleWidth      =   5025
         TabIndex        =   1
         Top             =   240
         Width           =   5055
         Begin Project1.chameleonButton cmddreport 
            Height          =   495
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            BTYPE           =   5
            TX              =   "DAILY REPORT"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            FCOL            =   8454143
         End
         Begin Project1.chameleonButton cmdmreport 
            Height          =   495
            Left            =   2760
            TabIndex        =   3
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            BTYPE           =   5
            TX              =   "MONTHLY REPORT"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            FCOL            =   8454143
         End
      End
      Begin MSComctlLib.ListView lst 
         Height          =   5895
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   10398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   16711680
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Qty"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Price"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label txttot 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   7440
         TabIndex        =   17
         Top             =   7920
         Width           =   7095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Total          :Php"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   240
         TabIndex        =   16
         Top             =   7920
         Width           =   7215
      End
   End
   Begin Project1.chameleonButton cmdprocess 
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   10200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Process"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      FCOL            =   0
   End
   Begin Project1.chameleonButton cmdout 
      Height          =   495
      Left            =   1920
      TabIndex        =   19
      Top             =   10200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Out"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      FCOL            =   0
   End
End
Attribute VB_Name = "frmrfc_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdout_Click()
Unload Me
mdimain.Show
End Sub

Private Sub chameleonButton1_Click()
rst.Open "Select * from tblrfc", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
Dim X
While rst.EOF = False
If cbomm = rst!mm And cbodd = rst!dd And txtyyyy = rst!yyyy Then
lst.ListItems.Add , , rst!qty
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!price
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!amount
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!mm + "/" + rst!dd + "/" + rst!yyyy
X = 1
txttot = Val(txttot) + Val(rst!amount)
cmdprocess.Enabled = True
End If
rst.MoveNext
Wend
rst.Close
If X <> 1 Then
MsgBox "Invalid month or year!", vbCritical, "System Message"
cbomm = ""
txtyyyy = ""
cbomm.SetFocus
End If

End Sub

Private Sub chameleonButton2_Click()
If cmdprint.Enabled = False Then
    rst.Close
    Call clear
    picprint.Visible = False
    picinfo.Enabled = True
    cmdout.Enabled = True
Else
    Call clear
    picprint.Visible = False
    picinfo.Enabled = True
    cmdout.Enabled = True
End If

End Sub

Private Sub cmddreport_Click()
picmonthly.Visible = False
picdaily.Visible = True
cbomm = "mm"
cbodd = "dd"
txtyyyy = "yyyy"

End Sub

Private Sub cmdenter_Click()
rst.Open "Select * from tblrfc", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
Dim X
While rst.EOF = False
If mm = rst!mm And yyyy = rst!yyyy Then
lst.ListItems.Add , , rst!qty
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!price
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!amount
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!mm + "/" + rst!dd + "/" + rst!yyyy
X = 1
txttot = Val(txttot) + Val(rst!amount)
cmdprocess.Enabled = True
End If
rst.MoveNext
Wend
rst.Close
If X <> 1 Then
MsgBox "Invalid month or year!", vbCritical, "System Message"
mm = ""
yyyy = ""
mm.SetFocus
End If

End Sub

Private Sub cmdmreport_Click()
picmonthly.Visible = True
picdaily.Visible = False
mm = "mm"
yyyy = "yyyy"

End Sub


Private Sub cmdprint_Click()
If picdaily.Visible = True Then
rst.Open "Select * from tblrfc where mm='" & cbomm & "' and dd='" & cbodd & "' and yyyy='" & txtyyyy & "'", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False

Set dtprfc_report.DataSource = rst
dtprfc_report.Sections("Section2").Controls.Item("lbldaily").Caption = "Daily"
dtprfc_report.Sections("Section5").Controls.Item("lbluser").Caption = mdimain.lblname
dtprfc_report.Sections("Section5").Controls.Item("lbldate").Caption = Date
dtprfc_report.Sections("Section5").Controls.Item("lbltot").Caption = txttot
dtprfc_report.Show
cmdprint.Enabled = False
rst.MoveNext
Wend
Else
rst.Open "Select * from tblrfc where mm='" & mm & "' and yyyy='" & yyyy & "'", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False

Set dtprfc_report.DataSource = rst
dtprfc_report.Sections("Section2").Controls.Item("lbldaily").Caption = "Monthly"
dtprfc_report.Sections("Section5").Controls.Item("lbluser").Caption = mdimain.lblname
dtprfc_report.Sections("Section5").Controls.Item("lbldate").Caption = Date
dtprfc_report.Sections("Section5").Controls.Item("lbltot").Caption = txttot
dtprfc_report.Show
cmdprint.Enabled = False
rst.MoveNext
Wend
End If


End Sub

Private Sub cmdprocess_Click()
picinfo.Enabled = False
picprint.Visible = True
cmdprocess.Enabled = False
cmdprint.Enabled = True
cmdout.Enabled = False

End Sub

Function clear()
cbomm = ""
cbodd = ""
txtyyyy = ""
lst.ListItems.clear
mm = ""
dd = ""
yyyy = ""
txttot = ""
End Function



Private Sub txtyyyy_Click()
txtyyyy = ""
txtyyyy.SetFocus

End Sub

Private Sub yyyy_Click()
yyyy = ""
yyyy.SetFocus
End Sub
