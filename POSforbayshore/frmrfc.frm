VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrfc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return from customer form"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   240
      ScaleHeight     =   9705
      ScaleWidth      =   14745
      TabIndex        =   0
      Top             =   360
      Width           =   14775
      Begin VB.TextBox txttot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1830
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   7800
         Width           =   9495
      End
      Begin VB.PictureBox piclist 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   1800
         ScaleHeight     =   3465
         ScaleWidth      =   5025
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   5055
         Begin MSComctlLib.ListView lst 
            Height          =   2775
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   4895
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
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Unit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Pcode"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Desc"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Price"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Amount"
               Object.Width           =   2540
            EndProperty
         End
         Begin Project1.chameleonButton return 
            Height          =   375
            Left            =   2760
            TabIndex        =   14
            Top             =   3000
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   5
            tx              =   "&Return"
            enab            =   -1  'True
            font            =   "frmrfc.frx":0000
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin Project1.chameleonButton cancel 
            Height          =   375
            Left            =   3960
            TabIndex        =   15
            Top             =   3000
            Width           =   975
            _extentx        =   1720
            _extenty        =   661
            btype           =   5
            tx              =   "&Cancel"
            enab            =   -1  'True
            font            =   "frmrfc.frx":002C
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
      End
      Begin VB.TextBox transid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtdate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox txttransid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   2895
      End
      Begin Project1.chameleonButton cmdenter 
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   720
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
         btype           =   5
         tx              =   "&ENTER"
         enab            =   -1  'True
         font            =   "frmrfc.frx":0058
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   14215660
         fcol            =   0
      End
      Begin Project1.chameleonButton cmdcancel 
         Height          =   375
         Left            =   5880
         TabIndex        =   10
         Top             =   720
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         btype           =   5
         tx              =   "&CANCEL"
         enab            =   -1  'True
         font            =   "frmrfc.frx":0084
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   14215660
         fcol            =   0
      End
      Begin MSComctlLib.ListView lstlist 
         Height          =   6255
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   11033
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
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pcode"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Desc"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Refund"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   7560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Return Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin Project1.chameleonButton cmdnew 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   10320
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      btype           =   3
      tx              =   "&New Trans."
      enab            =   -1  'True
      font            =   "frmrfc.frx":00B0
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      fcol            =   0
   End
   Begin Project1.chameleonButton cmdprocess 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   10320
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      btype           =   3
      tx              =   "&Process"
      enab            =   0   'False
      font            =   "frmrfc.frx":00DC
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      fcol            =   0
   End
   Begin Project1.chameleonButton cmdout 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   10320
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      btype           =   3
      tx              =   "&Out"
      enab            =   -1  'True
      font            =   "frmrfc.frx":0108
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      fcol            =   0
   End
End
Attribute VB_Name = "frmrfc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cancel_Click()
piclist.Visible = False
End Sub

Private Sub cmdcancel_Click()
piclist.Visible = False
transid = ""
cmdenter.Enabled = True
End Sub

Private Sub cmdnew_Click()
picinfo.Enabled = True
rst.Open "Select * from tblrfc", con, adOpenDynamic, adLockOptimistic
rst.MoveLast
txttransid = "RFC" + Format(Val(Right(rst!transid, 5)) + 1, "0000#")
rst.Close
transid.SetFocus
cmdenter.Enabled = True
End Sub

Private Sub cmdout_Click()
Unload Me
mdimain.Show
End Sub

Private Sub cmdenter_Click()
Dim x
rst.Open "Select * from tblsales", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
If transid = rst!transid Then
piclist.Visible = True
cmdenter.Enabled = False
lst.ListItems.Add , , rst!qty
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!unit
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!pcode
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!price
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!amount
x = 1
End If
rst.MoveNext
Wend
rst.Close
If x <> 1 Then
MsgBox "Invalid transaction id!", vbCritical, "System Message"
transid = ""
transid.SetFocus
End If
End Sub

Private Sub cmdprocess_Click()
picinfo.Enabled = False
MsgBox "Process Complete!", vbInformation, "System Message"
cmdprocess.Enabled = False
transid = ""
lst.ListItems.clear
lstlist.ListItems.clear
txttot = ""
cmdnew.SetFocus
End Sub

Private Sub Form_Load()
txtdate = Format(Date, "mm/dd/yyyy")
End Sub
Function reload()
rst.Open "Select * from tblrfc", con, adOpenDynamic, adLockOptimistic
lstlist.ListItems.clear
While rst.EOF = False
If txttransid = rst!transid Then
lstlist.ListItems.Add , , rst!qty
lstlist.ListItems(lstlist.ListItems.Count).ListSubItems.Add , , rst!pcode
lstlist.ListItems(lstlist.ListItems.Count).ListSubItems.Add , , rst!Desc
lstlist.ListItems(lstlist.ListItems.Count).ListSubItems.Add , , rst!price
lstlist.ListItems(lstlist.ListItems.Count).ListSubItems.Add , , rst!amount
End If
rst.MoveNext
Wend
rst.Close
End Function



Private Sub lst_DblClick()
return_Click
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then return_Click
End Sub

Private Sub return_Click()
rst.Open "Select * from tblrfc", con, adOpenDynamic, adLockPessimistic
rst.AddNew

rst!transid = txttransid
rst!qty = lst.SelectedItem
rst!pcode = lst.SelectedItem.SubItems(2)
rst!Desc = lst.SelectedItem.SubItems(3)
rst!price = lst.SelectedItem.SubItems(4)
rst!amount = lst.SelectedItem.SubItems(5)
rst!mm = Format(Date, "mm")
rst!dd = Format(Date, "dd")
rst!yyyy = Format(Date, "yyyy")
MsgBox "Sucessfully Return!", vbInformation, "Confirmation"
txttot = Val(txttot) + lst.SelectedItem.SubItems(5)
rst.Update
rst.Close
lst.SelectedItem = ""
lst.SelectedItem.SubItems(1) = ""
lst.SelectedItem.SubItems(2) = ""
lst.SelectedItem.SubItems(3) = ""
lst.SelectedItem.SubItems(4) = ""
lst.SelectedItem.SubItems(5) = ""
Call reload
cmdprocess.Enabled = True
End Sub
