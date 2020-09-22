VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return to Supplier Form"
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
   Begin VB.PictureBox picprint 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1065
      ScaleWidth      =   4785
      TabIndex        =   25
      Top             =   9120
      Visible         =   0   'False
      Width           =   4815
      Begin Project1.chameleonButton cmdprint 
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
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
      Begin Project1.chameleonButton chameleonButton1 
         Height          =   495
         Left            =   2640
         TabIndex        =   27
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
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
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   120
      ScaleHeight     =   9705
      ScaleWidth      =   14745
      TabIndex        =   0
      Top             =   240
      Width           =   14775
      Begin VB.PictureBox piclist 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   1800
         ScaleHeight     =   3465
         ScaleWidth      =   5625
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   5655
         Begin Project1.chameleonButton cancel 
            Height          =   375
            Left            =   4560
            TabIndex        =   31
            Top             =   3000
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            FCOL            =   0
         End
         Begin Project1.chameleonButton select 
            Height          =   375
            Left            =   3360
            TabIndex        =   32
            Top             =   3000
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "SELECT"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            FCOL            =   0
         End
         Begin MSComctlLib.ListView lst 
            Height          =   2775
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   5415
            _ExtentX        =   9551
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Pcode"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Desc"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Price"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Amount"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.TextBox txtsupp 
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
         TabIndex        =   28
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox txtrqty 
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
         Left            =   14760
         TabIndex        =   24
         Top             =   1680
         Width           =   210
      End
      Begin VB.TextBox txtprod 
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
         Left            =   15240
         TabIndex        =   23
         Top             =   1680
         Width           =   210
      End
      Begin VB.TextBox txtpcode 
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
         Left            =   15000
         TabIndex        =   21
         Top             =   1680
         Width           =   210
      End
      Begin VB.TextBox txtamount 
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
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtprice 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtqty 
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
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
         TabIndex        =   4
         Top             =   120
         Width           =   2895
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
         TabIndex        =   3
         Top             =   120
         Width           =   2415
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
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
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
         TabIndex        =   1
         Top             =   7800
         Width           =   9495
      End
      Begin Project1.chameleonButton cmdenter 
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&ENTER"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         FCOL            =   0
      End
      Begin Project1.chameleonButton cmdcancel 
         Height          =   375
         Left            =   6480
         TabIndex        =   6
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&CANCEL"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         FCOL            =   0
      End
      Begin MSComctlLib.ListView lstlist 
         Height          =   5175
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   9128
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
      Begin Project1.chameleonButton cmdreturn 
         Height          =   375
         Left            =   5400
         TabIndex        =   22
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&RETURN"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         FCOL            =   0
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Left            =   0
         TabIndex        =   29
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   3960
         TabIndex        =   20
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Left            =   2160
         TabIndex        =   18
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   1335
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
         TabIndex        =   11
         Top             =   240
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
         TabIndex        =   10
         Top             =   240
         Width           =   495
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
         TabIndex        =   9
         Top             =   720
         Width           =   1575
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
         TabIndex        =   8
         Top             =   7560
         Width           =   1575
      End
   End
   Begin Project1.chameleonButton cmdnew 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   10200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&New Trans."
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
   Begin Project1.chameleonButton cmdprocess 
      Height          =   495
      Left            =   1800
      TabIndex        =   13
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
      Left            =   3480
      TabIndex        =   14
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
Attribute VB_Name = "frmrts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
piclist.Visible = False
End Sub

Private Sub chameleonButton1_Click()
If cmdprint.Enabled = False Then
    rst.Close
    Call clear
    picprint.Visible = False
    cmdnew.Enabled = True
    cmdnew.SetFocus
    cmdout.Enabled = True
Else
    Call clear
    picprint.Visible = False
    cmdnew.SetFocus
    cmdnew.Enabled = True
cmdout.Enabled = True
End If

End Sub
Function clear()
txttransid = ""
transid = ""
lst.ListItems.clear
txtsupp = ""
txtqty = ""
txtprice = ""
txtamount = ""
txttot = ""
lstlist.ListItems.clear
End Function

Private Sub cmdenter_Click()
Dim X
rst.Open "Select * from tbldelivery", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
If transid = rst!transid Then
piclist.Visible = True
lst.ListItems.Add , , rst!qty
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!pcode
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!price
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!amount
txtsupp = rst!supp
X = 1
End If
rst.MoveNext
Wend
rst.Close
If X <> 1 Then
MsgBox "Invalid transaction id!", vbCritical, "System Message"
transid = ""
transid.SetFocus
End If
End Sub

Private Sub cmdnew_Click()
picinfo.Enabled = True
rst.Open "Select * from tblrts", con, adOpenDynamic, adLockOptimistic
rst.MoveLast
txttransid = "RTS" + Format(Val(Right(rst!transid, 5)) + 1, "0000#")
rst.Close
transid.SetFocus
cmdnew.Enabled = False
End Sub

Private Sub cmdout_Click()
Unload Me
mdimain.Show
End Sub

Private Sub cmdprint_Click()
rst.Open "Select * from tblrts where transid = '" & txttransid & "'", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If txttransid = rst!transid Then
Set dtprts.DataSource = rst
dtprts.Sections("Section2").Controls.Item("lblsupp").Caption = txtsupp
dtprts.Sections("Section5").Controls.Item("lbldate").Caption = Date
dtprts.Sections("Section5").Controls.Item("lbluser").Caption = mdimain.lblname
dtprts.Sections("Section5").Controls.Item("lbltransid").Caption = txttransid
dtprts.Sections("Section5").Controls.Item("lbltot").Caption = txttot
dtprts.Show
End If
rst.MoveNext
Wend
cmdprint.Enabled = False
End Sub

Private Sub cmdprocess_Click()
picinfo.Enabled = False
MsgBox "Process Complete!", vbInformation, "Confirmation"
picprint.Visible = True
cmdprocess.Enabled = False
cmdnew.Enabled = False
cmdout.Enabled = False
cmdprint.Enabled = True
End Sub

Private Sub cmdreturn_Click()
If txtqty = "" Then
    MsgBox "Please Enter Quantity", vbCritical, "System Message"
    txtqty.SetFocus
Else
rst.Open "Select * from tblrts", con, adOpenDynamic, adLockOptimistic
rst.AddNew
rst!transid = txttransid
rst!qty = txtqty
rst!pcode = txtpcode
rst!Desc = txtprod
rst!price = txtprice
rst!amount = txtamount
rst!mm = Format(Date, "mm")
rst!dd = Format(Date, "dd")
rst!yyyy = Format(Date, "yyyy")
rst!supp = txtsupp
rst.Update
rst.Close
txttot = Val(txttot) + Val(txtamount)
Call reload
txtqty = ""
txtprice = ""
txtamount = ""
cmdprocess.Enabled = True
End If
End Sub

Function reload()
rst.Open "Select * from tblrts", con, adOpenDynamic, adLockOptimistic
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

Private Sub Form_Load()
txtdate = Format(Date, "mm/dd/yyyy")
End Sub

Private Sub select_Click()
txtqty = ""
txtprice = ""
txtamount = ""
txtprice = lst.SelectedItem.SubItems(3)
txtpcode = lst.SelectedItem.SubItems(1)
txtprod = lst.SelectedItem.SubItems(2)
txtrqty = lst.SelectedItem

txtqty.SetFocus
piclist.Visible = False
End Sub



Private Sub txtqty_Change()
txtamount = Val(txtqty) * Val(txtprice)
End Sub

Private Sub txtqty_LostFocus()
If Val(txtqty) > Val(txtrqty) Then
MsgBox "Error: Invalid Input!", vbCritical, "System Message"
txtqty = ""
txtqty.SetFocus
End If
End Sub
