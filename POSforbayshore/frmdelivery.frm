VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdelivery 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product Delivery Form"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picprint 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1065
      ScaleWidth      =   3105
      TabIndex        =   29
      Top             =   5280
      Visible         =   0   'False
      Width           =   3135
      Begin Project1.chameleonButton cmdprint 
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   4
         TX              =   "PRINT"
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
      Begin Project1.chameleonButton cancel 
         Height          =   495
         Left            =   1680
         TabIndex        =   31
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   4
         TX              =   "CANCEL"
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
   Begin VB.PictureBox picinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      ScaleHeight     =   5985
      ScaleWidth      =   7065
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.PictureBox picprod 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   1800
         ScaleHeight     =   2985
         ScaleWidth      =   5145
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   5175
         Begin MSComctlLib.ListView lst 
            Height          =   2295
            Left            =   120
            TabIndex        =   35
            ToolTipText     =   "Double click the selected item to edit or delete"
            Top             =   120
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   4048
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
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
         End
         Begin Project1.chameleonButton select 
            Height          =   375
            Left            =   2520
            TabIndex        =   36
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   2
            TX              =   "Select"
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
         Begin Project1.chameleonButton chameleonButton2 
            Height          =   375
            Left            =   3840
            TabIndex        =   37
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   2
            TX              =   "Cancel"
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
      Begin VB.TextBox txtsupp 
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
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1560
         Width           =   5175
      End
      Begin VB.TextBox txttot 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   27
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox txtamount 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtpcode 
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
         Left            =   120
         TabIndex        =   24
         Top             =   6120
         Width           =   210
      End
      Begin VB.TextBox txtprice 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   19
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtpending 
         Alignment       =   1  'Right Justify
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtrqty 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtqty 
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
         TabIndex        =   14
         Top             =   6000
         Width           =   210
      End
      Begin VB.TextBox txtprod 
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
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtpoid 
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
         TabIndex        =   6
         Top             =   600
         Width           =   3255
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
         TabIndex        =   2
         Top             =   120
         Width           =   2295
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
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   2175
      End
      Begin Project1.chameleonButton cmdenter 
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         FCOL            =   0
      End
      Begin Project1.chameleonButton cmdselect 
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Select"
         ENAB            =   0   'False
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
         Left            =   6120
         TabIndex        =   11
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Cancel"
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
      Begin Project1.chameleonButton cmddeliver 
         Height          =   375
         Left            =   5160
         TabIndex        =   21
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Deliver"
         ENAB            =   0   'False
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
      Begin Project1.chameleonButton cmddelete 
         Height          =   375
         Left            =   6120
         TabIndex        =   22
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Delete"
         ENAB            =   0   'False
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
      Begin MSComctlLib.ListView lstdeliver 
         Height          =   2295
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Double click the selected item to edit or delete"
         Top             =   3000
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Quatity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pending"
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
      Begin VB.Label Label10 
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
         Left            =   -120
         TabIndex        =   33
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
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
         TabIndex        =   28
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   3360
         TabIndex        =   26
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label6 
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
         Left            =   1200
         TabIndex        =   20
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   2160
         Width           =   975
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
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
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
         Left            =   -120
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Order Id"
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
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
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
         TabIndex        =   4
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
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
   Begin Project1.chameleonButton cmdnew 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   6360
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
      Top             =   6360
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
End
Attribute VB_Name = "frmdelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
If cmdprint.Enabled = False Then
    rst.Close
    Call clear2
    cmdnew.Enabled = True
    cmdnew.SetFocus
    picinfo.Enabled = False
    picprint.Visible = False
Else
    Call clear2
     cmdnew.Enabled = True
    cmdnew.SetFocus
    picinfo.Enabled = False
    picprint.Visible = False
End If
End Sub
Function clear2()
txttransid = ""
txtpoid = ""
txtprod = ""
lst.ListItems.clear
txtqty = ""
txtprice = ""
txtamount = ""
txtpending = ""
txttot = ""
txtsupp = ""
lstdeliver.ListItems.clear
End Function

Private Sub chameleonButton2_Click()
picprod.Visible = False
End Sub

Private Sub cmdcancel_Click()
picprod.Visible = False
txtprod = ""
End Sub

Private Sub cmddelete_Click()
Dim q As VbMsgBoxResult
q = MsgBox("Delete?", vbQuestion + vbYesNo, "System Message")
If q = vbYes Then
    rst.Open "Select * from tbldelivery", con, adOpenDynamic, adLockOptimistic
    While rst.EOF = False
    If txttransid = rst!transid And lstdeliver.SelectedItem.SubItems(1) = rst!pcode Then
        rst.Delete
        rst.Update
        MsgBox "Sucessfully Deleted!", vbCritical, "Confirmation"
        txttot = Val(txttot) - Val(lstdeliver.SelectedItem.SubItems(5))
    End If
    rst.MoveNext
    Wend
    rst.Close
    Call reload
End If
End Sub

Private Sub cmddeliver_Click()
If txtprod = "" Then
    MsgBox "Please select product!", vbCritical, "System Message"
    cmdselect_Click
ElseIf txtrqty = "" Then
    MsgBox "Please enter received quantity!", vbCritical, "System Message"
    txtrqty.SetFocus
ElseIf txtprice = "" Then
    MsgBox "Please enter original price!", vbCritical, "System Message"
    txtprice.SetFocus
Else
rst.Open "Select * from tblprice", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
    If txtpcode = rst!pcode Then
    If IsNull(rst!stocks) Then
        rst!stocks = txtrqty
        rst!org_price = txtprice
    
    Else
        rst!stocks = Val(rst!stocks) + Val(txtrqty)
        rst!org_price = txtprice
    End If
    End If
    rst.MoveNext
Wend
rst.Close

rst.Open "Select * from tbldelivery", con, adOpenDynamic, adLockPessimistic
rst.AddNew
rst!transid = txttransid
rst!pcode = txtpcode
rst!Desc = txtprod
rst!qty = txtrqty
rst!price = txtprice
rst!amount = txtamount
rst!pending = txtpending
rst!mm = Format(Date, "mm")
rst!dd = Format(Date, "dd")
rst!yyyy = Format(Date, "yyyy")
rst!supp = txtsupp
rst.Update
rst.Close
txttot = Val(txttot) + Val(txtamount)
MsgBox "Delivered!", vbInformation, "Confirmation"
Call reload
cmddeliver.Enabled = False
txtprod = ""
txtrqty = ""
txtprice = ""
txtpending = ""
txtamount = ""
cmdprocess.Enabled = True

End If
End Sub
Function reload()
rst.Open "Select * from tbldelivery", con, adOpenDynamic, adLockOptimistic
lstdeliver.ListItems.clear
While rst.EOF = False
If txttransid = rst!transid Then
lstdeliver.ListItems.Add , , rst!qty
lstdeliver.ListItems(lstdeliver.ListItems.Count).ListSubItems.Add , , rst!pcode
lstdeliver.ListItems(lstdeliver.ListItems.Count).ListSubItems.Add , , rst!Desc
lstdeliver.ListItems(lstdeliver.ListItems.Count).ListSubItems.Add , , rst!pending
lstdeliver.ListItems(lstdeliver.ListItems.Count).ListSubItems.Add , , rst!price
lstdeliver.ListItems(lstdeliver.ListItems.Count).ListSubItems.Add , , rst!amount
End If
rst.MoveNext
Wend
rst.Close
End Function

Private Sub cmdenter_Click()
cmdselect.Enabled = True
cmdselect_Click
End Sub

Private Sub cmdnew_Click()
picinfo.Enabled = True
rst.Open "Select * from tbldelivery", con, adOpenDynamic, adLockOptimistic
rst.MoveLast
txttransid = "DEL" + Format(Val(Right(rst!transid, 5)) + 1, "0000#")
rst.Close
cmdnew.Enabled = False
txtpoid.SetFocus
End Sub

Private Sub cmdprint_Click()
rst.Open "Select * from tbldelivery where transid='" & txttransid & "'", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If txttransid = rst!transid Then
Set dtpdelivery.DataSource = rst
dtpdelivery.Sections("Section5").Controls.Item("lbldate").Caption = Date
dtpdelivery.Sections("Section5").Controls.Item("lbltransid").Caption = txttransid
dtpdelivery.Sections("Section5").Controls.Item("lbltot").Caption = txttot
dtpdelivery.Sections("Section5").Controls.Item("lbluser").Caption = mdimain.lblname
dtpdelivery.Show
cmdprint.Enabled = False
End If
rst.MoveNext
Wend
End Sub

Private Sub cmdprocess_Click()
MsgBox "Process Complete!", vbInformation, "System Message"
picinfo.Enabled = False
picprint.Visible = True
cmdprocess.Enabled = False
End Sub

Private Sub cmdselect_Click()
Dim x
rst.Open "Select * from tblorder", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
If txtpoid = rst!transid Then
picprod.Visible = True
lst.ListItems.Add , , rst!qty
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!unit
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!pcode
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
txtsupp = rst!supp
x = 1
End If
rst.MoveNext
Wend
rst.Close
If x <> 1 Then
    MsgBox "Please check product order id!", vbCritical, "System Message"
    txtpoid = ""
    txtpoid.SetFocus
End If

End Sub

Private Sub Form_Load()
txtdate = Format(Date, "mm/dd/yyyy")
End Sub



Private Sub lst_DblClick()
select_Click
End Sub



Private Sub lstdeliver_DblClick()
cmddelete.Enabled = True
cmddelete.SetFocus
End Sub

Private Sub select_Click()
txtrqty = ""
txtprice = ""
txtamount = ""
txtpending = ""
txtprod = lst.SelectedItem.SubItems(3)
txtpcode = lst.SelectedItem.SubItems(2)
txtqty = lst.SelectedItem
cmddeliver.Enabled = True
picprod.Visible = False
txtrqty.SetFocus
End Sub



Private Sub txtprice_Change()
txtamount = Val(txtrqty) * Val(txtprice)
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Is < 32
    Case 48 To 57
    Case 46
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtrqty_Change()
txtpending = ""
txtpending = Val(txtqty) - Val(txtrqty)
End Sub

Private Sub txtrqty_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Is < 32
    Case 48 To 57
    Case 46
    Case Else
        KeyAscii = 0
End Select
End Sub
