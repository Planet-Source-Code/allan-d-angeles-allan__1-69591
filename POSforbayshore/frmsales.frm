VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsales 
   BorderStyle     =   0  'None
   Caption         =   "Sales Invoice"
   ClientHeight    =   11160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox piccash 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   3840
      ScaleHeight     =   5505
      ScaleWidth      =   7545
      TabIndex        =   29
      Top             =   2520
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox txtcash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   4920
         TabIndex        =   31
         Top             =   1920
         Width           =   2415
      End
      Begin Project1.chameleonButton cmdprint 
         Height          =   615
         Left            =   240
         TabIndex        =   30
         Top             =   4560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "PRINT &RECEIPT"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
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
      Begin Project1.chameleonButton cancel 
         Height          =   615
         Left            =   4560
         TabIndex        =   32
         Top             =   4560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "CANCEL"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
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
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Total              : Php"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cash Tender : Php"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   36
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Change          : Php"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   35
         Top             =   3120
         Width           =   4695
      End
      Begin VB.Label lbltot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   4920
         TabIndex        =   34
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblchange 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   4920
         TabIndex        =   33
         Top             =   3120
         Width           =   2415
      End
   End
   Begin VB.PictureBox picinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   10215
      Left            =   240
      ScaleHeight     =   10215
      ScaleWidth      =   15015
      TabIndex        =   0
      Top             =   120
      Width           =   15015
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
         Height          =   1440
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   8640
         Width           =   9975
      End
      Begin VB.PictureBox picprod 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   1800
         ScaleHeight     =   3705
         ScaleWidth      =   5025
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   5055
         Begin MSComctlLib.ListView lst 
            Height          =   3015
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Double click the selected item to edit or delete"
            Top             =   120
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   5318
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Pcode"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   3618
            EndProperty
         End
         Begin Project1.chameleonButton select 
            Height          =   375
            Left            =   2400
            TabIndex        =   25
            Top             =   3240
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
            Left            =   3720
            TabIndex        =   26
            Top             =   3240
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
      Begin VB.ComboBox cbounit 
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
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtqty 
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
         TabIndex        =   14
         Top             =   1440
         Width           =   975
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
         Left            =   2880
         TabIndex        =   13
         Top             =   1440
         Width           =   855
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
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1440
         Width           =   855
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
         Left            =   3600
         TabIndex        =   9
         Top             =   10320
         Width           =   210
      End
      Begin Project1.chameleonButton cmdselect 
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&SELECT"
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
         TabIndex        =   6
         Top             =   720
         Width           =   3015
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
         Left            =   12480
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
      Begin Project1.chameleonButton cmdcancel 
         Height          =   375
         Left            =   5880
         TabIndex        =   8
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
      Begin Project1.chameleonButton cmdpurchase 
         Height          =   375
         Left            =   4800
         TabIndex        =   20
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&PURCHASE"
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
      Begin Project1.chameleonButton cmddelete 
         Height          =   375
         Left            =   5880
         TabIndex        =   21
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&DELETE"
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
         Height          =   6735
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Double click the selected item to edit or delete"
         Top             =   1800
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   11880
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Qty"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Unit"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Pcode"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Desc"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Price"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Happy to serve you!"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   6960
         TabIndex        =   40
         Top             =   1200
         Width           =   7935
      End
      Begin VB.Label lblname 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   8520
         TabIndex        =   39
         Top             =   720
         Width           =   6375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hi Im"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   6960
         TabIndex        =   38
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
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
         TabIndex        =   18
         Top             =   1200
         Width           =   615
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
         TabIndex        =   17
         Top             =   1200
         Width           =   855
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
         Left            =   2880
         TabIndex        =   16
         Top             =   1200
         Width           =   615
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
         Left            =   3720
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Products"
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
         TabIndex        =   5
         Top             =   720
         Width           =   1575
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
         Left            =   11880
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
   Begin Project1.chameleonButton cmdnew 
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   10440
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
      Left            =   1920
      TabIndex        =   11
      Top             =   10440
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
      Left            =   3600
      TabIndex        =   28
      Top             =   10440
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
Attribute VB_Name = "frmsales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click()
Call reload
picprod.Visible = True
End Sub

Private Sub cancel_Click()
If cmdprint.Enabled = False Then
    rst.Close
    piccash.Visible = False
    cmdnew.Enabled = True
    cmdnew.SetFocus
    cmdprocess.Enabled = False
    cmdout.Enabled = True
Else
    piccash.Visible = False
    cmdprocess.Enabled = True
    picinfo.Enabled = True
    cmdnew.Enabled = False
End If
End Sub
Function clear2()
txtprod = ""
txtqty = ""
txtprice = ""
cbounit = ""
txtamount = ""
txttot = ""
lbltot = ""
lblchange = ""
txtcash = ""
lstlist.ListItems.clear
End Function

Private Sub chameleonButton2_Click()
picprod.Visible = False
End Sub

Private Sub cmdcancel_Click()
picprod.Visible = False
txtprod = ""
End Sub

Private Sub cmddelete_Click()

rst.Open "Select * from tblsales", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If txttransid = rst!transid And lstlist.SelectedItem.SubItems(2) = rst!pcode Then
    rst.Delete
    rst.Update
    txttot = Val(txttot) - Val(lstlist.SelectedItem.SubItems(5))
    cmddelete.Enabled = False
End If
rst.MoveNext
Wend
rst.Close
Call reloadlist

End Sub
Function clear()
lstlist.ListItems.clear
lst.ListItems.clear
txttot = ""
txtprod = ""
txtqty = ""
txtprice = ""
txtamount = ""
End Function

Private Sub cmdnew_Click()
picinfo.Enabled = True
Call clear
rst1.Open "Select * from tblsales", con, adOpenDynamic, adLockOptimistic
rst1.MoveLast
txttransid = "SALE" + Format(Val(Right(rst1!transid, 5)) + 1, "00000#")
rst1.Close
End Sub

Function reload()
rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
lst.ListItems.Add , , rst!pcode
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
rst.MoveNext
Wend
rst.Close
End Function

Private Sub cmdout_Click()
Unload Me
mdimain.Show
End Sub

Private Sub cmdprint_Click()
If Val(txtcash) >= Val(lbltot) Then
cmdprint.Enabled = True

rst.Open "Select * from tblsales where transid='" & txttransid & "'", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
Set dtpreceipt.DataSource = rst
dtpreceipt.Sections("Section5").Controls.Item("lbldate").Caption = Date
dtpreceipt.Sections("Section5").Controls.Item("lbltransid").Caption = txttransid
dtpreceipt.Sections("Section5").Controls.Item("lbltot").Caption = lbltot
dtpreceipt.Sections("Section5").Controls.Item("lblcash").Caption = txtcash
dtpreceipt.Sections("Section5").Controls.Item("lblchange").Caption = lblchange
dtpreceipt.Sections("Section5").Controls.Item("lblcashier").Caption = lblname
dtpreceipt.Show
rst.MoveNext
Wend
cmdprint.Enabled = False
Else
MsgBox "Insuficient Amount!", vbCritical, "System Message"
txtcash.SetFocus
End If
End Sub

Private Sub cmdprocess_Click()
picinfo.Enabled = False
piccash.Visible = True
txtcash = ""
lblchange = ""
lbltot = txttot
txtcash.SetFocus
cmdprocess.Enabled = False
cmdnew.Enabled = False
cmdout.Enabled = False
cmdprint.Enabled = True

End Sub

Private Sub cmdpurchase_Click()
If txtprod = "" Then
    MsgBox "Please select product!", vbCritical, "System Message"
    cmdselect_Click
ElseIf txtqty = "" Then
    MsgBox "Please check product quantity!", vbCritical, "System Message"
    txtqty.SetFocus
ElseIf cbounit = "" Then
    MsgBox "Please check product unit!", vbCritical, "System Message"
    cbounit.SetFocus
Else
    rst.Open "Select * from tblsales", con, adOpenDynamic, adLockPessimistic
    rst.AddNew
    rst!transid = txttransid
    rst!qty = txtqty
    rst!unit = cbounit
    rst!pcode = txtpcode
    rst!Desc = txtprod
    rst!price = txtprice
    rst!amount = Format(txtamount, "0.00")
    rst!mm = Format(Date, "mm")
    rst!dd = Format(Date, "dd")
    rst!yyyy = Format(Date, "yyyy")
    MsgBox "Sucess!", vbInformation, "Confirmation"
    rst.Update
    rst.Close
    
    rst.Open "Select * from tblprice", con, adOpenDynamic, adLockOptimistic
    While rst.EOF = False
    If txtpcode = rst!pcode Then
        rst!stocks = Val(rst!stocks) - Val(txtqty)
    End If
    rst.MoveNext
    Wend
    rst.Close
    
    Call reloadlist
    txttot = Val(txttot) + Val(txtamount)
    txtprod = ""
    txtqty = ""
    cbounit = ""
    txtprice = ""
    txtamount = ""
    cmdprocess.Enabled = True
    cmdprocess.SetFocus
    
    cmdnew.Enabled = False
    cmdout.Enabled = False
    
End If
End Sub

Private Sub cmdselect_Click()
picprod.Visible = True
txtprod = ""
    txtqty = ""
    cbounit = ""
    txtprice = ""
    txtamount = ""
Call reload
End Sub

Private Sub Form_Load()
txtdate = Format(Date, "mm/dd/yyyy")
rst.Open "Select * from tblunit", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
cbounit.AddItem rst!Desc
rst.MoveNext
Wend
rst.Close
cmddelete.Enabled = False
End Sub
Function reloadlist()
rst.Open "Select * from tblsales", con, adOpenDynamic, adLockOptimistic
lstlist.ListItems.clear
While rst.EOF = False
If txttransid = rst!transid Then
lstlist.ListItems.Add , , rst!qty
lstlist.ListItems(lstlist.ListItems.Count).ListSubItems.Add , , rst!unit
lstlist.ListItems(lstlist.ListItems.Count).ListSubItems.Add , , rst!pcode
lstlist.ListItems(lstlist.ListItems.Count).ListSubItems.Add , , rst!Desc
lstlist.ListItems(lstlist.ListItems.Count).ListSubItems.Add , , rst!price
lstlist.ListItems(lstlist.ListItems.Count).ListSubItems.Add , , rst!amount
End If
rst.MoveNext
Wend
rst.Close

End Function



Private Sub Label13_Click()
piccash.Visible = False
End Sub

Private Sub lblchange_Click()
lblchange.Caption = Format(lblchange, "0.00")
End Sub

Private Sub lbltot_Click()
lbltot.Caption = Format(lbltot, "0.00")
End Sub

Private Sub lst_DblClick()
select_Click
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then select_Click
End Sub



Private Sub lstlist_DblClick()
If lstlist.SelectedItem Is Nothing Then
    cmddelete.Enabled = False
    cmdout.Enabled = True
    cmdprocess.Enabled = False
Else
cmddelete.Enabled = True
End If
End Sub

Private Sub lstlist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstlist.SelectedItem Is Nothing Then
    cmddelete.Enabled = False
    cmdout.Enabled = True
    cmdprocess.Enabled = False
Else
cmddelete.Enabled = True
End If
End Sub

Private Sub select_Click()

txtpcode = lst.SelectedItem
txtprod = lst.SelectedItem.SubItems(1)
picprod.Visible = False
txtqty.SetFocus
rst.Open "Select * from tblprice", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If txtpcode = rst!pcode Then
txtprice = rst!selling
End If
rst.MoveNext
Wend
rst.Close

If txtprice = "" Then
    MsgBox "Please add price in price maintenance form!", vbInformation, "System Message"
    frmsales.Hide
    frmprice.Show
End If
End Sub



Private Sub txtcash_Change()
lblchange = Val(txtcash) - Val(lbltot)
End Sub

Private Sub txtcash_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Is < 32
    Case 48 To 57
    Case 46
    Case Else
        KeyAscii = 0
End Select
If KeyAscii = 13 Then cmdprint_Click
End Sub





Private Sub txtqty_Change()
txtamount = Val(txtqty) * Val(txtprice)
End Sub


Private Sub txtqty_LostFocus()
rst.Open "Select * from tblprice", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If txtpcode = rst!pcode Then
    
    If Val(txtqty) > Val(rst!stocks) Then
        MsgBox "Error: Out of Stocks!", vbCritical, "Insuficient Stocks"
        txtqty = ""
        txtqty.SetFocus
    End If
    If Val(rst!stocks) <= Val(rst!critical) Then
        MsgBox "The product are in critical level!", vbInformation, "Information Message"
        cbounit.SetFocus
    End If
End If
rst.MoveNext
Wend
rst.Close
End Sub

Private Sub txttot_Change()
txttot = Format(txttot, ".00")
End Sub
