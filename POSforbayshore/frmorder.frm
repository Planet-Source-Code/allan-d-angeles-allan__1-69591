VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmorder 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product Order Form"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picprint 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1065
      ScaleWidth      =   3105
      TabIndex        =   28
      Top             =   4800
      Visible         =   0   'False
      Width           =   3135
      Begin Project1.chameleonButton cmdprint 
         Height          =   495
         Left            =   120
         TabIndex        =   29
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
         TabIndex        =   30
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
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5625
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
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   5175
         Begin MSComctlLib.ListView lst 
            Height          =   2295
            Left            =   120
            TabIndex        =   23
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
            Left            =   2520
            TabIndex        =   24
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
            TabIndex        =   25
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
         TabIndex        =   20
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtqty 
         Alignment       =   2  'Center
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
         TabIndex        =   19
         Top             =   2040
         Width           =   3255
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
         Left            =   1800
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtpcode 
         Height          =   375
         Left            =   9120
         TabIndex        =   10
         Top             =   1080
         Width           =   150
      End
      Begin Project1.chameleonButton cmdselect 
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Select"
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
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin VB.ComboBox cbosupp 
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
         TabIndex        =   4
         Top             =   600
         Width           =   1935
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
         Width           =   2295
      End
      Begin Project1.chameleonButton cmdcancel 
         Height          =   375
         Left            =   6120
         TabIndex        =   9
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
      Begin Project1.chameleonButton cmdorder 
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Order"
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
         Left            =   6120
         TabIndex        =   15
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Delete"
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
      Begin MSComctlLib.ListView lstorder 
         Height          =   2655
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "To delete item please double click the selected item"
         Top             =   2760
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4683
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
            Text            =   "Trans_Id"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pcode"
            Object.Width           =   2540
         EndProperty
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
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
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
         TabIndex        =   18
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "List Of Ordered Products"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label lblunit 
         BackStyle       =   0  'Transparent
         Caption         =   "unit desc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   1560
         Width           =   4935
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblsupp 
         BackStyle       =   0  'Transparent
         Caption         =   "supplier name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin Project1.chameleonButton cmdnew 
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   5880
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
      TabIndex        =   27
      Top             =   5880
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
Attribute VB_Name = "frmorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function clear()
txtprod = ""
cbounit = ""
lblunit = ""
txtqty = ""
End Function

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

Private Sub cbosupp_Click()
rst.Open "Select * from tblsupp", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
    If cbosupp = rst!suppid Then
        lblsupp.Caption = rst!Name
    End If
    rst.MoveNext
Wend
rst.Close
End Sub

Function clear2()
txttransid = ""
cbosupp = ""
lblsupp = ""
txtprod = ""
cbounit = ""
txtqty = ""
lblunit = ""
lstorder.ListItems.clear
End Function



Private Sub cbounit_Click()
rst.Open "Select * from tblunit", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
    If cbounit = rst!unitcode Then
        lblunit.Caption = rst!Desc
    End If
    rst.MoveNext
Wend
rst.Close
End Sub

Private Sub chameleonButton2_Click()
picprod.Visible = False
End Sub

Private Sub cmdcancel_Click()
picprod.Visible = False
txtprod = ""
End Sub

Private Sub cmddelete_Click()
Dim q As VbMsgBoxResult
q = MsgBox("Delete Item?", vbYesNo + vbQuestion, "System Message")
If q = vbYes Then
    rst.Open "Select * from tblorder", con, adOpenDynamic, adLockOptimistic
    While rst.EOF = False
    If txttransid = rst!transid And rst!pcode = lstorder.SelectedItem.SubItems(3) Then
    rst.Delete
    rst.Update
    End If
    rst.MoveNext
    Wend
    rst.Close
    Call reloadorder
End If
End Sub

Private Sub cmdnew_Click()
picinfo.Enabled = True
rst.Open "select * from tblorder", con, adOpenDynamic, adLockOptimistic
rst.MoveLast
txttransid = "ORDER" + Format(Val(Right(rst!transid, 5)) + 1, "0000#")
rst.Close
cbosupp.SetFocus
 cmdnew.Enabled = False
End Sub

Private Sub cmdorder_Click()
If cbosupp = "" Then
    MsgBox "Please select supplier!", vbCritical, "System Message"
    cbosupp.SetFocus
ElseIf txtprod = "" Then
    MsgBox "Please select product!", vbCritical, "System Message"
    cmdselect_Click
ElseIf cbounit = "" Then
    MsgBox "Please select unit!", vbCritical, "System Message"
    cbounit.SetFocus
ElseIf txtqty = "" Then
    MsgBox "Please check product quatnity!", vbCritical, "System Message"
    txtqty.SetFocus
Else
    Dim q As VbMsgBoxResult
    q = MsgBox("Order " + txtprod + "?", vbQuestion + vbYesNo, "System Message")
    If q = vbYes Then
        rst.Open "Select * from tblorder", con, adOpenDynamic, adLockPessimistic
        rst.AddNew
        
        rst!transid = txttransid
        rst!supp = lblsupp
        rst!pcode = txtpcode
        rst!Desc = txtprod
        rst!unit = lblunit
        rst!qty = txtqty
        rst!mm = Format(Date, "mm")
        rst!dd = Format(Date, "dd")
        rst!yyyy = Format(Date, "yyyy")
        rst.Update
        rst.Close
        MsgBox "Sucessfully added in order list!", vbInformation, "Confirmation"
        cmdprocess.Enabled = True
        Call reloadorder
        Call clear
    End If
End If
End Sub

Private Sub cmdprocess_Click()
MsgBox "Process Complete!", vbInformation, "System Message"
picinfo.Enabled = False
picprint.Visible = True
cmdprocess.Enabled = False
End Sub

Private Sub cmdselect_Click()
picprod.Visible = True
Call reload
End Sub

Private Sub Form_Load()
txtdate = Format(Date, "mm/dd/yyyy")
rst.Open "Select * from tblsupp", con, adOpenDynamic, adLockOptimistic
cbosupp.clear
While rst.EOF = False
cbosupp.AddItem rst!suppid
rst.MoveNext
Wend
rst.Close

rst.Open "Select * from tblunit", con, adOpenDynamic, adLockOptimistic
cbounit.clear
While rst.EOF = False
cbounit.AddItem rst!unitcode
rst.MoveNext
Wend
rst.Close

End Sub



Private Sub lst_DblClick()
select_Click
End Sub



Private Sub lstorder_DblClick()
cmddelete.Enabled = True
cmddelete.SetFocus
End Sub

Private Sub cmdprint_Click()
rst.Open "Select * from tblorder where transid='" & txttransid & "'", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If txttransid = rst!transid Then
Set dtporder.DataSource = rst
dtporder.Sections("Section5").Controls.Item("lbldate").Caption = Date
dtporder.Sections("Section2").Controls.Item("lblsupp").Caption = lblsupp
dtporder.Sections("Section5").Controls.Item("lbltransid").Caption = txttransid
dtporder.Sections("Section5").Controls.Item("lbluser").Caption = mdimain.lblname
dtporder.Show
cmdprint.Enabled = False
End If
rst.MoveNext
Wend
End Sub

Private Sub select_Click()
txtpcode = lst.SelectedItem
txtprod = lst.SelectedItem.SubItems(1)
picprod.Visible = False
End Sub

Function reload()
rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
    If cbosupp = rst!supp Then
    lst.ListItems.Add , , rst!pcode
    lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
    End If
    rst.MoveNext
Wend
rst.Close
End Function

Function reloadorder()
rst.Open "Select * from tblorder", con, adOpenDynamic, adLockOptimistic
lstorder.ListItems.clear
While rst.EOF = False
If txttransid = rst!transid Then
lstorder.ListItems.Add , , rst!transid
lstorder.ListItems(lstorder.ListItems.Count).ListSubItems.Add , , rst!qty
lstorder.ListItems(lstorder.ListItems.Count).ListSubItems.Add , , rst!unit
lstorder.ListItems(lstorder.ListItems.Count).ListSubItems.Add , , rst!pcode
End If
rst.MoveNext
Wend
rst.Close
End Function

