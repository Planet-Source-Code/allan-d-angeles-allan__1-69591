VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsupp 
   BackColor       =   &H80000010&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Supplier Maintenance"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2535
      ScaleWidth      =   7695
      TabIndex        =   2
      Top             =   240
      Width           =   7695
      Begin VB.TextBox txtno 
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
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   2040
         Width           =   5655
      End
      Begin VB.TextBox txtperson 
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
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1560
         Width           =   5655
      End
      Begin VB.TextBox txtaddress 
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
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox txtsuppid 
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
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   5655
      End
      Begin VB.TextBox txtname 
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
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   5655
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000009&
         Height          =   2535
         Left            =   0
         Top             =   0
         Width           =   7575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Height          =   135
         Left            =   360
         TabIndex        =   20
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Height          =   135
         Left            =   0
         TabIndex        =   19
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Height          =   135
         Left            =   720
         TabIndex        =   18
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Height          =   135
         Left            =   960
         TabIndex        =   17
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No."
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
         Left            =   480
         TabIndex        =   15
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
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
         TabIndex        =   14
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   720
         TabIndex        =   13
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lst 
      Height          =   3615
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Double click the selected item to edit or delete"
      Top             =   3000
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6376
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Supp_ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Contact_Person"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contact_No."
         Object.Width           =   2540
      EndProperty
   End
   Begin Project1.chameleonButton cmdadd 
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Add &New"
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
   Begin Project1.chameleonButton cmdsave 
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Save"
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
   Begin Project1.chameleonButton cmdedit 
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Edit"
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
   Begin Project1.chameleonButton cmddelete 
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Delete"
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
   Begin Project1.chameleonButton cmdupdate 
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   6120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Update"
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   3495
      Left            =   8040
      Top             =   3120
      Width           =   1815
   End
End
Attribute VB_Name = "frmsupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
picinfo.Enabled = True
Call clear
rst.Open "select * from tblsupp", con, adOpenDynamic, adLockOptimistic
rst.MoveLast
txtsuppid = "SUPP" + Format(Val(Right(rst!suppid, 3)) + 1, "00#")
rst.Close
txtname.SetFocus
cmdsave.Enabled = True
End Sub

Private Sub cmddelete_Click()
Dim q As VbMsgBoxResult
q = MsgBox("Delete record " + txtsuppid + " ? ", vbQuestion + vbYesNo, "System Message")
If q = vbYes Then
    rst.Open "Select * from tblsupp", con, adOpenDynamic, adLockOptimistic
    While rst.EOF = False
        If txtsuppid = rst!suppid Then
            rst.Delete
            rst.Update
            MsgBox "Deleted!", vbInformation, "Confirmation"
            Call clear
        End If
        rst.MoveNext
    Wend
    rst.Close
    Call reload
End If
cmddelete.Enabled = False
cmdedit.Enabled = False
End Sub

Private Sub cmdedit_Click()
picinfo.Enabled = True
cmdupdate.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
txtname.SetFocus
End Sub

Private Sub cmdsave_Click()
If txtsuppid = "" Then
    MsgBox "Please check supplier code!", vbCritical, "System Message"
    cmdadd.SetFocus
ElseIf txtname = "" Then
    MsgBox "Please check supplier name!", vbCritical, "System Message"
    txtname.SetFocus
ElseIf txtaddress = "" Then
    MsgBox "Please check supplier address!", vbCritical, "System Message"
    txtaddress.SetFocus
ElseIf txtperson = "" Then
    MsgBox "Please check supplier contact person!", vbCritical, "System Message"
    txtperson.SetFocus
ElseIf txtno = "" Then
    MsgBox "Please check supplier contact number!", vbCritical, "System Message"
    txtno.SetFocus
Else
    Dim q As VbMsgBoxResult
    q = MsgBox("Save record " + txtsuppid + " ? ", vbQuestion + vbYesNo, "System Message")
    If q = vbYes Then
    rst.Open " Select * from tblsupp", con, adOpenDynamic, adLockPessimistic
    rst.AddNew
    
    rst!suppid = txtsuppid
    rst!Name = txtname
    rst!address = txtaddress
    rst!person = txtperson
    rst!contact = txtno
    
    rst.Update
    rst.Close
    MsgBox "Saved!", vbInformation, "Confirmation"
    Call reload
    Call clear
    cmdsave.Enabled = False
    picinfo.Enabled = False
    End If
End If
End Sub

Private Sub cmdupdate_Click()
If txtsuppid = "" Then
    MsgBox "Please check supplier code!", vbCritical, "System Message"
    cmdadd.SetFocus
ElseIf txtname = "" Then
    MsgBox "Please check supplier name!", vbCritical, "System Message"
    txtname.SetFocus
ElseIf txtaddress = "" Then
    MsgBox "Please check supplier address!", vbCritical, "System Message"
    txtaddress.SetFocus
ElseIf txtperson = "" Then
    MsgBox "Please check supplier contact person!", vbCritical, "System Message"
    txtperson.SetFocus
ElseIf txtno = "" Then
    MsgBox "Please check supplier contact number!", vbCritical, "System Message"
    txtno.SetFocus
Else
    Dim q As VbMsgBoxResult
    q = MsgBox("Update record " + txtsuppid + " ? ", vbQuestion + vbYesNo, "System Message")
    If q = vbYes Then
    rst.Open " Select * from tblsupp", con, adOpenDynamic, adLockOptimistic
    While rst.EOF = False
    If rst!suppid = txtsuppid Then
    rst!Name = txtname
    rst!address = txtaddress
    rst!person = txtperson
    rst!contact = txtno
    
    rst.Update
    MsgBox "Saved!", vbInformation, "Confirmation"
    End If
    rst.MoveNext
    Wend
    rst.Close
    Call reload
    Call clear
    cmdupdate.Enabled = False
    picinfo.Enabled = False
    End If
End If
End Sub

Private Sub Form_Load()
Call reload
End Sub
Function reload()
rst.Open "Select * from tblsupp", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
    lst.ListItems.Add , , rst!suppid
    lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Name
    lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!address
    lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!person
    lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!contact
    rst.MoveNext
Wend
rst.Close
End Function
Function clear()
txtsuppid = ""
txtname = ""
txtaddress = ""
txtperson = ""
txtno = ""
End Function



Private Sub lst_Click()
txtsuppid = lst.SelectedItem
txtname = lst.SelectedItem.SubItems(1)
txtaddress = lst.SelectedItem.SubItems(2)
txtperson = lst.SelectedItem.SubItems(3)
txtno = lst.SelectedItem.SubItems(4)
End Sub

Private Sub lst_DblClick()
cmdedit.Enabled = True
cmddelete.Enabled = True
cmdedit.SetFocus
End Sub
