VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmprice 
   BackColor       =   &H80000010&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Price Maintenance "
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtcritical 
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
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   2520
         Width           =   2535
      End
      Begin VB.ComboBox cbomark 
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
         ItemData        =   "frmprice.frx":0000
         Left            =   1680
         List            =   "frmprice.frx":001F
         TabIndex        =   14
         Top             =   1560
         Width           =   2535
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
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox cbopcode 
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
         Left            =   1680
         TabIndex        =   6
         Top             =   120
         Width           =   2535
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
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtsell 
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
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Crtical level"
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
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Org. Price"
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
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price"
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
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Markup Price"
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
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000009&
         Height          =   3015
         Left            =   0
         Top             =   0
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin Project1.chameleonButton cmdadd 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      btype           =   3
      tx              =   "Add &New"
      enab            =   -1  'True
      font            =   "frmprice.frx":0059
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      fcol            =   0
   End
   Begin Project1.chameleonButton cmdsave 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      btype           =   3
      tx              =   "&Save"
      enab            =   0   'False
      font            =   "frmprice.frx":0085
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      fcol            =   0
   End
   Begin Project1.chameleonButton cmdedit 
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   3600
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      btype           =   3
      tx              =   "&Edit"
      enab            =   0   'False
      font            =   "frmprice.frx":00B1
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      fcol            =   0
   End
   Begin Project1.chameleonButton cmddelete 
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      btype           =   3
      tx              =   "&Delete"
      enab            =   0   'False
      font            =   "frmprice.frx":00DD
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      fcol            =   0
   End
   Begin Project1.chameleonButton cmdupdate 
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   3600
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      btype           =   3
      tx              =   "&Update"
      enab            =   0   'False
      font            =   "frmprice.frx":0109
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      fcol            =   0
   End
   Begin MSComctlLib.ListView lst 
      Height          =   3135
      Left            =   4560
      TabIndex        =   17
      ToolTipText     =   "Double click the selected item to edit or delete"
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
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
         Text            =   "Pcode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Original Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Markup_Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Selling Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Critical Level"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   120
      Top             =   3480
      Width           =   9015
   End
End
Attribute VB_Name = "frmprice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cbomark_Click()
txtsell = Val(cbomark) * Val(txtprice)
txtsell = Val(txtprice) + Val(txtsell)
End Sub

Private Sub cbopcode_Change()
rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If cbopcode = rst!pcode Then
txtpcode = rst!Desc
End If
rst.MoveNext
Wend
rst.Close
End Sub

Private Sub cbopcode_Click()
rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If cbopcode = rst!pcode Then
txtpcode = rst!Desc
End If
rst.MoveNext
Wend
rst.Close

End Sub

Private Sub cmdadd_Click()
picinfo.Enabled = True
Call clear
cbopcode.SetFocus
cmdsave.Enabled = True

End Sub

Private Sub cmddelete_Click()
Dim q As VbMsgBoxResult
q = MsgBox("Delete?", vbQuestion + vbYesNo, "System Message")
If q = vbYes Then
    rst.Open "Select * from tblprice", con, adOpenDynamic, adLockOptimistic
    While rst.EOF = False
    If cbopcode = rst!pcode Then
        rst.Delete
        rst.Update
        MsgBox "Deleted!", vbInformation, "Confirmation"
        cmddelete.Enabled = False
        cmdedit.Enabled = False
    End If
    rst.MoveNext
    Wend
    rst.Close
    Call clear
    Call reload
End If
End Sub

Private Sub cmdedit_Click()
picinfo.Enabled = True
cmdupdate.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
cbopcode.Locked = True
End Sub

Private Sub cmdsave_Click()
Dim x
rst.Open "Select * from tblprice", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If cbopcode = rst!pcode Then
    MsgBox "Product Code is already exist!", vbCritical, "System Message"
    cbopcode = ""
    txtpcode = ""
    Call clear
    cbopcode.SetFocus
    x = 1
End If
rst.MoveNext
Wend
rst.Close
If x <> 1 Then
If txtprice = "" Then
    MsgBox "Please check original price!", vbCritical, "System Message"
    txtprice.SetFocus
ElseIf cbomark = "" Then
    MsgBox "Please check markup price!", vbCritical, "System Message"
    cbomark.SetFocus
ElseIf txtcritical = "" Then
    MsgBox "Please enter critical level!", vbCritical, "System Message"
    txtcritical.SetFocus
Else
rst.Open "Select * from tblprice", con, adOpenDynamic, adLockPessimistic
rst.AddNew
rst!pcode = cbopcode
rst!org_price = txtprice
rst!markup = cbomark.Text
rst!selling = txtsell
rst!critical = txtcritical
rst.Update
rst.Close
MsgBox "Saved!", vbInformation, "Confirmation"
Call reload
Call clear
cmdsave.Enabled = False
End If
End If
End Sub

Private Sub cmdupdate_Click()
If txtprice = "" Then
    MsgBox "Please check original price!", vbCritical, "System Message"
    txtprice.SetFocus
ElseIf cbomark = "" Then
    MsgBox "Please check markup price!", vbCritical, "System Message"
    cbomark.SetFocus
ElseIf txtcritical = "" Then
    MsgBox "Please enter critical level!", vbCritical, "System Message"
    txtcritical.SetFocus
Else
rst.Open "Select * from tblprice", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If rst!pcode = cbopcode Then
rst!org_price = txtprice
rst!markup = cbomark
rst!selling = txtsell
rst!critical = txtcritical
rst.Update
MsgBox "Saved!", vbInformation, "Confirmation"
End If
rst.MoveNext
Wend
rst.Close
Call reload
Call clear
cmdupdate.Enabled = False
cbopcode.Locked = False

End If
End Sub

Private Sub Form_Load()
rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
cbopcode.clear
While rst.EOF = False
cbopcode.AddItem rst!pcode
rst.MoveNext
Wend
rst.Close
Call reload
End Sub
Function clear()
cbopcode = ""
txtpcode = ""
txtprice = ""
cbomark = ""
txtsell = ""
txtcritical = ""
End Function

Function reload()
rst.Open "Select * from tblprice", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
lst.ListItems.Add , , rst!pcode
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!org_price
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!markup
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!selling
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!critical
rst.MoveNext
Wend
rst.Close
End Function

Private Sub lst_Click()
cbopcode = lst.SelectedItem
txtprice = lst.SelectedItem.SubItems(1)
cbomark = lst.SelectedItem.SubItems(2)
txtsell = lst.SelectedItem.SubItems(3)
txtcritical = lst.SelectedItem.SubItems(4)
picinfo.Enabled = False
cmdsave.Enabled = False
End Sub

Private Sub lst_DblClick()
cmdedit.Enabled = True
cmddelete.Enabled = True
End Sub

