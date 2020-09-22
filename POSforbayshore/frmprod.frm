VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmprod 
   Appearance      =   0  'Flat
   BackColor       =   &H80000010&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product Maintenance"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1695
      ScaleWidth      =   7335
      TabIndex        =   7
      Top             =   120
      Width           =   7335
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   5655
      End
      Begin VB.TextBox txtdesc 
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
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   5655
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
         Left            =   1560
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000009&
         Height          =   1695
         Left            =   0
         Top             =   0
         Width           =   7335
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
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label4 
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
         TabIndex        =   15
         Top             =   1200
         Width           =   135
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
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   480
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblsupp 
         BackStyle       =   0  'Transparent
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
         Left            =   4080
         TabIndex        =   11
         Top             =   1080
         Width           =   3255
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   30
      Left            =   2880
      TabIndex        =   6
      Top             =   3480
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Project1.chameleonButton cmdadd 
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   2040
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
   Begin MSComctlLib.ListView lst 
      Height          =   3495
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Double click the selected item to edit or delete"
      Top             =   1920
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6165
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pcode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Supplier"
         Object.Width           =   2540
      EndProperty
   End
   Begin Project1.chameleonButton cmdsave 
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   2760
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
      Left            =   5880
      TabIndex        =   2
      Top             =   3480
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
      Left            =   5880
      TabIndex        =   3
      Top             =   4200
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
      Left            =   5880
      TabIndex        =   4
      Top             =   4920
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
      Left            =   5760
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "frmprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbosupp_Change()
rst.Open "Select * from tblsupp", con, adOpenDynamic
While rst.EOF = False
    If cbosupp = rst!suppid Then
        lblsupp = rst!Name
    End If
    rst.MoveNext
Wend
rst.Close

End Sub

Private Sub cbosupp_Click()
rst.Open "Select * from tblsupp", con, adOpenDynamic
While rst.EOF = False
    If cbosupp = rst!suppid Then
        lblsupp = rst!Name
    End If
    rst.MoveNext
Wend
rst.Close
End Sub

Private Sub cmdadd_Click()
picinfo.Enabled = True
Call clear
rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
rst.MoveLast
txtpcode = "PROD" + Format(Val(Right(rst!pcode, 3)) + 1, "00#")
txtdesc.SetFocus
rst.Close
cmdadd.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdupdate.Enabled = False
cmdsave.Enabled = True
End Sub

Private Sub cmddelete_Click()
Dim q As VbMsgBoxResult
q = MsgBox("Delete Record " + txtpcode + "?", vbQuestion + vbYesNo, "System Message")
If q = vbYes Then
    
    rst.Open "Select * from tblprod", con, adOpenDynamic
    While rst.EOF = False
    If txtpcode = rst!pcode Then
        rst.Delete
        rst.Update
        MsgBox "Sucessfully Deleted!", vbInformation, "Confirmation"
        cmddelete.Enabled = False
        cmdedit.Enabled = False
    End If
    rst.MoveNext
    Wend
    rst.Close
    Call reload
    txtpcode = ""
    txtdesc = ""
    cbosupp = ""
    lblsupp = ""
    
End If
End Sub

Private Sub cmdedit_Click()
picinfo.Enabled = True
txtdesc.SetFocus
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdupdate.Enabled = True
End Sub

Private Sub cmdsave_Click()
If txtpcode = "" Then
    MsgBox "Please check the product code!", vbCritical, "System Message"
    cmdadd.SetFocus
ElseIf txtdesc = "" Then
    MsgBox "Please check product description!", vbCritical, "System Message"
    txtdesc.SetFocus
ElseIf cbosupp = "" Then
    MsgBox "Please check product Supplier!", vbCritical, "System Message"
    txtdesc.SetFocus
Else
    Dim q As VbMsgBoxResult
    q = MsgBox("Save recorde " + txtpcode + " ?", vbQuestion + vbYesNo, "System Message")
    If q = vbYes Then
        rst.Open "Select * from tblprod", con, adOpenDynamic, adLockPessimistic
        rst.AddNew
        
        rst!pcode = txtpcode
        rst!Desc = txtdesc
        rst!supp = cbosupp
        rst.Update
        rst.Close
        
        txtpcode = ""
        txtdesc = ""
        cmdsave.Enabled = False
        MsgBox "Saved!", vbInformation, "Confirmation"
         Call reload
         picinfo.Enabled = False
    End If

End If



End Sub

Private Sub cmdupdate_Click()
If txtpcode = "" Then
    MsgBox "Please check the product code!", vbCritical, "System Message"
    cmdadd.SetFocus
ElseIf txtdesc = "" Then
    MsgBox "Please check product description!", vbCritical, "System Message"
    txtdesc.SetFocus
ElseIf cbosupp = "" Then
    MsgBox "Please check product Supplier!", vbCritical, "System Message"
    txtdesc.SetFocus
Else
    Dim q As VbMsgBoxResult
    q = MsgBox("Update recorde " + txtpcode + " ?", vbQuestion + vbYesNo, "System Message")
    If q = vbYes Then
        rst.Open "Select * from tblprod", con, adOpenDynamic, adLockPessimistic
        While rst.EOF = False
        If rst!pcode = txtpcode Then
        rst!Desc = txtdesc
        rst!supp = cbosupp
        rst.Update
        
        
        txtpcode = ""
        txtdesc = ""
        cmdupdate.Enabled = False
        MsgBox "Saved!", vbInformation, "Confirmation"
         
         End If
         rst.MoveNext
         Wend
         rst.Close
         Call reload
    End If

End If



End Sub

Private Sub Form_Load()
Call reload
Call cboreload
End Sub
Function clear()
txtpcode = ""
txtdesc = ""
cbosupp = ""
lblsupp = ""
End Function

Function reload()
rst.Open "Select * from tblprod", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear

While rst.EOF = False
lst.ListItems.Add , , rst!pcode
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!supp
rst.MoveNext
Wend
rst.Close
End Function

Function cboreload()
rst.Open "Select * from tblsupp", con, adOpenDynamic, adLockOptimistic
cbosupp.clear
While rst.EOF = False
cbosupp.AddItem rst!suppid
rst.MoveNext
Wend
rst.Close
End Function



Private Sub lst_Click()
txtpcode = lst.SelectedItem
txtdesc = lst.SelectedItem.SubItems(1)
cbosupp = lst.SelectedItem.SubItems(2)
End Sub

Private Sub lst_DblClick()
cmdedit.Enabled = True
cmddelete.Enabled = True
End Sub
