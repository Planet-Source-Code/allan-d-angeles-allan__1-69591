VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmuser 
   BackColor       =   &H80000010&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Maintenance"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtoption 
      Height          =   375
      Left            =   9960
      TabIndex        =   23
      Top             =   1080
      Width           =   150
   End
   Begin VB.PictureBox picinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   7335
      TabIndex        =   10
      Top             =   240
      Width           =   7335
      Begin VB.PictureBox Picture1 
         Height          =   615
         Left            =   1440
         ScaleHeight     =   555
         ScaleWidth      =   5595
         TabIndex        =   22
         Top             =   2520
         Width           =   5655
         Begin VB.OptionButton optcashier 
            Caption         =   "Cashier"
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
            Left            =   3480
            TabIndex        =   6
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton optadmin 
            Caption         =   "Administrator"
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
            Left            =   720
            TabIndex        =   5
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.TextBox txtconfirm 
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
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "#"
         TabIndex        =   4
         Top             =   2040
         Width           =   5655
      End
      Begin VB.TextBox txtpassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "#"
         TabIndex        =   3
         Top             =   1560
         Width           =   5655
      End
      Begin VB.TextBox txtuname 
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
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
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
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   5655
      End
      Begin VB.TextBox txtuserid 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   5655
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000009&
         Height          =   3255
         Left            =   0
         Top             =   0
         Width           =   7215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Level"
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
         Left            =   0
         TabIndex        =   21
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   1095
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
         Left            =   360
         TabIndex        =   19
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   18
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         TabIndex        =   17
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         TabIndex        =   15
         Top             =   1200
         Width           =   135
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
         Left            =   720
         TabIndex        =   14
         Top             =   720
         Width           =   615
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
         Left            =   600
         TabIndex        =   13
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Id"
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
         Left            =   -360
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
   End
   Begin Project1.chameleonButton cmdadd 
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   3720
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
      Left            =   7680
      TabIndex        =   7
      Top             =   4320
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
   Begin Project1.chameleonButton cmddelete 
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   4920
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
   Begin Project1.chameleonButton cmdclose 
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   6600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Close"
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
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "Double click the selected item to edit or delete"
      Top             =   3600
      Width           =   7215
      _ExtentX        =   12726
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User_Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "User_Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "User Level"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   3495
      Left            =   7560
      Top             =   3600
      Width           =   1815
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
picinfo.Enabled = True
Call clear
rst.Open "Select * from tbluser", con, adOpenDynamic, adLockOptimistic
rst.MoveLast
txtuserid = "USER" + Format(Val(Right(rst!userid, 3)) + 1, "00#")
rst.Close
txtname.SetFocus
cmdsave.Enabled = True
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Dim q As VbMsgBoxResult
q = MsgBox("Delete account " + lst.SelectedItem + " ?", vbQuestion + vbYesNo, "System Message")
If q = vbYes Then
    rst.Open "Select * from tbluser", con, adOpenDynamic, adLockOptimistic
    While rst.EOF = False
        If lst.SelectedItem = rst!userid Then
            rst.Delete
            rst.Update
            MsgBox "Deleted!", vbInformation, "Confirmation"
        End If
        rst.MoveNext
    Wend
    rst.Close
    Call reload
    cmddelete.Enabled = False
End If
End Sub

Private Sub cmdsave_Click()
If txtuserid = "" Then
    MsgBox "Please check user id!", vbCritical, "System Message"
    cmdadd.SetFocus
ElseIf txtname = "" Then
    MsgBox "Please check name!", vbCritical, "System Message"
    txtname.SetFocus
ElseIf txtuname = "" Then
    MsgBox "Please check user name!", vbCritical, "System Message"
    txtuname.SetFocus
ElseIf txtpassword = "" Then
    MsgBox "Please check user password!", vbCritical, "System Message"
    txtpassword.SetFocus
ElseIf txtoption = "" Then
    MsgBox "Please check user level!", vbCritical, "System Message"
    optadmin.SetFocus
Else
    Dim q As VbMsgBoxResult
    q = MsgBox("Save record " + txtuserid + " ? ", vbQuestion + vbYesNo, "System Message")
    If q = vbYes Then
    rst.Open "Select * from tbluser", con, adOpenDynamic, adLockPessimistic
    rst.AddNew
    rst!userid = txtuserid
    rst!Name = txtname
    rst!UserName = txtuname
    rst!Password = txtpassword
    rst!Level = txtoption
    rst.Update
    rst.Close
    MsgBox "Saved!", vbInformation, "Confirmation"
    Call clear
    Call reload
    picinfo.Enabled = False
    cmdsave.Enabled = False
    End If
End If
End Sub

Private Sub Form_Load()
Call reload
End Sub
Function clear()
txtuserid = ""
txtname = ""
txtuname = ""
txtpassword = ""
txtconfirm = ""
optadmin = False
optcashier = False
End Function
Function reload()
rst.Open "Select * from tbluser", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
lst.ListItems.Add , , rst!userid
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Name
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!UserName
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Level
rst.MoveNext
Wend
rst.Close
End Function



Private Sub lst_DblClick()
cmddelete.Enabled = True
cmddelete.SetFocus
End Sub

Private Sub optadmin_Click()
txtoption = ""
txtoption = 1
End Sub

Private Sub optcashier_Click()
txtoption = ""
txtoption = 2
End Sub



Private Sub txtconfirm_LostFocus()
If txtpassword <> txtconfirm Then
    MsgBox "Please check your password", vbCritical, "System Message"
    txtpassword = ""
    txtconfirm = ""
    txtpassword.SetFocus
End If
End Sub
