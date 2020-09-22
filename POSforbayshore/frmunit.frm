VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmunit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000010&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Unit Maintenance"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtunitid 
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
      TabIndex        =   1
      Top             =   240
      Width           =   4335
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
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
   Begin MSComctlLib.ListView lst 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Double click the selected item to edit or delete"
      Top             =   1440
      Width           =   4095
      _ExtentX        =   7223
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Unit Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2734
      EndProperty
   End
   Begin Project1.chameleonButton cmdadd 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1560
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
      Left            =   4440
      TabIndex        =   6
      Top             =   2280
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
      Left            =   4440
      TabIndex        =   7
      Top             =   3000
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
      Left            =   4440
      TabIndex        =   8
      Top             =   3720
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
      Left            =   4440
      TabIndex        =   9
      Top             =   4440
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1215
      Left            =   120
      Top             =   120
      Width           =   6015
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
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   3495
      Left            =   4320
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Code"
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
      TabIndex        =   3
      Top             =   360
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
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmunit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
Call clear
rst.Open "Select * from tblunit", con, adOpenDynamic, adLockOptimistic
rst.MoveLast
txtunitid = "UNIT" + Format(Val(Right(rst!unitcode, 3)) + 1, "00#")
txtdesc.SetFocus
rst.Close
cmdsave.Enabled = True

End Sub

Private Sub cmddelete_Click()
Dim q As VbMsgBoxResult
q = MsgBox("Delete ?", vbQuestion + vbYesNo, "System Message")
If q = vbYes Then
    rst.Open "Select * from tblunit", con, adOpenDynamic, adLockOptimistic
    While rst.EOF = False
    If txtunitid = rst!unitcode Then
        rst.Delete
        rst.Update
        txtunitid = ""
        txtdesc = ""
        MsgBox "Deleted!", vbInformation, "Confirmation"
        cmddelete.Enabled = False
        cmdedit.Enabled = False
        
    End If
    rst.MoveNext
    Wend
    rst.Close
    Call reload
End If
End Sub

Private Sub cmdedit_Click()
txtdesc.SetFocus
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdupdate.Enabled = True
End Sub

Private Sub cmdsave_Click()
If txtunitid = "" Then
    MsgBox "Please check unit code!", vbCritical, "System Message"
    txtunitid.SetFocus
ElseIf txtdesc = "" Then
    MsgBox "Please check unit description!", vbCritical, "System Message"
    txtdesc.SetFocus
Else
    Dim q As VbMsgBoxResult
    q = MsgBox("Save?", vbQuestion + vbYesNo, "System Message")
    If q = vbYes Then
        rst.Open "Select * from tblunit", con, adOpenDynamic, adLockOptimistic
        rst.AddNew
        rst!unitcode = txtunitid
        rst!Desc = txtdesc
        rst.Update
        rst.Close
        MsgBox "Saved!", vbInformation, "Confirmation"
        cmdsave.Enabled = False
        Call reload
        txtdesc = ""
        txtunitid = ""
    End If
End If
End Sub

Private Sub cmdupdate_Click()
If txtunitid = "" Then
    MsgBox "Please check unit code!", vbCritical, "System Message"
    txtunitid.SetFocus
ElseIf txtdesc = "" Then
    MsgBox "Please check unit description!", vbCritical, "System Message"
    txtdesc.SetFocus
Else
    Dim q As VbMsgBoxResult
    q = MsgBox("Save?", vbQuestion + vbYesNo, "System Message")
    If q = vbYes Then
        rst.Open "Select * from tblunit", con, adOpenDynamic, adLockOptimistic
        While rst.EOF = False
        If rst!unitcode = txtunitid Then
        rst!Desc = txtdesc
        rst.Update
        MsgBox "Updated!", vbInformation, "Confirmation"
        cmdsave.Enabled = False
        End If
        rst.MoveNext
        Wend
        rst.Close
        Call reload
        txtdesc = ""
        txtunitid = ""
        cmdupdate.Enabled = False
    End If
End If
End Sub

Private Sub Form_Load()
Call reload
End Sub
Function clear()
txtunitid = ""
txtdesc = ""
End Function

Function reload()
rst.Open "Select * from tblunit", con, adOpenDynamic, adLockOptimistic
lst.ListItems.clear
While rst.EOF = False
    lst.ListItems.Add , , rst!unitcode
    lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
    rst.MoveNext
Wend
rst.Close
End Function



Private Sub lst_Click()
txtunitid = lst.SelectedItem
txtdesc = lst.SelectedItem.SubItems(1)
End Sub

Private Sub lst_DblClick()
cmdedit.Enabled = True
cmddelete.Enabled = True
End Sub
