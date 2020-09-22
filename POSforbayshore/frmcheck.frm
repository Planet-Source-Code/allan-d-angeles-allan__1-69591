VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcheck 
   Appearance      =   0  'Flat
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Product Stocks"
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
   Begin MSComctlLib.ListView lst 
      Height          =   10695
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   18865
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
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Code"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Supplier"
         Object.Width           =   2540
      EndProperty
   End
   Begin Project1.chameleonButton cmdclose 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   10320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Close"
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
   Begin VB.PictureBox picinfo 
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin Project1.chameleonButton cmdsearch 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "SEARCH NOW"
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
         BCOL            =   8454143
         FCOL            =   0
      End
      Begin VB.TextBox txtoption 
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2400
         Width           =   165
      End
      Begin VB.TextBox txtsearch 
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
         Left            =   360
         TabIndex        =   1
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label lblpcode 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Search by: Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2880
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbldesc 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Search by: Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Product"
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
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Label lblcritical 
      BackColor       =   &H00C0FFFF&
      Caption         =   "crit"
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
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   9120
      Width           =   3615
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   7095
      Left            =   120
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label lblsell 
      BackColor       =   &H00C0FFFF&
      Caption         =   "see"
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
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   8040
      Width           =   3615
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label lblprice 
      BackColor       =   &H00C0FFFF&
      Caption         =   "price"
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
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   6960
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Original Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label lblstocks 
      BackColor       =   &H00C0FFFF&
      Caption         =   "stoks"
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
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label lblsupp 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   4800
      Width           =   3615
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   2415
   End
End
Attribute VB_Name = "frmcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmdclose_Click()
Unload Me
mdimain.Show
End Sub

Private Sub cmdsearch_Click()
If txtsearch = "" Then
    MsgBox "Invalid character!", vbCritical, "System Message"
    txtsearch.SetFocus
ElseIf txtoption = "" Then
    MsgBox "Please select search by!", vbCritical, "System Message"
    
Else
    Dim X
    rst.Open "Select * from tblprod where " & txtoption & " like '%" & txtsearch.Text & "%'", con, adOpenDynamic, adLockOptimistic
    lst.ListItems.clear
    While rst.EOF = False
    lst.ListItems.Add , , rst!pcode
    lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!Desc
    lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , rst!supp
    rst.MoveNext
    X = 1
    Wend
    rst.Close
    MsgBox "Search Complete!", vbInformation, "Confirmation"
    If X <> 1 Then
    MsgBox "No results!", vbInformation, "Confirmation"
    txtsearch.SetFocus
    End If
End If
End Sub

Function clear()
lblsupp = ""
lblstocks = ""
lblprice = ""
lblsell = ""
lblcritical = ""
lst.ListItems.clear
txtsearch = ""
txtoption = ""
End Function

Private Sub Form_Load()
Call clear
End Sub

Private Sub lbldesc_Click()
txtoption = ""
txtoption = "desc"
lblpcode.BorderStyle = 0
lbldesc.BorderStyle = 1
txtsearch.SetFocus
End Sub

Private Sub lbldesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldesc.BackStyle = 1
End Sub

Private Sub lblpcode_Click()
txtoption = ""
txtoption = "pcode"
lblpcode.BorderStyle = 1
lbldesc.BorderStyle = 0
txtsearch.SetFocus
End Sub

Private Sub lblpcode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblpcode.BackStyle = 1
End Sub



Private Sub lst_Click()
rst.Open "Select * from tblprice", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If lst.SelectedItem = rst!pcode Then
lblstocks = rst!stocks
lblprice = rst!org_price
lblsell = rst!selling
lblcritical = rst!critical
End If
rst.MoveNext
Wend
rst.Close

rst.Open "Select * from tblsupp", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
If lst.SelectedItem.SubItems(2) = rst!suppid Then
lblsupp = rst!Name
End If
rst.MoveNext
Wend
rst.Close
End Sub

Private Sub picinfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldesc.BackStyle = 0
lblpcode.BackStyle = 0
End Sub


Private Sub txtsearch_Click()
Call clear
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdsearch_Click
End Sub
