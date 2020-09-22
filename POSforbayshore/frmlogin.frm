VERSION 5.00
Begin VB.Form frmlogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000010&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Login"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.chameleonButton cmdlogin 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&LOGIN"
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
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtuname 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin Project1.chameleonButton cmdcancel 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&CANCEL"
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
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin VB.Label Label18 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
txtuname = ""
txtpass = ""
End Sub

Private Sub cmdlogin_Click()
connection.connect
rst.Open "Select * from tbluser", con, adOpenDynamic, adLockOptimistic
While rst.EOF = False
    Dim X
    If UCase(txtuname) = UCase(rst!UserName) And UCase(txtpass) = UCase(rst!Password) Then
    If rst!Level = 1 Then
        mdimain.Show
        mdimain.lblname = rst!Name
        mdimain.lbldep = "ADMIN"
        Unload Me
    ElseIf rst!Level = 2 Then
        mdimain.Show
        mdimain.mnufile.Enabled = False
        mdimain.lblname = rst!Name
        mdimain.lbldep = "Cashier"
        mdimain.mnufile.Enabled = False
        mdimain.mnureport.Enabled = False
        mdimain.lblprice.Enabled = False
        mdimain.lblsr.Enabled = False
        mdimain.lblor.Enabled = False
        mdimain.lbldr.Enabled = False
        Unload Me
    End If
    X = 1
    End If
    rst.MoveNext
Wend
rst.Close

If X <> 1 Then
    MsgBox "Please Check Your Username or Password!", vbCritical, "LOGIN FAILED"
    txtpass = ""
    txtuname = ""
    txtuname.SetFocus
End If
End Sub



Private Sub imgclose_Click()
End
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdlogin_Click
End Sub

Private Sub txtuname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdlogin_Click
End Sub
