VERSION 5.00
Begin VB.MDIForm mdimain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Point of Sales and Inventory System for Bayshore Water Refilling Station"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7980
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10215
      Left            =   0
      ScaleHeight     =   10215
      ScaleWidth      =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   3240
      Begin VB.Label lblcs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Stocks"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label lbll 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Logout"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   9600
         Width           =   2535
      End
      Begin VB.Label lblh 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   6720
         Width           =   2535
      End
      Begin VB.Label lbldr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Report"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   6120
         Width           =   2535
      End
      Begin VB.Label lblor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Orders Report"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   5520
         Width           =   2535
      End
      Begin VB.Label lblsr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Report"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Label lblrts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Return to supplier"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   4320
         Width           =   2535
      End
      Begin VB.Label lblrfc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Return from customer"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label lbld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label lblo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Orders"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label lblprice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Price Maintenance"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lblpm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Maintenance"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.PictureBox picuser 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000018&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   7920
      TabIndex        =   0
      Top             =   10215
      Width           =   7980
      Begin VB.Label lbldep 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Current User"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6360
         TabIndex        =   4
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " User Level"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblname 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Current User"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current User"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File Maintenance"
      Begin VB.Menu mnuprod 
         Caption         =   "Product Maintenance"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuunit 
         Caption         =   "Unit Maintenance"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnusupp 
         Caption         =   "Supplier Maintenance"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuuser 
         Caption         =   "User Maintenance"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuprice 
         Caption         =   "Price Maintenance"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnutrans 
      Caption         =   "&Transaction"
      Begin VB.Menu mnusales 
         Caption         =   "Sales"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuorder 
         Caption         =   "Orders"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnudelivery 
         Caption         =   "Delivery"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnureturn 
         Caption         =   "Return"
         Begin VB.Menu mnurfc 
            Caption         =   "Return From Custumer"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnusupplier 
            Caption         =   "Return To Supplier"
            Shortcut        =   {F5}
         End
      End
   End
   Begin VB.Menu mnustocks 
      Caption         =   "&Product"
      Begin VB.Menu mnucheck 
         Caption         =   "Check Stocks"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Reports"
      Begin VB.Menu mnusalesreport 
         Caption         =   "Sales Report"
      End
      Begin VB.Menu mnuordersreport 
         Caption         =   "Orders Report"
      End
      Begin VB.Menu mnudeliveryreport 
         Caption         =   "Delivery Report"
      End
      Begin VB.Menu mnurr 
         Caption         =   "Return Reports"
         Begin VB.Menu mnurfcreport 
            Caption         =   "Return from Customer Report"
         End
         Begin VB.Menu mnurtsreports 
            Caption         =   "Return to Supplier Reports"
         End
      End
   End
   Begin VB.Menu mnuutilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuhelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnudate 
         Caption         =   "Date and Time Setting"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About us"
      End
   End
   Begin VB.Menu mnulogout 
      Caption         =   "&Logout"
   End
End
Attribute VB_Name = "mdimain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblcs_Click()
mdimain.Hide
frmcheck.Show
End Sub

Private Sub lblcs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblcs.BackStyle = 1
lblcs.BorderStyle = 1
End Sub

Private Sub lbld_Click()
frmdelivery.Show
End Sub

Private Sub lbld_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbld.BackStyle = 1
lbld.BorderStyle = 1
End Sub

Private Sub lbldr_Click()
frmdreport.Show
mdimain.Hide

End Sub

Private Sub lbldr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldr.BackStyle = 1
lbldr.BorderStyle = 1
End Sub

Private Sub lblh_Click()
frmhelp.Show
mdimain.Hide

End Sub

Private Sub lblh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblh.BackStyle = 1
lblh.BorderStyle = 1
End Sub

Private Sub lbll_Click()
frmlogin.Show
Unload Me
End Sub

Private Sub lbll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbll.BackStyle = 1
lbll.BorderStyle = 1
End Sub

Private Sub lblo_Click()
frmorder.Show
End Sub

Private Sub lblo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblo.BackStyle = 1
lblo.BorderStyle = 1
End Sub

Private Sub lblor_Click()
frmoreport.Show
mdimain.Hide
End Sub

Private Sub lblor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblor.BackStyle = 1
lblor.BorderStyle = 1
End Sub

Private Sub lblpm_Click()
frmprod.Show
frmprod.cmdadd.SetFocus
End Sub

Private Sub lblpm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblpm.BackStyle = 1
lblpm.BorderStyle = 1
End Sub

Private Sub lblprice_Click()
frmprice.Show
End Sub

Private Sub lblprice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblprice.BackStyle = 1
lblprice.BorderStyle = 1
End Sub

Private Sub lblrfc_Click()
frmrfc.Show
mdimain.Hide

End Sub

Private Sub lblrfc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblrfc.BackStyle = 1
lblrfc.BorderStyle = 1
End Sub

Private Sub lblrts_Click()
frmrts.Show
mdimain.Hide
End Sub

Private Sub lblrts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblrts.BackStyle = 1
lblrts.BorderStyle = 1
End Sub

Private Sub lbls_Click()
frmsales.Show
frmsales.lblname = lblname
mdimain.Hide
End Sub

Private Sub lbls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbls.BackStyle = 1
lbls.BorderStyle = 1
End Sub

Private Sub lblsr_Click()
frmsalesreport.Show
mdimain.Hide


End Sub

Private Sub lblsr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblsr.BackStyle = 1
lblsr.BorderStyle = 1
End Sub

Private Sub MDIForm_Load()
connection.connect
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblcs.BackStyle = 0
lblcs.BorderStyle = 0
lblpm.BackStyle = 0
lblpm.BorderStyle = 0
lblprice.BackStyle = 0
lblprice.BorderStyle = 0
lbls.BackStyle = 0
lbls.BorderStyle = 0
lblo.BackStyle = 0
lblo.BorderStyle = 0
lbld.BackStyle = 0
lbld.BorderStyle = 0
lblrfc.BackStyle = 0
lblrfc.BorderStyle = 0
lblrts.BackStyle = 0
lblrts.BorderStyle = 0
lblsr.BackStyle = 0
lblsr.BorderStyle = 0
lblor.BackStyle = 0
lblor.BorderStyle = 0
lbll.BackStyle = 0
lbll.BorderStyle = 0
lbldr.BackStyle = 0
lbldr.BorderStyle = 0
lblh.BackStyle = 0
lblh.BorderStyle = 0

End Sub

Private Sub mnuabout_Click()
frmaboutus.Show
End Sub

Private Sub mnucheck_Click()
mdimain.Hide
frmcheck.Show
End Sub

Private Sub mnudate_Click()
Shell ("C:\WINDOWS\system32\control.exe date/time")
End Sub

Private Sub mnudelivery_Click()
frmdelivery.Show
End Sub



Private Sub mnudeliveryreport_Click()
frmdreport.Show
mdimain.Hide
End Sub

Private Sub mnuhelp_Click()
frmhelp.Show
mdimain.Hide
End Sub

Private Sub mnulogout_Click()
frmlogin.Show
Unload Me
End Sub

Private Sub mnuorder_Click()
frmorder.Show
End Sub

Private Sub mnuordersreport_Click()
frmoreport.Show
mdimain.Hide
End Sub

Private Sub mnuprice_Click()
frmprice.Show
End Sub

Private Sub mnuprod_Click()
frmprod.Show
frmprod.cmdadd.SetFocus
End Sub

Private Sub mnurfc_Click()
frmrfc.Show
mdimain.Hide
End Sub

Private Sub mnurfcreport_Click()
frmrfc_report.Show
mdimain.Hide
End Sub

Private Sub mnurtsreports_Click()
frmrts_report.Show
mdimain.Hide
End Sub

Private Sub mnusales_Click()
frmsales.Show
frmsales.lblname = lblname
mdimain.Hide
End Sub

Private Sub mnusalesreport_Click()
frmsalesreport.Show
mdimain.Hide
End Sub

Private Sub mnusupp_Click()
frmsupp.Show
End Sub

Private Sub mnusupplier_Click()
frmrts.Show
mdimain.Hide
End Sub

Private Sub mnuunit_Click()
frmunit.Show
frmunit.cmdadd.SetFocus
End Sub

Private Sub mnuuser_Click()
frmuser.Show
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblcs.BackStyle = 0
lblcs.BorderStyle = 0
lblpm.BackStyle = 0
lblpm.BorderStyle = 0
lblprice.BackStyle = 0
lblprice.BorderStyle = 0
lbls.BackStyle = 0
lbls.BorderStyle = 0
lblo.BackStyle = 0
lblo.BorderStyle = 0
lbld.BackStyle = 0
lbld.BorderStyle = 0
lblrfc.BackStyle = 0
lblrfc.BorderStyle = 0
lblrts.BackStyle = 0
lblrts.BorderStyle = 0
lblsr.BackStyle = 0
lblsr.BorderStyle = 0
lblor.BackStyle = 0
lblor.BorderStyle = 0
lbll.BackStyle = 0
lbll.BorderStyle = 0
lbldr.BackStyle = 0
lbldr.BorderStyle = 0
lblh.BackStyle = 0
lblh.BorderStyle = 0


End Sub
