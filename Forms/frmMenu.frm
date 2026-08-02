VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00000000&
   Caption         =   "Main Menu"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   720
   ClientWidth     =   17625
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   17625
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   3000
      Top             =   3240
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   3360
      ScaleHeight     =   915
      ScaleWidth      =   6315
      TabIndex        =   19
      Top             =   3600
      Width           =   6375
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Farmacia De Borja"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   600
      Top             =   4200
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   240
      TabIndex        =   15
      Top             =   5040
      Width           =   9975
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         FillColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   840
         ScaleHeight     =   675
         ScaleWidth      =   8595
         TabIndex        =   16
         Top             =   360
         Width           =   8655
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "ADVERTISING:"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Left            =   2760
            TabIndex        =   17
            Top             =   120
            Width           =   2895
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   19095
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Log-Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   16080
         MouseIcon       =   "frmMenu.frx":3E0CB
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         MaskColor       =   &H0080FFFF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "New User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Programer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11880
      Top             =   0
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Log-Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17040
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About Programmers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14160
      TabIndex        =   26
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User's Manual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11880
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "user name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   5640
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label l2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "H"" h ""m"" min ""s"" s """
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label l3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy,MMMM dd, dddd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunewstaff 
         Caption         =   "New Staff"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnumsales 
      Caption         =   "S&ales"
      Begin VB.Menu mnumsale 
         Caption         =   "Make Sales"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnustock 
      Caption         =   "&Stock"
      Begin VB.Menu mnunewstock 
         Caption         =   "Add New Stock"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnubalance 
         Caption         =   "Stock Balance"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuedit 
         Caption         =   "Edit Stock Master"
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuallsale_Click()
RptAllsales.Show vbModal
End Sub





Private Sub Command1_Click()
frmSale.Show vbModal
End Sub



Private Sub Command2_Click()
Module1.AdminTargetForm = "frmNewlog"
FrmAdminPass.Show vbModal
End Sub



Private Sub Command3_Click()
Module1.AdminTargetForm = "frmMaster"
FrmAdminPass.Show vbModal
End Sub



Private Sub Command4_Click()
frmFind.Show vbModal
End Sub



Private Sub Command5_Click()
frmHelp.Show
End Sub



Private Sub Command6_Click()
If MsgBox("Are you sure you want to exit?", vbOKCancel) = vbOK Then
  frmSplash.Show
  Unload Me
End If
End Sub



Private Sub Command7_Click()
Programmers.Show
End Sub



Private Sub mnuall_Click()
Module1.AdminTargetForm = "RptAllsales"
FrmAdminPass.Show vbModal
End Sub

Private Sub Form_Activate()

l3 = Date
Label2.Caption = "Welcome"
Label3.Caption = frmLogin.txtUserName.Text




End Sub

Private Sub Form_Load()
  UIStyle.ApplyModernTheme Me

End Sub





Private Sub mnubalance_Click()
frmFind.Show
End Sub

Private Sub mnuedit_Click()
Module1.AdminTargetForm = "frmEdit"
FrmAdminPass.Show vbModal
End Sub

Private Sub mnuexit_Click()
If MsgBox("Are you sure you want to exit?", vbOKCancel) = vbOK Then
    ' Unload all forms to safely close the application instead of using End
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next frm
End If
End Sub

Private Sub mnumsale_Click()
frmSale.Show vbModal
End Sub

Private Sub mnunewstaff_Click()
Module1.AdminTargetForm = "frmNewlog"
FrmAdminPass.Show vbModal
End Sub

Private Sub mnunewstock_Click()
Module1.AdminTargetForm = "frmMaster"
FrmAdminPass.Show vbModal
End Sub

Private Sub mnuoption_Click()
Module1.AdminTargetForm = "frmdate"
FrmAdminPass.Show vbModal
End Sub

Private Sub Timer1_Timer()
l2 = Time
End Sub

Private Sub Timer2_Timer()
If Label6.Left >= -2000 Then
    Label6.Left = Label6.Left - 100
ElseIf Label6.Left <= 15000 Then
    Label6.Left = 6000
End If

End Sub

Private Sub Timer3_Timer()
If Label7.Left >= -1000 Then
    Label7.Left = Label6.Left - 100
ElseIf Label7.Left <= 10000 Then
    Label7.Left = 6000
End If
End Sub
