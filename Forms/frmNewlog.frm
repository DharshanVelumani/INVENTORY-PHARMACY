VERSION 5.00
Begin VB.Form frmNewlog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Staff Registration"
   ClientHeight    =   3585
   ClientLeft      =   1695
   ClientTop       =   2625
   ClientWidth     =   9540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Staff Registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtusername 
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
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtname 
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
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtans 
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
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox txtquest 
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
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   2280
         Width           =   4215
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
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label Label5 
         Caption         =   "Answer"
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
         Left            =   480
         TabIndex        =   10
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Question"
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
         Left            =   360
         TabIndex        =   9
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Full Name"
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmNewlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.Open (Constring)

rs.Open "Select * from Login", con, adOpenKeyset, adLockOptimistic

If Me.txtname = "" Then
MsgBox "Enter name in the name field", vbOKOnly, "Empty field"
Exit Sub
End If

If txtUserName = "" Then
MsgBox "Enter username in the username field", vbOKOnly, "Empty field"
Exit Sub
End If

If txtpass = "" Then
MsgBox "Enter password in the password field", vbOKOnly, "Empty field"
Exit Sub
End If
If txtquest = "" Then
MsgBox "Enter Secret question in the question field", vbOKOnly, "Empty field"
Exit Sub
End If
If txtans = "" Then
MsgBox "Enter Secret answer in the answer field", vbOKOnly, "Empty field"
Exit Sub
End If


With rs
    .AddNew
    .Fields("Fullname") = Me.txtname
    .Fields("Username") = Me.txtUserName
    .Fields("Password") = Me.txtpass
    .Fields("Secquest") = Me.txtquest
    .Fields("Secanswer") = Me.txtans
    .Update
    .close
End With
Me.txtname = ""
Me.txtUserName = ""
Me.txtpass = ""
Me.txtquest = ""
Me.txtans = ""
Set rs = Nothing
Set con = Nothing
End Sub


Private Sub Form_Load()
  UIStyle.ApplyModernTheme Me
End Sub
