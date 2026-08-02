VERSION 5.00
Begin VB.Form FrmAdminPass 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin"
   ClientHeight    =   3390
   ClientLeft      =   2460
   ClientTop       =   4155
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   0
         Top             =   720
         Width           =   2925
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1320
         Width           =   2925
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please supply Admin username and password to creat new user's account."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   960
      End
   End
End
Attribute VB_Name = "FrmAdminPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
If MsgBox("Are you sure you want to exit?", vbOKCancel) = vbOK Then
    Unload Me
End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHandler
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim flag As Boolean
Dim inputHash As String
Dim storedSalt As String
Dim storedHash As String

con.Open (Constring)
rs.Open "Select * from Admin", con, adOpenKeyset, adLockOptimistic

While Not rs.EOF
    If Me.txtUserName.Text = rs!UserName Then
        storedSalt = rs!PasswordSalt & ""
        storedHash = rs!PasswordHash & ""
        
        If storedSalt <> "" Then
            ' Salted SHA-256 validation
            inputHash = Module1.HashPassword(Me.txtPassword.Text & storedSalt)
            If inputHash = storedHash Then
                flag = True
            End If
        ElseIf storedHash <> "" Then
            ' Unsalted SHA-256 validation (migration path)
            inputHash = Module1.HashPassword(Me.txtPassword.Text)
            If inputHash = storedHash Then
                flag = True
                ' Migrate to salted SHA-256 now
                Randomize
                Dim newSalt As String
                newSalt = CStr(Int((999999 - 100000 + 1) * Rnd() + 100000)) ' Generate 6-digit salt
                rs!PasswordSalt = newSalt
                rs!PasswordHash = Module1.HashPassword(Me.txtPassword.Text & newSalt)
                rs!Password = "" ' Clear plaintext password
                rs.Update
            End If
        ElseIf Me.txtPassword.Text = (rs!Password & "") Then
            ' Legacy plaintext fallback - Migrate to Salted SHA-256!
            flag = True
            Randomize
            Dim tempSalt As String
            tempSalt = CStr(Int((999999 - 100000 + 1) * Rnd() + 100000))
            rs!PasswordSalt = tempSalt
            rs!PasswordHash = Module1.HashPassword(Me.txtPassword.Text & tempSalt)
            rs!Password = "" ' Clear plaintext password
            rs.Update
        End If
    End If
    rs.MoveNext
Wend

rs.Close
con.Close
Set rs = Nothing
Set con = Nothing

If flag = True Then
    Unload Me
    Select Case Module1.AdminTargetForm
        Case "frmNewlog"
            frmNewlog.Show vbModal
        Case "frmMaster"
            frmMaster.Show vbModal
        Case "frmEdit"
            frmEdit.Show vbModal
        Case "frmdate"
            frmdate.Show vbModal
        Case "RptAllsales"
            RptAllsales.Show vbModal
        Case Else
            frmNewlog.Show vbModal
    End Select
Else
    MsgBox "Invalid username or password. Please try again.", vbCritical, "Authentication Failed"
End If
Exit Sub

ErrHandler:
MsgBox "An authentication error occurred: " & Err.Description, vbCritical, "Security Error"
End Sub

Private Sub Form_Load()
  UIStyle.ApplyModernTheme Me
End Sub
