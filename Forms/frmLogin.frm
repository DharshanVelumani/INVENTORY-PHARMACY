VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3375
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1994.062
   ScaleMode       =   0  'User
   ScaleWidth      =   6013.948
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   855
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   855
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtPassword 
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
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2925
      End
      Begin VB.TextBox txtUserName 
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
         Height          =   465
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   2925
      End
      Begin VB.CheckBox chkRemember 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remember Me"
         Height          =   350
         Left            =   1320
         TabIndex        =   3
         Top             =   1600
         Width           =   2000
      End
      Begin VB.CommandButton cmdShowPassword 
         Caption         =   "Show"
         Height          =   450
         Left            =   3800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   500
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
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   960
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
         TabIndex        =   3
         Top             =   600
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmLogin"
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
    Dim user As String
    Dim pwd As String
    Dim hashedPwd As String
    Dim rsUser As ADODB.Recordset
    Dim failedAttempts As Integer
    Dim lockoutVal As Variant
    Dim dbPwd As String
    Dim dbEncPwd As String
    Dim loginSuccess As Boolean
    
    user = Trim(Me.txtUserName.Text)
    pwd = Trim(Me.txtPassword.Text)
    
    If user = "" Or pwd = "" Then
        MsgBox "Please enter both username and password.", vbExclamation, "Authentication"
        Exit Sub
    End If
    
    ' Query database for user information
    Set rsUser = Module1.GetRecordset("SELECT Username, Password, EncryptedPassword, Role, FailedLogins, LockoutTime FROM Login WHERE Username='" & Module1.EscapeSQL(user) & "'")
    
    If rsUser Is Nothing Or rsUser.EOF Then
        ' Check Admin table if not found in Login table (legacy compatibility)
        Set rsUser = Module1.GetRecordset("SELECT Username, Password FROM Admin WHERE Username='" & Module1.EscapeSQL(user) & "'")
        If rsUser Is Nothing Or rsUser.EOF Then
            MsgBox "Invalid username or password. Please try again.", vbCritical, "Authentication Failed"
            ' Log failed login attempt
            Module1.ExecuteSQL "INSERT INTO LoginHistory (Username, LogTime, Status, IPAddress) VALUES ('" & Module1.EscapeSQL(user) & "', Now(), 'FAILED_USER_NOT_FOUND', '127.0.0.1')"
            Exit Sub
        End If
    End If
    
    ' Check lockout status if using Login table
    On Error Resume Next
    lockoutVal = rsUser.Fields("LockoutTime").Value
    failedAttempts = Val(rsUser.Fields("FailedLogins").Value)
    On Error GoTo ErrHandler
    
    If Not IsNull(lockoutVal) Then
        If CDate(lockoutVal) > Now() Then
            Dim diffMin As Long
            diffMin = DateDiff("n", Now(), CDate(lockoutVal))
            MsgBox "Account is locked due to multiple failed login attempts. Please try again in " & diffMin & " minutes.", vbCritical, "Account Locked"
            Exit Sub
        End If
    End If
    
    ' Retrieve passwords
    dbPwd = ""
    dbEncPwd = ""
    On Error Resume Next
    dbPwd = rsUser.Fields("Password").Value
    dbEncPwd = rsUser.Fields("EncryptedPassword").Value
    On Error GoTo ErrHandler
    
    ' Validate password
    hashedPwd = Module1.HashPassword(pwd)
    loginSuccess = False
    
    If dbEncPwd <> "" Then
        If dbEncPwd = hashedPwd Then
            loginSuccess = True
        End If
    End If
    
    ' Fallback to plain-text check for legacy accounts
    If Not loginSuccess And dbPwd <> "" Then
        If dbPwd = pwd Then
            loginSuccess = True
            ' Automatically migrate plain-text password to hashed password
            Module1.ExecuteSQL "UPDATE Login SET EncryptedPassword = '" & hashedPwd & "' WHERE Username = '" & Module1.EscapeSQL(user) & "'"
        End If
    End If
    
    If loginSuccess Then
        ' Handle Remember Me persistence
        If chkRemember.Value = vbChecked Then
            SaveSetting "PharmacyERP", "Login", "RememberedUser", user
        Else
            SaveSetting "PharmacyERP", "Login", "RememberedUser", ""
        End If
        
        ' Reset failed attempts
        Module1.ExecuteSQL "UPDATE Login SET FailedLogins = 0, LockoutTime = NULL WHERE Username = '" & Module1.EscapeSQL(user) & "'"
        
        ' Log success
        Module1.ExecuteSQL "INSERT INTO LoginHistory (Username, LogTime, Status, IPAddress) VALUES ('" & Module1.EscapeSQL(user) & "', Now(), 'SUCCESS', '127.0.0.1')"
        Module1.ExecuteSQL "INSERT INTO AuditLog (Username, Activity, LogTime) VALUES ('" & Module1.EscapeSQL(user) & "', 'Logged into the system successfully', Now())"
        
        Me.Hide
        frmMenu.Show
    Else
        ' Increment failed attempts
        failedAttempts = failedAttempts + 1
        If failedAttempts >= 3 Then
            Dim lockoutLimit As String
            lockoutLimit = Format(DateAdd("n", 15, Now()), "yyyy-mm-dd hh:nn:ss") ' 15 minutes lockout
            Module1.ExecuteSQL "UPDATE Login SET FailedLogins = " & failedAttempts & ", LockoutTime = #" & lockoutLimit & "# WHERE Username = '" & Module1.EscapeSQL(user) & "'"
            MsgBox "Invalid password. Your account has been locked for 15 minutes due to 3 failed attempts.", vbCritical, "Account Locked"
        Else
            Module1.ExecuteSQL "UPDATE Login SET FailedLogins = " & failedAttempts & " WHERE Username = '" & Module1.EscapeSQL(user) & "'"
            MsgBox "Invalid username or password. Remaining attempts: " & (3 - failedAttempts), vbCritical, "Authentication Failed"
        End If
        
        ' Log failed login attempt
        Module1.ExecuteSQL "INSERT INTO LoginHistory (Username, LogTime, Status, IPAddress) VALUES ('" & Module1.EscapeSQL(user) & "', Now(), 'FAILED_INVALID_PASSWORD', '127.0.0.1')"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "Login Execution Error: " & Err.Description, vbCritical, "System Error"



End Sub

Private Sub Form_Load()
    UIStyle.ApplyModernTheme Me
    Main
    
    Dim savedUser As String
    savedUser = GetSetting("PharmacyERP", "Login", "RememberedUser", "")
    
    If savedUser <> "" Then
        txtUserName.Text = savedUser
        chkRemember.Value = vbChecked
        txtPassword.Text = ""
        Me.Show
        txtPassword.SetFocus
    Else
        txtUserName.Text = ""
        txtPassword.Text = ""
    End If
End Sub

Private Sub cmdShowPassword_Click()
    If txtPassword.PasswordChar = "*" Then
        txtPassword.PasswordChar = ""
        cmdShowPassword.Caption = "Hide"
    Else
        txtPassword.PasswordChar = "*"
        cmdShowPassword.Caption = "Show"
    End If
    txtPassword.SetFocus
End Sub
