VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8835
   ClientLeft      =   210
   ClientTop       =   1275
   ClientWidth     =   11160
   ClipControls    =   0   'False
   DrawMode        =   6  'Mask Pen Not
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   4560
      Top             =   6360
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   720
      ScaleHeight     =   315
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   6600
      Width           =   4455
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VVD pvt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Automated Inventory System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      TabIndex        =   0
      Top             =   2040
      Width           =   3375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim userMsg As String





Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
  UIStyle.ApplyModernTheme Me

End Sub

Private Sub Timer1_Timer()
    Dim pct As Long
    pct = Int((Picture2.Width / 4000) * 100)
    If pct > 100 Then pct = 100
    
    If pct < 20 Then
        Label3.Caption = pct & "% - Loading Database..."
    ElseIf pct < 40 Then
        Label3.Caption = pct & "% - Loading Medicines..."
    ElseIf pct < 60 Then
        Label3.Caption = pct & "% - Initializing Inventory..."
    ElseIf pct < 80 Then
        Label3.Caption = pct & "% - Loading Reports..."
    Else
        Label3.Caption = pct & "% - Preparing Dashboard..."
    End If

    If Picture2.Width <= 4000 Then
        Picture2.Width = Picture2.Width + 60 ' slightly slower, smoother loading animation
    Else
        Unload Me
        frmLogin.Show
    End If
End Sub


