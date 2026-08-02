VERSION 5.00
Begin VB.Form frmFind 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   6405
   ClientLeft      =   3435
   ClientTop       =   2625
   ClientWidth     =   9855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      ItemData        =   "frmfind.frx":0000
      Left            =   240
      List            =   "frmfind.frx":0002
      Style           =   1  'Simple Combo
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   3120
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox Text1 
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
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   2535
      End
      Begin VB.TextBox txtPrice 
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
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox txtexdate 
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
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtbal 
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
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtshelf 
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
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtpdate 
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
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Product Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Price:"
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
         Left            =   1680
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Production Date:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Shelf:"
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
         Left            =   1800
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Expiry Date:"
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
         Left            =   840
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Stock Balance:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   2520
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   550
      Left            =   7920
      TabIndex        =   0
      Top             =   600
      Width           =   1200
      Style           =   1  'Graphical
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
If MsgBox("Are you sure you want to exit?", vbOKCancel) = vbOK Then
Unload Me
End If

End Sub

Private Sub Combo1_Click()
Frame1.Visible = 1
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select * from Master where Item = '" & Replace(Combo1, "'", "''") & "'", con, adOpenKeyset, adLockOptimistic
If rs.EOF = False And rs.BOF <> True Then
Me.txtpdate = rs!MfdDate
Me.txtbal = rs!Qty
Me.txtexdate = rs!ExpDate
Me.txtshelf = rs!Shelf
Me.txtPrice = rs!Price
Me.Text1 = rs!DrugName
End If
If Val(txtbal.Text) = 0 Then
MsgBox Me.Combo1 & " is not available in stock", vbInformation, "Stock Query"
Else: MsgBox "There are " & Val(txtbal.Text) & " " & Me.Combo1 & "(s) in stock", vbInformation, "Stock Query"
End If
Set con = Nothing

Set rs = Nothing
End Sub

Private Sub Combo1_GotFocus()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)
Combo1.Clear
rs.Open "Select Item from Master order by Item", con, adOpenKeyset, adLockOptimistic

While rs.EOF <> True And rs.BOF <> True
Combo1.AddItem rs!Item
rs.MoveNext
Wend
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_Load()
  UIStyle.ApplyModernTheme Me
Me.Combo1.Clear
End Sub



Private Sub Label4_Click()

End Sub
