VERSION 5.00
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit"
   ClientHeight    =   7245
   ClientLeft      =   1695
   ClientTop       =   3390
   ClientWidth     =   9210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   " "
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
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtprodname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtPPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   5040
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox txtPQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox txtexpiry 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txtpdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtshelf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3840
         Width           =   1935
      End
      Begin VB.ComboBox cmbpid 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Price:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Quantity:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Present Quantity:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Shelf:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Production Date:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6480
      TabIndex        =   15
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6480
      TabIndex        =   14
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblsave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6360
      TabIndex        =   13
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbpid_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select * from Master where DrugName = '" & Replace(Me.cmbpid, "'", "''") & "'", con, adOpenKeyset, adLockOptimistic

If rs.EOF <> True And rs.BOF <> True Then
Me.txtProdName = rs.Fields("Item")
Me.txtexpiry = rs.Fields("ExpDate")
Me.txtpdate = rs.Fields("MfdDate")
Me.txtshelf = rs.Fields("Shelf")
Me.txtPQty = rs.Fields("Qty")
Me.txtPPrice = rs.Fields("Price")
End If
End Sub

Private Sub cmbpid_GotFocus()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select DrugName from Master", con, adOpenKeyset, adLockOptimistic
While rs.EOF <> True
cmbpid.AddItem rs!DrugName
rs.MoveNext
Wend

End Sub



Private Sub lblDelete_Click()
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Delete * from Master where DrugName = '" & Replace(Me.cmbpid, "'", "''") & "'", con, adOpenKeyset, adLockOptimistic

    Me.cmbpid.Clear
    Me.txtProdName = ""
    Me.txtshelf = ""
    Me.txtpdate = ""
    Me.txtexpiry = ""
    Me.txtPQty = ""
    Me.txtPPrice = ""
    MsgBox "Item Deleted", vbInformation, "Deletion"
Set rs = Nothing
Set con = Nothing
End Sub

Private Sub lblEdit_Click()

End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblDelete.ForeColor = &H0& Then
lblDelete.ForeColor = &HFF&
End If
End Sub

Private Sub lblExit_Click()
If lblExit.Caption = "&Exit" Then
    If MsgBox("This program will be terminated", vbOKCancel, "Pharmacy") = vbOK Then
        Unload Me
    End If
ElseIf lblExit.Caption = "&Cancel" Then
      Me.cmbpid.Clear
      Me.txtProdName = ""
      Me.txtshelf = ""
     Me.txtpdate = ""
     Me.txtexpiry = ""
     Me.txtPQty = ""
     Me.txtPPrice = ""
     Me.Text2 = ""
    lblsave.Caption = "&Edit"
     lblDelete.Enabled = True
    Me.cmbpid.Locked = True
    Me.txtProdName.Locked = True
    Me.txtexpiry.Locked = True
    Me.txtpdate.Locked = True
    Me.txtshelf.Locked = True
    Me.txtPPrice.Locked = True
    lblExit.Caption = "&Exit"
End If
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblExit.ForeColor = &H0& Then
lblExit.ForeColor = &HFF&
End If
End Sub

Private Sub lblSave_Click()
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)
If lblsave.Caption = "&Edit" Then
    lblsave.Caption = "&Save"
   lblExit.Caption = "&Cancel"
    lblDelete.Enabled = True
    Me.cmbpid.Locked = False
    Me.txtProdName.Locked = False
    Me.txtexpiry.Locked = False
    Me.txtpdate.Locked = False
    Me.txtshelf.Locked = False
    Me.txtPPrice.Locked = False
    Me.Text2.Locked = False
    ElseIf lblsave.Caption = "&Save" Then
rs.Open "Select * from Master where DrugName = '" & Replace(Me.cmbpid, "'", "''") & "'", con, adOpenKeyset, adLockOptimistic
If rs.EOF <> True And rs.BOF <> True Then

With rs
    .Fields("Item") = Me.txtProdName
    .Fields("DrugName") = Me.cmbpid
    .Fields("Shelf") = Me.txtshelf
    .Fields("MfdDate") = Me.txtpdate
    .Fields("ExpDate") = Me.txtexpiry
    .Fields("Qty") = rs.Fields("Qty") + Val(Me.Text2)
    .Fields("Price") = Me.txtPPrice
    .Update
    .close
End With
    Me.cmbpid.Clear
    Me.txtProdName = ""
    Me.txtshelf = ""
    Me.txtpdate = ""
    Me.txtexpiry = ""
    Me.txtPQty = ""
    Me.txtPPrice = ""
    Me.Text2 = ""
    lblsave.Caption = "&Edit"
    lblExit.Caption = "&Exit"
    lblDelete.Enabled = True
    MsgBox "item qty update"
Set rs = Nothing
Set con = Nothing
ElseIf lblExit.Caption = "&Exit" Then
    Me.cmbpid.Clear
    Me.txtProdName = ""
    Me.txtshelf = ""
    Me.txtpdate = ""
    Me.txtexpiry = ""
    Me.txtPQty = ""
    Me.txtPPrice = ""
    Me.Text2 = ""
    lblsave.Caption = "&Edit"
     lblDelete.Enabled = True
    Me.cmbpid.Locked = True
    Me.txtexpiry.Locked = True
    Me.txtpdate.Locked = True
    Me.txtshelf.Locked = True
    Me.txtPPrice.Locked = True
    Me.Text2.Locked = True
    lblExit.Caption = "&Exit"
End If
End If
End Sub

Private Sub lblsave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblsave.ForeColor = &H0& Then
lblsave.ForeColor = &HFF&
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()
If IsNumeric(Text2.Text) = False Then
Text2.Text = ""
Text2.SetFocus
End If
End Sub

Private Sub txtPPrice_Change()
If IsNumeric(txtPPrice.Text) = False Then
txtPPrice.Text = ""
txtPPrice.SetFocus
End If
End Sub


Private Sub Form_Load()
  UIStyle.ApplyModernTheme Me
End Sub
