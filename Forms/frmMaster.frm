VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMaster 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Products"
   ClientHeight    =   7245
   ClientLeft      =   1260
   ClientTop       =   1530
   ClientWidth     =   8835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   4455
         Left            =   720
         ScaleHeight     =   4395
         ScaleWidth      =   7395
         TabIndex        =   3
         Top             =   1440
         Width           =   7455
         Begin VB.TextBox txtdname 
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
            Height          =   495
            Left            =   3840
            TabIndex        =   7
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox txtqty 
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
            Height          =   495
            Left            =   3840
            TabIndex        =   6
            Top             =   2520
            Width           =   3255
         End
         Begin VB.TextBox txtshelf 
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
            Height          =   495
            Left            =   3840
            TabIndex        =   5
            Top             =   3120
            Width           =   3255
         End
         Begin VB.TextBox txtPrice 
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
            Height          =   495
            Left            =   3840
            TabIndex        =   4
            Top             =   3720
            Width           =   3255
         End
         Begin MSComCtl2.DTPicker DTPed 
            Height          =   615
            Left            =   3840
            TabIndex        =   8
            Top             =   1800
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   1085
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            Format          =   72024064
            CurrentDate     =   39485
         End
         Begin MSComCtl2.DTPicker DTPpd 
            Height          =   615
            Left            =   3840
            TabIndex        =   9
            Top             =   1080
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   1085
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            Format          =   72024064
            CurrentDate     =   39485
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Shelf:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   15
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   14
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   13
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Production Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   12
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Category:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   11
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Price:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   3720
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6000
         Width           =   1695
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00C0FFC0&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6000
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
If MsgBox("This Program will be terminated", vbOKCancel, "Pharmacy") = vbOK Then
    Unload Me
End If

End Sub

Private Sub cmdsave_Click()
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

If Me.txtdname = "" Or Me.txtqty = "" Or Me.txtshelf = "" Then
    MsgBox "Please Complete the Data", vbInformation, "Pharmacy"
Else
rs.Open "Select * from Master where DrugName= '" & Replace(txtdname, "'", "''") & "'", con, adOpenKeyset, adLockOptimistic
If rs.EOF <> True And rs.BOF <> True Then
With rs
    .Fields("Qty") = rs.Fields("qty") + Val(Me.txtqty.Text)
    .Update
    .close
End With
Me.txtdname = ""
Me.txtqty = ""
Me.txtshelf = ""
Me.txtPrice = ""
Set rs = Nothing
Set con = Nothing

Else

rs1.Open "Select * from Master", con, adOpenKeyset, adLockOptimistic
With rs1
    .AddNew
    .Fields("Item") = Me.txtdname
    .Fields("DrugName") = Me.txtdname
    .Fields("MfdDate") = Me.DTPpd
    .Fields("ExpDate") = Me.DTPed
    .Fields("Shelf") = Me.txtshelf
    .Fields("Qty") = Me.txtqty
    .Fields("Price") = Me.txtPrice
    
    .Update
    .Close
MsgBox "Saved"
Me.txtdname = ""
Me.txtqty = ""
Me.txtshelf = ""
Me.txtPrice = ""
End With
Set rs1 = Nothing
Set con = Nothing
End If
End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
  UIStyle.ApplyModernTheme Me
txtqty.Text = "0"


End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub txtPrice_Change()
If IsNumeric(txtPrice.Text) = False Then
txtPrice.Text = ""
txtPrice.SetFocus
End If
End Sub

Private Sub txtqty_Change()
If IsNumeric(txtqty.Text) = False Then
txtqty.Text = ""
txtqty.SetFocus
End If
End Sub
