VERSION 5.00
Begin VB.Form p1 
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   9375
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Farmacia De Borja"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Sta.Cruz Laguna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   4320
         TabIndex        =   23
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "TIN# 221-934-853-000VAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         Top             =   1320
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   9375
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Height          =   3855
         Left            =   5280
         TabIndex        =   26
         Top             =   4080
         Width           =   3855
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   32
            Top             =   2760
            Width           =   2415
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   31
            Top             =   1800
            Width           =   2415
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   28
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cashier:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.TextBox txttotalprice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   9240
         TabIndex        =   25
         Text            =   "P0.00"
         Top             =   8760
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   7200
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   5880
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   4800
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ListBox List10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   3120
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ListBox List11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   1320
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ListBox List12 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00C0E0FF&
         Height          =   3855
         Left            =   120
         ScaleHeight     =   3795
         ScaleWidth      =   4995
         TabIndex        =   1
         Top             =   4080
         Width           =   5055
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tender:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   615
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Change:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   615
            Left            =   120
            TabIndex        =   5
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2280
            TabIndex        =   4
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2280
            TabIndex        =   3
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2280
            TabIndex        =   2
            Top             =   1680
            Width           =   2535
         End
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MARAMING SALAMAT PO!!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   34
         Top             =   8400
         Width           =   6375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "THIS SERVE AS YOUR OFFICIAL RECEIPT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   33
         Top             =   8040
         Width           =   6375
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Transaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3000
         TabIndex        =   20
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7200
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5880
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4800
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3120
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1320
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "p1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
cmdPrint.Visible = False

PrintForm

Unload Me
End Sub

Private Sub Form_Load()
  UIStyle.ApplyModernTheme Me
Label2.Caption = frmMenu.Label3.Caption
Label5.Caption = frmMenu.l3.Caption
Label6.Caption = frmMenu.l2.Caption


End Sub
