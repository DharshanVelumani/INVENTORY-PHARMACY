VERSION 5.00
Begin VB.Form frmSale 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Sales"
   ClientHeight    =   10815
   ClientLeft      =   1650
   ClientTop       =   2580
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10815
   ScaleWidth      =   16050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   10935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16095
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00000000&
         Height          =   2775
         Left            =   4680
         ScaleHeight     =   2715
         ScaleWidth      =   4755
         TabIndex        =   48
         Top             =   4320
         Visible         =   0   'False
         Width           =   4815
         Begin VB.CommandButton Command8 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   54
            Top             =   2280
            Width           =   2175
      Style           =   1  'Graphical
         End
         Begin VB.CommandButton Command6 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   2280
            Width           =   2295
      Style           =   1  'Graphical
         End
         Begin VB.TextBox Text1 
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
            Height          =   1455
            Left            =   120
            TabIndex        =   49
            Top             =   840
            Width           =   4455
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            Caption         =   "Enter Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   4455
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         Height          =   3495
         Left            =   5760
         ScaleHeight     =   3435
         ScaleWidth      =   10035
         TabIndex        =   39
         Top             =   2040
         Width           =   10095
         Begin VB.TextBox txtunitprice 
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
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   2280
            Width           =   2775
         End
         Begin VB.TextBox txtqty 
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
            Height          =   480
            Left            =   5760
            TabIndex        =   44
            Top             =   1440
            Width           =   2775
         End
         Begin VB.ComboBox cmbpid 
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
            Height          =   2460
            Left            =   120
            Style           =   1  'Simple Combo
            TabIndex        =   41
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txtTDate 
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
            Height          =   480
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Price:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3240
            TabIndex        =   47
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Quantity:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3240
            TabIndex        =   45
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Product Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   360
            TabIndex        =   43
            Top             =   120
            Width           =   2295
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3240
            TabIndex        =   42
            Top             =   600
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3435
         ScaleWidth      =   5475
         TabIndex        =   30
         Top             =   2040
         Width           =   5535
         Begin VB.TextBox txtProdName 
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
            Height          =   480
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtbalance 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
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
            Left            =   2880
            TabIndex        =   36
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox txtshelf 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   2640
            Width           =   2415
         End
         Begin VB.TextBox txtexpiry 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox txtpdate 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Category:"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   240
            TabIndex        =   59
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Stock Balance:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   240
            TabIndex        =   38
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Shelf:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   240
            TabIndex        =   37
            Top             =   2640
            Width           =   2415
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Production Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   240
            TabIndex        =   34
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Expiry Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   240
            TabIndex        =   33
            Top             =   1440
            Width           =   2415
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   15675
         TabIndex        =   24
         Top             =   9480
         Width           =   15735
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Print"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   11760
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   0
            Width           =   1935
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Void"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   9960
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   0
            Width           =   1815
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Amount"
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
            Left            =   8040
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   0
            Width           =   1935
         End
         Begin VB.CommandButton cmdexit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cancel"
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
            Left            =   13680
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   0
            Width           =   1815
         End
         Begin VB.CommandButton cmdsave 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Save"
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
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   0
            Width           =   1935
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
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   0
            Width           =   2055
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
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   0
            Width           =   2175
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "AddItem"
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
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   5520
         Width           =   10575
         Begin VB.ListBox List6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Left            =   8880
            TabIndex        =   16
            Top             =   1200
            Width           =   1575
         End
         Begin VB.ListBox List5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Left            =   7560
            TabIndex        =   15
            Top             =   1200
            Width           =   1335
         End
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Left            =   6360
            TabIndex        =   14
            Top             =   1200
            Width           =   1215
         End
         Begin VB.ListBox List3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Left            =   4200
            TabIndex        =   13
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Left            =   2040
            TabIndex        =   12
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label25 
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
            Left            =   3240
            TabIndex        =   23
            Top             =   120
            Width           =   4335
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
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
            Height          =   495
            Left            =   8880
            TabIndex        =   22
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
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
            Height          =   495
            Left            =   7560
            TabIndex        =   21
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Quantiry"
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
            Left            =   6360
            TabIndex        =   20
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
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
            Height          =   495
            Left            =   4200
            TabIndex        =   19
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
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
            Height          =   495
            Left            =   2040
            TabIndex        =   18
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   495
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   3855
         Left            =   10800
         ScaleHeight     =   3795
         ScaleWidth      =   4995
         TabIndex        =   3
         Top             =   5640
         Width           =   5055
         Begin VB.Label Label11 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label12 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label14 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   120
            TabIndex        =   7
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label15 
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
            Height          =   975
            Left            =   2280
            TabIndex        =   6
            Top             =   120
            Width           =   2535
         End
         Begin VB.Label Label16 
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
            Height          =   855
            Left            =   2280
            TabIndex        =   5
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label18 
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
            Height          =   855
            Left            =   2280
            TabIndex        =   4
            Top             =   2640
            Width           =   2535
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
         Left            =   9840
         TabIndex        =   1
         Text            =   "P0.00"
         Top             =   240
         Width           =   6135
      End
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   57
         Top             =   600
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3360
         TabIndex        =   56
         Top             =   960
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2880
         TabIndex        =   55
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   6000
         TabIndex        =   2
         Top             =   720
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbpid_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select * from Master where Item = '" & Replace(Me.cmbpid, "'", "''") & "'", con, adOpenKeyset, adLockOptimistic

If rs.EOF <> True And rs.BOF <> True Then
Me.txtProdName = rs.Fields("DrugName")
Me.txtexpiry = rs.Fields("ExpDate")
Me.txtpdate = rs.Fields("MfdDate")
'Me.txtpname = rs.Fields("DrugName")
'Me.txtprice = rs.Fields("Price")
'Me.txtqty = rs.Fields("Qty")
Me.txtshelf = rs.Fields("Shelf")
Me.txtbalance = rs.Fields("Qty")
Me.txtunitprice = rs.Fields("Price")
End If
End Sub

Private Sub cmbpid_GotFocus()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)
cmbpid.Clear
rs.Open "Select Item from Master", con, adOpenKeyset, adLockOptimistic
While rs.EOF <> True
cmbpid.AddItem rs!Item
rs.MoveNext
Wend
End Sub

Private Sub cmdexit_Click()
If MsgBox("Do you really want to terminate this program?", vbOKCancel) = vbOK Then
    Unload Me
End If

End Sub



Private Sub cmdPrint_Click()
p1.Show vbModal
Unload Me
End Sub

Private Sub cmdsave_Click()
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set con = New ADODB.Connection
Set rsbill = New ADODB.Recordset
con.Open (Constring)

rs1.Open "Select * from Sales", con, adOpenKeyset, adLockOptimistic
With rs1
    .AddNew
    .Fields("Total") = Me.txttotalprice
    .Fields("Tender") = Me.Label15
    .Fields("Due") = Me.Label16
    .Fields("Change") = Me.Label18
    .Update
    .Close
End With
MsgBox "Item Quantity Update"
Me.cmbpid.Clear
    Me.txtProdName = ""
    Me.txtunitprice = ""
    Me.txttotalprice = ""
    Me.txtqty = ""
    Me.txtshelf = ""
    Me.txtpdate = ""
    Me.txtexpiry = ""
    txtbalance = ""
    Me.List1.Clear
    Me.List2.Clear
    Me.List3.Clear
    Me.List4.Clear
    Me.List5.Clear
    Me.List6.Clear

End Sub

Private Sub Command1_Click()
'save
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set con = New ADODB.Connection
Set rsbill = New ADODB.Recordset
con.Open (Constring)

rs.Open "Select * from Master where Item = '" & Replace(Me.cmbpid, "'", "''") & "'", con, adOpenKeyset, adLockOptimistic
If rs.EOF <> True And rs.BOF <> True Then

With rs
    .Fields("Qty") = rs.Fields("Qty") - Val(Me.txtqty)
If rs.Fields("qty") <= -1 Then
MsgBox "THAT ITEM IS NOT AVIALABLE ", vbInformation
Exit Sub
End If
    .Update
End With

Set rs = Nothing

rs1.Open "Select * from Sales", con, adOpenKeyset, adLockOptimistic

With rs1
    .AddNew
    .Fields("DrugName") = Me.cmbpid
    .Fields("ProductName") = Me.txtProdName
    .Fields("Price") = Me.txtunitprice
    .Fields("Tprice") = Me.txttotalprice
    .Fields("Qty") = Me.txtqty
    .Fields("Shelf") = Me.txtshelf
    .Fields("ProdDate") = Me.txtpdate
    .Fields("ExpDate") = Me.txtexpiry
    .Fields("seller") = frmLogin.txtUserName.Text
    .Fields("Selldate") = Me.txtTDate
    .Update
    .close
End With

rsbill.Open "Select * from Bill", con, adOpenKeyset, adLockOptimistic

With rsbill
    .AddNew
    .Fields("Description") = Me.cmbpid
    .Fields("Qty") = Me.txtqty
    .Fields("UnitPrice") = Me.txtunitprice
    .Fields("TotalPrice") = Me.txttotalprice
    .Update
    .close
End With
    
Set rs = Nothing
Set con = Nothing
End If

'additem

If Val(txtbalance.Text) < Val(txtqty.Text) Then
    MsgBox "Dont have enough stock", vbInformation, "Confirmation"
    txtqty.Text = ""
ElseIf txtqty.Text = "" Then
     MsgBox "Insert Quantity first", vbInformation, "Confirmation"
     txtqty.SetFocus
Else
    
List1.AddItem txtTDate.Text
List2.AddItem cmbpid.Text
List3.AddItem txtProdName
List4.AddItem txtqty
List5.AddItem txtunitprice
List6.AddItem txttotalprice

'i2 ung list receipt
p1.List12.AddItem txtTDate.Text
p1.List11.AddItem cmbpid.Text
p1.List10.AddItem txtProdName
p1.List9.AddItem txtqty
p1.List8.AddItem txtunitprice
p1.List7.AddItem txttotalprice
End If


txtbalance.Text = Val(txtbalance.Text) - Val(txtqty.Text)
txttotalprice.Text = "0"
For X = 0 To List6.ListCount - 1
    txttotalprice.Text = Val(txttotalprice.Text) + Val(List6.List(X))
Next X

    
For i = 0 To List5.ListCount - 1
Next i
End Sub

Private Sub Command2_Click()
 
   
End Sub

Private Sub Command3_Click()
Picture5.Visible = True

Text1.SetFocus
End Sub

Private Sub Command4_Click()

frmFind.Show vbModal
End Sub

Private Sub Command5_Click()
frmHelp.Show vbModal
End Sub

Private Sub Command6_Click()




If Val(Text1.Text) < Val(txttotalprice.Text) Then
    MsgBox "Amount not enough", vbInformation, "Error"
   
    Text1.Text = ""
Else
Label18.Caption = Val(Text1.Text) - Val(txttotalprice)
Label16.Caption = Text1.Text
Label15.Caption = txttotalprice

'i2 print
p1.Label40.Caption = Val(Text1.Text) - Val(txttotalprice)
p1.Label39.Caption = Text1.Text
p1.Label38.Caption = txttotalprice




Picture5.Visible = False
cmdPrint.Enabled = True
Text1.Text = ""





End If
End Sub

Private Sub Command7_Click()
ans = MsgBox("Create New Transaction?", _
    vbYesNo + vbQuestion, "Confirm")
    
txtpdate.Text = ""
txtexpiry.Text = ""
txtbalance.Text = ""
txtshelf.Text = ""
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear

  cmbpid.SetFocus

If List1.ListCount = 0 Then
MsgBox "No more item!"
Else

txtbalance.Text = Val(txtbalance.Text) + Val(txtqty.Text)

For X = 0 To List1.ListCount
    
Next X

End If
End Sub

Private Sub Command8_Click()
Picture5.Visible = False
End Sub

Private Sub Form_Load()
  UIStyle.ApplyModernTheme Me
Set con = New ADODB.Connection
con.Open (Constring)

con.Execute "Delete * from Bill"
cmbpid.Clear
Me.txtTDate = Date


End Sub

Private Sub Frame1_Click()
Picture5.Visible = False
End Sub

Private Sub Text1_Change()
If IsNumeric(Text1.Text) = False Then
Text1.Text = ""


End If
End Sub

Private Sub txtqty_Change()
If IsNumeric(txtqty.Text) = False Then

txtqty.Text = ""
txtqty.SetFocus
End If
Dim dig$, i, digi$, digits$
If txtqty.Text <> "" Then
    dig$ = Mid(txtqty.Text, Len(txtqty.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtqty.Text) - 1
            digi$ = Mid(txtqty.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtqty.Text = digits$
        txtqty.SelStart = Len(txtqty.Text)
    End If
End If
txttotalprice.Text = Val(txtunitprice.Text) * Val(txtqty.Text)



End Sub
