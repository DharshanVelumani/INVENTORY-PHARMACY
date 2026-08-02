Attribute VB_Name = "UIStyle"
Option Explicit

' Modern Color Palette (BGR format for VB6 compatibility)
Public Const COLOR_BG As Long = &HFAF7F5&       ' Background Light Gray (#F5F7FA)
Public Const COLOR_WHITE As Long = &HFFFFFF&    ' Cards White (#FFFFFF)
Public Const COLOR_PRIMARY As Long = &HC06515&  ' Medical Blue (#1565C0)
Public Const COLOR_SECONDARY As Long = &HF5A542& ' Secondary Light Blue (#42A5F5)
Public Const COLOR_ACCENT As Long = &HC1AC00&    ' Accent Cyan (#00ACC1)
Public Const COLOR_SUCCESS As Long = &H327D2E&   ' Success Green (#2E7D32)
Public Const COLOR_WARNING As Long = &H25A8F9&   ' Warning Yellow (#F9A825)
Public Const COLOR_DANGER As Long = &H2F2FD3&    ' Danger Red (#D32F2F)
Public Const COLOR_TEXT As Long = &H372D25&     ' Dark Gray Text (#252D37)
Public Const COLOR_SIDEBAR As Long = &H1F1A15&  ' Sidebar Dark Blue/Gray (#151A1F)
Public Const COLOR_CANCEL_BG As Long = &H80726B& ' Cancel Neutral Gray (#6B7280)

Public Const FONT_NAME As String = "Segoe UI"

' GDI32 region API declarations for rounded controls
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' Helper to apply rounded corners to any control with an hWnd property
Public Sub RoundControl(ByVal ctl As Object, ByVal radius As Long)
    On Error Resume Next
    Dim hRgn As Long
    Dim wPx As Long
    Dim hPx As Long
    
    wPx = ctl.Width / Screen.TwipsPerPixelX
    hPx = ctl.Height / Screen.TwipsPerPixelY
    
    hRgn = CreateRoundRectRgn(0, 0, wPx, hPx, radius, radius)
    If hRgn <> 0 Then
        SetWindowRgn ctl.hWnd, hRgn, True
    End If
End Sub

' Entry point to apply modern styles to any form
Public Sub ApplyModernTheme(frm As Form)
    On Error Resume Next
    
    ' Apply general form modernization
    ModernizeForm frm
    
    ' Perform form-specific adjustments
    Select Case frm.Name
        Case "frmMenu"
            ModernizeMenu frm
        Case "frmLogin"
            ModernizeLogin frm
        Case "frmSplash"
            ModernizeSplash frm
        Case "frmSale"
            ModernizeSale frm
        Case "frmMaster"
            ModernizeMaster frm
        Case "frmEdit"
            ModernizeEdit frm
        Case "frmfind"
            ModernizeFind frm
        Case "frmHelp"
            ModernizeHelp frm
        Case "frmNewlog"
            ModernizeNewlog frm
        Case "frmdate"
            ModernizeDate frm
        Case "Programmers"
            ModernizeProgrammers frm
        Case "frmBill"
            ModernizeBill frm
        Case "p1"
            ModernizeP1 frm
        Case "FrmAdminPass", "FrmAdminedit", "FrmAdminvall", "FrmAdminvsales", "frmAddCategory"
            ModernizeAdminDialogs frm
    End Select
End Sub

' Styles controls on a form dynamically
Private Sub ModernizeForm(frm As Form)
    On Error Resume Next
    
    ' Style the form itself
    frm.BackColor = COLOR_BG
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        ' Enforce Segoe UI font and scale sizes dynamically
        ctl.Font.Name = FONT_NAME
        If ctl.Font.Size >= 14 Then
            ctl.Font.Size = 20
            ctl.Font.Bold = True
        ElseIf ctl.Font.Size >= 11 And ctl.Font.Size < 14 Then
            ctl.Font.Size = 11
            ctl.Font.Bold = True
        Else
            ctl.Font.Size = 9
            ctl.Font.Bold = False
        End If
        
        ' Apply control-specific styles
        If TypeOf ctl Is Label Then
            ctl.BackStyle = 0 ' Transparent
            If ctl.ForeColor = &HFF& Or ctl.ForeColor = &HC0& Or ctl.ForeColor = &HFF0000 Then
                ctl.ForeColor = COLOR_DANGER
            Else
                ctl.ForeColor = COLOR_TEXT
            End If
            ctl.Font.Underline = False
            
        ElseIf TypeOf ctl Is CommandButton Then
            ' Set size requirements: standard Height=600, Width=1800
            If frm.Name <> "frmMenu" And frm.Name <> "frmLogin" And frm.Name <> "frmNewlog" And frm.Name <> "frmdate" _
               And frm.Name <> "FrmAdminPass" And frm.Name <> "FrmAdminedit" And frm.Name <> "FrmAdminvall" _
               And frm.Name <> "FrmAdminvsales" And frm.Name <> "frmAddCategory" Then
                ctl.Height = 550
                ctl.Width = 1800
            End If
            
            ' Apply styling based on action categorization if Graphical
            If ctl.Style = 1 Then
                Dim cap As String
                cap = LCase(ctl.Caption)
                
                If cap Like "*save*" Or cap Like "*add*" Or cap Like "*update*" Or cap Like "*search*" Or cap Like "*submit*" Or cap Like "*ok*" Or cap Like "*print*" Or cap Like "*load*" Then
                    ctl.BackColor = COLOR_PRIMARY
                    ctl.ForeColor = COLOR_WHITE
                ElseIf cap Like "*cancel*" Or cap Like "*delete*" Or cap Like "*logout*" Or cap Like "*exit*" Or cap Like "*close*" Or cap Like "*void*" Then
                    ctl.BackColor = COLOR_CANCEL_BG
                    ctl.ForeColor = COLOR_WHITE
                Else
                    ctl.BackColor = COLOR_WHITE
                    ctl.ForeColor = COLOR_TEXT
                End If
                ctl.Font.Bold = True
                RoundControl ctl, 8
            End If
            
        ElseIf TypeOf ctl Is TextBox Then
            ctl.BackColor = COLOR_WHITE
            ctl.ForeColor = COLOR_TEXT
            ctl.Appearance = 0 ' Flat
            ctl.BorderStyle = 1 ' Fixed Single
            RoundControl ctl, 6
            
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.BackColor = COLOR_WHITE
            ctl.ForeColor = COLOR_TEXT
            
        ElseIf TypeOf ctl Is Frame Then
            ctl.BackColor = COLOR_WHITE
            ctl.ForeColor = COLOR_PRIMARY
            ctl.Font.Bold = True
            RoundControl ctl, 12
            
        ElseIf TypeOf ctl Is PictureBox Then
            ctl.BackColor = COLOR_WHITE
            ctl.Appearance = 0 ' Flat
            ctl.BorderStyle = 0 ' No border
            RoundControl ctl, 12
        End If
    Next ctl
End Sub

' Aligns data entry form fields (labels and input controls) to a perfect grid
Private Sub AlignDataEntryFields(frm As Form)
    On Error Resume Next
    Dim ctl As Control
    For Each ctl In frm.Controls
        If TypeOf ctl Is Label Then
            If ctl.Left < 2000 And ctl.Width < 3000 And ctl.Alignment = 0 Then
                ctl.Left = 300
                ctl.Width = 2000
            End If
        ElseIf TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeName(ctl) = "DTPicker" Then
            If ctl.Left >= 1500 And ctl.Left < 5000 Then
                ctl.Left = 2500
                ctl.Width = 3500
            End If
        End If
    Next ctl
End Sub

' Restructures the Main Menu form into a modern ERP Dashboard
Private Sub ModernizeMenu(frm As Form)
    On Error Resume Next
    
    ' Disable scrolling text timers
    frm.Timer2.Enabled = False
    frm.Timer3.Enabled = False
    
    ' Hide floating text guides
    frm.Label9.Visible = False
    frm.Label10.Visible = False
    frm.Label11.Visible = False
    frm.Label12.Visible = False
    frm.Label13.Visible = False
    frm.Label14.Visible = False
    
    ' 1. Left Sidebar Panel
    With frm.Frame1
        .Caption = ""
        .BackColor = COLOR_PRIMARY
        .BorderStyle = 0
        .Left = 0
        .Top = 0
        .Width = 3200
        .Height = frm.ScaleHeight
    End With
    RoundControl frm.Frame1, 0 ' No rounding on sidebar panel edges
    
    ' 2. Sidebar buttons positioning
    frm.Command1.Caption = "Sales POS"
    frm.Command2.Caption = "Manage Stock"
    frm.Command3.Caption = "Add Categories"
    frm.Command4.Caption = "Search Products"
    frm.Command5.Caption = "Staff Users"
    frm.Command7.Caption = "About Devs"
    frm.Command6.Caption = "Logout System"
    
    Dim btns As Variant
    btns = Array(frm.Command1, frm.Command2, frm.Command3, frm.Command4, frm.Command5, frm.Command7, frm.Command6)
    
    Dim i As Integer
    For i = 0 To UBound(btns)
        With btns(i)
            .Style = 1
            .Left = 200
            .Width = 2800
            .Height = 650
            .Top = 1500 + (i * 800)
            If i = 6 Then .Top = 1500 + (i * 800) + 600
            .Font.Name = FONT_NAME
            .Font.Size = 10
            .Font.Bold = True
            .BackColor = &HE5881E& ' Medium Blue (#1E88E5)
            .ForeColor = COLOR_WHITE
            RoundControl btns(i), 8
        End With
    Next i
    
    ' Sidebar Logo Header
    With frm.Label7
        .Alignment = 2 ' Center
        .Caption = "PHARMACY ERP"
        .Font.Name = FONT_NAME
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COLOR_WHITE
        .Left = 100
        .Top = 400
        .Width = 3000
        .Height = 800
        Set .Container = frm.Frame1
    End With
    
    ' 3. Top Header (Picture1)
    With frm.Picture1
        .BackColor = COLOR_WHITE
        .BorderStyle = 0
        .Left = 3200
        .Top = 0
        .Width = frm.ScaleWidth - 3200
        .Height = 1200
    End With
    RoundControl frm.Picture1, 0
    
    ' Header labels
    With frm.Label2
        .Caption = "Logged in as:"
        .Font.Name = FONT_NAME
        .Font.Size = 9
        .ForeColor = COLOR_TEXT
        .Left = 300
        .Top = 250
        .Width = 1500
        .Height = 300
        Set .Container = frm.Picture1
    End With
    
    With frm.Label3
        .Font.Name = FONT_NAME
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = COLOR_PRIMARY
        .Left = 300
        .Top = 550
        .Width = 3000
        .Height = 400
        Set .Container = frm.Picture1
    End With
    
    ' Clock / Date in Header
    With frm.l3
        .Alignment = 1 ' Right
        .Font.Name = FONT_NAME
        .Font.Size = 9
        .ForeColor = COLOR_TEXT
        .Left = frm.Picture1.ScaleWidth - 3200
        .Top = 250
        .Width = 2900
        .Height = 300
        Set .Container = frm.Picture1
    End With
    
    With frm.l2
        .Alignment = 1 ' Right
        .Font.Name = FONT_NAME
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = COLOR_TEXT
        .Left = frm.Picture1.ScaleWidth - 3200
        .Top = 550
        .Width = 2900
        .Height = 400
        Set .Container = frm.Picture1
    End With
    
    ' 4. Dashboard Content Panel (Frame2)
    With frm.Frame2
        .Caption = ""
        .Visible = True
        .BackColor = COLOR_BG
        .Left = 3200
        .Top = 1200
        .Width = frm.ScaleWidth - 3200
        .Height = frm.ScaleHeight - 1200
    End With
    RoundControl frm.Frame2, 0
    
    ' Card 1: Total Medicines (Picture3)
    With frm.Picture3
        .Visible = True
        .BackColor = COLOR_WHITE
        .BorderStyle = 0
        .Left = 400
        .Top = 400
        .Width = 3000
        .Height = 1600
        Set .Container = frm.Frame2
    End With
    RoundControl frm.Picture3, 12
    
    With frm.Label6
        .Alignment = 0
        .Caption = "Total Medicines"
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COLOR_PRIMARY
        .Left = 250
        .Top = 250
        .Width = 2500
        .Height = 300
        Set .Container = frm.Picture3
    End With
    
    With frm.Label9
        .Visible = True
        .Alignment = 0
        .Font.Name = FONT_NAME
        .Font.Size = 18
        .Font.Bold = True
        .ForeColor = COLOR_PRIMARY
        .Left = 250
        .Top = 650
        .Width = 2500
        .Height = 600
        Set .Container = frm.Picture3
    End With
    
    ' Card 2: Total Categories (Label10 card)
    ' Repurpose PictureBox or use Label directly. Let's make Label borderless with soft backcolor
    With frm.Label10
        .Visible = True
        .BackColor = COLOR_WHITE
        .ForeColor = COLOR_TEXT
        .BorderStyle = 0
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .Left = 3800
        .Top = 400
        .Width = 3000
        .Height = 1600
        Set .Container = frm.Frame2
    End With
    RoundControl frm.Label10, 12
    
    ' Card 3: Low Stock Warning (Label11 card)
    With frm.Label11
        .Visible = True
        .BackColor = COLOR_WHITE
        .ForeColor = COLOR_DANGER
        .BorderStyle = 0
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .Left = 7200
        .Top = 400
        .Width = 3000
        .Height = 1600
        Set .Container = frm.Frame2
    End With
    RoundControl frm.Label11, 12
    
    ' Card 4: Today's Sales (Label14 card)
    With frm.Label14
        .Visible = True
        .BackColor = COLOR_WHITE
        .ForeColor = COLOR_SUCCESS
        .BorderStyle = 0
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .Left = 10600
        .Top = 400
        .Width = 3000
        .Height = 1600
        Set .Container = frm.Frame2
    End With
    RoundControl frm.Label14, 12
    
    ' Welcome Header (Label12)
    With frm.Label12
        .Visible = True
        .Alignment = 0
        .Caption = "Welcome to the Healthcare ERP Dashboard"
        .Font.Name = FONT_NAME
        .Font.Size = 16
        .Font.Bold = True
        .ForeColor = COLOR_PRIMARY
        .Left = 400
        .Top = 2400
        .Width = 8000
        .Height = 500
        Set .Container = frm.Frame2
    End With
    
    ' Description (Label13)
    With frm.Label13
        .Visible = True
        .Alignment = 0
        .Caption = "Use the left sidebar navigation menu to access modules. You can process sales POS checkouts, manage inventory stock items, configure product categories, add system users, and print sales reports."
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .ForeColor = COLOR_TEXT
        .Left = 400
        .Top = 3000
        .Width = 9000
        .Height = 800
        Set .Container = frm.Frame2
    End With
    
    UpdateDashboard frm
End Sub

' Styles and hides graphic banners on Splash Form
Private Sub ModernizeSplash(frm As Form)
    On Error Resume Next
    
    ' Deep Medical Blue background for splash screen
    frm.BackColor = COLOR_PRIMARY
    RoundControl frm, 24
    
    ' Company Name Header (Label2)
    With frm.Label2
        .Caption = "VVD Private Limited"
        .Font.Name = FONT_NAME
        .Font.Size = 22
        .Font.Bold = True
        .ForeColor = COLOR_SECONDARY
        .Left = 1000
        .Top = 1500
        .Width = 9000
        .Height = 800
    End With
    
    ' Application Title and tagline (Label1)
    With frm.Label1
        .Caption = "Pharmacy Management System" & vbCrLf & "Smart Healthcare Management Solution"
        .Font.Name = FONT_NAME
        .Font.Size = 15
        .Font.Bold = True
        .ForeColor = COLOR_WHITE
        .Left = 1000
        .Top = 2600
        .Width = 9000
        .Height = 1500
    End With
    
    ' Progress text (Label3)
    With frm.Label3
        .ForeColor = COLOR_WHITE
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .Left = 1000
        .Top = 5000
        .Width = 9000
        .Height = 400
    End With
    
    ' Progress bar containers
    With frm.Picture1
        .BackColor = COLOR_SECONDARY
        .Left = 1000
        .Width = 9160
        .Height = 150
        .Top = 5600
    End With
    RoundControl frm.Picture1, 6
    
    With frm.Picture2
        .BackColor = COLOR_WHITE
        .Left = 0
        .Top = 0
        .Height = 150
    End With
    RoundControl frm.Picture2, 6
    
    ' Dynamically add copyright footer label
    Dim lblCopy As Object
    Set lblCopy = frm.Controls("lblCopy")
    If lblCopy Is Nothing Then
        Set lblCopy = frm.Controls.Add("VB.Label", "lblCopy")
        lblCopy.Alignment = 2 ' Center
        lblCopy.BackStyle = 0
        lblCopy.ForeColor = COLOR_WHITE
        lblCopy.Font.Name = FONT_NAME
        lblCopy.Font.Size = 8
        lblCopy.Caption = "Version 2.0  |  Copyright (c) 2026 VVD Private Limited  |  All Rights Reserved"
        lblCopy.Left = 0
        lblCopy.Width = frm.ScaleWidth
        lblCopy.Top = frm.ScaleHeight - 600
        lblCopy.Height = 400
        lblCopy.Visible = True
    End If
End Sub

' Styles and positions frmLogin split screen
Private Sub ModernizeLogin(frm As Form)
    On Error Resume Next
    
    Dim picLeft As Object
    Dim lblLeft1 As Object
    Dim lblLeft2 As Object
    
    ' Set standard form dimensions
    frm.ScaleMode = 1
    frm.Width = 9000
    frm.Height = 5500
    RoundControl frm, 20
    
    ' Create split left branding panel
    Set picLeft = frm.Controls("picLeft")
    If picLeft Is Nothing Then
        Set picLeft = frm.Controls.Add("VB.PictureBox", "picLeft")
        picLeft.BorderStyle = 0
        picLeft.BackColor = COLOR_PRIMARY
        picLeft.Visible = True
    End If
    picLeft.Move 0, 0, 4200, frm.ScaleHeight
    RoundControl picLeft, 0
    
    ' Add brand labels to the left panel
    Set lblLeft1 = frm.Controls("lblLeft1")
    If lblLeft1 Is Nothing Then
        Set lblLeft1 = frm.Controls.Add("VB.Label", "lblLeft1", picLeft)
        lblLeft1.BackStyle = 0
        lblLeft1.ForeColor = COLOR_WHITE
        lblLeft1.Font.Name = FONT_NAME
        lblLeft1.Font.Size = 20
        lblLeft1.Font.Bold = True
        lblLeft1.Caption = "Pharmacy ERP"
        lblLeft1.Visible = True
    End If
    lblLeft1.Move 400, 1500, 3400, 600
    
    Set lblLeft2 = frm.Controls("lblLeft2")
    If lblLeft2 Is Nothing Then
        Set lblLeft2 = frm.Controls.Add("VB.Label", "lblLeft2", picLeft)
        lblLeft2.BackStyle = 0
        lblLeft2.ForeColor = COLOR_SECONDARY
        lblLeft2.Font.Name = FONT_NAME
        lblLeft2.Font.Size = 9
        lblLeft2.Caption = "VVD Private Limited" & vbCrLf & "Smart Healthcare Solution"
        lblLeft2.Visible = True
    End If
    lblLeft2.Move 400, 2200, 3400, 800
    
    ' Style right credentials container card (Frame1)
    With frm.Frame1
        .Caption = ""
        .BackColor = COLOR_WHITE
        .BorderStyle = 0
        .Left = 4500
        .Top = 400
        .Width = 4000
        .Height = 3200
    End With
    RoundControl frm.Frame1, 16
    
    ' Style user fields
    With frm.txtUserName
        .BackColor = COLOR_BG
        .Left = 1400
        .Top = 800
        .Width = 2200
        .Height = 450
    End With
    RoundControl frm.txtUserName, 6
    
    With frm.txtPassword
        .BackColor = COLOR_BG
        .Left = 1400
        .Top = 1700
        .Width = 1500
        .Height = 450
    End With
    RoundControl frm.txtPassword, 6
    
    With frm.cmdShowPassword
        .Style = 1
        .BackColor = COLOR_BG
        .ForeColor = COLOR_PRIMARY
        .Font.Bold = True
        .Left = 3000
        .Top = 1700
        .Width = 600
        .Height = 450
        RoundControl frm.cmdShowPassword, 6
    End With
    
    With frm.chkRemember
        .BackColor = COLOR_WHITE
        .ForeColor = COLOR_TEXT
        .Left = 1400
        .Top = 2300
        .Width = 2200
        .Height = 350
        .Font.Name = FONT_NAME
        .Font.Size = 9
    End With
    
    ' Style OK & Cancel buttons below Frame1
    With frm.cmdOK
        .Style = 1
        .Left = 4500
        .Top = 3900
        .Width = 1900
        .Height = 550
        .BackColor = COLOR_PRIMARY
        .ForeColor = COLOR_WHITE
        RoundControl frm.cmdOK, 8
    End With
    
    With frm.cmdCancel
        .Style = 1
        .Left = 6600
        .Top = 3900
        .Width = 1900
        .Height = 550
        .BackColor = COLOR_CANCEL_BG
        .ForeColor = COLOR_WHITE
        RoundControl frm.cmdCancel, 8
    End With
End Sub

' Styles Master Product Entry Form
Private Sub ModernizeMaster(frm As Form)
    On Error Resume Next
    With frm.Picture1
        .BackColor = COLOR_WHITE
        .BorderStyle = 1
        .Appearance = 0
        RoundControl frm.Picture1, 12
    End With
    AlignDataEntryFields frm
End Sub

' Styles Edit Stock Form
Private Sub ModernizeEdit(frm As Form)
    On Error Resume Next
    With frm.Frame1
        .BackColor = COLOR_WHITE
        .BorderStyle = 1
        RoundControl frm.Frame1, 12
    End With
    AlignDataEntryFields frm
End Sub

' Styles New Staff Registration Form
Private Sub ModernizeNewlog(frm As Form)
    On Error Resume Next
    With frm.Frame1
        .BackColor = COLOR_WHITE
        .BorderStyle = 1
        .Caption = "New Staff Registration"
        .ForeColor = COLOR_PRIMARY
        RoundControl frm.Frame1, 12
    End With
    AlignDataEntryFields frm
End Sub

' Styles Sales POS Form
Private Sub ModernizeSale(frm As Form)
    On Error Resume Next
End Sub

' Styles Search Product Form
Private Sub ModernizeFind(frm As Form)
    On Error Resume Next
End Sub

' Styles User Help Form
Private Sub ModernizeHelp(frm As Form)
    On Error Resume Next
End Sub

' Styles Date Range Select Dialog
Private Sub ModernizeDate(frm As Form)
    On Error Resume Next
    frm.Frame1.BackColor = COLOR_BG
End Sub

' Styles About Programmers Form
Private Sub ModernizeProgrammers(frm As Form)
    On Error Resume Next
    frm.BackColor = COLOR_WHITE
    RoundControl frm, 16
End Sub

' Styles Billing Receipt Details Form
Private Sub ModernizeBill(frm As Form)
    On Error Resume Next
End Sub

' Styles Visual Billing Layout form
Private Sub ModernizeP1(frm As Form)
    On Error Resume Next
    frm.BackColor = COLOR_WHITE
End Sub

' Styles dialogue credentials gates (admin logins)
Private Sub ModernizeAdminDialogs(frm As Form)
    On Error Resume Next
    
    frm.Frame1.BackColor = COLOR_WHITE
    frm.Frame1.BorderStyle = 1
    frm.Frame1.Left = (frm.ScaleWidth - frm.Frame1.Width) / 2
    frm.Frame1.Top = (frm.ScaleHeight - frm.Frame1.Height - 1000) / 2
    RoundControl frm.Frame1, 12
    
    frm.cmdOK.Left = frm.Frame1.Left + 200
    frm.cmdOK.Top = frm.Frame1.Top + frm.Frame1.Height + 200
    frm.cmdOK.Height = 550
    frm.cmdOK.Width = 1800
    frm.cmdOK.Style = 1
    frm.cmdOK.BackColor = COLOR_PRIMARY
    frm.cmdOK.ForeColor = COLOR_WHITE
    RoundControl frm.cmdOK, 8
    
    frm.cmdCancel.Left = frm.Frame1.Left + frm.Frame1.Width - 2000
    frm.cmdCancel.Top = frm.cmdOK.Top
    frm.cmdCancel.Height = 550
    frm.cmdCancel.Width = 1800
    frm.cmdCancel.Style = 1
    frm.cmdCancel.BackColor = COLOR_CANCEL_BG
    frm.cmdCancel.ForeColor = COLOR_WHITE
    RoundControl frm.cmdCancel, 8
End Sub

' DB helper functions to retrieve dynamic dashboard counts
Public Function GetTotalMedicines() As Long
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim conTemp As New ADODB.Connection
    conTemp.Open Constring
    rsTemp.Open "SELECT COUNT(*) FROM Master", conTemp, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then
        GetTotalMedicines = rsTemp.Fields(0).Value
    End If
    rsTemp.Close
    conTemp.Close
End Function

Public Function GetTotalCategories() As Long
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim conTemp As New ADODB.Connection
    conTemp.Open Constring
    rsTemp.Open "SELECT DISTINCT DrugName FROM Master", conTemp, adOpenStatic, adLockReadOnly
    GetTotalCategories = rsTemp.RecordCount
    rsTemp.Close
    conTemp.Close
End Function

Public Function GetLowStockCount() As Long
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim conTemp As New ADODB.Connection
    conTemp.Open Constring
    rsTemp.Open "SELECT COUNT(*) FROM Master WHERE Qty <= 10", conTemp, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then
        GetLowStockCount = rsTemp.Fields(0).Value
    End If
    rsTemp.Close
    conTemp.Close
End Function

Public Function GetTodaySales() As Double
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim conTemp As New ADODB.Connection
    Dim total As Double
    Dim todayStr As String
    
    total = 0
    todayStr = Format(Date, "yyyy-mm-dd")
    
    conTemp.Open Constring
    rsTemp.Open "SELECT SUM(Val(TPrice)) FROM Sales WHERE SellDate >= #" & todayStr & " 00:00:00# AND SellDate <= #" & todayStr & " 23:59:59#", conTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp.Fields(0).Value) Then
            total = rsTemp.Fields(0).Value
        End If
    End If
    GetTodaySales = total
    rsTemp.Close
    conTemp.Close
End Function

Public Function GetTotalStockValue() As Double
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim conTemp As New ADODB.Connection
    Dim total As Double
    total = 0
    conTemp.Open Constring
    rsTemp.Open "SELECT SUM(Qty * Price) FROM Master", conTemp, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp.Fields(0).Value) Then
            total = rsTemp.Fields(0).Value
        End If
    End If
    GetTotalStockValue = total
    rsTemp.Close
    conTemp.Close
End Function

Public Function GetExpiredCount() As Long
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim conTemp As New ADODB.Connection
    conTemp.Open Constring
    rsTemp.Open "SELECT COUNT(*) FROM Master WHERE ExpDate < #" & Format(Date, "yyyy-mm-dd") & "#", conTemp, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then
        GetExpiredCount = rsTemp.Fields(0).Value
    End If
    rsTemp.Close
    conTemp.Close
End Function

Public Function GetMonthlySales() As Double
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim conTemp As New ADODB.Connection
    Dim total As Double
    total = 0
    conTemp.Open Constring
    rsTemp.Open "SELECT SUM(Val(TPrice)) FROM Sales WHERE Month(SellDate) = " & Month(Date) & " AND Year(SellDate) = " & Year(Date), conTemp, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp.Fields(0).Value) Then
            total = rsTemp.Fields(0).Value
        End If
    End If
    GetMonthlySales = total
    rsTemp.Close
    conTemp.Close
End Function

' Updates the four dashboard card text with real-time database numbers
Public Sub UpdateDashboard(frm As Form)
    On Error Resume Next
    
    Dim totalMeds As Long
    totalMeds = GetTotalMedicines()
    
    Dim totalCats As Long
    totalCats = GetTotalCategories()
    
    Dim lowStock As Long
    lowStock = GetLowStockCount()
    
    Dim todaySales As Double
    todaySales = GetTodaySales()
    
    Dim stockValue As Double
    stockValue = GetTotalStockValue()
    
    Dim expiredCount As Long
    expiredCount = GetExpiredCount()
    
    Dim monthlySales As Double
    monthlySales = GetMonthlySales()
    
    ' Card 1: Total Medicines & Stock Value
    frm.Label9.Caption = totalMeds & " items" & vbCrLf & "   Value: LKR " & Format(stockValue, "#,##0.00")
    
    ' Card 2: Total Categories
    frm.Label10.Caption = vbCrLf & "  Total Categories" & vbCrLf & vbCrLf & "   " & totalCats & " categories"
    
    ' Card 3: Stock Alerts (Low Stock & Expired)
    frm.Label11.Caption = vbCrLf & "  Stock Alerts" & vbCrLf & vbCrLf & "   Low Stock: " & lowStock & vbCrLf & "   Expired: " & expiredCount
    
    ' Card 4: Revenue Metrics (Today's & Monthly)
    frm.Label14.Caption = vbCrLf & "  Revenue Metrics" & vbCrLf & vbCrLf & "   Today: LKR " & Format(todaySales, "0.00") & vbCrLf & "   Month: LKR " & Format(monthlySales, "#,##0.00")
End Sub
