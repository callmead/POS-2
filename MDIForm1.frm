VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000000&
   Caption         =   ":: | :: NOORI ORDER PROCESSING SYSTEM :: | :."
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9960
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":08CA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   8943
            Text            =   "Software Version 1.0.0"
            TextSave        =   "Software Version 1.0.0"
            Object.ToolTipText     =   "Contact me at adeel_s90@hotmail.com for details and updates."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "11:19"
            Object.ToolTipText     =   "SystemTime"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "23/04/2005"
            Object.ToolTipText     =   "System Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnEdit 
         Caption         =   "&Items"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnCustomers 
         Caption         =   "&Customers"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnDrivers 
         Caption         =   "Drivers"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnOrder 
         Caption         =   "&Orders"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnInvoices 
         Caption         =   "Invoices"
         Shortcut        =   ^I
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnClose 
         Caption         =   "&Close"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnSepL 
         Caption         =   "-"
      End
      Begin VB.Menu mnLogOff 
         Caption         =   "Log Off"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnItemWise 
         Caption         =   "Items Wise"
      End
      Begin VB.Menu mnCustomer 
         Caption         =   "Customer Wise"
      End
      Begin VB.Menu SepRep 
         Caption         =   "-"
      End
      Begin VB.Menu mnSales 
         Caption         =   "Sales Wise"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnOthers 
      Caption         =   "O&thers"
      Begin VB.Menu mnDuas 
         Caption         =   "Duas"
      End
      Begin VB.Menu SepOth 
         Caption         =   "-"
      End
      Begin VB.Menu mnMail 
         Caption         =   "Mail"
      End
   End
   Begin VB.Menu mnOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnChangePwd 
         Caption         =   "Change Password"
      End
      Begin VB.Menu SepOpt 
         Caption         =   "-"
      End
      Begin VB.Menu mnUM 
         Caption         =   "&User Management"
      End
      Begin VB.Menu mnUsage 
         Caption         =   "&View Usage"
      End
   End
   Begin VB.Menu mnWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnArrange 
         Caption         =   "Arrange"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnCascade 
         Caption         =   "Cascade"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnHorizontal 
         Caption         =   "Horizontal"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnVertical 
         Caption         =   "Vertical"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnAbout 
      Caption         =   "&About"
      Begin VB.Menu mnDeveloper 
         Caption         =   "Developer"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnSoftware 
         Caption         =   "Software"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EndMe As Boolean

Private Sub MDIForm_Load()
EndMe = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Secure
    If (EndMe = True) Then
        End
    Else
    End If
End Sub

Private Sub mnArrange_Click()
MDIForm1.Arrange vbArrangeIcon
End Sub

Private Sub mnCascade_Click()
MDIForm1.Arrange vbCascade
End Sub

Private Sub mnChangePwd_Click()
frmChange.Show vbModal
End Sub

Private Sub mnClose_Click()
If ActiveForm Is Nothing Then Exit Sub
Unload ActiveForm
End Sub

Private Sub mnCustomer_Click()
frmCustomer.Show
End Sub

Private Sub mnCustomers_Click()
frmCustomer.Show
End Sub

Private Sub mnDeveloper_Click()
frmAbout2.Show vbModal
End Sub

Private Sub mnDrivers_Click()
frmDriver.Show
End Sub

Private Sub mnDuas_Click()
Shell App.Path + "\DUAS.exe", vbNormalFocus
End Sub

Private Sub mnEdit_Click()
frmItem.Show
End Sub

Private Sub mnExit_Click()
    Secure
    EndMe = True
    End
End Sub

Private Sub mnHorizontal_Click()
MDIForm1.Arrange vbTileHorizontal
End Sub

Private Sub mnInvoices_Click()
frmInvoice.Show
End Sub

Private Sub mnLabels_Click()
frmTest.Show
End Sub

Private Sub mnItemWise_Click()
frmItem.Show
End Sub

Private Sub mnLogOff_Click()
EndMe = False
Unload frmLogin
Unload Me
    frmLogin.Show
End Sub

Private Sub mnMail_Click()
frmMail.Show
End Sub

Private Sub mnOrder_Click()
frmOrder.Show
End Sub

Private Sub mnSales_Click()
frmSales.Show
End Sub

Private Sub mnSoftware_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnUM_Click()
frmSecurity.Show
End Sub

Private Sub mnUsage_Click()
    frmVU.Show
End Sub

Private Sub mnVertical_Click()
MDIForm1.Arrange vbVertical
End Sub
