VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmChange 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: | :: CHANGE PASSWORD :: | :."
   ClientHeight    =   2280
   ClientLeft      =   5025
   ClientTop       =   4965
   ClientWidth     =   6240
   Icon            =   "frmChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6000
      Begin VB.CommandButton mnChange 
         Caption         =   "&CHANGE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   450
         TabIndex        =   3
         Top             =   1620
         Width           =   2280
      End
      Begin VB.TextBox txtNP2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1110
         Width           =   2655
      End
      Begin VB.TextBox txtNP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   690
         Width           =   2655
      End
      Begin VB.TextBox txtCP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   270
         Width           =   2655
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CA&NCEL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3150
         TabIndex        =   4
         Top             =   1620
         Width           =   2280
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2655
         Left            =   330
         TabIndex        =   6
         Top             =   2400
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   4683
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   180
         Top             =   2310
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Conform New Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1170
         Width           =   2595
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   270
         TabIndex        =   8
         Top             =   750
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   270
         TabIndex        =   7
         Top             =   330
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancel_LostFocus()
txtCP.SetFocus
End Sub

Private Sub Form_Load()
    'Connect
    Adodc1.ConnectionString = cn
    Adodc1.CursorLocation = adUseClient
    Adodc1.RecordSource = "select * from Login order by User"
    Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub mnChange_Click()
Adodc1.Refresh
On Error Resume Next
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "User='" + user + "'"

If (txtCP.Text = password) Then
    ChangePassword
Else
    MsgBox "Current Password Mismatch !!!", vbCritical, "Change Password"
    txtNP2.Text = ""
    txtNP.Text = ""
    txtCP.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If

End Sub

Public Sub ChangePassword()
If (txtNP.Text = txtNP2.Text) Then
    Adodc1.Recordset.Update "Password", txtNP.Text
    MsgBox "Your Password has been changed, Remember to login with new password next time...", vbInformation, "Conformation"
    txtNP.Text = ""
    txtNP2.Text = ""
    txtCP.Text = ""
    Unload Me
Else
    MsgBox "Conform new password does not match the new password!!!", vbInformation, "Change Password"
    txtNP.SetFocus
    SendKeys "{Home}+{End}"
    txtNP2.Text = ""
    Exit Sub
End If
End Sub

Private Sub txtCP_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtNP_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtNP2_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

