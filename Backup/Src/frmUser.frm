VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSecurity 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: | :: USER MANAGEMENT :: | :."
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   9600
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   2505
      Left            =   0
      TabIndex        =   18
      Top             =   3105
      Width           =   9615
      Begin VB.TextBox txtUser 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "txtUser"
         Top             =   270
         Width           =   2385
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "txtPass"
         Top             =   690
         Width           =   2385
      End
      Begin VB.ComboBox UType 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmUser.frx":08CA
         Left            =   1590
         List            =   "frmUser.frx":08D4
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "User Type"
         Top             =   1110
         Width           =   2385
      End
      Begin VB.TextBox txtFN 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5820
         MaxLength       =   45
         TabIndex        =   2
         Text            =   "txtFN"
         Top             =   270
         Width           =   3525
      End
      Begin VB.TextBox txtAdd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         MaxLength       =   45
         TabIndex        =   7
         Text            =   "txtAdd"
         Top             =   1530
         Width           =   7755
      End
      Begin VB.TextBox txtDesg 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5820
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "txtDesg"
         Top             =   690
         Width           =   3525
      End
      Begin VB.TextBox txtCnt 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5820
         MaxLength       =   11
         TabIndex        =   6
         Text            =   "txtCnt"
         Top             =   1110
         Width           =   3525
      End
      Begin VB.TextBox txtR 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         MaxLength       =   45
         TabIndex        =   8
         Text            =   "txtR"
         Top             =   1950
         Width           =   7755
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   210
         TabIndex        =   26
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   210
         TabIndex        =   25
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label lblType 
         BackStyle       =   0  'Transparent
         Caption         =   "User Type"
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
         Left            =   210
         TabIndex        =   24
         Top             =   1170
         Width           =   1245
      End
      Begin VB.Label lblFN 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
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
         Left            =   4440
         TabIndex        =   23
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label lblAdd 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   210
         TabIndex        =   22
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label lblDesg 
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
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
         Left            =   4440
         TabIndex        =   21
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label lblCnt 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
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
         Left            =   4440
         TabIndex        =   20
         Top             =   1140
         Width           =   1245
      End
      Begin VB.Label lblR 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   210
         TabIndex        =   19
         Top             =   1980
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last Rec."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   1245
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5910
      Width           =   1245
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   1245
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First Rec."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   1245
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5910
      Width           =   1245
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5910
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5910
      Width           =   1245
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5910
      Width           =   1245
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   5370
      Visible         =   0   'False
      Width           =   9585
      _ExtentX        =   16907
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      DefColWidth     =   73
      Enabled         =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
            LCID            =   2057
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
            LCID            =   2057
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
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_Click()
On Error Resume Next
With DataGrid1
    .Col = 0
    txtUser.Text = .Text
    .Col = 1
    txtPass.Text = .Text
    .Col = 2
    UType.Text = .Text
    .Col = 3
    txtFN.Text = .Text
    .Col = 4
    txtAdd.Text = .Text
    .Col = 5
    txtDesg.Text = .Text
    .Col = 6
    txtCnt.Text = .Text
    .Col = 7
    txtR.Text = .Text
End With

End Sub

Private Sub Form_Load()
'Connect
    Adodc1.ConnectionString = cn
    Adodc1.CursorLocation = adUseClient
    Adodc1.RecordSource = "select * from Login order by User"
    Set DataGrid1.DataSource = Adodc1
    
    ClearText
    SetData
    cmdSave.Enabled = False
End Sub

Private Sub UType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdNew_Click()
'Clearing Text Boxes
    ClearText
    
    txtUser.SetFocus
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    cmdDel.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    UType.Text = "Select Type"
End Sub

Private Sub cmdSave_Click()
'Check
    If txtUser.Text = "" Then
    MsgBox "Enter User Name"
    txtUser.SetFocus
    Exit Sub

    ElseIf txtPass.Text = "" Then
    MsgBox "Enter User Password"
    txtPass.SetFocus
    Exit Sub
    
    ElseIf txtFN.Text = "" Then
    MsgBox "Enter User Full Name"
    txtFN.SetFocus
    Exit Sub
    
    ElseIf txtDesg.Text = "" Then
    MsgBox "Enter User Designation"
    txtDesg.SetFocus
    Exit Sub
        
    ElseIf UType.Text = "Select Type" Then
    MsgBox "Select User Type"
    UType.SetFocus
    Exit Sub
        
    ElseIf txtCnt.Text = "" Then
    MsgBox "Enter User Contact"
    txtCnt.SetFocus
    Exit Sub
    
    ElseIf txtAdd.Text = "" Then
    MsgBox "Enter User Address"
    txtAdd.SetFocus
    Exit Sub
    
    ElseIf txtR.Text = "" Then
    MsgBox "Please enter some remarks"
    txtR.SetFocus
    Exit Sub
    
    Else
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields("User") = txtUser.Text
    Adodc1.Recordset.Fields("Password") = txtPass.Text
    Adodc1.Recordset.Fields("Type") = UType.Text
    Adodc1.Recordset.Fields("Name") = txtFN.Text
    Adodc1.Recordset.Fields("Address") = txtAdd.Text
    Adodc1.Recordset.Fields("Designation") = txtDesg.Text
    Adodc1.Recordset.Fields("Contact") = txtCnt.Text
    Adodc1.Recordset.Fields("Remarks") = txtR.Text
        
    Adodc1.Recordset.Update
    Adodc1.Recordset.Requery
    MsgBox "User Database Updated !!!", vbInformation
    
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdDel.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdNew.SetFocus
    End If
End Sub

Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo SQLError
Adodc1.Recordset.MovePrevious
If (Adodc1.Recordset.BOF) Then
    MsgBox "This Is The First Record !!!", vbInformation, "Information"
    Adodc1.Recordset.MoveFirst
Else
    SetData
End If
Exit Sub

SQLError:
MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
Adodc1.RecordSource = "select * from Login order by User"
Adodc1.Refresh

    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdDel.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    SetData
'For CD Records Count
    'lblRecords4.Caption = (Adodc4.Recordset.RecordCount)
End Sub

Private Sub cmdNext_Click()
On Error GoTo SQLError
Adodc1.Recordset.MoveNext
If (Adodc1.Recordset.EOF) Then
    MsgBox "This Is The Last Record !!!", vbInformation, "Information"
    Adodc1.Recordset.MoveLast
Else
    SetData
End If
Exit Sub

SQLError:
MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
    SetData
End Sub

Private Sub cmdEdit_Click()
On Error GoTo SQLError
    Adodc1.Recordset.Update "Address", txtAdd.Text
    Adodc1.Recordset.Update "Contact", txtCnt.Text
    Adodc1.Recordset.Update "Designation", txtDesg.Text
    Adodc1.Recordset.Update "Type", UType.Text
    Adodc1.Recordset.Update "Name", txtFN.Text
    Adodc1.Recordset.Update "Password", txtPass.Text
    Adodc1.Recordset.Update "Remarks", txtR.Text
    Adodc1.Recordset.Update "User", txtUser.Text
    
    MsgBox "Record Updated Sucessfully!", vbInformation
    Exit Sub
SQLError:
MsgBox Err.Description
End Sub

Private Sub cmdDel_Click()
On Error GoTo SQLError
    If Adodc1.Recordset.BOF Then
        Exit Sub
    Else
        
        Adodc1.Recordset.Delete
        MsgBox "Record Deleted !!!", vbInformation
        Adodc1.RecordSource = "select * from Login order by User"
        Adodc1.Refresh
        ClearText
        SetData
    End If
    Exit Sub
SQLError:
MsgBox Err.Description
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub ClearText()
'Clear TextBoxes
    txtUser.Text = ""
    txtPass.Text = ""
    txtCnt.Text = ""
    txtFN.Text = ""
    txtDesg.Text = ""
    txtAdd.Text = ""
    txtR.Text = ""
End Sub

Public Sub SetData()
    'Put data in Text Fields from Adodc
    txtAdd.Text = Adodc1.Recordset.Fields("Address")
    txtCnt.Text = Adodc1.Recordset.Fields("Contact")
    txtDesg.Text = Adodc1.Recordset.Fields("Designation")
    txtFN.Text = Adodc1.Recordset.Fields("Name")
    txtPass.Text = Adodc1.Recordset.Fields("Password")
    txtR.Text = Adodc1.Recordset.Fields("Remarks")
    txtUser.Text = Adodc1.Recordset.Fields("User")
    UType.Text = Adodc1.Recordset.Fields("Type")
End Sub


