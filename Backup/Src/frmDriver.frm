VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmDriver 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: | :: DRIVERS :: | :."
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   Icon            =   "frmDriver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   8085
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "AMOUNT DEPOSIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   34
      Top             =   6480
      Width           =   7815
      Begin VB.TextBox txtD 
         Height          =   285
         Left            =   1200
         TabIndex        =   41
         Text            =   "txtD"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   255
         Left            =   4560
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   2040
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "Adodc4"
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
      Begin VB.ComboBox Drv_No 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmDriver.frx":08CA
         Left            =   1080
         List            =   "frmDriver.frx":08CC
         TabIndex        =   38
         Text            =   "Driver"
         ToolTipText     =   "How the Customer Will Pay"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdDep 
         Caption         =   "Deposit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtDeposit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   3480
         TabIndex        =   36
         Text            =   "txtDeposit"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   30
      Top             =   3720
      Width           =   7815
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   240
         TabIndex        =   19
         Text            =   "txtSearch"
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&FIND"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   5280
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "Adodc2"
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1335
         Left            =   4680
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2355
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1575
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2778
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   7815
      Begin VB.TextBox txtDCash 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   4440
         MaxLength       =   7
         TabIndex        =   7
         Text            =   "txtDCash"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdRDB 
         Caption         =   "Re&fresh DB"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   14
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   11
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdML 
         Caption         =   "Move &Last"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         TabIndex        =   18
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   4440
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "txtName"
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtDID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "txtDID"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtTelephone 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1440
         MaxLength       =   11
         TabIndex        =   3
         Text            =   "txtTelephone"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtAdd 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   795
         Left            =   4440
         MaxLength       =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmDriver.frx":08CE
         Top             =   840
         Width           =   3135
      End
      Begin VB.CommandButton cmdN 
         Caption         =   "Ne&xt"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdP 
         Caption         =   "&Previous"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   16
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdMF 
         Caption         =   "Move &First"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtMobile 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1440
         MaxLength       =   11
         TabIndex        =   5
         Text            =   "txtMobile"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtCash 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   6
         Text            =   "txtCash"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblR 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   3480
         TabIndex        =   28
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   3480
         TabIndex        =   27
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1440
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3960
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Adodc3"
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1335
      Left            =   3960
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2355
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   7080
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
End
Attribute VB_Name = "frmDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim D_id As String

Private Sub cmdDep_Click()
If (Drv_No.Text = "" Or Drv_No.Text = " " Or Drv_No.Text = "Driver") Then
    MsgBox "Please Select Driver !!!", vbInformation, ":: | :: ADMIN :: | :."
    Exit Sub
End If
If (txtDeposit.Text = "" Or txtDeposit.Text = " ") Then
    MsgBox "Please Enter the amount to be deducted from driver's account !!!", vbInformation, ":: | :: ADMIN :: | :."
    Exit Sub
Else
    Adodc1.ConnectionString = cn
    Adodc1.CursorLocation = adUseClient
    Adodc1.CursorType = adOpenDynamic
    Adodc1.RecordSource = "SELECT D_Cash FROM Driver WHERE Name='" + Drv_No.Text + "';"
    Adodc1.Refresh
    txtD.Text = Adodc1.Recordset.Fields(0)
    
    Adodc1.RecordSource = "SELECT * FROM Driver;"
    Adodc1.Refresh
    
    Adodc1.RecordSource = "SELECT * FROM Driver WHERE Name='" + Drv_No.Text + "';"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    txtD.Text = Val(txtD.Text) - Val(txtDeposit.Text)
    
    Adodc1.Recordset.Update "D_Cash", txtD.Text
    MsgBox "Driver Account Updated !!!", vbInformation, ":: | :: ADMIN :: | :."
    txtDeposit.Text = ""
    Drv_No.Text = "Driver"
    
    Normalize
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Form_Load()
    ClearFields
    ConnectDrivers
    Normalize
    txtSearch.Text = ""
    txtDeposit.Text = ""
    
    'Driver Combo
    Adodc4.ConnectionString = cn
    Adodc4.CursorLocation = adUseClient
    Adodc4.CursorType = adOpenDynamic
    Adodc4.RecordSource = "select DID,Name from Driver Order by Name;"
    Set DataGrid4.DataSource = Adodc4
    
    If (Adodc4.Recordset.BOF) Then
    Else
        Adodc4.Refresh
    End If
    
    GetDriverComboData
    RemoveDriverDuplicates
End Sub

Private Sub cmdNew_Click()
    EnterNewDriver
End Sub

Private Sub cmdAdd_Click()
'Checking Fields for Records
    If (txtDID.Text = "" Or txtDID.Text = " ") Then
        MsgBox "Enter Driver ID !!!", vbOKOnly, "Information Required"
        txtDID.SetFocus
        Exit Sub
    End If
    If (txtName.Text = "" Or txtName.Text = " ") Then
        MsgBox "Enter Driver Name !!!", vbOKOnly, "Information Required"
        txtName.SetFocus
        Exit Sub
    End If
    If (txtAdd.Text = "" Or txtAdd.Text = " ") Then
        MsgBox "Enter Driver Address !!!", vbOKOnly, "Information Required"
        txtAdd.SetFocus
        Exit Sub
    End If
    If (txtTelephone.Text = "" Or txtTelephone.Text = " ") Then
        MsgBox "Enter Telephone Number for " + CName.Text + " !!!", vbOKOnly, "Information Required"
        txtTelephone.SetFocus
        Exit Sub
    End If
    If (txtMobile.Text = "" Or txtMobile.Text = " ") Then
        txtMobile.Text = "-"
    End If
    
    If (txtCash.Text = "" Or txtCash.Text = " ") Then
        txtCash.Text = "0"
    End If
    If (txtDCash.Text = "" Or txtDCash.Text = " ") Then
        txtDCash.Text = "0"
    End If
        
    On Error GoTo AddError
    'Updating Database
    Adodc1.Recordset.AddNew

    Adodc1.Recordset.Fields("DID") = txtDID.Text
    Adodc1.Recordset.Fields("Name") = txtName.Text
    Adodc1.Recordset.Fields("Address") = txtAdd.Text
    Adodc1.Recordset.Fields("Telephone") = txtTelephone.Text
    Adodc1.Recordset.Fields("Mobile") = txtMobile.Text
    Adodc1.Recordset.Fields("Cash") = txtCash.Text
    Adodc1.Recordset.Fields("D_Cash") = txtDCash.Text

    Adodc1.Recordset.Update
    Adodc1.Recordset.Requery
    MsgBox "Record Added in Database Sucessfully !!!", vbInformation, "Conformation"
        
    Normalize
    cmdNew.SetFocus
    Exit Sub
    
AddError:
    ErrorMsg = "Add Error: " + Err.Description
    MsgBox "Add Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
End Sub

Private Sub cmdEdit_Click()
    EnableFields
    txtDID.SetFocus
    DisableButtons
    txtSearch.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Visible = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    On Error GoTo SaveError
    
    Adodc1.Recordset.Update "DID", txtDID.Text
    Adodc1.Recordset.Update "Name", txtName.Text
    Adodc1.Recordset.Update "Address", txtAdd.Text
    Adodc1.Recordset.Update "Telephone", txtTelephone.Text
    Adodc1.Recordset.Update "Mobile", txtMobile.Text
    Adodc1.Recordset.Update "Cash", txtCash.Text
    Adodc1.Recordset.Update "D_Cash", txtDCash.Text
    cmdRDB_Click
    Normalize
    Exit Sub

SaveError:
    ErrorMsg = "Save Error: " + Err.Description
    MsgBox "Save Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
End Sub

Private Sub cmdCancel_Click()
    Normalize
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DelError
    If (Adodc1.Recordset.BOF) Then
        Exit Sub
    Else
    
        Adodc1.Recordset.Delete
        Normalize
        
        MsgBox "Record Deleted !!!", vbInformation, ""
        ClearFields
        cmdRDB_Click
        Exit Sub
    End If

DelError:
    ErrorMsg = "Delete Error: " + Err.Description
    MsgBox "Delete Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdRDB_Click()
    Adodc1.RecordSource = "select * from Driver Order by Name;"
    Adodc1.Refresh
    If (Adodc1.Recordset.BOF) Then
    Else
        Adodc1.Recordset.MoveFirst
        GetDriverData
    End If
    GetDriverComboData
    RemoveDriverDuplicates
End Sub

Private Sub cmdMF_Click()
If (Adodc1.Recordset.RecordCount = 0) Then
    MsgBox "NO RECORDS FOUND !!!", vbInformation, ":: | :: ADMIN :: | :."
    Exit Sub
End If
On Error GoTo MError
    Adodc1.Recordset.MoveFirst
    GetDriverData
    Exit Sub
    
MError:
    ErrorMsg = "Movement Error: " + Err.Description
    MsgBox "Movement Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
End Sub

Private Sub cmdML_Click()
If (Adodc1.Recordset.RecordCount = 0) Then
    MsgBox "NO RECORDS FOUND !!!", vbInformation, ":: | :: ADMIN :: | :."
    Exit Sub
End If
On Error GoTo MError
    Adodc1.Recordset.MoveLast
    GetDriverData
    Exit Sub
    
MError:
    ErrorMsg = "Movement Error: " + Err.Description
    MsgBox "Movement Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
End Sub

Private Sub cmdN_Click()
If (Adodc1.Recordset.RecordCount = 0) Then
    MsgBox "NO RECORDS FOUND !!!", vbInformation, ":: | :: ADMIN :: | :."
    Exit Sub
End If
On Error GoTo MError
    Adodc1.Recordset.MoveNext
    
    If (Adodc1.Recordset.EOF) Then
        MsgBox "This is Last Record !!!", vbInformation, ":: | :: ADMIN :: | :."
        Adodc1.Recordset.MoveLast
    Else
        GetDriverData
    End If
Exit Sub
    
MError:
    ErrorMsg = "Movement Error: " + Err.Description
    MsgBox "Movement Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
End Sub

Private Sub cmdP_Click()
If (Adodc1.Recordset.RecordCount = 0) Then
    MsgBox "NO RECORDS FOUND !!!", vbInformation, ":: | :: ADMIN :: | :."
    Exit Sub
End If
On Error GoTo MError
    Adodc1.Recordset.MovePrevious
    
    If (Adodc1.Recordset.BOF) Then
        MsgBox "This is First Record !!!", vbInformation, ":: | :: ADMIN :: | :."
        Adodc1.Recordset.MoveFirst
    Else
        GetDriverData
    End If
Exit Sub
    
MError:
    ErrorMsg = "Movement Error: " + Err.Description
    MsgBox "Movement Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
End Sub

Private Sub cmdSearch_Click()
If (txtSearch.Text = "" Or txtSearch.Text = " ") Then
    MsgBox "Search what?", vbExclamation, ":: | :: ADMIN :: | :."
    txtSearch.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If

On Error GoTo SError

Adodc1.RecordSource = "select * from Driver where Name Like '" + txtSearch.Text + "%';"
Adodc1.Refresh

'Checking Either the Record is Present in Database or not
If Adodc1.Recordset.EOF Then
    MsgBox "Record for " + txtSearch.Text + " Not Found !!!", vbInformation, ":: | :: ADMIN :: | :."
    Adodc1.RecordSource = "select * from Driver Order by Name;"
    Adodc1.Refresh
    Exit Sub

Else

    'Getting Data in Text Fields
    On Error GoTo SError
    Adodc1.Recordset.MoveFirst
    
    Exit Sub
    
End If

SError:
    ErrorMsg = "Search Error: " + Err.Description
    MsgBox "Search Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError

End Sub

'********************************************************************
Private Sub ClearFields()
    txtDID.Text = ""
    txtName.Text = ""
    txtAdd.Text = ""
    txtTelephone.Text = ""
    txtMobile.Text = ""
    txtCash.Text = ""
    txtDCash.Text = ""
End Sub

Private Sub EnableFields()
    txtDID.Enabled = True
    txtName.Enabled = True
    txtAdd.Enabled = True
    txtTelephone.Enabled = True
    txtMobile.Enabled = True
    txtCash.Enabled = True
    txtDCash.Enabled = True
End Sub
Private Sub DisableFields()
    txtDID.Enabled = False
    txtName.Enabled = False
    txtAdd.Enabled = False
    txtTelephone.Enabled = False
    txtMobile.Enabled = False
    txtCash.Enabled = False
    txtDCash.Enabled = False
End Sub

Private Sub EnableButtons()
    cmdNew.Enabled = True
    cmdPrint.Enabled = True
    cmdDep.Enabled = True
    Command1.Enabled = True
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = True
    cmdRDB.Enabled = True
    cmdMF.Enabled = True
    cmdN.Enabled = True
    cmdP.Enabled = True
    cmdML.Enabled = True
    cmdSearch.Enabled = True
    cmdClose.Enabled = True
End Sub
Private Sub DisableButtons()
    cmdNew.Enabled = False
    cmdPrint.Enabled = False
    cmdDep.Enabled = False
    Command1.Enabled = False
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    cmdRDB.Enabled = False
    cmdMF.Enabled = False
    cmdN.Enabled = False
    cmdP.Enabled = False
    cmdML.Enabled = False
    cmdSearch.Enabled = False
    cmdClose.Enabled = False
End Sub

Private Sub Normalize()
    DisableFields
    EnableButtons
    cmdNew.Visible = True
    cmdEdit.Visible = True
    cmdDelete.Visible = True
    'cmdNew.SetFocus
    txtSearch.Enabled = True
    txtDeposit.Enabled = True
    cmdRDB_Click
End Sub

Private Sub cmdPrint_Click()
    'For Mouse
    Me.MousePointer = vbHourglass

    'Deleting Previous Records...
    DeleteDriverRecords
    
    'Adding New Records...
    AddDriverRecords
    
    'Showing Report...
        CrystalReport1.ReportFileName = App.Path & "\Reports\Rpt_Drv.rpt"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.WindowShowPrintSetupBtn = True
        CrystalReport1.WindowMaxButton = True
        CrystalReport1.WindowShowCloseBtn = True
        CrystalReport1.WindowShowRefreshBtn = True
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.Action = 1
    
    'For Mouse
        MousePointer = Default
        Exit Sub
End Sub
Private Sub RecordError()
'Connecting Database with ADODC1
    Adodc3.ConnectionString = cn
    Adodc3.CursorLocation = adUseClient
    Adodc3.CursorType = adOpenDynamic
    Adodc3.RecordSource = "select * from Error_Log;"
    Set DataGrid3.DataSource = Adodc3

    Dim Timing As String
    Timing = Now
    
    Adodc3.Recordset.AddNew
    
    Adodc3.Recordset.Fields("Timing") = Timing
    Adodc3.Recordset.Fields("Error") = ErrorMsg
    Adodc3.Recordset.Fields("User") = user
    Adodc3.Recordset.Fields("Form") = "frmDriver"
    
    Adodc3.Recordset.Update
    Adodc3.Recordset.Requery
    frmMail.txtBody.Text = "Error on frmDriver" + ErrorMsg + " USER: " + user
    frmMail.Show
    Exit Sub
End Sub
Private Sub DeleteDriverRecords()
    Dim i As Integer
    Dim Query As String
    
    On Error Resume Next
            Adodc3.ConnectionString = cn
            Adodc3.CursorLocation = adUseClient
            Adodc3.CursorType = adOpenDynamic
            Query = "Select * from Rpt_Drv"
            Adodc3.RecordSource = Query
            Adodc3.Refresh
            Set DataGrid3.DataSource = Adodc3
    
    If Adodc3.Recordset.BOF Then
        Exit Sub
    Else
        Adodc3.Recordset.MoveFirst
        For i = 1 To Adodc3.Recordset.RecordCount
            Adodc3.Recordset.Delete
            Adodc3.Recordset.Requery
            Adodc3.Refresh
        Next i
    End If
End Sub
Private Sub AddDriverRecords()
    Dim d, n, a, t, m, C, DC As String
    Dim i As Integer
    
    Adodc1.Refresh
    'Adodc1.Recordset.MoveFirst
        For i = 0 To Adodc1.Recordset.RecordCount - 1
            If (Adodc1.Recordset.EOF) Then
    
                Exit Sub
                'For Mouse
                    MousePointer = Default
                    Exit Sub
            
            Else
                d = Adodc1.Recordset.Fields(0)
                n = Adodc1.Recordset.Fields(1)
                a = Adodc1.Recordset.Fields(2)
                t = Adodc1.Recordset.Fields(3)
                m = Adodc1.Recordset.Fields(4)
                C = Adodc1.Recordset.Fields(5)
                DC = Adodc1.Recordset.Fields(6)

                Adodc3.Recordset.AddNew
                Adodc3.Recordset.Fields("DID") = d
                Adodc3.Recordset.Fields("Name") = n
                Adodc3.Recordset.Fields("Address") = a
                Adodc3.Recordset.Fields("Telephone") = t
                Adodc3.Recordset.Fields("Mobile") = m
                Adodc3.Recordset.Fields("Cash") = C
                Adodc3.Recordset.Fields("D_Cash") = DC
                
                Adodc3.Recordset.Update
                Adodc3.Recordset.Requery
                
                Adodc1.Recordset.MoveNext
                
            End If
        Next i

End Sub

Public Sub EnterNewDriver()
    EnableFields
    ClearFields
    
    DisableButtons
    txtSearch.Enabled = False
    txtDeposit.Enabled = False
    cmdNew.Visible = False
    cmdDelete.Visible = False
    cmdCancel.Enabled = True
    cmdAdd.Enabled = True
    GenerateID
    txtName.SetFocus
    
End Sub
Private Sub GenerateID()
    Adodc1.RecordSource = "SELECT * FROM Driver;"
    Adodc1.Refresh
    
    D_id = Adodc1.Recordset.RecordCount + 1
    D_id = "D" + D_id
    txtDID.Text = D_id
End Sub

Private Sub GetDriverComboData()
    Adodc4.ConnectionString = cn
    Adodc4.CursorLocation = adUseClient
    Adodc4.CursorType = adOpenDynamic
    Adodc4.RecordSource = "select DID,Name from Driver;"
    Set DataGrid4.DataSource = Adodc4
    If Adodc4.Recordset.BOF Then
        Exit Sub
    Else
        Adodc4.Refresh
    End If

'Removing Data First
    Dim a As Integer
    While Drv_No.ListCount <> 0
        a = a + 1
        Drv_No.RemoveItem (0)
    Wend

'For Item1 and Item Combo
    Dim X As Integer
    For X = 0 To (Adodc4.Recordset.RecordCount - 1)
    
        Drv_No.AddItem Adodc4.Recordset.Fields(1)
    
        If Adodc4.Recordset.EOF Then
            Exit Sub
        Else
            Adodc4.Recordset.MoveNext
        End If
    Next X
End Sub
Public Function RemoveDriverDuplicates()
    Dim Y As Integer
    Dim X As Integer
    Y = Drv_No.ListCount + 1
    For X = 1 To Drv_No.ListCount
        Y = Y - 1
        If Drv_No.List(Y) = Drv_No.List(Y - 1) Then
            Drv_No.RemoveItem (Y)
        End If
    Next
End Function
