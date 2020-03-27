VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: | :: EDIT MODE (CUSTOMER) :: | :."
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   8280
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1560
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   3855
      Left            =   240
      TabIndex        =   25
      Top             =   360
      Width           =   7815
      Begin VB.TextBox txtPC 
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
         MaxLength       =   9
         TabIndex        =   7
         Text            =   "txtPC"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtArea 
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
         Height          =   315
         Left            =   5280
         MaxLength       =   45
         TabIndex        =   6
         Text            =   "txtArea"
         Top             =   1080
         Width           =   2295
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
         Height          =   315
         Left            =   5280
         MaxLength       =   11
         TabIndex        =   4
         Text            =   "txtMobile"
         Top             =   720
         Width           =   2295
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
         TabIndex        =   13
         Top             =   3360
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
         TabIndex        =   14
         Top             =   3360
         Width           =   1215
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
         TabIndex        =   15
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtR 
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
         Height          =   1035
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "frmCustomer.frx":08CA
         Top             =   1920
         Width           =   3615
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
         Height          =   315
         Left            =   1440
         MaxLength       =   11
         TabIndex        =   3
         Text            =   "txtTelephone"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtCID 
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
         Height          =   315
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "txtCID"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtDate 
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
         Height          =   315
         Left            =   5280
         TabIndex        =   2
         Text            =   "txtDate"
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox CName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         ItemData        =   "frmCustomer.frx":08CF
         Left            =   1440
         List            =   "frmCustomer.frx":08D1
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "Name"
         Top             =   1080
         Width           =   2295
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
         Top             =   3000
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
         TabIndex        =   27
         Top             =   3000
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
         TabIndex        =   16
         Top             =   3360
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
         TabIndex        =   26
         Top             =   3000
         Width           =   1215
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
         TabIndex        =   12
         Top             =   3000
         Width           =   1215
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
         TabIndex        =   9
         Top             =   3000
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
         TabIndex        =   10
         Top             =   3000
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
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Post Code"
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
         TabIndex        =   35
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label7 
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
         Left            =   4080
         TabIndex        =   34
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   4080
         TabIndex        =   33
         Top             =   360
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
         Left            =   4080
         TabIndex        =   32
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         TabIndex        =   31
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
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
         TabIndex        =   29
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblR 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         TabIndex        =   28
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
      Left            =   240
      TabIndex        =   22
      Top             =   4320
      Width           =   7815
      Begin VB.ComboBox ST 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         ItemData        =   "frmCustomer.frx":08D3
         Left            =   2640
         List            =   "frmCustomer.frx":08E3
         Sorted          =   -1  'True
         TabIndex        =   18
         Text            =   "Name"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "&Label"
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
         Left            =   6000
         TabIndex        =   37
         Top             =   360
         Width           =   735
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
         TabIndex        =   36
         Top             =   360
         Width           =   855
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
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   735
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
         TabIndex        =   17
         Text            =   "txtSearch"
         Top             =   360
         Width           =   2295
      End
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
         Left            =   6840
         TabIndex        =   20
         Top             =   360
         Width           =   735
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
         TabIndex        =   23
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
         TabIndex        =   24
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   5400
      Top             =   240
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
      Left            =   4200
      TabIndex        =   21
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
      Left            =   600
      Top             =   6960
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
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nLastKeyAscii As Integer
Dim Cus_ID As String

Private Sub cmdLabel_Click()
    'For Mouse
    Me.MousePointer = vbHourglass

    'Deleting Previous Records...
    DeleteCustomerRecords
    
    'Adding New Records...
    AddCustomerRecords
    
    'Showing Report...
        CrystalReport1.ReportFileName = App.Path & "\Reports\Rpt_Add.rpt"
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

Private Sub Form_Load()
    ClearFields
    ConnectCustomers
    Normalize
    GetComboData
    RemoveNameDuplicates
    
    txtSearch.Text = ""
    
End Sub

Private Sub cmdNew_Click()
    EnterNewCustomer
End Sub

Private Sub cmdAdd_Click()
'Checking Fields for Records
    If (txtCID.Text = "" Or txtCID.Text = " ") Then
        MsgBox "Enter Customer ID !!!", vbOKOnly, "Information Required"
        txtCID.SetFocus
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
    If (CName.Text = "Name" Or CName.Text = "" Or CName.Text = " ") Then
        MsgBox "Enter Customer Name !!!", vbOKOnly, "Information Required"
        CName.SetFocus
        Exit Sub
    End If
    If (txtArea.Text = "" Or txtArea.Text = " ") Then
        MsgBox "Enter area for " + CName.Text + " !!!", vbOKOnly, "Information Required"
        txtArea.SetFocus
        Exit Sub
    End If
    If (txtPC.Text = "" Or txtPC.Text = " ") Then
        MsgBox "Enter post code for " + CName.Text + " !!!", vbOKOnly, "Information Required"
        txtPC.SetFocus
        Exit Sub
    End If
    
    If (txtR.Text = "") Then txtR.Text = "-"
    If (txtMobile.Text = "") Then txtR.Text = "-"
    
    
    On Error GoTo AddError
    'Updating Database
    Adodc1.Recordset.AddNew

    Adodc1.Recordset.Fields("CID") = txtCID.Text
    Adodc1.Recordset.Fields("Date") = txtDate.Text
    Adodc1.Recordset.Fields("Telephone") = txtTelephone.Text
    Adodc1.Recordset.Fields("Mobile") = txtMobile.Text
    Adodc1.Recordset.Fields("Name") = CName.Text
    Adodc1.Recordset.Fields("Area") = txtArea.Text
    Adodc1.Recordset.Fields("Post_Code") = txtPC.Text
    Adodc1.Recordset.Fields("Remarks") = txtR.Text

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
    txtCID.SetFocus
    DisableButtons
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Visible = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    On Error GoTo SaveError
    
    Adodc1.Recordset.Update "CID", txtCID.Text
    Adodc1.Recordset.Update "Date", txtDate.Text
    Adodc1.Recordset.Update "Telephone", txtTelephone.Text
    Adodc1.Recordset.Update "Mobile", txtMobile.Text
    Adodc1.Recordset.Update "Name", CName.Text
    Adodc1.Recordset.Update "Area", txtArea.Text
    Adodc1.Recordset.Update "Post_Code", txtPC.Text
    Adodc1.Recordset.Update "Remarks", txtR.Text
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
        cmdRDB_Click
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
    Adodc1.RecordSource = "select * from Customer Order by Name;"
    Adodc1.Refresh
    If (Adodc1.Recordset.BOF) Then
        Exit Sub
    Else
        Adodc1.Recordset.MoveFirst
        GetCustomerData
    End If
End Sub

Private Sub cmdMF_Click()
If (Adodc1.Recordset.RecordCount = 0) Then
    MsgBox "NO RECORDS FOUND !!!", vbInformation, ":: | :: ADMIN :: | :."
    Exit Sub
End If
On Error GoTo MError
    Adodc1.Recordset.MoveFirst
    GetCustomerData
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
    GetCustomerData
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
        GetCustomerData
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
        GetCustomerData
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

Adodc1.RecordSource = "select * from Customer where " + ST.Text + " Like '" + txtSearch.Text + "%'"
Adodc1.Refresh

'Checking Either the Record is Present in Database or not
If Adodc1.Recordset.EOF Then
    MsgBox "Record for " + txtSearch.Text + " Not Found !!!", vbInformation, ":: | :: ADMIN :: | :."
    Adodc1.RecordSource = "select * from Customer Order by Name;"
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

Private Sub CName_Change()
   Select Case nLastKeyAscii
      Case vbKeyBack
         Call Combo_Lookup(CName)
      Case vbKeyDelete
      Case Else
         Call Combo_Lookup(CName)
   End Select
End Sub

Private Sub CName_KeyDown(KeyCode As Integer, Shift As Integer)
   nLastKeyAscii = KeyCode
   
   If KeyCode = vbKeyBack And Len(CName.SelText) <> 0 And CName.SelStart > 0 Then
         CName.SelStart = CName.SelStart - 1
         CName.SelLength = CB_MAXLENGTH
   End If
End Sub

'********************************************************************
Private Sub ClearFields()
    txtCID.Text = ""
    txtDate.Text = ""
    txtTelephone.Text = ""
    txtMobile.Text = ""
    CName.Text = ""
    txtArea.Text = ""
    txtPC.Text = ""
    txtR.Text = ""
End Sub

Private Sub EnableFields()
    txtCID.Enabled = True
    txtDate.Enabled = True
    txtTelephone.Enabled = True
    txtMobile.Enabled = True
    CName.Enabled = True
    txtArea.Enabled = True
    txtPC.Enabled = True
    txtR.Enabled = True
End Sub
Private Sub DisableFields()
    txtCID.Enabled = False
    txtDate.Enabled = False
    txtTelephone.Enabled = False
    txtMobile.Enabled = False
    CName.Enabled = False
    txtArea.Enabled = False
    txtPC.Enabled = False
    txtR.Enabled = False
End Sub

Private Sub EnableButtons()
    cmdNew.Enabled = True
    cmdPrint.Enabled = True
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = True
    cmdLabel.Enabled = True
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
    cmdAdd.Enabled = False
    cmdLabel.Enabled = False
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
    ST.Enabled = True
    cmdRDB_Click
End Sub

Private Sub GetComboData()
    Adodc2.ConnectionString = cn
    Adodc2.CursorLocation = adUseClient
    Adodc2.CursorType = adOpenDynamic
    Adodc2.RecordSource = "select CID, Name from Customer order by Name;"
    Set DataGrid2.DataSource = Adodc2
    
    If Adodc2.Recordset.BOF Then
        Exit Sub
    Else

    'For Item1 and Item Combo
        Dim X As Integer
        For X = 0 To (Adodc2.Recordset.RecordCount - 1)
            CName.AddItem Adodc2.Recordset.Fields(1)
            Adodc2.Recordset.MoveNext
        Next X
    End If
    
End Sub

Public Function RemoveNameDuplicates()
    Dim Y As Integer
    Dim X As Integer
    Y = CName.ListCount + 1
    For X = 1 To CName.ListCount
        Y = Y - 1
        If CName.List(Y) = CName.List(Y - 1) Then
            CName.RemoveItem (Y)
        End If
    Next
End Function

Private Sub cmdPrint_Click()
    'For Mouse
    Me.MousePointer = vbHourglass

    'Deleting Previous Records...
    DeleteCustomerRecords
    
    'Adding New Records...
    AddCustomerRecords
    
    'Showing Report...
        CrystalReport1.ReportFileName = App.Path & "\Reports\Rpt_Cust.rpt"
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
    Adodc3.Recordset.Fields("Form") = "frmCustomer"
    
    Adodc3.Recordset.Update
    Adodc3.Recordset.Requery
    frmMail.txtBody.Text = "Error on frmCustomer" + ErrorMsg + " USER: " + user
    frmMail.Show
    Exit Sub
End Sub
Private Sub DeleteCustomerRecords()
    Dim i As Integer
    Dim Query As String
    
    On Error Resume Next
            Adodc3.ConnectionString = cn
            Adodc3.CursorLocation = adUseClient
            Adodc3.CursorType = adOpenDynamic
            Query = "Select * from Rpt_Customer"
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
Private Sub AddCustomerRecords()
    Dim CD, DT, TL, MB, NM, AR, PCD, RM As String
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
                CD = Adodc1.Recordset.Fields(0)
                DT = Adodc1.Recordset.Fields(1)
                TL = Adodc1.Recordset.Fields(2)
                MB = Adodc1.Recordset.Fields(3)
                NM = Adodc1.Recordset.Fields(4)
                AR = Adodc1.Recordset.Fields(5)
                PCD = Adodc1.Recordset.Fields(6)
                RM = Adodc1.Recordset.Fields(7)

                Adodc3.Recordset.AddNew
                Adodc3.Recordset.Fields("CID") = CD
                Adodc3.Recordset.Fields("Date") = DT
                Adodc3.Recordset.Fields("Telephone") = TL
                Adodc3.Recordset.Fields("Mobile") = MB
                Adodc3.Recordset.Fields("Name") = NM
                Adodc3.Recordset.Fields("Area") = AR
                Adodc3.Recordset.Fields("Post_Code") = PCD
                Adodc3.Recordset.Fields("Remarks") = RM
                
                Adodc3.Recordset.Update
                Adodc3.Recordset.Requery
                
                Adodc1.Recordset.MoveNext
                
            End If
        Next i

End Sub

Public Sub EnterNewCustomer()
    EnableFields
    ClearFields
    txtCID.Text = Cus_ID
    txtTelephone.SetFocus
    DisableButtons
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdNew.Visible = False
    cmdDelete.Visible = False
    cmdCancel.Enabled = True
    cmdAdd.Enabled = True
    GenerateID
    txtDate.Text = Date
End Sub
Private Sub GenerateID()
    Adodc1.RecordSource = "SELECT * FROM Customer;"
    Adodc1.Refresh
    
    Cus_ID = Adodc1.Recordset.RecordCount + 1
    Cus_ID = "C" + Cus_ID
    txtCID.Text = Cus_ID
End Sub

Private Sub ST_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
