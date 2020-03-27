VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmItem 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: | :: EDIT MODE (ITEMS) :: | :."
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "frmItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   8280
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3240
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4080
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   6360
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
      TabIndex        =   26
      Top             =   3720
      Width           =   7815
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
         TabIndex        =   18
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
         TabIndex        =   29
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
         TabIndex        =   19
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
         TabIndex        =   16
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
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1575
         Left            =   240
         TabIndex        =   27
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
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   7815
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
         TabIndex        =   11
         Top             =   2520
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
         TabIndex        =   8
         Top             =   2520
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
         TabIndex        =   15
         Top             =   2880
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
         TabIndex        =   10
         Top             =   2520
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
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox Item 
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
         ItemData        =   "frmItem.frx":08CA
         Left            =   1440
         List            =   "frmItem.frx":08CC
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "Item"
         Top             =   840
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
         TabIndex        =   6
         Text            =   "txtDate"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtIC 
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
         Text            =   "txtIC"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtUP 
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
         Height          =   315
         Left            =   5280
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "txtUP"
         Top             =   840
         Width           =   2295
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
         MaxLength       =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmItem.frx":08CE
         Top             =   1320
         Width           =   3615
      End
      Begin VB.ComboBox IType 
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
         ItemData        =   "frmItem.frx":08D3
         Left            =   5280
         List            =   "frmItem.frx":08D5
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "Type"
         Top             =   360
         Width           =   2295
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
         TabIndex        =   14
         Top             =   2880
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
         TabIndex        =   13
         Top             =   2880
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
         TabIndex        =   12
         Top             =   2880
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
         TabIndex        =   9
         Top             =   2520
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
         TabIndex        =   7
         Top             =   2520
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
         TabIndex        =   28
         Top             =   2520
         Width           =   1215
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
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
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
         TabIndex        =   24
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         TabIndex        =   22
         Top             =   360
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
         TabIndex        =   21
         Top             =   1320
         Width           =   1815
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1335
      Left            =   4080
      TabIndex        =   30
      Top             =   120
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
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nLastKeyAscii As Integer

Private Sub Form_Load()
    ClearFields
    ConnectItems
    Normalize
    GetComboData
    GetItemData
    
    RemoveITypeDuplicates
    RemoveItemDuplicates
    
    txtSearch.Text = ""
    
End Sub

Private Sub cmdNew_Click()
    EnableFields
    ClearFields
    txtIC.SetFocus
    DisableButtons
    txtSearch.Enabled = False
    cmdNew.Visible = False
    cmdDelete.Visible = False
    cmdCancel.Enabled = True
    cmdAdd.Enabled = True
    
    txtDate.Text = Date
End Sub

Private Sub cmdAdd_Click()
'Checking Fields for Records
    If (txtIC.Text = "" Or txtIC.Text = " ") Then
        MsgBox "Enter Item Code !!!", vbOKOnly, "Information Required"
        txtIC.SetFocus
        Exit Sub
    End If
    If (IType.Text = "Type" Or IType.Text = "" Or IType.Text = " ") Then
        MsgBox "Enter Item Group !!!", vbOKOnly, "Information Required"
        IType.SetFocus
        Exit Sub
    End If
    If (Item.Text = "Item" Or Item.Text = "" Or Item.Text = " ") Then
        MsgBox "Enter Item !!!", vbOKOnly, "Information Required"
        Item.SetFocus
        Exit Sub
    End If
    If (txtUP.Text = "" Or txtUP.Text = " ") Then
        MsgBox "Enter price per Kg for " + Item.Text + " !!!", vbOKOnly, "Information Required"
        txtUP.SetFocus
        Exit Sub
    End If
    
    If (txtR.Text = "") Then txtR.Text = "-"
    
    
    On Error GoTo AddError
    'Updating Database
    Adodc1.Recordset.AddNew

    Adodc1.Recordset.Fields("Item_Code") = txtIC.Text
    Adodc1.Recordset.Fields("Grp") = IType.Text
    Adodc1.Recordset.Fields("Item") = Item.Text
    Adodc1.Recordset.Fields("Price") = txtUP.Text
    Adodc1.Recordset.Fields("Remarks") = txtR.Text
    Adodc1.Recordset.Fields("Date") = txtDate.Text

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
    txtIC.SetFocus
    DisableButtons
    txtSearch.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Visible = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    On Error GoTo SaveError
    
    Adodc1.Recordset.Update "Item_Code", txtIC.Text
    Adodc1.Recordset.Update "Grp", IType.Text
    Adodc1.Recordset.Update "Item", Item.Text
    Adodc1.Recordset.Update "Price", txtUP.Text
    Adodc1.Recordset.Update "Remarks", txtR.Text
    Adodc1.Recordset.Update "Date", txtDate.Text
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
        Adodc1.Refresh
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
    Adodc1.RecordSource = "select * from Item Order by Item_Code;"
    Adodc1.Refresh
    
    If (Adodc1.Recordset.BOF) Then
        Exit Sub
    Else
        Adodc1.Recordset.MoveFirst
        GetItemData
    End If
End Sub

Private Sub cmdMF_Click()
If (Adodc1.Recordset.RecordCount = 0) Then
    MsgBox "NO RECORDS FOUND !!!", vbInformation, ":: | :: ADMIN :: | :."
    Exit Sub
End If
On Error GoTo MError
    Adodc1.Recordset.MoveFirst
    GetItemData
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
    GetItemData
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
        GetItemData
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
        GetItemData
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

Adodc1.RecordSource = "select * from Item where Item Like '" + txtSearch.Text + "%';"
Adodc1.Refresh

'Checking Either the Record is Present in Database or not
If Adodc1.Recordset.EOF Then
    MsgBox "Record for " + txtSearch.Text + " Not Found !!!", vbInformation, ":: | :: ADMIN :: | :."
    Adodc1.RecordSource = "select * from Item Order by Item_Code;"
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

Private Sub Item_Change()
   Select Case nLastKeyAscii
      Case vbKeyBack
         Call Combo_Lookup(Item)
      Case vbKeyDelete
      Case Else
         Call Combo_Lookup(Item)
   End Select
End Sub

Private Sub Item_KeyDown(KeyCode As Integer, Shift As Integer)
   nLastKeyAscii = KeyCode
   
   If KeyCode = vbKeyBack And Len(Item.SelText) <> 0 And Item.SelStart > 0 Then
         Item.SelStart = Item.SelStart - 1
         Item.SelLength = CB_MAXLENGTH
   End If
End Sub

Private Sub IType_Change()
   Select Case nLastKeyAscii
      Case vbKeyBack
         Call Combo_Lookup(IType)
      Case vbKeyDelete
      Case Else
         Call Combo_Lookup(IType)
   End Select
End Sub

Private Sub IType_KeyDown(KeyCode As Integer, Shift As Integer)
   nLastKeyAscii = KeyCode
   
   If KeyCode = vbKeyBack And Len(IType.SelText) <> 0 And IType.SelStart > 0 Then
         IType.SelStart = IType.SelStart - 1
         IType.SelLength = CB_MAXLENGTH
   End If
End Sub

Private Sub cmdPrint_Click()
    'For Mouse
    Me.MousePointer = vbHourglass

    'Deleting Previous Records...
    DeleteItemRecords
    
    'Adding New Records...
    AddItemRecords
    
    'Showing Report...
        CrystalReport1.ReportFileName = App.Path & "\Reports\Rpt_Items.rpt"
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

'********************************************************************
Private Sub ClearFields()
    txtIC.Text = ""
    IType.Text = ""
    Item.Text = ""
    txtUP.Text = ""
    txtR.Text = ""
    txtDate.Text = ""
    txtSearch.Text = ""
End Sub

Private Sub EnableFields()
    txtIC.Enabled = True
    IType.Enabled = True
    Item.Enabled = True
    txtUP.Enabled = True
    txtR.Enabled = True
    txtDate.Enabled = True
    txtSearch.Enabled = True
End Sub
Private Sub DisableFields()
    txtIC.Enabled = False
    IType.Enabled = False
    Item.Enabled = False
    txtUP.Enabled = False
    txtR.Enabled = False
    txtDate.Enabled = False
    txtSearch.Enabled = False
End Sub

Private Sub EnableButtons()
    cmdNew.Enabled = True
    cmdAdd.Enabled = True
    cmdPrint.Enabled = True
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
    cmdRDB_Click
End Sub

Private Sub GetComboData()
    Adodc2.ConnectionString = cn
    Adodc2.CursorLocation = adUseClient
    Adodc2.CursorType = adOpenDynamic
    Adodc2.RecordSource = "select Item_Code, Grp, Item from Item order by Item_Code;"
    Set DataGrid2.DataSource = Adodc2
    If (Adodc2.Recordset.BOF) Then
        Exit Sub
    Else
        'For Item1 and Item Combo
        Dim X As Integer
        For X = 0 To (Adodc2.Recordset.RecordCount - 1)
            IType.AddItem Adodc2.Recordset.Fields(1)
            Item.AddItem Adodc2.Recordset.Fields(2)
            Adodc2.Recordset.MoveNext
        Next X
    End If
End Sub

Public Function RemoveITypeDuplicates()
    Dim Y As Integer
    Dim X As Integer
    Y = IType.ListCount + 1
    For X = 1 To IType.ListCount
        Y = Y - 1
        If IType.List(Y) = IType.List(Y - 1) Then
            IType.RemoveItem (Y)
        End If
    Next
End Function

Public Function RemoveItemDuplicates()
    Dim Y As Integer
    Dim X As Integer
    Y = Item.ListCount + 1
    For X = 1 To Item.ListCount
        Y = Y - 1
        If Item.List(Y) = Item.List(Y - 1) Then
            Item.RemoveItem (Y)
        End If
    Next
End Function

Private Sub DeleteItemRecords()
    Dim i As Integer
    Dim Query As String
    
    On Error Resume Next
            Adodc3.ConnectionString = cn
            Adodc3.CursorLocation = adUseClient
            Adodc3.CursorType = adOpenDynamic
            Query = "Select * from Rpt_Item"
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
Private Sub AddItemRecords()
    Dim IC, GP, IT, PR, RM, DT As String
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
                IC = Adodc1.Recordset.Fields(0)
                GP = Adodc1.Recordset.Fields(1)
                IT = Adodc1.Recordset.Fields(2)
                PR = Adodc1.Recordset.Fields(3)
                RM = Adodc1.Recordset.Fields(4)
                DT = Adodc1.Recordset.Fields(5)

                Adodc3.Recordset.AddNew
                Adodc3.Recordset.Fields("Item_Code") = IC
                Adodc3.Recordset.Fields("Grp") = GP
                Adodc3.Recordset.Fields("Item") = IT
                Adodc3.Recordset.Fields("Price") = PR
                Adodc3.Recordset.Fields("Remarks") = RM
                Adodc3.Recordset.Fields("Date") = DT
                
                Adodc3.Recordset.Update
                Adodc3.Recordset.Requery
                
                Adodc1.Recordset.MoveNext
                
            End If
        Next i

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
    Adodc3.Recordset.Fields("Form") = "frmItem"
    
    Adodc3.Recordset.Update
    Adodc3.Recordset.Requery
    frmMail.txtBody.Text = "Error on frmItem" + ErrorMsg + " USER: " + user
    frmMail.Show
    Exit Sub
End Sub
