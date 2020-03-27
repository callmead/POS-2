VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmOrder 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: | :: ORDERS :: | :."
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10845
   Icon            =   "frmOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   10845
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Height          =   855
      Left            =   4560
      TabIndex        =   45
      Top             =   7560
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1508
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
      Left            =   3240
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      TabIndex        =   38
      Top             =   6480
      Width           =   10575
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   480
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSDataGridLib.DataGrid DataGrid5 
         Height          =   1695
         Left            =   240
         TabIndex        =   54
         ToolTipText     =   "Current Orders"
         Top             =   840
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   16777215
         DefColWidth     =   104
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
         Left            =   7800
         TabIndex        =   48
         Top             =   360
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   480
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Left            =   5160
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   240
         TabIndex        =   23
         Text            =   "txtSearch"
         Top             =   360
         Width           =   4815
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
         Left            =   6480
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   7440
         Top             =   1440
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
         Height          =   855
         Left            =   7320
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1508
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
         Height          =   855
         Left            =   240
         TabIndex        =   40
         ToolTipText     =   "Current Orders"
         Top             =   1680
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   1508
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
   Begin VB.Frame framee 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ORDER LIST"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   33
      Top             =   1680
      Width           =   10575
      Begin VB.TextBox txtGT 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "txtGT"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtOid 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2880
         TabIndex        =   7
         Text            =   "txtOID"
         ToolTipText     =   "Order ID"
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox txtPM 
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
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmOrder.frx":08CA
         Left            =   5640
         List            =   "frmOrder.frx":08D7
         Sorted          =   -1  'True
         TabIndex        =   8
         Text            =   "Payment Mode"
         ToolTipText     =   "How the Customer Will Pay"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtItem 
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
         Left            =   2040
         TabIndex        =   47
         Text            =   "txtItem"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtGroup 
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
         TabIndex        =   46
         Text            =   "txtGroup"
         Top             =   1440
         Width           =   1695
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
         Left            =   3360
         TabIndex        =   16
         Top             =   3840
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
         Left            =   2640
         TabIndex        =   19
         Top             =   4200
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
         Left            =   3960
         TabIndex        =   20
         Top             =   4200
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
         Left            =   5400
         TabIndex        =   21
         Top             =   4200
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
         Left            =   2040
         TabIndex        =   0
         Top             =   3840
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
         Left            =   6000
         TabIndex        =   17
         Top             =   3840
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
         Left            =   6720
         TabIndex        =   22
         Top             =   4200
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
         Left            =   7320
         TabIndex        =   18
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtP 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6360
         MaxLength       =   20
         TabIndex        =   14
         Text            =   "txtP"
         Top             =   2880
         Width           =   3975
      End
      Begin VB.TextBox txtC 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6360
         MaxLength       =   20
         TabIndex        =   13
         Text            =   "txtC"
         Top             =   2400
         Width           =   3975
      End
      Begin VB.ListBox List2 
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
         ForeColor       =   &H00000000&
         Height          =   1980
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Select Item"
         Top             =   1440
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.ListBox List1 
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
         ForeColor       =   &H00000000&
         Height          =   1980
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Select Item Group"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "txtCID"
         ToolTipText     =   "Current Customer's ID"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtQty 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6360
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "txtQty"
         ToolTipText     =   "For Example 6"
         Top             =   1440
         Width           =   975
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
         ForeColor       =   &H00000000&
         Height          =   795
         Left            =   7560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "frmOrder.frx":091C
         Top             =   1440
         Width           =   2775
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8040
         TabIndex        =   9
         Text            =   "txtDate"
         ToolTipText     =   "Current Date"
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Finish"
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
         Left            =   4680
         TabIndex        =   44
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label txtPrice 
         BackStyle       =   0  'Transparent
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
         Left            =   6360
         TabIndex        =   57
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label txtRem 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   56
         Top             =   3480
         Width           =   5055
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Price"
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
         Left            =   7440
         TabIndex        =   55
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Left            =   5520
         TabIndex        =   52
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label10 
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
         Left            =   8040
         TabIndex        =   51
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Mode"
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
         Left            =   5640
         TabIndex        =   50
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No."
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
         Left            =   2880
         TabIndex        =   49
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label16 
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
         Left            =   2040
         TabIndex        =   43
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label15 
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
         Left            =   240
         TabIndex        =   42
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         TabIndex        =   41
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Packing"
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
         Left            =   5520
         TabIndex        =   37
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Cutting"
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
         Left            =   5520
         TabIndex        =   36
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty/W"
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
         Left            =   5520
         TabIndex        =   35
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Instructions"
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
         Left            =   7560
         TabIndex        =   34
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   26
      Top             =   240
      Width           =   10575
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8040
         TabIndex        =   3
         Text            =   "txtName"
         Top             =   360
         Width           =   2295
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   1
         Text            =   "txtTelephone"
         Top             =   360
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4800
         MaxLength       =   11
         TabIndex        =   2
         Text            =   "txtMobile"
         Top             =   360
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1320
         MaxLength       =   45
         TabIndex        =   4
         Text            =   "txtArea"
         Top             =   720
         Width           =   2295
      End
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4800
         MaxLength       =   9
         TabIndex        =   5
         Text            =   "txtPC"
         Top             =   720
         Width           =   2295
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
         TabIndex        =   31
         Top             =   360
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
         Left            =   7200
         TabIndex        =   30
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
         Left            =   3720
         TabIndex        =   29
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
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
         Left            =   720
         TabIndex        =   28
         Top             =   720
         Width           =   1815
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
         Left            =   3720
         TabIndex        =   27
         Top             =   720
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   6120
      Top             =   480
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
      Height          =   855
      Left            =   6000
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1508
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
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isAdd As Boolean
    Dim p As Double
    Dim Rks As String

Private Sub cmdPrint_Click()
    'For Mouse
    Me.MousePointer = vbHourglass

    'Deleting Previous Records...
    DeleteOrderRecords
    
    'Adding New Records...
    AddOrderRecords
    
    'Showing Report...
        CrystalReport1.ReportFileName = App.Path & "\Reports\Rpt_Ord.rpt"
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
    ClearAllFields
    ConnectOrders
    GetListsReady
    GetOrderData
    isAdd = False
    txtSearch.Text = ""
    txtRem.Caption = ""

End Sub

Private Sub cmdNew_Click()
    EnableFields
    ClearAllFields
    txtTelephone.SetFocus
    DisableButtons
    txtSearch.Enabled = False
    cmdNew.Visible = False
    cmdDelete.Visible = False
    cmdCancel.Enabled = True
    cmdAdd.Enabled = True
    
    txtGroup.Visible = False
    txtItem.Visible = False
    List1.Visible = True
    List2.Visible = True
    txtSearch.Enabled = False
    txtGT.Visible = True
    txtGT.Text = 0
    cmdAdd.Enabled = True
End Sub

Private Sub cmdAdd_Click()
'Checking Fields for Records
    If (txtCID.Text = "" Or txtCID.Text = " ") Then
        MsgBox "Enter Customer ID !!!", vbOKOnly, "Information Required"
        txtCID.SetFocus
        Exit Sub
    End If
    If (txtPM.Text = "" Or txtPM.Text = " ") Then
        MsgBox "What will be the Payment Mode for " + txtName.Text + " ?", vbQuestion, "Information Required"
        txtPM.SetFocus
        Exit Sub
    End If
    If (txtDate.Text = "" Or txtDate.Text = " ") Then
        txtDate.Text = Date
        Exit Sub
    End If
    If (List1.Text = "") Then
        MsgBox "Select Item Group !!!", vbOKOnly, "Information Required"
        List1.SetFocus
        Exit Sub
    End If
    If (List2.Text = "") Then
        MsgBox "Select Item !!!", vbOKOnly, "Information Required"
        List2.SetFocus
        Exit Sub
    End If
    If (txtQty.Text = "" Or txtQty.Text = " ") Then
        MsgBox "How much " + List2.Text + " is required by " + txtName.Text + " ?", vbQuestion, "Information Required"
        txtQty.SetFocus
        Exit Sub
    End If

    If (txtC.Text = "" Or txtC.Text = " ") Then
        MsgBox "Enter Cutting for " + List2.Text + "", vbQuestion, "Information Required"
        txtC.SetFocus
        Exit Sub
    End If
    If (txtP.Text = "" Or txtP.Text = " ") Then
        MsgBox "How will " + List2.Text + " be packed ?", vbQuestion, "Information Required"
        txtP.SetFocus
        Exit Sub
    End If
    If (txtR.Text = "") Then
        txtR.Text = "-"
    End If
    
    
    On Error GoTo AddError
    'Updating Database
    Adodc1.Recordset.AddNew

    Adodc1.Recordset.Fields("Date") = txtDate.Text
    Adodc1.Recordset.Fields("CID") = txtCID.Text
    Adodc1.Recordset.Fields("OID") = txtOid.Text
    Adodc1.Recordset.Fields("Grp") = List1.Text
    Adodc1.Recordset.Fields("P_Mode") = txtPM.Text
    Adodc1.Recordset.Fields("Item") = List2.Text
    Adodc1.Recordset.Fields("Quantity") = txtQty.Text
    Adodc1.Recordset.Fields("Cutting") = txtC.Text
    Adodc1.Recordset.Fields("Packing") = txtP.Text
    Adodc1.Recordset.Fields("UP") = p
    Adodc1.Recordset.Fields("Price") = txtPrice.Caption
    Adodc1.Recordset.Fields("Remarks") = txtR.Text
    Adodc1.Recordset.Fields("Date") = txtDate.Text

    Adodc1.Recordset.Update
    Adodc1.Recordset.Requery
    Adodc5.Refresh
        
    txtGT.Text = Val(txtPrice.Caption) + Val(txtGT.Text)
    ClearOrderFields
    List1.SetFocus
    txtPM.Enabled = False
    isAdd = True
    Exit Sub
    
AddError:
    ErrorMsg = "Add Error: " + Err.Description
    MsgBox "Add Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
End Sub

Private Sub cmdCancel_Click()
    If (isAdd = True) Then
        MsgBox "Record Added in Database Sucessfully !!! Total amount for this delivery is £" + txtGT.Text, vbInformation, "Conformation"
        isAdd = False
    End If
    
    Normalize
    cmdAdd.Enabled = False
    txtGT.Visible = False
    List1.Visible = False
    List2.Visible = False
    txtGroup.Visible = True
    txtItem.Visible = True
    
    txtSearch.Enabled = True
    cmdNew.Visible = True
    cmdDelete.Visible = True
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DelError
    If (Adodc1.Recordset.BOF) Then
        Exit Sub
    Else
    
        Adodc1.Recordset.Delete
        Adodc5.Recordset.Delete
        Adodc1.Refresh
        Adodc5.Refresh
        
        Normalize
        
        MsgBox "Record Deleted !!!", vbInformation, ""
        ClearAllFields
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
    Adodc1.RecordSource = "select * from Ord Order by Date;"
    Adodc1.Refresh
    Adodc5.RecordSource = "SELECT Date, CID, OID, P_Mode, Item, Price FROM Ord ORDER BY Date;"
    Adodc5.Refresh
    
    If (Adodc1.Recordset.BOF Or Adodc5.Recordset.BOF) Then
        Exit Sub
    Else
        Adodc5.Recordset.MoveFirst
        Adodc1.Recordset.MoveFirst
        GetOrderData
    End If
End Sub

Private Sub cmdMF_Click()
If (Adodc1.Recordset.RecordCount = 0) Then
    MsgBox "NO RECORDS FOUND !!!", vbInformation, ":: | :: ADMIN :: | :."
    Exit Sub
End If
On Error GoTo MError
    Adodc1.Recordset.MoveFirst
    Adodc5.Recordset.MoveFirst
    GetOrderData
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
    Adodc5.Recordset.MoveLast
    GetOrderData
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
    Adodc5.Recordset.MoveNext
    
    If (Adodc1.Recordset.EOF Or Adodc5.Recordset.EOF) Then
        MsgBox "This is Last Record !!!", vbInformation, ":: | :: ADMIN :: | :."
        Adodc1.Recordset.MoveLast
        Adodc5.Recordset.MoveLast
    Else
        GetOrderData
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
    
    If (Adodc1.Recordset.BOF Or Adodc5.Recordset.BOF) Then
        MsgBox "This is First Record !!!", vbInformation, ":: | :: ADMIN :: | :."
        Adodc1.Recordset.MoveFirst
        Adodc5.Recordset.MoveFirst
    Else
        GetOrderData
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

Adodc1.RecordSource = "select * from Ord where OID Like '" + txtSearch.Text + "%';"
Adodc5.RecordSource = "select * from Ord where OID Like '" + txtSearch.Text + "%';"
Adodc1.Refresh
Adodc5.Refresh

'Checking Either the Record is Present in Database or not
If Adodc1.Recordset.EOF Then
    MsgBox "Record for " + txtSearch.Text + " Not Found !!!", vbInformation, ":: | :: ADMIN :: | :."
    Adodc1.RecordSource = "select * from Ord Order by CID;"
    Adodc1.Refresh
    
    Adodc5.RecordSource = "SELECT Date, CID, OID, P_Mode, Item, Price FROM Ord ORDER BY Date;"
    Adodc5.Refresh
    
    Exit Sub

Else

    'Getting Data in Text Fields
    On Error GoTo SError
    Adodc1.Recordset.MoveFirst
    Adodc5.Recordset.MoveFirst
    Exit Sub
    
End If

SError:
    ErrorMsg = "Search Error: " + Err.Description
    MsgBox "Search Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError

End Sub


Private Sub ClearAllFields()
    txtTelephone.Text = ""
    txtMobile.Text = ""
    txtName.Text = ""
    txtArea.Text = ""
    txtPC.Text = ""
    txtCID.Text = ""
    txtOid.Text = ""
    txtPM.Text = ""
    txtDate.Text = ""
    txtQty.Text = ""
    txtC.Text = ""
    txtP.Text = ""
    txtPrice.Caption = ""
    txtR.Text = ""
    txtSearch.Text = ""
    txtGroup.Text = ""
    txtItem.Text = ""
End Sub
Private Sub ClearOrderFields()
    txtQty.Text = ""
    txtC.Text = ""
    txtP.Text = ""
    txtPrice.Caption = ""
    txtR.Text = ""
    txtSearch.Text = ""
End Sub

Private Sub EnableFields()
    txtTelephone.Enabled = True
    
    txtPM.Enabled = True
    txtOid.Enabled = True
    txtDate.Enabled = True
    List1.Enabled = True
    List2.Enabled = True
    txtQty.Enabled = True
    txtC.Enabled = True
    txtP.Enabled = True
    txtR.Enabled = True
    
    txtSearch.Enabled = True
End Sub
Private Sub DisableFields()
    txtTelephone.Enabled = False
    
    txtPM.Enabled = False
    txtOid.Enabled = False
    txtDate.Enabled = False
    List1.Enabled = False
    List2.Enabled = False
    txtQty.Enabled = False
    txtC.Enabled = False
    txtP.Enabled = False
    txtR.Enabled = False
    
    txtSearch.Enabled = False
End Sub

Private Sub EnableButtons()
    cmdNew.Enabled = True
    cmdAdd.Enabled = True
    cmdPrint.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = True
    cmdRDB.Enabled = True
    cmdMF.Enabled = True
    cmdN.Enabled = True
    cmdP.Enabled = True
    cmdML.Enabled = True
    cmdSearch.Enabled = True
End Sub
Private Sub DisableButtons()
    cmdNew.Enabled = False
    cmdAdd.Enabled = False
    cmdPrint.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    cmdRDB.Enabled = False
    cmdMF.Enabled = False
    cmdN.Enabled = False
    cmdP.Enabled = False
    cmdML.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub Normalize()
    DisableFields
    EnableButtons
    txtSearch.Enabled = True
    cmdRDB_Click
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
    Adodc3.Recordset.Fields("Form") = "frmOrder"
    
    Adodc3.Recordset.Update
    Adodc3.Recordset.Requery
    frmMail.txtBody.Text = "Error on frmOrder" + ErrorMsg + " USER: " + user
    frmMail.Show
    Exit Sub
End Sub


Private Sub List2_Click()
GetItemPrice
End Sub

Private Sub txtP_LostFocus()
    txtPrice.Caption = (Val(txtQty.Text) * p)
End Sub

Private Sub txtPM_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtPM_LostFocus()
    If (txtPM.Text = "Payment Mode" Or txtPM.Text = "") Then
        MsgBox "Enter Payment Mode for " + txtName.Text + "", vbInformation, ":: | :: ADMIN :: | :."
        txtPM.SetFocus
        Exit Sub
    End If
    
    List1.SetFocus
End Sub

Private Sub txtTelephone_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtTelephone_LostFocus()
    If (txtTelephone.Text = "" Or txtTelephone.Text = " ") Then
        MsgBox "Enter Customer Data !!!", vbCritical, ":: | :: ADMIN :: | :."
        txtTelephone.SetFocus
        Exit Sub
    End If
    If (txtTelephone.Text = "q" Or txtTelephone.Text = "Q") Then
        Unload Me
        Exit Sub
    End If
    
    GetCustomerData
End Sub

Private Sub GetCustomerData()
    Adodc2.ConnectionString = cn
    Adodc2.CursorLocation = adUseClient
    Adodc2.RecordSource = "SELECT * FROM Customer;"
    Set DataGrid2.DataSource = Adodc2
    
    Adodc2.RecordSource = "select * from Customer where Telephone ='" + txtTelephone.Text + "';"
    Adodc2.Refresh

    'Checking Either the Record is Present in Database or not
    If Adodc2.Recordset.EOF Then
        NewTelephone = txtTelephone.Text
        MsgBox "Record for " + txtTelephone.Text + " Not Found !!! Please Enter Customer Data.", vbInformation, ":: | :: ADMIN :: | :."
        Adodc2.RecordSource = "select * from Customer;"
        Adodc2.Refresh
        Unload Me
        
        'frmCustomer.Show
        frmCustomer.EnterNewCustomer
        frmCustomer.txtTelephone.Text = NewTelephone
        frmCustomer.txtMobile.SetFocus
        Exit Sub
    
    Else
        txtMobile.Text = Adodc2.Recordset.Fields("Mobile")
        txtName.Text = Adodc2.Recordset.Fields("Name")
        txtArea.Text = Adodc2.Recordset.Fields("Area")
        txtPC.Text = Adodc2.Recordset.Fields("Post_Code")
        txtCID.Text = Adodc2.Recordset.Fields("CID")
        txtPM.SetFocus
        
        Dim t As String
        t = Time
        txtDate.Text = Date
        txtOid.Text = txtCID.Text + "-" + t + ""
    End If
End Sub

Private Sub GetListsReady()
    Adodc4.ConnectionString = cn
    Adodc4.CursorLocation = adUseClient
    Adodc4.CursorType = adOpenDynamic
    Adodc4.RecordSource = "select * from item;"
    Adodc4.Refresh
    Set DataGrid4.DataSource = Adodc4

    'For Item1 and Item Combo
    Dim X As Integer
    For X = 0 To (Adodc4.Recordset.RecordCount - 1)
    
        List1.AddItem Adodc4.Recordset.Fields(1)
        List2.AddItem Adodc4.Recordset.Fields(2)

        Adodc4.Recordset.MoveNext
    Next X
    RemoveList1Duplicates
    RemoveList2Duplicates
End Sub
Private Sub List1_Click()
    Dim s, t, u As Integer
    
    t = List2.ListCount
    s = 0
    
    If (t <> 0) Then
        For u = 1 To t
            List2.RemoveItem (s)
        Next
    End If
    
    GetGroupItemsReady
End Sub

Private Sub RemoveList1Duplicates()
    Dim Y As Integer
    Dim X As Integer
    Y = List1.ListCount + 1
    For X = 1 To List1.ListCount
        Y = Y - 1
        If List1.List(Y) = List1.List(Y - 1) Then
            List1.RemoveItem (Y)
        End If
    Next
End Sub

Private Sub RemoveList2Duplicates()
    Dim Y As Integer
    Dim X As Integer
    Y = List2.ListCount + 1
    For X = 1 To List2.ListCount
        Y = Y - 1
        If List2.List(Y) = List2.List(Y - 1) Then
            List2.RemoveItem (Y)
        End If
    Next
End Sub

Private Sub GetGroupItemsReady()
    Adodc4.ConnectionString = cn
    Adodc4.CursorLocation = adUseClient
    Adodc4.CursorType = adOpenDynamic
    
    Adodc4.RecordSource = "SELECT ITEM FROM ITEM WHERE Grp='" + List1.Text + "';"
    Adodc4.Refresh
    Dim X As Integer
    For X = 0 To (Adodc4.Recordset.RecordCount - 1)
    List2.AddItem Adodc4.Recordset.Fields(0)
    Adodc4.Recordset.MoveNext
    Next X

    RemoveList2Duplicates

End Sub

Private Sub GetItemPrice()
    If (List1.Text = "") Then
        MsgBox "Please Select Item Group First !!!", vbCritical, ":: | :: ADMIN :: | :."
        List1.SetFocus
        Exit Sub
    Else
    Adodc4.ConnectionString = cn
    Adodc4.CursorLocation = adUseClient
    Adodc4.CursorType = adOpenDynamic
    Adodc4.RecordSource = "SELECT Price, Remarks FROM Item WHERE Grp ='" + List1.Text + "' AND Item ='" + List2.Text + "';"
    Adodc4.Refresh
    Set DataGrid4.DataSource = Adodc4
    
    p = Adodc4.Recordset.Fields("Price")
    Rks = Adodc4.Recordset.Fields("Remarks")
    
    txtPrice.Caption = p
    txtRem.Caption = Rks
    End If
End Sub

Private Sub DeleteOrderRecords()
    Dim i As Integer
    Dim Query As String
    
    On Error Resume Next
            Adodc3.ConnectionString = cn
            Adodc3.CursorLocation = adUseClient
            Adodc3.CursorType = adOpenDynamic
            Query = "Select * from Rpt_Ord"
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
Private Sub AddOrderRecords()
    Dim UP, O_Date, O_Cid, O_Oid, O_Grp, O_PMode, O_Item, O_Qty, O_W, O_C, O_P, O_Prc, O_R As String
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
                O_Date = Adodc1.Recordset.Fields(0)
                O_Cid = Adodc1.Recordset.Fields(1)
                O_Oid = Adodc1.Recordset.Fields(2)
                O_Grp = Adodc1.Recordset.Fields(3)
                O_PMode = Adodc1.Recordset.Fields(4)
                O_Item = Adodc1.Recordset.Fields(5)
                O_Qty = Adodc1.Recordset.Fields(6)
                O_C = Adodc1.Recordset.Fields(7)
                O_P = Adodc1.Recordset.Fields(8)
                UP = Adodc1.Recordset.Fields(9)
                O_Prc = Adodc1.Recordset.Fields(10)
                O_R = Adodc1.Recordset.Fields(11)

                Adodc3.Recordset.AddNew
                Adodc3.Recordset.Fields("Date") = O_Date
                Adodc3.Recordset.Fields("CID") = O_Cid
                Adodc3.Recordset.Fields("OID") = O_Oid
                Adodc3.Recordset.Fields("Grp") = O_Grp
                Adodc3.Recordset.Fields("P_Mode") = O_PMode
                Adodc3.Recordset.Fields("Item") = O_Item
                Adodc3.Recordset.Fields("Quantity") = O_Qty
                Adodc3.Recordset.Fields("Cutting") = O_C
                Adodc3.Recordset.Fields("Packing") = O_P
                Adodc3.Recordset.Fields("UP") = UP
                Adodc3.Recordset.Fields("Price") = O_Prc
                Adodc3.Recordset.Fields("Remarks") = O_R
                
                Adodc3.Recordset.Update
                Adodc3.Recordset.Requery
                
                Adodc1.Recordset.MoveNext
                
            End If
        Next i

End Sub

