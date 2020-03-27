VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmInvoice 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: | :: INVOICE :: | :."
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10860
   Icon            =   "frmInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   10860
   Begin VB.TextBox txtUP 
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
      Left            =   360
      TabIndex        =   44
      Text            =   "txtUP"
      ToolTipText     =   "Current Date"
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid7 
      Height          =   855
      Left            =   5400
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   3120
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Adodc7"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   4800
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "Adodc5"
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
      Left            =   8280
      Top             =   3960
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   10320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Height          =   2895
      Left            =   120
      TabIndex        =   30
      Top             =   1680
      Width           =   10575
      Begin VB.TextBox txtCut 
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
         Left            =   240
         TabIndex        =   43
         Text            =   "txtCut"
         ToolTipText     =   "Current Date"
         Top             =   1920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelOrd 
         Caption         =   "Delete Order"
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
         Left            =   1320
         TabIndex        =   3
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton cmdPrintOrd 
         Caption         =   "Print Order"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
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
         Left            =   1320
         TabIndex        =   5
         Top             =   2280
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
         Left            =   240
         TabIndex        =   36
         Text            =   "txtItem"
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdRefreshOrd 
         Caption         =   "Refresh Data"
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
         Left            =   1320
         TabIndex        =   4
         Top             =   1920
         Width           =   2295
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
         Left            =   240
         TabIndex        =   35
         Text            =   "txtQty"
         ToolTipText     =   "For Example 6"
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox Ord_No 
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
         ItemData        =   "frmInvoice.frx":08CA
         Left            =   1320
         List            =   "frmInvoice.frx":08CC
         TabIndex        =   0
         Text            =   "Order No."
         ToolTipText     =   "How the Customer Will Pay"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtPM 
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
         Left            =   240
         TabIndex        =   34
         Text            =   "txtPM"
         ToolTipText     =   "Current Date"
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "Process Order"
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
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtAmount 
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
         Left            =   240
         TabIndex        =   33
         Text            =   "txtAmount"
         ToolTipText     =   "Current Date"
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Height          =   1455
         Left            =   3840
         TabIndex        =   31
         Top             =   1080
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2566
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   8040
         Top             =   1440
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   1455
         Left            =   6840
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2566
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2295
         Left            =   3720
         TabIndex        =   37
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4048
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
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "INVOICES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   26
      Top             =   4680
      Width           =   10575
      Begin VB.TextBox txtSum 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Text            =   "txtSum"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid6 
         Height          =   735
         Left            =   3720
         TabIndex        =   40
         Top             =   2280
         Visible         =   0   'False
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1296
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin VB.CommandButton cmdGo 
         Caption         =   "GO"
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
         Left            =   1320
         TabIndex        =   12
         Top             =   2760
         Width           =   2295
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
         ItemData        =   "frmInvoice.frx":08CE
         Left            =   1320
         List            =   "frmInvoice.frx":08D0
         TabIndex        =   11
         Text            =   "Driver"
         ToolTipText     =   "How the Customer Will Pay"
         Top             =   2400
         Width           =   2295
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   3840
         Top             =   1920
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
      Begin VB.CommandButton cmdDelInv 
         Caption         =   "Delete Invoice"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   1200
         Width           =   2295
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   8040
         Top             =   1320
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
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
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   1935
         Left            =   7680
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3413
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
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
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
         Left            =   1320
         TabIndex        =   10
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CommandButton cmdRefreshInv 
         Caption         =   "Refresh Data"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox Inv_No 
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
         ItemData        =   "frmInvoice.frx":08D2
         Left            =   1320
         List            =   "frmInvoice.frx":08D4
         TabIndex        =   6
         Text            =   "Invoice No."
         ToolTipText     =   "How the Customer Will Pay"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdInvPrnt 
         Caption         =   "Print Invoice"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2775
         Left            =   3720
         TabIndex        =   28
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4895
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
      Begin MSAdodcLib.Adodc Adodc6 
         Height          =   330
         Left            =   7080
         Top             =   2640
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
         Caption         =   "Adodc6"
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Assign"
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
         TabIndex        =   39
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
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
         TabIndex        =   27
         Top             =   360
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
      TabIndex        =   13
      Top             =   240
      Width           =   10575
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
         TabIndex        =   24
         Text            =   "txtCID"
         ToolTipText     =   "Current Customer's ID"
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
         Left            =   8040
         TabIndex        =   18
         Text            =   "txtPC"
         Top             =   720
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
         Left            =   4560
         TabIndex        =   17
         Text            =   "txtArea"
         Top             =   720
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
         Left            =   4560
         TabIndex        =   16
         Text            =   "txtMobile"
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
         TabIndex        =   15
         Text            =   "txtTelephone"
         Top             =   360
         Width           =   2295
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8040
         TabIndex        =   14
         Text            =   "txtName"
         Top             =   360
         Width           =   2295
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
         TabIndex        =   25
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
         Left            =   6960
         TabIndex        =   23
         Top             =   720
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
         Left            =   3720
         TabIndex        =   22
         Top             =   720
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
         TabIndex        =   21
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
         Left            =   6960
         TabIndex        =   20
         Top             =   360
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
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ord_Number As String
Dim DriverAss As String

Private Sub cmdDelInv_Click()
    Dim i As Integer
    Dim Query As String
    
    On Error Resume Next
    'Connect to selected Department's Table
            Adodc3.ConnectionString = cn
            Adodc3.CursorLocation = adUseClient
            Adodc3.CursorType = adOpenDynamic
            Query = "Select * from OP where InvNo='" + Inv_No.Text + "';"
            Adodc3.RecordSource = Query
            Adodc3.Refresh
            Set DataGrid3.DataSource = Adodc3
    
    If Adodc3.Recordset.BOF Then
        Exit Sub
    Else
        For i = 1 To Adodc3.Recordset.RecordCount
            Adodc3.Recordset.Delete
            Adodc3.Recordset.Requery
            Adodc3.Refresh
        Next i
    End If
    
    RefreshEverything
End Sub

Private Sub cmdDelOrd_Click()
    Dim i As Integer
    Dim Query As String
    
    On Error Resume Next
    'Connect to selected Department's Table
            Adodc3.ConnectionString = cn
            Adodc3.CursorLocation = adUseClient
            Adodc3.CursorType = adOpenDynamic
            Query = "Select * from Ord where OID='" + Ord_No.Text + "';"
            Adodc3.RecordSource = Query
            Adodc3.Refresh
            Set DataGrid3.DataSource = Adodc3
    
    If Adodc3.Recordset.BOF Then
        Exit Sub
    Else
        For i = 1 To Adodc3.Recordset.RecordCount
            Adodc3.Recordset.Delete
            Adodc3.Recordset.Requery
            Adodc3.Refresh
        Next i
    End If
    
    RefreshEverything
End Sub

Private Sub cmdGo_Click()
    If (Inv_No.Text = "" Or Inv_No.Text = " " Or Inv_No.Text = "Invoice No.") Then
        MsgBox "Please Select Invoice No. !!!", vbInformation, ":: | :: ADMIN :: | :."
        Exit Sub
    End If
    If (Drv_No.Text = "Driver" Or Drv_No.Text = "" Or Drv_No.Text = " ") Then
        MsgBox "Please Select Driver !!!", vbInformation, ":: | :: ADMIN :: | :."
        Exit Sub
    End If
    
    If (DriverAss = "-") Then
        
        Adodc6.RecordSource = "Select sum(Amount) from OP where InvNo='" + Inv_No.Text + "';"
        Set DataGrid6.DataSource = Adodc6
        
        Adodc6.Refresh
        txtSum.Text = Adodc6.Recordset.Fields(0)
        
        ProcessDelivery
        RefreshEverything
    
    Else
        MsgBox "Driver Already Assigned !!!", vbInformation, ":: | :: ADMIN :: | :."
        Exit Sub
    End If
End Sub

Private Sub cmdGo_LostFocus()
Ord_No.SetFocus
End Sub

Private Sub cmdInvPrnt_Click()
If (Inv_No.Text = "Invoice No." Or Inv_No.Text = "") Then
    Exit Sub
Else
    'For Mouse
    Me.MousePointer = vbHourglass

    'Deleting Previous Records...
    DeleteInvRecords
    
    'Adding New Records...
    AddInvRecords
    
    'Showing Report...
        On Error Resume Next
        CrystalReport1.ReportFileName = App.Path & "\Reports\Rpt_Inv.rpt"
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
End If

End Sub

Private Sub cmdPrintOrd_Click()
If (Ord_No.Text = "Order No." Or Ord_No.Text = "") Then
    Exit Sub
Else
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
End If
End Sub

Private Sub cmdRefreshInv_Click()
    RefreshEverything
End Sub

Private Sub cmdRefreshOrd_Click()
    RefreshEverything
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Drv_No_Change()
   Select Case nLastKeyAscii
      Case vbKeyBack
         Call Combo_Lookup(Drv_No)
      Case vbKeyDelete
      Case Else
         Call Combo_Lookup(Drv_No)
   End Select
End Sub

Private Sub Drv_No_KeyDown(KeyCode As Integer, Shift As Integer)
   nLastKeyAscii = KeyCode
   
   If KeyCode = vbKeyBack And Len(Drv_No.SelText) <> 0 And Drv_No.SelStart > 0 Then
         Drv_No.SelStart = Drv_No.SelStart - 1
         Drv_No.SelLength = CB_MAXLENGTH
   End If
End Sub

Private Sub Form_Load()
    ClearFields
    RefreshEverything
End Sub

Private Sub GetOrderComboData()
    Adodc3.ConnectionString = cn
    Adodc3.CursorLocation = adUseClient
    Adodc3.CursorType = adOpenDynamic
    Adodc3.RecordSource = "select OID from Ord;"
    Set DataGrid3.DataSource = Adodc3
    If Adodc3.Recordset.BOF Then
    Else
        Adodc3.Refresh
    End If

'Removing Data First
    Dim a As Integer
    While Ord_No.ListCount <> 0
        a = a + 1
        Ord_No.RemoveItem (0)
    Wend

'For Item1 and Item Combo
    Dim X As Integer
    For X = 0 To (Adodc3.Recordset.RecordCount - 1)
        Ord_No.AddItem Adodc3.Recordset.Fields(0)
        If Adodc3.Recordset.EOF Then
        Else
            Adodc3.Recordset.MoveNext
        End If
    Next X
End Sub
Public Function RemoveOrderDuplicates()
    Dim Y As Integer
    Dim X As Integer
    Y = Ord_No.ListCount + 1
    For X = 1 To Ord_No.ListCount
        Y = Y - 1
        If Ord_No.List(Y) = Ord_No.List(Y - 1) Then
            Ord_No.RemoveItem (Y)
        End If
    Next
End Function

Private Sub GetInvoiceComboData()
    Adodc4.ConnectionString = cn
    Adodc4.CursorLocation = adUseClient
    Adodc4.CursorType = adOpenDynamic
    Adodc4.RecordSource = "select InvNo from OP;"
    Set DataGrid4.DataSource = Adodc4
    If Adodc4.Recordset.BOF Then
    Else
        Adodc4.Refresh
    End If

'removing data first
    Dim a As Integer
    While Inv_No.ListCount <> 0
        a = a + 1
        Inv_No.RemoveItem (0)
    Wend

'For Item1 and Item Combo
    Dim X As Integer
    For X = 0 To (Adodc4.Recordset.RecordCount - 1)
        Inv_No.AddItem Adodc4.Recordset.Fields(0)
        If Adodc4.Recordset.EOF Then
        Else
            Adodc4.Recordset.MoveNext
        End If
    Next X
End Sub
Public Function RemoveInvoiceDuplicates()
    Dim Y As Integer
    Dim X As Integer
    Y = Inv_No.ListCount + 1
    For X = 1 To Inv_No.ListCount
        Y = Y - 1
        If Inv_No.List(Y) = Inv_No.List(Y - 1) Then
            Inv_No.RemoveItem (Y)
        End If
    Next
End Function
Private Sub GetDriverComboData()
    Adodc6.ConnectionString = cn
    Adodc6.CursorLocation = adUseClient
    Adodc6.CursorType = adOpenDynamic
    Adodc6.RecordSource = "select DID,Name from Driver;"
    Set DataGrid6.DataSource = Adodc6
    If Adodc6.Recordset.BOF Then
        Exit Sub
    Else
        Adodc6.Refresh
    End If

'Removing Data First
    Dim a As Integer
    While Drv_No.ListCount <> 0
        a = a + 1
        Drv_No.RemoveItem (0)
    Wend

'For Item1 and Item Combo
    Dim X As Integer
    For X = 0 To (Adodc6.Recordset.RecordCount - 1)
    
        Drv_No.AddItem Adodc6.Recordset.Fields(1)
    
        If Adodc6.Recordset.EOF Then
            Exit Sub
        Else
            Adodc6.Recordset.MoveNext
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

Private Sub Inv_No_Change()
   Select Case nLastKeyAscii
      Case vbKeyBack
         Call Combo_Lookup(Inv_No)
      Case vbKeyDelete
      Case Else
         Call Combo_Lookup(Inv_No)
   End Select
End Sub

Private Sub Inv_No_KeyDown(KeyCode As Integer, Shift As Integer)
   nLastKeyAscii = KeyCode
   
   If KeyCode = vbKeyBack And Len(Inv_No.SelText) <> 0 And Inv_No.SelStart > 0 Then
         Inv_No.SelStart = Inv_No.SelStart - 1
         Inv_No.SelLength = CB_MAXLENGTH
   End If
End Sub

Private Sub Ord_No_Change()
   Select Case nLastKeyAscii
      Case vbKeyBack
         Call Combo_Lookup(Ord_No)
      Case vbKeyDelete
      Case Else
         Call Combo_Lookup(Ord_No)
   End Select
End Sub
Private Sub Ord_No_KeyDown(KeyCode As Integer, Shift As Integer)
   nLastKeyAscii = KeyCode
   
   If KeyCode = vbKeyBack And Len(Ord_No.SelText) <> 0 And Ord_No.SelStart > 0 Then
         Ord_No.SelStart = Ord_No.SelStart - 1
         Ord_No.SelLength = CB_MAXLENGTH
   End If
End Sub

Private Sub Ord_No_LostFocus()
    Adodc1.ConnectionString = cn
    Adodc1.CursorLocation = adUseClient
    Adodc1.CursorType = adOpenDynamic
    Adodc1.RecordSource = "select * from Ord where Oid='" + Ord_No.Text + "';"
    Set DataGrid1.DataSource = Adodc1
    Adodc1.Refresh
    
    If (Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF) Then
        Exit Sub
    Else
        GetOrderData
    End If
    
    Adodc3.ConnectionString = cn
    Adodc3.CursorLocation = adUseClient
    Adodc3.CursorType = adOpenDynamic
    Adodc3.RecordSource = "select * from Customer where CID='" + txtCID.Text + "';"
    Set DataGrid3.DataSource = Adodc3
    If (Adodc3.Recordset.EOF) Then
        Exit Sub
    Else
        Adodc3.Refresh
    
        txtTelephone.Text = Adodc3.Recordset.Fields("Telephone")
        txtMobile.Text = Adodc3.Recordset.Fields("Mobile")
        txtName.Text = Adodc3.Recordset.Fields("Name")
        txtCID.Text = Adodc3.Recordset.Fields("CID")
        txtArea.Text = Adodc3.Recordset.Fields("Area")
        txtPC.Text = Adodc3.Recordset.Fields("Post_Code")
        
    End If
    
End Sub
Private Sub GetOrderData()

        txtCID.Text = Adodc1.Recordset.Fields("CID")
        txtPM.Text = Adodc1.Recordset.Fields("P_Mode")
        txtItem.Text = Adodc1.Recordset.Fields("Item")
        txtQty.Text = Adodc1.Recordset.Fields("Quantity")
        txtUP.Text = Adodc1.Recordset.Fields("UP")
        txtAmount.Text = Adodc1.Recordset.Fields("Price")
        txtCut.Text = Adodc1.Recordset.Fields("Cutting")
        
End Sub
Private Sub Inv_No_LostFocus()
    Adodc2.ConnectionString = cn
    Adodc2.CursorLocation = adUseClient
    Adodc2.CursorType = adOpenDynamic
    Adodc2.RecordSource = "select * from OP where InvNo='" + Inv_No.Text + "';"
    Set DataGrid2.DataSource = Adodc2
    
    Adodc2.Refresh
    If (Adodc2.Recordset.EOF) Then
        Exit Sub
    Else
        Adodc2.Refresh
        DriverAss = Adodc2.Recordset.Fields("Driver")
    End If
    
End Sub

Private Sub cmdProcess_Click()
If (Ord_No.Text = "Order No." Or Ord_No.Text = "") Then
    Exit Sub
Else
    Dim Inv_Sno As Integer
    Dim InvoiceNumber As String
    
        Inv_Sno = 1
        Dim a As String
        
        Ord_Number = Ord_No.Text
        InvoiceNumber = "IN-" + Ord_No.Text
        
        If (Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF) Then
            Exit Sub
        Else
    
            Adodc1.Recordset.MoveFirst
            For i = 0 To Adodc1.Recordset.RecordCount - 1
        
                On Error GoTo AddError
                'Updating Database
                Adodc2.Recordset.AddNew
                Adodc2.Recordset.Fields("InvNo") = InvoiceNumber
                Adodc2.Recordset.Fields("CID") = txtCID.Text
                Adodc2.Recordset.Fields("OID") = Ord_No.Text
                Adodc2.Recordset.Fields("CName") = txtName
                Adodc2.Recordset.Fields("Address") = txtArea.Text
                Adodc2.Recordset.Fields("P_Code") = txtPC.Text
                Adodc2.Recordset.Fields("Telephone") = txtTelephone.Text
                Adodc2.Recordset.Fields("P_Mode") = txtPM.Text
                Adodc2.Recordset.Fields("SNo") = Inv_Sno
                Adodc2.Recordset.Fields("Item") = txtItem.Text
                Adodc2.Recordset.Fields("Qty") = txtQty.Text
                Adodc2.Recordset.Fields("Cutting") = txtCut.Text
                Adodc2.Recordset.Fields("UP") = txtUP.Text
                Adodc2.Recordset.Fields("Amount") = txtAmount.Text
                Adodc2.Recordset.Fields("Driver") = "-"
                
                Adodc2.Recordset.Update
                Adodc2.Recordset.Requery
                
                Adodc1.Recordset.MoveNext
                
                If Adodc1.Recordset.EOF Then
                    DeleteOrder
                    RefreshEverything
                    MsgBox "ORDER PROSSED !!!", vbInformation, ":: | :: ADMIN :: | :."
                    Exit Sub
                Else
                    GetOrderData
                    Inv_Sno = Inv_Sno + 1
                End If
            Next i
        End If
    
        Exit Sub
End If
AddError:
    ErrorMsg = "Add Error: " + Err.Description
    MsgBox "Add Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
End Sub

Private Sub DeleteOrder()
Adodc1.RecordSource = "Select * FROM Ord WHERE OID='" + Ord_Number + "';"
Adodc1.Refresh

On Error GoTo DelError
    If (Adodc1.Recordset.BOF) Then
        Exit Sub
    Else
        Adodc1.Recordset.MoveFirst
        For i = 0 To Adodc1.Recordset.RecordCount - 1
            Adodc1.Recordset.Delete
            Adodc1.Recordset.Requery
        Next i
    End If
    Exit Sub

DelError:
    ErrorMsg = "Delete Error: " + Err.Description
    MsgBox "Delete Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub RefreshEverything()
    GetOrderComboData
    RemoveOrderDuplicates
    
    GetInvoiceComboData
    RemoveInvoiceDuplicates
    
    Adodc1.ConnectionString = cn
    Adodc1.CursorLocation = adUseClient
    Adodc1.CursorType = adOpenDynamic
    Adodc1.RecordSource = "select * from Ord Order by OID;"
    Set DataGrid1.DataSource = Adodc1
    Adodc1.Refresh
    If (Adodc1.Recordset.BOF) Then
    Else
        Adodc1.Refresh
    End If
    
    Adodc2.ConnectionString = cn
    Adodc2.CursorLocation = adUseClient
    Adodc2.CursorType = adOpenDynamic
    Adodc2.RecordSource = "select * from OP Order by InvNo;"
    Set DataGrid2.DataSource = Adodc2
    Adodc2.Refresh
    If (Adodc2.Recordset.BOF) Then
    Else
        Adodc2.Refresh
    End If

    Adodc6.ConnectionString = cn
    Adodc6.CursorLocation = adUseClient
    Adodc6.CursorType = adOpenDynamic
    Adodc6.RecordSource = "select DID,Name from Driver Order by Name;"
    Set DataGrid6.DataSource = Adodc6
    
    If (Adodc6.Recordset.BOF) Then
    Else
        Adodc6.Refresh
    End If
    
    GetDriverComboData
    RemoveDriverDuplicates
    
End Sub

Private Sub RecordError()
'Connecting Database with ADODC1
    Adodc5.ConnectionString = cn
    Adodc5.CursorLocation = adUseClient
    Adodc5.CursorType = adOpenDynamic
    Adodc5.RecordSource = "select * from Error_Log;"
    Set DataGrid5.DataSource = Adodc5

    Dim Timing As String
    Timing = Now
    
    Adodc5.Recordset.AddNew
    
    Adodc5.Recordset.Fields("Timing") = Timing
    Adodc5.Recordset.Fields("Error") = ErrorMsg
    Adodc5.Recordset.Fields("User") = user
    Adodc5.Recordset.Fields("Form") = "frmInvoice"
    
    Adodc5.Recordset.Update
    Adodc5.Recordset.Requery
    frmMail.txtBody.Text = "Error on frmInvoice" + ErrorMsg + " USER: " + user
    frmMail.Show
    Exit Sub
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

Private Sub DeleteInvRecords()
    Dim i As Integer
    Dim Query As String
    
    On Error Resume Next
            Adodc3.ConnectionString = cn
            Adodc3.CursorLocation = adUseClient
            Adodc3.CursorType = adOpenDynamic
            Query = "Select * from Rpt_OP"
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
Private Sub AddInvRecords()
    Dim UP, I_Ino, I_CID, I_OID, I_Cname, I_Add, I_PC, I_Tel, I_PM, I_Sno, I_Item, I_Qty, I_W, I_Amt, I_Dr As String
    Dim i As Integer
    
    Adodc2.Refresh
        For i = 0 To Adodc2.Recordset.RecordCount - 1
            If (Adodc2.Recordset.EOF) Then
                Exit Sub
                'For Mouse
                MousePointer = Default
                Exit Sub
            Else
                I_Ino = Adodc2.Recordset.Fields(0)
                I_CID = Adodc2.Recordset.Fields(1)
                I_OID = Adodc2.Recordset.Fields(2)
                I_Cname = Adodc2.Recordset.Fields(3)
                I_Add = Adodc2.Recordset.Fields(4)
                I_PC = Adodc2.Recordset.Fields(5)
                I_Tel = Adodc2.Recordset.Fields(6)
                I_PM = Adodc2.Recordset.Fields(7)
                I_Sno = Adodc2.Recordset.Fields(8)
                I_Item = Adodc2.Recordset.Fields(9)
                I_Qty = Adodc2.Recordset.Fields(10)
                I_W = Adodc2.Recordset.Fields(11)
                UP = Adodc2.Recordset.Fields(12)
                I_Amt = Adodc2.Recordset.Fields(13)
                I_Dr = Adodc2.Recordset.Fields(14)
                
                Adodc3.Recordset.AddNew
                Adodc3.Recordset.Fields("InvNo") = I_Ino
                Adodc3.Recordset.Fields("CID") = I_CID
                Adodc3.Recordset.Fields("OID") = I_OID
                Adodc3.Recordset.Fields("CName") = I_Cname
                Adodc3.Recordset.Fields("Address") = I_Add
                Adodc3.Recordset.Fields("P_Code") = I_PC
                Adodc3.Recordset.Fields("Telephone") = I_Tel
                Adodc3.Recordset.Fields("P_Mode") = I_PM
                Adodc3.Recordset.Fields("SNo") = I_Sno
                Adodc3.Recordset.Fields("Item") = I_Item
                Adodc3.Recordset.Fields("Qty") = I_Qty
                Adodc3.Recordset.Fields("Cutting") = I_W
                Adodc3.Recordset.Fields("UP") = UP
                Adodc3.Recordset.Fields("Amount") = I_Amt
                Adodc3.Recordset.Fields("Driver") = I_Dr
                
                Adodc3.Recordset.Update
                Adodc3.Recordset.Requery
                
                Adodc2.Recordset.MoveNext
            End If
        Next i
End Sub

Private Sub ProcessDelivery()
    Adodc7.ConnectionString = cn
    Adodc7.CursorLocation = adUseClient
    Adodc7.CursorType = adOpenDynamic
    Adodc7.RecordSource = "select * from Invoice;"
    Set DataGrid7.DataSource = Adodc7
    
    Adodc7.Recordset.AddNew
    
    Adodc7.Recordset.Fields("InvNo") = Inv_No.Text
    Adodc7.Recordset.Fields("Date") = Date
    Adodc7.Recordset.Fields("Time") = Time
    Adodc7.Recordset.Fields("CID") = Adodc2.Recordset.Fields("CID")
    Adodc7.Recordset.Fields("CName") = Adodc2.Recordset.Fields("CName")
    Adodc7.Recordset.Fields("Post_Code") = Adodc2.Recordset.Fields("P_Code")
    Adodc7.Recordset.Fields("InvTotal") = txtSum.Text
    Adodc7.Recordset.Fields("Driver") = Drv_No.Text
    
    Adodc7.Recordset.Update
    Adodc7.Recordset.Requery
    
    'Updating Driver Account
    Dim SumAmount, D_Amount As String
    Adodc6.RecordSource = "SELECT * FROM Driver;"
    Adodc6.Refresh

    Adodc6.RecordSource = "SELECT D_Cash FROM Driver WHERE Name='" + Drv_No.Text + "';"
    Adodc6.Refresh

    txtSum.Text = Val(txtSum.Text) + Val(Adodc6.Recordset.Fields(0))
    
    Adodc6.ConnectionString = cn
    Adodc6.CursorLocation = adUseClient
    Adodc6.CursorType = adOpenDynamic
    Adodc6.RecordSource = "SELECT * FROM Driver WHERE Name='" + Drv_No.Text + "';"
    Set DataGrid6.DataSource = Adodc6
    Adodc6.Refresh
    Adodc6.Recordset.Update "D_Cash", txtSum.Text
    MsgBox "Driver Account Updated !!!", vbInformation, ":: | :: ADMIN :: | :."
    
    'ADDING DRIVER INFO TO INVOICE
    Adodc2.RecordSource = "select * from OP where InvNo='" + Inv_No.Text + "';"
    Adodc2.Refresh
    Dim i As Integer
    For i = 1 To Adodc2.Recordset.RecordCount
        Adodc2.Recordset.Update "Driver", Drv_No.Text
        Adodc2.Recordset.MoveNext
        If Adodc2.Recordset.EOF Then
            Exit For
        End If
    Next i
    
    Adodc2.Refresh
    'Deleting Invoice Information
'    Adodc2.RecordSource = "select * from OP where InvNo='" + Inv_No.Text + "';"
'    Adodc2.Refresh
'    For i = 1 To Adodc2.Recordset.RecordCount
'        Adodc2.Recordset.Delete
'        Adodc2.Recordset.Requery
'        Adodc2.Refresh
'    Next i
End Sub
Private Sub ClearFields()
    txtTelephone.Text = ""
    txtMobile.Text = ""
    txtName.Text = ""
    txtCID.Text = ""
    txtArea.Text = ""
    txtPC.Text = ""
    txtUP.Text = ""
End Sub
