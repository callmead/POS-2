VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSales 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: | :: SALES :: | :."
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11340
   ForeColor       =   &H00000000&
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11340
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1935
      Left            =   5280
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3413
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4920
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   14208
      _Version        =   393216
      MousePointer    =   99
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&SEARCH && PRINT"
      TabPicture(0)   =   "frmSales.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&COMPARISONS"
      TabPicture(1)   =   "frmSales.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   24
         Top             =   4320
         Width           =   9255
         Begin VB.CommandButton cmdClose1 
            BackColor       =   &H00808080&
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   32
            ToolTipText     =   "Close This Form"
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdLast1 
            BackColor       =   &H00808080&
            Caption         =   "&Last Record"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7680
            TabIndex        =   31
            ToolTipText     =   "Move to Last Item"
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdNext1 
            BackColor       =   &H00808080&
            Caption         =   "Ne&xt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7680
            TabIndex        =   30
            ToolTipText     =   "Move to Next Item"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdPrevious1 
            BackColor       =   &H00808080&
            Caption         =   "&Previous"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   29
            ToolTipText     =   "Move to Previous Item"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdFirst1 
            BackColor       =   &H00808080&
            Caption         =   "&First Record"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   28
            ToolTipText     =   "Move to First Item"
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox Item1 
            ForeColor       =   &H00C00000&
            Height          =   315
            ItemData        =   "frmSales.frx":0902
            Left            =   240
            List            =   "frmSales.frx":0904
            Sorted          =   -1  'True
            TabIndex        =   27
            Text            =   "Select Item"
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdShow1 
            BackColor       =   &H00808080&
            Caption         =   "Show &Records"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   26
            ToolTipText     =   "If you want to see the records of specific Item then Chose Item and Click Me...."
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdRefresh1 
            BackColor       =   &H00808080&
            Caption         =   "&Refresh"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6240
            TabIndex        =   25
            ToolTipText     =   "Refresh Database"
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblRecords 
            Caption         =   "Total Records Available:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblItemshow1 
            Caption         =   "From DB"
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
            Left            =   2280
            TabIndex        =   33
            ToolTipText     =   "Total Categories Available in Database."
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   10815
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
            Left            =   2880
            TabIndex        =   6
            Top             =   6960
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
            Left            =   4200
            TabIndex        =   7
            Top             =   6960
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
            Left            =   5520
            TabIndex        =   8
            Top             =   6960
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
            Left            =   1560
            TabIndex        =   5
            Top             =   6960
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
            Left            =   6840
            TabIndex        =   9
            Top             =   6960
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
            Left            =   8160
            TabIndex        =   10
            Top             =   6960
            Width           =   1215
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
            Left            =   9120
            TabIndex        =   4
            Top             =   360
            Width           =   1335
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
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   240
            TabIndex        =   0
            Text            =   "txtSearch"
            Top             =   360
            Width           =   3855
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
            Left            =   6240
            TabIndex        =   2
            Top             =   360
            Width           =   1335
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
            Left            =   7680
            TabIndex        =   3
            Top             =   360
            Width           =   1335
         End
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
            ItemData        =   "frmSales.frx":0906
            Left            =   4200
            List            =   "frmSales.frx":0919
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   360
            Width           =   1935
         End
         Begin MSAdodcLib.Adodc Adodc1 
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
            Height          =   5895
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   10398
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   10815
         Begin VB.TextBox txtS2 
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
            Left            =   3360
            TabIndex        =   12
            Text            =   "txtS2"
            Top             =   360
            Width           =   1815
         End
         Begin VB.ComboBox ST2 
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
            ItemData        =   "frmSales.frx":0942
            Left            =   5280
            List            =   "frmSales.frx":094C
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdPrint2 
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
            Left            =   8280
            TabIndex        =   15
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdFind2 
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
            Left            =   7080
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtS1 
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
            Left            =   1200
            TabIndex        =   11
            Text            =   "txtS1"
            Top             =   360
            Width           =   1815
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
            Left            =   9480
            TabIndex        =   16
            Top             =   360
            Width           =   1095
         End
         Begin MSAdodcLib.Adodc Adodc3 
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
            Height          =   5895
            Left            =   240
            TabIndex        =   36
            Top             =   840
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   10398
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "&&"
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
            Left            =   3120
            TabIndex        =   21
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "BETWEEN"
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
            TabIndex        =   20
            Top             =   360
            Width           =   975
         End
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   -74280
         Top             =   4200
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
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
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   35
         ToolTipText     =   "Showing All the Items Records Available...."
         Top             =   480
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   4210752
         ForeColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFind2_Click()
If (txtS1.Text = "" Or txtS1.Text = " ") Then
    MsgBox "Please Specify Complete Information !!!", vbExclamation, ":: | :: ADMIN :: | :."
    txtS1.SetFocus
    Exit Sub
End If
If (txtS2.Text = "" Or txtS2.Text = " ") Then
    MsgBox "Please Specify Complete Information !!!", vbExclamation, ":: | :: ADMIN :: | :."
    txtS2.SetFocus
    Exit Sub
Else
    Adodc1.RecordSource = "SELECT * FROM Invoice WHERE " + ST2.Text + " BETWEEN '" + txtS1.Text + "' AND '" + txtS2.Text + "';"
    Set DataGrid3.DataSource = Adodc1
    If (Adodc1.Recordset.BOF) Then
    Else
        Adodc1.Refresh
    End If
End If

End Sub

Private Sub cmdPrint_Click()
    'For Mouse
    Me.MousePointer = vbHourglass

    'Deleting Previous Records...
    DeleteInvRecords
    
    'Adding New Records...
     AddInvRecords
    
    'Showing Report...
        CrystalReport1.ReportFileName = App.Path & "\Reports\Rpt_InvM.rpt"
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

Private Sub cmdRDB_LostFocus()
txtSearch.SetFocus
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1_LostFocus()
txtSearch.SetFocus
End Sub

Private Sub Form_Load()
    txtSearch.Text = ""
    txtS1.Text = ""
    txtS2.Text = ""
    
    
    Adodc1.ConnectionString = cn
    Adodc1.CursorLocation = adUseClient
    Adodc1.CursorType = adOpenDynamic
    Adodc1.RecordSource = "SELECT * FROM Invoice ORDER BY Date;"
    Set DataGrid1.DataSource = Adodc1
    Set DataGrid3.DataSource = Adodc1
    If (Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF) Then
    Else
        Adodc1.Refresh
    End If
    
End Sub

Private Sub cmdSearch_Click()
If (txtSearch.Text = "" Or txtSearch.Text = " ") Then
    MsgBox "Search what?", vbExclamation, ":: | :: ADMIN :: | :."
    txtSearch.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If

If (ST.Text = "" Or ST.Text = " ") Then
    MsgBox "Select Search Type !!!", vbInformation, ":: | :: ADMIN :: | :."
    ST.SetFocus
    Exit Sub
End If

Adodc1.RecordSource = "select * from Invoice where " + ST.Text + " Like '" + txtSearch.Text + "%'"
Adodc1.Refresh

'Checking Either the Record is Present in Database or not
If Adodc1.Recordset.EOF Then
    MsgBox "Record for " + txtSearch.Text + " Not Found !!!", vbInformation, ":: | :: ADMIN :: | :."
    Adodc1.RecordSource = "select * from Invoice Order by Date;"
    Adodc1.Refresh
    Exit Sub

Else

    'Getting Data in Text Fields
    Adodc1.Recordset.MoveFirst
    
    Exit Sub
    
End If
End Sub

Private Sub DeleteInvRecords()
    Dim i As Integer
    Dim Query As String
    
    On Error Resume Next
            Adodc2.ConnectionString = cn
            Adodc2.CursorLocation = adUseClient
            Adodc2.CursorType = adOpenDynamic
            Query = "SELECT * FROM Rpt_Inv;"
            Adodc2.RecordSource = Query
            Adodc2.Refresh
            Set DataGrid2.DataSource = Adodc2
    
    If Adodc2.Recordset.BOF Then
        Exit Sub
    Else
        Adodc2.Recordset.MoveFirst
        For i = 1 To Adodc2.Recordset.RecordCount
            Adodc2.Recordset.Delete
            Adodc2.Recordset.Requery
            Adodc2.Refresh
        Next i
    End If
    Exit Sub
End Sub
Private Sub AddInvRecords()
    On Error GoTo AddError
    Dim INV, DT, TM, CI, cn, PC, SM, DR As String
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
                INV = Adodc1.Recordset.Fields(0)
                DT = Adodc1.Recordset.Fields(1)
                TM = Adodc1.Recordset.Fields(2)
                CI = Adodc1.Recordset.Fields(3)
                cn = Adodc1.Recordset.Fields(4)
                PC = Adodc1.Recordset.Fields(5)
                SM = Adodc1.Recordset.Fields(6)
                DR = Adodc1.Recordset.Fields(7)

                Adodc2.Recordset.AddNew
                Adodc2.Recordset.Fields("InvNo") = INV
                Adodc2.Recordset.Fields("Date") = DT
                Adodc2.Recordset.Fields("Time") = TM
                Adodc2.Recordset.Fields("CID") = CI
                Adodc2.Recordset.Fields("CName") = cn
                Adodc2.Recordset.Fields("Post_Code") = PC
                Adodc2.Recordset.Fields("InvTotal") = SM
                Adodc2.Recordset.Fields("Driver") = DR
                
                Adodc2.Recordset.Update
                Adodc2.Recordset.Requery
                
                Adodc1.Recordset.MoveNext
                
            End If
        Next i
    Exit Sub
AddError:
MsgBox Err.Description
Exit Sub
End Sub

Private Sub ST_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cmdRDB_Click()
    Adodc1.RecordSource = "select * from Invoice Order by Date;"
    Adodc1.Refresh
    If (Adodc1.Recordset.BOF) Then
        Exit Sub
    Else
        Adodc1.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdMF_Click()
If (Adodc1.Recordset.RecordCount = 0) Then
    MsgBox "NO RECORDS FOUND !!!", vbInformation, ":: | :: ADMIN :: | :."
    Exit Sub
End If
On Error GoTo MError
    Adodc1.Recordset.MoveFirst
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
    End If
Exit Sub
    
MError:
    ErrorMsg = "Movement Error: " + Err.Description
    MsgBox "Movement Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
End Sub
Private Sub cmdDelete_Click()
On Error GoTo DelError
    If (Adodc1.Recordset.BOF) Then
        Exit Sub
    Else
    
        Adodc1.Recordset.Delete
        cmdRDB_Click
        
        MsgBox "Record Deleted !!!", vbInformation, ""
        Adodc1.Refresh
        Exit Sub
    End If

DelError:
    ErrorMsg = "Delete Error: " + Err.Description
    MsgBox "Delete Error: " + Err.Description, vbCritical, ":: | :: ADMIN :: | :."
    RecordError
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
    Adodc3.Recordset.Fields("Form") = "frmSales"
    
    Adodc3.Recordset.Update
    Adodc3.Recordset.Requery
    frmMail.txtBody.Text = "Error on frmSales" + ErrorMsg + " USER: " + user
    frmMail.Show
    Exit Sub
End Sub

