VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMail 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: | :: MAIL :: | :."
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "frmMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9975
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9735
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
         Height          =   315
         Left            =   8280
         TabIndex        =   13
         Top             =   3165
         Width           =   1215
      End
      Begin VB.TextBox txtOS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "txtOS"
         ToolTipText     =   "Your Operating Syatem"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtVer 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "txtVer"
         ToolTipText     =   "OS Version"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtIP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "txtIP"
         ToolTipText     =   "Your IP Address"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtHost 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "txtHost"
         ToolTipText     =   "Your Host Name"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset Fields"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   3165
         Width           =   1215
      End
      Begin VB.TextBox txtToName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         ToolTipText     =   "The Person Name to whom the mail is being send"
         Top             =   360
         Width           =   4695
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   6
         ToolTipText     =   "Password"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtUser 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "Your Default Email Account User Name."
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtBody 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2415
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         ToolTipText     =   "Type your mail here."
         Top             =   3480
         Width           =   9255
      End
      Begin VB.TextBox txtToAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         ToolTipText     =   "Email address of the person to whom the mail is being send."
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtSubject 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "Mail Subject"
         Top             =   1080
         Width           =   4695
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send Mail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   285
         TabIndex        =   1
         Top             =   3165
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   17
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   660
      End
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   8520
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSMAPI.MAPIMessages mpmErrorMail 
      Left            =   9045
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession mpsErrorMail 
      Left            =   7920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   0   'False
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IP, Host, OS, Ver, Build As String

'Getting IP
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Const ERROR_SUCCESS       As Long = 0
Private Const WS_VERSION_REQD     As Long = &H101
Private Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD    As Long = 1
Private Const SOCKET_ERROR        As Long = -1

Private Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Private Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type

Private Declare Function WSAGetLastError Lib "wsock32" () As Long

Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Declare Function gethostname Lib "wsock32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   
Private Declare Function gethostbyname Lib "wsock32" _
  (ByVal szHost As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" (hpvDest As Any, _
   ByVal hpvSource As Long, _
   ByVal cbCopy As Long)
   
Private Function GetIPAddress() As String

   Dim sHostName    As String * 256
   Dim lpHost    As Long
   Dim Host      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
   End If
    
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
              " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
    
  'to extract the returned IP address, we have to copy
  'the HOST structure and its members
   CopyMemory Host, lpHost, Len(Host)
   CopyMemory dwIPAddr, Host.hAddrList, 4
   
  'create an array to hold the result
   ReDim tmpIPAddr(1 To Host.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, Host.hLen
   
  'and with the array, build the actual address,
  'appending a period between members
   For i = 1 To Host.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
  
  'the routine adds a period to the end of the
  'string, so remove it here
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
    
End Function

Private Function GetIPHostName() As String

    Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function

Private Function HiByte(ByVal wParam As Integer) As Byte
  
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function

Private Function LoByte(ByVal wParam As Integer) As Byte

  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function

Private Sub SocketsCleanup()

    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub

Private Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
 
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function
'********************************************************************

Private Sub cmdClose_Click()
Unload Me
End Sub

'****************************************************************************
Private Sub cmdReset_Click()
    txtToName.Text = ""
    txtToAddress.Text = ""
    txtSubject.Text = ""
    txtUser.Text = ""
    txtPassword.Text = ""
    txtBody.Text = ""
    
    txtToName.SetFocus
End Sub

Private Sub cmdSend_Click()
    On Error GoTo SendErrorMailError

    ' Sign on to the mail system.
    mpsErrorMail.UserName = txtUser.Text
    mpsErrorMail.password = txtPassword.Text
    mpsErrorMail.SignOn

    ' Send the message.
    mpmErrorMail.SessionID = mpsErrorMail.SessionID
    mpmErrorMail.Compose
    mpmErrorMail.RecipDisplayName = txtToName.Text
    mpmErrorMail.RecipAddress = txtToAddress.Text
    mpmErrorMail.AddressResolveUI = False
    mpmErrorMail.MsgSubject = txtSubject.Text
    mpmErrorMail.MsgNoteText = txtBody.Text + Host + IP + OS + Ver + " || " + Build
    mpmErrorMail.Send False

    ' Sign off of the mail system.
    mpsErrorMail.SignOff
    MsgBox "Your Mail has been send to " + txtToAddress.Text, vbInformation, ":: | :: ADMIN :: | :."
    Unload Me
    Exit Sub

SendErrorMailError:
    ' There was an error sending the mail.
    ' Just present the message.
    MsgBox "Error " & Format$(Err.Number) & _
        " sending mail." & vbCrLf & _
        Err.Description
    Exit Sub
End Sub

Private Sub Form_Load()

txtToName.Text = "Developer"
txtToAddress.Text = "adeel_s90@hotmail.com"
txtSubject.Text = "About Your Software"
txtUser.Text = ""

   Host = "Host Name: " + GetIPHostName()
   IP = "IP Address: " + GetIPAddress()

'Getting System Info
        Select Case SysInfo.OSPlatform
                Case 0
                        OS = "OS Platform: Unknown 32-Bit Windows"
                Case 1
                        OS = "OS Platform: Windows 95"
                Case 2
                        OS = "OS Platform: Windows NT"
        End Select
        Ver = "OS Version: " & SysInfo.OSVersion
        Build = "OS Build: " & SysInfo.OSBuild

'Puting Data in TextFields

txtHost.Text = Host
txtIP.Text = IP
txtOS.Text = OS
txtVer.Text = Ver + " || " + Build

End Sub


