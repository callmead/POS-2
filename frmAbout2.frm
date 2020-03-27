VERSION 5.00
Begin VB.Form frmAbout2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  ..........::::::::::  A b o u t _ T h e _ D e v e l o p e r  ::::::::::.........."
   ClientHeight    =   3855
   ClientLeft      =   2955
   ClientTop       =   2490
   ClientWidth     =   8985
   Icon            =   "frmAbout2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout2.frx":08CA
   ScaleHeight     =   3855
   ScaleWidth      =   8985
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   2250
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   480
      Top             =   2250
   End
   Begin VB.Timer Timer3 
      Left            =   960
      Top             =   2250
   End
   Begin VB.Label lblheading 
      BackStyle       =   0  'Transparent
      Caption         =   "About Developer ..."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   210
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   360
      Y1              =   90
      Y2              =   1050
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4200
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   2115
      Left            =   555
      TabIndex        =   1
      Top             =   1050
      Width           =   7950
   End
   Begin VB.Label lblMail 
      BackStyle       =   0  'Transparent
      Caption         =   "Mail me adeel_s90@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   3450
      Width           =   3255
   End
End
Attribute VB_Name = "frmAbout2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'For Label
Dim s, s1 As String
Dim l As Integer
'LoadEffect
Dim fheight As Integer, step As Integer, fwidth As Integer
'For Form layout
Private Declare Function GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
'On Top Decleration
Private Const HWND_TOPMOST = -1

Public Function AddOfficeBorder(ByVal hwnd As Long)
    
    Dim lngRetVal As Long
    'Retrieve the current border style
    lngRetVal = GetWindowLong(hwnd, GWL_EXSTYLE)
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    'Apply the changes
    SetWindowLong hwnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Function

Public Sub changeForm(ByRef frmChng As Object)
    On Error GoTo EH
    Dim objCtrl As Control
    
    frmChng.Appearance = 0
    AddOfficeBorder (frmChng.hwnd)
    frmChng.BackColor = &H80000016
    For Each objCtrl In frmChng.Controls
    
        If Not TypeOf objCtrl Is Label Then
            objCtrl.Appearance = 0
            
            If TypeOf objCtrl Is TextBox Or _
                TypeOf objCtrl Is CommandButton Or _
                TypeOf objCtrl Is ComboBox Then
                                
                AddOfficeBorder (objCtrl.hwnd)
            End If
            If TypeOf objCtrl Is TextBox Or _
                TypeOf objCtrl Is CommandButton Or _
                TypeOf objCtrl Is ComboBox Or _
                TypeOf objCtrl Is CheckBox Or _
                TypeOf objCtrl Is OptionButton Then
                
                objCtrl.BackColor = &H80000016
                objCtrl.BorderStyle = 0
            End If
            
            If TypeOf objCtrl Is CheckBox Or _
                TypeOf objCtrl Is OptionButton Then
                
                objCtrl.BackColor = &H8000000F
            End If
            If TypeOf objCtrl Is Frame Then
                objCtrl.BackColor = &H8000000F
            End If
        End If
    Next
    
EH:
    If Err.Number = 438 Then
        Resume Next
    End If
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Change Form
    Call changeForm(Me)
'For Label
    s = "By the Grace of Almighty Allah I have completed this software. For any Further Changes Please do not hesitate to contact me on adeel_s90@hotmail.com. Your suggestions will be appriciated.                                                                                                                                                     Warning: This software is protected by copyright law and international treaties. Unauthorized reproduction and distribution of this software, or any portion of it, may result in severe civil and criminal penalties, and will be prosecuted to the maximum extend possible under law.                                                        "
    l = 1
'Form Load Effect
    fheight = 4365
    fwidth = 9615
    frmAbout2.Height = 0
    frmAbout2.Width = 0
    step = 50
    Timer3.Interval = 10
'On TOP
    Call SetWindowPos(frmAbout2.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Sub

Private Sub lblMail_Click()
ShellExecute Me.hwnd, vbNullString, _
"mailto:adeel_s90@hotmail.com", _
vbNullString, _
Left$(CurDir, 3), 1
lblMail.ForeColor = vbWhite
lblMail.Font.Underline = False
End Sub

Private Sub lblMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMail.ForeColor = vbWhite
lblMail.Font.Underline = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMail.ForeColor = vbBlue
lblMail.Font.Underline = False
End Sub

Private Sub Timer1_Timer()
frmAbout2.Caption = Right(Trim(frmAbout2.Caption), Len(Trim(frmAbout2.Caption)) - 1) + Trim(Left(frmAbout2.Caption, 1))
frmAbout2.Caption = Left(Trim(frmAbout2.Caption), Len(Trim(frmAbout2.Caption)) - 1) + Trim(Right(frmAbout2.Caption, 1))
End Sub

Private Sub Timer2_Timer()
s1 = Left(s, l)
lblText.Caption = s1
l = l + 1
If l >= Len(s) + 3 Then
    l = 1
    s1 = ""
End If
End Sub

Private Sub Timer3_Timer()
    If step >= fheight Then
        frmAbout2.Height = fheight
    Else
    step = step + 300
    frmAbout2.Height = step
    End If
    
    If step >= fwidth Then
        frmAbout2.Width = fwidth
        Exit Sub
    Else
    step = step + 300
    frmAbout2.Width = step
    End If
End Sub

