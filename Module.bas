Attribute VB_Name = "Module"
Option Explicit
Public NewTelephone As String
Public sno, user, password, winuser, logindate, logintime, logouttime, UserType, IP, Host, OS, Ver, Build
Public UserTypeUsing As String
Public Starting As Boolean

'Combo autoComplete
' SendMessage API
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'CB Constants
Public Const CB_MAXLENGTH = 50
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_LIMITTEXT = &H141
'............................................................

Public ErrorMsg As String 'Error Loging

'Mail
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'***********************************************************************
'Database
'***********************************************************************
Public cn As New ADODB.Connection

Sub Main()

    cn.Open ("Provider=MSDASQL.1;Persist Security Info=False;Data Source=Noori_Halal")
    frmSplash.Show

End Sub

'Combo
Public Sub Combo_Lookup(ctlCombo As ComboBox)
   Dim lngItemPos As Long
   Dim strCombo As String

   strCombo = ctlCombo.Text

   ' Use SendMessage() API to Find Combobox Values
   lngItemPos = SendMessage(ctlCombo.hwnd, CB_FINDSTRING, -1, ByVal strCombo)

   If lngItemPos >= 0 Then
      ctlCombo.ListIndex = lngItemPos
   End If

   ctlCombo.SelStart = Len(strCombo)
   ctlCombo.SelLength = CB_MAXLENGTH
End Sub
Public Sub SendErrorReport()
'Mailing the Error
    Dim Error As String
    Error = Err.Description
    frmMail.txtBody.Text = "Please Specify on which form this Error Came and when it came so that is should be debug. Error Message ( " + Error + " )"
    frmMail.txtSubject.Text = "Error Report"
    frmMail.txtToName.Text = "The Developers"
    frmMail.txtToAddress.Text = "adeel_s90@hotmail.com"
    frmMail.Show
End Sub
'*****************************************************************************

Public Sub ConnectItems()
'Connecting Database with ADODC1
On Error GoTo ConnectionError
    frmItem.Adodc1.ConnectionString = cn
    frmItem.Adodc1.CursorLocation = adUseClient
    frmItem.Adodc1.CursorType = adOpenDynamic
    frmItem.Adodc1.RecordSource = "select * from Item order by Item_Code;"
    Set frmItem.DataGrid1.DataSource = frmItem.Adodc1
    
    If (frmItem.Adodc1.Recordset.BOF) Then
        Exit Sub
    End If
    frmItem.Adodc1.Refresh
    Exit Sub

ConnectionError:
    MsgBox "Unable to Connect", vbCritical, ":: | :: ADMIN :: | :."
    Exit Sub
End Sub
Public Sub GetItemData()
    If (frmItem.Adodc1.Recordset.BOF) Then
        Exit Sub
    End If

    frmItem.txtIC.Text = frmItem.Adodc1.Recordset.Fields("Item_Code")
    frmItem.IType.Text = frmItem.Adodc1.Recordset.Fields("GRp")
    frmItem.Item.Text = frmItem.Adodc1.Recordset.Fields("Item")
    frmItem.txtUP.Text = frmItem.Adodc1.Recordset.Fields("Price")
    frmItem.txtR.Text = frmItem.Adodc1.Recordset.Fields("Remarks")
    frmItem.txtDate.Text = frmItem.Adodc1.Recordset.Fields("Date")
End Sub

Public Sub ConnectDrivers()
'Connecting Database with ADODC1
On Error GoTo ConnectionError
    frmDriver.Adodc1.ConnectionString = cn
    frmDriver.Adodc1.CursorLocation = adUseClient
    frmDriver.Adodc1.CursorType = adOpenDynamic
    frmDriver.Adodc1.RecordSource = "select * from Item order by Item_Code;"
    Set frmDriver.DataGrid1.DataSource = frmDriver.Adodc1
    
    If (frmDriver.Adodc1.Recordset.BOF) Then
        Exit Sub
    End If
    frmDriver.Adodc1.Refresh
    Exit Sub

ConnectionError:
    MsgBox "Unable to Connect", vbCritical, ":: | :: ADMIN :: | :."
    Exit Sub
End Sub
Public Sub GetDriverData()
    If (frmDriver.Adodc1.Recordset.BOF) Then
        Exit Sub
    End If

    frmDriver.txtDID.Text = frmDriver.Adodc1.Recordset.Fields("DID")
    frmDriver.txtName.Text = frmDriver.Adodc1.Recordset.Fields("Name")
    frmDriver.txtAdd.Text = frmDriver.Adodc1.Recordset.Fields("Address")
    frmDriver.txtTelephone.Text = frmDriver.Adodc1.Recordset.Fields("Telephone")
    frmDriver.txtMobile.Text = frmDriver.Adodc1.Recordset.Fields("Mobile")
    frmDriver.txtCash.Text = frmDriver.Adodc1.Recordset.Fields("Cash")
    frmDriver.txtDCash.Text = frmDriver.Adodc1.Recordset.Fields("D_Cash")
End Sub
Public Sub ConnectCustomers()
'Connecting Database with ADODC1
On Error GoTo ConnectionError
    frmCustomer.Adodc1.ConnectionString = cn
    frmCustomer.Adodc1.CursorLocation = adUseClient
    frmCustomer.Adodc1.CursorType = adOpenDynamic
    frmCustomer.Adodc1.RecordSource = "select * from Customer order by Name;"
    Set frmCustomer.DataGrid1.DataSource = frmCustomer.Adodc1
    
    If (frmCustomer.Adodc1.Recordset.BOF) Then
        Exit Sub
    Else
        frmCustomer.Adodc1.Refresh
    End If
    
    Exit Sub
ConnectionError:
    MsgBox "Unable to Connect", vbCritical, ":: | :: ADMIN :: | :."
    Exit Sub
End Sub
Public Sub GetCustomerData()
    If (frmCustomer.Adodc1.Recordset.BOF) Then
        Exit Sub
    End If
    
    frmCustomer.txtCID.Text = frmCustomer.Adodc1.Recordset.Fields("CID")
    frmCustomer.txtDate.Text = frmCustomer.Adodc1.Recordset.Fields("Date")
    frmCustomer.txtTelephone.Text = frmCustomer.Adodc1.Recordset.Fields("Telephone")
    frmCustomer.txtMobile.Text = frmCustomer.Adodc1.Recordset.Fields("Mobile")
    frmCustomer.CName.Text = frmCustomer.Adodc1.Recordset.Fields("Name")
    frmCustomer.txtArea.Text = frmCustomer.Adodc1.Recordset.Fields("Area")
    frmCustomer.txtPC.Text = frmCustomer.Adodc1.Recordset.Fields("Post_Code")
    frmCustomer.txtR.Text = frmCustomer.Adodc1.Recordset.Fields("Remarks")
End Sub

Public Sub ConnectOrders()
'Connecting Database with ADODC1
On Error GoTo ConnectionError
    frmOrder.Adodc1.ConnectionString = cn
    frmOrder.Adodc1.CursorLocation = adUseClient
    frmOrder.Adodc1.CursorType = adOpenDynamic
    frmOrder.Adodc1.RecordSource = "select * from Ord order by Date;"
    Set frmOrder.DataGrid1.DataSource = frmOrder.Adodc1
    
    frmOrder.Adodc5.ConnectionString = cn
    frmOrder.Adodc5.CursorLocation = adUseClient
    frmOrder.Adodc5.CursorType = adOpenDynamic
    frmOrder.Adodc5.RecordSource = "SELECT Date, CID, OID, P_Mode, Item, Price FROM Ord ORDER BY Date;"
    frmOrder.Adodc5.Refresh
    Set frmOrder.DataGrid5.DataSource = frmOrder.Adodc5
    
    If (frmOrder.Adodc1.Recordset.BOF Or frmOrder.Adodc5.Recordset.BOF) Then
        Exit Sub
    Else
        frmOrder.Adodc1.Refresh
        frmOrder.Adodc5.Refresh
    End If
    
    Exit Sub

ConnectionError:
    MsgBox "Unable to Connect", vbCritical, ":: | :: ADMIN :: | :."
    Exit Sub
End Sub
Public Sub GetOrderData()
    If (frmOrder.Adodc1.Recordset.BOF) Then
        Exit Sub
    End If

    frmOrder.txtCID.Text = frmOrder.Adodc1.Recordset.Fields("CID")
    frmOrder.txtOid.Text = frmOrder.Adodc1.Recordset.Fields("OID")
    frmOrder.txtDate.Text = frmOrder.Adodc1.Recordset.Fields("Date")
    frmOrder.txtGroup.Text = frmOrder.Adodc1.Recordset.Fields("Grp")
    frmOrder.txtPM.Text = frmOrder.Adodc1.Recordset.Fields("P_Mode")
    frmOrder.txtItem.Text = frmOrder.Adodc1.Recordset.Fields("Item")
    frmOrder.txtQty.Text = frmOrder.Adodc1.Recordset.Fields("Quantity")
    frmOrder.txtC.Text = frmOrder.Adodc1.Recordset.Fields("Cutting")
    frmOrder.txtP.Text = frmOrder.Adodc1.Recordset.Fields("Packing")
    frmOrder.txtPrice.Caption = frmOrder.Adodc1.Recordset.Fields("Price")
    frmOrder.txtR.Text = frmOrder.Adodc1.Recordset.Fields("Remarks")
End Sub
Public Sub CheckUser()
    'Getting User Type from DB
    frmLogin.Adodc2.Recordset.MoveFirst
    frmLogin.Adodc2.RecordSource = "select * from Login where User='" + user + "'"
    frmLogin.Adodc2.Refresh
    UserTypeUsing = frmLogin.Adodc2.Recordset.Fields(2)
            
    If (UserTypeUsing = "Local") Then
        MDIForm1.mnUM.Enabled = False
        MDIForm1.mnUsage.Enabled = False
        MDIForm1.mnEdit.Enabled = False
        MDIForm1.mnReports.Enabled = False
    Else
    End If
'End If
End Sub
Public Sub CheckWorking()
Dim Working As String
    If (Month(Date) = 6) Then
        Dim i As Integer
        frmLogin.Adodc1.ConnectionString = cn
        frmLogin.Adodc1.CursorLocation = adUseClient
        frmLogin.Adodc1.CursorType = adOpenDynamic
        frmLogin.Adodc1.RecordSource = "SELECT Rem FROM Conn;"
        Set frmLogin.DataGrid1.DataSource = frmLogin.Adodc1
            
        For i = 1 To frmLogin.Adodc1.Recordset.RecordCount
            frmLogin.Adodc1.Recordset.Update "Rem", "Noori_Halal"
            frmLogin.Adodc1.Recordset.Requery
            frmLogin.Adodc1.Refresh
        Next i
        MsgBox "Software Expired !!!", vbCritical, "AdEeL"
        End
        Exit Sub
    End If

    frmLogin.Adodc1.ConnectionString = cn
    frmLogin.Adodc1.CursorLocation = adUseClient
    frmLogin.Adodc1.CursorType = adOpenDynamic
    frmLogin.Adodc1.RecordSource = "SELECT Rem FROM Conn;"
    Set frmLogin.DataGrid1.DataSource = frmLogin.Adodc1
    Working = frmLogin.Adodc1.Recordset.Fields("Rem")
    If (Working = "Noori_Halal") Then
        MsgBox "Software Expired !!!", vbCritical, "AdEeL"
        End
        Exit Sub
    End If
End Sub
Public Sub Secure()
    frmLogin.Adodc1.Recordset.AddNew
    frmLogin.Adodc1.Recordset.Fields("SNo") = sno
    frmLogin.Adodc1.Recordset.Fields("User") = user
    frmLogin.Adodc1.Recordset.Fields("Password") = password
    frmLogin.Adodc1.Recordset.Fields("WinUser") = winuser
    frmLogin.Adodc1.Recordset.Fields("LoginDate") = logindate
    frmLogin.Adodc1.Recordset.Fields("LoginTime") = logintime
    frmLogin.Adodc1.Recordset.Fields("LogoutTime") = Time
    frmLogin.Adodc1.Recordset.Fields("Computer") = Host
    frmLogin.Adodc1.Recordset.Fields("IP") = IP
    frmLogin.Adodc1.Recordset.Fields("OS") = OS
    frmLogin.Adodc1.Recordset.Fields("Version") = Ver
    frmLogin.Adodc1.Recordset.Fields("Build") = Build
        
    frmLogin.Adodc1.Recordset.Update
    frmLogin.Adodc1.Recordset.Requery
    MsgBox "Thank you for using this software", vbInformation, ":: | :: ADMIN :: | :."
End Sub
