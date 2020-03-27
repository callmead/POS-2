VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "TESTING"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   11385
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    RemoveList1Duplicates
End Sub

Private Sub Command3_Click()
Dim Y, m, d, h As Integer
Dim yy, mm, dd, hh As String

Y = Year(Date)
m = Month(Date)
d = Day(Date)
h = Hour(Time)

yy = Y
mm = m
dd = d
hh = h
Text1.Text = "O"
Text1.Text = yy
End Sub

Private Sub Command4_Click()
MsgBox List1.Text
End Sub

Private Sub Form_Load()
    Adodc1.ConnectionString = CN
    Adodc1.CursorLocation = adUseClient
    Adodc1.CursorType = adOpenDynamic
    Adodc1.RecordSource = "select * from item;"
    Set DataGrid1.DataSource = Adodc1
    
    GetList1Ready
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

GetList2Ready

End Sub

Private Sub GetList1Ready()
'For Item1 and Item Combo
    Dim X As Integer
    For X = 0 To (Adodc1.Recordset.RecordCount - 1)
    
    List1.AddItem Adodc1.Recordset.Fields(1)

    Adodc1.Recordset.MoveNext
    Next X
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

Private Sub GetList2Ready()
    Adodc1.RecordSource = "SELECT ITEM FROM ITEM WHERE Grp='" + List1.Text + "';"
    Adodc1.Refresh
    Dim X As Integer
    For X = 0 To (Adodc1.Recordset.RecordCount - 1)
    List2.AddItem Adodc1.Recordset.Fields(0)
    Adodc1.Recordset.MoveNext
    Next X

    RemoveList2Duplicates

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

