VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Db As Database
Dim Rs As Recordset
Dim sMode As String
Dim AppPath As String
Dim bShowMsg As Boolean

Public Function GetFld(N As Integer) As String
On Error GoTo Err_GetFld
  If IsNull(Rs(N)) Then
    GetFld = ""
  Else
    GetFld = Rs(N)
  End If
  Exit Function
  
Err_GetFld:
  MsgBox Err.Description
End Function

Public Sub SetFld(N As Integer, S As String)
On Error GoTo Err_SetFld
  Rs(N) = S
  Exit Sub
  
Err_SetFld:
  MsgBox Err.Description
End Sub

Public Sub MNext()
  Rs.MoveNext
  If Rs.EOF Then
    If bShowMsg Then MsgBox "You are at the Last Record"
    Rs.MoveLast
  End If
End Sub

Public Sub MF()
  Rs.MoveFirst
End Sub

Public Sub ML()
  Rs.MoveLast
End Sub

Public Sub MPrevious()
  Rs.MovePrevious
  If Rs.BOF Then
    If bShowMsg Then MsgBox "You are at the First Record"
    Rs.MoveFirst
  End If
End Sub

Public Sub OpenDB(sTbl As String)
On Error GoTo Err_OpenDB
  Set Db = OpenDatabase(AppPath & "\Excursion97.mdb")
  Set Rs = Db.OpenRecordset("SELECT * FROM " & sTbl)
  Rs.MoveLast
  Rs.MoveFirst
  Exit Sub
  
Err_OpenDB:
  MsgBox Err.Description
End Sub

Public Function Find(sFld As String, nVal As String) As Boolean
On Error GoTo Err_Find
  Rs.FindFirst (sFld & nVal)
  If Rs.NoMatch Then
    Find = False
  Else
    Find = True
  End If
  Exit Function
  
Err_Find:
  MsgBox Err.Description
End Function

Public Function RecCount() As Integer
  RecCount = Rs.RecordCount
End Function

Public Property Get Mode() As String
  Mode = sMode
End Property

Public Property Let Mode(ByVal vNewValue As String)
  sMode = vNewValue
End Property

Private Sub Class_Initialize()
  sMode = ""
  ShowMsg = True
  AppPath = App.Path
End Sub

Public Function Edit() As Boolean
  On Error GoTo Err_Edit
  Rs.Edit
  Edit = True
  Exit Function
  
Err_Edit:
  Edit = False
  MsgBox Err.Description
End Function

Public Sub Update()
On Error GoTo Err
  Rs.Update
  Exit Sub
  
Err:
  MsgBox Err.Description
  Rs.Requery
End Sub

Public Sub Del()
On Error GoTo Err
  If MsgBox("Are you sure to Delete?", vbYesNo, "Delete?") _
                                        = vbYes Then
    Rs.Delete
    Rs.Requery
  End If
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Public Function AddNew() As Boolean
On Error GoTo Err_AddNew
  Rs.AddNew
  AddNew = True
  Exit Function
  
Err_AddNew:
  AddNew = False
  MsgBox Err.Description
End Function

Public Property Let ShowMsg(ByVal vNewValue As Boolean)
  bShowMsg = vNewValue
End Property
