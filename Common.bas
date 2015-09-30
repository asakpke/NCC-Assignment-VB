Attribute VB_Name = "mCommon"
Option Explicit

Public Sub ShowRptDateExc()
On Error GoTo Err_ShowRptDateExc

  Dim sExc As String
  Dim sDate As String
  sExc = InputBox("Enter Excursion", "Excursion?", "Zoo")
  If sExc <> "" Then
    sDate = InputBox("Enter Date", "Date?", "9/27/2003")
    If IsDate(sDate) Then
      DE.DateExc sExc, CDate(sDate)
      rptDateExc.Show
    Else
      MsgBox "Not a valid date"
    End If
  Else
    MsgBox "Not a valid Excursion"
  End If
  Exit Sub
  
Err_ShowRptDateExc:
  MsgBox Err.Description
End Sub

Public Sub ShowRptBkName()
On Error GoTo Err_ShowRptBkName
  Dim sN As String
  sN = InputBox("Enter name", "Name?", "Khurram")
  If sN <> "" Then
    DE.BkName (sN)
    rptBkName.Show
  End If
  Exit Sub
  
Err_ShowRptBkName:
  MsgBox Err.Description
End Sub

Public Sub ShowRptRemTicket()
On Error GoTo Err_ShowRptRemTicket
  Dim sExc As String
  Dim sDate As String
  sExc = InputBox("Enter Excursion", "Excursion?", "Zoo")
  If sExc <> "" Then
    sDate = InputBox("Enter Date", "Date?", "9/28/2003")
    If IsDate(sDate) Then
      DE.RemTicket sExc, CDate(sDate)
      rptRemTicket.Show
    Else
      MsgBox "Not a valid date"
    End If
  Else
    MsgBox "Not a valid Excursion"
  End If
  Exit Sub
  
Err_ShowRptRemTicket:
  MsgBox Err.Description
End Sub



