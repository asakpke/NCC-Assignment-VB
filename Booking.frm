VERSION 5.00
Begin VB.Form frmBk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Booking"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4410
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3480
      TabIndex        =   29
      Top             =   4800
      Width           =   840
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2880
      TabIndex        =   28
      Top             =   2100
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   27
      Top             =   4800
      Width           =   840
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   4800
      Width           =   840
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Add New"
      Height          =   495
      Left            =   960
      TabIndex        =   25
      Top             =   4800
      Width           =   840
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   24
      Top             =   4800
      Width           =   840
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   8
      Left            =   1560
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdMNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   2880
      TabIndex        =   20
      Top             =   1110
      Width           =   1335
   End
   Begin VB.CommandButton cmdMPrev 
      Caption         =   "Previous"
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   615
      Width           =   1335
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   1605
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "--> Calculate"
      Height          =   495
      Left            =   2880
      TabIndex        =   16
      Top             =   3570
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcDetail 
      Caption         =   "--> Detail"
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Top             =   4065
      Width           =   1335
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   7
      Left            =   1560
      TabIndex        =   7
      Top             =   3585
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   6
      Left            =   1560
      TabIndex        =   6
      Top             =   3090
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   1560
      TabIndex        =   5
      Top             =   2595
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   2100
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   1605
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   1110
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   615
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "No. of Children"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   1605
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "No. of  Student"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "No. of Young"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   3090
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Excursion ID"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Total Price"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   3585
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Number of Older"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   2595
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Booking Date"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Booking Persion Name"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   615
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Booking Number"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oBk As cDb
Dim oExc As cDb

Dim LV As Integer

Private Sub cmdCalc_Click()
On Error GoTo Err
  frmCal.Show
  
  frmCal.cmdCpyR.Visible = True
  
  Dim V As Integer
  For V = 0 To 3
    frmCal.txtStrength(V) = txtFld(V + 3)
  Next V

  If IsNumeric(txtFld(8)) Then
    If oExc.Find("ExcursionID=", txtFld(8)) Then
      frmCal.cboExc.Text = oExc.GetFld(1)
    End If
  End If
  
  frmCal.Calc
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
  oBk.Mode = ""
  
  Call SetTxt
  Call SetBtn
  Call GetFld
  
End Sub

Private Sub cmdDel_Click()
  oBk.Del
  Call GetFld
End Sub

Private Sub cmdExcDetail_Click()
  frmExc.Show
  
  frmExc.cmdCancel.Visible = True
  frmExc.cmdCpyR.Visible = True
  
  If IsNumeric(txtFld(8)) Then
    If frmExc.oExc.Find("ExcursionID=", txtFld(8)) Then
      frmExc.GetFld
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Set oBk = New cDb
  Set oExc = New cDb
  
  oBk.OpenDB ("Booking")
  oExc.OpenDB ("Excursion")
  
  Call GetFld
End Sub

Private Sub GetFld()
On Error GoTo Err
  Dim V As Integer
  For V = 0 To 8
    txtFld(V) = oBk.GetFld(V)
  Next V
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Private Sub cmdFirst_Click()
  oBk.MF
  Call GetFld
End Sub

Private Sub cmdLast_Click()
  oBk.ML
  Call GetFld
End Sub

Private Sub cmdMNext_Click()
  oBk.MNext
  Call GetFld
End Sub

Private Sub cmdMPrev_Click()
  oBk.MPrevious
  Call GetFld
End Sub

Public Sub cmdEdit_Click()
  oBk.Mode = "Edit"
   
  Call SetTxt
   
  Call SetBtn
End Sub

Private Sub cmdNew_Click()
  For LV = 0 To 8
    txtFld(LV) = ""
  Next LV
  Call SetTxt
  
  Call SetBtn
  
  oBk.Mode = "AddNew"
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err
  Select Case oBk.Mode
    Case "Edit"
      If oBk.Edit Then
        For LV = 0 To 8
          oBk.SetFld LV, txtFld(LV)
        Next LV
        oBk.Update
      End If
    Case "AddNew"
      If oBk.AddNew Then
        For LV = 0 To 8
          oBk.SetFld LV, txtFld(LV)
        Next LV
        oBk.Update
      End If
  End Select
  Call SetTxt
  Call GetFld
  Call SetBtn
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Private Sub SetTxt()
  For LV = 0 To 8
    txtFld(LV).Enabled = Not txtFld(LV).Enabled
  Next LV
End Sub

Public Sub SetBtn()
  Dim C As Control
  For Each C In Controls
    If TypeOf C Is CommandButton Then
      C.Enabled = Not C.Enabled
    End If
  Next C
End Sub

