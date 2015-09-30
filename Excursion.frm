VERSION 5.00
Begin VB.Form frmExc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excursion"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4260
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   495
      Left            =   2760
      TabIndex        =   24
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2760
      TabIndex        =   23
      Top             =   2595
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2070
      TabIndex        =   22
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Add New"
      Height          =   495
      Left            =   1095
      TabIndex        =   21
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3045
      TabIndex        =   19
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdCpyR 
      Caption         =   "Copy && Return"
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   2100
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   1605
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   495
      Left            =   2760
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdMPrev 
      Caption         =   "Previous"
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   615
      Width           =   1215
   End
   Begin VB.CommandButton cmdMNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   1110
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   615
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   1110
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   1605
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Top             =   2100
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   1440
      TabIndex        =   5
      Top             =   2595
      Width           =   1215
   End
   Begin VB.TextBox txtFld 
      Enabled         =   0   'False
      Height          =   495
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   3090
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Standard Price"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   3090
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Maximum Tickets"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2595
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Date of Excursion"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Time Return"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1605
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Time Depart"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Excursion Name"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   615
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Excursion ID"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmExc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oExc As cDb
Dim LV As Integer

Private Sub cmdCancel_Click()
  oExc.Mode = ""
  
  Call SetTxt
  Call SetBtn
  Call GetFld
End Sub

Private Sub cmdCpyR_Click()
  frmBk.txtFld(8) = txtFld(0)
  
  Call frmBk.cmdEdit_Click
  frmBk.cmdSave.SetFocus
  
  Unload Me
End Sub

Private Sub cmdDel_Click()
  oExc.Del
  Call GetFld
End Sub

Private Sub cmdEdit_Click()
  oExc.Mode = "Edit"
   
  Call SetTxt
   
  Call SetBtn
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  oExc.MF
  Call GetFld
End Sub

Private Sub cmdLast_Click()
  oExc.ML
  Call GetFld
End Sub

Private Sub cmdMNext_Click()
  oExc.MNext
  Call GetFld
End Sub

Private Sub cmdMPrev_Click()
  oExc.MPrevious
  Call GetFld
End Sub

Private Sub cmdNew_Click()
  For LV = 0 To 6
    txtFld(LV) = ""
  Next LV
  Call SetTxt
  Call SetBtn
  
  oExc.Mode = "AddNew"
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err
  Select Case oExc.Mode
    Case "Edit"
      If oExc.Edit Then
        For LV = 0 To 6
          oExc.SetFld LV, txtFld(LV)
        Next LV
        oExc.Update
      End If
    Case "AddNew"
      If oExc.AddNew Then
        For LV = 0 To 6
          oExc.SetFld LV, txtFld(LV)
        Next LV
        oExc.Update
      End If
  End Select
  
  Call SetTxt
  Call SetBtn
  Call GetFld
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Private Sub Form_Load()
  Set oExc = New cDb
  oExc.OpenDB ("Excursion")
  Call GetFld
End Sub

Public Sub GetFld()
On Error GoTo Err
  Dim V As Integer
  For V = 0 To 6
    txtFld(V) = oExc.GetFld(V)
  Next V
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Private Sub SetTxt()
Dim V As Integer
  For V = 0 To 6
    txtFld(V).Enabled = Not txtFld(V).Enabled
  Next V
End Sub

Public Sub SetBtn()
  Dim C As Control
  For Each C In Controls
    If TypeOf C Is CommandButton Then
      C.Enabled = Not C.Enabled
    End If
  Next C
End Sub

