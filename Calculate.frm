VERSION 5.00
Begin VB.Form frmCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6960
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1575
      TabIndex        =   36
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtStrength 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   35
      Text            =   "0"
      Top             =   1410
      Width           =   1215
   End
   Begin VB.TextBox txtStrength 
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   34
      Text            =   "0"
      Top             =   2025
      Width           =   1215
   End
   Begin VB.TextBox txtSPrice 
      Height          =   495
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "0"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtDiscount 
      Height          =   495
      Index           =   0
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1425
      Width           =   1215
   End
   Begin VB.TextBox txtDiscount 
      Height          =   495
      Index           =   1
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   2025
      Width           =   1215
   End
   Begin VB.TextBox txtStrength 
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   30
      Text            =   "0"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtDiscount 
      Height          =   495
      Index           =   2
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2625
      Width           =   1215
   End
   Begin VB.TextBox txtDiscount 
      Height          =   495
      Index           =   3
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3210
      Width           =   1215
   End
   Begin VB.TextBox txtStrength 
      Height          =   495
      Index           =   3
      Left            =   1455
      TabIndex        =   27
      Text            =   "0"
      Top             =   3225
      Width           =   1215
   End
   Begin VB.TextBox txtStrength 
      Height          =   495
      Index           =   4
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0"
      Top             =   3825
      Width           =   1215
   End
   Begin VB.TextBox txtDiscount 
      Height          =   495
      Index           =   4
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3825
      Width           =   1215
   End
   Begin VB.CommandButton cmdCpyR 
      Caption         =   "Copy && Return"
      Height          =   495
      Left            =   255
      TabIndex        =   24
      Top             =   4560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   2430
      TabIndex        =   23
      Top             =   4560
      Width           =   975
   End
   Begin VB.ComboBox cboExc 
      Height          =   315
      ItemData        =   "Calculate.frx":0000
      Left            =   480
      List            =   "Calculate.frx":0002
      TabIndex        =   22
      Text            =   "Zoo"
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtTCalPrice 
      Height          =   495
      Left            =   5415
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   4575
      Width           =   1215
   End
   Begin VB.TextBox txtTPrice 
      Height          =   495
      Index           =   4
      Left            =   5415
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3855
      Width           =   1215
   End
   Begin VB.TextBox txtDPrice 
      Height          =   495
      Index           =   4
      Left            =   4095
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3855
      Width           =   1215
   End
   Begin VB.TextBox txtDPrice 
      Height          =   495
      Index           =   3
      Left            =   4095
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtTPrice 
      Height          =   495
      Index           =   3
      Left            =   5415
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3255
      Width           =   1215
   End
   Begin VB.TextBox txtTPrice 
      Height          =   495
      Index           =   2
      Left            =   5415
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2655
      Width           =   1215
   End
   Begin VB.TextBox txtTPrice 
      Height          =   495
      Index           =   1
      Left            =   5415
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2055
      Width           =   1215
   End
   Begin VB.TextBox txtTPrice 
      Height          =   495
      Index           =   0
      Left            =   5415
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1455
      Width           =   1215
   End
   Begin VB.TextBox txtDPrice 
      Height          =   495
      Index           =   2
      Left            =   4095
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2655
      Width           =   1215
   End
   Begin VB.TextBox txtDPrice 
      Height          =   495
      Index           =   1
      Left            =   4095
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2055
      Width           =   1215
   End
   Begin VB.TextBox txtDPrice 
      Height          =   495
      Index           =   0
      Left            =   4095
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1455
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Total Calculated Price"
      Height          =   495
      Left            =   3735
      TabIndex        =   20
      Top             =   4575
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Total Persons"
      Height          =   495
      Left            =   135
      TabIndex        =   17
      Top             =   3855
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "No. of Young"
      Height          =   495
      Left            =   135
      TabIndex        =   14
      Top             =   3255
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Total"
      Height          =   495
      Left            =   5415
      TabIndex        =   10
      Top             =   855
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "No. of  Older"
      Height          =   495
      Left            =   135
      TabIndex        =   6
      Top             =   2655
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "No. of Student"
      Height          =   495
      Left            =   135
      TabIndex        =   5
      Top             =   2055
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Discounted Ticket Price"
      Height          =   495
      Left            =   4095
      TabIndex        =   4
      Top             =   855
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Discount %"
      Height          =   495
      Left            =   2775
      TabIndex        =   3
      Top             =   855
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "No. of Children"
      Height          =   495
      Left            =   135
      TabIndex        =   2
      Top             =   1455
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Standard Ticket Price"
      Height          =   495
      Left            =   2820
      TabIndex        =   1
      Top             =   150
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Strength"
      Height          =   495
      Left            =   1455
      TabIndex        =   0
      Top             =   855
      Width           =   1215
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oBk As cDb
Dim oDis As cDb
Dim LV As Integer

Private Sub cboExc_Click()
  If oBk.Find("Description='", cboExc.Text & "'") Then
    txtSPrice = oBk.GetFld(6)
    'jsdfsdjkfjsdkfjsdfkjsdkf
    Call Calc
  End If
End Sub

Private Sub cmdCalculate_Click()
  Call Calc
End Sub

Private Sub cmdCpyR_Click()
On Error GoTo Err
  ' Copy Data back to frmBk
  For LV = 0 To 3
    frmBk.txtFld(LV + 3) = txtStrength(LV)
  Next LV
  frmBk.txtFld(7) = txtTCalPrice
  If oBk.Find("Excursion='", cboExc.Text & "'") Then
    frmBk.txtFld(8) = oBk.GetFld(0)
  End If
  
  Call frmBk.cmdEdit_Click
  frmBk.cmdSave.SetFocus
  
  Unload Me
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Err
  Set oBk = New cDb
  Set oDis = New cDb
  
  oBk.OpenDB ("Excursion")
  oDis.OpenDB ("Discount")
  
  Dim V As Integer
  
  oBk.ShowMsg = False
  For V = 1 To oBk.RecCount
    Dim L As Long
    cboExc.AddItem oBk.GetFld(1)
    oBk.MNext
  Next V
  oBk.ShowMsg = True
  
  oDis.ShowMsg = False
  For LV = 0 To oDis.RecCount - 1
    txtDiscount(LV) = oDis.GetFld(2)
    oDis.MNext
  Next LV
  oDis.ShowMsg = True
  
  If oBk.Find("Excursion='", cboExc.Text & "'") Then
    txtSPrice = oBk.GetFld(6)
  End If
  
  Call Calc
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Private Sub DPrice(Index As Integer)
On Error GoTo Err
  If txtDiscount(Index) = 0 Then
    If Index = 4 Then
      txtDPrice(Index) = 0
    Else
      txtDPrice(Index) = txtSPrice
    End If
  Else
  ' checking cint( a * b / c)
    txtDPrice(Index) = CInt(CInt(txtSPrice) _
      * CInt(txtDiscount(Index)) / 100)
  End If
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Private Sub Strength(Index As Integer)
On Error GoTo Err
  If Index = 4 Then
    Dim V As Integer
    txtStrength(4) = 0
    For V = 0 To 3
      txtStrength(4) = CInt(txtStrength(4)) _
        + CInt(txtStrength(V))
    Next V
    
    If CInt(txtStrength(4)) > 20 Then
      txtDiscount(4) = 25
    Else
      txtDiscount(4) = 0
    End If
  End If
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Private Sub TCalPrice()
On Error GoTo Err
  Dim V As Integer
  txtTCalPrice = 0
  For V = 0 To 3
    txtTCalPrice = CInt(txtTCalPrice) + CInt(txtTPrice(V))
  Next V
  
  txtTCalPrice = CInt(txtTCalPrice) - CInt(txtTPrice(4))
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Private Sub TPrice(Index As Integer)
On Error GoTo Err
  txtTPrice(Index) = txtDPrice(Index) * txtStrength(Index)
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Public Sub Calc()
  Call Strength(4)
  
  Dim V As Integer
  For V = 0 To 4
    Call DPrice(V)
    
    Call TPrice(V)
  Next V
    
  Call TCalPrice
End Sub
