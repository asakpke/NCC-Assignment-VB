VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excursion"
   ClientHeight    =   2805
   ClientLeft      =   4425
   ClientTop       =   2865
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4755
   Begin VB.CommandButton cmdRptBk 
      Caption         =   "--> Detail Report"
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   735
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   232
      TabIndex        =   7
      Top             =   1965
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemTicket 
      Caption         =   "Report Remaning Ticket"
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdDateExc 
      Caption         =   "Report Excrusion by Date"
      Height          =   615
      Left            =   3150
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdRptBkName 
      Caption         =   " Report by Name"
      Height          =   615
      Left            =   3135
      TabIndex        =   4
      Top             =   735
      Width           =   1455
   End
   Begin VB.CommandButton cmdRptExc 
      Caption         =   "--> Detail Report"
      Height          =   615
      Left            =   1695
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   232
      TabIndex        =   2
      Top             =   1350
      Width           =   1455
   End
   Begin VB.CommandButton cmdBk 
      Caption         =   "Booking"
      Height          =   615
      Left            =   232
      TabIndex        =   1
      Top             =   735
      Width           =   1455
   End
   Begin VB.CommandButton cmdExc 
      Caption         =   "Excursion"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBk_Click()
  frmBk.Show
End Sub

Private Sub cmdCal_Click()
  frmCal.Show
End Sub

Private Sub cmdDateExc_Click()
  Call ShowRptDateExc
End Sub

Private Sub cmdExc_Click()
  frmExc.Show
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdRemTicket_Click()
  Call ShowRptRemTicket
End Sub

Private Sub cmdRptBk_Click()
  rptBk.Show
End Sub

Private Sub cmdRptBkName_Click()
  Call ShowRptBkName
End Sub

Private Sub cmdRptExc_Click()
    rptExc.Show
End Sub

Private Sub Form_Load()
  frmMain.Left = mfrmExcBk.ScaleWidth / 2
  frmMain.Top = mfrmExcBk.ScaleHeight / 2
End Sub
