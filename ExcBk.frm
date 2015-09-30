VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mfrmExcBk 
   BackColor       =   &H8000000C&
   Caption         =   "Excursion Booking"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6390
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "ExcBk.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tbRpt 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   4305
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   1111
      ButtonWidth     =   2461
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports --->"
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "sp"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excursion"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remaning Tickets"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "sp"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Booking"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "By Name"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "By Date"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "sp"
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbFrm 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   1111
      ButtonWidth     =   1402
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forms --->"
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "sp"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Main"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "sp"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excursion"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Booking"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calculate"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "sp"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Splash"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFrm 
      Caption         =   "Forms"
      Begin VB.Menu mnuFrmExc 
         Caption         =   "Excrusion"
      End
      Begin VB.Menu mnuFrmBk 
         Caption         =   "Booking"
      End
      Begin VB.Menu mnuFrmCal 
         Caption         =   "Calculation"
      End
   End
   Begin VB.Menu mnuRpt 
      Caption         =   "Reports"
      Begin VB.Menu mnuRptExc 
         Caption         =   "Excrusion"
      End
      Begin VB.Menu mnuRptBk 
         Caption         =   "Booking"
      End
      Begin VB.Menu mnuRptBkName 
         Caption         =   "Booking by Name"
      End
      Begin VB.Menu mnuRptDateExc 
         Caption         =   "Excrusion by Date"
      End
      Begin VB.Menu mnuRptRemTicket 
         Caption         =   "Remaning Ticket"
      End
   End
End
Attribute VB_Name = "mfrmExcBk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuFExit_Click()
  End
End Sub

Private Sub mnuFrmBk_Click()
  frmBk.Show
End Sub

Private Sub mnuFrmCal_Click()
  frmCal.Show
End Sub

Private Sub mnuFrmExc_Click()
  frmExc.Show
End Sub

Private Sub mnuRptBk_Click()
  rptBk.Show
End Sub

Private Sub mnuRptBkName_Click()
  Call ShowRptBkName
End Sub

Private Sub mnuRptDateExc_Click()
  Call ShowRptDateExc
End Sub

Private Sub mnuRptExc_Click()
  rptExc.Show
End Sub

Private Sub mnuRptRemTicket_Click()
  Call ShowRptRemTicket
End Sub

Private Sub tbFrm_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button
    Case "Excursion"
      frmExc.Show
      
    Case "Booking"
      frmBk.Show
      
    Case "Calculate"
      frmCal.Show
      
    Case "Main"
      frmMain.Show
    Case "Splash"
      frmSplash.Show
  End Select
End Sub

Private Sub tbRpt_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button
    Case "Excursion"
      rptExc.Show
      
    Case "Remaning Tickets"
      Call ShowRptRemTicket
      
    Case "Booking"
      rptBk.Show
      
    Case "By Name"
      Call ShowRptBkName
  
    Case "By Date"
      Call ShowRptDateExc
  End Select
End Sub
