VERSION 5.00
Begin VB.Form frmSIS100 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rights Issue Menu"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS100.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Print RI &Application "
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Print Rights Issue Application Forms"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Print RI &Offer Letter"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "Prints Rights Issue Allotment Offer letters"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&JCSD Mail Merge"
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   4
      ToolTipText     =   "Create Mail Merge File from JCSD Ledger"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   375
      Index           =   8
      Left            =   3240
      MaskColor       =   &H000000FF&
      TabIndex        =   7
      ToolTipText     =   "Returns to main menu"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Main  Mail Merge"
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "Create Mial Merge File from Main Ledger"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Issue Allocation Report"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Lists Rights Issue Allocations with fractions"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Calcualte Allocations"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Calculates Rights Issue based on Entry Information"
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Rights Issue &Entry"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Activates the Rights Issue Information Entry Form"
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmSIS100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim repSISRept As New SISRepts
Private Sub cmdBtn_Click(Index As Integer)
' set status msg to wait...
Screen.MousePointer = vbHourglass
' set common SIS print properties
'--------------------------------
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
Set repSISRept = New SISRepts
'repSISRept.DSN = gblDSN
repSISRept.LoginId = gblFileName
repSISRept.siteid = gblSiteId
repSISRept.ReportType = 9
'--
Select Case Index
Case 0
frmSIS101.Show 0 ' Rights Issue Entry
'--
Case 1
frmSIS102.Show 0 ' Calculate Payments
'--
Case 2 ' print allocation Report
repSISRept.ReportNumber = 103
repSISRept.RunShareHolderReport
'--
Case 3  ' Create Mail Merge File from Main Legder
  Call CreateRIOfferLetter(1)
Case 4 ' Create Mail merge File from JCSD Ledger
  Call CreateRIOfferLetter(2)
Case 5 ' Print Allotment Offer Letter
   Call PrintRIOLetter(1)
Case 6 ' Print Application Form
   Call PrintRIOLetter(2)
'--
Case 8
  cnnClose
  Unload Me
Case Else
End Select
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
End Sub
Private Sub Form_Load()
 
 csvCenterForm Me, gblMDIFORM
 '--  disable menu items
 '----------------------
 frmMDI.mnuFile.Enabled = False
 frmMDI.btnClose.Enabled = False
 frmMDI.mnuLists.Enabled = False
 frmMDI.mnuAct.Enabled = False
 frmMDI.mnuAdm.Enabled = False
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmMDI.mnuFile.Enabled = True
 frmMDI.btnClose.Enabled = True
 frmMDI.mnuLists.Enabled = True
 frmMDI.mnuAct.Enabled = True
 If gblUserLevel = 1 Then frmMDI.mnuAdm.Enabled = True
 cnnClose
End Sub
