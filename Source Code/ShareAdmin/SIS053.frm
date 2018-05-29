VERSION 5.00
Begin VB.Form frmSIS053 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate Production Menu"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS053.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5550
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&5. Brokers Register Report     "
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   7
      ToolTipText     =   "Prints Brokers Register"
      Top             =   1560
      Width           =   2200
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&7.    Stockholder Verification"
      Height          =   375
      Index           =   7
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Checks that control Shares matches active certificates"
      Top             =   2160
      Width           =   2200
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&6.  Ownership Report    "
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Prints the Ownership Report"
      Top             =   2160
      Width           =   2200
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&4.Brokers Summary       "
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Prints Brokers Summary report"
      Top             =   1560
      Width           =   2200
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4080
      MaskColor       =   &H000000FF&
      TabIndex        =   4
      ToolTipText     =   "Returns to main menu"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&8.Close Cycle"
      Height          =   375
      Index           =   8
      Left            =   1560
      TabIndex        =   8
      ToolTipText     =   "Clears the Register Workfiles and Reset the System Workfiles."
      Top             =   2760
      Width           =   2200
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&3.         Certificates               "
      Height          =   375
      Index           =   3
      Left            =   2900
      TabIndex        =   2
      ToolTipText     =   "Prints new certificates to be issued."
      Top             =   960
      Width           =   2200
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&2. Certificate Register   "
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Prints The periods certificate register."
      Top             =   960
      Width           =   2200
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&1. Audit Trail Report     "
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Prints the System Audit Trail Report with optional Purging"
      Top             =   360
      Width           =   2200
   End
End
Attribute VB_Name = "frmSIS053"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCmp As New ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim repSISRept As New SISRepts
Private Sub cmdBtn_Click(Index As Integer)
' set status msg to wait...
Dim X As Integer
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'--
Select Case Index
Case 1 ' Print audit Trail Report
'Setup Call to SISRepts
'----------------------
repSISRept.ReportType = 9
repSISRept.ReportNumber = 55
repSISRept.RunShareHolderReport


'gblOptions = 55
'frmReportEngine.Show 0
'--
Case 2 'Print Certificate Register
'Setup Call to SISRepts
'----------------------
repSISRept.ReportType = 9
repSISRept.ReportNumber = 9
repSISRept.RunShareHolderReport

'--
Case 3 ' print certificates
'frmSIS056.Show 0
'rsCmp.Open "Company", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
'Setup Call to SISRepts
'----------------------
'On Error Resume Next
repSISRept.ReportType = 1
repSISRept.ReportNumber = 0
repSISRept.RunShareHolderReport
'rsCmp.Close
'--
Case 4 ' Print Brokers summary Reports
'gblOptions = 0
'frmReportEngine.Show 0 ' print Brokers summary
'Setup Call to SISRepts
'----------------------
repSISRept.ReportType = 9
repSISRept.ReportNumber = 0
repSISRept.RunShareHolderReport
'--
Case 5 ' Print Brokers register
'gblOptions = 1
'frmReportEngine.Show 0
'Setup Call to SISRepts
'----------------------
repSISRept.ReportType = 9
repSISRept.ReportNumber = 1
repSISRept.RunShareHolderReport
'--
On Error GoTo Case5_Err

Set SpCon = New ADODB.Connection
With SpCon
     .ConnectionString = gblFileName
     .CursorLocation = adUseServer
     .ConnectionTimeout = 0
     '.Provider = "SQLOLEDB.1"
End With
SpCon.Open , , , adAsyncConnect
Do While SpCon.State = adStateConnecting
   Screen.MousePointer = vbHourglass
   frmMDI.txtStatusMsg.SimpleText = "Connecting, Please wait......"
   frmMDI.txtStatusMsg.Refresh
Loop
Screen.MousePointer = vbDefault

X = RunSP(SpCon, "usp_IndicatorUpd", 0, 4)

'--
Case 6 ' Print Percentage Ownership Exception report
'gblOptions = 2
'frmReportEngine.Show 0
'Setup Call to SISRepts
'----------------------

repSISRept.ReportType = 9
repSISRept.ReportNumber = 2
repSISRept.RunShareHolderReport
'--
Case 7 ' Verification
frmSIS053.Visible = False
frmSIS057.Show 0
'--
Case 8 ' Close Certificate Production Cycle
frmSIS053.Visible = False
frmSIS058.Show 0
'--
Case 0
  '--
 Unload Me
Case Else
End Select
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Exit Sub
Case5_Err:
 MsgBox "SIS053"
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
 Set repSISRept = New SISRepts
 repSISRept.LoginId = gblFileName
 repSISRept.siteid = Trim(gblSiteId)
 'repSISRept.DSN = gblDSN
   ' ready message
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
Set rsCmp = Nothing
Set repSISRept = Nothing

End Sub
