VERSION 5.00
Begin VB.Form frmSIS025 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Transfer  Menu"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS025.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&4. Broker to Stockholder           "
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Activates Broker to Stockholder Transfers"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   375
      Index           =   9
      Left            =   3240
      MaskColor       =   &H000000FF&
      TabIndex        =   4
      ToolTipText     =   "Returns to main menu"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&3. Stockholder to Broker         "
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Activates Stockholder to Broker Transfers"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&2. Broker to Broker Transfer     "
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Activates Broker to Broker Transfers and Certifications"
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&1. Stockholder to Stockholder"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Activates the Stockholder to Stockholder Transfer"
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmSIS025"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBtn_Click(Index As Integer)
' set status msg to wait...
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'--
Select Case Index
Case 0
frmSIS025.Visible = False
frmSIS024.Show 0 ' Stockholder to Stockholder
'--
Case 1
frmSIS025.Visible = False
frmSIS027.Show 0  ' Brokers Certification..
'--
Case 2
frmSIS025.Visible = False
frmSIS031.Show 0 'Stockholder to Broker
'--
Case 3
frmSIS025.Visible = False
frmSIS033.Show 0 ' Broker to Stockholder
'--


Case 9
  
 Unload Me
 Set frmSIS025 = Nothing
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
 frmMDI.mnuAct.Enabled = False
 frmMDI.mnuAdm.Enabled = False
 ' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
'--  enable menu items
 '----------------------
 frmMDI.mnuFile.Enabled = True
 frmMDI.btnClose.Enabled = True
 frmMDI.mnuAct.Enabled = True
 If gblUserLevel = 1 Then frmMDI.mnuAdm.Enabled = True
End Sub
