VERSION 5.00
Begin VB.Form frmSIS013I 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interest Payment Menu"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS013I.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5190
   Begin VB.CommandButton CmdBtn 
      Caption         =   "&List Interest Payments"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Lists Payment Information by Query"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   375
      Index           =   6
      Left            =   1680
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      ToolTipText     =   "Returns to main menu"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "&Post Interest Payments"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "Post Payments to Bank Reconciliation System"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "&Make Cheques"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Prints Payment Cheques"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "Interests &Summary"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Lists Payment Information by Query"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "&Calcualte Interests"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Calculates Payments based on Entry Information"
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "Interest Payment &Entry"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Activates the Payment Information Entry Form"
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmSIS013I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim repSISRept As New SISRepts
Private Sub cmdBtn_Click(Index As Integer)
' set status msg to wait...

Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
Set repSISRept = New SISRepts
repSISRept.ReportType = 9
repSISRept.LoginId = gblFileName
repSISRept.siteid = gblSiteId

'--
Select Case Index
Case 0
frmSIS013I.Visible = False
frmSIS014I.Show 0 ' Payment Entry
'--
Case 1
frmSIS013I.Visible = False
frmSIS015I.Show 0 ' Calculate Payments
'--
Case 2 ' Print Interest Payment Summary
repSISRept.ReportNumber = 161
repSISRept.RunShareHolderReport

'--
Case 3  ' Print cheques
frmSIS013I.Visible = False
frmSIS018I.Show 0

Case 4 ' Print Interest Payment Detail Report
repSISRept.ReportNumber = 174
repSISRept.RunShareHolderReport
           
'--
Case 5 'Post Payments
frmSIS013I.Visible = False
frmSIS019I.Show 0
'--
Case 6
 '--
 Unload Me
Case Else
End Select
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Set repSISRept = Nothing

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

End Sub
