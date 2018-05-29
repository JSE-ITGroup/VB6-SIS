VERSION 5.00
Begin VB.Form frmSIS013 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payments Menu"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS013.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   4575
   Begin VB.CommandButton CmdBtn 
      Caption         =   "View Non RTGS payments above threshold"
      Height          =   495
      Index           =   10
      Left            =   480
      TabIndex        =   10
      ToolTipText     =   "Lists Payment Information by Query"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "Update Cheque Inventory"
      Height          =   495
      Index           =   8
      Left            =   480
      TabIndex        =   9
      ToolTipText     =   "Should only be selected at the end of the dividend run"
      Top             =   5040
      Width           =   3615
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "Dividend Chqs Reconciliation Report"
      Height          =   495
      Index           =   6
      Left            =   480
      TabIndex        =   8
      ToolTipText     =   "Post Payments to Bank Reconciliation System"
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "Print Ta&x Cheque"
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   7
      ToolTipText     =   "Lists Payment Information by Query"
      Top             =   3960
      Width           =   3615
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "&List Payments"
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Lists Payment Information by Query"
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "E&xit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1680
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      ToolTipText     =   "Returns to main menu"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "&Post Payments"
      Height          =   375
      Index           =   7
      Left            =   480
      TabIndex        =   5
      ToolTipText     =   "Post Payments to Bank Reconciliation System"
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "&Make Cheques"
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "Prints Payment Cheques"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "Payments &Summary"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Lists Payment Information by Query"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "&Calcualte Payments"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "Calculates Payments based on Entry Information"
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton CmdBtn 
      Caption         =   "Payment &Entry"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Activates the Payment Information Entry Form"
      Top             =   120
      Width           =   3735
   End
   Begin VB.Frame FmeEODRun 
      BackColor       =   &H0080C0FF&
      Caption         =   "End of Dividend Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      TabIndex        =   11
      Top             =   3720
      Width           =   3855
   End
End
Attribute VB_Name = "frmSIS013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim repSISRept As New SISRepts
Dim SpCon As ADODB.Connection

Private Sub cmdBtn_Click(Index As Integer)
' set status msg to wait...
Dim i As Integer

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
frmSIS013.Visible = False
frmSIS014.Show 0 ' Payment Entry
'--
Case 1
frmSIS013.Visible = False
frmSIS015.Show 0 ' Calculate Payments
'--
Case 2 ' print payment summary
repSISRept.ReportNumber = 16
repSISRept.RunShareHolderReport

'--
Case 3  ' Make cheques
frmSIS013.Visible = False
frmSIS018.Show 0

Case 4 ' Print Payment Report
repSISRept.OptNo = 0
repSISRept.ReportNumber = 17
repSISRept.RunShareHolderReport
     
Case 5 ' Print Tax Cheque
     frmSIS013.Visible = False
     FrmPrintTaxChq.Show 0
'--
Case 6 ' Dividend Cheques Reconciliation report
     frmSIS013.Visible = False
     gblOptions = 4
     FrmSelectAccount.Show 0
     
Case 7 'Post Payments
frmSIS013.Visible = False
frmSIS019.Show 0
'--
Case 8
   frmMDI.txtStatusMsg.SimpleText = "Cheque Inventory being updated. Please wait"
   frmMDI.txtStatusMsg.Refresh
   i = RunSP(SpCon, "usp_DividendChqs", 0, gblLoginName)
   If i <> 0 Then
      MsgBox "Error on updating inventory. Update abondoned"
      GoTo Exit_CmdBtn_Click
   Else
      MsgBox "Cheque Inventory updated"
   End If
Case 10
     FrmNonRTGS.Show 0
Case 9
 
 '--
 Unload Me

Case Else
End Select
Exit_CmdBtn_Click:
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Set repSISRept = Nothing

End Sub

Private Sub Form_Activate()
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_AlreadyPosted", 1)
If adoRst.State = adStateClosed Or adoRst.EOF Then
   cmdBtn(5).Enabled = False
   cmdBtn(6).Enabled = False
   cmdBtn(8).Enabled = False
   GoTo Exit_Form_Activate
End If

If adoRst!IsEnabled = "F" Then
   cmdBtn(5).Enabled = False
   cmdBtn(6).Enabled = False
   cmdBtn(8).Enabled = False
Else
   cmdBtn(5).Enabled = True
   cmdBtn(6).Enabled = True
   cmdBtn(8).Enabled = True
   cmdBtn(7).Enabled = False
End If
If adoRst!Posted = 1 Then
   cmdBtn(7).Enabled = False
Else
   cmdBtn(7).Enabled = True
End If
adoRst.Close
Set adoRst = Nothing

Exit_Form_Activate:
Exit Sub

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
   Set SpCon = New ADODB.Connection
   With SpCon
        .ConnectionString = gblFileName
        .CursorLocation = adUseClient
        .ConnectionTimeout = 0
        '.Provider = "SQLOLEDB.1"
   End With
   SpCon.Open , , , adAsyncConnect
   Do While SpCon.State = adStateConnecting
      Screen.MousePointer = vbHourglass
      frmMDI.txtStatusMsg.SimpleText = "Connecting, Please wait......"
      frmMDI.txtStatusMsg.Refresh
   Loop
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
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = vbFormControlMenu Or UnloadMode = 1 Then
        'the X has been clicked or the user has pressed Alt+F4
'        cmdBtn(8) = True
'    End If
'End Sub
