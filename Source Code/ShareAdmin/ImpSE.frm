VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImpSE 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Stock Exchange Payments"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "ImpSE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   300
      Left            =   2520
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   3600
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   4332
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "display Key Field"
      Height          =   372
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   3372
   End
End
Attribute VB_Name = "ImpSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpCon As ADODB.Connection

Private Sub cmdCancel_Click()
Unload Me
Set ImpSE = Nothing
End Sub

Private Sub CmdStart_Click()
On Error GoTo Err_CmdStart_Click

Dim iResp As Integer
Dim sMsg As String
Dim sTitle As String
Dim sErrMsg As String

'--
sMsg = "WARNING: This procedure will will delete existing " & gblHold & " Sub Ledger Records"
sMsg = sMsg & "  then recreate them from the XL Import file. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Building Stock Exchange data"
iResp = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If iResp = vbNo Then
  cmdCancel_Click
  Exit Sub
End If
'--
lbl.Caption = "Clearing existing files"
lbl.Visible = True

sErrMsg = "Procedure failed when trying to activate EXCEL"
lbl.Caption = "Recreating " & gblHold & " SUB Ledger for"

'txtfile = frmMDI.CmnDialog.FileName
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Importing " & gblHold & " data..."
frmMDI.txtStatusMsg.Refresh
iResp = ImportExcel2(frmMDI.CmnDialog.FileName, "ShareBook")
frmMDI.txtStatusMsg.SimpleText = "Importing " & gblHold & " data...Done"
frmMDI.txtStatusMsg.Refresh

'--
sErrMsg = "Procedure failed/check if " & gblHold & " changed the format of the XL Sheet"

'-- display success message
lbl.Caption = ""
ProgressBar1.Visible = False
frmMDI.txtStatusMsg.SimpleText = "Converting " & gblHold & " data..."
frmMDI.txtStatusMsg.Refresh
iResp = RunSP(SpCon, "usp_ImportStockExchangeData", 0, gblReply)

MsgBox "Update successfull. Select Ok to clear this message, then Cancel to end."
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
CmdStart.Enabled = False
SpCon.Close
Exit_CmdStart_Click:
Exit Sub
Err_CmdStart_Click:
MsgBox Err & " " & Err.Description, vbOKOnly
cmdCancel_Click
Exit Sub
End Sub
Private Sub InitProgressBar(max As Long)
  If max = 0 Then Exit Sub
    ProgressBar1.Min = 0
    ProgressBar1.max = max
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.Min

End Sub
Private Sub Form_Load()
csvCenterForm Me, gblMDIFORM
'ready Message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
ProgressBar1.Visible = False
lbl.Caption = ""
lbl.Visible = False
'--

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
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Me.Caption = "Import " & gblHold & " Payments"

FL_Exit:
Exit Sub
End Sub


