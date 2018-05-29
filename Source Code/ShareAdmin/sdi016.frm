VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSDI016 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Error Messages"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "sdi016.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4770
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
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
End
Attribute VB_Name = "frmSDI016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As ADODB.Recordset
Dim iRecs As Long, iRecsAdded As Long
Dim AppExcl As Excel.Application
Dim X As Integer, y As Integer, curCell As Object, nextCol As Object
Dim nextCell As Object, txtfile As String
Dim qSQL As String, OpenErr As Integer

Private Sub cmdCancel_Click()
Unload Me
Set frmSDI016 = Nothing
End Sub

Private Sub cmdStart_Click()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
Dim iAcnt As Long
Dim sERRDES As String, sAlert As String, sERRDES2 As String
sErrMsgL1 = "Update failed while erasing error messages  "
sErrMsgL2 = " Note this error."
Set rsMain = New ADODB.Recordset
'--
sMsg = "WARNING: This procedure will will delete existing Error Messages"
sMsg = sMsg & "  then recreate them from the XL Import file. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Re-Building Error Messages"
iResp = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If iResp = vbNo Then
  cmdCancel_Click
  Exit Sub
End If
'--

lbl.Caption = "Clearing existing file"
lbl.Visible = True
qSQL = "DELETE From ERRMSG "
If csvADODML(qSQL, cnn) = False Then
  sErrMsg = sErrMsgL1 & "USERS." & sErrMsgL2
  GoTo SDI016_Fail
  Exit Sub
End If
'-- Open sub Ledger
On Error GoTo Open_err
cnn.Open
rsMain.Open "ERRMSG", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
'--
On Error GoTo SDI016_Fail
sErrMsg = "Procedure failed when trying to activate EXCEL"
Set AppExcl = CreateObject("Excel.application")
lbl.Caption = "Recreating Error messages for"

txtfile = frmMDI.cmnDialog.filename
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Recreating Error Messages..."
frmMDI.txtStatusMsg.Refresh
'--
sErrMsg = "Procedure failed. Call or check Support"
With AppExcl
  .Workbooks.Open (txtfile)
  Set curCell = .Worksheets(1).Range("A2")
  Do While Not IsEmpty(curCell)
    iRecs = iRecs + 1
    Set nextCell = curCell.Offset(1, 0)
    Set curCell = nextCell
  Loop
  InitProgressBar (iRecs)
  iRecs = 0: iRecsAdded = 0
  ProgressBar1.Visible = True
  '--Updating subledger routine
  '--
  cnn.BeginTrans
  Set curCell = .Worksheets(1).Range("A2")
  Do While Not IsEmpty(curCell)
    sErrMsg = "Procedure failed while formatting"
    Set nextCol = curCell.Offset(0, 0) '- error number
    iAcnt = nextCol.Value
    lbl.Caption = "Recreating Error Message for " & iAcnt
    lbl.Refresh
    Set nextCol = curCell.Offset(0, 1) '- ERRDES
    sERRDES = nextCol.Value
    Set nextCol = curCell.Offset(0, 2) '- Alert
    sAlert = nextCol.Value
    Set nextCol = curCell.Offset(0, 3) '- ERRDES2
    sERRDES2 = nextCol.Value
    sErrMsg = "Procedure failed in writing to ERRMSG"
    rsMain.AddNew
    rsMain!errcde = iAcnt
    rsMain!errdes = sERRDES
    rsMain!alert = sAlert
    rsMain!errdes2 = sERRDES2
    rsMain.Update
    iRecsAdded = iRecsAdded + 1  'count records added
    If iRecsAdded = 500 Then
       cnn.CommitTrans
       cnn.BeginTrans
       iRecsAdded = 0
    End If
    iRecs = iRecs + 1
    ProgressBar1.Value = iRecs
    Set nextCell = curCell.Offset(1, 0)
    Set curCell = nextCell
  Loop
  If iRecsAdded > 0 Then
       cnn.CommitTrans
  End If
  sErrMsg = "Procedure failed in closing EXCEL spread sheet"
  .ActiveWorkbook.Save
 .Workbooks.Close
 AppExcl.Quit
End With
rsMain.Close
'--
Set rsMain = Nothing
cnn.Close
cmdStart.Enabled = False
'-- display success message
lbl.Caption = ""
ProgressBar1.Visible = False
MsgBox "Update successfull. Select Ok to clear this message, then Cancel to end."
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Exit Sub
Open_err:
  MsgBox "Fail to Open Existing Error Messages; update aborting. "
 
  Exit Sub
SDI016_Fail:
   MsgBox sErrMsg
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

Private Sub Form_Activate()
If OpenErr = True Then
   '''cnn.close
   '''set cnn = nothing
   Unload Me
   Exit Sub
End If
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
'''Set cnn = New ADODB.Connection
'''cnn.Open cnn
OpenErr = 0
Set rsMain = New ADODB.Recordset
On Error GoTo FL_ERR
cnn.Open
rsMain.Open "ERRMSG", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
rsMain.Close
cnn.Close
FL_Exit:
Exit Sub
FL_ERR:
   MsgBox "SDI016/Form_Load"
   OpenErr = True
   Resume FL_Exit
End Sub
