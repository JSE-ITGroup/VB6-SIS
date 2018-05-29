VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImpTTSE 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import TTSE Payments"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "ImpTTSE.frx":0000
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
Attribute VB_Name = "ImpTTSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As ADODB.Recordset
Dim iRecs As Long, iRecsAdded As Long
Dim X As Integer, y As Integer, curCell As Object, nextCol As Object
Dim nextCell As Object, txtfile As String
Dim qSQL As String

Private Sub cmdCancel_Click()
Unload Me
Set ImpTTSE = Nothing
End Sub

Private Sub cmdStart_Click()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
Dim iAcnt As Long, iCBL As Long, iRat As Single
Dim sNam As String, sAD1 As String, sAD2 As String, sAd3 As String, sTax As String
Dim sAD4 As String
Dim sAD5 As String
Dim SelString As String, ttseID As String
Dim CntryRst As ADODB.Recordset
Dim StrSql As String


sErrMsgL1 = "Update failed while clearing balances of "
sErrMsgL2 = " Note this error."
Set rsMain = New ADODB.Recordset
Set CntryRst = New ADODB.Recordset
'--
sMsg = "WARNING: This procedure will will delete existing TTSE Sub Ledger Records"
sMsg = sMsg & "  then recreate them from the XL Import file. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Building TTSE Sub Ledger"
iResp = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If iResp = vbNo Then
  cmdCancel_Click
  Exit Sub
End If
'--

lbl.Caption = "Clearing existing file"
lbl.Visible = True
qSQL = "DELETE From TTSESUB "
If csvADODML(qSQL, cnn) = False Then
  sErrMsg = sErrMsgL1 & "USERS." & sErrMsgL2
  GoTo ImpTTSE_Fail
  Exit Sub
End If
'-- Open sub Ledger
On Error GoTo Open_err
cnn.Open
rsMain.Open "TTSESub", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
'rsMnDte.Open "TTSEMndte", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
'--

On Error GoTo ImpTTSE_Fail
sErrMsg = "Procedure failed when trying to activate EXCEL"
Open gblHold For Input As #1
iRecs = 0
Do Until EOF(1) = True
   iRecs = iRecs + 1
   Line Input #1, SelString
Loop
Close 1

Open gblHold For Input As #1

lbl.Caption = "Recreating TTSE SUB Ledger for"

txtfile = frmMDI.cmnDialog.filename
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Recreating TTSE Payments..."
frmMDI.txtStatusMsg.Refresh
'--
sErrMsg = "Procedure failed/check if TTSE changed the format of the XL Sheet"
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
'--Updating subledger routine
'--
iAcnt = 0
cnn.BeginTrans
Do While Not EOF(1) ' do until u have reached end of file 10
   Line Input #1, SelString ' go to next line
   lbl.Caption = "Recreating TTSE SUB Ledger for " & iAcnt
   lbl.Refresh
   ttseID = Trim(Mid(SelString, 37, 15))
   
   iAcnt = iAcnt + 1
      
   sNam = Trim(Mid(SelString, 107, 40))
   If IsEmpty(iAcnt) = True Then
      sNam = " "
   End If
   
   sAD1 = Trim(Mid(SelString, 458, 40))
   If IsEmpty(sAD1) = True Then
      sAD1 = " "
   End If
   
   sAD2 = Trim(Mid(SelString, 498, 40))
   If IsEmpty(sAD2) = True Then
      sAD2 = " "
   End If
   
   sAd3 = Trim(Mid(SelString, 538, 40))
   If IsEmpty(sAd3) = True Then
      sAd3 = " "
   End If
   
   sAD4 = Trim(Mid(SelString, 588, 25))
   If IsEmpty(sAD4) Then
      sAD4 = " "
   End If
   
   sAD5 = Trim(Mid(SelString, 616, 3))
   If IsEmpty(sAD5) Then
      sAD5 = " "
   End If
   sAd3 = sAd3 & " " & sAD4 & " " & sAD5
   
   iCBL = CLng(Mid(SelString, 690, 15))
   If IsEmpty(iCBL) = True Then
      iCBL = 0
   End If
   
   sErrMsg = "Procedure failed in writing to TTSE Sub ledger"
   rsMain.AddNew
   rsMain!GR8NIN = iAcnt
   rsMain!GR8NAM = sNam
   If Not IsNothing(sAD1) Then
      rsMain!GR8Ad1 = sAD1
   Else
      rsMain!GR8Ad1 = " "
   End If
   If Not IsNothing(sAD2) Then
      rsMain!GR8Ad2 = sAD2
   Else
      rsMain!GR8Ad2 = " "
   End If
   If Not IsNothing(sAd3) Then
     rsMain!GR8Ad3 = sAd3
   Else
     rsMain!GR8Ad3 = " "
   End If
   rsMain!GR8CBL = iCBL
    'rsMain! = iRat
   rsMain!Cat = "SH"
   rsMain!Tax = "JA"
   rsMain!ttseID = ttseID
   rsMain!GR8RAT = 0
   rsMain.Update
   iRecsAdded = iRecsAdded + 1  'count records added
    
    
Commit_Check:
    If iRecsAdded = 500 Then
       cnn.CommitTrans
       cnn.BeginTrans
       iRecsAdded = 0
    End If
    iRecs = iRecs + 1
    ProgressBar1.Value = iRecs
Loop
  If iRecsAdded > 0 Then
       cnn.CommitTrans
  End If
  sErrMsg = "Procedure failed in closing EXCEL spread sheet"
 ' .ActiveWorkbook.Save
Close 1


rsMain.Close
'--
Set rsMain = Nothing
CntryRst.Close
Set CntryRst = Nothing

cmdStart.Enabled = False
'-- display success message
lbl.Caption = ""
ProgressBar1.Visible = False
MsgBox "Update successfull. Select Ok to clear this message, then Cancel to end."
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
cnn.Close
''''set cnn = nothing
Exit Sub
Open_err:
  MsgBox "Fail to Open Existing TTSE Sub-Legder; update aborting. "
 
  Exit Sub
ImpTTSE_Fail:
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

Set rsMain = New ADODB.Recordset
On Error GoTo Create_TTSETable
cnn.Open
rsMain.Open "TTSESub", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
rsMain.Close
cnn.Close
FL_Exit:
Exit Sub
Create_TTSETable:
 qSQL = "Create Table TTSESUB (" _
        & "GR8NIN  long not null CONSTRAINT pkGR8NIN PRIMARY KEY, " _
        & "GR8NAM  text(50), " _
        & "GR8AD1  text(50), " _
        & "GR8AD2  text(50) null, " _
        & "GR8AD3  text(50) null, " _
        & "GR8CBL  long not null, " _
        & "CAT  text(2), " _
        & "TAX  text(2))"
 X = csvADODML(qSQL, cnn)
 Resume 0
End Sub


