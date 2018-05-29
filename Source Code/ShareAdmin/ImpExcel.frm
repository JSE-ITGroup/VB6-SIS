VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImpExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Excel Payments"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "ImpExcel.frx":0000
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
Attribute VB_Name = "ImpExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRecs As Long, iRecsAdded As Long
Dim AppExcl As Excel.Application
Dim X As Integer, y As Integer, curCell As Object, nextCol As Object
Dim nextCell As Object, txtfile As String
Dim qSQL As String
Dim SpCon As ADODB.Connection

Private Sub cmdCancel_Click()
Unload Me
Set ImpExcel = Nothing
End Sub

Private Sub cmdStart_Click()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
Dim iAcnt As Long, iCBL As Long, iRat As Single, WrkAddr As String
Dim sNam As String, sAD1 As String, sAD2 As String, sAd3 As String, sTax As String
Dim CliType As String, sAD4 As String, sAD5 As String
Dim sLen As Integer, pos As Integer
sErrMsgL1 = "Update failed while clearing balances of "
sErrMsgL2 = " Note this error."
'--
sMsg = "WARNING: This procedure will will delete existing Records"
sMsg = sMsg & "  then recreate them from the XL Import file. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Building Ledger"
iResp = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If iResp = vbNo Then
  cmdCancel_Click
  Exit Sub
End If
'--

lbl.Caption = "Clearing existing files"
lbl.Visible = True
iResp = RunSP(SpCon, "usp_DeleteStkName", 0)
iResp = RunSP(SpCon, "usp_DeleteCertMSt", 0)

'--
'''On Error GoTo ImpJCSD_Fail
sErrMsg = "Procedure failed when trying to activate EXCEL"
Set AppExcl = CreateObject("Excel.application")
lbl.Caption = "Recreating Ledger for"

txtfile = frmMDI.CmnDialog.filename
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Recreating Payments..."
frmMDI.txtStatusMsg.Refresh
'--
sErrMsg = "Procedure failed/check if the client changed the format of the XL Sheet"
With AppExcl
  .Workbooks.Open (txtfile)
  Set curCell = .Worksheets(1).Range("A4")
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
  'cnn.BeginTrans
  Set curCell = .Worksheets(1).Range("A4")
  Do While Not IsEmpty(curCell)
    sErrMsg = "Procedure failed while formatting/check if the client changed the format of the XL Sheet"
    iAcnt = iAcnt + 1
    lbl.Caption = "Recreating Ledger for " & iAcnt
    lbl.Refresh
    Set nextCol = curCell.Offset(0, 0) '- account name
    sNam = nextCol.Value
    
    Set nextCol = curCell.Offset(0, 1) '-address
    WrkAddr = nextCol.Value
    Set nextCol = curCell.Offset(0, 2) '-shares
    iCBL = nextCol.Value
    
    sErrMsg = "Procedure failed in writing to ledger"
    iResp = InStr(1, sNam, ",")
    If iResp > 0 Then
       CliType = "P"
    Else
       CliType = "C"
    End If
     
    sAD1 = " "
    sAD2 = " "
    sAd3 = " "
    sAD4 = " "
    sAD5 = " "
    
    pos = Len(WrkAddr)
    sLen = 0
    iResp = InStr(1, WrkAddr, ",")
    If iResp > 0 Then
       sAD1 = Mid(WrkAddr, 1, iResp - 1)
       sLen = iResp + 1
    End If
    If iResp = 0 And pos > sLen Then
       sAD1 = WrkAddr
       GoTo AddOthers
    End If
    
    iResp = InStr(sLen, WrkAddr, ",")
    If iResp > 0 Then
       sAD2 = Mid(WrkAddr, sLen, iResp - sLen)
       sLen = iResp + 1
    End If
    If iResp = 0 And pos > sLen Then
       sAD2 = Mid(WrkAddr, sLen, pos - sLen + 1)
       GoTo AddOthers
    End If
    
    iResp = InStr(sLen, WrkAddr, ",")
    If iResp > 0 Then
       sAd3 = Mid(WrkAddr, sLen, iResp - sLen)
       sLen = iResp + 1
    End If
    If iResp = 0 And pos > sLen Then
       sAd3 = Mid(WrkAddr, sLen, pos - sLen + 1)
       GoTo AddOthers
    End If
    
    iResp = InStr(sLen, WrkAddr, ",")
    If iResp > 0 Then
       sAD4 = Mid(WrkAddr, sLen, iResp - sLen)
       sLen = iResp + 1
    End If
    If iResp = 0 And pos > sLen Then
       sAD4 = Mid(WrkAddr, sLen, pos - sLen + 1)
       GoTo AddOthers
    End If
    
    If pos > sLen Then
       sAD5 = Mid(WrkAddr, sLen, pos - sLen + 1)
       GoTo AddOthers
    End If
       
AddOthers:
    iResp = InStr(1, WrkAddr, "USA")
    If iResp > 0 Then
       sTax = "US"
       GoTo Cert_Time
    End If
    
    iResp = InStr(1, WrkAddr, "U.S.A.")
    If iResp > 0 Then
       sTax = "US"
       GoTo Cert_Time
    End If
    
    iResp = InStr(1, WrkAddr, "CANADA")
    If iResp > 0 Then
       sTax = "CN"
       GoTo Cert_Time
    End If
    iResp = InStr(1, WrkAddr, "ENGLAND")
    If iResp > 0 Then
       sTax = "UK"
       GoTo Cert_Time
    End If
    iResp = InStr(1, WrkAddr, "AUSTRALIA")
    If iResp > 0 Then
       sTax = "AU"
       GoTo Cert_Time
    End If
    
    sTax = "JA"
    
Cert_Time:
    iResp = RunSP(SpCon, "usp_ImportExcel", 0, iAcnt, CliType, sNam, iCBL, sTax, sAD1, sAD2, sAd3, sAD4, sAD5)
    
Commit_Check:
    'If iRecsAdded = 500 Then
     '  cnn.CommitTrans
      ' cnn.BeginTrans
       'iRecsAdded = 0
    'End If
    iRecs = iRecs + 1
    ProgressBar1.Value = iRecs

    Set nextCell = curCell.Offset(1, 0)
    Set curCell = nextCell
  Loop
  'If iRecsAdded > 0 Then
  '     cnn.CommitTrans
  'End If
  sErrMsg = "Procedure failed in closing EXCEL spread sheet"
 ' .ActiveWorkbook.Save
 .Workbooks.Close
 AppExcl.Quit
End With
'--
cmdStart.Enabled = False
SpCon.Close
'-- display success message
lbl.Caption = ""
ProgressBar1.Visible = False
MsgBox "Update successfull. Select Ok to clear this message, then Cancel to end."
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Exit Sub
Open_err:
  MsgBox "Fail to Open Existing the clients database; update aborting. "
 
  Exit Sub
ImpExcel_Fail:
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
On Error GoTo FL_Exit
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

FL_Exit:
Exit Sub
End Sub


