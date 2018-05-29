VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImpACHFinIns 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import ACH Financial Institutions  "
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "ImpACHFinIns.frx":0000
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
Attribute VB_Name = "ImpACHFinIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRecs As Long, iRecsAdded As Long
Dim SpCon As ADODB.Connection
Dim AppExcl As Excel.Application
Dim X As Integer, Y As Integer, curCell As Object, nextCol As Object
Dim nextCell As Object, txtfile As String
Dim qSQL As String

Private Sub cmdCancel_Click()
Unload Me
Set ImpACHFinIns = Nothing
End Sub

Private Sub cmdStart_Click()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
Dim FinID As String, BranchID As String, BankID As String
Dim BranchName As String, BankName As String
sErrMsgL1 = "Import failed while update reject code "
sErrMsgL2 = " Note this error."

'--
sMsg = "WARNING: This procedure will load ACH Financial Institutions"
sMsg = sMsg & "  FROM the EXCEL file selected. Please click No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Building ACH Financial Institutions"
iResp = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If iResp = vbNo Then
  cmdCancel_Click
  Exit Sub
End If
'--

sErrMsg = "Procedure failed when trying to activate EXCEL"
Set AppExcl = CreateObject("Excel.application")
lbl.Caption = "Updating Financial Institutions"

txtfile = frmMDI.cmnDialog.filename
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Updating ACH Financial Institutions"
frmMDI.txtStatusMsg.Refresh
'--
sErrMsg = "Procedure failed/check if ACH Financial Institutions file format of the XL Sheet has changed"
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
  'cnn.BeginTrans
  Set curCell = .Worksheets(1).Range("A2")
  Do While Not IsEmpty(curCell)
    sErrMsg = "Procedure failed/check if ACH Financial Instituions file format of the XL Sheet has changed"
    Set nextCol = curCell.Offset(0, 0) '- FinID
    FinID = nextCol.Value
    lbl.Caption = "Updating Description for " & FinID
    lbl.Refresh
    If Len(FinID) <> 8 Then
       .Workbooks.Close
       AppExcl.Quit
       MsgBox "Financial Institution ID " & FinID & " should be 8 characters, please correct and resubmit"
       GoTo Open_err
    End If
    
    Set nextCol = curCell.Offset(0, 1) '- Branch ID
    BranchID = nextCol.Value
    If Len(BranchID) <> 5 Then
       .Workbooks.Close
       AppExcl.Quit
       MsgBox "Branch ID " & BranchID & " should be 5 characters, please correct and resubmit"
       GoTo Open_err
    End If
    
    Set nextCol = curCell.Offset(0, 2) '- Bank ID
    BankID = nextCol.Value
    If Len(BankID) <> 3 Then
       .Workbooks.Close
       AppExcl.Quit
       MsgBox "Bank ID " & BranchID & " should be 3 characters, please correct and resubmit"
       GoTo Open_err
    End If
    
    Set nextCol = curCell.Offset(0, 3) '- Bank Name
    BankName = nextCol.Value
    If Len(BankName) > 50 Then
       .Workbooks.Close
       AppExcl.Quit
       MsgBox "Bank Name " & BankName & " cannot be greater than 50 characters, please correct and resubmit"
       GoTo Open_err
    End If
    If Len(BankName) < 1 Then
       .Workbooks.Close
       AppExcl.Quit
       MsgBox "Bank Name " & BankName & " cannot be blank, please correct and resubmit"
       GoTo Open_err
    End If
     
    Set nextCol = curCell.Offset(0, 4) '- Branch Name
    BranchName = nextCol.Value
    If Len(BranchName) > 50 Then
       .Workbooks.Close
       AppExcl.Quit
       MsgBox "Branch Name " & BranchName & " cannot be greater than 50 characters, please correct and resubmit"
       GoTo Open_err
    End If
    If Len(BranchName) < 1 Then
       .Workbooks.Close
       AppExcl.Quit
       MsgBox "Branch Name " & BranchName & " cannot be blank, please correct and resubmit"
       GoTo Open_err
    End If
  
    iResp = RunSP(SpCon, "usp_ImportACHFinIns", 0, FinID, BranchID, BankID, BankName, BranchName)
    If iResp <> 0 Then
       .Workbooks.Close
       AppExcl.Quit
       MsgBox "Financial Institution " & FinID & " was not updated/added"
       GoTo Open_err
    End If
    iRecsAdded = iRecsAdded + 1  'count records added
    
    If nextCol = "" Then
       GoTo Commit_Check
    End If
    
Commit_Check:
    iRecs = iRecs + 1
    ProgressBar1.Value = iRecs
    Set nextCell = curCell.Offset(1, 0)
    Set curCell = nextCell
  Loop
  sErrMsg = "Procedure failed in closing EXCEL spread sheet"
 ' .ActiveWorkbook.Save
 .Workbooks.Close
 AppExcl.Quit
End With

'--
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
  MsgBox "Fail to Open ACH Financial Institutions file; update aborting. "
 
  Exit Sub
CmdStart_Fail:
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
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg


FL_Exit:
Exit Sub
End Sub


