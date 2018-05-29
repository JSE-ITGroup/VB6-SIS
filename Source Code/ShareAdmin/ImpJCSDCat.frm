VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImpJCSDCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Stock Exchange Shareholders Categories"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "ImpJCSDCat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5220
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
Attribute VB_Name = "ImpJCSDCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As ADODB.Recordset
Dim rsMnDte As ADODB.Recordset
Dim iRecs As Long, iRecsAdded As Long
Dim SpCon As ADODB.Connection
Dim AppExcl As Excel.Application
Dim X As Integer, y As Integer, curCell As Object, nextCol As Object
Dim nextCell As Object, txtfile As String
Dim qSQL As String

Private Sub cmdCancel_Click()
Unload Me
Set ImpJCSDCat = Nothing
End Sub

Private Sub CmdStart_Click()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
Dim iAcnt As String
Dim sNam As String, sTax As String
sErrMsgL1 = "Update failed while clearing balances of "
sErrMsgL2 = " Note this error."

'--
sMsg = "WARNING: This procedure will Change the " & gblHold & " Categories of Shareholders"
sMsg = sMsg & "  FROM SH to the category in the EXCEL file OR Chaning the Tax code from JA. Please click No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Building Category/Tax List"
iResp = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If iResp = vbNo Then
  cmdCancel_Click
  Exit Sub
End If
'--


'''On Error GoTo ImpJCSD_Fail
sErrMsg = "Procedure failed when trying to activate EXCEL"
Set AppExcl = CreateObject("Excel.application")
lbl.Caption = "Updating Categories"

txtfile = frmMDI.cmnDialog.filename
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Updating " & gblHold & " Categories"
frmMDI.txtStatusMsg.Refresh
'--
sErrMsg = "Procedure failed/check if Category file format of the XL Sheet has changed"
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
    sErrMsg = "Procedure failed/check if Category file format of the XL Sheet has changed"
    Set nextCol = curCell.Offset(0, 0) '- account number
    If IsNumber(nextCol.Value) Then
       iAcnt = nextCol.Value
    Else
       iAcnt = ConvertToNumber(nextCol.Value)
    End If
    lbl.Caption = "Updating Category for " & iAcnt
    lbl.Refresh
    Set nextCol = curCell.Offset(0, 1) '-Category
    sNam = nextCol.Value
    If Len(sNam) > 2 Then
       .Workbooks.Close
       AppExcl.Quit
       MsgBox "Category code for " * iAcnt & " is greater than 2 characters"
       GoTo Open_err
    End If
    Set nextCol = curCell.Offset(0, 2) '-Tax
    sTax = nextCol.Value
    If Len(sTax) > 2 Then
       .Workbooks.Close
       AppExcl.Quit
       MsgBox "Tax code for " * iAcnt & " is greater than 2 characters"
       GoTo Open_err
    End If
    iResp = RunSP(SpCon, "usp_ImportSECategory", 0, iAcnt, sNam, sTax, gblReply)
    
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
CmdStart.Enabled = False

'-- display success message
lbl.Caption = ""
ProgressBar1.Visible = False
MsgBox "Update successfull. Select Ok to clear this message, then Cancel to end."
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Exit Sub
Open_err:
  MsgBox "Fail to Open Existing " & gblHold & " Sub-Legder; update aborting. "
 
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
Me.Caption = "Update " & gblHold & " Shareholders Categories/Tax group"

FL_Exit:
Exit Sub
End Sub


