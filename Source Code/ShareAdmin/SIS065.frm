VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSIS065 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Stocks with Bonus Certs"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS065.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6735
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Begin"
      Default         =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   4320
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      ToolTipText     =   "Returns to main menu"
      Top             =   3360
      Width           =   975
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   6132
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      Height          =   300
      Index           =   0
      Left            =   5400
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      ToolTipText     =   "Returns to main menu"
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ver:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Display program Information"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS065"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, iEOF As Integer
Dim rsMain As ADODB.Recordset
Dim rsCert As ADODB.Recordset
Dim rsComp As ADODB.Recordset
Dim rsAct As ADODB.Recordset
Dim OpenErr As Integer
Dim iOpenMain As Integer
Dim iOpenCert As Integer
Dim iOpenAct As Integer
Dim iOpenCmp As Integer
Dim iCommit As Integer, iErr As Integer
Dim sql As String, iTotBon As Long
Dim iRecs As Long, sConst As String
Private Sub cmdBtn_Click(Index As Integer)

On Error GoTo cmdBtn_Click_Err
Select Case Index
Case 0 'Cancel
    Shutdown
    '--
    Unload Me
    frmSIS060.Visible = True
     '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
Case 1 'Perform Calculations
'--
' wait & hourglass message
'--------------
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.Refresh
OpenFiles
If OpenErr = True Then GoTo cmdBtn_Click_Exit
If iEOF = True Then GoTo cmdBtn_Click_Exit
sConst = "Update now being performed "
lbl = sConst
sConst = sConst & "for account "
If rsMain.RecordCount < 0 Then
    iRecs = 0
    With rsMain
       .MoveFirst
       While Not .EOF
          iRecs = iRecs + 1
          .MoveNext
       Wend
    End With
 Else
  iRecs = rsMain.RecordCount
 End If
 '--
 InitProgressBar (iRecs)
 ProgressBar1.Visible = True
 lbl.Visible = True
 iRecs = 1
 '--
 With rsMain
  .MoveFirst
  If Not .EOF Then
    '-- open certmst
     sql = "SELECT * from CERTMST where CERTNO = " _
        & !certno
     rsCert.Open sql, cnn, adOpenDynamic, adLockPessimistic, adCmdText
     iOpenCert = True
     '---
  BonusUpdate
  '-- update nextcert on company record
  rsComp!Totstocks = rsComp!Totstocks - iTotBon
  rsComp!issStocks = rsComp!issStocks + iTotBon
  rsComp.Update
  '--- Remove StkBONUS records to eliminate duplicate updating
  sql = "Delete from stkbonus"
  X = csvADODML(sql, cnn)
  '--
  If iCommit > 0 And iCommit <> 1000 Then
    cnn.CommitTrans
    cmdBtn_Click (0)
  End If
 End If
End With
Case Else
End Select
cmdBtn_Click_Exit:
Exit Sub
cmdBtn_Click_Err:
  MsgBox "SIS065/Load"
  cmdBtn_Click (0)
End Sub
Private Sub Form_Load()
'--
Dim i As Integer
Dim strTmp As String
On Error GoTo FL_ERR
iEOF = False
iCommit = 0
'--
   csvCenterForm Me, gblMDIFORM
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '----------------------
   '--  disable menu items
   '----------------------
   frmMDI.mnuFile.Enabled = False
   frmMDI.btnClose.Enabled = False
   frmMDI.mnuLists.Enabled = False
   frmMDI.mnuAct.Enabled = False
   frmMDI.mnuAdm.Enabled = False
   '--
   ProgressBar1.Visible = False
   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS065/Load"
  Unload Me
End Sub

Private Sub InitProgressBar(max As Long)
    ProgressBar1.max = max
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.Min

End Sub

Private Sub Form_Unload(Cancel As Integer)
cnnClose
If iEOF = False Then
  Cancel = -1
End If
End Sub

Private Sub OpenFiles()
Dim icount As Integer, i As Integer
On Error GoTo OpenFiles_Err
Dim qSQL As String, qView As String, sql As String
Dim iRecs As Integer
'''Set cnn = New ADODB.Connection
cnn.Open
OpenErr = False
iOpenCert = False
iOpenMain = False
iOpenCmp = False
iOpenAct = False
'__
Set rsMain = New ADODB.Recordset
Set rsCert = New ADODB.Recordset
Set rsComp = New ADODB.Recordset
Set rsAct = New ADODB.Recordset
'------------------------------
'--  Issue warning message
'---------------------------
iErr = 181
i = csvYesNo(iErr, "Update Bonus Stocks")
If i = False Then
    OpenErr = True
    cmdBtn_Click (0)
    Exit Sub
End If
'-----------------------
'-- open BONREF table --
'-----------------------

'qSQL = "Select * from StkBonus where CertNO > 0"
qSQL = "SELECT * FROM StkBonus where CertNo > 0"
rsMain.Open qSQL, cnn, adOpenStatic
iOpenMain = True
If rsMain.EOF = True Then
     iErr = 180
     csvShowUsrErr iErr, "Calculate Bonus"
     rsMain.Close
     iOpenMain = False
     GoTo OpenFiles_Close
End If
'---------------------------
'-- Open Company File --
'---------------------------
rsComp.Open "Company", cnn, adOpenDynamic, adLockPessimistic, adCmdTable
iOpenCmp = True
'--
rsAct.Open "STKACTIV", cnn, adOpenDynamic, adLockPessimistic, adCmdTable
iOpenAct = True
'--
OpenFiles_Exit:
   Exit Sub
OpenFiles_Close:
   Set rsMain = Nothing
   Set rsComp = Nothing
   '''set cnn = nothing
   Set frmSIS065 = Nothing
   iEOF = True
   Unload Me
   frmSIS060.Show
   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
   GoTo OpenFiles_Exit
   
OpenFiles_Err:
  MsgBox "SIS065/OpenFiles"
  OpenErr = True
  cmdBtn_Click (0)
  GoTo OpenFiles_Exit
  
End Sub

Private Sub Shutdown()
If iOpenMain = True Then rsMain.Close
If iOpenCmp = True Then rsComp.Close
If iOpenCert = True Then rsCert.Close
If iOpenAct = True Then rsAct.Close
Set rsMain = Nothing
Set rsCert = Nothing
Set rsComp = Nothing
Set rsAct = Nothing
'''set cnn = nothing
Set frmSIS065 = Nothing
iEOF = True
End Sub



Private Sub BonusUpdate()
Dim iWork As Double
Dim iLine As Long
Dim sBatch As String
On Error GoTo BonusUpdate_Err
'---
sBatch = "BN" & Year(Now) & Month(Now)
iLine = 0
With rsMain
 .MoveFirst
 While Not .EOF
   If iCommit = 0 Or iCommit = 1000 Then
     cnn.BeginTrans
     iCommit = 0
   End If
   lbl = sConst & !ClientId
   lbl.Refresh
   ProgressBar1.Value = iRecs
   '-- add to stkactiv for printing
    rsAct.AddNew
    rsAct!TrnBatch = sBatch
    rsAct!Form = "BONUS"
    iLine = iLine + 1
    rsAct!stklineno = iLine
    rsAct!ClientId = !ClientId
    rsAct!TRNDATE = Now
    rsAct!TrnCode = "G"  'Bonus Issues
    rsAct!Status = "O"
    rsAct!FRCERT = 0
    rsAct!IssDate = !BonIss
    rsAct!shares = !allocated
    rsAct!certno = !certno
    rsAct!BrokerBuy = 0
    rsAct!BROKERID = 0
    rsAct.Update
   '-- Add the Bonus cert
    rsCert.AddNew
    rsCert!ClientId = !ClientId
    rsCert!certno = !certno
    rsCert!IssDate = !BonIss
    rsCert!shares = !allocated
    rsCert!Status = "A"
    rsCert!assigned = 0
    rsCert!TrnBatch = "BONUS"
    rsCert!TRNDATE = Now
    rsCert!Remarks = "Bonus Issue"
    rsCert.Update
    '-- Update stkname
    sql = "Update stkname set shares = shares + " & !allocated _
          & " where clientid = " & !ClientId
    X = csvADODML(sql, cnn)
    iTotBon = iTotBon + !allocated
    '-- count total Bonus applied
    iCommit = iCommit + 1
    If iCommit = 1000 Then
        cnn.CommitTrans
    End If
    iRecs = iRecs + 1
    .MoveNext
 Wend
End With
'--
BonusUpdate_Exit:
  Exit Sub
BonusUpdate_Err:
  MsgBox "SIS065/bonusupdate"
  cmdBtn_Click (0)
End Sub
