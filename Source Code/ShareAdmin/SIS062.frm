VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSIS062 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Bonus Stocks"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS062.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6720
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
Attribute VB_Name = "frmSIS062"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, iEOF As Integer
Dim rsMain As ADODB.Recordset
Dim rsBon As ADODB.Recordset
Dim rsCert As ADODB.Recordset
Dim rsComp As ADODB.Recordset
Dim rsDup As ADODB.Recordset
Dim OpenErr As Integer
Dim iOpenMain As Integer
Dim iOpenBon As Integer
Dim iOpenCert As Integer
Dim iOpenCmp As Integer
Dim iOpenDup As Integer
Dim iCommit As Integer
Dim qDMLQry As String
Dim iClient As Long
Dim iShares As Long, icert As Long
Private Sub cmdBtn_Click(Index As Integer)
Dim iRecs As Long, sConst As String
On Error GoTo cmdBtn_Click_Err
Select Case Index
Case 0 'Cancel
    If iOpenBon = True Then rsBon.Close
    If iOpenMain = True Then rsMain.Close
    If iOpenCert = True Then rsCert.Close
    If iOpenCmp = True Then rsComp.Close
    If iOpenDup = True Then rsDup.Close
    '--
    Set rsBon = Nothing
    Set rsMain = Nothing
    Set rsCert = Nothing
    Set rsDup = Nothing
    Set rsComp = Nothing
    '''set cnn = nothing
    Set frmSIS062 = Nothing
    iEOF = True
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
sConst = "Calculations now being performed "
lbl = sConst
sConst = sConst & "for account "
If rsCert.RecordCount < 0 Then
    iRecs = 0
    With rsCert
       .MoveFirst
       While Not .EOF
          iRecs = iRecs + 1
          .MoveNext
       Wend
    End With
 Else
  iRecs = rsCert.RecordCount
 End If
 '--
 InitProgressBar (iRecs)
 ProgressBar1.Visible = True
 iRecs = 1: iClient = 0: iShares = 0
 '--
 With rsCert
    .MoveFirst
    While Not .EOF
       lbl = sConst & !ClientId
       lbl.Refresh
       ProgressBar1.Value = iRecs
       If iClient = 0 Then
         iClient = !ClientId
       End If
       If iClient <> !ClientId Then CalcPay
       iShares = iShares + !shares
       iRecs = iRecs + 1
       .MoveNext
    Wend
 End With
 CalcPay
 '-- update nextcert on company record
  rsComp!nextcert = icert
  rsComp.Update
 cnn.CommitTrans
 cmdBtn_Click (0)
Case Else
End Select
cmdBtn_Click_Exit:
Exit Sub
cmdBtn_Click_Err:
  MsgBox "SIS062/Load"
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
  MsgBox "SIS062/Load"
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
Dim iErr As Integer, iRecs As Integer
'''Set cnn = New ADODB.Connection
cnn.Open
OpenErr = False
iOpenCert = False
iOpenBon = False
iOpenMain = False
iOpenCmp = False
iOpenDup = False
'__
Set rsBon = New ADODB.Recordset
Set rsMain = New ADODB.Recordset
Set rsCert = New ADODB.Recordset
Set rsComp = New ADODB.Recordset
Set rsDup = New ADODB.Recordset
'-----------------------
'-- open BONREF table --
'-----------------------
rsBon.Open "BONUSREF", cnn, , adLockOptimistic, adCmdTable
iOpenBon = True
If rsBon.EOF = True Then
     iErr = 180
     csvShowUsrErr iErr, "Calculate Bonus"
     rsBon.Close
     iOpenBon = False
     GoTo OpenFiles_Close
End If
'---------------------------
'-- Open Certificate View --
'---------------------------
qView = "SELECT CertNo, a.ClientId, a.shares, issdate, " _
        & " b.Cliname " _
        & "from CERTMST a, STKNAME b " _
        & "where a.Status <> 'C' and " _
        & "IssDate <= '" & Format(rsBon!RECDAT, "yyyy-mm-dd") & " 00:00:00' " _
        & "and a.shares > 0 " _
        & " and a.ClientId = b.ClientId " _
        & "order by CliName, " _
        & "a.ClientId, CertNo"
rsCert.Open qView, cnn, , adLockOptimistic, adCmdText
iOpenCert = True
If rsCert.EOF And rsCert.BOF Then
    iErr = 119
    csvShowUsrErr iErr, "Calculate Bonus"
    rsBon.Close
    iOpenBon = False
    rsCert.Close
    iOpenCert = False
    GoTo OpenFiles_Close
End If
'---------------------------
'-- Open Company File --
'---------------------------
rsComp.Open "Company", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
iOpenCmp = True
icert = rsComp!nextcert
'-- check for duplicate number on certmst
sql = "SELECT CERTNO from CERTMST where CERTNO = " _
      & icert
rsDup.Open sql, cnn, , adLockOptimistic, adCmdText
iOpenDup = True
If Not rsDup.EOF Then
  iErr = 182
  csvShowUsrErr 182, "Calculate Bonus"
  cmdBtn_Click (0)
  Exit Sub
End If
rsDup.Close
iOpenDup = 0
'-- open Bonus & clear existing data
'--------------------------------------
lbl = "Deleting Previous Bonus Information"
lbl.Visible = True
'----------
qDMLQry = "DELETE FROM STKBONUS"
X = csvADODML(qDMLQry, cnn)
rsMain.Open "STKBONUS", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
iOpenMain = True
'--
OpenFiles_Exit:
   Exit Sub
OpenFiles_Close:
   cmdBtn_Click (0)
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
  MsgBox "SIS062/OpenFiles"
  OpenErr = True
  cmdBtn_Click (0)
  GoTo OpenFiles_Exit
  
End Sub
Private Static Sub CalcPay()
Dim iWork As Double
Dim allocated As Double
On Error GoTo Calcpay_Err
'---
If iCommit = 0 Or iCommit = 1000 Then
     cnn.BeginTrans
     iCommit = 0
End If

If iShares < rsBon!STKBASE Then
  With rsMain
  .AddNew
  !ClientId = iClient
  !certno = 0
  !BonIss = rsBon!RECDAT
  !shares = iShares
  !allocated = 0
  !unallocated = (iShares * rsBon!STKSTO / rsBon!STKBASE) - !allocated
  .Update
  End With
GoTo Commit_Trans
End If

With rsMain
  .AddNew
  !ClientId = iClient
  !certno = icert
  !BonIss = rsBon!RECDAT
  !shares = iShares
  allocated = (iShares / rsBon!STKBASE) * rsBon!STKSTO
  !allocated = Int(allocated)
  !unallocated = (iShares * rsBon!STKSTO / rsBon!STKBASE) - !allocated
  .Update
End With
icert = icert + 1
'--
Commit_Trans:
iCommit = iCommit + 1
If iCommit = 1000 Then
     cnn.CommitTrans
End If
'--- update stored information
With rsCert
   If Not .EOF Then iClient = !ClientId
   iShares = 0
End With

calcpay_Exit:
  Exit Sub
Calcpay_Err:
  MsgBox "SIS062/calcpay"
  cmdBtn_Click (0)
End Sub

