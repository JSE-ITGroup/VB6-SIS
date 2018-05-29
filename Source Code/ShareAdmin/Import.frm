VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Import 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Dataflex Text files"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4335
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
Attribute VB_Name = "Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As ADODB.Recordset
Dim rsName As ADODB.Recordset
Dim rsCert As ADODB.Recordset
Dim fs, F, iRecs As Long, iRecsAdded As Long
Dim sInRec As String, txtfile As String, sQuote As String
Dim X As Integer, Y As Integer
Dim qSQL As String

Private Sub cmdCancel_Click()
Unload Me

Set Import = Nothing
End Sub

Private Sub cmdStart_Click()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
sErrMsgL1 = "Import failed during Purge of "
sErrMsgL2 = " Note this error."
Set rsMain = New ADODB.Recordset
Set rsName = New ADODB.Recordset
Set rsCert = New ADODB.Recordset
'--
sMsg = "WARNING: Importing Dataflex Text files will erase"
sMsg = sMsg & " your existing data. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Import Dataflex Text Files"
iResp = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If iResp = vbNo Then
  cmdCancel_Click
  Exit Sub
End If
'--
lbl.Caption = "Clearing database files"
lbl.Visible = True
qSQL = "DELETE * FROM USERS where SYSTEMNAME <> 'ADMIN'"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "USERS." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM UNUSEDNOS"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "UUNUSEDNOS." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "Delete from stkcat"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKCAT." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "Delete from stktaxr"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKTAXR." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKPYMNTS"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKPYMNTS." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKBRKTRN"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKBRKTRN." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKBRKHIS"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKBRKHIS." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKBKCRT"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKBKCRT." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKBRKPL"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKBRKPL." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKBONUS"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKBONUS." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKBANK"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKBANK." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKASSGN"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKASSGN." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKACTIV"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKACTIV." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM MNDPAYMNTS"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "MNDPAYMENTS." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM DIVREF"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "DIVREF." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM CHQTRN"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "CHQTRN." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM CERTMST"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "CERTMST." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKJOINT"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKJOINT." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKNAME"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKNAME." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM STKMNDTE"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "STKMNDTE." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM BONUSREF"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "BONUSREF." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM BATHDR"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "BATHDR." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM BNKREF"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "BNKREF." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM BNKLODGE"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "BNKLODGE." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM AUDTRN"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "AUDTRN." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--
qSQL = "DELETE * FROM ARCHAUDIT"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "ARCHAUDIT." & sErrMsgL2
  GoTo IMPORT_Fail
  Exit Sub
End If
'--

lbl.Caption = "Importing Company Details"
txtfile = App.Path & "\stkctrl.txt"
lbl.Refresh
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Importing text files..."
frmMDI.txtStatusMsg.Refresh
'--
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
qSQL = "Select * from Company"
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
Import_Company
F.Close
'--
txtfile = App.Path & "\stktype.txt"
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
Import_STKTYPE
rsMain.Close
F.Close
'--

lbl.Caption = "Importing Stockholder Category Info"
lbl.Refresh
txtfile = App.Path & "\stkcate.txt"
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
iRecs = 0
Do Until F.atendofstream = True
    iRecs = iRecs + 1
    sInRec = F.readline
Loop
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
F.Close
Set F = fs.opentextfile(txtfile)
'--import stkcate routine
'On Error GoTo Open_err
'--
qSQL = "SELECT * FROM STKCAT"
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
sInRec = F.readline
If F.atendofstream = True Then GoTo Open_err
cnn.BeginTrans
Do Until F.atendofstream = True
  Import_STKCATE
  iRecs = iRecs + 1
  ProgressBar1.Value = iRecs
  sInRec = F.readline
Loop
If iRecsAdded > 0 Then
       cnn.CommitTrans
       iRecsAdded = 0
End If
F.Close
rsMain.Close

'--
lbl.Caption = "Importing Tax & Residency Codes Info"
lbl.Refresh
txtfile = App.Path & "\stktaxr.txt"
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
iRecs = 0
Do Until F.atendofstream = True
    iRecs = iRecs + 1
    sInRec = F.readline
Loop
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
F.Close
Set F = fs.opentextfile(txtfile)
'--import stkcate routine
'On Error GoTo Open_err
'--
qSQL = "SELECT * FROM STKTAXR"
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
sInRec = F.readline
If F.atendofstream = True Then GoTo Open_err
cnn.BeginTrans
Do Until F.atendofstream = True
  Import_STKTAXR
  iRecs = iRecs + 1
  ProgressBar1.Value = iRecs
  sInRec = F.readline
Loop
If iRecsAdded > 0 Then
       cnn.CommitTrans
       iRecsAdded = 0
End If
F.Close
rsMain.Close
'--
lbl.Caption = "Importing Stockholder Name & Address Info"
lbl.Refresh
txtfile = App.Path & "\stkname.txt"
'-- count recs to process
'--
iRecs = 0
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
Do Until F.atendofstream = True
    iRecs = iRecs + 1
    sInRec = F.readline
Loop
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
F.Close
Set F = fs.opentextfile(txtfile)
'--import stkname routine
'On Error GoTo Open_err
'--
qSQL = "SELECT * FROM STKNAME"
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
sInRec = F.readline
If F.atendofstream = True Then GoTo Open_err
cnn.BeginTrans
Do Until F.atendofstream = True
  Import_STKNAME
  iRecs = iRecs + 1
  ProgressBar1.Value = iRecs
  sInRec = F.readline
Loop
If iRecsAdded > 0 Then
       cnn.CommitTrans
       iRecsAdded = 0
End If
F.Close
rsMain.Close
'--
lbl.Caption = "Importing Stockholder Certificate Info"
lbl.Refresh
txtfile = App.Path & "\stkcert.txt"
'-- count recs to process
'--
iRecs = 0
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
Do Until F.atendofstream = True
    iRecs = iRecs + 1
    sInRec = F.readline
Loop
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
F.Close
Set F = fs.opentextfile(txtfile)
'--import stkcerts routine
'On Error GoTo Open_err
'--
qSQL = "SELECT * FROM CERTMST"  'clean out table
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
sInRec = F.readline
If F.atendofstream = True Then GoTo Open_err
cnn.BeginTrans
Do Until F.atendofstream = True
  Import_STKCERT
  iRecs = iRecs + 1
  ProgressBar1.Value = iRecs
  sInRec = F.readline
Loop
If iRecsAdded > 0 Then
       cnn.CommitTrans
       iRecsAdded = 0
End If
F.Close
rsMain.Close
'--
lbl.Caption = "Importing Joint Stockholders Info"
lbl.Refresh
txtfile = App.Path & "\stkjoint.txt"
'-- count recs to process
'--
iRecs = 0
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
Do Until F.atendofstream = True
    iRecs = iRecs + 1
    sInRec = F.readline
Loop
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
F.Close
Set F = fs.opentextfile(txtfile)
'--import stkjoint routine
'On Error GoTo Open_err
'--
qSQL = "SELECT * FROM STKJOINT"
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
sInRec = F.readline
If F.atendofstream = True Then GoTo STKJOINT_END
cnn.BeginTrans
Do Until F.atendofstream = True
    Import_STKJOINT
  iRecs = iRecs + 1
  ProgressBar1.Value = iRecs
  sInRec = F.readline
Loop
If iRecsAdded > 0 Then
       cnn.CommitTrans
       iRecsAdded = 0
End If
STKJOINT_END:
F.Close
rsMain.Close
'--
lbl.Caption = "Importing Bank Mandate Info"
lbl.Refresh
txtfile = App.Path & "\stkmndte.txt"
'-- count recs to process
'--
iRecs = 0
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
Do Until F.atendofstream = True
    iRecs = iRecs + 1
    sInRec = F.readline
Loop
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
F.Close
Set F = fs.opentextfile(txtfile)
'--import stkmndte routine
'On Error GoTo Open_err
'--
qSQL = "SELECT * FROM STKMNDTE"  'clean out table
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
sInRec = F.readline
If F.atendofstream = True Then GoTo STKMNDTE_END
cnn.BeginTrans
Do Until F.atendofstream = True
    Import_STKMNDTE
  iRecs = iRecs + 1
  ProgressBar1.Value = iRecs
  sInRec = F.readline
Loop
If iRecsAdded > 0 Then
       cnn.CommitTrans
       iRecsAdded = 0
End If
STKMNDTE_END:
F.Close
rsMain.Close
'--
lbl.Caption = "Importing Brokers Pool Info"
lbl.Refresh
txtfile = App.Path & "\stkbrkpl.txt"
'-- count recs to process
'--
iRecs = 0
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
Do Until F.atendofstream = True
    iRecs = iRecs + 1
    sInRec = F.readline
Loop
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
F.Close
Set F = fs.opentextfile(txtfile)
'--import stkbrkpl routine
'On Error GoTo Open_err
'--
qSQL = "SELECT * FROM STKBRKPL"  'clean out table
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
sInRec = F.readline
If F.atendofstream = True Then GoTo STKBRKPL_END
cnn.BeginTrans
Do Until F.atendofstream = True
    Import_STKBRKPL
  iRecs = iRecs + 1
  ProgressBar1.Value = iRecs
  sInRec = F.readline
Loop
If iRecsAdded > 0 Then
       cnn.CommitTrans
       iRecsAdded = 0
End If
STKBRKPL_END:
F.Close
rsMain.Close
'--
lbl.Caption = "Importing Brokers Certification Info"
lbl.Refresh
txtfile = App.Path & "\stkcrtbk.txt"
'-- count recs to process
'--
iRecs = 0
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
Do Until F.atendofstream = True
    iRecs = iRecs + 1
    sInRec = F.readline
Loop
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
F.Close
Set F = fs.opentextfile(txtfile)
'--import stkcrtbk routine
'On Error GoTo Open_err
'--
qSQL = "SELECT * FROM STKBKCRT"  'clean out table
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
sInRec = F.readline
If F.atendofstream = True Then GoTo STKBKCRT_END
cnn.BeginTrans
Do Until F.atendofstream = True
    Import_STKCRTBK
  iRecs = iRecs + 1
  ProgressBar1.Value = iRecs
  sInRec = F.readline
Loop
If iRecsAdded > 0 Then
       cnn.CommitTrans
       iRecsAdded = 0
End If
STKBKCRT_END:
F.Close
rsMain.Close
'--
lbl.Caption = "Importing Bank Reconciliation Info"
lbl.Refresh
txtfile = App.Path & "\stkbank.txt"
'-- count recs to process
'--
iRecs = 0
Set F = fs.opentextfile(txtfile)
If F.atendofstream = True Then GoTo STKBANK_END
sInRec = F.readline
Do Until F.atendofstream = True
    iRecs = iRecs + 1
    sInRec = F.readline
Loop
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
F.Close
Set F = fs.opentextfile(txtfile)
'--import stkbank routine
'On Error GoTo Open_err
'--
qSQL = "SELECT * FROM STKBANK"
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
sInRec = F.readline
If F.atendofstream = True Then GoTo STKBANK_END
cnn.BeginTrans
Do Until F.atendofstream = True
  Import_STKBANK
  iRecs = iRecs + 1
  ProgressBar1.Value = iRecs
  sInRec = F.readline
Loop
If iRecsAdded > 0 Then
       cnn.CommitTrans
       iRecsAdded = 0
End If
rsMain.Close
STKBANK_END:
F.Close

'--
Set rsMain = Nothing
Set rsName = Nothing
cmdStart.Enabled = False
'-- display success message
lbl.Caption = ""
ProgressBar1.Visible = False
MsgBox "Conversion successful. Select Ok to clear this message, then Cancel to end."
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Exit Sub
Open_err:
  MsgBox "Input Text File " & txtfile & " is blank; conversion aborting. "
 
  Exit Sub
IMPORT_Fail:
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
sQuote = Chr(34)
'--
Set cnn = New ADODB.Connection
cnn.Open gblFileName
End Sub

Private Sub Import_STKNAME()
 '--
With rsMain
    .AddNew
    ' -- extract client id
    '---------------------
    X = InStr(1, sInRec, ",", vbTextCompare)
    If X = 0 Then GoTo STKNAME_BADDATA
    !CLIENTID = Val(left(sInRec, X - 1))
    '-- extract cliname
    '------------------
    X = X + 2 ' position pointer on first char of the name
    Y = InStr(X, sInRec, sQuote, vbTextCompare)
    If Y = 0 Then GoTo STKNAME_BADDATA
    !cliname = Mid(sInRec, X, Y - X)
    '-- determine if company or person in name field
    '------------------------------------------------
    X = InStr(X, sInRec, ",", vbTextCompare)
    If X = 0 Then GoTo STKNAME_BADDATA
    If X > Y Then ' it is a company
      !CLITYPE = "C"
    Else
      !CLITYPE = "P"
    End If
    '-- extract address line 1
    '-------------------------
    X = Y   'point to quote at the end of the name
    X = InStr(X + 1, sInRec, sQuote, vbTextCompare) 'skip to the next quote
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare) ' find the next quote
    If Y = 0 Then GoTo STKNAME_BADDATA
    If Y - X = 1 Then 'there is no addr1
     !CLIADDR1 = " "
    Else
     X = X + 1
     !CLIADDR1 = Mid(sInRec, X, Y - X)
    End If
    '-- extract address line 2
    '-------------------------
    X = Y ' POINT TO THE END OF the line 1
    X = InStr(X + 1, sInRec, sQuote, vbTextCompare) 'skip to the next quote
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare) ' find the next quote
    If Y = 0 Then GoTo STKNAME_BADDATA
    If Y - X = 1 Then 'there is no addr2
     !CLIADDR2 = " "
    Else
      X = X + 1
     !CLIADDR2 = Mid(sInRec, X, Y - X)
    End If
    '--extract line 3 if entered
    '----------------------------
    X = Y  ' set x to the quote at end of line 2
    X = InStr(X + 1, sInRec, sQuote, vbTextCompare) 'skip to the next quote
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare) ' find the next quote
    If Y = 0 Then GoTo STKNAME_BADDATA
    If Y - X = 3 Then 'there is no addr3
    Else
       X = X + 1 'advance to the first char of addr3
       !CLIADDR3 = Mid(sInRec, X, Y - X)
    End If
    '--extract line 4 if entered
    '----------------------------
    X = Y   ' set x to the quote at end of line 3
    X = InStr(X + 1, sInRec, sQuote, vbTextCompare) 'skip to the next quote
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare) ' find the next quote
    If Y = 0 Then GoTo STKNAME_BADDATA
    If Y - X = 1 Then 'there is no addr4
    Else
       X = X + 1 'advance to the first char of addr4
       !CLIADDR4 = Mid(sInRec, X, Y - X)
    End If
    '--extract line 5 if entered
    '----------------------------
    X = Y   ' set x to the quote at end of line 4
    X = InStr(X + 1, sInRec, sQuote, vbTextCompare) 'skip to the next quote
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare) ' find the next quote
    If Y = 0 Then GoTo STKNAME_BADDATA
    If Y - X = 1 Then 'there is no addr5
    Else
       X = X + 1 'advance to the first char of addr5
       !CLIADDR5 = Mid(sInRec, X, Y - X)
    End If
    '--extract catcode if entered
    '----------------------------
    X = Y  ' set x to the quote at end of line 5
    X = InStr(X + 1, sInRec, sQuote, vbTextCompare) 'skip to the next quote
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare) ' find the next quote
    If Y = 0 Then GoTo STKNAME_BADDATA
    If Y - X = 1 Then 'there is no category code
    Else
       X = X + 1 'advance to the first char of category
       !catcode = Mid(sInRec, X, Y - X)
    End If
    '--extract tax code if entered
    '----------------------------
    X = Y ' set x to the quote at end of catcode
    X = InStr(X + 1, sInRec, sQuote, vbTextCompare) 'skip to the next quote
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare) ' find the next quote
    If Y = 0 Then GoTo STKNAME_BADDATA
    If Y - X = 1 Then 'there is no tax code
    Else
       X = X + 1 'advance to the first char of tax code
       !ResCode = Mid(sInRec, X, Y - X)
    End If
    '--extract shares
    '----------------
    X = Y + 2 ' skip to the first digit of stocks
    Y = InStr(X, sInRec, ",", vbTextCompare)  ' find the comma at the end of stocks
    If Y = 0 Then GoTo STKNAME_BADDATA
    !shares = Val(Mid(sInRec, X, Y - X))
    '-- extract remarks
    '------------------
    X = Y + 1 ' set x to the quote in remarks
    
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare) ' find the next quote
    If Y = 0 Then GoTo STKNAME_BADDATA
    If Y - X = 1 Then 'there is no remark
       !remarks = " "
    Else
       X = X + 1 'advance to the first char of remarks
       !remarks = Mid(sInRec, X, Y - X)
    End If
    '-- extract date account opened
    '------------------------------
    X = Y + 3 ' set x to the start of date account opened
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare) ' find the next quote to end the date
    If Y - X <> 8 Then
     !dteopened = Now()
    Else
    !dteopened = DateValue(Mid(sInRec, X, Y - X))
    End If
    '-- extract Joint
    X = X + 11 'set x to position of joint indicator
    If Mid(sInRec, X, 1) = "Y" Then
      !JOINT = True
    Else
      !JOINT = False
    End If
    !YTDPYMNT = 0 ' initialize new field
    .Update
    iRecsAdded = iRecsAdded + 1  'count records added
    If iRecsAdded = 500 Then
       cnn.CommitTrans
       cnn.BeginTrans
       iRecsAdded = 0
    End If
End With
Exit Sub
STKNAME_BADDATA:
   MsgBox "BAD DATA/Import failed during conversion of STKNAME. Note this error." _
   & " " & sInRec
   Exit Sub
End Sub

Private Sub Import_Company()
Dim X As Long, Y As Long
'--
With rsMain
    '-- extract company name
    X = 1
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
    If Y - X = 1 Then 'no name present
      !compname = "A "
    Else
      X = X + 1
      !compname = Mid(sInRec, X, Y - X)
    End If
    '-- EXTRACT Next Cheque No
    X = Y + 2 'postion pointer at starting digit of chq #
    Y = InStr(X, sInRec, ",", vbTextCompare)
    !NEXTCHQ = Val(Mid(sInRec, X, Y - X))
    '-- extract next shareholder No
    X = Y + 1 'postion pointer at starting digit of acct#
    Y = InStr(X, sInRec, ",", vbTextCompare)
    !NEXTACCT = Val(Mid(sInRec, X, Y - X))
    '-- extract cusip #
    X = Y + 1 'position x at the quote preceeding the cusip
    Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
    If Y - X = 1 Then 'there is no cusip
      !CUSIP = " "
    Else
       !CUSIP = Mid(sInRec, X + 1, Y - X - 1)
    End If
    '-- set up defaults for other fields
    !COMPSTREET = " "
    !COMPPOB = " "
    !COMPCITY = " "
    !COMPSTREET = " "
    !TRNUMBER = " "
    !SERIALNUMBER = "A320-001"
    !NextApp = 1
    !archivedata = 0
    !auditind = 0
    !REGISTERIND = 0
    !CERTIND = 0
    !OTHERREPIND = 0
    ' record is not updated at this point.
    ' More information comes from STKTYPE
End With
End Sub

Private Sub Import_STKTYPE()
With rsMain
    '-- extract next cert #
    X = InStr(1, sInRec, ",", vbTextCompare)
    If X = 0 Then GoTo STKTYPE_BADDATA
    !nextcert = Val(left(sInRec, X - 1))
    '-- extract issuable shares
    X = X + 1
    Y = InStr(X, sInRec, ",", vbTextCompare)
    !TOTSTOCKS = Val(Mid(sInRec, X, Y - X))
    '-- extract Issued stocks
    X = Y + 1
    Y = InStr(X, sInRec, ",", vbTextCompare)
    !ISSSTOCKS = Val(Mid(sInRec, X, Y - X))
    '-- extract parvalue
    X = Y + 1
    Y = InStr(X, sInRec, ",", vbTextCompare)
    !PARVALUE = Val(Mid(sInRec, X, Y - X))
    '-- extract tax free limit
    X = Y + 1
    Y = Len(RTrim(sInRec))
    !TAXFREELIMIT = Val(right(sInRec, Y - X + 1))
    .Update
End With
Exit Sub
STKTYPE_BADDATA:
   MsgBox "BAD DATA/Import failed during conversion of STKTYPE. Note this error."
   cmdCancel_Click
   Exit Sub
End Sub

Private Sub Import_STKCATE()
'--
With rsMain
  .AddNew
  '-- extract category code
  X = 1
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'no code
    Exit Sub
  Else
    X = X + 1
    !catcode = Mid(sInRec, X, Y - X)
  End If
  '-- extract category description
  X = Y + 2
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'no description
    !catdesc = " "
  Else
    X = X + 1
    !catdesc = Mid(sInRec, X, Y - X)
  End If
  '-- extract taxable code
  X = Y + 2
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'no taxable code
    !cattax = True
  Else
    X = X + 1
    If Mid(sInRec, X, Y - X) = "Y" Then
       !cattax = True
    Else
       !cattax = False
    End If
 End If
 .Update
 iRecsAdded = iRecsAdded + 1  'count records added
 If iRecsAdded = 500 Then
    cnn.CommitTrans
    cnn.BeginTrans
    iRecsAdded = 0
 End If
End With
End Sub

Private Sub Import_STKTAXR()
'--
With rsMain
  .AddNew
  '-- extract tax code
  X = 1
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'no code
    Exit Sub
  Else
    X = X + 1
    !ResCode = Mid(sInRec, X, Y - X)
  End If
  '-- extract country description
  X = Y + 2
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'no description
    !RESCTRY = " "
  Else
    X = X + 1
    !RESCTRY = Mid(sInRec, X, Y - X)
  End If
  '-- extract tax rate
  X = Y + 2
  Y = Len(RTrim(sInRec))
  !taxrate = Val(right(sInRec, Y - X + 1))
 .Update
 iRecsAdded = iRecsAdded + 1  'count records added
 If iRecsAdded = 500 Then
    cnn.CommitTrans
    cnn.BeginTrans
    iRecsAdded = 0
 End If
End With
End Sub

Private Sub Import_STKCERT()
Dim iCertno As Long, iClientid As Long, sql As String, iChg As Integer
iChg = False
'-- extract shareholder number
 X = 1
 Y = InStr(X, sInRec, ",", vbTextCompare)
 iClientid = Val(left(sInRec, Y - X))
 '-- extract cert #
 X = Y + 1
 Y = InStr(X, sInRec, ",", vbTextCompare)
 iCertno = Val(Mid(sInRec, X, Y - X))
 '--
  With rsMain
   .AddNew
  !CLIENTID = iClientid
  !certno = iCertno
  '-- extract issdate
  X = InStr(Y, sInRec, sQuote, vbTextCompare)
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'no issue date
   !IssDate = Now()
  Else
   X = X + 1
   !IssDate = DateValue(Mid(sInRec, X, Y - X))
  End If
  '-- extract shares
  X = Y + 2
  Y = InStr(X, sInRec, ",", vbTextCompare)
  !shares = Val(Mid(sInRec, X, Y - X))
  '-- extract status
  X = Y + 1
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'cert transferred
   !Status = "C"
  Else
   X = X + 1
   !Status = Mid(sInRec, X, Y - X)
  End If
  !assigned = False ' users must manually load assignments.
  '-- extract remarks
  X = Y + 2
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'no remarks
    !remarks = " "
  Else
   X = X + 1
   !remarks = Mid(sInRec, X, Y - X)
  End If
  '-- extract Form/Batch
  X = Y + 2
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'no form no
    !FORMNO = " "
  Else
   X = X + 1
   !FORMNO = Mid(sInRec, X, Y - X)
  End If
  '-- extract transfer date
  X = Y + 2
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'no tfr date
  Else
   X = X + 1
   !TRNDATE = DateValue(Mid(sInRec, X, Y - X))
  End If
  '-- extract transfer batch
  X = Y + 2
  Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
  If Y - X = 1 Then 'no batch
    !TRNBATCH = " "
  Else
   X = X + 1
   !TRNBATCH = Mid(sInRec, X, Y - X)
  End If
  .Update
  iRecsAdded = iRecsAdded + 1  'count records added
    If iRecsAdded = 500 Then
       cnn.CommitTrans
       cnn.BeginTrans
       iRecsAdded = 0
    End If
End With
Exit Sub
STKCERT_BADDATA:
   MsgBox "BAD DATA/Import failed during conversion of STKCERT. Note this error." _
   & " " & sInRec
   Exit Sub
End Sub


Private Sub Import_STKJOINT()
Dim sql As String, iClient As Long
'-- extract client id
 X = 1
 Y = InStr(X, sInRec, ",", vbTextCompare)
 If Y = 0 Then ' bad data
  GoTo STKJOINT_BADDATA
 Else
  iClient = Val(left(sInRec, Y - X))
  'sql = "Select clientid from stkname where clientid = " & iClient
  'rsName.Open sql, gblFileName, , , adCmdText
  'If rsName.EOF Then 'parent name missing so skip joint
   ' rsName.Close
  '  Exit Sub
  'Else
  '  rsName.Close
  'End If
 End If
With rsMain
 .AddNew
 !CLIENTID = iClient
 X = Y + 1
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y - X = 1 Then ' no date entered use default
  !JNTSTADTE = Now
 Else
   X = X + 1
  !JNTSTADTE = DateValue(Mid(sInRec, X, Y - X))
 End If
 !jntcreated = !JNTSTADTE
 '-- extract name 1
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKJOINT_BADDATA
 If Y - X = 1 Then GoTo STKJOINT_BADDATA
 X = X + 1
 !JNTNAME1 = Mid(sInRec, X, Y - X)
 '-- extract jnt name 2 if present
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKJOINT_BADDATA
 If Y - X = 1 Then ' blank snd joint
 Else
  X = X + 1
  !JNTNAME2 = Mid(sInRec, X, Y - X)
 End If
 '-- extract beneficiary as jnt name 3
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKJOINT_BADDATA
 If Y - X = 1 Then ' blank snd joint
 Else
  X = X + 1
  !JNTNAME2 = Mid(sInRec, X, Y - X)
 End If
 .Update
 iRecsAdded = iRecsAdded + 1  'count records added
    If iRecsAdded = 500 Then
       cnn.CommitTrans
       cnn.BeginTrans
       iRecsAdded = 0
    End If
End With
Exit Sub
STKJOINT_BADDATA:
   MsgBox "BAD DATA/Import failed during conversion of STKJOINT. Note this error." _
   & " " & sInRec
Exit Sub
End Sub

Private Sub Import_STKBANK()
With rsMain
  .AddNew
  '-- extract cheque number
  X = 1
  Y = InStr(X, sInRec, ",", vbTextCompare)
 If Y = 0 Then ' bad data
  GoTo STKBANK_BADDATA
 Else
  !CHQNUM = Val(left(sInRec, Y - X))
 End If
 '-- extract client id
 X = Y
 Y = InStr(X + 1, sInRec, ",", vbTextCompare)
 If Y = 0 Then ' bad data
  GoTo STKBANK_BADDATA
 Else
  X = X + 1
  !CLIENTID = Val(Mid(sInRec, X, Y - X))
 End If
 '-- extract chq date
 X = Y + 1
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y - X = 1 Then 'bad data
  GoTo STKBANK_BADDATA
 Else
   X = X + 1
  !CHQDAT = DateValue(Mid(sInRec, X, Y - X))
 End If
 '-- extract chqamt
 X = Y + 1
 Y = InStr(X + 1, sInRec, ",", vbTextCompare)
 If Y = 0 Then ' bad data
  GoTo STKBANK_BADDATA
 Else
   X = X + 1
  !CHQAMT = Val(Mid(sInRec, X, Y - X))
 End If
 '-- extract reconind
 X = Y + 1
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y - X = 1 Then 'bad data
  GoTo STKBANK_BADDATA
 Else
   X = X + 1
   If (Mid(sInRec, X, Y - X)) = "Y" Then
     !reconind = True
   Else
     !reconind = False
   End If
 End If
 '-- extract dividend date
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y - X = 1 Then 'bad data
  GoTo STKBANK_BADDATA
 Else
   X = X + 1
  !DECDATE = DateValue(Mid(sInRec, X, Y - X))
 End If
 '-- extract new chq no
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y - X = 1 Then 'no reval chq
 Else
   X = X + 1
  !RepChqNo = Val(Mid(sInRec, X, Y - X))
 End If
 '-- extract reval date
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y - X = 1 Then 'no reval date
 Else
   X = X + 1
  !REVALDAT = DateValue(Mid(sInRec, X, Y - X))
 End If
 '-- extract recon mth
 !PAYTYP = "D"
 If !reconind = False Then GoTo UPD_ATE
  X = Y + 1
 Y = InStr(X + 1, sInRec, ",", vbTextCompare)
 If Y = 0 Then GoTo STKBANK_BADDATA ' bad data
 X = X + 1
 If Mid(sInRec, X, Y - X) = "0" Then ' not reconciled
 Else
  !FOLIOMTH = Year(!CHQDAT) & Mid(sInRec, X, Y - X)
 End If
 '--extract actual recondate
 X = Y + 1
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y - X = 1 Then 'no recon date
 Else
   X = X + 1
  !RECONDAT = DateValue(Mid(sInRec, X, Y - X))
 End If
 '--
UPD_ATE:
 .Update
  iRecsAdded = iRecsAdded + 1  'count records added
    If iRecsAdded = 500 Then
       cnn.CommitTrans
       cnn.BeginTrans
       iRecsAdded = 0
    End If
End With
Exit Sub
STKBANK_BADDATA:
   MsgBox "BAD DATA/Import failed during conversion of STKBANK. Note this error." _
          & " " & sInRec
   
   Exit Sub
End Sub

Private Sub Import_STKMNDTE()
With rsMain
  .AddNew
 '-- extract client id
 X = 1
 Y = InStr(X, sInRec, ",", vbTextCompare)
 If Y = 0 Then ' bad data
  GoTo STKMNDTE_BADDATA
 Else
  !CLIENTID = Val(left(sInRec, Y - X))
 End If
 '-- extract bank name
 X = Y + 1
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y - X = 1 Then ' no bank name entered
  GoTo STKMNDTE_BADDATA
 Else
   X = X + 1
  !MNDNAME = Mid(sInRec, X, Y - X)
 End If
  '-- extract address 1
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKMNDTE_BADDATA
 If Y - X = 1 Then
  !MNDADDR1 = "."
 Else
  X = X + 1
  !MNDADDR1 = Mid(sInRec, X, Y - X)
 End If
 '-- extract address  2 if present
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKMNDTE_BADDATA
 If Y - X = 1 Then ' no address 2
 Else
  X = X + 1
  !MNDADDR2 = Mid(sInRec, X, Y - X)
 End If
 '-- extract bank account
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKMNDTE_BADDATA
 If Y - X = 1 Then ' blank snd joint
 Else
  X = X + 1
  !MNDACNTNME = Mid(sInRec, X, Y - X)
 End If
 '-- set defaults
 !MNDSTADTE = Now
 !MNDCREATED = Now
 !MNDMET = "CHQ"
 .Update
 iRecsAdded = iRecsAdded + 1  'count records added
    If iRecsAdded = 500 Then
       cnn.CommitTrans
       cnn.BeginTrans
       iRecsAdded = 0
    End If
End With
Exit Sub
STKMNDTE_BADDATA:
   MsgBox "BAD DATA/Import failed during conversion of STKMNDTE. Note this error." _
   & " " & sInRec
   Exit Sub
End Sub

Private Sub Import_STKBRKPL()
With rsMain
  .AddNew
 '-- extract broker id
 X = 1
 Y = InStr(X, sInRec, ",", vbTextCompare)
 If Y = 0 Then ' bad data
  GoTo STKBRKPL_BADDATA
 Else
  !BROKERID = Val(left(sInRec, Y - X))
 End If
 '-- extract certificate #
 X = Y
 Y = InStr(X + 1, sInRec, ",", vbTextCompare)
 If Y = 0 Then  ' bad data
   GoTo STKBRKPL_BADDATA
 Else
   X = X + 1
  !certno = Val(Mid(sInRec, X, Y - X))
 End If
  '-- extract bal-start-of-mth
 X = Y
 Y = InStr(X + 1, sInRec, ",", vbTextCompare)
 If Y = 0 Then GoTo STKBRKPL_BADDATA
 X = X + 1
 !BalStrtPer = Val(Mid(sInRec, X, Y - X))
 '-- extract shares buy
 X = Y
 Y = InStr(X + 1, sInRec, ",", vbTextCompare)
 If Y = 0 Then GoTo STKBRKPL_BADDATA
 X = X + 1
 !ShrBuy = Val(Mid(sInRec, X, Y - X))
 '-- extract shares sell
 X = Y
 Y = InStr(X + 1, sInRec, ",", vbTextCompare)
 If Y = 0 Then GoTo STKBRKPL_BADDATA
 X = X + 1
 !ShrSell = Val(Mid(sInRec, X, Y - X))
 '-- extract shares held
 X = Y
 Y = Len(sInRec)
 If Y = X Then GoTo STKBRKPL_BADDATA
 !SHRHELD = Val(right(sInRec, Y - X))
 .Update
 iRecsAdded = iRecsAdded + 1  'count records added
    If iRecsAdded = 500 Then
       cnn.CommitTrans
       cnn.BeginTrans
       iRecsAdded = 0
    End If
End With
Exit Sub
STKBRKPL_BADDATA:
   MsgBox "BAD DATA/Import failed during conversion of STKBRKPL. Note this error." _
   & " " & sInRec
   Exit Sub
End Sub

Private Sub Import_STKCRTBK()
With rsMain
  .AddNew
 '-- extract broker id
 X = 1
 Y = InStr(X, sInRec, ",", vbTextCompare)
 If Y = 0 Then ' bad data
  GoTo STKCRTBK_BADDATA
 Else
  !BROKERID = Val(left(sInRec, Y - X))
 End If
 '-- extract form number
 X = Y + 1
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKCRTBK_BADDATA
 If Y - X = 1 Then GoTo STKCRTBK_BADDATA
 X = X + 1
 !FORMNO = Mid(sInRec, X, Y - X)
 '-- extract date certified
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKCRTBK_BADDATA
 If Y - X = 1 Then GoTo STKCRTBK_BADDATA
 X = X + 1
 !DTECRTFY = DateValue(Mid(sInRec, X, Y - X))
 '-- extract to broker id
 X = Y + 1
 Y = InStr(X + 1, sInRec, ",", vbTextCompare)
 If Y = 0 Then  ' bad data
   GoTo STKCRTBK_BADDATA
 Else
   X = X + 1
  !TOBROKERID = Val(Mid(sInRec, X, Y - X))
 End If
 '-- extract shares held
 X = Y
 Y = InStr(X + 1, sInRec, ",", vbTextCompare)
 If Y = 0 Then GoTo STKCRTBK_BADDATA
 X = X + 1
 !shares = Val(Mid(sInRec, X, Y - X))
 '-- extract STATUS
 X = Y + 1
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKCRTBK_BADDATA
 X = X + 1
 !Status = Mid(sInRec, X, Y - X)
 '-- extract date changed
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKCRTBK_BADDATA
 X = X + 1
 !STACHGDTE = DateValue(Mid(sInRec, X, Y - X))
 '-- extract Batch
 X = Y + 2
 Y = InStr(X + 1, sInRec, sQuote, vbTextCompare)
 If Y = 0 Then GoTo STKCRTBK_BADDATA
 X = X + 1
 !batch = Mid(sInRec, X, Y - X)
 .Update
 iRecsAdded = iRecsAdded + 1  'count records added
    If iRecsAdded = 500 Then
       cnn.CommitTrans
       cnn.BeginTrans
       iRecsAdded = 0
    End If
End With
Exit Sub
STKCRTBK_BADDATA:
   MsgBox "BAD DATA/Import failed during conversion of STKCRTBK. Note this error." _
   & " " & sInRec
   Exit Sub
End Sub


