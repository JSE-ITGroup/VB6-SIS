VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSIS092 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Stock Splits"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS092.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
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
      Left            =   360
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
Attribute VB_Name = "frmSIS092"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, iEOF As Integer
Dim rsMain As ADODB.Recordset
Dim rsBon As ADODB.Recordset
Dim rsName As ADODB.Recordset
Dim rsComp As ADODB.Recordset
Dim iOpenComp As Integer
Dim OpenErr As Integer
Dim iOpenMain As Integer
Dim iOpenBon As Integer
Dim iOpenName As Integer
Dim cmdChange As ADODB.Command
Dim iCommit As Integer, icert As Long
Dim qDMLQry As String
Private Sub cmdBtn_Click(Index As Integer)
Dim iRecs As Long, sConst As String
On Error GoTo cmdBtn_Click_Err
Select Case Index
Case 0 'Cancel
    If iOpenBon = True Then rsBon.Close
    If iOpenMain = True Then rsMain.Close
    If iOpenName = True Then rsName.Close
    If iOpenComp = True Then rsComp.Close
    '--
    Set rsBon = Nothing
    Set rsMain = Nothing
    Set rsName = Nothing
    Set rsComp = Nothing
    '''set cnn = nothing
    Set frmSIS092 = Nothing
    iEOF = True
    Unload Me
    frmSIS090.Visible = True
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
If rsName.RecordCount < 0 Then
    iRecs = 0
    With rsName
       .MoveFirst
       While Not .EOF
          iRecs = iRecs + 1
          .MoveNext
       Wend
    End With
 Else
  iRecs = rsName.RecordCount
 End If
 '--
 InitProgressBar (iRecs)
 ProgressBar1.Visible = True
 iRecs = 1
 '--
 With rsName
    .MoveFirst
    While Not .EOF
       lbl = sConst & !ClientId
       lbl.Refresh
       ProgressBar1.Value = iRecs
       CalcPay
       icert = icert + 1
       iRecs = iRecs + 1
       .MoveNext
    Wend
 End With
 rsComp!nextcert = icert
 rsComp.Update
 cnn.CommitTrans
 cmdBtn_Click (0)
Case Else
End Select
cmdBtn_Click_Exit:
Exit Sub
cmdBtn_Click_Err:
  MsgBox "SIS092/Load"
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
  MsgBox "SIS092/Load"
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
'''cnn.Open cnn
OpenErr = False
iOpenName = False
iOpenBon = False
iOpenMain = False
iOpenComp = False
'__
Set rsBon = New ADODB.Recordset
Set rsMain = New ADODB.Recordset
Set rsName = New ADODB.Recordset
Set rsComp = New ADODB.Recordset
cnn.Open
'-----------------------
'-- open BONREF table --
'-----------------------
rsBon.Open "BONUSREF", cnn, , , adCmdTable
iOpenBon = True
If rsBon.EOF = True Then
     iErr = 180
     csvShowUsrErr iErr, "Calculate Bonus"
     rsBon.Close
     iOpenBon = False
     GoTo OpenFiles_Close
End If
'--
rsComp.Open "Company", cnn, adOpenDynamic, adLockPessimistic, adCmdTable
iOpenComp = True
icert = rsComp!nextcert
'---------------------------
'-- Open STKNAME View --
'---------------------------
qView = "SELECT ClientId, shares " _
        & "from STKNAME " _
        & "where shares > 0 " _
        & " order by CliName,ClientId"
rsName.Open qView, cnn, , , adCmdText
iOpenName = True
If rsName.EOF And rsName.BOF Then
    iErr = 119
    csvShowUsrErr iErr, "Calculate Stock Split"
    rsBon.Close
    iOpenBon = False
    rsName.Close
    iOpenName = False
    GoTo OpenFiles_Close
End If
'-- open Bonus & clear existing data
'--------------------------------------
lbl = "Deleting Previous Split Information"
lbl.Visible = True
'----------
qDMLQry = "DELETE FROM STKBONUS"
X = csvADODML(qDMLQry, cnn)
rsMain.Open "STKBONUS", cnn, adOpenDynamic, adLockPessimistic, adCmdTable
iOpenMain = True
'--
OpenFiles_Exit:
   Exit Sub
OpenFiles_Close:
   Set rsComp = Nothing
   Set rsBon = Nothing
   Set rsMain = Nothing
   Set rsName = Nothing
   '''set cnn = nothing
   Set frmSIS092 = Nothing
   iEOF = True
   Unload Me
   frmSIS090.Show
   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
   GoTo OpenFiles_Exit
   
OpenFiles_Err:
  MsgBox "SIS092/OpenFiles"
  OpenErr = True
  cmdBtn_Click (0)
  GoTo OpenFiles_Exit
  
End Sub
Private Static Sub CalcPay()
Dim iWork As Double
On Error GoTo Calcpay_Err
'---
If iCommit = 0 Or iCommit = 1000 Then
     cnn.BeginTrans
     iCommit = 0
End If
With rsMain
  .AddNew
  !ClientId = rsName!ClientId
  !certno = icert
  !BonIss = rsBon!RECDAT
  !shares = rsName!shares
  !allocated = Int(!shares * rsBon!STKSTO / rsBon!STKBASE)
  !unallocated = (!shares * rsBon!STKSTO / rsBon!STKBASE) - !allocated
  .Update
End With
'--
iCommit = iCommit + 1
If iCommit = 1000 Then
        cnn.CommitTrans
End If
calcpay_Exit:
  Exit Sub
Calcpay_Err:
  MsgBox "SIS092/calcpay"
  cmdBtn_Click (0)
End Sub

