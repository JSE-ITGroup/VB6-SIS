VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmSIS102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate RI Stock Allotment"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS102.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
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
      Top             =   2400
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
   Begin SSDataWidgets_A.SSDBOptSet Opt 
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   840
      Width           =   2325
      _Version        =   196611
      _ExtentX        =   4207
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Main Ledger"
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      IndexSelected   =   0
      NumberOfButtons =   2
      Buttons.Button(0).OptionValue=   "-1"
      Buttons.Button(0).Caption=   "Main Ledger"
      Buttons.Button(0).Mnemonic=   77
      Buttons.Button(0).Value=   -1  'True
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   74
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   76
      Buttons.Button(0).PictureRight=   75
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   158
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
      Buttons.Button(1).OptionValue=   "0"
      Buttons.Button(1).Caption=   "JCSD Ledger"
      Buttons.Button(1).Mnemonic=   74
      Buttons.Button(1).TextLeft=   15
      Buttons.Button(1).TextTop=   16
      Buttons.Button(1).TextRight=   78
      Buttons.Button(1).TextBottom=   30
      Buttons.Button(1).ButtonTop=   16
      Buttons.Button(1).ButtonRight=   13
      Buttons.Button(1).ButtonBottom=   29
      Buttons.Button(1).PictureLeft=   80
      Buttons.Button(1).PictureTop=   16
      Buttons.Button(1).PictureRight=   79
      Buttons.Button(1).PictureBottom=   30
      Buttons.Button(1).ButtonToColTop=   16
      Buttons.Button(1).ButtonToColRight=   158
      Buttons.Button(1).ButtonToColBottom=   30
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
      Left            =   120
      TabIndex        =   3
      Top             =   1800
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
Attribute VB_Name = "frmSIS102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, iEOF As Integer
Dim rsMain As ADODB.Recordset
Dim rsBon As ADODB.Recordset
Dim rsCert As ADODB.Recordset
Dim OpenErr As Integer
Dim iOpenMain As Integer
Dim iOpenBon As Integer
Dim iOpenCert As Integer
Dim iCommit As Integer
Dim qDMLQry As String
Dim iClient As Long
Dim iShares As Long
Private Sub cmdBtn_Click(Index As Integer)
Dim iRecs As Long, sConst As String
On Error GoTo cmdBtn_Click_Err
Select Case Index
Case 0 'Cancel
    If iOpenBon = True Then rsBon.Close
    If iOpenMain = True Then rsMain.Close
    If iOpenCert = True Then rsCert.Close
    '--
    Set rsBon = Nothing
    Set rsMain = Nothing
    Set rsCert = Nothing
    '''set cnn = nothing
    Set frmSIS102 = Nothing
    iEOF = True
    Unload Me
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
       lbl.Caption = sConst & !ClientId
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
 cnn.CommitTrans
 cmdBtn_Click (0)
Case Else
End Select
cmdBtn_Click_Exit:
Exit Sub
cmdBtn_Click_Err:
  MsgBox "SIS102/Load"
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
  MsgBox "SIS102/Load"
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
iOpenCert = False
iOpenBon = False
iOpenMain = False
'__
Set rsBon = New ADODB.Recordset
Set rsCert = New ADODB.Recordset
cnn.Open
'-----------------------
'-- open BONREF table --
'-----------------------
rsBon.Open "BONUSREF", cnn, , , adCmdTable
iOpenBon = True
If rsBon.EOF = True Then
     iErr = 180
     csvShowUsrErr iErr, "Calculate Rights Issue"
     rsBon.Close
     iOpenBon = False
     GoTo OpenFiles_Close
End If
If Opt.IndexSelected = 0 Then
   '---------------------------
   '-- Open Certificate View --
   '---------------------------
   qView = "SELECT CertNo, ClientId, shares, issdate " _
        & "from CERTMST " _
        & "where Status <> 'C' and " _
        & "IssDate <= " & Format(rsBon!RECDAT, "mm/dd/yyyy") & " " _
        & "and shares > 0 order by " _
        & "ClientId, CertNo"
Else
   qView = "SELECT GR8NIN as ClientId, GR8CBL as shares, " _
   & "GR8NAM as CliName from JCSDSUB " _
          & "order by GR8NAM, GR8NIN"
End If
rsCert.Open qView, cnn, , , adCmdText
iOpenCert = True
If rsCert.EOF And rsCert.BOF Then
    iErr = 119
    csvShowUsrErr iErr, "Calculate Rights Issue"
    rsBon.Close
    iOpenBon = False
    rsCert.Close
    iOpenCert = False
    GoTo OpenFiles_Close
End If
'-- open Rights Issue & clear existing data
'--------------------------------------
If Opt.IndexSelected = 0 Then
  lbl = "Deleting Previous Rights Issue Information"
  lbl.Visible = True
  '----------
  qDMLQry = "DELETE FROM STKRIWRK"
  X = csvADODML(qDMLQry, cnn)
End If
Set rsMain = New ADODB.Recordset
sql = "Select * from STKRIWRK where Ledger = 'S'"
rsMain.Open sql, cnn, adOpenDynamic, adLockPessimistic, adCmdText
iOpenMain = True
rsMain.Requery
If rsMain.EOF And rsMain.BOF Then
Else
   MsgBox "The JCSD Allocations have already been calculated." _
          & vbCrLf & "Rerun allocations for main legder", vbCritical, "RI Allocations"
   
   GoTo OpenFiles_Close
 End If
   
'--
OpenFiles_Exit:
   Exit Sub
OpenFiles_Close:
   cmdBtn_Click (0)
   iEOF = True
   Unload Me
   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
   GoTo OpenFiles_Exit
   
OpenFiles_Err:
  MsgBox "SIS102/OpenFiles"
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
  !ClientId = iClient
  !shares = iShares
  !offer = Int(iShares * rsBon!STKSTO / rsBon!STKBASE)
  !UNUSED = !shares - (!offer * rsBon!STKBASE / rsBon!STKSTO)
  !Cost = !offer * rsBon!RIPRICE
  If Opt.IndexSelected = 0 Then
    !Ledger = "M"   '-- Main Ledger
  Else
    !Ledger = "S"   '-- JCSD sub ledger
  End If
  .Update
End With
'--
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
  MsgBox "SIS102/calcpay"
  cmdBtn_Click (0)
End Sub

