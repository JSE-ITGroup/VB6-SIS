VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS057 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stockholder Verification"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7320
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   7200
      _Version        =   196616
      DataMode        =   2
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   1958
      Columns(0).Caption=   "Client Id"
      Columns(0).Name =   "ClientId"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   10
      Columns(1).Width=   5477
      Columns(1).Caption=   "Client Name"
      Columns(1).Name =   "CliName"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   50
      Columns(2).Width=   2487
      Columns(2).Caption=   "Control Shares"
      Columns(2).Name =   "CShares"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   10
      Columns(3).Width=   2566
      Columns(3).Caption=   "Cert Shares"
      Columns(3).Name =   "Shares"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   3
      Columns(3).FieldLen=   10
      _ExtentX        =   12700
      _ExtentY        =   3836
      _StockProps     =   79
      ForeColor       =   0
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Begin"
      Default         =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   5280
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      ToolTipText     =   "Returns to main menu"
      Top             =   4080
      Width           =   975
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      Height          =   300
      Index           =   0
      Left            =   6360
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      ToolTipText     =   "Returns to main menu"
      Top             =   4080
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
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   7200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Display program Information"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
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
      Width           =   7335
   End
End
Attribute VB_Name = "frmSIS057"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, iEOF As Integer
Dim sRowinfo As String
Dim rsCmp As ADODB.Recordset
Dim rsName As ADODB.Recordset
Dim rsCert As ADODB.Recordset
Dim rsAcnt As ADODB.Recordset
Dim OpenErr As Integer
Dim iOpenCmp As Integer
Dim iOpenName As Integer
Dim iOpenCert As Integer
Dim iClient As Long
Dim iShares As Long
Dim sCliName As String
Dim iConShares As Long, iCertTot As Long
Dim iRecs As Long, sConst As String, Qview As String
Private Sub cmdBtn_Click(Index As Integer)
Dim sql As String, iRep As Integer
On Error GoTo cmdBtn_Click_Err
Select Case Index
Case 0 'Cancel
    If iOpenCmp = True Then rsCmp.Close
    If iOpenName = True Then rsName.Close
    If iOpenCert = True Then rsCert.Close
    '--
    Set rsCmp = Nothing
    Set rsName = Nothing
    Set rsCert = Nothing
    Set frmSIS057 = Nothing
    iEOF = True
    Unload Me
    frmSIS053.Visible = True
   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
Case 1 'Perform Verifications
iRep = MsgBox("Select Yes for two pass verification;" _
       & vbCrLf & "No for one pass verification. ", vbInformation + vbYesNo, _
       "Stockholder Verification")
    
'--
' wait & hourglass message
'--------------
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.Refresh
OpenFiles
If OpenErr = True Then GoTo cmdBtn_Click_Exit
If iEOF = True Then GoTo cmdBtn_Click_Exit
sConst = "Verification now being performed "
lbl = sConst
sConst = sConst & "for "
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
    rsName.MoveFirst
    lbl.Visible = True
    While Not .EOF
       lbl = sConst & !CLIENTID & " - First Pass"
       lbl.Refresh
       ProgressBar1.Value = iRecs
       If iClient = 0 Then
         iClient = !CLIENTID
       End If
       If iClient <> !CLIENTID Then
         CalcPay
         iShares = 0
         iClient = !CLIENTID
       End If
       iShares = iShares + !shares
       iRecs = iRecs + 1
       .MoveNext
    Wend
 End With
 iEOF = True
 CalcPay
 If iRep = vbYes Then
 '-- Start second pass using Stkname to drive
 '--  the verification
 '--------------------------------------------
 
 If iOpenName = True Then
     rsName.Close
      iOpenName = 0
 End If
   If iOpenCert = True Then
     rsCert.Close
     iOpenCert = 0
   End If
   '--
   sql = "Select sum(shares) as crtshares, clientid " _
             & "from CERTMST GROUP by CLIENTID, STATUS" _
             & " having STATUS <> 'C'" _
             & " order by clientid "
   rsCert.Open sql, gblFileName, , , adCmdText
   iOpenCert = True
 '--
 rsName.Open Qview, gblFileName, , , adCmdText
 iOpenName = True
 '--
 Process_name
 End If
 PrintGrid
Case Else
End Select
cmdBtn_Click_Exit:
Exit Sub
cmdBtn_Click_Err:
  csvShowError "SIS057/Load"
  cmdBtn_Click (0)
End Sub
Private Sub Form_Load()
'--
Dim i As Integer
Dim strTmp As String
On Error GoTo FL_ERR
iEOF = False
iConShares = 0: iCertTot = 0
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
   grd.RemoveAll
   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
FL_Exit:
  Exit Sub
FL_ERR:
  csvShowError "SIS057/Load"
  Unload Me
End Sub

Private Sub InitProgressBar(max As Long)
    ProgressBar1.Min = 0
    ProgressBar1.max = max
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.Min

End Sub

Private Sub Form_Unload(Cancel As Integer)
If iEOF = False Then
  Cancel = -1
End If
End Sub

Private Sub OpenFiles()
Dim iErr As Integer
On Error GoTo OpenFiles_Err
OpenErr = False
iOpenCmp = False
iOpenName = False
iOpenCert = False
'--
Set rsCmp = New ADODB.Recordset
Set rsName = New ADODB.Recordset
Set rsCert = New ADODB.Recordset
'---------------------------
'-- Open Certificate View --
'---------------------------
Qview = "SELECT CertNo, ClientId, " _
        & "shares " _
        & " from CERTMST " _
        & " where (Status <> 'C') " _
        & " order by ClientId, CertNo"
rsCert.Open Qview, gblFileName, , , adCmdText
iOpenCert = True
If rsCert.EOF And rsCert.BOF Then
  iErr = 119
  csvShowUsrErr iErr, "Stockholder Verification"
  rsCert.Close
  iOpenCert = False
  OpenErr = True
  cmdBtn_Click (0)
  GoTo OpenFiles_Exit
 End If
'--
Qview = "SELECT ClientID, CLINAME, Shares from STKNAME order by CLIENTID"
rsName.Open Qview, gblFileName, , , adCmdText
iOpenName = True
'--
OpenFiles_Exit:
   Exit Sub
OpenFiles_Err:
  csvShowError "SIS057/OpenFiles"
  OpenErr = True
  cmdBtn_Click (0)
  GoTo OpenFiles_Exit
End Sub

Private Static Sub CalcPay()
Dim sql As String
Dim iCshares As Long
On Error GoTo Calcpay_Err
'---
   'get details from stkname file
   '-----------------------------
   'sql = "SELECT CLINAME, SHARES FROM STKNAME WHERE " _
   '          & "CLIENTID = " & iClient
   'rsName.Open sql, gblFileName, , , adCmdText
   'iOpenName = True
   If iOpenName = True Then
    If Not rsName.EOF Then
       While rsName!CLIENTID < iClient
          rsName.MoveNext
       Wend
       If Not rsName.EOF And rsName!CLIENTID = iClient Then
          sCliName = rsName!cliname
          iCshares = rsName!shares
       Else
          sCliName = "No Name & Address record found"
          iCshares = 0
       End If
     End If
     If rsName.EOF Then
        iOpenName = False
        rsName.Close
     End If
   Else
     sCliName = "No Name & Address record found"
     iCshares = 0
   End If
   '---
   iConShares = iConShares + iCshares
   iCertTot = iCertTot + iShares
   If iShares <> iCshares Then
       sRowinfo = iClient & Chr(9) & sCliName & Chr(9) & iCshares & Chr(9) & iShares
       grd.AddItem sRowinfo
   End If
   If iEOF = True Then
      sRowinfo = 0 & Chr(9) & "Total Register Based on Certs" & Chr(9) & iConShares & Chr(9) _
                 & iCertTot
      grd.AddItem sRowinfo
   End If
calcpay_Exit:
  Exit Sub
Calcpay_Err:
  csvShowError "SIS057/calcpay"
  cmdBtn_Click (0)
End Sub


Public Sub PrintGrid()
grd.PrintData ssPrintAllRows, False, True
cmdBtn_Click (0)
End Sub



Private Sub Process_name()
Dim sql As String
iConShares = 0: iCertTot = 0
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
 iRecs = 1: iClient = 0: iShares = 0
 '--
 rsCert.MoveFirst
 With rsName
    .MoveFirst
    While Not .EOF
       lbl = sConst & !CLIENTID & " - Final Pass"
       lbl.Refresh
       ProgressBar1.Value = iRecs
       
       '--
       If iOpenCert = True Then
          If Not rsCert.EOF Then
             While rsCert!CLIENTID < !CLIENTID
                rsCert.MoveNext
             Wend
             If Not rsCert.EOF And rsCert!CLIENTID = !CLIENTID Then
               iShares = rsCert!crtshares
             Else
               iShares = 0
             End If
          End If
          If rsCert.EOF Then
            iOpenCert = False
            rsCert.Close
          End If
      Else
         iShares = 0
      End If
      iCertTot = iCertTot + iShares
       iConShares = iConShares + !shares
       If iShares <> !shares Then
          sRowinfo = !CLIENTID & Chr(9) & !cliname & Chr(9) & !shares & Chr(9) & iShares
          grd.AddItem sRowinfo
       End If
       iRecs = iRecs + 1
       .MoveNext
    Wend
    .Close
    iOpenName = 0
 End With
 sRowinfo = 0 & Chr(9) & "Total Register Based on Account" & Chr(9) & iConShares & Chr(9) _
                 & iCertTot
 grd.AddItem sRowinfo
 iEOF = True
End Sub

Private Sub grd_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
Dim sHeader As String
rsCmp.Open "COMPANY", gblFileName, , adCmdTable
sHeader = "Date <date>" & Chr(9) & "STOCKHOLDER VERIFICATION for " & rsCmp!compname & Chr(9) & "Page: <page number> "
rsCmp.Close
ssPrintInfo.PageHeader = sHeader


End Sub
