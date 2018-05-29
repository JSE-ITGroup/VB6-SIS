VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form ImpRegisterExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import New Register Details"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "ImpRegisterExcel.frx":0000
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
Attribute VB_Name = "ImpRegisterExcel"
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
Dim iAcnt As Double, iCBL As Long, iRat As Single, WrkAddr As String
Dim sNam As String, sAD1 As String, sAD2 As String, sAd3 As String, sTax As String
Dim CliType As String, sAD4 As String, sAD5 As String
Dim CompInd As String
Dim Jnt As Integer
Dim CatCode As String
Dim sLen As Integer, pos As Integer
sErrMsgL1 = "Update failed while clearing balances of "
sErrMsgL2 = " Note this error."
'--
sMsg = "WARNING: This procedure will will delete existing Records"
sMsg = sMsg & "  then recreate them from the XL Import file. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Building Register"
iResp = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If iResp = vbNo Then
  cmdCancel_Click
  Exit Sub
End If
'--

lbl.Caption = "Clearing existing files"
lbl.Visible = True
'iResp = RunSP(SpCon, "usp_DeleteRegisterTables", 0)
GoTo Import_Cert
'--
'''On Error GoTo ImpJCSD_Fail
sErrMsg = "Procedure failed when trying to activate EXCEL"
Set AppExcl = CreateObject("Excel.application")
lbl.Caption = "Recreating Ledger for"

txtfile = frmMDI.CmnDialog.filename
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Creating Records..."
frmMDI.txtStatusMsg.Refresh
'--
sErrMsg = "Procedure failed/check if the client changed the format of the XL Sheet"
With AppExcl
  .Workbooks.Open (txtfile)
  Set curCell = .Worksheets(1).Range("A1")
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
  Set curCell = .Worksheets(1).Range("A1")
  Do While Not IsEmpty(curCell)
    sErrMsg = "Procedure failed while formatting/check if the client changed the format of the XL Sheet"
    Set nextCol = curCell.Offset(0, 2) '- ClientId
    WrkAddr = nextCol.Value
    If Len(WrkAddr) > 9 Then
       iAcnt = CDbl(Mid(WrkAddr, 1, 9))
    Else
       iAcnt = CDbl(WrkAddr)
    End If
    
    'iAcnt = nextCol.Value
    lbl.Caption = "Recreating Ledger for " & iAcnt
    lbl.Refresh
    Set nextCol = curCell.Offset(0, 3) '- Company Indicator
    CompInd = nextCol.Value
    Set nextCol = curCell.Offset(0, 4) '- account name
    sNam = nextCol.Value
    sNam = Trim(sNam)
    Set nextCol = curCell.Offset(0, 5) '-shares
    iCBL = nextCol.Value
    
    Set nextCol = curCell.Offset(0, 7) '-address1
    sAD1 = nextCol.Value
    Set nextCol = curCell.Offset(0, 8) '-address2
    sAD2 = nextCol.Value
    Set nextCol = curCell.Offset(0, 9) '-address3
    sAd3 = nextCol.Value
    Set nextCol = curCell.Offset(0, 10) '-address4
    sAD4 = nextCol.Value
    
    Set nextCol = curCell.Offset(0, 11) '-address5
    sAD5 = nextCol.Value
    
    Set nextCol = curCell.Offset(0, 13) '-Country Code
    WrkAddr = nextCol.Value
    
    Set nextCol = curCell.Offset(0, 14) '-Joint Holder Indicator
    Jnt = nextCol.Value
    
    If sAD5 = "Stockbroker" Or sAD5 = "Stockbrokers" Then
       CatCode = "SB"
    Else
       CatCode = "SH"
    End If
    
    If CompInd = "N" Then
       CliType = "P"
    Else
       CliType = "C"
    End If
     
    If CliType = "P" Then
       sAD5 = TrimSpace(sNam)
       pos = Len(sAD5)
       iResp = InStr(1, sAD5, " ")
       sNam = Mid(sAD5, 1, iResp - 1) & "," & Mid(sAD5, iResp + 1, pos - iResp)
    End If
       
AddOthers:
    If WrkAddr = "US" Or WrkAddr = "USA" Then
       sTax = "US"
       GoTo Cert_Time
    End If
    
    If WrkAddr = "JM" Or WrkAddr = "JAM" Then
       If CliType = "P" Then
          sTax = "JA"
       Else
          sTax = "JC"
       End If
       GoTo Cert_Time
    End If
    
    If WrkAddr = "CA" Or WrkAddr = "CAN" Then
       sTax = "CN"
       GoTo Cert_Time
    End If
    
    If WrkAddr = "GB" Or WrkAddr = "ENG" Then
       sTax = "UK"
       GoTo Cert_Time
    End If
    
    If WrkAddr = "BB" Or WrkAddr = "BAR" Then
       sTax = "BB"
       GoTo Cert_Time
    End If
    
    If WrkAddr = "BAH" Then
       sTax = "BS"
       GoTo Cert_Time
    End If
    
    If WrkAddr = "BZ" Or WrkAddr = "BLZ" Then
       sTax = "BZ"
       GoTo Cert_Time
    End If
    
    If WrkAddr = "CYM" Or WrkAddr = "KY" Then
       sTax = "KY"
       GoTo Cert_Time
    End If
    
     If WrkAddr = "DE" Then
       sTax = "DE"
       GoTo Cert_Time
    End If
    
    If WrkAddr = "EGT" Then
       sTax = "EG"
       GoTo Cert_Time
    End If
    
    If WrkAddr = "MA" Then
       sTax = "SP"
       GoTo Cert_Time
    End If
    
    If WrkAddr = "SC" Then
       sTax = "SE"
       GoTo Cert_Time
    End If
   
    If WrkAddr = "T&T" Or WrkAddr = "TT" Then
       sTax = "TT"
       GoTo Cert_Time
    End If
    
    sTax = "JA"
    
Cert_Time:
    iResp = RunSP(SpCon, "usp_ImportNewRegisterData", 0, iAcnt, CliType, sNam, iCBL, sTax, sAD1, sAD2, sAd3, sAD4, Jnt, CatCode)
    
Commit_Check:
    iRecs = iRecs + 1
    ProgressBar1.Value = iRecs

    Set nextCell = curCell.Offset(1, 0)
    Set curCell = nextCell
  Loop
  sErrMsg = "Procedure failed in closing EXCEL spread sheet"
 .Workbooks.Close
 AppExcl.Quit
End With
'--
Import_Cert:
ImportCert
'ImportStkJoint
'ImportDivRef
'ImportReconData
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

Function TrimSpace(strInput As String) As String
   ' This procedure trims extra space from any part of
   ' a string.

   Dim astrInput()     As String
   Dim astrText()      As String
   Dim strElement      As String
   Dim lngCount        As Long
   Dim lngIncr         As Long
   
   ' Split passed-in string.
   astrInput = Split(strInput)
   
   ' Resize second array to be same size.
   ReDim astrText(UBound(astrInput))
   
   ' Initialize counter variable for second array.
   lngIncr = LBound(astrInput)
   ' Loop through split array, looking for
   ' non-zero-length strings.
   For lngCount = LBound(astrInput) To UBound(astrInput)
      strElement = astrInput(lngCount)
      If Len(strElement) > 0 Then
         ' Store in second array.
         astrText(lngIncr) = strElement
         lngIncr = lngIncr + 1
      End If
   Next
   ' Resize new array.
   ReDim Preserve astrText(LBound(astrText) To lngIncr - 1)

   ' Join new array to return string.
   TrimSpace = Join(astrText)
End Function
Private Sub ImportCert()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
Dim iAcnt As Double, iCBL As Long
Dim sLen As Integer, pos As Long
Dim certno As Long
Dim IssDate As Date
Dim CanCert As Integer
Dim WrkStr As String
Dim Mth As String

sErrMsgL1 = "Update failed while clearing balances of "
sErrMsgL2 = " Note this error."
'--
sMsg = "WARNING: This procedure will will delete existing Records"
sMsg = sMsg & "  then recreate them from the XL Import file. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Building Register"
'--
'--
'''On Error GoTo ImpJCSD_Fail
frmMDI.CmnDialog.DialogTitle = "Import Register Data XL File"
frmMDI.CmnDialog.Filter = "XLS(*.xls)|*.xls"
frmMDI.CmnDialog.DefaultExt = "XLS"
frmMDI.CmnDialog.ShowOpen
If Len(frmMDI.CmnDialog.filename) > 0 Then
     txtfile = frmMDI.CmnDialog.filename
Else
    GoTo Exit_Sub
End If
sErrMsg = "Procedure failed when trying to activate EXCEL"
Set AppExcl = CreateObject("Excel.application")
lbl.Caption = "Creating Certificate for"

Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Creating Records..."
frmMDI.txtStatusMsg.Refresh
'--
sErrMsg = "Procedure failed/check if the client changed the format of the XL Sheet"
With AppExcl
  .Workbooks.Open (txtfile)
  Set curCell = .Worksheets(1).Range("A1")
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
  Set curCell = .Worksheets(1).Range("A1")
  Do While Not IsEmpty(curCell)
    sErrMsg = "Procedure failed while formatting/check if the client changed the format of the XL Sheet"
    Set nextCol = curCell.Offset(0, 2) '- ClientId
    WrkStr = nextCol.Value
    If Len(WrkStr) > 9 Then
       iAcnt = CDbl(Mid(WrkStr, 1, 9))
    Else
       iAcnt = CDbl(WrkStr)
    End If
    
    'iAcnt = nextCol.Value
    lbl.Caption = "Creating Certificates for " & iAcnt
    lbl.Refresh
    Set nextCol = curCell.Offset(0, 3) '- Certificate Number
    certno = nextCol.Value
    Set nextCol = curCell.Offset(0, 4) '- Issue Date
    WrkStr = nextCol.Value
    If Len(WrkStr) < 6 Then
       WrkStr = "0" & WrkStr
    End If
    
    pos = CInt(Mid(WrkStr, 3, 2))
    
    Select Case pos
           Case 1
                Mth = "Jan"
           Case 2
                Mth = "Feb"
           Case 3
                Mth = "Mar"
           Case 4
                Mth = "Apr"
           Case 5
                Mth = "May"
           Case 6
                Mth = "Jun"
           Case 7
                Mth = "Jul"
           Case 8
                Mth = "Aug"
           Case 9
                Mth = "Sep"
           Case 10
                Mth = "Oct"
           Case 11
                Mth = "Nov"
           Case 12
                Mth = "Dec"
     End Select
       
    sErrMsg = Mid(WrkStr, 1, 2) & "-" & Mth & "-" & Mid(WrkStr, 5, 2)
    IssDate = CDate(sErrMsg)
    
    Set nextCol = curCell.Offset(0, 5) '- Cancel Date
    WrkStr = nextCol.Value
    If WrkStr = "0" Then
       CanCert = 0
    Else
       CanCert = 1
    End If
    
    Set nextCol = curCell.Offset(0, 6) '- Shares
    iCBL = nextCol.Value
    
    Set nextCol = curCell.Offset(0, 7) '- Duplicate Indicator
    pos = nextCol.Value
    'If pos <> 0 Then
    '   GoTo Commit_Check
    'End If
    
Cert_Time:
    iResp = RunSP(SpCon, "usp_ImportCertificateData", 0, iAcnt, certno, Format(IssDate, "dd-mmm-yyyy"), iCBL, CanCert)
    
Commit_Check:
    iRecs = iRecs + 1
    ProgressBar1.Value = iRecs

    Set nextCell = curCell.Offset(1, 0)
    Set curCell = nextCell
  Loop
  sErrMsg = "Procedure failed in closing EXCEL spread sheet"
 .Workbooks.Close
 AppExcl.Quit
End With
Exit_Sub:
Exit Sub
End Sub
Private Sub ImportStkJoint()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
Dim iAcnt As Double, iCBL As Long
Dim sLen As Integer, pos As Long
Dim Jnt As String
Dim IssDate As Date
Dim CanDate As Date
Dim WrkStr As String
Dim ComPId As String

sErrMsgL1 = "Update failed while clearing balances of "
sErrMsgL2 = " Note this error."
'--
sMsg = "WARNING: This procedure will will delete existing Records"
sMsg = sMsg & "  then recreate them from the XL Import file. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Building Register"
'--
'--
'''On Error GoTo ImpJCSD_Fail
frmMDI.CmnDialog.DialogTitle = "Import Register Data XL File"
frmMDI.CmnDialog.Filter = "XLS(*.xls)|*.xls"
frmMDI.CmnDialog.DefaultExt = "XLS"
frmMDI.CmnDialog.ShowOpen
If Len(frmMDI.CmnDialog.filename) > 0 Then
     txtfile = frmMDI.CmnDialog.filename
Else
    GoTo Exit_Sub
End If
sErrMsg = "Procedure failed when trying to activate EXCEL"
Set AppExcl = CreateObject("Excel.application")
lbl.Caption = "Creating Joint Holder Records for"

Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Creating Records..."
frmMDI.txtStatusMsg.Refresh
'--
sErrMsg = "Procedure failed/check if the client changed the format of the XL Sheet"
With AppExcl
  .Workbooks.Open (txtfile)
  Set curCell = .Worksheets(1).Range("A1")
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
  Set curCell = .Worksheets(1).Range("A1")
  Do While Not IsEmpty(curCell)
    sErrMsg = "Procedure failed while formatting/check if the client changed the format of the XL Sheet"
    Set nextCol = curCell.Offset(0, 2) '- ClientId
    WrkStr = nextCol.Value
    If Len(WrkStr) > 9 Then
       iAcnt = CDbl(Mid(WrkStr, 1, 9))
    Else
       iAcnt = CDbl(WrkStr)
    End If
    
    'iAcnt = nextCol.Value
    lbl.Caption = "Creating Joint Holder Record for " & iAcnt
    lbl.Refresh
    Set nextCol = curCell.Offset(0, 4) '- Company Indicator
    ComPId = nextCol.Value
    
    Set nextCol = curCell.Offset(0, 5) '- Joint Holder Name
    Jnt = nextCol.Value
    If Len(Jnt) < 1 Then
       GoTo Commit_Check
    End If
    
    If ComPId = "N" Then
       WrkStr = TrimSpace(Jnt)
       pos = Len(WrkStr)
       iResp = InStr(1, WrkStr, " ")
       Jnt = Mid(WrkStr, 1, iResp - 1) & "," & Mid(WrkStr, iResp + 1, pos - iResp)
    End If
    
Joint_Time:
    iResp = RunSP(SpCon, "usp_ImportJointHolderData", 0, iAcnt, Jnt)
    
Commit_Check:
    iRecs = iRecs + 1
    ProgressBar1.Value = iRecs

    Set nextCell = curCell.Offset(1, 0)
    Set curCell = nextCell
  Loop
  sErrMsg = "Procedure failed in closing EXCEL spread sheet"
 .Workbooks.Close
 AppExcl.Quit
End With
Exit_Sub:
Exit Sub
End Sub

Private Sub ImportDivRef()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
Dim PayType As String, Mth As String
Dim sLen As Integer, pos As Long
Dim ChqDate As Date
Dim RecDate As Date
Dim WrkStr As String
Dim DiRefer As String

sErrMsgL1 = "Update failed while clearing balances of "
sErrMsgL2 = " Note this error."
'--
sMsg = "WARNING: This procedure will will delete existing Records"
sMsg = sMsg & "  then recreate them from the XL Import file. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Building Register"
'--
'--
'''On Error GoTo ImpJCSD_Fail
frmMDI.CmnDialog.DialogTitle = "Import Register Data XL File"
frmMDI.CmnDialog.Filter = "XLS(*.xls)|*.xls"
frmMDI.CmnDialog.DefaultExt = "XLS"
frmMDI.CmnDialog.ShowOpen
If Len(frmMDI.CmnDialog.filename) > 0 Then
     txtfile = frmMDI.CmnDialog.filename
Else
    GoTo Exit_Sub
End If
sErrMsg = "Procedure failed when trying to activate EXCEL"
Set AppExcl = CreateObject("Excel.application")
lbl.Caption = "Creating Dividend Ref for"

Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Creating Records..."
frmMDI.txtStatusMsg.Refresh
'--
sErrMsg = "Procedure failed/check if the client changed the format of the XL Sheet"
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
    sErrMsg = "Procedure failed while formatting/check if the client changed the format of the XL Sheet"
    Set nextCol = curCell.Offset(0, 2) '- Reference
    DiRefer = nextCol.Value
    
    Set nextCol = curCell.Offset(0, 3) '- Cheque Date
    WrkStr = nextCol.Value
    If WrkStr = "0" Then
       ChqDate = ""
    Else
    If Len(WrkStr) < 6 Then
       WrkStr = "0" & WrkStr
    End If
    pos = CInt(Mid(WrkStr, 3, 2))
    
    Select Case pos
           Case 1
                Mth = "Jan"
           Case 2
                Mth = "Feb"
           Case 3
                Mth = "Mar"
           Case 4
                Mth = "Apr"
           Case 5
                Mth = "May"
           Case 6
                Mth = "Jun"
           Case 7
                Mth = "Jul"
           Case 8
                Mth = "Aug"
           Case 9
                Mth = "Sep"
           Case 10
                Mth = "Oct"
           Case 11
                Mth = "Nov"
           Case 12
                Mth = "Dec"
     End Select
       
    sErrMsg = Mid(WrkStr, 1, 2) & "-" & Mth & "-" & Mid(WrkStr, 5, 2)
    ChqDate = CDate(sErrMsg)
    
    End If
    
    Set nextCol = curCell.Offset(0, 5) '- Description
    WrkStr = nextCol.Value
    
    If InStr(1, WrkStr, "CAP") > 0 Then
       PayType = "C"
    Else
       If InStr(1, WrkStr, "DIV") > 0 Then
         PayType = "D"
       Else
          GoTo Commit_Check
       End If
    End If
    Set nextCol = curCell.Offset(0, 6) '- Record Date
    WrkStr = nextCol.Value
    If WrkStr = "0" Then
       RecDate = ""
    Else
    If Len(WrkStr) < 6 Then
       WrkStr = "0" & WrkStr
    End If
    pos = CInt(Mid(WrkStr, 3, 2))
    
    Select Case pos
           Case 1
                Mth = "Jan"
           Case 2
                Mth = "Feb"
           Case 3
                Mth = "Mar"
           Case 4
                Mth = "Apr"
           Case 5
                Mth = "May"
           Case 6
                Mth = "Jun"
           Case 7
                Mth = "Jul"
           Case 8
                Mth = "Aug"
           Case 9
                Mth = "Sep"
           Case 10
                Mth = "Oct"
           Case 11
                Mth = "Nov"
           Case 12
                Mth = "Dec"
     End Select
       
    sErrMsg = Mid(WrkStr, 1, 2) & "-" & Mth & "-" & Mid(WrkStr, 5, 2)
    RecDate = CDate(sErrMsg)
    End If
    If IsNull(ChqDate) Then
       ChqDate = RecDate
    End If
    If IsNull(RecDate) Then
       RecDate = ChqDate
    End If
    'If ChqDate > "1/1/2006" Then
    '   MsgBox "hey"
    'End If
Cert_Time:
    iResp = RunSP(SpCon, "usp_ImportDivRefData", 0, Format(ChqDate, "dd-mmm-yyyy"), Format(RecDate, "dd-mmm-yyyy"), PayType, DiRefer)
    
Commit_Check:
    iRecs = iRecs + 1
    ProgressBar1.Value = iRecs

    Set nextCell = curCell.Offset(1, 0)
    Set curCell = nextCell
  Loop
  sErrMsg = "Procedure failed in closing EXCEL spread sheet"
 .Workbooks.Close
 AppExcl.Quit
End With
Exit_Sub:
Exit Sub

End Sub

Private Sub ImportReconData()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
Dim PayType As String, Mth As String
Dim sLen As Integer, pos As Long, iAcnt As Double
Dim ChqDate As Date
Dim DecDate As Date
Dim WrkStr As String
Dim DiRefer As String
Dim ChqNum As Integer
Dim ReconInd As Integer
Dim GrossPay As Currency
Dim Tax As Currency
Dim ChqAmt As Currency
Dim FolioMth As String

sErrMsgL1 = "Update failed while clearing balances of "
sErrMsgL2 = " Note this error."
'--
sMsg = "WARNING: This procedure will will delete existing Records"
sMsg = sMsg & "  then recreate them from the XL Import file. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Building Register"
'--
'--
'''On Error GoTo ImpJCSD_Fail
frmMDI.CmnDialog.DialogTitle = "Import Register Data XL File"
frmMDI.CmnDialog.Filter = "XLS(*.xls)|*.xls"
frmMDI.CmnDialog.DefaultExt = "XLS"
frmMDI.CmnDialog.ShowOpen
If Len(frmMDI.CmnDialog.filename) > 0 Then
     txtfile = frmMDI.CmnDialog.filename
Else
    GoTo Exit_Sub
End If
sErrMsg = "Procedure failed when trying to activate EXCEL"
Set AppExcl = CreateObject("Excel.application")
lbl.Caption = "Creating Dividend Ref for"

Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Creating Records..."
frmMDI.txtStatusMsg.Refresh
'--
sErrMsg = "Procedure failed/check if the client changed the format of the XL Sheet"
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
    sErrMsg = "Procedure failed while formatting/check if the client changed the format of the XL Sheet"
    Set nextCol = curCell.Offset(0, 2) '- Reference
    DiRefer = nextCol.Value
    
    Set nextCol = curCell.Offset(0, 3) '- ClientId
    WrkStr = nextCol.Value
    If Len(WrkStr) > 9 Then
       iAcnt = CDbl(Mid(WrkStr, 1, 9))
    Else
       iAcnt = CDbl(WrkStr)
    End If
    
    'iAcnt = nextCol.Value
    lbl.Caption = "Recreating Ledger for " & iAcnt
    lbl.Refresh
    
    Set nextCol = curCell.Offset(0, 4) '- Cheque Date
    ChqNum = nextCol.Value
    
    Set nextCol = curCell.Offset(0, 5) '- Cheque Date
    WrkStr = nextCol.Value
    If WrkStr = "0" Then
       ChqDate = ""
    Else
    If Len(WrkStr) < 6 Then
       WrkStr = "0" & WrkStr
    End If
    pos = CInt(Mid(WrkStr, 3, 2))
    
    Select Case pos
           Case 1
                Mth = "Jan"
           Case 2
                Mth = "Feb"
           Case 3
                Mth = "Mar"
           Case 4
                Mth = "Apr"
           Case 5
                Mth = "May"
           Case 6
                Mth = "Jun"
           Case 7
                Mth = "Jul"
           Case 8
                Mth = "Aug"
           Case 9
                Mth = "Sep"
           Case 10
                Mth = "Oct"
           Case 11
                Mth = "Nov"
           Case 12
                Mth = "Dec"
     End Select
       
    sErrMsg = Mid(WrkStr, 1, 2) & "-" & Mth & "-" & Mid(WrkStr, 5, 2)
    ChqDate = CDate(sErrMsg)
    
    End If
    
    Set nextCol = curCell.Offset(0, 7) '- Reconciliation Indicator
    WrkStr = nextCol.Value
    If WrkStr = "N" Then
       ReconInd = 0
    Else
       ReconInd = 1
    End If
    
    Set nextCol = curCell.Offset(0, 8) '- Declaration Date
    WrkStr = nextCol.Value
    If WrkStr = "0" Then
       DecDate = ""
    Else
    If Len(WrkStr) < 6 Then
       WrkStr = "0" & WrkStr
    End If
    pos = CInt(Mid(WrkStr, 3, 2))
    
    Select Case pos
           Case 1
                Mth = "Jan"
           Case 2
                Mth = "Feb"
           Case 3
                Mth = "Mar"
           Case 4
                Mth = "Apr"
           Case 5
                Mth = "May"
           Case 6
                Mth = "Jun"
           Case 7
                Mth = "Jul"
           Case 8
                Mth = "Aug"
           Case 9
                Mth = "Sep"
           Case 10
                Mth = "Oct"
           Case 11
                Mth = "Nov"
           Case 12
                Mth = "Dec"
     End Select
       
    sErrMsg = Mid(WrkStr, 1, 2) & "-" & Mth & "-" & Mid(WrkStr, 5, 2)
    DecDate = CDate(sErrMsg)
    End If
    If IsNull(ChqDate) Then
       ChqDate = DecDate
    End If
    If IsNull(DecDate) Then
       DecDate = ChqDate
    End If
    
    Set nextCol = curCell.Offset(0, 9) '- Gross Payment
    WrkStr = nextCol.Value
    GrossPay = CCur(WrkStr)

    Set nextCol = curCell.Offset(0, 10) '- Net Payment
    WrkStr = nextCol.Value
    Tax = CCur(WrkStr)
    ChqAmt = GrossPay = Tax
    FolioMth = Format("yyyydd", Date)
Cert_Time:
    iResp = RunSP(SpCon, "usp_ImportPaymentsData", 0, iAcnt, ChqNum, Format(ChqDate, "dd-mmm-yyyy"), ChqAmt, Format(DecDate, "dd-mmm-yyyy"), ReconInd, DiRefer, FolioMth)
    
Commit_Check:
    iRecs = iRecs + 1
    ProgressBar1.Value = iRecs

    Set nextCell = curCell.Offset(1, 0)
    Set curCell = nextCell
  Loop
  sErrMsg = "Procedure failed in closing EXCEL spread sheet"
 .Workbooks.Close
 AppExcl.Quit
End With
Exit_Sub:
Exit Sub

End Sub
