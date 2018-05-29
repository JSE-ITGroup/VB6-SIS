Attribute VB_Name = "modMain"
Option Explicit

Sub Main()

SDILogin.Show 1
If (Isloaded("SDILogin")) Then
        frmMDI.Show
        frmMDI.Enabled = True
        Unload SDILogin
        
Else
        ' Login form was closed by System menu or Cancel button.
        gblLoginName = ""
        gblPassword = ""
    
End If
End Sub
Function updCmpShares(stocks As Double)
Dim rsCmp As New ADODB.Recordset
Dim iErr As Integer
Dim errLoop As Error
Dim errs1 As Error
On Error GoTo UpdCmpShares_Err
updCmpShares = False
If stocks = 0 Then
  updCmpShares = True
  GoTo UpdCmpShares_Exit
End If
'-- open company table -------
Set rsCmp = New ADODB.Recordset
rsCmp.Open "COMPANY", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
'--
With rsCmp
  If Not .EOF Then
     If stocks > !Totstocks Then
        iErr = 134
        csvShowUsrErr iErr, "Allocate Shares"
        GoTo UpdCmpShares_Exit
     End If
     !Totstocks = !Totstocks - stocks
     !issStocks = !issStocks + stocks
     .Update
  End If
End With
updCmpShares = True
UpdCmpShares_Exit:
  Exit Function
UpdCmpShares_Err:
  MsgBox "UpdCmpShares"
  GoTo UpdCmpShares_Exit
  
End Function
Function CancelCert(client As Long, certno As Long, trn, batch _
       , tdt As Date, formno, SpCon As ADODB.Connection)
'-----------------------------------------------------------------
'-- Cancels active certificate and sets reply to true if successful
'-----------------------------------------------------------------
Dim rsCert As ADODB.Recordset
Dim qSQL As String, iErr As Integer, sCert As String
Dim iShares As Double
CancelCert = False
On Error GoTo CancelCert_Trap
Set rsCert = New ADODB.Recordset
'--
qSQL = "Select * from CERTMST where CERTNO = " & certno
rsCert.Open qSQL, SpCon, adOpenDynamic, adLockOptimistic, adCmdText
With rsCert
  If .EOF Then
     iErr = 144
     sCert = "CERT No = " & certno
     csvShowUsrErr iErr, "CancelCert", sCert
     GoTo CancelCert_Exit
  End If
  !TrnCode = trn
  !TrnBatch = batch
  !TrnDate = tdt
  !formno = formno
  iShares = !shares
  !Status = "C"
  .Update
 End With
 iShares = iShares * -1  'negate shares to reduce stkname:shares
 If gblOptions = 1 Then 'This is a false test. Actual is commented out below
 'If UpdStocks(client, iShares, SpCon) Then
   CancelCert = True
 End If
CancelCert_Exit:
 rsCert.Close
 Set rsCert = Nothing
 Exit Function

CancelCert_err:
  csvShowUsrErr iErr, "CreateCert"
  GoTo CancelCert_Exit
CancelCert_Trap:
  MsgBox "CreateCert"
End Function
Function CreateCert(client As Long, iss, shares As Double, trn, batch _
         , tdte As Date, formno, SpCon As ADODB.Connection) As Long
'---------------------------------------
'-- creates a certificate record and returns
'-- a certno if successful
'---------------------------------------
Dim rsCert As ADODB.Recordset, X As Integer
Dim qSQL As String, iErr As Integer, sCert As String
CreateCert = 0
On Error GoTo CreateCert_Err
Set rsCert = New ADODB.Recordset
CreateCert = GetNextCert(SpCon)
'--
qSQL = "Select * from CERTMST where CERTNO = " & CreateCert
rsCert.Open qSQL, SpCon, adOpenDynamic, adLockOptimistic, adCmdText
With rsCert
  If Not .EOF Then
     iErr = 142
     sCert = "CERT No = " & CreateCert
     csvShowUsrErr iErr, "CreateCert", sCert
     CreateCert = 0
     GoTo CreateCert_Exit
  End If
  '--
  .AddNew
  !certno = CreateCert
  !ClientID = client
  !IssDate = iss
  !shares = shares
  !TrnCode = trn
  !TrnBatch = batch
  !TrnDate = tdte
  !formno = formno
  !Status = "A"
  .Update
 End With
 'X = UpdStocks(client, shares, SpCon) this should be uncommented
CreateCert_Exit:
 rsCert.Close
 Set rsCert = Nothing
  Exit Function
CreateCert_Err:
MsgBox "CreateCert"
  GoTo CreateCert_Exit
End Function

Public Function GetNextCert(SpCon As ADODB.Connection) As Long
Dim qSQL As String
Dim rsUnused As ADODB.Recordset
Dim rsCmp As ADODB.Recordset
On Error GoTo GetNextCert_Err
Set rsCmp = New ADODB.Recordset
Set rsUnused = New ADODB.Recordset
'--
qSQL = "SELECT * from UNUSEDNOS where SEQTYP = 'C' "
qSQL = qSQL & " order by UNUSED"
rsCmp.Open "COMPANY", SpCon, adOpenDynamic, adLockOptimistic, adCmdTable
rsUnused.Open qSQL, SpCon, adOpenDynamic, adLockOptimistic, adCmdText
GetNextCert = 0
'-----------------------
'-- GET NEXT CERT NUMBER
'-----------------------
If Not rsUnused.EOF Then
     With rsUnused
        .MoveFirst
        GetNextCert = !UNUSED
        .Delete
      End With
Else
    With rsCmp
       If Not .EOF Then
          GetNextCert = !nextcert
          !nextcert = !nextcert + 1
          .Update
       End If
     End With
End If
rsCmp.Close
rsUnused.Close
GetNextCert_Exit:
 Set rsCmp = Nothing
 Set rsUnused = Nothing
 Exit Function
GetNextCert_Err:
  'put logic here to trap file being used error
  MsgBox "GetNextCert"
  GoTo GetNextCert_Exit
End Function
Public Function CreateBrokerPool(broker As Long, cert As Long)
Dim qSQL As String
Dim rsBroker As ADODB.Recordset
On Error GoTo CreateBrokerPool_Err
CreateBrokerPool = False
Set rsBroker = New ADODB.Recordset
qSQL = "SELECT * from STKBRKPL where BROKERID = "
qSQL = qSQL & broker
rsBroker.Open qSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
With rsBroker
   If .EOF Then
      .AddNew
      !BROKERID = broker
      !certno = cert
      .Update
   End If
   .Close
End With
CreateBrokerPool = True
rsBroker.Close
CreateBrokerPool_Exit:
  Set rsBroker = Nothing
  Exit Function
CreateBrokerPool_Err:
 MsgBox "CreateBrokerPool"
 GoTo CreateBrokerPool_Exit
End Function

Public Function GetBrokerCert(broker As Long, SpCon As ADODB.Connection) As Long
Dim qSQL As String
Dim rsBroker As ADODB.Recordset
On Error GoTo GetBrokerCert_Err
GetBrokerCert = False
Set rsBroker = New ADODB.Recordset
qSQL = "SELECT CERTNO FROM STKBRKPL where BROKERID = "
qSQL = qSQL & broker
rsBroker.Open qSQL, SpCon, adOpenStatic
If Not rsBroker.EOF Then GetBrokerCert = rsBroker!certno
rsBroker.Close
Set rsBroker = Nothing

GetBrokerCert_Exit:
 Exit Function
GetBrokerCert_Err:
  MsgBox "GetBrokerCert"
 GoTo GetBrokerCert_Exit
End Function

Public Function ImpBankRecon(textfile)
On Error GoTo Data_Err

Dim fs, F, iRecs As Long
Dim sInRec As String
Dim sTranDate As String, sType As String
Dim sText As String, sChqNo As String, sChqAmnt As String
Dim sBatDte As String, sBankAcnt As String, sDesc As String
Dim rectype As Integer, X As Integer
Dim sDD As String, sMM As String, sYY As String
Dim SDate As String, EDate As String
Dim SAmt As Currency, EAmt As Currency
Dim sBalType As String, eBalType As String
Dim SpCon As ADODB.Connection
'--
ImpBankRecon = False
frmMDI.txtStatusMsg.SimpleText = "Importing Bank Recon Ascii file..."
frmMDI.txtStatusMsg.Refresh
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

'Read File
'--------------------------
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.opentextfile(textfile)
sInRec = F.readline
iRecs = 0
If F.atendofstream = True Then
GoTo Open_err
End If
If MsgBox("Are you sure you wish to import the current file?", vbYesNo) = vbYes Then
   GoTo ContinueImport
Else
   GoTo IMPBankRecon_Exit
End If

ContinueImport:
'Collect the date
sYY = Trim(Mid(sInRec, 32, 4))
sMM = FindMonth(Trim(Mid(sInRec, 37, 2)))
If Len(sMM) < 1 Then
   ImpBankRecon2 (textfile)
   GoTo IMPBankRecon_Exit
End If

sDD = Trim(Mid(sInRec, 40, 2))
SDate = sDD & "-" & sMM & "-" & sYY

'collect bank acccount
sBankAcnt = Trim(Left(sInRec, 17))
sType = Mid(sInRec, 18, 1)
If sType = "1" Then
   sBalType = "C"
Else
   sBalType = "D"
End If
sChqAmnt = Trim(Mid(sInRec, 19, 13))
  If IsNumeric(sChqAmnt) Then
        sChqAmnt = Format(sChqAmnt, "#.##")
  Else
        sChqAmnt = Format(0, "#.99")
  End If
SAmt = CCur(sChqAmnt)

sInRec = F.readline

Do Until F.atendofstream = True
  'collect chqnumber and batchdate
  rectype = Len(sInRec)
  If rectype > 43 Then
  sChqNo = Trim(Mid(sInRec, 89, 7))
  sDD = Trim(Left(sInRec, 2))
  sMM = FindMonth(Trim(Mid(sInRec, 4, 2)))
  sYY = Trim(Mid(sInRec, 7, 2))
  sTranDate = sDD & "-" & sMM & "-" & sYY
  sDesc = Trim(Mid(sInRec, 10, 79))
  'validate trandate
  If IsDate(sTranDate) = True Then
     sTranDate = Format(sTranDate, "dd-mmm-yyyy")
    Else
    sTranDate = ""
  End If
  'collect and validate chqamnt
  sChqAmnt = Trim(Mid(sInRec, 96, 13))
  If IsNumeric(sChqAmnt) Then
        sChqAmnt = Format(sChqAmnt, "#.##")
  Else
        sChqAmnt = ""
  End If
  
  'collect the tran type
  sType = Trim(Mid(sInRec, 9, 1))
  
  
If sChqNo = "" Then
   sChqNo = 0
End If
If sType = "2" Then
   sType = "D"
Else
   sType = "C"
End If
   X = RunSP(SpCon, "usp_InsertBankItem", 0, sBankAcnt, sChqNo, CCur(sChqAmnt), sTranDate, sType, sDesc)
End If
iRecs = iRecs + 1

sInRec = F.readline
Loop
sYY = Trim(Mid(sInRec, 32, 4))
sMM = FindMonth(Trim(Mid(sInRec, 37, 2)))
sDD = Trim(Mid(sInRec, 40, 2))
EDate = sDD & "-" & sMM & "-" & sYY

sType = Mid(sInRec, 18, 1)
If sType = "1" Then
      eBalType = "C"
Else
      eBalType = "D"
End If
sChqAmnt = Trim(Mid(sInRec, 19, 13))
If IsNumeric(sChqAmnt) Then
        sChqAmnt = Format(sChqAmnt, "#.##")
Else
        sChqAmnt = Format(0, "#.99")
End If
EAmt = CCur(sChqAmnt)
   
X = RunSP(SpCon, "usp_BalanceUpdate", 0, sBankAcnt, SDate, SAmt, sBalType, EDate, EAmt, eBalType)

F.Close
ImpBankRecon = True
MsgBox "Imporatation Complete", vbInformation
IMPBankRecon_Exit:
  SpCon.Close
  Screen.MousePointer = vbDefault
  frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   frmMDI.txtStatusMsg.Refresh
   
  Exit Function
Open_err:
  MsgBox "Input Text File " & textfile & " is blank; import aborting... "
  F.Close
  GoTo IMPBankRecon_Exit
Data_Err:
  MsgBox Err & " " & Err.Description, vbOKOnly, "ImpBankRecon"
  GoTo IMPBankRecon_Exit
  
IMPBankRecon_NotFound:
   MsgBox ("ImpBankRecon")
End Function
Public Function ImpBankReconXL(textfile)
Dim X As Integer, sTranDate As String, sType As String
Dim sChqNo As String, sChqAmnt As Currency
Dim sBatDte As String, sBankAcnt As String, sDesc As String
Dim AppExcl As Excel.Application
Dim curCell As Object, nextCol As Object
Dim nextCell As Object, txtfile As String
Dim SpCon As ADODB.Connection
'--
'On Error GoTo IMPBankReconXL_NotFound
ImpBankReconXL = False
frmMDI.txtStatusMsg.SimpleText = "Importing Bank Recon Excel file..."
frmMDI.txtStatusMsg.Refresh
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

If MsgBox("Importation will clear existing records: do you wish to continue?", vbYesNo) = vbYes Then
   X = RunSP(SpCon, "usp_DeleteRetChqDr", 0)
Else
   GoTo IMPBankReconXL_Exit:
End If

'Read File
'--------------------------
Set AppExcl = CreateObject("Excel.application")
With AppExcl
  .Workbooks.Open (textfile)
  Set curCell = .Worksheets(1).Range("A3")
  Do While Not IsEmpty(curCell)
    'sErrMsg = "Procedure failed while formatting/check if the client changed the format of the XL Sheet"
    Set nextCol = curCell.Offset(0, 0) '- Cheque Number
    sChqNo = nextCol.Value
    Set nextCol = curCell.Offset(0, 1) ' - Bank Account
    sBankAcnt = nextCol.Value
    Set nextCol = curCell.Offset(0, 2) '- Cheque Amount
    sChqAmnt = nextCol.Value
    Set nextCol = curCell.Offset(0, 3) '- Trans Date
    sTranDate = nextCol.Value
    Set nextCol = curCell.Offset(0, 5) '- Description
    sDesc = nextCol.Value
    sType = "D"
    sBatDte = Format(Date, "yyyy/mm/dd")

X = RunSP(SpCon, "usp_InsertRetChqDr", 0, sBankAcnt, sBatDte, sChqNo, CCur(sChqAmnt), sTranDate, sType, sDesc)

Set nextCell = curCell.Offset(1, 0)
Set curCell = nextCell
Loop
.Workbooks.Close
AppExcl.Quit
End With

ImpBankReconXL = True
MsgBox "Imporatation Complete", vbInformation
IMPBankReconXL_Exit:
  SpCon.Close
  Screen.MousePointer = vbDefault
  frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
  frmMDI.txtStatusMsg.Refresh
  Exit Function

Open_err:
  MsgBox "Input Excel File " & textfile & " is blank; import aborting... "
  GoTo IMPBankReconXL_Exit

Data_Err:
  MsgBox Err & " " & Err.Description, vbOKOnly, "ImpBankRecon"
  GoTo IMPBankReconXL_Exit
  
IMPBankReconXL_NotFound:
   MsgBox ("ImpBankReconXL")
End Function

Public Sub CreateMandateLetter()
'''Set cnn = New ADODB.Connection
Dim unPk As cUnPackName
cnnClose
cnn.Open
Set unPk = New cUnPackName
Dim iReply As Integer
Dim rsClients As New ADODB.Recordset
Dim sql As String, sNewfile
Dim fso, txtfile, iErr As Integer
Dim sOutRec As String, msg As String
sNewfile = App.Path & "\bankmndte.txt"
Set rsClients = New ADODB.Recordset
'--
sql = "SELECT CHQTRN.ClientId, PayeeName, MndAddr1, MndAddr2, MndAddr3, mndAccnt, "
sql = sql & " mndacntnme, CliName, GrossPymnt - WhldTax as Amount, CHQTRN.CHQNUM "
sql = sql & " FROM (CHQTRN INNER JOIN StkName ON CHQTRN.ClientId = StkName.ClientId) "
sql = sql & " INNER JOIN STKPYMNTS ON STKNAME.CLIENTID = STKPYMNTS.CLIENTID "
sql = sql & " WHERE mndAccnt <>'NONE' and mndaccnt is not null and chqtrn.chqnum <> 0 "
sql = sql & " Order by CHQTRN.ChqNum"
On Error GoTo Open_err
rsClients.Open sql, cnn, , , adCmdText
'create txt file for printing bank mandate letters
'-------------------------------------------------
With rsClients
   If .EOF Then ' SELECT RETURNED NO RECORDS
   '---
     iErr = 164
     csvShowUsrErr iErr, "CreateMandateLetters"
   Else
   ' process returned recordset
   '--
   iErr = 0
   frmMDI.txtStatusMsg.SimpleText = "Creating Merge File....."      'open file for output & write header record
   '--
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set txtfile = fso.CreateTextFile(sNewfile, True)
      sOutRec = ""
      sOutRec = sOutRec & "CLIENTID" & Chr(9)
      sOutRec = sOutRec & "CLINAME" & Chr(9)
      sOutRec = sOutRec & "PAYEENAME" & Chr(9)
      sOutRec = sOutRec & "ADDRESS1" & Chr(9)
      sOutRec = sOutRec & "ADDRESS2" & Chr(9)
      sOutRec = sOutRec & "ADDRESS3" & Chr(9)
      sOutRec = sOutRec & "Amount" & Chr(9)
      sOutRec = sOutRec & "Account" & Chr(9)
      sOutRec = sOutRec & "AcntName" & Chr(9)
      sOutRec = sOutRec & "CHQNUM"
      txtfile.writeline (sOutRec)
      ' BUILD DATA FILE FROM RECORDSET
      '--
      .MoveFirst
      Do While Not .EOF
        sOutRec = ""
        sOutRec = sOutRec & !ClientID & Chr(9)
        If unPk.Unpack(!CliName) = True Then
            sOutRec = sOutRec & unPk.FirstName & _
                   " " & unPk.LastName & Chr(9)
        Else
           sOutRec = sOutRec & !CliName & Chr(9)
        End If
        sOutRec = sOutRec & !PayeeName & Chr(9)
        sOutRec = sOutRec & !MndAddr1 & Chr(9)
        sOutRec = sOutRec & !MndAddr2 & Chr(9)
        sOutRec = sOutRec & !MndAddr3 & Chr(9)
        sOutRec = sOutRec & Format(!Amount, "$#,##0.00") & Chr(9)
        sOutRec = sOutRec & !mndaccnt & Chr(9)
        sOutRec = sOutRec & !MndAcntNme & Chr(9)
        sOutRec = sOutRec & !ChqNum
        txtfile.writeline (sOutRec)
        .MoveNext
      Loop
      txtfile.Close
   End If
   .Close
End With
'-------  USE MICROSOFT MAIL MERGE TO PRINT LETTERS ---------
'------------------------------------------------------------
msg = "Select Yes to Activate Predefined Microsoft Word Document or No to end"
iReply = MsgBox(msg, vbQuestion + vbYesNo, "CreateMandateLetters")
If iReply = vbYes Then
 Dim AppWord As Word.Application
 Dim sDoc As String, iWarn As String
 On Error GoTo Openword_error
 iWarn = 168
 csvShowUsrErr iWarn, "Create Mandate Letters"
 frmMDI.txtStatusMsg.SimpleText = "Running Word to Print Mandate Letters..."
 sDoc = App.Path & "\STKPYMNTLET.doc"
 
 Set AppWord = CreateObject("Word.application")
 Screen.MousePointer = vbDefault
 With AppWord
  .Documents.Open (sDoc)
  .ActiveDocument.MailMerge.OpenDataSource (sNewfile)
 End With
 With AppWord.ActiveDocument.MailMerge
  .Destination = wdSendToPrinter
  .MailAsAttachment = False
  .MailAddressFieldName = ""
  .MailSubject = ""
  .SuppressBlankLines = True
  With .DataSource
      .FirstRecord = wdDefaultFirstRecord
      .LastRecord = wdDefaultLastRecord
  End With

 .Execute
End With
AppWord.ActiveDocument.Close
AppWord.DisplayAlerts = False
AppWord.Quit

cnn.Close
Exit Sub
Openword_error:
  
  MsgBox "Word Open Error " & Err.Number & "-" & Err.Description, , "CreateMandateLetters"
 
 Exit Sub
End If
Exit_CreateMandateLetter:
Exit Sub
Open_err:
   MsgBox "CreateMandateLetter"
   Exit Sub
End Sub


Public Sub CreateRIOfferLetter(Opt)
Dim unPk As New cUnPackName
Dim rsRI As New ADODB.Recordset
Dim sql As String, sNewfile As String
Dim fso, txtfile, iErr As Integer
Dim sOutRec As String, msg As String, iReply As Integer
'--
Set rsRI = New ADODB.Recordset
If Opt = 1 Then
   sNewfile = App.Path & "\MMRIOffer.doc"
   sql = "SELECT a.Clientid, a.shares, a.Offer, a.Cost, " _
      & "CliName, CliAddr1, CliAddr2, CliAddr3, CliAddr4, CliAddr5, " _
      & "(Select c.JNTNAME1 from STKJOINT c where a.CLIENTID = c.CLIENTID and JNTENDDTE is null) as JNT1, " _
      & "(Select c.JNTNAME2 from STKJOINT c where b.CLIENTID = c.CLIENTID and JNTENDDTE is null) as JNT2, " _
      & "(Select c.JNTNAME3 from STKJOINT c where b.CLIENTID = c.CLIENTID and JNTENDDTE is null) as JNT3 " _
      & "From STKRIWRK a INNER Join STKNAME b ON " _
      & "a.ClientId = B.ClientId " _
      & "Where Ledger = 'M' " _
      & "order by CliName, a.ClientId "
 Else
   sNewfile = App.Path & "\MMRIOffer.doc"
   sql = "SELECT a.Clientid, a.shares, a.Offer, a.Cost, " _
      & "GR8NAM as CliName, GR8AD1 as CliAddr1, GR8AD2 as CliAddr2, GR8Ad3 as CliAddr3, ' ' as CliAddr4, ' ' as CliAddr5, " _
      & "null as JNT1, null as JNT2, null as JNT3 " _
      & "From STKRIWRK a INNER Join JCSDSUB b ON " _
      & "a.Clientid = b.GR8NIN " _
      & "Where Ledger = 'S' "
      '& "Order by GR8NAM, ClientId "
 End If
cnn.Open
 rsRI.Open sql, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
'create txt file for printing RI Offer
With rsRI
   .Requery
   If .EOF Then ' SELECT RETURNED NO RECORDS
   '---
     iErr = 164
     csvShowUsrErr iErr, "RI Offer Selection"
   Else
   ' process returned recordset
   '--
      iErr = 0
      frmMDI.txtStatusMsg.SimpleText = "Creating Merge File....."      'open file for output & write header record
      '--
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set txtfile = fso.CreateTextFile(sNewfile, True)
      sOutRec = ""
      sOutRec = sOutRec & "CLINAME" & Chr(9)
      sOutRec = sOutRec & "JOINT1" & Chr(9)
      sOutRec = sOutRec & "JOINT2" & Chr(9)
      sOutRec = sOutRec & "JOINT3" & Chr(9)
      sOutRec = sOutRec & "ADDRESS1" & Chr(9)
      sOutRec = sOutRec & "ADDRESS2" & Chr(9)
      sOutRec = sOutRec & "ADDRESS3" & Chr(9)
      sOutRec = sOutRec & "ADDRESS4" & Chr(9)
      sOutRec = sOutRec & "ADDRESS5" & Chr(9)
      sOutRec = sOutRec & "CLIENTID" & Chr(9)
      sOutRec = sOutRec & "SHARES" & Chr(9)
      sOutRec = sOutRec & "OFFER" & Chr(9)
      sOutRec = sOutRec & "COST"
      txtfile.writeline (sOutRec)
      ' BUILD DATA FILE FROM RECORDSET
      '--
      .MoveFirst
      Do While Not .EOF
        sOutRec = ""
        If unPk.Unpack(!CliName) = True Then
            sOutRec = sOutRec & unPk.FirstName & _
                   " " & unPk.LastName & Chr(9)
        Else
           sOutRec = sOutRec & RTrim(!CliName) & Chr(9)
        End If
        If Not IsNull(!Jnt1) Then
          If unPk.Unpack(!Jnt1) = True Then
             sOutRec = sOutRec & "& " & unPk.FirstName & _
               " " & unPk.LastName & Chr(9)
          Else
             sOutRec = sOutRec & "& " & RTrim(!Jnt1) & Chr(9)
          End If
        Else
            sOutRec = sOutRec & !Jnt1 & Chr(9)
        End If
          
        If Not IsNull(!Jnt2) Then
           If unPk.Unpack(!Jnt2) = True Then
             sOutRec = sOutRec & "& " & unPk.FirstName & _
               " " & unPk.LastName & Chr(9)
           Else
             sOutRec = sOutRec & "& " & RTrim(!Jnt2) & Chr(9)
           End If
        Else
            sOutRec = sOutRec & !Jnt2 & Chr(9)
        End If
        If Not IsNull(!Jnt3) Then
           If unPk.Unpack(!Jnt3) = True Then
             sOutRec = sOutRec & "& " & unPk.FirstName & _
               " " & unPk.LastName & Chr(9)
           Else
             sOutRec = sOutRec & "& " & RTrim(!Jnt3) & Chr(9)
           End If
        Else
           sOutRec = sOutRec & !Jnt3 & Chr(9)
        End If
        sOutRec = sOutRec & !CliAddr1 & Chr(9)
        sOutRec = sOutRec & !CliAddr2 & Chr(9)
        sOutRec = sOutRec & IIf(IsNothing(!CliAddr3), ",", !CliAddr3) & Chr(9)
        sOutRec = sOutRec & IIf(IsNothing(!CliAddr4), ",", !CliAddr4) & Chr(9)
        sOutRec = sOutRec & IIf(IsNothing(!CliAddr5), ",", !CliAddr5) & Chr(9)
        sOutRec = sOutRec & !ClientID & Chr(9)
        sOutRec = sOutRec & Format(!shares, "#,##0") & Chr(9)
        sOutRec = sOutRec & Format(!offer, "#,##0") & Chr(9)
        sOutRec = sOutRec & Format(!Cost, "$#,##0.00")
        txtfile.writeline (sOutRec)
        .MoveNext
      Loop
      txtfile.Close
   End If
   .Close
End With
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
cnnClose
End Sub

Public Sub PrintRIOLetter(Opt)
Dim AppWord As Word.Application
Dim sNewfile As String
Dim sDoc As String, iWarn As String
On Error GoTo Open_error
csvShowUsrErr iWarn, "RI Offer Letter"
frmMDI.txtStatusMsg.SimpleText = "Running Word to Print RI Offer Letter..."
If Opt = 1 Then
  sDoc = App.Path & "\RIOffer.doc"
  MsgBox "Load Offer letters in printer", , "RI Offer letter"
Else
  sDoc = App.Path & "\RIAppForm.doc"
  MsgBox "Load Application Forms in printer", , "RI Application Form"
End If
sNewfile = App.Path & "\MMRIOffer.doc"
Set AppWord = CreateObject("Word.application")
Screen.MousePointer = vbDefault
With AppWord
 .Documents.Open (sDoc)
 .ActiveDocument.MailMerge.OpenDataSource (sNewfile)
End With
With AppWord.ActiveDocument.MailMerge
  .Destination = wdSendToPrinter
  .MailAsAttachment = False
  .MailAddressFieldName = ""
  .MailSubject = ""
  .SuppressBlankLines = True
  With .DataSource
      .FirstRecord = wdDefaultFirstRecord
      .LastRecord = wdDefaultLastRecord
  End With

 .Execute
End With
AppWord.ActiveDocument.Close
AppWord.DisplayAlerts = False
AppWord.Quit

Exit Sub
Open_error:
  
  MsgBox "Word Open Error " & Err.Number & "-" & Err.Description, , "Labels Selection"
 

End Sub

Public Sub cnnClose()
If cnn.State = 1 Then
   cnn.Close
End If
End Sub
Public Sub cnnOpen()
If cnn.State = 0 Then
   cnn.Open
End If

End Sub
Public Sub cnn1Close()
If cnn1.State = 1 Then
   cnn1.Close
End If
End Sub
Public Sub cnn1Open()
If cnn1.State = 0 Then
   cnn1.Open
End If

End Sub

Public Sub CreateFinacleFile(CreationType As String)
On Error GoTo Open_err

Dim iReply As Integer
Dim rsClients As New ADODB.Recordset
Dim rsBnkLodge As New ADODB.Recordset
Dim rsFinacle As New ADODB.Recordset
Dim sql As String, sNewfile
Dim fso, txtfile, iErr As Integer
Dim sOutRec As String, msg As String
Dim LastNumber As Long, FileSeqNo As Long
Dim FinNo As String
Dim pos As Integer
Dim CrAmt As Currency
Dim WrkChqDat As Date
Dim iCurrencyCode As String
Dim iComment As String
Dim Acct As String
Dim Amt As String
Dim SumAmt As Double
Dim SpCon As ADODB.Connection

Set rsClients = New ADODB.Recordset
Set rsFinacle = New ADODB.Recordset
Set rsBnkLodge = New ADODB.Recordset

'--
Set SpCon = New ADODB.Connection
With SpCon
     .ConnectionString = gblFileName
     .CursorLocation = adUseClient
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

'--
If CreationType = "D" Then
   Set rsFinacle = RunSP(SpCon, "usp_FinacleFileQry", 1)
Else
   Set rsFinacle = RunSP(SpCon, "usp_FinacleFileQryR", 1)
End If

If rsFinacle.EOF Then
   LastNumber = 1
   FileSeqNo = 1
Else
   LastNumber = rsFinacle!AccountNo
   FileSeqNo = rsFinacle!FileSeqNo
End If
Set rsBnkLodge = rsFinacle.NextRecordset
Set rsClients = rsFinacle.NextRecordset
rsFinacle.Close
Set rsFinacle = Nothing


rsBnkLodge.MoveFirst
If rsBnkLodge.EOF Then
   MsgBox "No Finacle Payments were found"
   GoTo Exit_CreateFinancleFile
End If
CrAmt = 0
With rsBnkLodge
     Do While Not .EOF
     Acct = Right(!AccountNo, 4)
     'If IsNumeric(Acct) Then
     pos = InStr(1, CStr(!ChqAmt), ".")
     If pos = 0 Then
        Amt = Right(CStr(!ChqAmt), 4)
     End If
     
     If pos > 0 And pos < 5 Then
        Amt = Mid(CStr(!ChqAmt), 1, pos - 1)
     End If
     If pos > 4 Then
        Amt = Mid(CStr(!ChqAmt), pos - 4, 4)
     End If
     
     SumAmt = SumAmt + CLng(Acct) + CLng(Amt)
     CrAmt = CrAmt + !ChqAmt
     .MoveNext
     Loop
End With
rsBnkLodge.Close
Set rsBnkLodge = Nothing

Acct = Right(LastNumber, 4)
pos = InStr(1, CStr(CrAmt), ".")
If pos = 0 Then
   Amt = Right(CStr(CrAmt), 4)
End If
     
If pos > 0 And pos < 5 Then
   Amt = Mid(CStr(CrAmt), 1, pos - 1)
End If
If pos > 4 Then
   Amt = Mid(CStr(CrAmt), pos - 4, 4)
End If
SumAmt = SumAmt + CLng(Acct) + CLng(Amt)
pos = 0
CrAmt = 0


'Create txt file for upload into Financle Core
'-------------------------------------------------
   
With rsClients
   If .EOF Then ' SELECT RETURNED NO RECORDS
   '---
     iErr = 164
     csvShowUsrErr iErr, "CreateFinancleFile"
   Else
   ' process returned recordset
   '--
   iErr = 0
   frmMDI.txtStatusMsg.SimpleText = "Creating Financle Direct Lodgement File....."      'open file for output & write header record
   '--
   
   sNewfile = frmSIS018.CmnDialog.FileName
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set txtfile = fso.CreateTextFile(sNewfile, True)
      sOutRec = ""
      sOutRec = sOutRec & Format(Date, "ddmmyyyy")
      sOutRec = sOutRec & Format(SumAmt, "000000000000000")
      sOutRec = sOutRec & Format(FileSeqNo, "00000000")
      txtfile.writeline (sOutRec)
      CrAmt = 0
      ' BUILD DATA FILE FROM RECORDSET
      '--
      .MoveFirst
      Do While Not .EOF
        sOutRec = ""
        sOutRec = sOutRec & "CI"
        pos = Len(CStr(!AccountNo))
        pos = 16 - pos
        sOutRec = sOutRec & Left("                ", pos) & !AccountNo
     
        sOutRec = sOutRec & "C"
        pos = Len(!Comment)
        sOutRec = sOutRec & !Comment & Right("                                                  ", pos)
        iComment = !Comment
        'sOutRec = sOutRec & "Dividend Payment                                  "
        sOutRec = sOutRec & !ReferenceNo
        sOutRec = sOutRec & Format(!ChqAmt, "0000000000000.00")
        sOutRec = sOutRec & !CurrencyCode
        iCurrencyCode = !CurrencyCode
        sOutRec = sOutRec & "   "
        sOutRec = sOutRec & Format(!ChqDate, "ddmmyyyy")
        WrkChqDat = !ChqDate
        sOutRec = sOutRec & "                              "
        sOutRec = sOutRec & "N"
        sOutRec = sOutRec & "                              "
        sOutRec = sOutRec & "                              "
        sOutRec = sOutRec & "                              "
        txtfile.writeline (sOutRec)
        CrAmt = CrAmt + !ChqAmt
        .MoveNext
      Loop
      sOutRec = ""
      sOutRec = sOutRec & "CI"
      pos = Len(CStr(LastNumber))
      pos = 16 - pos
      sOutRec = sOutRec & Left("                ", pos) & LastNumber
      sOutRec = sOutRec & "D"
      'sOutRec = sOutRec & "Dividend Payment                                  "
      pos = Len(iComment)
      sOutRec = sOutRec & iComment & Right("                                                  ", pos)
      sOutRec = sOutRec & "DATED " & Format(WrkChqDat, "yyyy mmm dd") & "   "
      sOutRec = sOutRec & Format(CrAmt, "0000000000000.00")
      sOutRec = sOutRec & iCurrencyCode
      sOutRec = sOutRec & "   "
      sOutRec = sOutRec & Format(WrkChqDat, "ddmmyyyy")
      sOutRec = sOutRec & "                              "
      sOutRec = sOutRec & "Y"
      sOutRec = sOutRec & "                              "
      sOutRec = sOutRec & "                              "
      sOutRec = sOutRec & "                              "
      txtfile.writeline (sOutRec)
      
      txtfile.Close
   
   End If
   .Close
End With

SpCon.Close
frmMDI.txtStatusMsg.SimpleText = "Creating Financle Direct Lodgement File.....Completed"
Exit_CreateFinancleFile:
Exit Sub
Open_err:
   MsgBox Err & " " & Err.Description, vbOKOnly, "Error on CreateFinancleFile"
   Exit Sub
End Sub

Public Function RunSP(SpCon As ADODB.Connection, sProcName As String, SpType As Integer, _
ParamArray aPara() As Variant) As Variant

Dim spCmd As ADODB.Command
Dim spPara As ADODB.Parameter
Dim spDataType As ADODB.DataTypeEnum
Dim i As Integer
Dim X As Integer
Dim intLength As Integer

Set spCmd = New ADODB.Command

If SpCon.State = 1 Then
   With spCmd
        .CommandText = sProcName
        .CommandType = adCmdStoredProc
        .ActiveConnection = SpCon
   End With
End If

Set spPara = spCmd.CreateParameter("Return", adInteger, adParamReturnValue)
spCmd.Parameters.Append spPara

X = UBound(aPara())

If X < 0 Then
   GoTo TypeTest
End If

SetupParameters:
For i = 0 To X
spDataType = VarType(aPara(i))

'If IsNumeric(aPara(i)) Then
If spDataType > 1 And spDataType < 7 Then
   Set spPara = spCmd.CreateParameter(, spDataType, adParamInput)
Else
intLength = Len(aPara(i))
Set spPara = spCmd.CreateParameter(, spDataType, adParamInput, intLength)
End If
spCmd.Parameters.Append spPara
spPara.Value = aPara(i)
Next i


TypeTest:
If SpType = 0 Then
      spCmd.Execute , , adAsyncExecute
Else
   Set RunSP = spCmd.Execute(, , adAsyncExecute)
End If

Do While spCmd.State = adStateExecuting
   Screen.MousePointer = vbHourglass
   frmMDI.txtStatusMsg.SimpleText = "Processing, Please wait......."
Loop
Screen.MousePointer = vbDefault

If SpType = 0 Then
   RunSP = spCmd.Parameters("Return").Value
Else
   gblReply = spCmd.Parameters("return").Value
End If


Set spCmd = Nothing
Set spPara = Nothing

End Function
Public Function IsNullMove(a As Variant) As String
If IsNull(a) Or a = "" Or a = " " Then
   IsNullMove = " "
   Else
   IsNullMove = a
End If

End Function
Function ValidEmail(ByVal strCheck As String) As Boolean
'Created by Chad M. Kovac
'Tech Knowledgey, Inc.
'http://www.TechKnowledgeyInc.com

Dim bCK As Boolean
Dim strDomainType As String
Dim strDomainName As String
Const sInvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
Dim i As Integer

bCK = Not InStr(1, strCheck, Chr(34)) > 0 'Check to see if there is a double quote
If Not bCK Then GoTo ExitFunction

bCK = Not InStr(1, strCheck, "..") > 0 'Check to see if there are consecutive dots
If Not bCK Then GoTo ExitFunction

' Check for invalid characters.
If Len(strCheck) > Len(sInvalidChars) Then
    For i = 1 To Len(sInvalidChars)
        If InStr(strCheck, Mid(sInvalidChars, i, 1)) > 0 Then
            bCK = False
            GoTo ExitFunction
        End If
    Next
Else
    For i = 1 To Len(strCheck)
        If InStr(sInvalidChars, Mid(strCheck, i, 1)) > 0 Then
            bCK = False
            GoTo ExitFunction
        End If
    Next
End If

If InStr(1, strCheck, "@") > 1 Then 'Check for an @ symbol
    bCK = Len(Left(strCheck, InStr(1, strCheck, "@") - 1)) > 0
Else
    bCK = False
End If
If Not bCK Then GoTo ExitFunction

strCheck = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "@"))
bCK = Not InStr(1, strCheck, "@") > 0 'Check to see if there are too many @'s
If Not bCK Then GoTo ExitFunction

strDomainType = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "."))
bCK = Len(strDomainType) > 0 And InStr(1, strCheck, ".") < Len(strCheck)
If Not bCK Then GoTo ExitFunction

strCheck = Left(strCheck, Len(strCheck) - Len(strDomainType) - 1)
Do Until InStr(1, strCheck, ".") <= 1
    If Len(strCheck) >= InStr(1, strCheck, ".") Then
        strCheck = Left(strCheck, Len(strCheck) - (InStr(1, strCheck, ".") - 1))
    Else
        bCK = False
        GoTo ExitFunction
    End If
Loop
If strCheck = "." Or Len(strCheck) = 0 Then bCK = False

'If InStr(1, strCheck, "jncb") > 0 Then
'   bCK = True
'Else
'   bCK = False
'End If

ExitFunction:
ValidEmail = bCK
End Function
Function FindMonth(MthNo As String)
Select Case MthNo
       Case "01"
            FindMonth = "Jan"
       Case "02"
            FindMonth = "Feb"
       Case "03"
            FindMonth = "Mar"
       Case "04"
            FindMonth = "Apr"
       Case "05"
            FindMonth = "May"
       Case "06"
            FindMonth = "Jun"
       Case "07"
            FindMonth = "Jul"
       Case "08"
            FindMonth = "Aug"
       Case "09"
            FindMonth = "Sep"
       Case "10"
            FindMonth = "Oct"
       Case "11"
            FindMonth = "Nov"
       Case "12"
            FindMonth = "Dec"
End Select
            
End Function
Public Sub CreateACHFile(CreationType As String)
On Error GoTo Open_err
Dim fso, txtfile
Dim sOutRec As String
Dim pos As Integer
Dim SpCon As ADODB.Connection
Dim rsACHHeader As New ADODB.Recordset
Dim rsACHDetails As New ADODB.Recordset
Dim StrSql As String
Dim CrAmt As Currency
Dim CrCnt As Double
Dim iOriginatorBank As String
Dim sNewfile As String

'--
Set SpCon = New ADODB.Connection
With SpCon
     .ConnectionString = gblFileName
     .CursorLocation = adUseClient
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

'--
If CreationType = "D" Then
   Set rsACHHeader = RunSP(SpCon, "usp_ACHHeaderQry", 1)
   Set rsACHDetails = RunSP(SpCon, "usp_ACHDetails", 1)
Else
   Set rsACHHeader = RunSP(SpCon, "usp_ACHHeaderQryR", 1)
   Set rsACHDetails = RunSP(SpCon, "usp_ACHDetailsR", 1)
End If

frmMDI.txtStatusMsg.SimpleText = "Creating ACH Direct Lodgement File....."
If rsACHDetails.EOF Then
    MsgBox "No ACH Payments were found"
    rsACHHeader.Close
    rsACHDetails.Close
    Set rsACHHeader = Nothing
    Set rsACHDetails = Nothing
    SpCon.Close
    GoTo Exit_CreateACHFile
End If
'open file for output
sNewfile = frmSIS018.CmnDialog.FileName
Set fso = CreateObject("Scripting.FileSystemObject")
Set txtfile = fso.CreateTextFile(sNewfile, True)

' Write header record
sOutRec = ""
sOutRec = sOutRec & "H"
With rsACHHeader
     sOutRec = sOutRec & Pad(!AccountNo, !LenAccountFld, "Right", " ")
     sOutRec = sOutRec & !CompanyID & "PPD" & !OriginatorID & !EffectiveDate
     pos = Len(sOutRec) + !LenReserved
     sOutRec = Pad(sOutRec, pos, "Right", " ")
     txtfile.writeline (sOutRec)
     iOriginatorBank = !OriginatorBank
End With
rsACHHeader.Close
Set rsACHHeader = Nothing

' End of header writing section

' Write transaction record(s)
'Set rsACHDetails = RunSP(SpCon, "usp_ACHDetails", 1)
CrAmt = 0
CrCnt = 0

With rsACHDetails ' Transaction Record
     Do While Not .EOF
        sOutRec = ""
        sOutRec = sOutRec & "T" 'Identifier for transactions
        StrSql = !AccountNo ' Beneficiary Account No
        StrSql = Pad(StrSql, 17, "Right", " ")
        sOutRec = sOutRec & StrSql
        sOutRec = sOutRec & !AccountType 'Accoun type of beneficiary Account
        StrSql = Format(!PaymentAmt, "#,##0.00") 'Amount to be credited to beneficiary
        CrAmt = CrAmt + !PaymentAmt
        CrCnt = CrCnt + 1
        StrSql = Pad(StrSql, 18, "Right", " ")
        sOutRec = sOutRec & StrSql
        sOutRec = sOutRec & !BranchID & !BankID ' Beneficiary's Branch and bank ID where account is held
        StrSql = !ClientID ' Participant ID i.e. ClientIDin StkName
        StrSql = Pad(StrSql, 15, "Right", " ")
        sOutRec = sOutRec & StrSql
        StrSql = !AccountName ' Participant Name
        StrSql = Pad(StrSql, 22, "Right", " ")
        sOutRec = sOutRec & "CR0" & "00000" & iOriginatorBank ' transaction type (CR), Notification (0) and originator bank
        StrSql = !ChqNum ' TraceNo assigned by the payment process
        StrSql = Pad(StrSql, 7, "Left", "0")
        sOutRec = sOutRec & StrSql
        StrSql = ""
        StrSql = Pad(StrSql, 10, "Right", " ")
        sOutRec = sOutRec & StrSql
        StrSql = !PaymentNarrative 'Addenda
        StrSql = Pad(StrSql, 50, "Right", " ")
        sOutRec = sOutRec & StrSql
        txtfile.writeline (sOutRec)
        .MoveNext
    Loop
End With
rsACHDetails.Close
Set rsACHDetails = Nothing
' End of Transaction section

' Start of Control section
sOutRec = ""
sOutRec = sOutRec & "C"
StrSql = ""
StrSql = Pad(StrSql, 10, "Right", " ")
sOutRec = sOutRec & StrSql
StrSql = "0.00"
StrSql = Pad(StrSql, 18, "Left", " ")
sOutRec = sOutRec & StrSql
StrSql = Format(CrAmt, "#,##0.00")
StrSql = Pad(StrSql, 18, "Left", " ")
sOutRec = sOutRec & StrSql
StrSql = "0"
StrSql = Pad(StrSql, 6, "Left", "0")
sOutRec = sOutRec & StrSql
StrSql = CrCnt
StrSql = Pad(StrSql, 6, "Left", "0")
sOutRec = sOutRec & StrSql
StrSql = CrCnt
StrSql = Pad(StrSql, 7, "Left", "0")
sOutRec = sOutRec & StrSql
StrSql = ""
StrSql = Pad(StrSql, 94, "Right", " ")
sOutRec = sOutRec & StrSql
txtfile.writeline (sOutRec)
' End of Control Section

txtfile.Close
SpCon.Close
frmMDI.txtStatusMsg.SimpleText = "Creating ACH File.....Completed"
Exit_CreateACHFile:
Exit Sub
Open_err:
   MsgBox "Error on CreateACHFile"
   Exit Sub

End Sub
Public Function Pad(StringToPad As String, TotalLengthOfReturnString As Integer, LeftOrRight As String, Optional CharacterToPadWith As String = " ") As String
On Error GoTo PadError
Dim StringLength As Integer, TotalLength As Integer, Difference As Integer
StringLength = Len(StringToPad)
Difference = TotalLengthOfReturnString - StringLength
If Difference > 0 Then
   If LeftOrRight = "Left" Then
      Pad = String(Difference, CharacterToPadWith) & StringToPad
   Else
      Pad = StringToPad & String(Difference, CharacterToPadWith)
   End If
Else
   Pad = StringToPad
End If
Exit Function
PadError:
  MsgBox Err.Description
End Function
Public Function ConvertToNumber(InString As String)
Dim i As Integer
Dim NewNumber As Long
Dim Y As String
Dim m As Integer

NewNumber = 0
For i = 1 To Len(InString)
    Y = Mid(InString, i, 1)
    m = Asc(Y)
    NewNumber = NewNumber + m
Next i
ConvertToNumber = NewNumber
    
End Function
Public Function ImpBankRecon2(textfile)
On Error GoTo Data_Err

Dim fs, F, iRecs As Long
Dim sInRec As String
Dim sTranDate As String, sType As String
Dim sText As String, sChqNo As String, sChqAmnt As String
Dim sBatDte As String, sBankAcnt As String, sDesc As String
Dim rectype As Integer, X As Integer
Dim sDD As String, sMM As String, sYY As String
Dim SDate As String, EDate As String
Dim SAmt As Currency, EAmt As Currency
Dim sBalType As String, eBalType As String, eChqAmnt As String
Dim SpCon As ADODB.Connection
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

'Read File
'--------------------------
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.opentextfile(textfile)
ContinueImport:
X = 0
Do While X < 5
   sInRec = F.readline
   X = X + 1
Loop
sBankAcnt = Mid(sInRec, 67, 9)
'Collect the date
sInRec = F.readline
sInRec = F.readline
sYY = Trim(Mid(sInRec, 74, 4))
X = 0
Do While X < 16
   sInRec = F.readline
   X = X + 1
Loop
sMM = Trim(Mid(sInRec, 4, 3))
sDD = Left(sInRec, 2)
SDate = sDD & "-" & sMM & "-" & sYY

'collect bank acccount

sChqAmnt = Trim(Mid(sInRec, 62, 18))
If IsNumeric(sChqAmnt) Then
   sChqAmnt = Format(sChqAmnt, "#.##")
Else
   sChqAmnt = Format(0, "#.99")
End If
SAmt = CCur(sChqAmnt)
If SAmt < 0 Then
   sBalType = "D"
   sType = "2"
Else
   sBalType = "C"
   sType = "1"
End If

sInRec = F.readline
sInRec = F.readline

Do Until F.atendofstream = True
  'collect chqnumber and batchdate
  rectype = Len(Trim(Left(sInRec, 6)))
  If rectype <> 6 Then
     sInRec = F.readline
     GoTo CheckAgain
  End If
  rectype = Len(Trim(Mid(sInRec, 9, 3)))
  If rectype <> 3 Then
     sInRec = F.readline
     GoTo CheckAgain
  End If
  
  sMM = Trim(Mid(sInRec, 4, 3))
  sDD = Left(sInRec, 2)
  sTranDate = sDD & "-" & sMM & "-" & sYY
  If IsDate(sTranDate) = True Then
     sTranDate = Format(sTranDate, "dd-mmm-yyyy")
  Else
     sTranDate = ""
  End If
  
  sChqNo = Trim(Mid(sInRec, 89, 7))
  sDesc = Trim(Mid(sInRec, 9, 37))
  X = InstrNum(1, sDesc)
  If X = 0 Then
     sChqNo = 0
     GoTo FindAmount
  End If
  sChqNo = Mid(sDesc, X, 7)
  If IsNumeric(sChqNo) Then
     sChqNo = sChqNo
  Else
     sChqNo = 0
  End If
  
  'validate trandate
  'collect and validate chqamnt
FindAmount:
  sChqAmnt = Trim(Mid(sInRec, 46, 18))
  If IsNumeric(sChqAmnt) Then
     sChqAmnt = Format(sChqAmnt, "#.##")
  Else
     sChqAmnt = "0"
  End If
  If sChqAmnt < 0 Then
     sBalType = "D"
     sType = "2"
     sChqAmnt = sChqAmnt * -1
  Else
     sBalType = "C"
     sType = "1"
  End If
  eChqAmnt = Trim(Mid(sInRec, 63, 17))
  If IsNumeric(eChqAmnt) Then
     eChqAmnt = Format(eChqAmnt, "#.##")
  Else
     eChqAmnt = "0"
  End If
  
  X = RunSP(SpCon, "usp_InsertBankItem", 0, sBankAcnt, sChqNo, CCur(sChqAmnt), sTranDate, sType, sDesc)
  iRecs = iRecs + 1
CheckAgain:
  If iRecs = 28 Then
     iRecs = iRecs
  End If
  sInRec = F.readline
Loop

EDate = sDD & "-" & sMM & "-" & sYY

If eChqAmnt < 0 Then
     eBalType = "D"
  Else
     eBalType = "C"
  End If

EAmt = CCur(eChqAmnt)
   
X = RunSP(SpCon, "usp_BalanceUpdate", 0, sBankAcnt, SDate, SAmt, sBalType, EDate, EAmt, eBalType)

F.Close
ImpBankRecon2 = True
MsgBox "Importation Complete", vbInformation
IMPBankRecon2_Exit:
  SpCon.Close
  Screen.MousePointer = vbDefault
  frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   frmMDI.txtStatusMsg.Refresh
   
  Exit Function
Open_err:
  MsgBox "Input Text File " & textfile & " is blank; import aborting... "
  F.Close
  GoTo IMPBankRecon2_Exit
Data_Err:
  MsgBox Err & " " & Err.Description, vbOKOnly, "ImpBankRecon"
  GoTo IMPBankRecon2_Exit
  
IMPBankRecon2_NotFound:
   MsgBox ("ImpBankRecon2")
End Function

Public Function InstrNum(Optional Start = 1, Optional strTest As String) As Long
Dim i As Long
InstrNum = 0
For i = Start To Len(strTest)
 If Mid$(strTest, i, 1) Like "#" Then
 InstrNum = i
 Exit For
 End If
Next
End Function
 
