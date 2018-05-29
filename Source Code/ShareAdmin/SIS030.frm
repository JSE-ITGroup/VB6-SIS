VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS030 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Broker to Broker Transfer"
   ClientHeight    =   3255
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "SIS030.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6840
   Begin VB.TextBox tbFld 
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   2
      ToolTipText     =   "Assign a unique form number. Duplicates will be rejected."
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5760
      TabIndex        =   5
      ToolTipText     =   "Returns to the previous menu."
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Saves the edit/new certification"
      Top             =   2880
      Width           =   975
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   3
      ToolTipText     =   "Enter the certification date."
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   1
      ToolTipText     =   "Enter the date of the new batch"
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Select from existing batches or enter a new batch number"
      Top             =   600
      Width           =   1815
      DataFieldList   =   "Column 0"
      AllowNull       =   0   'False
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2090
      Columns(0).Caption=   "Batch #"
      Columns(0).Name =   "Batch #"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   8
      Columns(1).Width=   2170
      Columns(1).Caption=   "Date"
      Columns(1).Name =   "Date"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd-mmm-yyyy"
      Columns(1).FieldLen=   11
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Label lbl 
      Caption         =   "Certified Shares"
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
      Index           =   0
      Left            =   1920
      TabIndex        =   19
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label lbl 
      Caption         =   "Broker's name displayed here"
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
      Index           =   4
      Left            =   1920
      TabIndex        =   18
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   6840
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lbl 
      Caption         =   "Broker's name displayed here"
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
      Index           =   8
      Left            =   1920
      TabIndex        =   17
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Shares to Transfer:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "To Broker:"
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
      Index           =   2
      Left            =   720
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Batch Date:"
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
      Index           =   11
      Left            =   3840
      TabIndex        =   14
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
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
      Index           =   3
      Left            =   4200
      TabIndex        =   13
      Top             =   1150
      Width           =   540
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Transfer"
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
      Index           =   7
      Left            =   3840
      TabIndex        =   12
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "From Broker:"
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
      Index           =   12
      Left            =   720
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Form No:"
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
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   480
      Y2              =   480
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
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
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
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Batch No:"
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
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   1020
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
      TabIndex        =   9
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X, iNew As Integer, iEOF As Integer
Dim rsClient As ADODB.Recordset
Dim rsBrkCrt As New ADODB.Recordset
Dim rsActiv As New ADODB.Recordset
Dim rsBat As ADODB.Recordset
Dim rsVerBat As ADODB.Recordset
Dim rsVerFrm As ADODB.Recordset
Dim errLoop As Error
Dim errs1 As Error
Dim iStocks As Double
Dim iBrokerId As Long, iToBrkId As Long, sForm As String
Function IsValid() As Integer
Dim iErr As Integer, qSQL As String
Dim sElable As String
sElable = "Broker To Broker Transfer"
IsValid = False
iErr = 0
'--
If dbc(0) = "" Then ' batch
    iErr = 132
    csvShowUsrErr iErr, sElable
    dbc(0).SetFocus
    GoTo Validate_Exit
End If
dbc(0) = UCase(dbc(0))
'--
If meb(0) = "" Then 'batch date
   iErr = 139
   csvShowUsrErr iErr, sElable
   meb(0).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(0)) Then
      iErr = 14
      csvShowUsrErr iErr, sElable
      meb(0).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If tbfld(1).Text = "" Then 'Invalid form
   iErr = 140
   csvShowUsrErr iErr, sElable
   tbfld(1).SetFocus
   GoTo Validate_Exit
End If
tbfld(1).Text = UCase(tbfld(1).Text)
'--
If gblOptions = 1 Then 'check for duplicate form
  
  qSQL = "SELECT FORM from STKACTIV where FORM = '"
  qSQL = qSQL & tbfld(1) & "'"
  rsVerFrm.Open qSQL, cnn, , , adCmdText
  If Not rsVerFrm.EOF Then
     iErr = 141
     csvShowUsrErr iErr, sElable
     tbfld(1).SetFocus
     rsVerFrm.Close
     GoTo Validate_Exit
  End If
  rsVerFrm.Close
  
End If
'--
If meb(1) = "" Then 'Transfer date
   iErr = 129
   csvShowUsrErr iErr, sElable
   meb(1).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(1)) Then
      iErr = 14
      csvShowUsrErr iErr, sElable
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 IsValid = True
Validate_Exit:
   
   Exit Function
End Function
Private Sub cmdCancel_Click()
rsBrkCrt.Close
'''set cnn = nothing
Set rsClient = Nothing
Set rsActiv = Nothing
Set rsBat = Nothing
Set rsVerBat = Nothing
Set rsVerFrm = Nothing
frmSIS028.Visible = True
cnn.Close
Unload Me
End Sub
Private Sub cmdUpdate_Click()
Dim imsg As Integer, X As Integer, iLines As Integer
Dim iss As Date, iShares As Double
Dim strncode As String, strnbatch As String, StrnDate As Date
Dim iFCert As Long, iClient As Long, iTCert As Long
Dim qSQL As String, qBat As String
On Error GoTo cmdUpdate_Err
If Not IsValid Then Exit Sub
iLines = 0
iss = DateValue(meb(1).Text)
strncode = "B"
strnbatch = dbc(0)
StrnDate = DateValue(meb(0).Text)
cnn.BeginTrans
'----------------------------
'reduce selling brookers pool
'----------------------------
iStocks = rsBrkCrt!shares
iFCert = GetBrokerCert(iBrokerId, cnn)
If iFCert = 0 Then Exit Sub
'X = UpdBrokerPool(3, iBrokerId, iStocks * -1, cnn) ' reduce shrheld
If X = 0 Then Exit Sub ' error occured
'X = UpdBrokerPool(2, iBrokerId, iStocks, cnn) 'increase shrsell
If X = 0 Then ' reverse the reduction of shares held
   'X = UpdBrokerPool(3, iBrokerId, iStocks, cnn) ' increase shrheld
   Exit Sub
End If
'--------------------------------------------
'-- update certmas & stkname of selling broker
'X = UpdateCert(iBrokerId, iFCert, iStocks * -1, cnn) ' reduce pool cert.
If X = 0 Then ' reverse previous updates
 '  X = UpdBrokerPool(3, iBrokerId, iStocks, cnn) ' increase shrheld
  ' X = UpdBrokerPool(2, iBrokerId, iStocks * -1, cnn) 'decrease shrsell
   Exit Sub
End If
'-------------------------------------
'-- create stkactiv for selling broker
'-------------------------------------
With rsActiv
   .AddNew
   !TrnBatch = dbc(0)
   !TRNDATE = StrnDate
   iLines = iLines + 1
   !stklineno = iLines
   !Form = tbfld(1)
   !ClientID = rsBrkCrt!BROKERID
   !certno = 0
   !FRCERT = iFCert
   !CANDATE = DateValue(meb(1).Text)
   !FRSHARES = Val(lbl(0))
   !TrnCode = strncode
   !Status = "O"
   !BROKERID = iBrokerId
   .Update
End With
'--------------------------------
'increase receiving broker's pool
'--------------------------------
iClient = rsBrkCrt!TOBROKERID
iTCert = GetBrokerCert(iClient, cnn)
If iTCert = 0 Then ' create broker certmas & pool
  iTCert = CreateCert(iClient, iss, iStocks, strncode, strnbatch, StrnDate, tbfld(1), cnn)
  If iTCert = 0 Then ' fail to add new cert hense reverse all
     'X = UpdBrokerPool(3, iBrokerId, iStocks, cnn) ' increase shrheld
     'X = UpdBrokerPool(2, iBrokerId, iStocks * -1, cnn) 'decrease shrsell
     'X = UpdateCert(iBrokerId, iFCert, iStocks, cnn) ' increase pool cert.
     cnn.RollbackTrans
     Exit Sub
  Else
     X = iTCert
  End If
Else
   'X = UpdateCert(iClient, iTCert, iStocks, cnn) 'increase to broker certs
End If
If X = 0 Then 'reverse all
  'X = UpdBrokerPool(3, iBrokerId, iStocks, cnn) ' increase shrheld
  'X = UpdBrokerPool(2, iBrokerId, iStocks * -1, cnn) 'decrease shrsell
  'X = UpdateCert(iBrokerId, iFCert, iStocks, cnn) ' increase pool cert.
  cnn.RollbackTrans
  Exit Sub
End If
'X = UpdBrokerPool(1, iClient, iStocks, cnn) 'increase shrbuy
If X = 0 Then ' reverse all
   ' X = UpdBrokerPool(3, iBrokerId, iStocks, cnn) ' increase shrheld
   ' X = UpdBrokerPool(2, iBrokerId, iStocks * -1, cnn) 'decrease shrsell
   ' X = UpdateCert(iBrokerId, iFCert, iStocks, cnn) ' increase pool cert.
   ' X = UpdateCert(iClient, iTCert, iStocks * -1, cnn) 'decrease to broker pool
    cnn.RollbackTrans
    Exit Sub
End If
'---------------------------------------
'-- create stkactiv register transaction
'---------------------------------------
With rsActiv
   .AddNew
   !TrnBatch = dbc(0)
   !TRNDATE = StrnDate
   iLines = iLines + 1
   !stklineno = iLines
   !Form = tbfld(1)
   !ClientID = iClient
   !certno = iTCert
   !IssDate = iss
   !shares = Val(lbl(0))
   !TrnCode = strncode
   !Status = "O"
   !BrokerBuy = 1
   !BROKERID = iBrokerId
   .Update
   .Close
End With
'-- update batch header if new batch
'-----------
If iNew Then
     With rsBat
         qBat = "SELECT BATCHNO FROM BATHDR WHERE BATCHNO = '"
         qBat = qBat & dbc(0) & "'"
         rsVerBat.Open qBat, cnn, , , adCmdText
         If rsVerBat.EOF Then .AddNew
           !BATCHNO = dbc(0)
           !BATDATE = DateValue(meb(0).Text)
           .Update
           rsVerBat.Close
      End With
End If
'--- CHANGE CERTIFICATION STATUS
rsBrkCrt!Status = "T" ' Transferred
rsBrkCrt.Update
cnn.CommitTrans
'--
cmdCancel_Click
'---
Done:
Exit Sub
'--
cmdUpdate_Err:
  MsgBox "SIS030/cmdUpdate"
  csvLogError "SIS030/cmdUpdate", Err.Number, Err.Description
  cmdCancel_Click
End Sub
Private Sub dbc_InitColumnProps(Index As Integer)
Dim sRowinfo As String
Select Case Index
Case 0 ' Load Open Batches
rsBat.Requery
With rsBat
  If Not .EOF And Not .BOF Then
     .MoveFirst
     dbc(0).RemoveAll
     Do While Not .EOF
       sRowinfo = !BATCHNO & Chr(9) & !BATDATE
       dbc(0).AddItem sRowinfo
       .MoveNext
     Loop
  End If
End With
Case Else
End Select
End Sub

Private Sub dbc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  Select Case Index
    Case 0
      If meb(0).Enabled = True Then
         meb(0).SetFocus
      Else
         tbfld(1).SetFocus
      End If
     Case Else
  End Select
Case Else
End Select
End Sub
Private Sub dbc_LostFocus(Index As Integer)
Dim qDMLQry As String, i As Integer
Dim sRowinfo As String, bm As Variant
Select Case Index
Case 0
  '-----------------------------------------
  '-- get batch date if existing batch keyed
  '-- if not set focus to get date
  '-----------------------------------------
  If dbc(0) = "" Then
     dbc(0).SetFocus
  Else
     iNew = True
     For i = 0 To dbc(0).Rows - 1
        bm = dbc(0).GetBookmark(i)
        If dbc(0).Columns(0).CellText(bm) = dbc(0) Then
           meb(0).Mask = ""
           meb(0).Text = dbc(0).Columns(1).CellText(bm)
           meb(0).Enabled = False
           If Not IsDate(meb(1)) Then meb(1) = meb(0)
           iNew = False
           Exit For
        End If
    Next
    If iNew Then
      meb(0).Enabled = True
      meb(0).SetFocus
    End If
 End If
Case Else
End Select
End Sub
Private Sub Form_Activate()
Dim i As Integer
 On Error GoTo Form_Activate_Err
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
'--
UpdateScreen

'--
Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
 If Err = -2147168242 Then ' no current transactions
   Resume 0
 Else
   MsgBox "SIS030/Activate"
   csvLogError "SIS030/Activate", Err.Number, Err.Description
   Exit Sub
 End If
End Sub

Private Sub Form_Load()
Dim iDay As Integer
Dim qSQL As String, qMain As String
Dim qSql1 As String
Dim X, Y As Integer
Dim strTmp As String
On Error GoTo FL_ERR
iEOF = False
'--
   csvCenterForm Me, gblMDIFORM
   Set rsClient = New ADODB.Recordset
   Set rsActiv = New ADODB.Recordset
   Set rsBat = New ADODB.Recordset
   Set rsVerBat = New ADODB.Recordset
   Set rsVerFrm = New ADODB.Recordset
   '''Set cnn = New ADODB.Connection
   '----------------
   '-- Unpack gblfilekey
   '--------------------
   X = InStr(1, gblFileKey, ";", 1)
   iBrokerId = Val(Mid(gblFileKey, 1, X - 1))
   Y = InStr(X + 1, gblFileKey, ";", 1)
   sForm = Mid(gblFileKey, X + 1, Y - X - 1)
   iToBrkId = Val(Mid(gblFileKey, Y + 1, (Len(gblFileKey) - Y)))
   lbl(8).Caption = frmSIS028.lbl(2).Caption
   '-----------------------
   '-- open tables --------
   '-----------------------
   cnn.Open
   qSQL = "SELECT a.*, b.SHRHELD from STKBKCRT as a" _
          & " inner join STKBRKPL as b on" _
          & " a.BROKERID  = b.BROKERID" _
          & " Where a.BROKERID = " & iBrokerId & " and" _
          & " FORMNO = '" & sForm & "' and" _
          & " TOBROKERID = " & iToBrkId
   rsBrkCrt.Open qSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
   rsBat.Open "BATHDR", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
   '''cnn.Open cnn
   qSql1 = "SELECT * from STKACTIV WHERE FORM = '" & sForm _
          & "' and TRNCODE = 'B' and STATUS = 'O' and " _
          & "(CLIENTID = " & iBrokerId & " or CLIENTID = " _
          & iToBrkId & ") ORDER by FORM, stklineno"
   rsActiv.Open qSql1, cnn, adOpenDynamic, adLockOptimistic, adCmdText
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS030/Load"
  csvLogError "SIS030/Load", Err.Number, Err.Description
  Unload Me
   
End Sub
Private Sub UpdateScreen()
Dim qSQL As String
Dim i As Integer, bm As Variant

With rsActiv
  If Not .EOF Then
     dbc(0) = !TrnBatch
     meb(0) = !TRNDATE
     tbfld(1) = sForm
     meb(1) = !CANDATE
     dbc(0).Enabled = False
     meb(0).Enabled = False
     tbfld(1).Enabled = False
     meb(1).Enabled = False
     Me.Caption = "Add Broker to Broker Transfer"
   
  End If
End With
With rsBrkCrt
  If Not .EOF Then
      qSQL = "SELECT CLINAME,  CLIENTID FROM STKNAME "
      qSQL = qSQL & " WHERE CLIENTID = " & !TOBROKERID
      rsClient.Open qSQL, cnn, , , adCmdText
      lbl(4).Caption = rsClient!CliName
      rsClient.Close
      lbl(0).Caption = !shares
      If gblOptions = 2 Then cmdUpdate.Enabled = False
  End If
End With
End Sub
Private Sub meb_GotFocus(Index As Integer)
Select Case Index
Case 0
  meb(Index).Mask = "##-???-####"
Case Else
End Select
End Sub

Private Sub meb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  Select Case Index
  Case 0
     tbfld(1).SetFocus
  Case 1
     cmdUpdate.SetFocus
   End Select
Case vbKeyUp
KeyCode = 0
Select Case Index
Case 0
   If dbc(0).Enabled = True Then dbc(0).SetFocus
Case 1
   tbfld(1).SetFocus
End Select
Case Else
End Select
End Sub
Private Sub meb_LostFocus(Index As Integer)
Select Case Index
Case 0
  If Not IsDate(meb(1)) Then meb(1) = meb(0)
Case Else
End Select
End Sub
Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
 KeyCode = 0
 If Index = 1 Then
    meb(1).SetFocus
 End If
Case vbKeyUp
 If Index = 1 Then
   If meb(0).Enabled = True Then meb(0).SetFocus
 End If
Case Else
End Select
End Sub
