VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS023 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Issue"
   ClientHeight    =   4470
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "SIS023.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6810
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   300
      Left            =   1440
      TabIndex        =   23
      ToolTipText     =   "Pressing this button will activate the search program to locate a shareholder."
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox tbFld 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Assign a document number."
      Top             =   1080
      Width           =   1500
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   3600
      TabIndex        =   20
      Top             =   4080
      Width           =   975
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   0
      Left            =   1560
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
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   2520
      TabIndex        =   13
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5760
      TabIndex        =   7
      ToolTipText     =   "Cancels all processing and exits program."
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4680
      TabIndex        =   6
      ToolTipText     =   "Saves the screen information to the database."
      Top             =   4080
      Width           =   975
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   4920
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
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      ToolTipText     =   "Enter the number of units purchased."
      Top             =   3360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "Enter the date of  issue you want to appear on the certificate.."
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "Select a share holder from the list."
      Top             =   1800
      Width           =   2895
      DataFieldList   =   "Column 1"
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
      Columns(0).Width=   5636
      Columns(0).Caption=   "Client Name"
      Columns(0).Name =   "Client Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   50
      Columns(1).Width=   3200
      Columns(1).Caption=   "Client Id"
      Columns(1).Name =   "Client Id"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   10
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      Enabled         =   0   'False
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Holder #3:"
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
      TabIndex        =   25
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Joint Holder #3 Name:"
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
      Left            =   1800
      TabIndex        =   24
      Top             =   2880
      Width           =   3015
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
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label lblLabels 
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
      Index           =   7
      Left            =   3720
      TabIndex        =   21
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Joint Holder #1 Name:"
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
      Left            =   1800
      TabIndex        =   19
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label lblLabels 
      Caption         =   "Joint Holder #2 Name:"
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
      Index           =   13
      Left            =   1800
      TabIndex        =   18
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Holder #2:"
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
      Left            =   0
      TabIndex        =   17
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblLabels 
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
      Index           =   14
      Left            =   480
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Holder #1:"
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
      Left            =   0
      TabIndex        =   15
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "No of Units:"
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
      Left            =   0
      TabIndex        =   14
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   3960
      Y2              =   3960
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
      TabIndex        =   11
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Lead Share Holder:"
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
      Index           =   16
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   1740
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
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date of Issue:"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
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
      TabIndex        =   12
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS023"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim X As Integer, iEOF As Integer
Dim iNewApp As Integer, iNew As Integer
Dim rsCmp As ADODB.Recordset
Dim rsClient As ADODB.Recordset
Dim rsMain As ADODB.Recordset
Dim rsJoint As ADODB.Recordset
Dim rsUnused As ADODB.Recordset
Dim rsBat As ADODB.Recordset
Dim rsVerBat As ADODB.Recordset
Dim rsVerFrm As ADODB.Recordset
Dim cmdChange As ADODB.Command
Dim cmdDel As ADODB.Command
Dim errLoop As Error
Dim errs1 As Error
Dim strTable As String
Dim strRecNO As String
Dim iStocks As Double, icert As Long
Dim bm As Variant
Function IsValid() As Integer
Dim iErr As Integer, dtefld As Date
Dim sElable, qSQL As String
IsValid = False
sElable = "Stock Issue Entry"
iErr = 0
'--
If dbc(0) = "" Then ' batch
    iErr = 132
    csvShowUsrErr iErr, sElable
    tbFld(0).SetFocus
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
 If tbFld(1).Text = "" Then 'Invalid form
   iErr = 140
   csvShowUsrErr iErr, sElable
   tbFld(1).SetFocus
   GoTo Validate_Exit
End If
tbFld(1).Text = UCase(tbFld(1).Text)
'--
If gblOptions = 1 Then 'check for duplicate form
  
  qSQL = "SELECT FORM from STKACTIV where TRNBATCH = '"
  qSQL = qSQL & dbc(0) & "' and "
  qSQL = qSQL & "FORM = '" & tbFld(1) & "'"
  rsVerFrm.Open qSQL, cnn, , , adCmdText
  If Not rsVerFrm.EOF Then
     iErr = 141
     csvShowUsrErr iErr, sElable
     tbFld(1).SetFocus
     rsVerFrm.Close
     GoTo Validate_Exit
  End If
  rsVerFrm.Close
End If
'--
If dbc(1) = "" Then   ' Client name
   iErr = 130
   csvShowUsrErr iErr, sElable
   dbc(1).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If meb(2) = "" Then 'issue date
   iErr = 129
   csvShowUsrErr iErr, sElable
   meb(2).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(2)) Then
      iErr = 14
      csvShowUsrErr iErr, sElable
      meb(2).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If meb(1) = "" Then ' shares
   iErr = 131
   csvShowUsrErr iErr, sElable
   meb(1).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsNumeric(meb(1)) Then
      iErr = 28
      csvShowUsrErr iErr, "Stock Issue  Entry"
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
Dim iSeqKey As Integer
rsCmp.Close
rsMain.Close
rsClient.Close
Set rsMain = Nothing
Set rsCmp = Nothing
Set rsClient = Nothing
Set rsJoint = Nothing
Set rsVerBat = Nothing
Set rsVerFrm = Nothing
cnn.Close
'''set cnn = nothing
Unload Me
End Sub

Private Sub cmdClear_Click()
ClearScreen
If gblOptions = 2 Then UpdateScreen
End Sub

Private Sub cmdDelete_Click()
Dim imsg As Integer, X As Integer
Dim qDMLQry As String
imsg = 133
If csvYesNo(imsg, "Stock Issue") Then
   cnn.BeginTrans
   iStocks = rsMain!shares * -1
   If updCmpShares(iStocks) Then
     qDMLQry = "DELETE FROM CERTMST WHERE CLIENTID = "
     qDMLQry = qDMLQry & rsMain!ClientID & " and CERTNO = "
     qDMLQry = qDMLQry & rsMain!certno
     X = csvADODML(qDMLQry, cnn)
     '--
     qDMLQry = "UPDATE STKNAME SET SHARES = SHARES - " & rsMain!shares
     qDMLQry = qDMLQry & " WHERE CLIENTID = " & rsMain!ClientID
     X = csvADODML(qDMLQry, cnn)
     '--
     rsMain.Delete
     cnn.CommitTrans
     cmdCancel_Click
     Set cmdChange = Nothing
     Set cmdDel = Nothing
   End If
End If
End Sub

Private Sub cmdFind_Click()
Dim i As Integer, X As Integer
Dim cChk As Integer, qCli As String
Dim sWhere As String

Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
     .cbWhere.Clear
    .cbWhere.AddItem "CliName"
    .cbWhere.AddItem "ClientId"
    .cbWhere.ListIndex = 0
    .cbOptions.ListIndex = 0
    .optBtn.Visible = False
    .lbl(3).Visible = False
    .Show vbModal
    '----------------------------
    '-------- main line ---------
    '----------------------------
    If .tbFind.Text = vbNullString Then
    Else
      rsClient.Close
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      sWhere = Trim(.tbFind.Text)
      X = .cbWhere.ListIndex
      qCli = "SELECT CLINAME,  CLIENTID FROM STKNAME WHERE "
      qCli = qCli & "CATCODE <> 'SB' and "
      qCli = qCli & .cbWhere
      '---
      If sWhere <> "" Then
          Select Case .cbOptions.ListIndex
          Case 0 ' Exact Match
              If X = 0 Then
                qCli = qCli & " = '" & sWhere & "'"
              Else
                qCli = qCli & " = " & Val(.tbFind.Text)
              End If
          Case 1 ' Starts With
              sWhere = Trim(.tbFind.Text) & "%"
              qCli = qCli & " like '" & sWhere & "'"
              qCli = qCli & " ORDER BY CLINAME, CLIENTID"
            Case 2 ' Ends With
               sWhere = "%" & Trim(.tbFind.Text)
               qCli = qCli & " like '" & sWhere & "'"
               qCli = qCli & " ORDER BY CLINAME, CLIENTID"
             Case 3 ' AnyWhere
               sWhere = "%" & Trim(.tbFind.Text) & "%"
               qCli = qCli & " like '" & sWhere & "'"
               qCli = qCli & " ORDER BY CLINAME, CLIENTID"
             End Select
      End If
      rsClient.Open qCli, cnn, , , adCmdText
      If Not rsClient.EOF Then
         dbc_InitColumnProps (1)
         dbc(1).Enabled = True
         dbc(1).SetFocus
         cmdUpdate.Enabled = True
      End If
    End If
  End With
Unload frmFind
Set frmFind = Nothing
Exit Sub
cmdFind_Click_err:
  MsgBox "SIS023/CmdFind"
End Sub
Private Sub cmdUpdate_Click()
Dim qBat As String
Dim iClient As Long, iss As Date, iShares As Double
Dim strncode As String, strnbatch As String, StrnDate As Date
On Error GoTo cmdUpdate_Err
' wait message
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.Refresh
If IsValid Then
  '--
  cnn.BeginTrans
  '--
  iClient = dbc(1).Columns(1).Text
  iss = DateValue(meb(2).Text)
  iShares = Val(meb(1).Text)
  strncode = "I"
  strnbatch = dbc(0)
  StrnDate = DateValue(meb(0).Text)
  iStocks = Val(meb(1).Text)
  '--
  If gblOptions = 2 Then
     iStocks = iStocks - rsMain!shares
     icert = rsMain!certno
  End If
  '--
  If updCmpShares(iStocks) Then
    With rsMain
    '---
      If gblOptions = 1 Then
        ' icert = CreateCert(iClient, iss, iStocks, strncode, strnbatch, StrnDate, tbFld(1), cnn)
         .AddNew
      Else
         If gblOptions = 1 Then
         'If UpdateCert(iClient, icert, iStocks, cnn) Then
         Else
            GoTo Done
         End If
      End If
      !TrnBatch = strnbatch
      !TRNDATE = StrnDate
      !stklineno = 1
      !Form = tbFld(1).Text
      !ClientID = iClient
      !IssDate = iss
      !shares = iShares
      !TrnCode = strncode ' Application Issue
      !Status = "O"
      !certno = icert
      .Update
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
    '--
    cnn.CommitTrans
    If gblOptions = 1 Then
       dbc_InitColumnProps (0)
       InitAddNew
    Else
       cmdCancel_Click
      End If
   End If
End If
'---
Done:
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
 Exit Sub
'--
cmdUpdate_Err:
  cnn.RollbackTrans
  MsgBox "SIS023/cmdUpdate"
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
Case 1 'Load Client Names from recordset
   rsClient.Requery
   With rsClient
     If Not .EOF And Not .BOF Then
       .MoveFirst
       dbc(1).RemoveAll
       Do While Not .EOF
          sRowinfo = !CliName & Chr(9) & !ClientID
          dbc(1).AddItem sRowinfo
          If dbc(1).Row = 0 Then dbc(1) = !CliName
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
  If Index = 1 Then
     meb(1).SetFocus
  End If
Case vbKeyUp
  KeyCode = 0
  If Index = 1 Then meb(2).SetFocus
End Select
End Sub

Private Sub dbc_LostFocus(Index As Integer)
Dim qDMLQry As String, i As Integer

Select Case Index
Case 0
  '-----------------------------------------
  '-- get batch date if existing batch keyed
  '-- if not set focus to get date
  '-----------------------------------------
  If dbc(0) <> "" Then
   iNew = True
   For i = 0 To dbc(0).Rows - 1
     bm = dbc(0).GetBookmark(i)
     If dbc(0).Columns(0).CellText(bm) = dbc(0) Then
       meb(0).Text = Format(dbc(0).Columns(1).CellText(bm), "dd-mmm-yyyy")
       meb(0).Enabled = False
       tbFld(1).SetFocus
       iNew = False
       Exit For
     End If
   Next
   If iNew Then
    meb(0).Enabled = True
    meb(0).SetFocus
   End If
  End If
Case 1
  '---------------------------------------
  '-- get corresponding joint record for
  '-- displaying
  '----------------------------------------
  qDMLQry = "SELECT JNTNAME1,JNTNAME2, JNTNAME3 FROM STKJOINT WHERE "
  qDMLQry = qDMLQry & " CLIENTID = " & dbc(1).Columns(1).Text
  qDMLQry = qDMLQry & " and JNTENDDTE  is NULL"
  rsJoint.Open qDMLQry, cnn, , , adCmdText
  If Not rsJoint.EOF Then
    lblLabels(12).Caption = rsJoint!JNTNAME1
    If Not IsNull(rsJoint!JNTNAME2) Then lblLabels(13).Caption = rsJoint!JNTNAME2
    If Not IsNull(rsJoint!jntname3) Then lblLabels(4).Caption = rsJoint!jntname3
  Else
    lblLabels(12).Caption = " "
    lblLabels(13).Caption = " "
    lblLabels(4).Caption = " "
  End If
  rsJoint.Close
Case Else
End Select
End Sub
Private Sub Form_Activate()
 On Error GoTo Form_Activate_Err
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
'--
If gblOptions = 2 Then
  UpdateScreen
   Me.Caption = "Edit Stock Issue"
   meb(0).Enabled = False   ' Batch Date
   dbc(0).Enabled = False   ' Batch No
End If
'--
Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
 If Err = -2147168242 Then ' no current transactions
   Resume 0
 Else
   MsgBox "SIS023/Activate"
   Exit Sub
 End If
End Sub

Private Sub Form_Load()
Dim iDay As Integer, ipos As Long
Dim qSQL As String, qMain As String
Dim qSql1 As String, iClient As Long, icert As Long
Dim i As Integer
Dim strTmp As String
On Error GoTo FL_ERR
iEOF = False
'--
   csvCenterForm Me, gblMDIFORM
   '''Set cnn = New ADODB.Connection
   cnn.Open
   Set rsCmp = New ADODB.Recordset
   Set rsClient = New ADODB.Recordset
   Set rsJoint = New ADODB.Recordset
   Set rsMain = New ADODB.Recordset
   Set rsUnused = New ADODB.Recordset
   Set rsBat = New ADODB.Recordset
   Set rsVerBat = New ADODB.Recordset
   Set rsVerFrm = New ADODB.Recordset
   '-----------------------
   '-- open tables --------
   '-----------------------
   If gblOptions = 2 Then
       ipos = InStr(1, gblFileKey, ";", 1)
       iClient = Val(Mid(gblFileKey, 1, ipos - 1))
       icert = Val(Mid(gblFileKey, ipos + 1, (Len(gblFileKey) - ipos)))
   Else
       iClient = 0
       icert = 0
   End If
   qMain = "SELECT * FROM STKACTIV WHERE CLIENTID = "
   qMain = qMain & iClient & " and CERTNO = " & icert & " and "
   qMain = qMain & " TRNCODE = 'I' and STATUS = 'O'"
   rsMain.Open qMain, cnn, adOpenDynamic, adLockPessimistic, adCmdText
   rsBat.Open "BATHDR", cnn, adOpenDynamic, adLockPessimistic, adCmdTable
   rsCmp.Open "COMPANY", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
   qSQL = "SELECT CLINAME,  CLIENTID FROM STKNAME "
   qSQL = qSQL & " WHERE CLIENTID = " & iClient
   qSQL = qSQL & " ORDER BY CLINAME, CLIENTID"
   rsClient.Open qSQL, cnn, , , adCmdText
  
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '--
   If gblOptions = 1 Then
      qSql1 = "SELECT * FROM UNUSEDNOS WHERE SEQTYP = 'C' "
      qSql1 = qSql1 & "order by SEQTYP, UNUSED"
      rsUnused.Open qSql1, cnn, adOpenDynamic, adLockOptimistic, adCmdText
      InitAddNew
   End If
   '--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS023/Load"
   Unload Me
   
End Sub
Private Sub UpdateScreen()
Dim i As Integer
With rsMain
  If Not .EOF Then
      For i = 0 To dbc(0).Rows - 1
        bm = dbc(0).GetBookmark(i)
        If dbc(0).Columns(0).CellText(bm) = !TrnBatch Then
          dbc(0).Bookmark = dbc(0).GetBookmark(i)
          dbc(0) = dbc(0).Columns(0).CellText(bm)
          dbc(0).Enabled = False
          meb(0).Text = dbc(0).Columns(1).CellText(bm)
          Exit For
        End If
      Next i
      tbFld(1).Text = !Form
      dbc(1).MoveFirst
      For i = 0 To dbc(1).Rows - 1
        bm = dbc(1).GetBookmark(i)
        If dbc(1).Columns(1).CellText(bm) = !ClientID Then
          dbc(1).Bookmark = dbc(1).GetBookmark(i)
          dbc(1) = dbc(1).Columns(0).CellText(bm)
          Exit For
        End If
      Next i
      dbc_LostFocus (1)
      meb(2).Text = !IssDate
      meb(1).Text = !shares
      
  End If
End With
End Sub

Private Sub meb_GotFocus(Index As Integer)

Select Case Index
Case 0, 2
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
   tbFld(1).SetFocus
  Case 1
    cmdUpdate.Enabled = True
    cmdUpdate.SetFocus
  Case 2
    If dbc(1).Enabled = True Then
       dbc(1).SetFocus
    Else
       cmdFind.SetFocus
    End If
  Case Else
  End Select
Case vbKeyUp
KeyCode = 0
  Select Case Index
  Case 0
    dbc(0).SetFocus
  Case 1
    If dbc(1).Enabled = True Then
       dbc(1).SetFocus
    Else
       cmdFind.SetFocus
    End If
  Case 2
    tbFld(1).SetFocus
  Case Else
  End Select
Case Else
End Select
End Sub


Private Sub ClearScreen()
  dbc(1) = ""
  meb(1).Mask = ""
  meb(1).Text = ""
  meb(2).Mask = ""
  meb(2).Text = ""
  tbFld(1) = ""
  lblLabels(12).Caption = ""
  lblLabels(13).Caption = ""
  lblLabels(4).Caption = ""
  
End Sub
Private Sub InitAddNew()
  ClearScreen
  Me.Caption = "New Stock Issue"
  cmdDelete.Enabled = False
  cmdUpdate.Enabled = False
End Sub
Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
 KeyCode = 0
 If Index = 1 Then
    meb(2).SetFocus
 End If
Case vbKeyUp
 If gblOptions = 1 Then
  If Index = 1 Then
    If iNew Then
       meb(1).SetFocus
    Else
       dbc(0).SetFocus
    End If
  End If
 End If
Case Else
End Select
End Sub

