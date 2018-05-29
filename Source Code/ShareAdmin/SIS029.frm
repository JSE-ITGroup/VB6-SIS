VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS029 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Broker's Certification"
   ClientHeight    =   3615
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "SIS029.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6795
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   300
      Left            =   1440
      TabIndex        =   24
      ToolTipText     =   "Pressing this button will activate the search program to locate a broker."
      Top             =   3240
      Width           =   975
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      ToolTipText     =   "Enter the number of shares to certify. Must be less than or equal the available shares."
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "#,###"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Reverse"
      Height          =   300
      Left            =   3600
      TabIndex        =   15
      ToolTipText     =   "Releases the certification by deleting all reference from the system."
      Top             =   3240
      Width           =   975
   End
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
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Select a a Broker from the list"
      Top             =   2400
      Width           =   3735
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
      _ExtentX        =   6588
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   2520
      TabIndex        =   12
      ToolTipText     =   "Clears the screen on insertions."
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5760
      TabIndex        =   7
      ToolTipText     =   "Returns to the previous menu."
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Enabled         =   0   'False
      Height          =   300
      Left            =   4680
      TabIndex        =   6
      ToolTipText     =   "Saves the edit/new certification"
      Top             =   3240
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
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   6840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lbl 
      Caption         =   "Available Shares"
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
      Index           =   10
      Left            =   1920
      TabIndex        =   23
      Top             =   1920
      Width           =   1455
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
      TabIndex        =   22
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Shares to Certify:"
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
      Left            =   480
      TabIndex        =   21
      Top             =   2760
      Width           =   1575
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
      Left            =   960
      TabIndex        =   20
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   6840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Available Shares:"
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
      Left            =   240
      TabIndex        =   19
      Top             =   1920
      Width           =   1575
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
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   1150
      Width           =   540
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Certification "
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
      TabIndex        =   16
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   11
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS029"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X, iNew As Integer, iEOF As Integer
Dim rsClient As ADODB.Recordset
Dim rsBrkCrt As ADODB.Recordset
Dim rsPool As ADODB.Recordset
Dim rsBat As ADODB.Recordset
Dim rsVerBat As ADODB.Recordset
Dim rsVerFrm As ADODB.Recordset
Dim rsVerCrt As ADODB.Recordset
Dim errLoop As Error
Dim errs1 As Error
Dim iStocks As Double
Dim iBrokerId As Long, iToBrkId As Long, sForm As String
Private Sub CallRefresh()
  frmSIS027.Visible = True 'used to trigger activate event
  frmSIS027.Visible = False
  frmSIS028.Visible = True
  frmSIS028.Visible = False
End Sub
Private Sub FillCombo(i As Integer)
Dim sRowinfo As String
With rsClient
    .Requery
    If Not .EOF And Not .BOF Then
      .MoveFirst
      dbc(i).RemoveAll
      Do While Not .EOF
         sRowinfo = !CliName & Chr(9) & !ClientID
         dbc(i).AddItem sRowinfo
         If dbc(i).Row = 0 Then dbc(i) = !CliName
         .MoveNext
      Loop
    End If
End With
End Sub
Function IsValid() As Integer
Dim iErr As Integer, dtefld As Date, qSQL
Dim sElable As String
sElable = "Broker Certification"
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
 If tbFld(1).Text = "" Then 'Invalid form
   iErr = 140
   csvShowUsrErr iErr, sElable
   tbFld(1).SetFocus
   GoTo Validate_Exit
End If
tbFld(1).Text = UCase(tbFld(1).Text)
'--
If gblOptions = 1 Then 'check for duplicate form
  
  qSQL = "SELECT FORM from STKACTIV where FORM = '"
  qSQL = qSQL & tbFld(1) & "'"
  rsVerFrm.Open qSQL, cnn, , , adCmdText
  If Not rsVerFrm.EOF Then
     iErr = 141
     csvShowUsrErr iErr, sElable
     tbFld(1).SetFocus
     rsVerFrm.Close
     GoTo Validate_Exit
  End If
  rsVerFrm.Close
  qSQL = "SELECT FORMNO FROM STKBKCRT WHERE FORMNO = '"
  qSQL = qSQL & tbFld(1) & "'"
  rsVerCrt.Open qSQL, cnn, , , adCmdText
  If Not rsVerCrt.EOF Then
     iErr = 141
     csvShowUsrErr iErr, sElable
     tbFld(1).SetFocus
     rsVerCrt.Close
     GoTo Validate_Exit
  End If
  rsVerCrt.Close
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
 If dbc(1) = "" Then  ' no broker selected
   iErr = 153
   csvShowUsrErr iErr, sElable
   dbc(1).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If meb(2) = "" Then
    iErr = 159
    csvShowUsrErr iErr, sElable
    meb(2).SetFocus
    GoTo Validate_Exit
 End If
 '--
 If meb(2) > Val(lbl(10).Caption) Then
   iErr = 160
   csvShowUsrErr iErr, sElable
   meb(2).SetFocus
   GoTo Validate_Exit
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
Set rsPool = Nothing
Set rsBrkCrt = Nothing
Set rsBat = Nothing
Set rsVerBat = Nothing
Set rsVerFrm = Nothing
Set rsVerCrt = Nothing
frmSIS028.Visible = True
cnn.clsoe
Unload Me
End Sub

Private Sub cmdClear_Click()
ClearScreen
If gblOptions = 2 Then UpdateScreen
End Sub

Private Sub cmdDelete_Click()
Dim imsg As Integer, X As Integer
imsg = 133
If csvYesNo(imsg, "Broker Certification") Then

   iStocks = rsBrkCrt!shares * -1
   'If UpdBrokerPool(3, iBrokerId, iStocks, cnn) Then
     rsBrkCrt!Status = "R"
     rsBrkCrt.Update
     cmdCancel_Click
   'End If
End If
End Sub

Private Sub cmdFind_Click()
Dim X As Integer
Dim cChk As Integer, qCli As String
Dim sWhere As String, sRowinfo As String

Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
     .cbWhere.Clear
    .cbWhere.AddItem "CliName"
    .cbWhere.AddItem "ClientId"
    .cbWhere.ListIndex = 0
    .cbOptions.ListIndex = 0
    .lbl(3).Visible = False
    .optBtn.Visible = False
    .Show vbModal
    '----------------------------
    '-------- main line ---------
    '----------------------------
    If .tbFind.Text = vbNullString Then
    Else
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      sWhere = Trim(.tbFind.Text)
      X = .cbWhere.ListIndex
      qCli = "SELECT CLINAME,  CLIENTID  FROM STKNAME WHERE "
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
          Case 2 ' Ends With
               sWhere = "%" & Trim(.tbFind.Text)
               qCli = qCli & " like '" & sWhere & "'"
           Case 3 ' AnyWhere
               sWhere = "%" & Trim(.tbFind.Text) & "%"
               qCli = qCli & " like '" & sWhere & "'"
           End Select
      End If
      qCli = qCli & " AND CATCODE = 'SB' AND CLIENTID <> "
      qCli = qCli & iBrokerId
      qCli = qCli & " ORDER BY CLINAME, CLIENTID"
      rsClient.Open qCli, cnn, , , adCmdText
      If Not rsClient.EOF Then
          X = 1
          FillCombo (X) ' Client selling
          cmdUpdate.Enabled = True
      End If
      
      rsClient.Close
    End If
  End With
Unload frmFind
Set frmFind = Nothing
Exit Sub
cmdFind_Click_err:
  MsgBox "SIS029/CmdFind"
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo cmdUpdate_Err
If IsValid Then
  '--
  iStocks = Val(meb(2).Text)
  If gblOptions = 2 Then
     iStocks = iStocks - rsBrkCrt!shares
  End If
  If gblOptions = 1 Then
  'If UpdBrokerPool(3, iBrokerId, iStocks, cnn) Then
    If gblOptions = 1 Then
           rsBrkCrt.AddNew
    End If
    With rsBrkCrt
    '---
      !BROKERID = iBrokerId
      !formno = tbFld(1)
      !DTECRTFY = DateValue(meb(1).Text)
      !TOBROKERID = Val(dbc(1).Columns(1).Text)
      !shares = Val(meb(2).Text)
      !STACHGDTE = DateValue(meb(0).Text)
      !batch = dbc(0).Text
      !Status = "H"  ' Held pending transfer
      .Update
    End With
   
    If gblOptions = 1 Then
      CallRefresh
      Form_Activate
      InitAddNew
    Else
      CallRefresh
      cmdCancel_Click
    End If
 
  End If
End If
'---
Done:
 Exit Sub
'--
cmdUpdate_Err:
  
  MsgBox "SIS029/cmdUpdate"
  csvLogError "SIS029/cmdUpdate", Err.Number, Err.Description
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
         tbFld(1).SetFocus
      End If
    Case 1
      meb(2).SetFocus
    Case Else
  End Select
     
 Case vbKeyUp
  KeyCode = 0
  Select Case Index
    Case 1
       meb(1).SetFocus
    Case Else
  End Select
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
If gblOptions = 2 Then
  UpdateScreen
   Me.Caption = "Edit Broker Certification"
   tbFld(1).Enabled = False
   For i = 0 To 1
     dbc(i).Enabled = False
     meb(i).Enabled = False
   Next
   cmdUpdate.Enabled = True
   cmdDelete.Enabled = True
   cmdFind.Enabled = False
End If
lbl(10).Caption = frmSIS028.lbl(4).Caption
'--
Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
 If Err = -2147168242 Then ' no current transactions
   Resume 0
 Else
   MsgBox "SIS029/Activate"
   csvLogError "SIS029/Activate", Err.Number, Err.Description
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
   Set rsPool = New ADODB.Recordset
   Set rsBrkCrt = New ADODB.Recordset
   Set rsBat = New ADODB.Recordset
   Set rsVerBat = New ADODB.Recordset
   Set rsVerFrm = New ADODB.Recordset
   Set rsVerCrt = New ADODB.Recordset
   '''Set cnn = New ADODB.Connection
   '----------------
   '-- Unpack gblfilekey
   '--------------------
   X = InStr(1, gblFileKey, ";", 1)
   iBrokerId = Val(Mid(gblFileKey, 1, X - 1))
   If gblOptions = 2 Then
       Y = InStr(X + 1, gblFileKey, ";", 1)
       sForm = Mid(gblFileKey, X + 1, Y - X - 1)
       iToBrkId = Val(Mid(gblFileKey, Y + 1, (Len(gblFileKey) - Y)))
   Else
       InitAddNew
   End If
   lbl(8).Caption = frmSIS028.lbl(2).Caption
   
   '-----------------------
   '-- open tables --------
   '-----------------------
   cnn.Open
   qSQL = "SELECT a.*, b.SHRHELD from STKBKCRT as a"
   qSQL = qSQL & " inner join STKBRKPL as b on "
   qSQL = qSQL & " a.BROKERID  = b.BROKERID "
   qSQL = qSQL & "  Where a.BROKERID = " & iBrokerId & " and"
   qSQL = qSQL & " FORMNO = '" & sForm & "' and"
   qSQL = qSQL & " TOBROKERID = " & iToBrkId
   rsBrkCrt.Open qSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
   rsBat.Open "BATHDR", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
   '''cnn.Open cnn
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS029/Load"
  csvLogError "SIS029/Load", Err.Number, Err.Description
  Unload Me
   
End Sub
Private Sub UpdateScreen()
Dim qSQL As String
Dim i As Integer, bm As Variant
With rsBrkCrt
  If Not .EOF Then
      dbc(0) = !batch
      meb(0).Text = !STACHGDTE
      tbFld(1) = !formno
      meb(1).Text = !DTECRTFY
      dbc(1) = !TOBROKERID
      qSQL = "SELECT CLINAME,  CLIENTID FROM STKNAME "
      qSQL = qSQL & " WHERE CLIENTID = " & !TOBROKERID
      rsClient.Open qSQL, cnn, , , adCmdText
      i = 1
      FillCombo (i)
      rsClient.Close
      meb(2) = !shares
  End If
End With
End Sub
Private Sub meb_GotFocus(Index As Integer)
Select Case Index
Case 0
  meb(Index).Mask = "##-???-####"
Case 1
  If Not IsDate(meb(1)) Then meb(Index).Mask = "##-???-####"
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
     dbc(1).SetFocus
  Case 2
     cmdUpdate.SetFocus
  End Select
Case vbKeyUp
KeyCode = 0
Select Case Index
Case 0
   If dbc(0).Enabled = True Then dbc(0).SetFocus
Case 1
   tbFld(1).SetFocus
Case 2
   dbc(1).SetFocus
End Select
Case Else
End Select
End Sub



Private Sub ClearScreen()
Dim i As Integer
    dbc(1) = ""
  For i = 1 To 2
   meb(i).Mask = ""
   meb(i).Text = ""
  Next
  tbFld(1) = ""
  
End Sub



Private Sub meb_LostFocus(Index As Integer)
Select Case Index
Case 0
  If Not IsDate(meb(1)) Then meb(1) = meb(0)
Case Else
End Select
End Sub


Private Sub InitAddNew()
  ClearScreen
  Me.Caption = "New Broker Certification"
  sForm = ""
  iToBrkId = 0
  cmdDelete.Enabled = False
  cmdUpdate.Enabled = False
  meb(0).Enabled = False
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
