VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3A32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3B32.OCX"
Begin VB.Form frmSIS082 
   Caption         =   "Bank Reconciliation Data Entry"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Enter a cheque bumber to edit or use find key to locate payment"
      Top             =   600
      Width           =   1935
      DataFieldList   =   "Column 0"
      AllowNull       =   0   'False
      _Version        =   196616
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
      Columns.Count   =   5
      Columns(0).Width=   2037
      Columns(0).Caption=   "ChequeNo"
      Columns(0).Name =   "ChequeNo"
      Columns(0).DataField=   "Column 4"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   8
      Columns(1).Width=   3572
      Columns(1).Caption=   "CliName"
      Columns(1).Name =   "Client Name"
      Columns(1).DataField=   "Column 0"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   50
      Columns(2).Width=   1852
      Columns(2).Caption=   "ClientId"
      Columns(2).Name =   "Client Id"
      Columns(2).Alignment=   1
      Columns(2).DataField=   "Column 1"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   11
      Columns(3).Width=   4498
      Columns(3).Caption=   "PayeeName"
      Columns(3).Name =   "Payee Name"
      Columns(3).DataField=   "Column 2"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   50
      Columns(4).Width=   2037
      Columns(4).Caption=   "Amount"
      Columns(4).Name =   "Amount"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 3"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   12
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_A.SSDBOptSet OptBtn 
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      ToolTipText     =   "Indicates whether the cheque will be reconciled or not"
      Top             =   960
      Width           =   1245
      _Version        =   196611
      _ExtentX        =   2196
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "&Yes"
      BackColor       =   -2147483643
      IndexSelected   =   1
      NumberOfButtons =   2
      Buttons.Button(0).OptionValue=   "0"
      Buttons.Button(0).Caption=   "&No"
      Buttons.Button(0).Mnemonic=   78
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   29
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   31
      Buttons.Button(0).PictureRight=   30
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   82
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(1).OptionValue=   "1"
      Buttons.Button(1).Caption=   "&Yes"
      Buttons.Button(1).Mnemonic=   89
      Buttons.Button(1).Value=   -1  'True
      Buttons.Button(1).TextLeft=   15
      Buttons.Button(1).TextTop=   16
      Buttons.Button(1).TextRight=   33
      Buttons.Button(1).TextBottom=   30
      Buttons.Button(1).ButtonTop=   16
      Buttons.Button(1).ButtonRight=   13
      Buttons.Button(1).ButtonBottom=   29
      Buttons.Button(1).PictureLeft=   35
      Buttons.Button(1).PictureTop=   16
      Buttons.Button(1).PictureRight=   34
      Buttons.Button(1).PictureBottom=   30
      Buttons.Button(1).ButtonToColTop=   16
      Buttons.Button(1).ButtonToColRight=   82
      Buttons.Button(1).ButtonToColBottom=   30
      Buttons.Button(1).ButtonBitmapID=   2
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   2280
      TabIndex        =   6
      ToolTipText     =   "Locates all records for the payee entered"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Enabled         =   0   'False
      Height          =   300
      Left            =   4440
      TabIndex        =   4
      ToolTipText     =   "saves any changes to the Bank recon file"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   5520
      TabIndex        =   7
      ToolTipText     =   "terminates the process with terminates the process"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   3360
      TabIndex        =   5
      ToolTipText     =   "Clears the screen"
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox tb 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   1
      ToolTipText     =   "Format YYYYMM Eg 199902 "
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "DecDate"
      Height          =   255
      Index           =   12
      Left            =   4800
      TabIndex        =   22
      Top             =   2280
      Width           =   1140
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Dividend Date:"
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
      Left            =   3120
      TabIndex        =   21
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Accounting Period:"
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
      TabIndex        =   20
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Cheque Number:"
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
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   1740
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Stockholder: "
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
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Payee Name:"
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
      TabIndex        =   17
      Top             =   1920
      Width           =   1860
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Cheque  Date:"
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
      Left            =   75
      TabIndex        =   16
      Top             =   2280
      Width           =   1785
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Cheque Amount:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   1620
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
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
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6495
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
      Index           =   3
      Left            =   0
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Reconciled:"
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
      Left            =   3480
      TabIndex        =   11
      Top             =   1080
      Width           =   1185
   End
   Begin VB.Label lbl 
      Caption         =   "CHQAMT"
      Height          =   255
      Index           =   10
      Left            =   1920
      TabIndex        =   10
      Top             =   2640
      Width           =   1740
   End
   Begin VB.Label lbl 
      Caption         =   "CHQDAT"
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   9
      Top             =   2280
      Width           =   1140
   End
   Begin VB.Label lbl 
      Caption         =   "PayeeName"
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   8
      Top             =   1920
      Width           =   4620
   End
   Begin VB.Label lbl 
      Caption         =   "ClientId && CliName"
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   4620
   End
End
Attribute VB_Name = "frmSIS082"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As ADODB.Recordset
Dim rsLookup As ADODB.Recordset
Dim iOpenLookup As Integer
Dim iErr As Integer
Dim iInputRecon As Integer
Dim sPostDte As String

Private Sub cmdCancel_Click()
ShutDown
Unload Me
End Sub

Private Sub cmdClear_Click()
tb(1) = ""
dbc = " "
Clear_Display
dbc.SetFocus
End Sub

Private Sub cmdFind_Click()
'-- do not use the find key to do a correction
'-- use it to record a new reconciliation.
'-- if the recon date needs correcting,
'-- enter the cheque number directly in the combo box
'------------------------------------------------
Dim x As Integer
Dim cChk As Integer, qCli As String
Dim sWhere As String, sRowinfo As String
On Error GoTo cmdFind_Click_err
'--
Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
    .cbWhere.AddItem "CliName"
    .cbWhere.AddItem "ClientId"
    .cbWhere.AddItem "PayeeName"
    .cbWhere.ListIndex = 0
    .cbOptions.ListIndex = 0
    .lbl(3).Caption = " "
    .OptBtn.Buttons(0).Visible = False
    .OptBtn.Buttons(1).Visible = False
    .Show vbModal
    '----------------------------
    '-------- main line ---------
    '----------------------------
    If .tbFind.Text = vbNullString Then
    Else
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      sWhere = Trim(.tbFind.Text)
      x = .cbWhere.ListIndex
      qCli = "SELECT a.*, cliname FROM STKBANK a, " _
             & "STKNAME b WHERE " _
      '---
      If sWhere <> "" Then
          Select Case .cbOptions.ListIndex
          Case 0 ' Exact Match
              If x = 1 Then
                qCli = qCli & "a." & .cbWhere
                qCli = qCli & " = " & Val(.tbFind.Text)
              End If
              If x = 0 Then
                qCli = qCli & "b." & .cbWhere
                qCli = qCli & " = '" & sWhere & "'"
              End If
              If x = 2 Then
                qCli = qCli & "a." & .cbWhere
                qCli = qCli & " = '" & sWhere & "'"
              End If
          Case 1 ' Starts With
              sWhere = Trim(.tbFind.Text) & "%"
              If x = 1 Or x = 2 Then
                 qCli = qCli & "a." & .cbWhere
              End If
              If x = 0 Then
                 qCli = qCli & "b." & .cbWhere
              End If
              qCli = qCli & " like '" & sWhere & "'"
          Case 2 ' Ends With
               sWhere = "%" & Trim(.tbFind.Text)
               If x = 1 Or x = 2 Then
                 qCli = qCli & "a." & .cbWhere
              End If
              If x = 0 Then
                 qCli = qCli & "b." & .cbWhere
              End If
              qCli = qCli & " like '" & sWhere & "'"
           Case 3 ' AnyWhere
               sWhere = "%" & Trim(.tbFind.Text) & "%"
               If x = 1 Or x = 2 Then
                 qCli = qCli & "a." & .cbWhere
              End If
              If x = 0 Then
                 qCli = qCli & "b." & .cbWhere
              End If
              qCli = qCli & " like '" & sWhere & "'"
           End Select
      End If
      qCli = qCli & " and a.clientid = b.clientid " _
             & "and reconind = false "
      qCli = qCli & " ORDER BY CLINAME, a.CLIENTID"
      rsLookup.Open qCli, gblFileName, , , adCmdText
      iOpenLookup = True
      If Not rsLookup.EOF Then
         FillCombo
      End If
      rsLookup.Close
      iOpenLookup = False
    End If
  End With
Unload frmFind
Set frmFind = Nothing
Exit Sub
cmdFind_Click_err:
  csvShowError "SIS082/CmdFind"
  ShutDown
  Unload Me
End Sub

Private Sub cmdUpdate_Click()
Dim sql As String
Dim x As Integer
If iInputRecon <> OptBtn.IndexSelected Then
  sql = "Update stkbank set reconind = " _
        & OptBtn.IndexSelected & ", foliomth = '" _
        & tb(1) & "', recondat = '" _
        & Format(Now(), "dd-mmm-yyyy") & "' where chqnum = " & dbc
  x = csvADODML(sql)
Else
  If sPostDte <> tb(1) And Not IsNothing(sPostDte) Then
    sql = "Update stkbank set foliomth = " & tb(1) _
          & " where chqnum = " & dbc
    x = csvADODML(sql)
  Else
   iErr = 172
   csvShowUsrErr iErr, "Bank Recon Data Entry"
   OptBtn.SetFocus
   Exit Sub
  End If
End If
cmdClear_Click
End Sub

Private Sub dbc_Validate(Cancel As Boolean)
Dim qSQL As String
Dim i As Integer, bm As Variant
'--
If IsNothing(dbc) Then
 iErr = 28
 Cancel = True
 GoTo dbc_Validate_Err
End If
'--- validate cheque number
'--------------------------
qSQL = "SELECT a.*, cliname FROM STKBANK a, " _
       & "STKNAME b WHERE a.chqnum = " _
       & Val(dbc) & " and a.clientid = b.clientid"
rsLookup.Open qSQL, gblFileName, , , adCmdText
iOpenLookup = True
If rsLookup.EOF Then ' no match
   iErr = 166
   Cancel = True
   rsLookup.Close
   iOpenLookup = False
   GoTo dbc_Validate_Err
End If
'-- populate fields for display
'------------------------------
With rsLookup
  lbl(4) = !clientid & "- " & Trim(!cliname)
  If Not IsNull(!payeename) Then
    lbl(6) = !payeename
  Else
    lbl(6) = !cliname
  End If
  If Not IsNull(!foliomth) Then
    tb(1) = !foliomth
    sPostDte = tb(1)
  End If
    
  lbl(8) = Format(!CHQDAT, "dd-mmm-yyyy")
  lbl(10) = Format(!CHQAMT, "$##,###.00")
  lbl(12) = Format(!DECDATE, "dd-mmm-yyyy")
  If !reconind = True Then
     OptBtn.IndexSelected = 1
     iInputRecon = 1
  Else
     OptBtn.IndexSelected = 0
     iInputRecon = 0
  End If
  .Close
  iOpenLookup = False
End With

dbc_Validate_Exit:
  Exit Sub
dbc_Validate_Err:
  csvShowUsrErr iErr, "Bank Recon Data Entry"
  GoTo dbc_Validate_Exit
'--
End Sub

Private Sub Form_Load()
Dim sql As String, indx As Integer
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = gblVersn
'--
csvCenterForm Me, gblMDIFORM
For indx = 4 To 12 Step 2
 lbl(indx) = " "
Next
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Set cnn = New ADODB.Connection
cnn.Open gblFileName
Set rsLookup = New ADODB.Recordset
iOpenLookup = False
End Sub

 Private Sub FillCombo()
Dim sRowinfo As String
With rsLookup
    .Requery
    If Not .EOF And Not .BOF Then
      .MoveFirst
      dbc.RemoveAll
      Do While Not .EOF
         sRowinfo = !CHQNUM! & Chr(9) & !cliname & Chr(9) _
                    & !clientid & Chr(9) & !payeename _
                    & Chr(9) & !CHQAMT
         dbc.AddItem sRowinfo
         If dbc.Row = 0 Then dbc = !CHQNUM
         .MoveNext
      Loop
    End If
End With
End Sub
Private Sub ShutDown()
 If iOpenLookup = True Then rsLookup.Close
 Set cnn = Nothing
 Set rsMain = Nothing
 Set rsLookup = Nothing
 Set frmSIS082 = Nothing
End Sub


Private Sub tb_Validate(Index As Integer, Cancel As Boolean)
Dim sYear As String, sMonth As String
Select Case Index
Case 1 ' accounting period
  If IsNothing(tb(1)) Then
    iErr = 28
    Cancel = True
    GoTo Validate_Err
  End If
  '-- Validate year portion of accounting period
  sYear = left(tb(1), 4)
  If Val(sYear) < 1999 Or Val(sYear) > 2100 Then
     iErr = 170
     Cancel = True
     GoTo Validate_Err
  End If
  '-- validate month portion of accounting period
  sMonth = right(tb(1), 2)
  If Val(sMonth) < 1 Or Val(sMonth) > 12 Then
    iErr = 171
    Cancel = True
    GoTo Validate_Err
  End If
'--
cmdUpdate.Enabled = True
cmdUpdate.SetFocus
lbl(11).Visible = True
OptBtn.Visible = True
Validate_Exit:
  Exit Sub
Validate_Err:
  csvShowUsrErr iErr, "Bank Recon Data Entry"
  GoTo Validate_Exit
'--
End Select
End Sub

Private Sub Clear_Display()
Dim indx As Integer
For indx = 4 To 12 Step 2
 lbl(indx) = " "
Next
cmdUpdate.Enabled = False
End Sub



