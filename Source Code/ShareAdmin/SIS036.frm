VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS036 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merge Accounts"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS036.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7890
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H80000004&
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3600
      MaskColor       =   &H000000FF&
      TabIndex        =   15
      ToolTipText     =   "Search for Account number"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H80000004&
      Caption         =   "&Find"
      Height          =   300
      Left            =   4680
      MaskColor       =   &H000000FF&
      TabIndex        =   14
      ToolTipText     =   "Search for Account number"
      Top             =   4800
      Width           =   975
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Client Account number to merge from"
      Top             =   600
      Width           =   1695
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
      RowHeight       =   397
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "ClientId"
      Columns(0).Name =   "ClientId"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   10
      Columns(1).Width=   5715
      Columns(1).Caption=   "CliName"
      Columns(1).Name =   "CliName"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   50
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   93
      ForeColor       =   4194304
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   252
      Left            =   5400
      TabIndex        =   12
      Top             =   4320
      Width           =   1332
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   2652
      Left            =   600
      TabIndex        =   7
      Top             =   1560
      Width           =   6252
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   2540
      Columns(0).Caption=   "Cert No"
      Columns(0).Name =   "CertNo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   10
      Columns(1).Width=   2709
      Columns(1).Caption=   "DateIssue"
      Columns(1).Name =   "DateIssue"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd-mmm-yyyy"
      Columns(1).FieldLen=   11
      Columns(2).Width=   2032
      Columns(2).Caption=   "Status"
      Columns(2).Name =   "Status"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   10
      Columns(3).Width=   3200
      Columns(3).Caption=   "Shares"
      Columns(3).Name =   "Shares"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   3
      Columns(3).FieldLen=   10
      _ExtentX        =   11028
      _ExtentY        =   4678
      _StockProps     =   79
      Caption         =   "Merged Certificates"
      ForeColor       =   4194304
      Enabled         =   0   'False
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
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H80000004&
      Caption         =   "C&ommit"
      Enabled         =   0   'False
      Height          =   300
      Left            =   5760
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      ToolTipText     =   "Accepts input and merge accounts"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   6840
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      ToolTipText     =   "Returns to main menu"
      Top             =   4800
      Width           =   975
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Client Account Number to Merge to"
      Top             =   1080
      Width           =   1695
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
      RowHeight       =   397
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "ClientId"
      Columns(0).Name =   "ClientId"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   10
      Columns(1).Width=   5715
      Columns(1).Caption=   "CliName"
      Columns(1).Name =   "CliName"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   50
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   93
      ForeColor       =   4194304
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Active Shares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   3840
      TabIndex        =   13
      Top             =   4320
      Width           =   1572
   End
   Begin VB.Label lbllabels 
      Caption         =   "CliName"
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
      Index           =   6
      Left            =   3360
      TabIndex        =   11
      Top             =   1080
      Width           =   4260
   End
   Begin VB.Label lbllabels 
      Caption         =   "CliName:"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   600
      Width           =   4140
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Merge To Client Number:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Merge From Client Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1380
   End
   Begin VB.Label lbllabels 
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
   Begin VB.Label lbllabels 
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
      X2              =   7920
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lbllabels 
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
      Height          =   372
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7812
   End
End
Attribute VB_Name = "frmSIS036"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, sql As String
Dim rsName As ADODB.Recordset
Dim rsCert As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim OpenErr As Integer
Dim iOpenName As Integer
Dim iOpenCert As Integer
Dim iErr As Integer
Dim ToAcnt As Long
Dim FrAcnt As Long
Dim iYtd As Currency
Dim aFrJoint(2) As String, aToJoint(2) As String

Private Sub cmdCancel_Click()
Shutdown
Unload Me
End Sub

Private Sub cmdClear_Click()
ClearScreen
End Sub

Private Sub cmdFind_Click()
Dim X As Integer, i As Integer
Dim cChk As Integer, qCli As String
Dim sWhere As String, sRowinfo As String

Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
     .cbWhere.Clear
    .cbWhere.AddItem "ClientId"
    .cbWhere.AddItem "CliName"
    .cbWhere.ListIndex = 0
    .cbOptions.ListIndex = 0
    .lbl(3).Caption = "Find Client"
    .Optbtn.Buttons(0).Caption = "Merge &From"
    .Optbtn.Buttons(1).Caption = "Merge &To"
    .Show vbModal
    '----------------------------
    '-------- main line ---------
    '----------------------------
    If .tbFind.Text = vbNullString Then
    Else
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      sWhere = Trim(.tbFind.Text)
      X = .cbWhere.ListIndex
      qCli = "SELECT CLINAME,  CLIENTID FROM STKNAME WHERE "
      qCli = qCli & .cbWhere
      '---
      If sWhere <> "" Then
          Select Case .cbOptions.ListIndex
          Case 0 ' Exact Match
              If X = 0 Then
                qCli = qCli & " = " & Val(.tbFind.Text)
              Else
                qCli = qCli & " = '" & sWhere & "'"
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
      If .Optbtn.IndexSelected = 0 Then
         i = 0
      Else
         i = 1
      End If
      qCli = qCli & " ORDER BY CLINAME, CLIENTID"
      rsName.Open qCli, SpCon, , , adCmdText
      iOpenName = True
      If Not rsName.EOF Then
          FillCombo (i)
      End If
      rsName.Close
      iOpenName = 0
    End If
  End With
Unload frmFind
Set frmFind = Nothing
Exit Sub
cmdFind_Click_err:
  MsgBox "SIS036/CmdFind"
End Sub

Private Sub cmdOk_Click()
Dim i As Integer, bm As Variant
On Error GoTo CmdOk_click_err
'-- Update the certs with the TOACNT Client Id
'---------------------------------------------
If Not IsValid Then Exit Sub
With grd
 For i = 0 To .Rows - 1
   bm = .GetBookmark(i)
   sql = "update certmst set clientid = " & Val(dbc(1))
   sql = sql & " where certno = " & Val(.Columns(0).CellText(bm))
   X = csvADODML(sql, SpCon)
  Next i
  .RemoveAll
End With
'-- Delete the FromAcnt
'----------------------
sql = "delete from stkname where clientid = " & Val(dbc(0))
X = csvADODML(sql, SpCon)
'-- Update the ToACNT with shares and  payment information
'---------------------------------------------------------
sql = "update stkname set shares = shares + " & Val(meb) _
        & ", ytdpymnt = ytdpymnt + " & iYtd _
        & " where clientid = " & Val(dbc(1))
X = csvADODML(sql, SpCon)
'-- Update the Payment History with the TOACNT clientid
'------------------------------------------------------
sql = "update stkbank set clientid = " & Val(dbc(1)) _
      & " where clientid = " & Val(dbc(0))
X = csvADODML(sql, SpCon)
ClearScreen
CmdOk_click_exit:
Exit Sub
CmdOk_click_err:
  MsgBox "frmSIS036/CmdOk_Click"
  cmdCancel_Click
End Sub

Private Sub dbc_CloseUp(Index As Integer)
Dim bm As Variant, i As Integer
Select Case Index
Case 0, 1
For i = 0 To dbc(Index).Rows - 1
     bm = dbc(Index).GetBookmark(i)
     If dbc(Index).Columns(0).CellText(bm) = dbc(Index) Then
       If Index = 0 Then
        FrAcnt = dbc(Index).Columns(0).CellText(bm)
        lblLabels(9).Caption = dbc(Index).Columns(1).CellText(bm)
       Else
        ToAcnt = dbc(Index).Columns(0).CellText(bm)
        lblLabels(6).Caption = dbc(Index).Columns(1).CellText(bm)
       End If
       Exit For
     End If
   Next
End Select
End Sub

Private Sub dbc_Validate(Index As Integer, Cancel As Boolean)

If dbc(Index) = "" Then
   Exit Sub
End If
'--
sql = "Select a.Clientid, cliname, ytdpymnt, jntname1, jntname2, jntname3 " _
      & "from StkName a left join stkjoint b " _
      & "on a.clientid = b.clientid " _
      & "where a.clientId = " & Val(dbc(Index))
rsName.Open sql, SpCon, , , adCmdText
iOpenName = True
With rsName
   If .EOF Then
    iErr = 184
    csvShowUsrErr iErr, "Merge Accounts"
    rsName.Close
    iOpenName = 0
    Cancel = True
    Exit Sub
  End If
  If Index = 0 Then
    lblLabels(9).Caption = !CliName
    iYtd = !YtdPymnt
    If Not IsNull(!JNTNAME1) Then aFrJoint(0) = !JNTNAME1
    If Not IsNull(!JNTNAME2) Then aFrJoint(1) = !JNTNAME2
    If Not IsNull(!jntname3) Then aFrJoint(2) = !jntname3
  Else
    lblLabels(6).Caption = !CliName
    If Not IsNull(!JNTNAME1) Then aToJoint(0) = !JNTNAME1
    If Not IsNull(!JNTNAME2) Then aToJoint(1) = !JNTNAME2
    If Not IsNull(!jntname3) Then aToJoint(2) = !jntname3
    cmdOk.Enabled = True
  End If
  rsName.Close
  iOpenName = 0
  If Index = 0 Then GetCerts
  
  
End With

End Sub

Private Sub Form_Load()
'--
Dim i As Integer
Dim strTmp As String
On Error GoTo FL_ERR
'--
   csvCenterForm Me, gblMDIFORM
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   grd.RemoveAll
   '--
   OpenErr = 0
   iOpenName = 0
   iOpenCert = 0
   '''Set cnn = New ADODB.Connection
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
Loop
Screen.MousePointer = vbDefault

   Set rsCert = New ADODB.Recordset
   Set rsName = New ADODB.Recordset
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS036/Load"
  Unload Me
End Sub

Private Sub FillCombo(i As Integer)
Dim sRowinfo As String
With rsName
    .Requery
    If Not .EOF And Not .BOF Then
      .MoveFirst
      dbc(i).RemoveAll
      Do While Not .EOF
         sRowinfo = !ClientID & Chr(9) & !CliName
         dbc(i).AddItem sRowinfo
         .MoveNext
      Loop
    End If
End With
End Sub

Public Sub GetCerts()
Dim iShares As Long
On Error GoTo GetCert_Err
sql = "select * from certmst where clientid = " & Val(dbc(0))
rsCert.Open sql, SpCon, , , adCmdText
iOpenCert = True
iShares = 0
grd.RemoveAll
With rsCert
If Not .EOF Then
  .MoveFirst
  While Not .EOF
     grd.AddItem !certno & Chr(9) & !IssDate & Chr(9) _
                & !Status & Chr(9) & !shares
     If !Status <> "C" Then iShares = iShares + !shares
    .MoveNext
  Wend
  meb = iShares
End If
End With
rsCert.Close
iOpenCert = 0
GetCert_Exit:
Exit Sub
GetCert_Err:
  MsgBox "SIS036/GetCerts"
  cmdCancel_Click
End Sub

Private Sub Shutdown()
If iOpenName = True Then rsName.Close
If iOpenCert = True Then rsCert.Close
Set rsCert = Nothing
Set rsCert = Nothing
'''set cnn = nothing
Set frmSIS036 = Nothing
End Sub

Private Sub ClearScreen()
grd.RemoveAll
dbc(0).RemoveAll
dbc(1).RemoveAll
meb = ""
dbc(0) = ""
dbc(1) = ""
lblLabels(9).Caption = ""
lblLabels(6).Caption = ""
grd.Refresh

End Sub

Private Function IsValid()
Dim i As Integer
IsValid = False
For i = 0 To 2
 If Trim(aFrJoint(i)) <> Trim(aToJoint(i)) Then
  iErr = 185
  csvShowUsrErr iErr, "Merge Accounts"
  Exit Function
 End If
Next i
If dbc(0) = dbc(1) Then
   iErr = 186
   csvShowUsrErr iErr, "Merge Accounts"
  Exit Function
End If
IsValid = True
End Function



Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub
