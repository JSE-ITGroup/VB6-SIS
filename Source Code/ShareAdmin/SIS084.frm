VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSIS084 
   Caption         =   "Replace Cheque Data Entry"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   Icon            =   "SIS084.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPChqDate 
      Height          =   375
      Left            =   2160
      TabIndex        =   33
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   52690947
      CurrentDate     =   40729
   End
   Begin VB.OptionButton OptPayment 
      Caption         =   "Add to ACH file"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   32
      Top             =   3480
      Width           =   1815
   End
   Begin VB.OptionButton OptPayment 
      Caption         =   "Add to Finacle file"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   31
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox TxtCurrency 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   30
      Top             =   795
      Width           =   615
   End
   Begin VB.TextBox TxtAccountNo 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   29
      Top             =   810
      Width           =   2775
   End
   Begin VB.CommandButton CmdGo 
      Caption         =   "Go"
      Height          =   255
      Left            =   4800
      TabIndex        =   28
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Con&vert"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   300
      Left            =   3600
      TabIndex        =   26
      ToolTipText     =   "Locates all records for the payee entered"
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton OptPayment 
      Caption         =   "Cheque"
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   25
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox TxtAmount 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   24
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "ADD to List"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   22
      Top             =   2160
      Width           =   1695
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBCheque 
      Height          =   2295
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   7335
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   5
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   2566
      Columns(0).Caption=   "Cheque No"
      Columns(0).Name =   "Cheque No"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2752
      Columns(1).Caption=   "Cheque Amount"
      Columns(1).Name =   "Cheque Amount"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2275
      Columns(2).Caption=   "Payment Date"
      Columns(2).Name =   "Payment Date"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   7
      Columns(2).FieldLen=   256
      Columns(3).Width=   4604
      Columns(3).Caption=   "Payee"
      Columns(3).Name =   "Payee"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "AccountNo"
      Columns(4).Name =   "AccountNo"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      _ExtentX        =   12938
      _ExtentY        =   4048
      _StockProps     =   79
      Caption         =   "List of Cheques to be replaced"
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
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Enter a cheque number to edit"
      Top             =   480
      Width           =   2700
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
      Columns.Count   =   8
      Columns(0).Width=   2037
      Columns(0).Caption=   "ChequeNo"
      Columns(0).Name =   "ChequeNo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   8
      Columns(1).Width=   3572
      Columns(1).Caption=   "CliName"
      Columns(1).Name =   "Client Name"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   50
      Columns(2).Width=   1852
      Columns(2).Caption=   "ClientId"
      Columns(2).Name =   "Client Id"
      Columns(2).Alignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   11
      Columns(3).Width=   4498
      Columns(3).Caption=   "PayeeName"
      Columns(3).Name =   "Payee Name"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   50
      Columns(4).Width=   2037
      Columns(4).Caption=   "Amount"
      Columns(4).Name =   "Amount"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   12
      Columns(5).Width=   3200
      Columns(5).Caption=   "AccountNo"
      Columns(5).Name =   "AccountNo"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1402
      Columns(6).Caption=   "Currency"
      Columns(6).Name =   "Currency"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Caption=   "Pay Type"
      Columns(7).Name =   "Pay Type"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      _ExtentX        =   4762
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "saves any changes to the Bank recon file"
      Top             =   3780
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   4080
      TabIndex        =   5
      ToolTipText     =   "terminates the process with terminates the process"
      Top             =   3780
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "Clears the screen"
      Top             =   3780
      Width           =   975
   End
   Begin VB.TextBox tb 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   1
      ToolTipText     =   "Format YYYYMM Eg 199902 "
      Top             =   840
      Width           =   1215
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBCurrency 
      Height          =   255
      Left            =   4680
      TabIndex        =   27
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
      DataFieldList   =   "Column 0"
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
      Columns(0).Width=   3200
      Columns(0).Caption=   "Currency"
      Columns(0).Name =   "Currency"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      Enabled         =   0   'False
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Label Lbl 
      Caption         =   "New Cheque Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   0
      TabIndex        =   23
      Top             =   3240
      Width           =   1935
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "New Cheque  Date:"
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
      TabIndex        =   20
      Top             =   2760
      Width           =   1785
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   7680
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   7680
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Lbl 
      Caption         =   "DecDate"
      Height          =   255
      Index           =   12
      Left            =   4800
      TabIndex        =   19
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Label Lbl 
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
      TabIndex        =   18
      Top             =   1920
      Width           =   1620
   End
   Begin VB.Label Lbl 
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
      TabIndex        =   17
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label Lbl 
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
      TabIndex        =   16
      Top             =   480
      Width           =   1740
   End
   Begin VB.Label Lbl 
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
      TabIndex        =   15
      Top             =   1200
      Width           =   1740
   End
   Begin VB.Label Lbl 
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
      TabIndex        =   14
      Top             =   1560
      Width           =   1860
   End
   Begin VB.Label Lbl 
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
      TabIndex        =   13
      Top             =   1920
      Width           =   1785
   End
   Begin VB.Label Lbl 
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
      TabIndex        =   12
      Top             =   2280
      Width           =   1620
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
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6495
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   -120
      X2              =   7680
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Lbl 
      Caption         =   "CHQAMT"
      Height          =   255
      Index           =   10
      Left            =   1920
      TabIndex        =   8
      Top             =   2280
      Width           =   1740
   End
   Begin VB.Label Lbl 
      Caption         =   "CHQDAT"
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   7
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Label Lbl 
      Caption         =   "PayeeName"
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   4620
   End
   Begin VB.Label Lbl 
      Caption         =   "ClientId && CliName"
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   4620
   End
End
Attribute VB_Name = "frmSIS084"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As ADODB.Recordset
Dim rsLookup As ADODB.Recordset
Dim iOpenLookup As Integer
Dim iErr As String
Dim iInputRecon As Integer
Dim SpCon As ADODB.Connection
Dim sCHQAMT As String
Dim ChqAmt As Currency
Dim KeptAmt As Currency
Dim oCurr As String
Dim iCurr As String
Dim iClientID As Long, sCliName As String, sPayTYP As String
Dim ChqNo As String, oClientID As Long
Dim sql As String
Dim LocalCurrency As String

Private Sub CmdAdd_Click()
On Error GoTo Err_CmdAdd_Click
Dim StrSql As String

If iClientID <> oClientID Then
   MsgBox "This is not the same shareholder", vbOKOnly, "Shareholder Details"
   GoTo Exit_CmdAdd_Click
End If
If iCurr <> oCurr Then
   If ChqAmt = 0 Then
      iCurr = oCurr
   Else
      MsgBox "This is not the same currency", vbOKOnly, "Shareholder Details"
      GoTo Exit_CmdAdd_Click
   End If
End If

If dbc <> ChqNo Then
   StrSql = dbc & vbTab & lbl(10) & vbTab & lbl(8) & vbTab & lbl(6) & vbTab & TxtAccountNo
   SSDBCheque.AddItem StrSql
   ChqAmt = ChqAmt + KeptAmt
   TxtAmount = ChqAmt
   ChqNo = dbc
Else
    MsgBox "This cheque number has already been added", vbOKOnly, "Duplicated Chq No being added"
End If

Exit_CmdAdd_Click:
CmdAdd.Enabled = False
cmdConvert.Enabled = False
Exit Sub

Err_CmdAdd_Click:
MsgBox Err.Description, vbOKOnly, "Adding Cheque to be reconciled"
GoTo Exit_CmdAdd_Click

End Sub

Private Sub cmdCancel_Click()
Shutdown
Unload Me
End Sub

Private Sub cmdClear_Click()
DTPChqDate = Date
dbc = " "
Clear_Display
dbc.SetFocus
SSDBCheque.RemoveAll
TxtAmount = 0
End Sub

Private Sub cmdConvert_Click()
Dim adoRst As ADODB.Recordset
Dim DecDate As Date

DecDate = CDate(lbl(12))
Set adoRst = RunSP(SpCon, "usp_ConversionRate", 1, TxtCurrency, SSDBCurrency.Columns(0).Text, Mid(dbc.Columns(7).Text, 1, 1), Format(DecDate, "dd-mmm-yyyy"))
If adoRst!ExchRate = 0 Then
   MsgBox "No Rate was setup. This probably means you do not pay dividend in this currency for this register"
   GoTo Exit_CmdConvert_Click
End If

lbl(10) = Round((lbl(10) / adoRst!ExchRate), 2)
KeptAmt = lbl(10)
oCurr = SSDBCurrency.Columns(0).Text

Exit_CmdConvert_Click:
Exit Sub

Err_CmdConvert_Click:
MsgBox Err.Description, vbOKOnly, "Convert Process"
Resume Exit_CmdConvert_Click

End Sub

Private Sub CmdGo_Click()
On Error GoTo Err_CmdGo_Click
Dim qSQL As String
Dim i As Integer, bm As Variant
'--
If CmdAdd.Enabled = True Then
   GoTo Exit_CmdGo_Click
End If

If IsNothing(dbc) Then
 iErr = "No Cheque number was selected/entered"
  MsgBox iErr
 GoTo Exit_CmdGo_Click
End If
'--- validate cheque number
'--------------------------

Set rsLookup = RunSP(SpCon, "usp_SIS084Validate", 1, dbc)

frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
iOpenLookup = True
If rsLookup.State = 0 Then  ' no match
   iErr = "Cheque Number was not found"
   MsgBox iErr
'   rsLookup.Close
   iOpenLookup = False
   GoTo Exit_CmdGo_Click
End If
'-- populate fields for display
'------------------------------
With rsLookup
  dbc.RemoveAll
  Do While Not .EOF
    qSQL = !ChqNum & vbTab & !CliName & vbTab & !ClientID & vbTab
    qSQL = qSQL & !PayeeName & vbTab & !ChqAmt & vbTab & !AccountNo & vbTab & !Currency & vbTab
    If !PayTyp = "D" Then
       qSQL = qSQL & "Dividend"
    Else
       qSQL = qSQL & "Capital Distribution"
    End If
    dbc.AddItem qSQL
  .MoveNext
  Loop
  iOpenLookup = False
End With

Exit_CmdGo_Click:
Exit Sub
Err_CmdGo_Click:
sql = "The system could not complete the required action because of the following: " & vbCrLf
   sql = sql & Err.Description & " " & vbCrLf
   sql = sql & "We apologise for the inconvenience. Please Try again." & vbCrLf
   sql = sql & "If that fails, please contact your Systems Adminstrator"
   MsgBox sql
   GoTo Exit_CmdGo_Click
End Sub

Private Sub cmdUpdate_Click()
Dim X As Integer
Dim i As Integer
Dim iAccountNos As String
Dim NoLines As Integer
On Error GoTo cmdUpdate_Err

NoLines = 0
With SSDBCheque
     .MoveFirst
     sql = ""
     iAccountNos = ""
     For X = 0 To .Rows - 1
         sql = sql & .Columns(0).Text & ";"
         iAccountNos = iAccountNos & .Columns(4).Text & ";"
         NoLines = NoLines + 1
         .MoveNext
     Next X
End With

If NoLines < 1 Then
   MsgBox "Use the ADD Button to select the cheque(s) to be replaced"
   GoTo cmdUpdate_exit
End If
For i = 0 To 3
    If OptPayment(i).Value = True Then
       GoTo NextCommand
    End If
Next i

NextCommand:
X = RunSP(SpCon, "usp_ReplacementUpdate", 0, sql, NoLines, Format(DTPChqDate, "dd-mmm-yyyy"), _
     iClientID, ChqAmt, gblLoginName, i, iAccountNos, iCurr)
If X = 0 Then
   MsgBox "Replacement successfully concluded"
Else
    If X = 2 Then
       sql = "This cheque is awaiting confirmation to be replaced" & vbCrLf
       sql = sql & "Please have this approved or revoked before attempting this option"
       MsgBox sql, vbOKOnly, "Replacement process already started for this cheque"
    Else
       sql = "The system could not complete the replacement because of the following: " & vbCrLf
       sql = sql & Err.Description & " " & vbCrLf
       sql = sql & "We apologise for the inconvenience. Please Try again." & vbCrLf
       sql = sql & "If that fails, please contact your systems adminstrator"
       MsgBox sql
   End If
End If
iInputRecon = 0
ChqNo = " "
iCurr = ""
cmdClear_Click
cmdUpdate_exit:
 Exit Sub
cmdUpdate_Err:
sql = "The system could not complete the replacement because of the following: " & vbCrLf
sql = sql & Err.Description & " " & vbCrLf
sql = sql & "We apologise for the inconveneniece. Please Try again." & vbCrLf
sql = sql & "If that fails, please contact your systems adminstrator"
MsgBox sql, vbOKOnly, "Replacement Cheque Error"
  Shutdown
  Unload Me
End Sub

Private Sub dbc_Click()
Dim adoCurrency As ADODB.Recordset
Dim StrSql As String

With rsLookup
    .Filter = "ClientID = " & CDbl(dbc.Columns(2).Text) & " AND ChqAmt = " & CCur(dbc.Columns(4).Text)
  lbl(4) = !ClientID & "- " & Trim(!CliName)
  If iInputRecon = 0 Then
     oClientID = !ClientID
     oCurr = !Currency
     iInputRecon = 1
  End If
  iClientID = !ClientID
  iCurr = !Currency
  sCliName = !CliName
  sPayTYP = !PayTyp
  If Not IsNull(!PayeeName) Then
    lbl(6) = !PayeeName
  Else
    lbl(6) = !CliName
  End If
  If Not IsNull(!FolioMth) Then
    tb(1) = !FolioMth
  End If
    
  lbl(8) = Format(!ChqDat, "dd-mmm-yyyy")
  lbl(10) = Format(!ChqAmt, "$##,###.00")
  KeptAmt = !ChqAmt
  sCHQAMT = !ChqAmt
  lbl(12) = Format(!DecDate, "dd-mmm-yyyy")
  TxtAccountNo = !AccountNo
  TxtCurrency = !Currency
  CmdAdd.Enabled = True
End With

Set adoCurrency = RunSP(SpCon, "usp_SelectCurrencyList", 1)

Do While Not adoCurrency.EOF
   SSDBCurrency.Enabled = True
   If adoCurrency!Currency <> TxtCurrency Then
      With SSDBCurrency
           StrSql = adoCurrency!Currency & vbTab
           .AddItem StrSql
      End With
   End If
   If adoCurrency!CurrencyType = "L" Then
      LocalCurrency = adoCurrency!Currency
   End If
   adoCurrency.MoveNext
Loop

adoCurrency.Close
Set adoCurrency = Nothing
End Sub

Private Sub Form_Activate()
tb(1) = Format(Date, "dd-mmm-yyyy")
End Sub

Private Sub Form_Load()
Dim sql As String, indx As Integer
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
'--
csvCenterForm Me, gblMDIFORM
ChqAmt = 0
For indx = 4 To 12 Step 2
 lbl(indx) = " "
Next
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
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
iOpenLookup = False
DTPChqDate = Date
iInputRecon = 0
End Sub

Private Sub FillCombo()
Dim sRowinfo As String
With rsLookup
    '.Requery
    If Not .EOF And Not .BOF Then
      .MoveFirst
      dbc.RemoveAll
      Do While Not .EOF
         sRowinfo = !ChqNum! & Chr(9) & !CliName & Chr(9) _
                    & !ClientID & Chr(9) & !PayeeName _
                    & Chr(9) & !ChqAmt
         dbc.AddItem sRowinfo
         If dbc.Row = 0 Then dbc = !ChqNum
         .MoveNext
      Loop
    End If
End With
End Sub
Private Sub Shutdown()
 If iOpenLookup = True Then rsLookup.Close
 Set rsMain = Nothing
 Set rsLookup = Nothing
 Set frmSIS084 = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub lbl_Click(Index As Integer)
gblFileKey = iClientID
frmSIS073.Show 0
End Sub

Private Sub OptPayment_Click(Index As Integer)
On Error GoTo Err_OptPayment_Click
Dim adoRst As ADODB.Recordset

If Index <> 0 Then
   If iCurr <> LocalCurrency Then
      MsgBox "Only the local currency is allowed for this option", vbOKOnly, "Currency Error"
      GoTo Exit_OptPayment_Click
   End If
   If Index = 1 Then
      Set adoRst = RunSP(SpCon, "usp_IsAFinAccountPresent", 1, iClientID)
   Else
      If Index = 2 Then
         Set adoRst = RunSP(SpCon, "usp_IsACHAccountPresent", 1, iClientID)
      End If
   End If
   If adoRst.EOF Or adoRst!MndAcnt = 0 Then
      MsgBox "No Financial Institution Account was found"
      OptPayment(0).Value = True
      GoTo Exit_OptPayment_Click
   End If
End If

Exit_OptPayment_Click:
Exit Sub

Err_OptPayment_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Validating Payment Option selection"
Resume Exit_OptPayment_Click

End Sub

Private Sub SSDBCurrency_Click()
cmdConvert.Enabled = True
End Sub

Private Sub tb_Validate(Index As Integer, Cancel As Boolean)
Dim sYear As String, sMonth As String
Select Case Index
Case 1 ' accounting period
  If IsNothing(tb(1)) Then
    iErr = "Accounting Period Missing"
    Cancel = True
    GoTo Validate_Err
  End If
Validate_Exit:
  Exit Sub
Validate_Err:
  MsgBox iErr, vbOKOnly, "Replace Cheque Data Entry"
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
Function ValidateAccount()
On Error GoTo Err_ValidateAccount
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_IsAFinAccountPresent", 1, iClientID)
If adoRst.EOF Then
   MsgBox "No Financial Institution Account was found"
   GoTo Exit_ValidateAccount
End If

Exit_ValidateAccount:
Exit Function

Err_ValidateAccount:
MsgBox Err & " " & Err.Description, vbOKOnly, "Validation function failed"
Resume Exit_ValidateAccount

End Function
