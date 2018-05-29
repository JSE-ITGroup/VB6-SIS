VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmSelectAccount 
   Caption         =   "Select Account"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3090
   Icon            =   "FrmSelectAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBAccount 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
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
      RowHeight       =   503
      Columns.Count   =   2
      Columns(0).Width=   5741
      Columns(0).Caption=   "Account Number"
      Columns(0).Name =   "Account Number"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2275
      Columns(1).Caption=   "Currency"
      Columns(1).Name =   "Currency"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Column 0"
   End
End
Attribute VB_Name = "FrmSelectAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
On Error GoTo Err_CmdExit_Click
Unload Me
SpCon.Close
Exit_CmdExit_Click:
Exit Sub

Err_CmdExit_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on closing screen"
Resume Exit_CmdExit_Click

End Sub

Private Sub CmdStart_Click()
On Error GoTo Err_CmdStart_Click
Dim OptTitle As String
Dim i As Integer

If SSDBAccount = "" Then
   MsgBox "Please select an account first"
   GoTo Exit_CmdStart_Click
End If

If gblOptions = 1 Then
   i = RunSP(SpCon, "usp_PostBulkReturns", 0, SSDBAccount.Columns(0).Text)
   If i = 0 Then
      MsgBox "Returns successfully posted"
      GoTo Exit_CmdStart_Click
   End If
End If

If gblOptions = 2 Then
   i = RunSP(SpCon, "usp_PostBulkReplacements", 0, SSDBAccount.Columns(0).Text)
   If i = 0 Then
      MsgBox "Replacements successfully posted"
      GoTo Exit_CmdStart_Click
   End If
End If
If gblOptions = 4 Then
   gblFileKey = SSDBAccount.Columns(0).Text
   FrmReportView.Show 0
   GoTo Exit_CmdStart_Click
End If

Exit_CmdStart_Click:
Exit Sub

Err_CmdStart_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, OptTitle
Resume Exit_CmdStart_Click

End Sub

Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
If gblOptions = 1 Then
   OptTitle = "Post Bulk Returns"
End If
If gblOptions = 2 Then
   OptTitle = "Post Bulk Replacements"
End If
If gblOptions = 4 Then
   OptTitle = "Dividend Reconciliation Report"
End If
Exit_Form_Activate:
Exit Sub

Err_Form_Activate:
MsgBox Err & " " & Err.Description, vbOKOnly, "Form Activate Error"
Resume Exit_Form_Activate
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load
csvCenterForm Me, gblMDIFORM
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
frmMDI.txtStatusMsg.Refresh

Exit_Form_Load:
Exit Sub

Err_Form_Load:
MsgBox Err & " " & Err.Description, vbOKOnly, "Form Load Error"
Resume Exit_Form_Load

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmSelectAccount = Nothing
End Sub

Private Sub SSDBAccount_InitColumnProps()
On Error GoTo Err_SSDBAccount_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset
Dim i As Integer

Set adoRst = RunSP(SpCon, "usp_ListAllActiveAccounts", 1)
If adoRst.EOF Then
   MsgBox "Accounts were not setup" & vbCrLf & "Please do so now", vbCritical, "Account Error"
   GoTo Exit_SSDBAccount_InitColumnProps
End If

'adoRst.MoveFirst
With SSDBAccount
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst!AccountNo & vbTab & adoRst!Currency & vbTab
     .AddItem StrSql
     'If adoRst!CurrencyType = "L" Then
     '   i = .Rows
     'End If
     adoRst.MoveNext
     Loop
     '.Bookmark = .GetBookmark(i - 1)
     ' SSDBAccount = .Columns(0).CellText(i - 1)
End With

adoRst.Close
Set adoRst = Nothing
Exit_SSDBAccount_InitColumnProps:
Exit Sub

Err_SSDBAccount_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing active accounts"
Resume Exit_SSDBAccount_InitColumnProps
End Sub

