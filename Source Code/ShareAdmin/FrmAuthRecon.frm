VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmAuthRecon 
   Caption         =   "Authorise Bank Items View"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "FrmAuthRecon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_B.SSDBGrid SSDBPostings 
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   9615
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   8
      RowHeight       =   423
      Columns.Count   =   8
      Columns(0).Width=   2064
      Columns(0).Caption=   "Trans Date"
      Columns(0).Name =   "Trans Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).NumberFormat=   "dd-mmm-yyyy"
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2275
      Columns(2).Caption=   "Chq No"
      Columns(2).Name =   "Chq No"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Amount"
      Columns(3).Name =   "Narration"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1138
      Columns(4).Caption=   "DBCR"
      Columns(4).Name =   "DBCR"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   256
      Columns(5).Width=   2223
      Columns(5).Caption=   "Recon Date"
      Columns(5).Name =   "Recon Date"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   7
      Columns(5).NumberFormat=   "dd-mmm-yyyy"
      Columns(5).FieldLen=   256
      Columns(6).Width=   1746
      Columns(6).Caption=   "Authorise"
      Columns(6).Name =   "Authorise"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   11
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      Columns(7).Width=   1482
      Columns(7).Caption=   "ItemID"
      Columns(7).Name =   "ItemID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      _ExtentX        =   16960
      _ExtentY        =   8070
      _StockProps     =   79
      Caption         =   "Reconciling Items Awaiting Authorisation"
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
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton CmdAuth 
      Caption         =   "&Authorise"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton CmdSelectAll 
      Caption         =   "&Select All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5760
      Width           =   1695
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBAccount 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   5535
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
      Columns(0).Width=   2646
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
      _ExtentX        =   9763
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
   Begin VB.Label Label2 
      Caption         =   "Select Account:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15
   End
End
Attribute VB_Name = "FrmAuthRecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoRst As ADODB.Recordset
Dim SpCon As ADODB.Connection

Private Sub CmdAuth_Click()
On Error GoTo Err_CmdAuth_Click
Dim StrSql As String

With SSDBPostings
     .MoveFirst
     adoRst.MoveFirst
     'For i = 1 To .Rows
      '   If .Columns(6).Value = True Then
     'Next i
End With
CmdSelectAll.Enabled = False
CmdAuth.Enabled = False

Exit_CmdAuth_Click:
Exit Sub
Err_CmdAuth_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on Authorisation Command"
Resume Exit_CmdAuth_Click

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdSelectAll_Click()
On Error GoTo Err_CmdSelectAll_Click
Dim i As Integer
With SSDBPostings
     .MoveFirst
     For i = 1 To .Rows
     .Columns(6).Value = True
     .MoveNext
     Next i
End With

SSDBPostings.Refresh
Exit_CmdSelectAll_Click:
Exit Sub
Err_CmdSelectAll_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on Select All Command"
Resume Exit_CmdSelectAll_Click
End Sub

Private Sub Form_Load()
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
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
frmMDI.txtStatusMsg.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmAuthRecon = Nothing
If adoRst.State <> 0 Then
   adoRst.Close
End If
Set adoRst = Nothing
SpCon.Close
End Sub

Private Sub SSDBAccount_Click()
On Error GoTo Err_SSDBAccount_Click
Dim StrSql As String
Dim i As Integer

If IsEmpty(SSDBAccount.SelBookmarks(0)) Then
   Beep
   MsgBox "Select an account "
   SSDBAccount.SetFocus
   GoTo Exit_SSDBAccount_Click
End If


Set adoRst = RunSP(SpCon, "usp_SelectBankItems", 1, 1, CDbl(SSDBAccount.Columns(0).Text))
If adoRst.EOF Then
   MsgBox "Sorry, no records requiring authorisation were found"
   GoTo Exit_SSDBAccount_Click
End If
With SSDBPostings
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = ""
     For i = 0 To 5
         StrSql = StrSql & adoRst(i) & vbTab
         Next i
     StrSql = StrSql & 0 & vbTab & adoRst(6)
     .AddItem StrSql
     adoRst.MoveNext
Loop
End With
CmdSelectAll.Enabled = True
CmdAuth.Enabled = True

Exit_SSDBAccount_Click:
Exit Sub
Err_SSDBAccount_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Authorisation Error"
Resume Exit_SSDBAccount_Click

End Sub

Private Sub SSDBAccount_InitColumnProps()
On Error GoTo Err_SSDBAccount_InitColumnProps
Dim StrSql As String

Set adoRst = RunSP(SpCon, "usp_SelectAccounts", 1, 1)
If adoRst.EOF Then
   MsgBox "Accounts were not setup" & vbCrLf & "Please do so now", vbCritical, "Account Error"
   GoTo Exit_SSDBAccount_InitColumnProps
End If

adoRst.MoveFirst
With SSDBAccount
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst(0) & vbTab & adoRst(1) & vbTab
     .AddItem StrSql
     adoRst.MoveNext
     StrSql = ""
     Loop
End With
adoRst.Close
Set adoRst = Nothing
Exit_SSDBAccount_InitColumnProps:
Exit Sub

Err_SSDBAccount_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "SSDB Combo Box load Error"
Resume Exit_SSDBAccount_InitColumnProps

End Sub

