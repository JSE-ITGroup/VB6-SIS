VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmAuthRepChq 
   Caption         =   "Authorise Replaced Cheques View"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   Icon            =   "FrmAuthRepChq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_B.SSDBGrid SSDBPostings 
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10455
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   9
      BackColorEven   =   12648447
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   9
      Columns(0).Width=   3200
      Columns(0).Caption=   "Client Name"
      Columns(0).Name =   "Client Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2275
      Columns(1).Caption=   "Old Chq No"
      Columns(1).Name =   "Old Chq No"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2064
      Columns(2).Caption=   "Old Chq Amt"
      Columns(2).Name =   "Old Chq Amt"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   6
      Columns(2).NumberFormat=   "CURRENCY"
      Columns(2).FieldLen=   256
      Columns(3).Width=   2699
      Columns(3).Caption=   "New Chq No"
      Columns(3).Name =   "New Chq No"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1984
      Columns(4).Caption=   "New Chq Amt"
      Columns(4).Name =   "New Chq Amt"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   256
      Columns(5).Width=   2223
      Columns(5).Caption=   "Replacement Date"
      Columns(5).Name =   "Replacement Date"
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
      Columns(7).Caption=   "User ID"
      Columns(7).Name =   "User ID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1270
      Columns(8).Caption=   "TXID"
      Columns(8).Name =   "TXID"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      _ExtentX        =   18441
      _ExtentY        =   9128
      _StockProps     =   79
      Caption         =   "Replaced Cheques Requiring Authorisation"
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
      TabIndex        =   2
      Top             =   5400
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
      TabIndex        =   1
      Top             =   5400
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
      TabIndex        =   0
      Top             =   5400
      Width           =   1695
   End
End
Attribute VB_Name = "FrmAuthRepChq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoRst As ADODB.Recordset
Dim SpCon As ADODB.Connection

Private Sub CmdAuth_Click()
On Error GoTo Err_CmdAuth_Click
Dim StrSql As String
Dim X As Integer
Dim iLines As Integer

StrSql = ""
iLines = 0
With SSDBPostings
     .MoveFirst
     For i = 1 To .Rows
         If .Columns(6).Value = True Then
            StrSql = StrSql & .Columns(8).Value & ";"
            iLines = iLines + 1
         End If
     .MoveNext
     Next i
End With
X = RunSP(SpCon, "usp_RepChqUpdate", 0, StrSql, iLines, gblLoginName)
If X = 0 Then
   CmdSelectAll.Enabled = False
   CmdAuth.Enabled = False
   MsgBox "Authorisation was successful"
Else
   GoTo Err_CmdAuth_Click
End If

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

Private Sub SSDBPostings_InitColumnProps()
On Error GoTo Err_SSDBPostings_InitColumnProps
Dim StrSql As String
Dim i As Integer

Set adoRst = RunSP(SpCon, "usp_SelectRepChq", 1, 0)
If adoRst.EOF Then
   MsgBox "Sorry, no records requiring authorisation were found"
   GoTo Exit_SSDBPostings_InitColumnProps
End If
With SSDBPostings
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = ""
     For i = 0 To 5
         StrSql = StrSql & adoRst(i) & vbTab
         Next i
     StrSql = StrSql & 0 & vbTab & adoRst(6) & vbTab & adoRst(7)
     .AddItem StrSql
     adoRst.MoveNext
Loop
End With
CmdSelectAll.Enabled = True
CmdAuth.Enabled = True

Exit_SSDBPostings_InitColumnProps:
Exit Sub

Err_SSDBPostings_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Authorisation Error"
Resume Exit_SSDBPostings_InitColumnProps

End Sub
