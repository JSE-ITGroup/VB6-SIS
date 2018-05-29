VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmBankRTGSDetails 
   BackColor       =   &H00808080&
   Caption         =   "List of Payments to be credited to accounts at "
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   Icon            =   "FrmBankRTGSDetails.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExport 
      Caption         =   "Export To Excel"
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
      Left            =   840
      TabIndex        =   2
      Top             =   4800
      Width           =   1935
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBRTGSList 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11535
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   5
      RowHeight       =   503
      ExtraHeight     =   132
      Columns.Count   =   5
      Columns(0).Width=   4551
      Columns(0).Caption=   "Shareholder"
      Columns(0).Name =   "Shareholder"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "ClientID"
      Columns(1).Name =   "ClientID"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3810
      Columns(2).Caption=   "Gross Payment"
      Columns(2).Name =   "Gross Payment"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Tax"
      Columns(3).Name =   "Tax"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   5636
      Columns(4).Caption=   "Financial Institution"
      Columns(4).Name =   "Financial Institution"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      _ExtentX        =   20346
      _ExtentY        =   8070
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   9360
      TabIndex        =   0
      Top             =   4800
      Width           =   1455
   End
End
Attribute VB_Name = "FrmBankRTGSDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoRst As ADODB.Recordset
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdExport_Click()
Call ExportToExcel(adoRst)
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
adoRst.Close
Set adoRst = Nothing
Set FrmBankRTGSDetails = Nothing
SpCon.Close
End Sub

Private Sub SSDBRTGSList_InitColumnProps()
On Error GoTo Err_SSDBRTGSList_InitColumnProps
Dim StrSql As String

Set adoRst = RunSP(SpCon, "usp_ShowRTGSAccounts", 1, gblFileKey)
If adoRst.EOF Then
   MsgBox "This is not an other bank account"
   GoTo Exit_SSDBRTGSList_InitColumnProps
End If

adoRst.MoveFirst
With SSDBRTGSList
     .RemoveAll
     Me.Caption = Me.Caption & " " & adoRst!BnkName
     Do While Not adoRst.EOF
        StrSql = adoRst!CliName & vbTab & adoRst!ClientID & vbTab & Format(adoRst!GrossPymnt, "#,##0.00") & vbTab
        StrSql = StrSql & Format(adoRst!WhldTax, "#,##0.00") & vbTab & adoRst!BnkName
        .AddItem StrSql
        adoRst.MoveNext
        StrSql = ""
     Loop
End With

Exit_SSDBRTGSList_InitColumnProps:
Exit Sub

Err_SSDBRTGSList_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on loading RTGS list of accounts"
Resume Exit_SSDBRTGSList_InitColumnProps
End Sub
