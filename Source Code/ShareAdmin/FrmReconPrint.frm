VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReconPrint 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reconciliation Report Generator"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmReconPrint.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7080
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   5160
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.OptionButton OptSelection 
      BackColor       =   &H0080FFFF&
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton OptSelection 
      BackColor       =   &H0080FFFF&
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   51511299
      CurrentDate     =   38581
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBAccounts 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   3855
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
      _ExtentX        =   6800
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
      BackColor       =   &H0080FFFF&
      Caption         =   "Reconciliation Date:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Select an account"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FrmReconPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim repSISRept As New SISRepts

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdStart_Click()

If OptSelection(0).Value = True Then
   gblOptions = 0
Else
   gblOptions = 3
End If
gblDate = DTPicker1
gblDate1 = DTPicker1
gblFileKey = SSDBAccounts.Columns(0).Text
FrmReportView.Show 0
End Sub

Private Sub Form_Load()
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
DTPicker1 = Date
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
frmMDI.txtStatusMsg.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub SSDBAccounts_InitColumnProps()
On Error GoTo Err_SSDBAccounts_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset
Set adoRst = RunSP(SpCon, "usp_SelectAccounts", 1, 1)
If adoRst.EOF Then
   MsgBox "Accounts were not setup" & vbCrLf & "Please do so now", vbCritical, "Account Error"
   GoTo Exit_SSDBAccounts_InitColumnProps
End If

'adoRst.MoveFirst
With SSDBAccounts
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
Exit_SSDBAccounts_InitColumnProps:
Exit Sub

Err_SSDBAccounts_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "SSDB Combo Box load Error"
Resume Exit_SSDBAccounts_InitColumnProps

End Sub
