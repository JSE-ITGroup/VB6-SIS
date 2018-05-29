VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReportOption 
   Caption         =   "Report Options"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_B.SSDBCombo SSDBDiv 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Width           =   2535
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
      Columns(0).Caption=   "Account No"
      Columns(0).Name =   "Account No"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).NumberFormat=   "dd-mmm-yyyy"
      Columns(0).FieldLen=   256
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton CmdGo 
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   55377923
      CurrentDate     =   38811
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBAccounts 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2415
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
      _ExtentX        =   4260
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
      Caption         =   "Declaration Date:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Report Date:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FrmReportOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim repSISRept As New SISRepts

Private Sub CmdExit_Click()
On Error GoTo Err_CmdExit_Click

Unload Me
Exit_cmdExit_Click:
Exit Sub

Err_CmdExit_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Report Options - Exit Click"
GoTo Exit_cmdExit_Click

End Sub

Private Sub CmdGo_Click()
On Error GoTo Err_CmdGo_Click
If IsEmpty(SSDBAccounts.SelBookmarks(0)) Then
   Beep
   MsgBox "Select an account "
   SSDBAccounts.SetFocus
   GoTo Exit_CmdGo_Click
End If
Set repSISRept = New SISRepts
repSISRept.LoginId = gblFileName
repSISRept.ReportType = 9

repSISRept.ClientID = CLng(SSDBAccounts.Columns(0).Text)
'repSISRept.DSN = DTPicker1.Value
repSISRept.siteid = SSDBDiv.Columns(0).Text
repSISRept.ReportNumber = 85
repSISRept.RunShareHolderReport

  
Exit_CmdGo_Click:
Unload Me
Exit Sub

Err_CmdGo_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Report Options - Go Click"
GoTo Exit_CmdGo_Click

End Sub

Private Sub Form_Activate()
DTPicker1 = Date
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbArrowHourglass
Set SpCon = New ADODB.Connection
With SpCon
     .ConnectionString = gblFileName
     .CursorLocation = adUseClient
     '.Provider = "SQLOLEDB.1"
End With
SpCon.Open , , , adAsyncConnect
Do While SpCon.State = adStateConnecting
   Screen.MousePointer = vbHourglass
Loop
'''MsgBox "Connected"
Screen.MousePointer = vbDefault
gblReply = 0
gblDSN = vbNullString

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

adoRst.MoveFirst
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

Private Sub SSDBDiv_InitColumnProps()
On Error GoTo Err_SSDBDiv_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset
Set adoRst = RunSP(SpCon, "usp_DivRef", 1)
If adoRst.EOF Then
   MsgBox "Dividend Reference Table is missing or empty" & vbCrLf & "Please set this up now", vbCritical, "Dividend Error"
   GoTo Exit_SSDBDiv_InitColumnProps
End If

adoRst.MoveFirst
With SSDBDiv
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst(3) & vbTab
     .AddItem StrSql
     adoRst.MoveNext
     StrSql = ""
     Loop
End With
adoRst.Close
Set adoRst = Nothing
Exit_SSDBDiv_InitColumnProps:
Exit Sub

Err_SSDBDiv_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "SSDB Combo Box load Error"
Resume Exit_SSDBDiv_InitColumnProps

End Sub
