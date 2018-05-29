VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmCurrencyRate 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Set Currency Exchange Rate"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FrmCurrencyRate.frx":0000
   ScaleHeight     =   3795
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FmePaymentDetails 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Current Payment Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   4215
      Begin VB.TextBox TxtChqDate 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxtDividendType 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Payment Date:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Payment Type:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox TxtExchangeRate 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBCurrency 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   2175
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
      Columns.Count   =   2
      Columns(0).Width=   1508
      Columns(0).Caption=   "Currency"
      Columns(0).Name =   "Currency"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Account No"
      Columns(1).Name =   "Account No"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Exchange Rate:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Currency"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "FrmCurrencyRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
On Error GoTo Err_CmdSave_Click
Dim StrSql As String
Dim i As Integer
Dim iPayTyp As String

If SSDBCurrency = "" Then
   StrSql = "You are required to select a currency before click SAVE" & vbCrLf
   StrSql = StrSql & "Please do so before continuing"
   MsgBox StrSql, vbOKOnly, "No Currency Selected"
   SSDBCurrency.SetFocus
   GoTo Exit_CmdSave_Click
End If

If Not IsNumber(TxtExchangeRate) Then
   StrSql = "You are required to enter an exchange rate before clicking SAVE" & vbCrLf
   StrSql = StrSql & "Please correct"
   MsgBox StrSql, vbOKOnly, "Exchange Rate Missing"
   GoTo Exit_CmdSave_Click
End If
If CCur(TxtExchangeRate) < 1 Then
   StrSql = "This is not a valid amount" & vbCrLf
   StrSql = StrSql & "Please correct and then click SAVE again"
   MsgBox StrSql, vbOKOnly, "Invalid Rate"
   GoTo Exit_CmdSave_Click
End If

If TxtDividendType = "Dividend" Then
   iPayTyp = "D"
Else
   iPayTyp = "C"
End If
i = RunSP(SpCon, "usp_UpdateExchangeRate", 0, SSDBCurrency.Columns(0).Text, iPayTyp, TxtChqDate, CCur(TxtExchangeRate))
If i <> 0 Then
   StrSql = "An error occurred while updating the rate." & vbCrLf
   StrSql = StrSql & "Please contact your Systems Administrator"
   MsgBox StrSql, vbOKOnly, "Error on Currency Rate update"
   GoTo Exit_CmdSave_Click
Else
   MsgBox "Exchange Rate successfully saved", vbOKOnly, "Save Completed"
End If

Exit_CmdSave_Click:
Exit Sub

Err_CmdSave_Click:
MsgBox Err.Description, vbOKOnly, "Error trying to save"
GoTo Exit_CmdSave_Click

End Sub

Private Sub Form_Activate()
Dim adoRst As New ADODB.Recordset
Dim i As Integer
Dim bm As Variant

Set adoRst = RunSP(SpCon, "usp_ExchangeRateDetail", 1, gblFileKey)
With adoRst
     If !PayType = 0 Then
         MsgBox "There are no active dividends. This option will terminate"
      GoTo Exit_Form_Activate
     End If
     TxtDividendType = !PayType
     TxtChqDate = Format(!ChqDate, "dd-mmm-yyyy")
End With
If gblFileKey = "0" Then
   TxtExchangeRate = 0
   SSDBCurrency.Enabled = True
Else
   TxtExchangeRate = adoRst!ExchRate
   With SSDBCurrency
     .MoveFirst
     For i = 0 To .Rows - 1
         bm = .GetBookmark(i)
         If .Columns(0).CellText(bm) = gblFileKey Then
            .Bookmark = .GetBookmark(i)
             SSDBCurrency = .Columns(0).CellText(bm)
         Exit For
         End If
     Next i
   End With
   SSDBCurrency.Enabled = False
End If


adoRst.Close
Set adoRst = Nothing

Exit_Form_Activate:
Exit Sub

End Sub

Private Sub Form_Load()
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

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmCurrencyRate = Nothing
SpCon.Close
End Sub

Private Sub SSDBCurrency_InitColumnProps()
On Error GoTo Err_SSDBCurrency_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_ListForeignCurrencies", 1)
If adoRst.EOF Then
   MsgBox "No foreign currency accounts were not setup" & vbCrLf & "Please set them up now", vbCritical, "List Foreign Currencies error"
   GoTo Exit_SSDBCurrency_InitColumnProps
End If

'adoRst.MoveFirst
With SSDBCurrency
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
Exit_SSDBCurrency_InitColumnProps:
Exit Sub

Err_SSDBCurrency_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error loading currencies"
Resume Exit_SSDBCurrency_InitColumnProps

End Sub
