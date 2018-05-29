VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmStockExchange 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Stock Exchange Maintenance"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   Icon            =   "FrmStockExchange.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkAdd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add New Exchange"
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
      Left            =   3600
      TabIndex        =   10
      Top             =   840
      Width           =   2415
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBExchanges 
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   7095
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   5715
      Columns(0).Caption=   "Stock Exchange"
      Columns(0).Name =   "Stock Exchange"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Abbreviation"
      Columns(1).Name =   "Abbreviation"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Currencies"
      Columns(2).Name =   "Currencies"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "Exchange ID"
      Columns(3).Name =   "Exchange ID"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   12515
      _ExtentY        =   4048
      _StockProps     =   79
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
   Begin VB.TextBox TxtExchID 
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin SSDataWidgets_A.SSDBCommand CmdSave 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
      _Version        =   196612
      _ExtentX        =   2778
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12648384
      BevelColorHighlight=   16761087
   End
   Begin VB.Frame FmeCurrency 
      Caption         =   "Payments made in the following currencies (ticked)"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   7095
      Begin VB.CheckBox ChkCurrency 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtABBR 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox TxtStockExchange 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
   Begin SSDataWidgets_A.SSDBCommand CmdExit 
      Height          =   735
      Left            =   5520
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
      _Version        =   196612
      _ExtentX        =   2778
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   16761024
      BevelColorHighlight=   12582912
   End
   Begin SSDataWidgets_A.SSDBCommand CmdDelete 
      Height          =   735
      Left            =   3000
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
      _Version        =   196612
      _ExtentX        =   2778
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   255
      BevelColorHighlight=   16761087
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Abbreviation:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stock Exchange:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FrmStockExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub ChkAdd_Click()
Dim i As Integer

If ChkAdd.Value = 1 Then
   TxtABBR = ""
   TxtStockExchange = ""
   TxtExchID = ""
   For i = 0 To ChkCurrency.Count - 1
       ChkCurrency(i).Value = 0
   Next i
Else
   SSDBExchanges_Click
End If

   
End Sub

Private Sub CmdDelete_Click()
On Error GoTo Err_CmdDelete_Click
Dim StrSql As String
Dim i As Integer

If Len(TxtExchID) < 1 Then
   MsgBox "Please select a exchange from the list below to delete"
   GoTo Exit_CmdDelete_Click
End If
If TxtExchID = "0" Then
   MsgBox "You are not allowed to delete the Main register!", vbOKOnly
   GoTo Exit_CmdDelete_Click
End If

StrSql = "You are about to Delete " & TxtStockExchange & "!" & vbCrLf
StrSql = StrSql & "Are you sure you want to do this?"
i = MsgBox(StrSql, vbYesNo)
If i = vbYes Then
   i = RunSP(SpCon, "usp_DeleteStockExchange", 0, CInt(TxtExchID))
   If i = 1 Then
      StrSql = "Unable to delete!" & vbCrLf
      StrSql = StrSql & "There are currently shareholders attached to this stock exchange"
      MsgBox StrSql, vbOKOnly
      GoTo Exit_CmdDelete_Click
   Else
      If i = 0 Then
         MsgBox "Stock Exchange deleted"
         GoTo Exit_CmdDelete_Click
      Else
         GoTo Err_CmdDelete_Click
      End If
   End If
Else
   MsgBox "Deletion abandoned"
End If
      
Exit_CmdDelete_Click:
Exit Sub

Err_CmdDelete_Click:
MsgBox Err.Description, vbOKOnly, "Error on deleting stock exchange"
GoTo Exit_CmdDelete_Click
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
On Error GoTo Err_CmdSave_Click
Dim CurrencyList As String
Dim i As Integer
Dim iExchID As String

If Len(TxtStockExchange) < 10 Then
   MsgBox "Please enter a name for this stock exchange"
   TxtStockExchange.SetFocus
   GoTo Exit_CmdSave_Click
End If
If Len(TxtABBR) < 4 Then
   MsgBox "Please enter an abbreviation for this stock exchange"
   TxtABBR.SetFocus
   GoTo Exit_CmdSave_Click
End If
CurrencyList = ""
For i = 0 To ChkCurrency.Count - 1
    If ChkCurrency(i).Value = 1 Then
       CurrencyList = CurrencyList & ChkCurrency(i).Caption & ";"
    End If
Next i
If Len(CurrencyList) < 4 Then
   MsgBox "At least one currency must be selected before saving"
   ChkCurrency(0).SetFocus
   GoTo Exit_CmdSave_Click
End If
If ChkAdd.Value = 1 Then
   iExchID = "A"
Else
   iExchID = TxtExchID
End If

i = RunSP(SpCon, "usp_UpdateStockExchange", 0, iExchID, TxtStockExchange, TxtABBR, CurrencyList, gblLoginName)
If i = 0 Then
   MsgBox "Stock Exchange update was succesful"
   GoTo Exit_CmdSave_Click
Else
   MsgBox "Error on stock exchnage update"
End If


Exit_CmdSave_Click:
Exit Sub

Err_CmdSave_Click:
MsgBox Err.Description, vbOKOnly, "Error on saving changes"
GoTo Exit_CmdSave_Click
End Sub

Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
Dim adoRst As ADODB.Recordset
Dim i As Integer

Set adoRst = RunSP(SpCon, "usp_SelectAvailableCurrencies", 1)
i = 1
With adoRst
     Do While Not .EOF
        If !CurrencyType = "L" Then
           ChkCurrency(0).Caption = !Currency
        Else
           Load ChkCurrency(i)
           ChkCurrency(i).Caption = !Currency
           ChkCurrency(i).Visible = True
           ChkCurrency(i).Top = ChkCurrency(i - 1).Top
           ChkCurrency(i).Left = ChkCurrency(i - 1).Left + ChkCurrency(i - 1).Width + 525
           i = i + 1
           If i > 4 Then
              MsgBox "No more currencies can be loaded"
              GoTo Exit_Form_Activate
           End If
        End If
        .MoveNext
     Loop
End With

adoRst.Close
Set adoRst = Nothing

Exit_Form_Activate:
Exit Sub

Err_Form_Activate:
MsgBox Err.Description, vbOKOnly, "Error on listing currencies"
GoTo Exit_Form_Activate
End Sub

Private Sub SSDBExchanges_Click()
On Error GoTo Err_SSDBExchanges_Click
Dim pos As Integer
Dim i As Integer

With SSDBExchanges
     TxtExchID = .Columns(3).Text
     TxtStockExchange = .Columns(0).Text
     TxtABBR = .Columns(1).Text
For i = 0 To ChkCurrency.Count - 1
    pos = InStr(1, .Columns(2).Text, ChkCurrency(i).Caption)
    If pos > 0 Then
       ChkCurrency(i).Value = 1
    Else
       ChkCurrency(i).Value = 0
    End If
Next i
End With

Exit_SSDBExchanges_Click:
Exit Sub

Err_SSDBExchanges_Click:
MsgBox Err.Description, vbOKOnly, "Error on displaying stock exchange details"
GoTo Exit_SSDBExchanges_Click
End Sub

Private Sub SSDBExchanges_InitColumnProps()
On Error GoTo Err_SSDBExchanges_InitColumnProps
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_SelectAvailableExchanges", 1)
If adoRst.EOF Then
   MsgBox "No exchanges have been setup. This must be corrected", vbOKOnly
   GoTo Exit_SSDBExchanges_InitColumnProps
End If
With SSDBExchanges
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst!StockExchange & vbTab & adoRst!ExchangeABBR & vbTab & adoRst!Currencies & vbTab & adoRst!StockExchangeID
     .AddItem StrSql
     adoRst.MoveNext
     Loop
End With
Exit_SSDBExchanges_InitColumnProps:
Exit Sub

Err_SSDBExchanges_InitColumnProps:
MsgBox Err.Description, vbOKOnly, "Error on setting up stock exchanges"
GoTo Exit_SSDBExchanges_InitColumnProps
End Sub
Private Sub Form_Load()
csvCenterForm Me, gblMDIFORM
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
