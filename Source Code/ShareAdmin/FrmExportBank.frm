VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmExportBank 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Export Selected Bank Payments"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   Icon            =   "FrmExportBank.frx":0000
   ScaleHeight     =   2100
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3480
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton CmdExport 
      Caption         =   "Export"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   3495
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
      Columns(0).Width=   3836
      Columns(0).Caption=   "Branch Name"
      Columns(0).Name =   "Branch Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1535
      Columns(1).Caption=   "Bank Id"
      Columns(1).Name =   "Bank Id"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   6165
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Bank Id:"
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
      Index           =   12
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1500
   End
End
Attribute VB_Name = "FrmExportBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdExport_Click()
Dim rsMain As ADODB.Recordset

If dbc(0) = vbNullString Then
   MsgBox "Please select a Bank", vbOKOnly, "Bank not Selected"
   dbc(0).SetFocus
   GoTo Exit_CmdExport_Click
End If

Set rsMain = RunSP(SpCon, "usp_BankPaymentsExport", 1, dbc(0).Columns(1).Text)

Call ExportToExcel(rsMain)
rsMain.Close
Set rsMain = Nothing

Exit_CmdExport_Click:
Exit Sub

Err_CmdExport_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error On Bank Payment Export"
GoTo Exit_CmdExport_Click

End Sub

Private Sub dbc_InitColumnProps(Index As Integer)
On Error GoTo dbc_InitColumnProps_Err
Dim sRowinfo As String
Dim rsBank As ADODB.Recordset

Select Case Index
Case 0
  dbc(0).RemoveAll
  Set rsBank = RunSP(SpCon, "usp_Banks", 1)
  If rsBank.EOF Then
     GoTo dbc_InitColumnProps_Err
  End If

  With rsBank
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        sRowinfo = !BnkName & vbTab & !BankID
        dbc(0).AddItem sRowinfo
       .MoveNext
      Loop
    End If
  End With
  '--
End Select

rsBank.Close

Exit_dbc_InitColumnProps:
Set rsBank = Nothing
Exit Sub

dbc_InitColumnProps_Err:
  MsgBox "Bank List not found", vbOKOnly, "Export Bank payments"
  GoTo Exit_dbc_InitColumnProps
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
   'frmMDI.txtStatusMsg.Refresh
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

End Sub
