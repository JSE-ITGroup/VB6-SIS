VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmSelectBatch 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Select certificate register to view"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4290
   Icon            =   "FrmSelectBatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBBatches 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
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
      Columns(0).Width=   3200
      Columns(0).Caption=   "Batch Date"
      Columns(0).Name =   "Batch Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Batch Number"
      Columns(1).Name =   "Batch Number"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   7011
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
End
Attribute VB_Name = "FrmSelectBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
gblFileKey = ""
Unload Me
End Sub


Private Sub CmdView_Click()
On Error GoTo Exit_CmdView_Click
Dim adoRst As ADODB.Recordset
Dim StrSql As String
Dim i As Integer

If SSDBBatches = vbNullString Then
   MsgBox "Please select a batch to re-open first"
   SSDBBatches.SetFocus
   GoTo Exit_CmdView_Click
End If
gblFileKey = SSDBBatches.Columns(1).Text
Unload Me
Exit_CmdView_Click:
Exit Sub
Err_CmdView_Click:
MsgBox Err.Description, vbOKOnly, "Error on attempt to setup view conditions"
Resume Exit_CmdView_Click

End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load
csvCenterForm Me, gblMDIFORM
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
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
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

Exit_Form_Load:
Exit Sub
Err_Form_Load:
MsgBox Err.Description, vbOKOnly, "Closed Certificate view open error"
GoTo Exit_Form_Load

End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
Set FrmSelectBatch = Nothing
End Sub

Private Sub SSDBBatches_InitColumnProps()
Dim adoBatches As ADODB.Recordset
Dim StrSql As String

Set adoBatches = RunSP(SpCon, "usp_Selectbatchlist", 1)

Do While Not adoBatches.EOF
   With SSDBBatches
        StrSql = Format(adoBatches!TrnDate, "dd-mmm-yyyy") & vbTab & adoBatches!TrnBatch
        .AddItem StrSql
   End With
   adoBatches.MoveNext
Loop

adoBatches.Close
Set adoBatches = Nothing

End Sub
