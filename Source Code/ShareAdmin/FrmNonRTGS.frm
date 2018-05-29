VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmNonRTGS 
   Caption         =   "Non RTGS with amounts over threshold"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   Icon            =   "FrmNonRTGS.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
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
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBNonRTGS 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   3
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Caption=   "ClientID"
      Columns(0).Name =   "ClientID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5689
      Columns(1).Caption=   "Client Name"
      Columns(1).Name =   "Client Name"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3572
      Columns(2).Caption=   "Payment Amount"
      Columns(2).Name =   "Payment Amount"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   13335
      _ExtentY        =   7858
      _StockProps     =   79
      BackColor       =   12640511
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
End
Attribute VB_Name = "FrmNonRTGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
On Error GoTo Err_CmdExit_Click

Unload Me

Exit_cmdExit_Click:
Exit Sub

Err_CmdExit_Click:
MsgBox Err.Description, vbOKOnly, "Returned Cheques Exit"
GoTo Exit_cmdExit_Click
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
Loop
Screen.MousePointer = vbDefault
  
   '-------------------------------------
   '-- Initialize License Details -------
   '-------------------------------------
   '--
 '--
Exit_Form_Load::
Exit Sub

Err_Form_Load:
MsgBox Err.Description, vbOKOnly, "RTGS threshold Form Load error"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set FrmNonRTGS = Nothing

End Sub

Private Sub SSDBNonRTGS_DblClick()
gblFileKey = SSDBNonRTGS.Columns(0).Text
FrmBankRTGSDetails.Show 0
End Sub

Private Sub SSDBNonRTGS_InitColumnProps()
On Error GoTo Err_SSDBNonRTGS_InitColumnProps
Dim adoRst As ADODB.Recordset
Dim StrSql As String

Set adoRst = RunSP(SpCon, "usp_GetNonRTGS", 1)
If adoRst.EOF Then
   'MsgBox "There are no payments above the RTGS threshold but not setup for RTGS"
   GoTo Exit_SSDBNonRTGS_InitColumnProps
End If
SSDBNonRTGS.RemoveAll

Do While Not adoRst.EOF
   With SSDBNonRTGS
        StrSql = adoRst!ClientID & vbTab & adoRst!PayeeName & vbTab & Format(adoRst!PayAmt, "#,###.00")
        .AddItem StrSql
    adoRst.MoveNext
   End With
Loop

Exit_SSDBNonRTGS_InitColumnProps:
Exit Sub

Err_SSDBNonRTGS_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on retrieving RTGS threshold"
Resume Exit_SSDBNonRTGS_InitColumnProps
End Sub
