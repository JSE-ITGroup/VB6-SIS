VERSION 5.00
Begin VB.Form FrmPrintBnkCheques 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Bank Cheques"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FrmPrintBnkCheques.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmPrintBnkCheques.frx":030A
   ScaleHeight     =   2265
   ScaleWidth      =   4680
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
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtRegister 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmPrintBnkCheques"
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
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on closing create cheque print file screen"
Resume Exit_cmdExit_Click
End Sub

Private Sub CmdStart_Click()
On Error GoTo Err_CmdStart_Click
If SSDBAccount = "" Then
   MsgBox "Please select an account first"
   SSDBAccount.SetFocus
   GoTo Exit_CmdStart_Click
End If

gblFileKey = SSDBAccount.Columns(1).Text
CmdExit_Click

Exit_CmdStart_Click:
Exit Sub

Err_CmdStart_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on retrieving current ledger"
Resume Exit_CmdStart_Click
End Sub

Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_SelectLedger", 1)
If adoRst.EOF Then
   MsgBox "No ledger found. Unable to proceed"
   GoTo Exit_Form_Activate
Else
   TxtRegister = adoRst!StockExchange
End If
SSDBAccount.SetFocus

Exit_Form_Activate:
Exit Sub

Err_Form_Activate:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on retrieving current ledger"
Resume Exit_Form_Activate

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

