VERSION 5.00
Begin VB.Form FrmCreateBnkPrintFile 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Bank Print File"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FrmCreateBnkPrintFile.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkSkip 
      BackColor       =   &H0080C0FF&
      Caption         =   "Skip first cheque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Tick to allow the print to skip te first cheque in the sequence"
      Top             =   1440
      Width           =   1935
   End
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
      TabIndex        =   4
      Top             =   1920
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
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox TxtChqNum 
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   2415
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
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Starting Chq No:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "FrmCreateBnkPrintFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim OriginalNo As Long
Dim iExchABBR As String

Private Sub ChkSkip_Click()
On Error GoTo Err_ChkSkip_Click

If ChkSkip = 1 Then
   TxtChqNum = OriginalNo + 1
Else
   TxtChqNum = OriginalNo
End If
Exit_ChkSkip_Click:
Exit Sub

Err_ChkSkip_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on skipping next account number"
Resume Exit_ChkSkip_Click
End Sub

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
Dim i As Integer
Dim tSkip As Boolean

If OriginalNo = 0 Then
   MsgBox "Unable to process as no cheque number was found"
   GoTo Exit_CmdStart_Click
End If

If ChkSkip = 1 Then
   tSkip = True
Else
   tSkip = False
End If

i = RunSP(SpCon, "usp_MakeBnkChq", 0, CLng(TxtChqNum), iExchABBR, tSkip)
If i <> 0 Then
   MsgBox "Error on generating and numbering Bank payments. Please try again or advise your Sys Ad"
   GoTo Exit_CmdStart_Click
Else
   MsgBox "Bank Payments have been generated and numbered. You may now print cheques"
End If

Exit_CmdStart_Click:
Exit Sub

Err_CmdStart_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on retrieving current ledger"
Resume Exit_CmdStart_Click
End Sub

Private Sub Form_Activate()
'On Error GoTo Err_Form_Activate
Dim adoRst As ADODB.Recordset
Dim StrSql As String
Dim iAccountNo As String

Set adoRst = RunSP(SpCon, "usp_SelectLedger", 1)
If adoRst.EOF Then
   MsgBox "No ledger found. Unable to proceed"
   GoTo Exit_Form_Activate
Else
   TxtRegister = adoRst!StockExchange
   iExchABBR = adoRst!ExchangeABBR
End If
OriginalNo = 0
Set adoRst = RunSP(SpCon, "usp_SelectLocalAccount", 1)
If adoRst.EOF Then
   StrSql = "Currencies not setup"
   MsgBox StrSql
End If
iAccountNo = adoRst!AccountNo

Set adoRst = RunSP(SpCon, "usp_NextAvailableChqNo", 1, iAccountNo, iExchABBR, "B")
If IsNull(adoRst!StartingNo) Or adoRst!StartingNo = "No Chqs found" Then
   MsgBox "Divdend inventory is empty. Unable to proceed"
Else
   If adoRst!StartingNo = "Already exists" Then
      StrSql = "Cheque numbers already assigned for " & iAccountNo & vbCrLf
      StrSql = StrSql & "in " & TxtRegister
      MsgBox StrSql
   Else
      OriginalNo = adoRst!StartingNo
      TxtChqNum = OriginalNo
   End If
End If

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

