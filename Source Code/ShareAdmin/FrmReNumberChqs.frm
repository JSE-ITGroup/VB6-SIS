VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmReNumberChqs 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ReNumber Cheque ranges"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FrmReNumberChqs.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter New Cheque Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   13
      Top             =   3240
      Width           =   4575
      Begin VB.TextBox TxtNewStartNo 
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
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox TxtNewEndNo 
         Enabled         =   0   'False
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "New No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Ending No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Enter Range to be renumbered"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   4575
      Begin VB.TextBox TxtNoOfChqs 
         Enabled         =   0   'False
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox TxtOldEndNo 
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
         Left            =   1320
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox TxtOldStartNo 
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
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "No  of Chqs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Ending No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Starting No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
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
      TabIndex        =   6
      Top             =   5160
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
      TabIndex        =   5
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox TxtRegister 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
   Begin SSDataWidgets_B.SSDBCombo SSDBAccount 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   840
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
      Columns(0).Width=   5741
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
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Select the currency"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "FrmReNumberChqs"
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

Exit_CmdExit_Click:
Exit Sub

Err_CmdExit_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on closing create cheque print file screen"
Resume Exit_CmdExit_Click
End Sub

Private Sub CmdStart_Click()
On Error GoTo Err_CmdStart_Click
Dim i As Integer

If SSDBAccount = "" Then
   MsgBox "Please select a currency first"
   GoTo Exit_CmdStart_Click
End If
If Len(TxtOldStartNo) < 1 Then
   MsgBox "Please recheck the values entered in the Old Start number"
   TxtOldStartNo.SetFocus
   GoTo Exit_CmdStart_Click
End If
If Len(TxtOldEndNo) < 1 Then
   MsgBox "Please recheck the values entered in the Old End number"
   TxtOldEndNo.SetFocus
   GoTo Exit_CmdStart_Click
End If
If Len(TxtNewStartNo) < 1 Then
   MsgBox "Please recheck the values entered in the New Start number"
   TxtNewStartNo.SetFocus
   GoTo Exit_CmdStart_Click
End If

If TxtNoOfChqs = "0" Then
   MsgBox "Please recheck the values entered in the Old Start and Old End numbers"
   TxtOldEndNo.SetFocus
   GoTo Exit_CmdStart_Click
End If

If TxtNewEndNo = "0" Then
   MsgBox "Please recheck the value entered in the New Start number"
   TxtNewStartNo.SetFocus
   GoTo Exit_CmdStart_Click
End If

i = RunSP(SpCon, "usp_ReNumberChqs", 0, SSDBAccount.Columns(0).Text, CLng(TxtOldStartNo), CLng(TxtOldEndNo), CLng(TxtNoOfChqs), CLng(TxtNewStartNo), CLng(TxtNewEndNo), gblLoginName)
If i = 1 Then
   MsgBox "The Cheques to be renumbered were not found. Please verify the numbers and resubmit"
   GoTo Exit_CmdStart_Click
End If
If i = 2 Or i = 3 Then
   MsgBox "The New cheques numbers were not found in the dividend inventory. Please make the necessary transfers and try again."
   GoTo Exit_CmdStart_Click
End If

MsgBox "Cheque renumbering was successfull. You may now print cheques"

Exit_CmdStart_Click:
Exit Sub

Err_CmdStart_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on renumbering"
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
   iExchABBR = adoRst!ExchangeABBR
End If
SSDBAccount.SetFocus
OriginalNo = 0

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

Private Sub SSDBAccount_InitColumnProps()
On Error GoTo Err_SSDBAccount_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset
Dim i As Integer

Set adoRst = RunSP(SpCon, "usp_ListActiveAccounts", 1)
If adoRst.EOF Then
   MsgBox "Accounts were not setup" & vbCrLf & "Please do so now", vbCritical, "Account Error"
   GoTo Exit_SSDBAccount_InitColumnProps
End If

With SSDBAccount
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst!AccountNo & vbTab & adoRst!Currency & vbTab
     .AddItem StrSql
     adoRst.MoveNext
     StrSql = ""
     Loop
End With

adoRst.Close
Set adoRst = Nothing
Exit_SSDBAccount_InitColumnProps:
Exit Sub

Err_SSDBAccount_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing active accounts"
Resume Exit_SSDBAccount_InitColumnProps
End Sub

Private Sub TxtNewStartNo_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
   Else
   KeyAscii = 0
   GoTo Exit_TxtNewStartNo_KeyPress
End If

Exit_TxtNewStartNo_KeyPress:
Exit Sub

Err_TxtNewStartNo_KeyPress:
MsgBox Err.Number & " " & Err.Description, vbOKOnly, "New start number input error"
Resume Exit_TxtNewStartNo_KeyPress
End Sub

Private Sub TxtNewStartNo_LostFocus()
If Len(TxtOldStartNo) < 1 Then
   GoTo Exit_TxtNewStartNo_LostFocus
End If
If Len(TxtNewStartNo) < 1 Then
   MsgBox "New Starting number is required"
   GoTo Exit_TxtNewStartNo_LostFocus
End If
If Len(TxtNoOfChqs) < 1 Then
   MsgBox "Number of cheques missing. Please address"
   GoTo Exit_TxtNewStartNo_LostFocus
End If

If CLng(TxtNewStartNo) < CLng(TxtOldStartNo) Then
   MsgBox "The new number is less than the old number. Please correct"
   TxtNewEndNo = 0
   TxtNewStartNo.SetFocus
Else
   TxtNewEndNo = CLng(TxtNewStartNo) + CLng(TxtNoOfChqs) - 1
End If
Exit_TxtNewStartNo_LostFocus:
Exit Sub

End Sub

Private Sub TxtOldEndNo_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
   Else
   KeyAscii = 0
   GoTo Exit_TxtOldEndNo_KeyPress
End If

Exit_TxtOldEndNo_KeyPress:
Exit Sub

Err_TxtOldEndNo_KeyPress:
MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Old end number input error"
Resume Exit_TxtOldEndNo_KeyPress
End Sub

Private Sub TxtOldEndNo_LostFocus()
If Len(TxtOldStartNo) < 1 Then
   GoTo ExitTxtOldEndNo_LostFocus
End If
If Len(TxtOldEndNo) < 1 Then
   GoTo ExitTxtOldEndNo_LostFocus
End If

If CLng(TxtOldEndNo) < CLng(TxtOldStartNo) Then
   MsgBox "End number is lower than start number. This is not allowed"
   TxtOldEndNo.SetFocus
   TxtNoOfChqs = 0
Else
   TxtNoOfChqs = 1 + (CLng(TxtOldEndNo) - CLng(TxtOldStartNo))
End If
ExitTxtOldEndNo_LostFocus:
Exit Sub

End Sub

Private Sub TxtOldStartNo_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
   Else
   KeyAscii = 0
   GoTo Exit_TxtOldStartNo_KeyPress
End If

Exit_TxtOldStartNo_KeyPress:
Exit Sub

Err_TxtOldStartNo_KeyPress:
MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Old start number input error"
Resume Exit_TxtOldStartNo_KeyPress
End Sub
