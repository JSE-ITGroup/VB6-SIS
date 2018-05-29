VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmAutoRecon 
   BackColor       =   &H80000013&
   Caption         =   "Automatic Reconciliation"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FrmAutoRecon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
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
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start Now"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBAccount 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   2895
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
      _ExtentX        =   5106
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
      BackColor       =   &H80000013&
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "FrmAutoRecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdStart_Click()
On Error GoTo Err_CmdStart_Click
Dim i As Integer
Dim StrSql As String

If SSDBAccount = "" Then
   MsgBox "An Account must be selected before starting the reconciliation", vbOKOnly
   GoTo Exit_CmdStart_Click
End If
If MsgBox("Confirm Automatic Reconciliation", vbYesNo) = vbYes Then
    frmMDI.txtStatusMsg.SimpleText = "Processing reconciliation file, Please wait......"
    frmMDI.txtStatusMsg.Refresh
    i = RunSP(SpCon, "usp_AutoRecon", 0, SSDBAccount.Columns(0).Text, gblLoginName)
    If i <> 0 Then
       StrSql = "An error occurred which stopped the update" & vbCrLf
       StrSql = StrSql & "The server returned error number " & i & "." & vbCrLf
       StrSql = StrSql & "Please report this error number to the techical team."
       MsgBox StrSql, vbOKOnly, "Auto Reconciliation Error"
       GoTo Exit_CmdStart_Click
    Else
        MsgBox "Process successfully completed", vbOKOnly
    End If
End If

Exit_CmdStart_Click:
Exit Sub

Err_CmdStart_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on starting automatic reconciliation"
Resume Exit_CmdStart_Click

End Sub

Private Sub Form_Activate()
' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 '--
End Sub

Private Sub Form_Load()
On Error GoTo FL_ERR
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 '--
csvCenterForm Me, gblMDIFORM
'-----------------------------------
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

'--
   
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox Err & " " & Err.Description, , "Error on loading Auto Reconciliation Screen"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
   
End Sub
Private Sub SSDBAccount_InitColumnProps()
On Error GoTo Err_SSDBAccount_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_ListActiveAccounts", 1)
If adoRst.EOF Then
   MsgBox "Accounts were not setup" & vbCrLf & "Please do so now", vbCritical, "Account Error"
   GoTo Exit_SSDBAccount_InitColumnProps
End If

With SSDBAccount
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
Exit_SSDBAccount_InitColumnProps:
Exit Sub

Err_SSDBAccount_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing active accounts"
Resume Exit_SSDBAccount_InitColumnProps

End Sub
Private Sub CmdExit_Click()
SpCon.Close
Unload Me
End Sub

