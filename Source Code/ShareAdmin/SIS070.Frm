VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS070 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Accounts List"
   ClientHeight    =   3870
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   7170
   Icon            =   "SIS070.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3870
   ScaleWidth      =   7170
   Begin VB.CommandButton CmdFindCert 
      Caption         =   "Find By Certificate"
      Height          =   328
      Left            =   5160
      TabIndex        =   10
      Top             =   3120
      Width           =   1605
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Unpaid Chq"
      Enabled         =   0   'False
      Height          =   328
      Index           =   6
      Left            =   4080
      TabIndex        =   9
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Di&v History"
      Enabled         =   0   'False
      Height          =   328
      Index           =   5
      Left            =   3000
      TabIndex        =   8
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Cert &History"
      Enabled         =   0   'False
      Height          =   328
      Index           =   4
      Left            =   1920
      TabIndex        =   7
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Certificates"
      Enabled         =   0   'False
      Height          =   328
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   3480
      Width           =   1005
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Mandates"
      Enabled         =   0   'False
      Height          =   328
      Index           =   3
      Left            =   3000
      TabIndex        =   5
      Top             =   3480
      Width           =   1005
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   328
      Left            =   840
      TabIndex        =   4
      Top             =   3480
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   328
      Left            =   6240
      TabIndex        =   2
      Top             =   3480
      Width           =   885
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Details"
      Enabled         =   0   'False
      Height          =   328
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Display Name,  Address  and total shares..."
      Top             =   3480
      Width           =   1005
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Joint"
      Enabled         =   0   'False
      Height          =   328
      Index           =   2
      Left            =   5160
      TabIndex        =   0
      Top             =   3480
      Width           =   1005
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3045
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7080
      ScrollBars      =   2
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   3
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   6244
      Columns(0).Caption=   "Client Name"
      Columns(0).Name =   "Client Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   30
      Columns(1).Width=   2514
      Columns(1).Caption=   "Account No"
      Columns(1).Name =   "Account No"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   10
      Columns(2).Width=   3200
      Columns(2).Caption=   "Shares"
      Columns(2).Name =   "Shares"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   5
      Columns(2).NumberFormat=   "#,####"
      Columns(2).FieldLen=   11
      _ExtentX        =   12488
      _ExtentY        =   5371
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
End
Attribute VB_Name = "frmSIS070"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim iOpen As Integer
Dim ErrWarn As New cLstWarn
Dim X As Integer
Dim qIndex As Integer
Dim rsClients As New ADODB.Recordset
Dim qSQL As String
Dim SpCon As ADODB.Connection

Private Sub cmdBtn_Click(Index As Integer)
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
gblFileKey = grd.Columns(1).Text
frmSIS070.Visible = False
Select Case Index
Case 0 'Name & address details
    frmSIS071.Show 0 ' Display Name & Address details
Case 1 ' certificate enquiry
    frmSIS072.Show 0
Case 2 ' Joint
    frmSIS075.Show 0
Case 3 ' Mandates
    frmSIS076.Show 0
Case 4 'certificate history
    frmSIS074.Show 0
Case 5 ' Dividend History
   frmSIS073.Show 0
Case 6 ' Unpaid Cheques
   frmSIS073Unp.Show 0

Case Else
End Select

End Sub

Private Sub cmdClose_Click()

Set rsClients = Nothing
Unload Me
Set frmSIS070 = Nothing
'---------
'-- enableable menu options ---
'---------------------------
If gblUserLevel <> 3 Then
  frmMDI.mnuAct.Enabled = True
  frmMDI.mnuLists.Enabled = True
  frmMDI.mnuReports.Enabled = True
End If
If gblUserLevel = 1 Then frmMDI.mnuAdm.Enabled = True
End Sub

Private Sub cmdDet_Click()
frmSIS070.Visible = False
frmSIS071.Show 0 ' Display Name & Address details
frmSIS070.Visible = True
End Sub

Private Sub cmdFind_Click()
Dim i As Integer
Dim sWhere As String
On Error GoTo cmdFind_Err

Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
    .cbWhere.Clear
    .cbWhere.AddItem "CliName"
    .cbWhere.AddItem "ClientId"
    .cbWhere.ListIndex = 0
    .cbOptions.ListIndex = 0
    .lbl(3).Visible = False
    .optBtn.Visible = False
    .Show vbModal
    '----------------------------
    '-------- main line ---------
    '----------------------------
    If .tbFind.Text = vbNullString Then
    Else
      If iOpen = True Then
          rsClients.Close
          iOpen = False
      End If
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      qSQL = Trim(.tbFind.Text)
      X = .cbWhere.ListIndex
      '--------------------
      '-- display selection
      '--------------------
      qIndex = .cbOptions.ListIndex
      UpdateScreen
       '--------------------
       If grd.Rows > 0 Then
         For i = 0 To 6
             cmdBtn(i).Enabled = True
         Next
       Else
       For i = 0 To 6
             cmdBtn(i).Enabled = False
         Next
       End If
       '---------
       '-- disable menu options ---
       '---------------------------
       frmMDI.mnuAct.Enabled = False
       frmMDI.mnuLists.Enabled = False
       frmMDI.mnuAdm.Enabled = False
       End If
  End With
  Unload frmFind
  Set frmFind = Nothing
  
cmdFind_Exit:
  Exit Sub
cmdFind_Err:
  MsgBox "SIS070/cmdFind"
  
End Sub

Private Sub cmdMndate_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'--
gblFileKey = grd.Columns(1).Text
frmSIS010.Show 0
End Sub

Private Sub cmdNew_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

gblOptions = 1
gblFileKey = "0"
frmSIS002.Show 0
End Sub

Private Sub CmdFindCert_Click()
On Error GoTo Err_CmdFindCert_Click
Dim sMsg As String
Dim sTitle As String
Dim adoCert As ADODB.Recordset
Dim CertType As String
Dim i As Integer

sMsg = "Please enter Certificate Number"
sTitle = "Find By Certificate Number"

qSQL = 1
qSQL = InputBox(sMsg, sTitle, qSQL)
If IsEmpty(qSQL) Or qSQL = "" Then
    GoTo Exit_CmdFindCert_Click
End If
If IsNumeric(qSQL) Then
   Set adoCert = RunSP(SpCon, "usp_FindByCert", 1, qSQL)
   If adoCert.EOF Then
      MsgBox "Sorry, that Certificate Number was not found"
      GoTo Exit_CmdFindCert_Click
   Else
       qSQL = adoCert!ClientID
       CertType = adoCert!Status
       If CertType = "C" Then
          MsgBox "Please click the Certificate History Button to view list with certificate"
       Else
          MsgBox "Please click the Certificates Button to view list with certificate"
       End If
       qIndex = 0
       X = 1
       UpdateScreen
       If grd.Rows > 0 Then
         For i = 0 To 6
             cmdBtn(i).Enabled = True
         Next
       Else
       For i = 0 To 6
             cmdBtn(i).Enabled = False
         Next
       End If
       '---------
       '-- disable menu options ---
       '---------------------------
       frmMDI.mnuAct.Enabled = False
       frmMDI.mnuLists.Enabled = False
       frmMDI.mnuAdm.Enabled = False
   End If
Else
   MsgBox "Certificate Number must be Numeric"
   GoTo Exit_CmdFindCert_Click
End If
Exit_CmdFindCert_Click:
Exit Sub
Err_CmdFindCert_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error executing Find by Certificate"
GoTo Exit_CmdFindCert_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub grd_Click()
gblFileKey = grd.Columns(1).CellValue(grd.GetBookmark(0))
End Sub
Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If gblUserLevel <> gblViewOnly Then
  frmMDI.txtStatusMsg.SimpleText = "Double Click to Edit"
  frmMDI.txtStatusMsg.Refresh
End If
End Sub

Private Sub grd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If gblUserLevel <> gblViewOnly Then
  frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
  frmMDI.txtStatusMsg.Refresh
End If
End Sub


Private Sub Form_Activate()
' status ready
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim strTmp As String
' readymsg
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
csvCenterForm Me, gblMDIFORM
'--
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
Set rsClients = New ADODB.Recordset
iOpen = False

End Sub
Private Sub UpdateScreen()
Dim sRowinfo As String
If X = 1 Then
   Set rsClients = RunSP(SpCon, "usp_ClientFind", 1, CLng(qSQL), qIndex, X, 1)
Else
   Set rsClients = RunSP(SpCon, "usp_ClientFind", 1, qSQL, qIndex, X, 1)
End If
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

With rsClients
      grd.RemoveAll
      If Not .EOF Then
         Do While Not .EOF
          sRowinfo = !CliName & vbTab & !ClientID & _
                      vbTab & !shares
          grd.AddItem sRowinfo
         .MoveNext
        Loop
      End If
End With
End Sub


