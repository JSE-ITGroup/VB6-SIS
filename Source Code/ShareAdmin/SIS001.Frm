VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS001 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Accounts List"
   ClientHeight    =   3525
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   6540
   Icon            =   "SIS001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   6540
   Begin VB.CommandButton cmdCert 
      Caption         =   "Ce&rtificates"
      Height          =   328
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Certificates Splits and Assignments"
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdMndate 
      Caption         =   "&Mandates"
      Height          =   328
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "Assign Mandate Accounts for Payments"
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   328
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Find Clients by Name or Accounts"
      Top             =   3120
      Width           =   1005
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6360
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   3
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   7567
      Columns(0).Caption=   "Client Name"
      Columns(0).Name =   "Client Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   30
      Columns(1).Width=   3043
      Columns(1).Caption=   "Account No"
      Columns(1).Name =   "Account No"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   10
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "CatCode"
      Columns(2).Name =   "CatCode"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   2
      _ExtentX        =   11218
      _ExtentY        =   5318
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   328
      Left            =   5520
      TabIndex        =   2
      ToolTipText     =   "Close this form"
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   328
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Edit Client Name and Adress Details"
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   328
      Left            =   4440
      TabIndex        =   0
      ToolTipText     =   "Adds new Client name and address detaits"
      Top             =   3120
      Width           =   1005
   End
End
Attribute VB_Name = "frmSIS001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim ErrWarn As New cLstWarn
Dim qSQL As String
Dim qIndex As Integer
Dim X As Integer, iOpen As Integer, OpenErr As Integer
Dim rsClients As New ADODB.Recordset
Dim SpCon As ADODB.Connection

Private Sub cmdCert_Click()
If grd.Columns(2).Text = "SB" Then 'can't assign
  cmdCert.SetFocus
  Exit Sub
End If
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

gblFileKey = grd.Columns(1).Text
frmSIS011.Show 0
End Sub

Private Sub cmdClose_Click()
If Isloaded("frmSIS002") Then
  X = ErrWarn.ListWarn()
  frmSIS002.Show
  Exit Sub
Else
   If iOpen = True Then rsClients.Close
   Set rsClients = Nothing
  Unload Me
End If
End Sub


Private Sub cmdEdit_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

gblFileKey = grd.Columns(1).Text
gblOptions = 2
frmSIS002.Show 0
End Sub

Private Sub cmdFind_Click()
Dim i As Integer
'On Error GoTo cmdFind_Err
Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
     .cbWhere.Clear
    .cbWhere.AddItem "Client Name"
    .cbWhere.AddItem "Account No"
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
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      qSQL = UCase(Trim(.tbFind.Text))
      X = .cbWhere.ListIndex
      '---
      If qSQL <> "" Then
        grd.RemoveAll
        If iOpen = True Then
          rsClients.Close
          iOpen = False
        End If
        qIndex = .cbOptions.ListIndex
        iOpen = True
        UpdateScreen
        
        If grd.Rows > 0 Then
           cmdCert.Enabled = True
           cmdMndate.Enabled = True
           cmdEdit.Enabled = True
        Else
          cmdCert.Enabled = False
          cmdMndate.Enabled = False
          cmdEdit.Enabled = False
        End If
      End If
    End If
  End With
  Unload frmFind
  Set frmFind = Nothing
  If cmdCert.Enabled = True Then cmdCert.SetFocus Else CmdFind.SetFocus
  
cmdFind_Exit:
  Exit Sub
cmdFind_Err:
  MsgBox "Error on Find Option"
  cmdClose_Click
End Sub

Private Sub cmdMndate_Click()
If grd.Columns(2).Text = "SB" Then 'can't assign
  cmdMndate.SetFocus
  Exit Sub
End If
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'--
gblFileKey = grd.Columns(1).Text
gblHold = grd.Columns(0).Text
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
Private Sub Form_GotFocus()
UpdateScreen
End Sub

Private Sub grd_Click()
gblFileKey = grd.Columns(1).CellValue(grd.GetBookmark(0))
End Sub

Private Sub grd_DblClick()
' disable dblclick for viewonly users
If gblUserLevel <> gblViewOnly Then
  ' wait message & hourglass
  Screen.MousePointer = vbHourglass
  frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
  frmMDI.txtStatusMsg.Refresh

  gblFileKey = grd.Columns(1).CellValue(grd.GetBookmark(0))
  gblOptions = 2
  frmSIS002.Show 0
End If
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

If OpenErr = True Then
  Unload Me
Else
  If iOpen = True Then
    UpdateScreen
  Else
    If gblFileKey <> "0" Then
      qIndex = 0
      X = 1
      UpdateScreen
      iOpen = True
   End If
 End If
End If
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim strTmp As String
   ' readymsg
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   csvCenterForm Me, gblMDIFORM
   iOpen = False: OpenErr = False
   '--
   Set rsClients = New ADODB.Recordset
   '--
   On Error GoTo Form_Load_Err:
   '--
   gblFileKey = "0"
   cmdCert.Enabled = False
   cmdMndate.Enabled = False
   cmdEdit.Enabled = False
   '--
   'disable edit buttons if viewonly user
   If gblUserLevel = gblViewOnly Then
     cmdEdit.Enabled = False
     cmdNew.Enabled = False
   End If
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
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

FL_Exit:
Exit Sub
Form_Load_Err:
  MsgBox Err & " " & Err.Description, vbOKOnly, "Error on Loading SIS001"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit

End Sub


Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
If Isloaded("frmSIS002") Then
  X = ErrWarn.ListWarn()
  Cancel = -1
  frmSIS002.Show
  Exit Sub
Else
  Unload Me
End If
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
          sRowinfo = !CliName & vbTab & !ClientID & Chr(9)
          sRowinfo = sRowinfo & !CatCode
          grd.AddItem sRowinfo
          grd.Redraw = True
         .MoveNext
        Loop
      End If
End With
End Sub


