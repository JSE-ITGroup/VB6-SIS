VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS045 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stockholder Profile Selection"
   ClientHeight    =   3555
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   6660
   Icon            =   "SIS045.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   6660
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   328
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   1005
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3012
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6360
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   3
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
      Left            =   5400
      TabIndex        =   1
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   328
      Left            =   4320
      TabIndex        =   0
      Top             =   3120
      Width           =   1005
   End
End
Attribute VB_Name = "frmSIS045"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim ErrWarn As New cLstWarn
Dim repSISRept As New SISRepts
Dim qSQL As String
Dim X As Integer
Dim qIndex As Integer
Dim SpCon As New ADODB.Connection
Dim rsClients As New ADODB.Recordset
Dim errLoop As Error
Dim errs1 As Error
Private Sub cmdClose_Click()
Set rsClients = Nothing
SpCon.Close
Unload Me
Set frmSIS045 = Nothing
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
      sWhere = UCase(Trim(.tbFind.Text))
      X = .cbWhere.ListIndex
      '---
      qIndex = .cbOptions.ListIndex
      qSQL = sWhere
      UpdateScreen
      cmdPrint.Enabled = True
    End If
  End With
  Unload frmFind
  Set frmFind = Nothing
cmdFind_Exit:
  Exit Sub
cmdFind_Err:
  MsgBox "SIS045/Load"
  Unload Me
End Sub

Private Sub cmdPrint_Click()
gblFileKey = grd.Columns(1).CellValue(grd.GetBookmark(0))
Set repSISRept = New SISRepts
repSISRept.ClientID = gblFileKey
repSISRept.LoginId = gblFileName
repSISRept.ReportNumber = 5
repSISRept.ReportType = 9
repSISRept.siteid = gblSiteId
'repSISRept.DSN = gblDSN
repSISRept.RunShareHolderReport
End Sub
Private Sub grd_DblClick()
  'Screen.MousePointer = vbHourglass
  frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
  frmMDI.txtStatusMsg.Refresh
  cmdPrint_Click

End Sub


Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If gblUserLevel <> gblViewOnly Then
  frmMDI.txtStatusMsg.SimpleText = "Double Click to Print"
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
Set rsClients = New ADODB.Recordset
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
On Error GoTo Form_Load_Err:
'--
gblFileKey = "0"
'--
Exit Sub
Form_Load_Err:
  MsgBox "SIS045/Load"
  Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub UpdateScreen()
Dim sRowinfo As String
Set rsClients = RunSP(SpCon, "usp_ClientFind", 1, qSQL, qIndex, X, 1)
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

With rsClients
      grd.RemoveAll
      If Not .EOF Then
        Do While Not .EOF
          sRowinfo = !CliName & Chr(9) & !ClientID & Chr(9)
          sRowinfo = sRowinfo & !CatCode
          grd.AddItem sRowinfo
          grd.Redraw = True
         .MoveNext
        Loop
      End If
End With
End Sub


