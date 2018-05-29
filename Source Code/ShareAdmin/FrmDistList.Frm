VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmDistList 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Available Payment Distribution Methods"
   ClientHeight    =   3555
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   5475
   Icon            =   "FrmDistList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   5475
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   328
      Left            =   600
      TabIndex        =   5
      Top             =   3120
      Width           =   889
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   328
      Left            =   1560
      TabIndex        =   4
      Top             =   3120
      Width           =   889
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3012
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5280
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   3200
      Columns(0).Caption=   "Type"
      Columns(0).Name =   "Type"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Currency"
      Columns(1).Name =   "Currency"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Code"
      Columns(2).Name =   "Code"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Editable"
      Columns(3).Name =   "Editable"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      _ExtentX        =   9313
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
      Left            =   4440
      TabIndex        =   2
      Top             =   3120
      Width           =   888
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   328
      Left            =   2520
      TabIndex        =   1
      Top             =   3120
      Width           =   889
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   328
      Left            =   3480
      TabIndex        =   0
      Top             =   3120
      Width           =   888
   End
End
Attribute VB_Name = "frmDistList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ErrWarn As New cLstWarn
Dim X As Integer
Dim SpCon As ADODB.Connection
Dim rsMain As New ADODB.Recordset
Dim OpenErr As Integer
Private Sub cmdClose_Click()
If Isloaded("FrmDistMaint") Then
  X = ErrWarn.ListWarn()
  FrmDistMaint.Show
  Exit Sub
Else
   rsMain.Close
   Set rsMain = Nothing
  Unload Me
End If
End Sub
Private Sub cmdEdit_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

If grd.Columns(2).Text = "0" Then
   MsgBox "This type cannot be edited", vbOKOnly, "Distribution Type: Cheque"
   Screen.MousePointer = vbDefault
   GoTo Exit_CmdEdit
End If
gblFileKey = grd.Columns(2).Text
gblOptions = 2
FrmDistMaint.Show 0
Exit_CmdEdit:
Exit Sub
End Sub

Private Sub cmdNew_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

gblOptions = 1
gblFileKey = "0"
FrmDistMaint.Show 0
End Sub

Private Sub cmdPrint_Click()
frmSIS049.Show
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
  FrmDistMaint.Show 0
End If
End Sub


Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If gblUserLevel <> gblViewOnly Then
  frmMDI.txtStatusMsg.SimpleText = "Double Click to Edit"
  frmMDI.txtStatusMsg.Refresh
End If
End Sub

Private Sub grd_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
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
  UpdateScreen
End If
End Sub
Private Sub Form_Load()
   Dim qSQL As String
   Dim i As Integer
   Dim strTmp As String
   ' readymsg
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   csvCenterForm Me, gblMDIFORM
   OpenErr = False
   '--
   On Error GoTo Form_Load_Err:
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
  
  ' disable edit buttons if viewonly user
   If gblUserLevel = gblViewOnly Then
     cmdEdit.Enabled = False
     cmdNew.Enabled = False
   End If
FL_Exit:
Exit Sub
Form_Load_Err:
  MsgBox Err & " " & Err.Description, vbOKOnly, "FrmDistCode/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Isloaded("frmDistMaint") Then
  X = ErrWarn.ListWarn()
  Cancel = -1
  FrmDistMaint.Show
  Exit Sub
Else
  SpCon.Close
  Unload Me
End If
End Sub
Private Sub UpdateScreen()
Dim sRowinfo As String
Dim qSQL As String

Set rsMain = RunSP(SpCon, "usp_DistCodes", 1)

With rsMain
      grd.RemoveAll
      If Not .EOF Then
        Do While Not .EOF
          sRowinfo = !DistDesc & vbTab & !Currency & vbTab & !DistCode & vbTab & !Editable
          grd.AddItem sRowinfo
         .MoveNext
        Loop
      End If
End With
End Sub

