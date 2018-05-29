VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSDI014 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registered Users"
   ClientHeight    =   3525
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   4875
   Icon            =   "SDI014.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   4875
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3012
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4620
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   3
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3228
      Columns(0).Caption=   "User Name"
      Columns(0).Name =   "User Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   30
      Columns(1).Width=   2831
      Columns(1).Caption=   "Login Id"
      Columns(1).Name =   "Login Id"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   10
      Columns(2).Width=   1482
      Columns(2).Caption=   "Access"
      Columns(2).Name =   "Access"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   2
      Columns(2).FieldLen=   1
      _ExtentX        =   8149
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
      Left            =   3840
      TabIndex        =   2
      Top             =   3120
      Width           =   888
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   328
      Left            =   1920
      TabIndex        =   1
      Top             =   3120
      Width           =   889
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   328
      Left            =   2880
      TabIndex        =   0
      Top             =   3120
      Width           =   888
   End
End
Attribute VB_Name = "frmSDI014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ErrWarn As New cLstWarn
Dim X As Integer
Dim rsUsers As New ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim OpenErr As Integer


Private Sub cmdClose_Click()
If Isloaded("frmSDI015") Then
  X = ErrWarn.ListWarn()
  frmSDI015.Show
  Exit Sub
Else
   rsUsers.Close
   Set rsUsers = Nothing
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
frmSDI015.Show 0
End Sub

Private Sub cmdNew_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

gblOptions = 1
gblFileKey = "0"
frmSDI015.Show 0
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
  frmSDI015.Show 0
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
If OpenErr = True Then
  Unload Me
Else
  frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
  Screen.MousePointer = vbDefault
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
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

  ' disable edit buttons if viewonly user
   If gblUserLevel = gblViewOnly Then
     cmdEdit.Enabled = False
     cmdNew.Enabled = False
   End If
   If gblUserLevel = 2 Then
     cmdNew.Visible = False
   End If
   OpenErr = False
Form_Load_Exit:
Exit Sub
Form_Load_Err:
  MsgBox "SDI014/Load"
  OpenErr = True
  On Error Resume Next
  Resume Form_Load_Exit
  Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
If Isloaded("frmSDI015") Then
  X = ErrWarn.ListWarn()
  Cancel = -1
  frmSDI015.Show
  Exit Sub
Else
  Unload Me
End If
End Sub

Private Sub UpdateScreen()
Dim TempPOP As String
Set rsUsers = RunSP(SpCon, "usp_SelectUsers", 1)
  With rsUsers
      grd.RemoveAll
      Do While Not .EOF
        TempPOP = ""
         TempPOP = !UserName & vbTab & !SystemName & vbTab & !rowinfo
         grd.AddItem TempPOP
         .MoveNext
      Loop
   End With
 End Sub
