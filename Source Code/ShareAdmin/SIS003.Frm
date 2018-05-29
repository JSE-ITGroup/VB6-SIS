VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS003 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Who Is Logged On"
   ClientHeight    =   3525
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   4860
   Icon            =   "SIS003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   4860
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3012
      Left            =   120
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   3120
      Width           =   888
   End
End
Attribute VB_Name = "frmSIS003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ErrWarn As New cLstWarn
Dim X As Integer, OpenErr As Integer
Dim rsUsers As New ADODB.Recordset
Dim SpCon As ADODB.Connection

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub Form_GotFocus()
UpdateScreen
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
Loop
Screen.MousePointer = vbDefault
   OpenErr = False
   '--
   On Error GoTo Form_Load_Err:
   Set rsUsers = RunSP(SpCon, "usp_WhoLoggedOn", 1)
FL_Exit:
   Exit Sub

Form_Load_Err:
   MsgBox "SIS003/Load"
   OpenErr = True
   On Error Resume Next
   Resume FL_Exit
End Sub
Private Sub UpdateScreen()
Dim qSQL As String

With rsUsers
      grd.RemoveAll
      .MoveFirst
      Do While Not .EOF
         qSQL = !UserName & vbTab & !SystemName & vbTab & !userlevel
         grd.AddItem qSQL
         .MoveNext
      Loop
      
   End With
   grd.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsUsers.Close
Set rsUsers = Nothing
SpCon.Close
End Sub
