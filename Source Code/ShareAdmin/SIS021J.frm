VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS021J 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JCSD Bank & Branch Maintenance"
   ClientHeight    =   3765
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SIS021J.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6795
   Begin VB.TextBox TxtBankName 
      BackColor       =   &H00E0E0E0&
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
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   6615
   End
   Begin VB.TextBox txtBranch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Enter a name Branch name as it appears in the JCSD file"
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   3
      ToolTipText     =   "Clear screen"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Cancel current change and exit."
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4680
      TabIndex        =   2
      ToolTipText     =   "Save changes"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox TxtInstitution 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   0
      ToolTipText     =   "Enter a name Institution name as it appears in the JCSD file"
      Top             =   2040
      Width           =   4335
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBBankID 
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   600
      Width           =   1455
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
      Columns(0).Width=   1455
      Columns(0).Caption=   "BankID"
      Columns(0).Name =   "Account Number"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   8123
      Columns(1).Caption=   "Institution"
      Columns(1).Name =   "Currency"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2566
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
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   9480
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblLabels 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Institution:"
      Height          =   255
      Index           =   16
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   1620
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ver:"
      Height          =   255
      Index           =   20
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Bank Id:"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Branch:"
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS021J"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, ifirstime As Integer
Dim rsMain As ADODB.Recordset
Dim rsAdt As ADODB.Recordset
Dim rsBank As ADODB.Recordset 'used to check for duplicates
Dim iOpenMain As Integer
Dim iOpenAdt As Integer
Dim OpenErr As Integer
Dim iOpenBank As Integer
Dim strTable As String
Dim strRecNO As String
Dim iNewAcct As Long
Dim SpCon As ADODB.Connection
Function IsValid() As Integer
Dim iErr As String, sql As String
IsValid = True
'--
If SSDBBankID = "" Then  ' Bank Number
   iErr = "bank ID Missing"
   SSDBBankID.SetFocus
   GoTo Validate_Err
 End If
 '--
 If TxtInstitution = "" Then 'Bank name
       iErr = "Bank Name is missing"
       TxtInstitution.SetFocus
       GoTo Validate_Err
 End If
 '--
 
 '--

 txtBranch = Trim(txtBranch)
 '--
Validate_Exit:
   '''Set rsBank = Nothing
   Exit Function
'--
Validate_Err:
  'MsgBox msg, vbInformation, "Clients"
  MsgBox iErr, vbOKOnly, "Bank Name & Address"
  IsValid = False
  GoTo Validate_Exit
'--
End Function

Private Sub cmdCancel_Click()
  Shutdown
  Unload Me

End Sub

Private Sub cmdClear_Click()

If gblOptions = 1 Then
   ClearScreen
   SSDBBankID.SetFocus
Else
   ClearScreen
   TxtInstitution.SetFocus
End If
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer
Dim i As Integer, imsg As Integer
On Error GoTo cmdUpdate_Err

If IsValid Then
  '--
  strRecNO = SSDBBankID.Columns(0).Text
  strChg = 0
  i = RunSP(SpCon, "usp_SIS021JB", 0, gblOptions, SSDBBankID.Columns(0).Text, TxtInstitution, txtBranch, _
      strRecNO, CInt(gblFileKey), gblLoginName)
  If i <> 0 Then
     MsgBox "Update failed", vbCritical + vbOKOnly
  Else
     MsgBox "Update was successful"
  End If
End If
'---
Done:
 Shutdown
 Unload Me
 Exit Sub
'--
cmdUpdate_Err:
  MsgBox "Error on Update", vbOKOnly, "SIS021J/cmdUpdate"
  GoTo Done
End Sub
Private Sub Form_Activate()
On Error GoTo Form_Activate_Err
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
'--
If OpenErr = True Then
  Shutdown
  Unload Me
Else
 If gblOptions = 2 Then
   UpdateScreen
   Me.Caption = "Edit Bank Name & Address"
   SSDBBankID.Enabled = False
 End If
End If
'--
Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
   MsgBox "SIS021/Activate"
   Shutdown
   Unload Me
      Exit Sub
End Sub

Private Sub Form_Load()
Dim iDay As Integer
Dim qSQL As String
Dim indx As Integer
Dim strTmp As String
On Error GoTo FL_ERR
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
 ifirstime = 0
'--
csvCenterForm Me, gblMDIFORM
'-----------------------------------
OpenErr = False
iOpenMain = False
iOpenAdt = False
iOpenBank = False
'-----
Set rsMain = New ADODB.Recordset
Set rsAdt = New ADODB.Recordset
Set rsBank = New ADODB.Recordset
'----------------------------
'---- open recordsets -----
'----------------------------
'''Set rsAdt.ActiveConnection = Nothing
iOpenAdt = True
'----------------------------------------
' create SQL for selecting record to edit
'----------------------------------------
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


'---
Set rsMain = RunSP(SpCon, "usp_SIS021JA", 1, CInt(gblFileKey))
iOpenMain = True
If gblOptions = 1 Then
    InitAddNew
End If
'--

FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS021/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
End Sub
Private Sub UpdateScreen()
Dim i As Integer, bm As Variant
With SSDBBankID
     .MoveFirst
     For i = 0 To .Rows - 1
         bm = .GetBookmark(i)
         If .Columns(0).CellText(bm) = rsMain!BankID Then
            .Bookmark = .GetBookmark(i)
             SSDBBankID = .Columns(0).CellText(bm)
         Exit For
         End If
     Next i
 End With

TxtBankName = SSDBBankID.Columns(1).Text
TxtInstitution = rsMain!Institution
If Not IsNull(rsMain!Branch) Then
   txtBranch = rsMain!Branch
End If
End Sub

Private Sub SSDBBankID_Click()
TxtBankName = SSDBBankID.Columns(1).Text
End Sub

Private Sub SSDBBankID_InitColumnProps()
On Error GoTo Err_SSDBBankID_InitColumnProps
Dim adoRst As New ADODB.Recordset
Dim StrSql As String

Set adoRst = RunSP(SpCon, "usp_Sis020", 1)

With SSDBBankID
     .RemoveAll
      StrSql = ""
Do While Not adoRst.EOF
        StrSql = adoRst(1) & vbTab & adoRst(0) & vbTab
        .AddItem StrSql
        adoRst.MoveNext
Loop
End With
adoRst.Close
Set adoRst = Nothing

Exit_SSDBBankID_InitColumnProps:
Exit Sub

Err_SSDBBankID_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error while populating Bank IDs"
Resume Exit_SSDBBankID_InitColumnProps

End Sub

Private Sub ClearScreen()
Dim qSQL As String

TxtInstitution = ""
txtBranch = ""
SSDBBankID = ""
If gblOptions = 2 Then
     UpdateScreen
End If
End Sub
Private Sub InitAddNew()
  ClearScreen
  Me.Caption = "New JCSD Institution Name & Branch"
  End Sub

Private Sub Shutdown()
If iOpenMain = True Then rsMain.Close
Set rsMain = Nothing

End Sub

