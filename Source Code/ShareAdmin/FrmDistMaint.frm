VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmDistMaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Distribution Maintenance"
   ClientHeight    =   2625
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6660
   Icon            =   "FrmDistMaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6660
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3480
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5640
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4560
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   1
      ToolTipText     =   "Enter the Company's Tax Reference Number"
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox tbfld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   0
      ToolTipText     =   "Use generate number or enter your own unique client Number"
      Top             =   720
      Width           =   855
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBCurrency 
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
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
      Columns.Count   =   3
      Columns(0).Width=   1111
      Columns(0).Caption=   "Currency"
      Columns(0).Name =   "Currency"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3413
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Currency Code"
      Columns(2).Name =   "Currency Code"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   1931
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
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Currency:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1740
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblLabels 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ver:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Distribution Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "FrmDistMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim rsMain As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim OpenErr As Integer
Dim iOpenMain As Integer
Dim iOpenAdt As Integer
Dim iOpenCmp As Integer
Function IsValid() As Integer
Dim iErr As Integer
Dim msg As String

IsValid = True
 If tbfld(1) = "" Then ' Distribution Description
   tbfld(1).SetFocus
   msg = "Distribution Description is missing"
   GoTo Validate_Err
 End If
 tbfld(1) = Trim(tbfld(1))
 '--
 If SSDBCurrency = "" Then  ' Tax Rate
   msg = "Currency Code is missing"
   SSDBCurrency.SetFocus
   GoTo Validate_Err
 End If
 
Validate_Exit:
   Exit Function
'--
Validate_Err:
  MsgBox msg, vbInformation, "Payment Distribution Rates"
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
   tbfld(0).SetFocus
Else
   ClearScreen
   tbfld(1).SetFocus
End If
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer
Dim i As Integer
On Error GoTo cmdUpdate_Err
If IsValid Then
  '--
  i = RunSP(SpCon, "usp_DistCodeUpdate", 0, tbfld(0), tbfld(1), CInt(SSDBCurrency.Columns(3).Text), gblLoginName)
  If i <> 0 Then
      MsgBox "There was an error saving the changes. Please re-try"
      GoTo Done
  End If
  If gblOptions = 1 Then
     ClearScreen
  Else
     Shutdown
     Unload Me
  End If
End If
'---

Done:
 Exit Sub
'--
cmdUpdate_Err:
  Shutdown
  Unload Me
End Sub
Private Sub Form_Activate()
' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 
 If gblOptions = 1 Then
  Me.Caption = "New Payment Distribution"
  tbfld(0).Text = "0"
''Else
''  UpdateScreen
End If
 '--
 If OpenErr = True Then
  Shutdown
  Unload Me
 Else
   If gblOptions = 2 Then
      Me.Caption = "Edit Payment Distribution"
      UpdateScreen
   End If
End If
End Sub

Private Sub Form_Load()
Dim indx As Integer
Dim strTmp As String
On Error GoTo FL_ERR
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
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

OpenErr = False
iOpenCmp = False
iOpenMain = False
iOpenAdt = False
iOpenMain = True

'--
   
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "Error on loading FrmDistMaint"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
   
End Sub
Private Sub UpdateScreen()
Dim i As Integer, bm As Variant
Set rsMain = RunSP(SpCon, "usp_DistCodeEdit", 1, CInt(gblFileKey))
 With rsMain
    tbfld(0).Text = !DistCode
    tbfld(1).Text = !DistDesc
    For i = 0 To SSDBCurrency.Rows - 1
       bm = SSDBCurrency.GetBookmark(i)
       If SSDBCurrency.Columns(3).CellText(bm) = !CurrencyCode Then
          SSDBCurrency.Bookmark = SSDBCurrency.GetBookmark(i)
          SSDBCurrency = SSDBCurrency.Columns(0).CellText(bm)
          Exit For
       End If
    Next i
 End With
End Sub



Private Sub meb_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  cmdUpdate.SetFocus
Case vbKeyUp
  tbfld(1).SetFocus
End Select
End Sub

Private Sub SSDBCurrency_InitColumnProps()
On Error GoTo Err_SSDBCurrency_InitColumnProps
Dim Strsql As String
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_SelectCurrency", 1)
If adoRst.EOF Then
   MsgBox "Currencies were not setup" & vbCrLf & "Please do so now", vbCritical, "Currency Error"
   GoTo Exit_SSDBCurrency_InitColumnProps
End If

''adoRst.MoveFirst
With SSDBCurrency
     .RemoveAll
     Do While Not adoRst.EOF
     Strsql = adoRst(0) & vbTab & adoRst(1) & vbTab & adoRst(2)
     .AddItem Strsql
     adoRst.MoveNext
     Strsql = ""
     Loop
End With
adoRst.Close
Set adoRst = Nothing
Exit_SSDBCurrency_InitColumnProps:
Exit Sub

Err_SSDBCurrency_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Currency load Error"
Resume Exit_SSDBCurrency_InitColumnProps
End Sub

Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
   Select Case Index
   Case 0
      tbfld(1).SetFocus
   Case 1
      SSDBCurrency.SetFocus
   End Select
Case vbKeyUp
   Select Case Index
   Case 1
     If gblOptions = 1 Then tbfld(0).SetFocus
   Case Else
   End Select
Case Else
End Select
End Sub

Private Sub ClearScreen()
  For X = 0 To 1
    tbfld(X).Text = ""
  Next
  SSDBCurrency.Text = ""
  If gblOptions = 2 Then
     UpdateScreen
     tbfld(1).SetFocus
  End If
End Sub

Private Sub Shutdown()
SpCon.Close
End Sub
