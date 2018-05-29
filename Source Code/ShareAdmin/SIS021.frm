VERSION 5.00
Begin VB.Form frmSIS021 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank & Branch Maintenance"
   ClientHeight    =   3765
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "SIS021.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6795
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   2
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   2
      ToolTipText     =   "Mandatory line 1 of address."
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3600
      TabIndex        =   8
      ToolTipText     =   "Clear screen"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5760
      TabIndex        =   9
      ToolTipText     =   "Cancel current change and exit."
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4680
      TabIndex        =   7
      ToolTipText     =   "Save changes"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Enter a name for the branch or if individual lastname,first names"
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox tbfld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   0
      ToolTipText     =   "Enter unique bank identifier."
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   6
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   6
      ToolTipText     =   "Optional line 5 of address."
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   5
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   5
      ToolTipText     =   "Optional line 3 of address"
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   4
      Left            =   1920
      MaxLength       =   25
      TabIndex        =   4
      ToolTipText     =   "Optional line 4 of address."
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   3
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   3
      ToolTipText     =   "Mandatory line 2 of address"
      Top             =   2040
      Width           =   3375
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
      TabIndex        =   14
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Bank & Branch Name:"
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
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   1860
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
      TabIndex        =   12
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Bank Id:"
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
      Left            =   360
      TabIndex        =   11
      Top             =   720
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
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
      Index           =   3
      Left            =   1080
      TabIndex        =   10
      Top             =   1800
      Width           =   735
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
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS021"
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
If tbfld(0) = "" Then  ' Bank Number
   iErr = "bank ID Missing"
   tbfld(0).SetFocus
   GoTo Validate_Err
 End If
 tbfld(0) = UCase(tbfld(0))
 '--
 If tbfld(1) = "" Then 'Bank name
       iErr = "Bank Name is missing"
       tbfld(1).SetFocus
       GoTo Validate_Err
 End If
 '--
 
Set rsBank = RunSP(SpCon, "usp_SIS021A", 1, Trim(tbfld(0)))
 '''Set rsBank.ActiveConnection = Nothing
 '
 If Not rsBank.EOF And gblOptions = 1 Then
   iErr = "Bank Id already in use"
   ''tbfld(0).SetFocus
   rsBank.Close
   GoTo Validate_Err
 End If
 rsBank.Close
 
 '--
 If tbfld(2) = "" Then ' Address Line 1
   iErr = "Address Line 1 is required"
   tbfld(2).SetFocus
   GoTo Validate_Err
 End If
 tbfld(2) = Trim(tbfld(2))
 '--
 If tbfld(2) = "" Then  ' Address Line 2
   iErr = "Address Line 2 is required"
   tbfld(2).SetFocus
   GoTo Validate_Err
 End If
 tbfld(3) = Trim(tbfld(3))
 tbfld(4) = Trim(tbfld(4))
 tbfld(5) = Trim(tbfld(5))
 tbfld(6) = Trim(tbfld(6))
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
   tbfld(0).Text = iNewAcct
   tbfld(0).SetFocus
Else
   ClearScreen
   tbfld(1).SetFocus
End If
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer
Dim i As Integer, imsg As Integer
On Error GoTo cmdUpdate_Err

If IsValid Then
  '--
  strRecNO = tbfld(0)
  strChg = 0
  i = RunSP(SpCon, "usp_SIS021B", 0, gblOptions, tbfld(0), tbfld(1), tbfld(2), _
      tbfld(3), tbfld(4), tbfld(5), tbfld(6), strRecNO, gblLoginName)
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
  MsgBox "SIS021/cmdUpdate"
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
   tbfld(0).Enabled = False
 End If
End If
'--
Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
   MsgBox Err.Description, vbOKOnly, "Error on SIS021 Activate"
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
qSQL = "Select * from [BNKREF] "
qSQL = qSQL & "where [BANKID] = '"
qSQL = qSQL & gblFileKey & "'"
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
Set rsMain = RunSP(SpCon, "usp_SIS021A", 1, gblFileKey)
iOpenMain = True
If gblOptions = 1 Then
    InitAddNew
End If
'--

FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox Err.Description, vbOKOnly, "SIS021/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
End Sub
Private Sub UpdateScreen()
Dim i As Integer, bm As Variant
 With rsMain
    tbfld(0).Text = !BankId
    tbfld(1).Text = !BnkName
    tbfld(2).Text = !BNKADDR1
    tbfld(3).Text = !BNKADDR2
    If Not IsNull(!BNKADDR3) Then
        tbfld(4).Text = !BNKADDR3
    End If
    If Not IsNull(!BNKADDR4) Then
      tbfld(5).Text = !BNKADDR4
    End If
    If Not IsNull(!BNKADDR5) Then
       tbfld(6).Text = !BNKADDR5
    End If
   
 End With
End Sub


Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
   KeyCode = 0
   Select Case Index
   Case 0 To 5
      tbfld(Index + 1).SetFocus
   Case 6
      cmdUpdate.SetFocus
   Case Else
   End Select
Case vbKeyUp
   KeyCode = 0
   Select Case Index
   Case 1
      If gblOptions = 1 Then tbfld(0).SetFocus
   Case 2 To 6
     tbfld(Index - 1).SetFocus
   Case Else
   End Select
Case Else
End Select
End Sub

Private Sub ClearScreen()
Dim qSQL As String

  For X = 0 To 6
    tbfld(X).Text = ""
  Next
  If gblOptions = 2 Then
     UpdateScreen
  End If
End Sub
Private Sub InitAddNew()
  ClearScreen
  Me.Caption = "New Bank Name & Address"
  End Sub

Private Sub Shutdown()
If iOpenMain = True Then rsMain.Close
Set rsMain = Nothing

End Sub

