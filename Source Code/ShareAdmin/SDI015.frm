VERSION 5.00
Begin VB.Form frmSDI015 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   3075
   ClientLeft      =   510
   ClientTop       =   1920
   ClientWidth     =   5610
   Icon            =   "SDI015.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3075
   ScaleWidth      =   5610
   Begin VB.CommandButton CmdEnable 
      Caption         =   "&Enable"
      Height          =   300
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Remove a user from the list of registered Users."
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "UserName"
      DataSource      =   "datDataCtl"
      Height          =   285
      Index           =   3
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton cmdRevok 
      Caption         =   "&Revoke"
      Height          =   300
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Remove a user from the list of registered Users."
      Top             =   2640
      Width           =   975
   End
   Begin VB.ComboBox dbc 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select User Access Level from List"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Saves the changes made."
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   4560
      TabIndex        =   8
      ToolTipText     =   "Cancels all pending updates and exits."
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password"
      DataSource      =   "datDataCtl"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2040
      MaxLength       =   32
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "UserName"
      DataSource      =   "datDataCtl"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "SystemName"
      DataSource      =   "datDataCtl"
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   0
      Top             =   40
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   0
      TabIndex        =   13
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "User's Access Level:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1225
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Login Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmSDI015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iErr As String, sMsg As String, Criteria As String
Dim rsUserLevel As New ADODB.Recordset
Dim rsUsers As New ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim OpenErr As Integer
Dim iOpenUser As Integer
Dim iOpenUserLevel As Integer
Dim i As Integer
Dim iLevel As Integer
Dim qSQL As String
Dim OldPassword As String


Function IsValid() As Integer
IsValid = True
If txtFields(0) = "" Then  ' System Name
   'msg = "Entry required for System Name"
   iErr = "UserID cannot be blank"
   txtFields(0).SetFocus
   GoTo Validate_Err
 End If
 If gblOptions = 1 Then 'Check for duplicates
    gblFileKey = UCase(Trim(txtFields(0)))
    On Error GoTo IV_ERR
    Set rsUsers = RunSP(SpCon, "usp_UserSelect", 1, gblFileKey)
    iOpenUser = True
    If Not rsUsers.EOF Then  'duplicate
       iErr = "User already exists"
       txtFields(0).SetFocus
       rsUsers.Close
       iOpenUser = 0
       GoTo Validate_Err
    End If
 End If
 If txtFields(2) = "" Then ' Password
   'msg = "Entry Required for User's Password"
   iErr = "Password cannot be blank"
   txtFields(2).SetFocus
   GoTo Validate_Err
 End If
 If dbc = "" Then ' User Level
   'msg = "Entry Required for User's Access Level"
   iErr = "Please select an access level"
   dbc.SetFocus
   GoTo Validate_Err
 End If
 If txtFields(1) = "" Then  ' User Name
   'msg = "Entry required for User's Real Name"
   iErr = "Users full name must be entered"
   txtFields(1).SetFocus
   GoTo Validate_Err
 End If
 If Len(txtFields(3)) = 0 Then
    iErr = "Please enter an NCB email address"
    txtFields(3).SetFocus
    GoTo Validate_Err
 End If
 If Not ValidEmail(txtFields(3)) Then
    iErr = "This is not a valid email address"
    txtFields(3).SetFocus
    GoTo Validate_Err
 End If
 
Validate_Exit:
   Exit Function

Validate_Err:
  MsgBox iErr, vbOKOnly, "Users"
  IsValid = False
  GoTo Validate_Exit
  
IV_ERR:
  MsgBox "SDI015/IsValid"
  cmdClose_Click

End Function

Private Sub CmdEnable_Click()
On Error GoTo Err_CmdEnable_Click
i = RunSP(SpCon, "usp_EnableAccount", 0, txtFields(0))
If i = 0 Then
   MsgBox "Account re-enabled"
Else
   MsgBox "Account was not enabled due an application error"
End If

Exit_CmdEnable_Click:
Unload Me
Exit Sub

Err_CmdEnable_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error enabling user account"
Resume Exit_CmdEnable_Click
End Sub

Private Sub cmdRevok_Click()
Dim MsG As String, iReply As Integer
On Error GoTo Revok_Err

If rsUsers!IsLoggedOn <> 0 Then
   MsG = "User is currently logged on cannot activate revoke. "
   MsG = MsG & vbCr & "Ask User to log off then revoke."
   MsgBox MsG, vbExclamation, "Revoke User"
   txtFields(1).SetFocus
   GoTo Done
End If
'--
MsG = "Please Confirm"
iReply = MsgBox(MsG, vbQuestion + vbYesNo, "Revoke User")
If iReply = vbNo Then
  txtFields(1).SetFocus
  GoTo Done
End If
iReply = RunSP(SpCon, "usp_DisableAccount", 0, UCase(Trim(txtFields(0))))
rsUsers.Close
iOpenUser = False
Set rsUsers = Nothing
Unload Me
Done:

Exit Sub
Revok_Err:
MsgBox "SDI015/Revoke"
cmdClose_Click
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo CUC_Err
Dim iReply As Integer
Dim Hash As New MD5Hash
Dim bytBlock() As Byte

If OldPassword <> txtFields(2) Then
   bytBlock = StrConv(txtFields(2).Text, vbFromUnicode)
   OldPassword = Hash.HashBytes(bytBlock)
End If
 If IsValid() Then
    If gblOptions = 1 Then
       iReply = RunSP(SpCon, "usp_UserUpdate", 0, 1, UCase(Trim(txtFields(0))), Trim(txtFields(1)), UCase(Trim(OldPassword)), dbc.ListIndex + 1, txtFields(3))
    Else
       iReply = RunSP(SpCon, "usp_UserUpdate", 0, 2, UCase(Trim(txtFields(0))), Trim(txtFields(1)), UCase(Trim(OldPassword)), dbc.ListIndex + 1, txtFields(3))
    End If
    rsUsers.Close
    iOpenUser = False
    If gblOptions = 1 Then
      ClearScreen
      txtFields(0).SetFocus
    Else
      Set rsUsers = Nothing
      Unload Me
    End If
 End If
cmdUpdateClick_Exit:
 Exit Sub
CUC_Err:
  MsgBox "SDI015/cmdUpdate"
  cmdClose_Click
End Sub

Private Sub cmdClose_Click()
  If iOpenUser = True Then rsUsers.Close
  Set rsUsers = Nothing
  Set rsUserLevel = Nothing
  Unload Me
  Set frmSDI015 = Nothing
End Sub

Private Sub Form_Activate()
If OpenErr = True Then
  If iOpenUserLevel = True Then rsUserLevel.Close
  If iOpenUser = True Then rsUsers.Close
  Set rsUsers = Nothing
  Unload Me
End If
End Sub

Private Sub Form_Load()
Dim indx As Integer
OpenErr = False
iOpenUser = False
iOpenUserLevel = False
'--
  ' On Error GoTo FL_ERR
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
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

   Set rsUserLevel = RunSP(SpCon, "usp_SelectUserLevels", 1)
   iOpenUserLevel = True
   '-----------------------------------
   '-- load combo with access levels---
   '-----------------------------------
   With rsUserLevel
       iLevel = 0 'initialize counter
      Do While Not .EOF
         dbc.AddItem !Description
         dbc.ItemData(dbc.NewIndex) = !userlevel
         iLevel = iLevel + 1
         .MoveNext
      Loop
      .Close
      iOpenUserLevel = False
   End With
   '----------------------------------------
   ' create SQL for selecting record to edit
   '----------------------------------------
    '---
 OldPassword = ""
 If gblOptions = 1 Then cmdRevok.Visible = False
 If gblUserLevel > 1 Then
      cmdRevok.Visible = False
      lblLabels(1).Visible = False
      txtFields(1).Enabled = False
      dbc.Visible = False
 End If
 If gblOptions = 2 Then
   Set rsUsers = RunSP(SpCon, "usp_UserSelect", 1, gblFileKey)
   iOpenUser = True
   Me.Caption = "Edit Registered User"
   txtFields(0).Enabled = False
   With rsUsers
     txtFields(0).Text = !SystemName
     txtFields(1).Text = !UserName
     txtFields(2).Text = !PWord
     OldPassword = !PWord
     txtFields(3).Text = IsNullMove(!Email)
     For indx = 1 To iLevel
       If dbc.ItemData(indx - 1) = !userlevel Then
          dbc.ListIndex = indx - 1
          Exit For
       End If
     Next
   End With
 End If
 '--
 ' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 
 If gblOptions = 1 Then
   ClearScreen
   Me.Caption = "New User"
 End If
 csvCenterForm Me, gblMDIFORM
  
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "frmSDI015/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
  
 End Sub
Private Sub ClearScreen()
Dim i As Integer
For i = 0 To 2
  txtFields(i) = ""
Next

End Sub

