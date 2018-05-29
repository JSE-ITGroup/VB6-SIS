VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3825
   ClientLeft      =   2265
   ClientTop       =   2325
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   5475
   Begin VB.CommandButton btnOk 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   328
      Left            =   2760
      TabIndex        =   3
      Top             =   3480
      Width           =   1365
   End
   Begin VB.TextBox txtField 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1920
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Select the data source name assigned to the application. "
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   328
      Left            =   4200
      TabIndex        =   4
      Top             =   3480
      Width           =   1365
   End
   Begin VB.TextBox txtField 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1920
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter a valid pass word that is associated with your user Id."
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   0
      ToolTipText     =   "Enter a user Id to identify yourself."
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data Source Name:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Filled in at run time"
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label Label2 
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
      Left            =   840
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Dim i As Integer
 Unload Me
 i = LogOff()
End Sub

Private Function IsValid() As Integer
'--
Dim msg As String
On Error GoTo IsValid_Err
   IsValid = False
   '--
   'If no LoginName or Password was entered then beep and exit.
   '--
   If Len(txtField(0).Text) = 0 Then
        msg = "Please enter a Login Name"
        txtField(0).SetFocus
        GoTo IsValid_Err
    End If
    '--
    If Len(txtField(1).Text) = 0 Then
        msg = "Please enter a Password"
        txtField(1).SetFocus
        GoTo IsValid_Err
    End If
    '--
     If Len(txtField(2).Text) = 0 Then
        msg = "Please enter a Data Source Name"
        txtField(2).SetFocus
        GoTo IsValid_Err
    End If
    IsValid = True
IsValid_Exit:
    Exit Function
IsValid_Err:
    MsgBox msg, vbCritical, "Login"
    GoTo IsValid_Exit
End Function
Private Sub btnOk_Click()
Dim cnn As New cADOAccess
Dim rs As ADODB.Recordset
Dim i As Integer, UserExists As Integer, isloggedon As Integer
Dim n As Integer
Dim sMsg1 As String, sMsg2 As String, msg As String
Dim qSQL As String

If Not IsValid() Then GoTo Done
'-- Open connection
'------------------
cnn.UserId = txtField(0)
cnn.Password = txtField(1)
cnn.DSN = txtField(2)
'gblDSN = txtField(2)
cnn.SDICnn  ' call connection routine
'-- test if connection successful
'--------
If cnn.Reply = 0 Then 'open connection failed
    txtField(0).SetFocus
Else
    '--connect successfull
    '---------------------
    gblFileName = cnn.OpenData
    qSQL = "SELECT * from USERS "
    qSQL = qSQL & "where SYSTEMNAME = '" & UCase(cnn.UserId) & "'"
    Set rs = New ADODB.Recordset
    sMsg1 = "Can't Open USERS file"
    On Error GoTo TableErrorHandler
    rs.Open qSQL, gblFileName, adOpenKeyset, adLockOptimistic, adCmdText
    ' See if the user exists.
    '--
    UserExists = True
    With rs
        If .EOF And .BOF Then
          UserExists = False
        End If
        ' If the user exists, then validate the password.
        '--
        If (UserExists) Then
            If UCase(![Password]) = UCase(cnn.Password) Then
                gblLoginName = Trim(UCase(cnn.UserId))
                gblPassword = Trim(UCase(cnn.Password))
                gblUserLevel = ![userlevel]
                isloggedon = ![isloggedon]
                If isloggedon = True Then
                   n = MsgBox("Another user appears to be logged on under that name and password." & Chr(13) & Chr(10) & "This may be the result of your having earlier exited vbSIS abnormally." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Are you sure there are no other users currently logged on" & Chr(13) & Chr(10) & "with that name and password?", 36, "Login")
                   If n <> 6 Then
                        txtField(0).SetFocus
                        Exit Sub
                   End If
                End If
                ![isloggedon] = -1
               .Update
            Else
                msg = "Invalid Password"
                txtField(1).SetFocus
                GoTo Entry_Error
            End If
        Else

        ' The user does not exist.
         Screen.MousePointer = 0
        msg = "Invalid Login Name."
        txtField(0).SetFocus
        GoTo Entry_Error
      End If
      .Close
    End With
    ' Extract Company Name from Company File
    '--
    sMsg1 = "Can't open Company file."
    sMsg2 = "Can't read Company Record."
    On Error GoTo TableErrorHandler
    rs.Open "Company", gblFileName, , , adCmdTable
    On Error GoTo ReadErrorHandler
    rs.MoveFirst
    If Not rs.EOF Then
       gblCompName = rs!compname
    End If
    rs.Close
    sMsg1 = "Can't open Configuration file"
    On Error GoTo TableErrorHandler
    qSQL = "SELECT value from CONFIG where ID = 1"
    rs.Open qSQL, gblFileName, , , adCmdText
    If Not rs.EOF Then
      gblVersn = rs!Value
    End If
    rs.Close
    Set rs = Nothing
    Me.Visible = False
  
End If
Done:
Exit Sub
Entry_Error:
    MsgBox msg, vbCritical, "Login"
    Exit Sub
TableErrorHandler:
    MsgBox sMsg1, vbCritical, "Login"
    Resume Done
    
ReadErrorHandler:
    MsgBox sMsg2, vbCritical, "Login"
    Resume Done
End Sub

Private Sub Form_Load()
Dim msg As String
lblMessage.Caption = ""
msg = "Your Login Name may be different from your real name."
msg = msg & vbCrLf & "It can be any combination of up to 10 characters"
msg = msg & vbCrLf & "Your Password may be up to 10 characters."
msg = msg & vbCrLf & "Neither your Login Name"
msg = msg & " or Password are case sensitive." & vbCrLf
msg = msg & "If this is your first logon, use 'Admin' as your Login Name "
msg = msg & "and also as your Password."
lblMessage.Caption = msg
End Sub


Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
Case vbKeyReturn, vbKeyDown
   KeyAscii = 0
   If Index = 2 Then
      btnOk.SetFocus
   Else
      txtField(Index + 1).SetFocus
   End If
Case vbKeyUp
  If Index > 0 Then
    txtField(Index - 1).SetFocus
  End If
End Select
End Sub


