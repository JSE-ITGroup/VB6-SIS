VERSION 5.00
Begin VB.Form SDILogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NCB - Registrar (SIS)  Login Prompt"
   ClientHeight    =   3945
   ClientLeft      =   2265
   ClientTop       =   2325
   ClientWidth     =   5655
   ControlBox      =   0   'False
   Icon            =   "SDILogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "SDILogin.frx":030A
   ScaleHeight     =   3945
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbDbase 
      Height          =   315
      ItemData        =   "SDILogin.frx":159A
      Left            =   2160
      List            =   "SDILogin.frx":159C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "Click the down arrow to see a list of available Registers"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   0
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   0
      ToolTipText     =   "Enter a user Id to identify yourself."
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtField 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2160
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter a valid pass word that is associated with your user Id."
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton btnCancel 
      BackColor       =   &H00C0E0FF&
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
      Height          =   450
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1125
   End
   Begin VB.CommandButton btnOk 
      BackColor       =   &H00C0FFFF&
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
      Height          =   450
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Filled in at run time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Register:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "SDILogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
 Unload Me
 End
End Sub

Private Function IsValid() As Integer
'--
Dim msg As String
On Error GoTo IsValid_Err
   IsValid = False
   '--
   'If no LoginName or Password was entered then beep and exit.
   '--
   If IsNothing(txtField(0).Text) Then
        msg = "Please enter a Login Name"
        txtField(0).SetFocus
        GoTo IsValid_Err
    End If
    '--
    If IsNothing(txtField(1).Text) Then
        msg = "Please enter a Password"
        txtField(1).SetFocus
        GoTo IsValid_Err
    End If
    '--
     If IsNothing(txtField(2).Text) Then
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
Dim Conn As ADODB.Connection
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim i As Integer, UserExists As Integer, isloggedon As Integer
Dim n As Integer
Dim sMsg1 As String, sMsg2 As String, msg As String
Dim qSQL As String

If Not IsValid() Then GoTo Done
    ''' This connection strings needs adjustment to work for all registers
    '''gblFileName = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=matrixx;Initial Catalog=SISNCBJ"
    gblFileName = "DSN=" & txtField(2)
    Set Conn = New ADODB.Connection
    Conn.ConnectionString = gblFileName
    Conn.Open , , , adAsyncConnect
    Do While Conn.State = adStateConnecting
       Screen.MousePointer = vbHourglass
    Loop
    Screen.MousePointer = vbDefault
    Set cnn = New ADODB.Connection
    cnn.Open gblFileName
    '''Set rs = New ADODB.Recordset
    sMsg1 = "Can't Open USERS file"
    On Error GoTo TableErrorHandler
    Set rs = RunSP(cnn, "usp_Login", 1, txtField(0))
    '****qSQL = "SELECT * FROM USers WHERE SYSTEMNAME = '" & txtField(0) & "'"
    '****rs.Open qSQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    
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
            If UCase(!Password) = UCase(txtField(1)) Then
                gblLoginName = Trim(UCase(txtField(0)))
                gblPassword = Trim(UCase(txtField(1)))
                gblUserLevel = ![userlevel]
                isloggedon = ![isloggedon]
                If isloggedon = True Then
                   n = MsgBox("Another user appears to be logged on under that name and password." & Chr(13) & Chr(10) & "This may be the result of your having earlier exited vbSIS abnormally." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Are you sure there are no other users currently logged on" & Chr(13) & Chr(10) & "with that name and password?", 36, "Login")
                       If n <> 6 Then
                        txtField(0).SetFocus
                        frmMDI.Enabled = False
                        Exit Sub
                   End If
                   
                End If
                n = RunSP(cnn, "usp_Logged", 0, 0, gblLoginName)
            frmMDI.Enabled = True
            gblOpenComp = "O"
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
    Set rs = RunSP(Conn, "usp_Company", 1)
    '''rs.Open "Company", Conn, , , adCmdTable
    On Error GoTo ReadErrorHandler
    '''rs.MoveFirst
    If Not rs.EOF Then
       gblCompName = rs!compname
       gblReadyMsg = rs!compname & " is ready"
       gblSiteId = rs!siteid
    End If
    rs.Close
    Set rs = Nothing
    Me.Visible = False
    Conn.Close
    cnn.Close
    'End If
Done:
Exit Sub
Entry_Error:
    MsgBox msg, vbCritical, "Login"
    Exit Sub
TableErrorHandler:
    MsgBox Err & " " & Err.Description & vbCrLf & sMsg1, vbCritical, "Login"
    Resume Done
    
ReadErrorHandler:
    MsgBox sMsg2, vbCritical, "Login"
    Resume Done
End Sub

Private Sub Form_Load()
Dim msg As String
Dim SpCon As ADODB.Connection
Dim adoRst As ADODB.Recordset

lblMessage.Caption = ""
msg = "Your Login Name may be different from your real name."
msg = msg & vbCrLf & "It can be any combination of up to 10 characters"
msg = msg & vbCrLf & "Your Password may be up to 10 characters."
msg = msg & vbCrLf & "Neither your Login Name"
msg = msg & " or Password are case sensitive." & vbCrLf
''msg = msg & "If this is your first logon, use 'Admin' as your Login Name "
''msg = msg & "and also as your Password."
lblMessage.Caption = msg
gblFileName = "DSN=SIS"
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
Loop
Screen.MousePointer = vbDefault
Set adoRst = RunSP(SpCon, "usp_Registers", 1)
If adoRst.EOF Then
   MsgBox "There are no registers defined. Please inform your Sys Ad"
   GoTo Exit_Form_Load
End If
Do While Not adoRst.EOF
   CmbDbase.AddItem adoRst(1)
   adoRst.MoveNext
Loop

Exit_Form_Load:
Exit Sub

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


