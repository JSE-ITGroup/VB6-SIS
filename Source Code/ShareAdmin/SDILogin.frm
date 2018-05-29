VERSION 5.00
Begin VB.Form SDILogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "NCB - Registrar (SIS)  Login Prompt"
   ClientHeight    =   3945
   ClientLeft      =   2220
   ClientTop       =   1890
   ClientWidth     =   5655
   ControlBox      =   0   'False
   Icon            =   "SDILogin.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "SDILogin.frx":030A
   ScaleHeight     =   3945
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   4
      Top             =   3360
      Width           =   1125
   End
   Begin VB.TextBox txtField 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2160
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Select the data source name assigned to the application. "
      Top             =   1200
      Width           =   2175
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
      TabIndex        =   3
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Label LblForgetPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Forget Password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   240
      MouseIcon       =   "SDILogin.frx":2AA1
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Source Name:"
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
      TabIndex        =   5
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
'On Error GoTo LoginErrorHandler
Dim Conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim i As Integer, IsLoggedOn As Integer
Dim n As Integer
Dim sMsg1 As String
Dim qSQL As String
Dim Hash As New MD5Hash
Dim bytBlock() As Byte
Dim TodaysDate As Date

If Not IsValid() Then GoTo Done
gblFileName = "DSN=" & txtField(2)
Set Conn = New ADODB.Connection
Conn.ConnectionString = gblFileName
Conn.Open , , , adAsyncConnect
Do While Conn.State = adStateConnecting
   Screen.MousePointer = vbHourglass
Loop
Screen.MousePointer = vbDefault
bytBlock = StrConv(txtField(1).Text, vbFromUnicode)
qSQL = Hash.HashBytes(bytBlock)
Set rs = RunSP(Conn, "usp_Login", 1, txtField(0).Text, qSQL)
If rs.State = adStateClosed Then
   MsgBox "UserID does not exist or is disabled"
   GoTo Done
End If
If rs.EOF Then
   MsgBox "UserID does not exist or is disabled"
   rs.Close
   GoTo Done
End If

With rs
     If !LoginFails = 3 Then
        sMsg1 = "You have exhausted the number of tries. Account will be disabled"
        MsgBox sMsg1, vbOKOnly
        i = RunSP(Conn, "usp_DisableAccount", 0, txtField(0))
        GoTo Done
     End If
     If !IsLoggedOn = 1 Then
        sMsg1 = "Sorry, the system is unable to log you in." & vbCrLf
        sMsg1 = sMsg1 & "This indicates that the login name is invalid or disabled" & vbCrLf
        sMsg1 = sMsg1 & "or that the password is incorrect in this database"
        MsgBox sMsg1, vbOKOnly
        GoTo Done
     End If
    
     If !IsLoggedOn = 2 Then
        sMsg1 = "Sorry, you are already logged in" & vbCrLf
        sMsg1 = sMsg1 & "Log out at the other location and then try logging in again"
        MsgBox sMsg1, vbOKOnly
        GoTo Done
     End If
     TodaysDate = Date
     If !PassWDExpiry <= TodaysDate Then
        frmChangePassword.txtUserName = Trim(UCase(txtField(0)))
        gblUserName = !UserName
        frmChangePassword.Show vbModal
        If gblYesNo = False Then
           GoTo Done
        Else
           GoTo ContinueRegular
        End If
     End If
     i = !PassWDExpiry - TodaysDate
     If i < 10 Then
        sMsg1 = "Your password will expire in " & i & " days" & vbCrLf
        sMsg1 = sMsg1 & "You can change it at anytime under the File --> Utilities Option"
        MsgBox sMsg1, vbOKOnly
     End If
ContinueRegular:
     gblLoginName = Trim(UCase(txtField(0)))
     gblPassword = !PWord
     gblUserLevel = !userlevel
     IsLoggedOn = !IsLoggedOn
     gblUserName = !UserName
     n = RunSP(Conn, "usp_Logged", 0, 0, gblLoginName)
     frmMDI.Enabled = True
     gblOpenComp = "O"
     .Close
End With
Set rs = RunSP(Conn, "usp_GetCompanyBasics", 1)
If Not rs.EOF Then
   gblCompName = rs!compname
   gblReadyMsg = rs!compname & " is ready"
   gblSiteId = rs!siteid
End If
rs.Close
Set rs = Nothing
Me.Visible = False
Conn.Close

Done:
Exit Sub
LoginErrorHandler:
    MsgBox Err & " " & Err.Description & vbCrLf & sMsg1, vbCritical, "Login"
    Resume Done
End Sub

Private Sub Form_Load()
Dim msg As String
lblMessage.Caption = ""
msg = "Your Password is case sensitive. Please bear this in mind as you enter the password."

lblMessage.Caption = msg
End Sub

Private Sub LblForgetPassword_Click()
Dim adoUserEmail As ADODB.Recordset
Dim iString As String
Dim i As Integer
Dim iText As String
Dim ToAddress As String
Dim Hash As New MD5Hash
Dim bytBlock() As Byte

If Len(txtField(2)) < 1 Then
   MsgBox "Enter a data source first"
   GoTo Exit_LblForgetPassword_Click
End If
If Len(txtField(0)) < 1 Then
   MsgBox "Enter your userID"
   GoTo Exit_LblForgetPassword_Click
End If

gblFileName = "DSN=" & txtField(2)
Set cnn = New ADODB.Connection
cnn.ConnectionString = gblFileName
cnn.Open , , , adAsyncConnect
Do While cnn.State = adStateConnecting
   Screen.MousePointer = vbHourglass
Loop
Screen.MousePointer = vbDefault
iString = RandomPassword

Set adoUserEmail = RunSP(cnn, "usp_GetEmailAddress", 1, txtField(0))
With adoUserEmail
     If .State = adStateClosed Then
        GoTo Exit_LblForgetPassword_Click
     End If
     ToAddress = Chr(34) & !UserName & Chr(34) & "<" & !Email & ">"
     iText = "Your temporary password is " & iString & " . Please log in and reset immediately."
     i = SendEmail(cnn, "scliffordallen@gmail.com", ToAddress, "SIS password reset", iText)
     If i = True Then
        bytBlock = StrConv(iString, vbFromUnicode)
        iString = Hash.HashBytes(bytBlock)
        i = RunSP(cnn, "usp_PasswordReset", 0, txtField(0), iString)
        If i = 0 Then
           MsgBox "Please cheeck your email for your new temporary password"
        End If
     End If
    .Close
End With

Set adoUserEmail = Nothing
'cnn.Close

Exit_LblForgetPassword_Click:
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

Function RandomPassword()
Dim All_Chars() As Variant
Dim i As Integer
Dim iPassword As String
Dim Random_Index As Integer

All_Chars = Array("A", "a", "b", "B", "C", "c", "D", "d", "E", "e", "F", "f", 0, "G", "g", 1, "H", "h", 2, "i", 3, "I", "j", "J", "k", "K", 4, 5, "L", "l", "M", "m", "n", "N", "O", "o", "p", "P", 6, "q", "Q", "R", "r", "s", "S", 7, "t", "T", "U", "u", "V", "v", 8, "W", "w", "X", "x", "Y", "y", "Z", "z", 9)
Randomize
For i = 1 To 10
Random_Index = Int(Rnd() * 25)
iPassword = iPassword & All_Chars(Random_Index)
Next
RandomPassword = iPassword
End Function

