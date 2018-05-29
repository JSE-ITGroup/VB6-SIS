VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2550
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5640
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1506.624
   ScaleMode       =   0  'User
   ScaleWidth      =   5295.655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdRequirements 
      Caption         =   "Password Requirements"
      Default         =   -1  'True
      Height          =   390
      Left            =   1800
      TabIndex        =   10
      Top             =   2040
      Width           =   2100
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2865
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1440
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2865
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   960
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Enabled         =   0   'False
      Height          =   345
      Left            =   2850
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4200
      TabIndex        =   9
      Top             =   2040
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2850
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Confirm New Password:"
      Height          =   270
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1455
      Width           =   2160
   End
   Begin VB.Label lblLabels 
      Caption         =   "&New Password:"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   975
      Width           =   2160
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Old Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As ADODB.Connection
Dim NoOfPasswords As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_cmdOK_Click
Dim i As Integer
Dim msg As String
Dim Hash As New MD5Hash
Dim bytBlock() As Byte

gblYesNo = False
If Not IsValid Then GoTo Exit_cmdOK_Click
If Not MeetsRequirements Then GoTo Exit_cmdOK_Click
If txtPassword(1) <> txtPassword(2) Then
   MsgBox "New password and Confirmation password do not match", vbCritical, "New Passwords differ"
   GoTo Exit_cmdOK_Click
End If
bytBlock = StrConv(txtPassword(1).Text, vbFromUnicode)
msg = Hash.HashBytes(bytBlock)
i = RunSP(Conn, "usp_PasswordUpdate", 0, txtUserName, msg)
If i = 0 Then
   MsgBox "Password successfully changed"
   gblYesNo = True
Else
   If i = 1 Then
      msg = "Password was previously used" & vbCrLf
      msg = msg & "You are not allowed to reused any of the last " & NoOfPasswords & " passwords"
      MsgBox msg, vbOKOnly, "Password already used"
   Else
      MsgBox Err & " " & Err.Description, vbOKOnly, "Your password was not changed. Please try again"
   End If
End If
   
Exit_cmdOK_Click:
Exit Sub

Err_cmdOK_Click:
MsgBox Err & " " & Err.Description, vbOKOnly
Resume Exit_cmdOK_Click
End Sub
Private Function IsValid() As Boolean
'--
Dim msg As String
On Error GoTo IsValid_Err
IsValid = False
If IsNothing(txtUserName.Text) Then
   msg = "Please enter a Login Name"
   txtUserName.SetFocus
   GoTo IsValid_Err
End If
    '--
If IsNothing(txtPassword(0).Text) Then
   msg = "Please enter the Old Password"
   txtPassword(0).SetFocus
   GoTo IsValid_Err
End If
If IsNothing(txtPassword(1).Text) Then
   msg = "Please enter the New Password"
   txtPassword(1).SetFocus
   GoTo IsValid_Err
End If
If IsNothing(txtPassword(2).Text) Then
   msg = "Please confirm the New Password"
   txtPassword(2).SetFocus
   GoTo IsValid_Err
End If
    '--
IsValid = True
IsValid_Exit:
    Exit Function
IsValid_Err:
    MsgBox msg, vbCritical, "Change Password"
    GoTo IsValid_Exit
End Function

Function MeetsRequirements() As Boolean
Dim adoRs As New ADODB.Recordset
Dim i As Integer
Dim msg As String

Set Conn = New ADODB.Connection
Conn.ConnectionString = gblFileName
Conn.Open , , , adAsyncConnect
Do While Conn.State = adStateConnecting
   Screen.MousePointer = vbHourglass
Loop
Screen.MousePointer = vbDefault

MeetsRequirements = False
Set adoRs = RunSP(Conn, "usp_GetPasswordParameters", 1)
If adoRs.State = adStateClosed Then
   MsgBox "Password Parameters missing"
   GoTo MeetsRequirements_Exit
End If
With adoRs
     NoOfPasswords = !NoOfPasswords
     If Len(txtPassword(1).Text) < !PasswordLength Then
        msg = "Password should be at least " & !PasswordLength & " characters long"
        MsgBox msg, vbCritical, "Password Entered is too short"
        GoTo MeetsRequirements_Exit
     End If
     If !NumberChr = True Then
        If Not HasNumber(txtPassword(1).Text) Then
           msg = "Password should have at least 1 number"
           MsgBox msg, vbCritical, "Number required in Password"
          GoTo MeetsRequirements_Exit
        End If
     End If
     If !UpperCase = True Then
        If Not HasUpper(txtPassword(1).Text) Then
           msg = "Password should have at least 1 character being uppercase"
           MsgBox msg, vbCritical, "Uppercase character required in Password"
          GoTo MeetsRequirements_Exit
        End If
     End If
     If !LowerCase = True Then
        If Not HasLower(txtPassword(1).Text) Then
           msg = "Password should have at least 1 character being lowercase"
           MsgBox msg, vbCritical, "Lowercase character required in Password"
          GoTo MeetsRequirements_Exit
        End If
     End If
     If !SpecialChr = True Then
        If Not HasSpecial(txtPassword(1).Text) Then
           msg = "Password should have at least 1 special character"
           MsgBox msg, vbCritical, "Special character required in Password"
          GoTo MeetsRequirements_Exit
        End If
     End If
     If !ProhibitUserID = True Then
        i = InStr(1, txtPassword(1).Text, txtUserName)
        If i > 0 Then
           msg = "Password should not contain your user id"
           MsgBox msg, vbCritical, "UserID not allowed in Password"
          GoTo MeetsRequirements_Exit
        End If
     End If
     If !ProhibitName = True Then
        i = InStr(1, txtPassword(1).Text, gblUserName)
        If i > 0 Then
           msg = "Password should not contain your name"
           MsgBox msg, vbCritical, "Your name is not allowed in Password"
          GoTo MeetsRequirements_Exit
        End If
     End If
End With

MeetsRequirements = True
MeetsRequirements_Exit:
If adoRs.State = adStateOpen Then
   adoRs.Close
End If
Set adoRs = Nothing

Exit Function

End Function
Function HasNumber(StringToCheck As String) As Boolean
Dim X As Integer
HasNumber = False
For X = 1 To Len(StringToCheck)
If Asc(Mid(StringToCheck, X, 1)) > 47 And Asc(Mid(StringToCheck, X, 1)) < 58 Then
   HasNumber = True
   Exit For
End If
Next X

End Function
Function HasUpper(StringToCheck As String) As Boolean
Dim X As Integer
HasUpper = False
For X = 1 To Len(StringToCheck)
If Asc(Mid(StringToCheck, X, 1)) > 64 And Asc(Mid(StringToCheck, X, 1)) < 91 Then
   HasUpper = True
   Exit For
End If
Next X

End Function
Function HasLower(StringToCheck As String) As Boolean
Dim X As Integer
HasLower = False
For X = 1 To Len(StringToCheck)
If Asc(Mid(StringToCheck, X, 1)) > 96 And Asc(Mid(StringToCheck, X, 1)) < 123 Then
   HasLower = True
   Exit For
End If
Next X

End Function
Function HasSpecial(StringToCheck As String) As Boolean
Dim X As Integer
HasSpecial = False
For X = 1 To Len(StringToCheck)
If Asc(Mid(StringToCheck, X, 1)) > 32 And Asc(Mid(StringToCheck, X, 1)) < 48 Or Asc(Mid(StringToCheck, X, 1)) > 57 And Asc(Mid(StringToCheck, X, 1)) < 65 Then
   HasSpecial = True
   Exit For
End If
Next X

End Function

Private Sub CmdRequirements_Click()
FrmPasswordDialog.Show vbModal
End Sub

Private Sub txtPassword_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Then
   If KeyAscii = 34 Or KeyAscii = 39 Or KeyAscii = 96 Then
      KeyAscii = 0
      MsgBox "That special character is not allowed"
   End If
End If
   

End Sub
