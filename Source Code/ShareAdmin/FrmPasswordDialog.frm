VERSION 5.00
Begin VB.Form FrmPasswordDialog 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Password Requirements"
   ClientHeight    =   3735
   ClientLeft      =   2715
   ClientTop       =   3315
   ClientWidth     =   8805
   Icon            =   "FrmPasswordDialog.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LblRequirements 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "FrmPasswordDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Dim Conn As ADODB.Connection
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

Set adoRs = RunSP(Conn, "usp_GetPasswordParameters", 1)
If adoRs.State = adStateClosed Then
   MsgBox "Password Parameters missing"
   Unload Me
   GoTo Exit_Form_Load
End If
With adoRs
     msg = "You are not allowed to used any of your last " & !NoOfPasswords & " passwords" & vbCrLf & vbCrLf
     msg = msg & "Your password should be at least " & !PasswordLength & " characters long" & vbCrLf & vbCrLf
     If !NumberChr = True Then
        msg = msg & "Password should have at least 1 number" & vbCrLf & vbCrLf
     End If
     If !UpperCase = True Then
        msg = msg & "Password should have at least 1 character being uppercase" & vbCrLf & vbCrLf
     End If
     If !LowerCase = True Then
        msg = msg & "Password should have at least 1 character being lowercase" & vbCrLf & vbCrLf
     End If
     If !SpecialChr = True Then
        msg = msg & "Password should have at least 1 special character" & vbCrLf & vbCrLf
     End If
     If !ProhibitUserID = True Then
        msg = msg & "Password should not contain your user id" & vbCrLf & vbCrLf
     End If
     If !ProhibitName = True Then
        msg = msg & "Password should not contain your name"
     End If
End With
LblRequirements.Caption = msg
adoRs.Close
Set adoRs = Nothing
Conn.Close
Set Conn = Nothing

Exit_Form_Load:
Exit Sub

End Sub



Private Sub OKButton_Click()
Unload Me
End Sub
