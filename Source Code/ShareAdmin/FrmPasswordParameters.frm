VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form FrmPasswordParameters 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Password Parameters"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   Icon            =   "FrmPasswordParameters.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   11
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox TxtDays 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox TxtNoOfPasswords 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox TxtLength 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin SSDataWidgets_A.SSDBOptSet OptUpperCase 
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Select the required tax option for the category."
      Top             =   600
      Width           =   1950
      _Version        =   196611
      _ExtentX        =   3440
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "&No"
      BackColor       =   16777215
      BevelColorFace  =   12648447
      Cols            =   2
      IndexSelected   =   0
      NumberOfButtons =   2
      Buttons.Button(0).OptionValue=   "0"
      Buttons.Button(0).Caption=   "&No"
      Buttons.Button(0).Mnemonic=   78
      Buttons.Button(0).Value=   -1  'True
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   29
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   31
      Buttons.Button(0).PictureRight=   30
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   64
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
      Buttons.Button(1).Caption=   "&Yes"
      Buttons.Button(1).Mnemonic=   89
      Buttons.Button(1).TextLeft=   80
      Buttons.Button(1).TextRight=   98
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   65
      Buttons.Button(1).ButtonRight=   78
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   100
      Buttons.Button(1).PictureRight=   99
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   65
      Buttons.Button(1).ButtonToColRight=   129
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).Column=   1
   End
   Begin SSDataWidgets_A.SSDBOptSet OptLowerCase 
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      ToolTipText     =   "Select the required tax option for the category."
      Top             =   1080
      Width           =   1950
      _Version        =   196611
      _ExtentX        =   3440
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No"
      BackColor       =   -2147483643
      BevelColorFace  =   12648447
      Cols            =   2
      IndexSelected   =   0
      NumberOfButtons =   2
      Buttons.Button(0).Caption=   "No"
      Buttons.Button(0).Mnemonic=   78
      Buttons.Button(0).Value=   -1  'True
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   29
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   31
      Buttons.Button(0).PictureRight=   30
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   64
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
      Buttons.Button(1).Caption=   "Yes"
      Buttons.Button(1).Mnemonic=   89
      Buttons.Button(1).TextLeft=   80
      Buttons.Button(1).TextRight=   98
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   65
      Buttons.Button(1).ButtonRight=   78
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   100
      Buttons.Button(1).PictureRight=   99
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   65
      Buttons.Button(1).ButtonToColRight=   129
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).Column=   1
   End
   Begin SSDataWidgets_A.SSDBOptSet OptSpecialChar 
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      ToolTipText     =   "Select the required tax option for the category."
      Top             =   1560
      Width           =   1950
      _Version        =   196611
      _ExtentX        =   3440
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No"
      BackColor       =   -2147483643
      BevelColorFace  =   12648447
      Cols            =   2
      IndexSelected   =   0
      NumberOfButtons =   2
      Buttons.Button(0).Caption=   "No"
      Buttons.Button(0).Mnemonic=   78
      Buttons.Button(0).Value=   -1  'True
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   29
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   31
      Buttons.Button(0).PictureRight=   30
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   64
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
      Buttons.Button(1).Caption=   "Yes"
      Buttons.Button(1).Mnemonic=   89
      Buttons.Button(1).TextLeft=   80
      Buttons.Button(1).TextRight=   98
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   65
      Buttons.Button(1).ButtonRight=   78
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   100
      Buttons.Button(1).PictureRight=   99
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   65
      Buttons.Button(1).ButtonToColRight=   129
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).Column=   1
   End
   Begin SSDataWidgets_A.SSDBOptSet OptNumbers 
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      ToolTipText     =   "Select the required tax option for the category."
      Top             =   2040
      Width           =   1950
      _Version        =   196611
      _ExtentX        =   3440
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No"
      BackColor       =   -2147483643
      BevelColorFace  =   12648447
      Cols            =   2
      IndexSelected   =   0
      NumberOfButtons =   2
      Buttons.Button(0).Caption=   "No"
      Buttons.Button(0).Mnemonic=   78
      Buttons.Button(0).Value=   -1  'True
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   29
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   31
      Buttons.Button(0).PictureRight=   30
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   64
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
      Buttons.Button(1).Caption=   "Yes"
      Buttons.Button(1).Mnemonic=   89
      Buttons.Button(1).TextLeft=   80
      Buttons.Button(1).TextRight=   98
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   65
      Buttons.Button(1).ButtonRight=   78
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   100
      Buttons.Button(1).PictureRight=   99
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   65
      Buttons.Button(1).ButtonToColRight=   129
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).Column=   1
   End
   Begin SSDataWidgets_A.SSDBOptSet OptUserID 
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      ToolTipText     =   "Select the required tax option for the category."
      Top             =   2640
      Width           =   1950
      _Version        =   196611
      _ExtentX        =   3440
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No"
      BackColor       =   -2147483643
      BevelColorFace  =   12648447
      Cols            =   2
      IndexSelected   =   0
      NumberOfButtons =   2
      Buttons.Button(0).Caption=   "No"
      Buttons.Button(0).Mnemonic=   78
      Buttons.Button(0).Value=   -1  'True
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   29
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   31
      Buttons.Button(0).PictureRight=   30
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   64
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
      Buttons.Button(1).Caption=   "Yes"
      Buttons.Button(1).Mnemonic=   89
      Buttons.Button(1).TextLeft=   80
      Buttons.Button(1).TextRight=   98
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   65
      Buttons.Button(1).ButtonRight=   78
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   100
      Buttons.Button(1).PictureRight=   99
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   65
      Buttons.Button(1).ButtonToColRight=   129
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).Column=   1
   End
   Begin SSDataWidgets_A.SSDBOptSet OptFullName 
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      ToolTipText     =   "Select the required tax option for the category."
      Top             =   3240
      Width           =   1950
      _Version        =   196611
      _ExtentX        =   3440
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No"
      BackColor       =   -2147483643
      BevelColorFace  =   12648447
      Cols            =   2
      IndexSelected   =   0
      NumberOfButtons =   2
      Buttons.Button(0).Caption=   "No"
      Buttons.Button(0).Mnemonic=   78
      Buttons.Button(0).Value=   -1  'True
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   29
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   31
      Buttons.Button(0).PictureRight=   30
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   64
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
      Buttons.Button(1).Caption=   "Yes"
      Buttons.Button(1).Mnemonic=   89
      Buttons.Button(1).TextLeft=   80
      Buttons.Button(1).TextRight=   98
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   65
      Buttons.Button(1).ButtonRight=   78
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   100
      Buttons.Button(1).PictureRight=   99
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   65
      Buttons.Button(1).ButtonToColRight=   129
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).Column=   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Days password remainds active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "No of passwords before reuse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Do not allow user's full name as password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Do not allow UserID as password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Insist on Numbers?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Insist on Special Characters?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Insist on Lowercase letters?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Insist on Uppercase letters?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Minimum length of password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FrmPasswordParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
On Error GoTo Err_CmdSave_Click
Dim StrSql As String
Dim i As Integer

If Len(TxtLength) < 1 Then
   StrSql = "You are required to enter a valid minimum length for a password before clicking SAVE" & vbCrLf
   StrSql = StrSql & "Please do so before continuing"
   MsgBox StrSql, vbOKOnly, "Minimum length was not entered"
   TxtLength.SetFocus
   GoTo Exit_CmdSave_Click
End If

If Not IsNumber(TxtLength) Then
   StrSql = "You are required to enter a valid number for minimum password length before clicking SAVE" & vbCrLf
   StrSql = StrSql & "Please correct"
   MsgBox StrSql, vbOKOnly, "Minimum Password length is not a number"
   TxtLength.SetFocus
   GoTo Exit_CmdSave_Click
End If

If TxtLength = "0" Then
   StrSql = "You are required to enter a valid number for minimum password length before clicking SAVE" & vbCrLf
   StrSql = StrSql & "Please correct"
   MsgBox StrSql, vbOKOnly, "Minimum Password length cannot be 0"
   TxtLength.SetFocus
   GoTo Exit_CmdSave_Click
End If

If Len(TxtNoOfPasswords) < 1 Then
   StrSql = "You are required to enter a valid No. of Passwords before clicking SAVE" & vbCrLf
   StrSql = StrSql & "Please do so before continuing"
   MsgBox StrSql, vbOKOnly, "No. Of Passwords before reuse was not entered"
   TxtNoOfPasswords.SetFocus
   GoTo Exit_CmdSave_Click
End If

If Not IsNumber(TxtNoOfPasswords) Then
   StrSql = "You are required to enter a valid number for No. of Passwords before clicking SAVE" & vbCrLf
   StrSql = StrSql & "Please correct"
   MsgBox StrSql, vbOKOnly, "No. of Passwords must be a number"
   TxtNoOfPasswords.SetFocus
   GoTo Exit_CmdSave_Click
End If

If TxtNoOfPasswords = "0" Then
   StrSql = "You are required to enter a valid number for No. of Passwords before clicking SAVE" & vbCrLf
   StrSql = StrSql & "Please correct"
   MsgBox StrSql, vbOKOnly, "No. Of Passwords cannot be 0"
   TxtNoOfPasswords.SetFocus
   GoTo Exit_CmdSave_Click
End If

If Len(TxtDays) < 1 Then
   StrSql = "You are required to enter a valid duration in days for passwords before clicking SAVE" & vbCrLf
   StrSql = StrSql & "Please do so before continuing"
   MsgBox StrSql, vbOKOnly, "Password duration was not entered"
   TxtDays.SetFocus
   GoTo Exit_CmdSave_Click
End If

If Not IsNumber(TxtDays) Then
   StrSql = "You are required to enter a valid number for Password duration beore clicking SAVE" & vbCrLf
   StrSql = StrSql & "Please correct"
   MsgBox StrSql, vbOKOnly, "Password duration must be a number in days"
   TxtDays.SetFocus
   GoTo Exit_CmdSave_Click
End If

If TxtDays = "0" Then
   StrSql = "You are required to enter a valid number for Password duration beore clicking SAVE" & vbCrLf
   StrSql = StrSql & "Please correct"
   MsgBox StrSql, vbOKOnly, "Password duration cannot be 0"
   TxtDays.SetFocus
   GoTo Exit_CmdSave_Click
End If

i = RunSP(SpCon, "usp_UpdatePasswordParameters", 0, CInt(TxtLength), OptUpperCase.IndexSelected, OptLowerCase.IndexSelected, OptSpecialChar.IndexSelected, _
    OptNumbers.IndexSelected, CInt(TxtNoOfPasswords), CInt(TxtDays), OptUserID.IndexSelected, OptFullName.IndexSelected, gblLoginName)
If i <> 0 Then
   StrSql = "An error occurred while updating the Password Parameters." & vbCrLf
   StrSql = StrSql & "Please contact your Systems Administrator"
   MsgBox StrSql, vbOKOnly, "Error on Password Parameter update"
   GoTo Exit_CmdSave_Click
Else
   MsgBox "Password Parameters successfully saved", vbOKOnly, "Save Completed"
End If

Exit_CmdSave_Click:
Exit Sub

Err_CmdSave_Click:
MsgBox Err.Description, vbOKOnly, "Error trying to save password parameters"
GoTo Exit_CmdSave_Click
End Sub

Private Sub Form_Activate()
Dim adoRst As New ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_GetPasswordParameters", 1)
With adoRst
     If .EOF Then
        TxtLength = 8
        TxtNoOfPasswords = 1
        TxtDays = 90
        OptUpperCase.IndexSelected = 0
        OptLowerCase.IndexSelected = 0
        OptSpecialChar.IndexSelected = 0
        OptNumbers.IndexSelected = 0
        OptUserID.IndexSelected = 0
        OptFullName.IndexSelected = 0
     Else
         TxtLength = !PasswordLength
         TxtNoOfPasswords = !NoOfPasswords
         TxtDays = !PasswordPeriod
         If !UpperCase = True Then OptUpperCase.IndexSelected = 1
         If !LowerCase = True Then OptLowerCase.IndexSelected = 1
         If !SpecialChr = True Then OptSpecialChar.IndexSelected = 1
         If !NumberChr = True Then OptNumbers.IndexSelected = 1
         If !ProhibitUserID = True Then OptUserID.IndexSelected = 1
         If !ProhibitName = True Then OptFullName.IndexSelected = 1
     End If
End With

adoRst.Close
Set adoRst = Nothing

Exit_Form_Activate:
Exit Sub
End Sub

Private Sub Form_Load()
csvCenterForm Me, gblMDIFORM
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
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
frmMDI.txtStatusMsg.Refresh

End Sub
