VERSION 5.00
Begin VB.Form FrmRTGSThreshold 
   BackColor       =   &H008080FF&
   Caption         =   "Maintain RTGS Threshold"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4590
   Icon            =   "FrmRTGSThreshold.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4590
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
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
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
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox TxtThreshold 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "Threshold"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmRTGSThreshold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
On Error GoTo Err_CmdExit_Click

Unload Me

Exit_cmdExit_Click:
Exit Sub

Err_CmdExit_Click:
MsgBox Err.Description, vbOKOnly, "Returned Cheques Exit"
GoTo Exit_cmdExit_Click
End Sub

Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_GetRTGSThreshold", 1)
If Not adoRst.EOF Then
   TxtThreshold = Format(adoRst!Amount, "#,###.00")
End If

Exit_Form_Activate:
Exit Sub

Err_Form_Activate:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on retrieving RTGS threshold"
Resume Exit_Form_Activate

End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load

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
Loop
Screen.MousePointer = vbDefault
  
   '-------------------------------------
   '-- Initialize License Details -------
   '-------------------------------------
   '--
 '--
Exit_Form_Load::
Exit Sub

Err_Form_Load:
MsgBox Err.Description, vbOKOnly, "RTGS threshold Form Load error"

End Sub
Private Sub CmdSave_Click()
On Error GoTo Err_CmdSave_Click
Dim i As Integer

If Len(TxtThreshold) < 1 Then
   MsgBox "Please enter an amount"
   TxtThreshold.SetFocus
   GoTo Exit_CmdSave_Click
End If

If Not IsNumeric(TxtThreshold) Then
   MsgBox "Only numbers are allowed"
   TxtThreshold.SetFocus
   GoTo Exit_CmdSave_Click
End If

i = RunSP(SpCon, "usp_UpdateRTGSThreshold", 0, CCur(TxtThreshold))
If i = 0 Then
   MsgBox "RTGS threshold updated"
Else
   MsgBox "RTGS update failed"
End If

Exit_CmdSave_Click:
Exit Sub

Err_CmdSave_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on saving RTGS threshold"
Resume Exit_CmdSave_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set FrmRTGSThreshold = Nothing

End Sub
