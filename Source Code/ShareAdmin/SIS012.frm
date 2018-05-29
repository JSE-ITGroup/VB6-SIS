VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSIS012 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lien on Certification"
   ClientHeight    =   3690
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "SIS012.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6930
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Enter the date the lien is placed on the certificate."
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3720
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5880
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4800
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox tbfld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Enter the name of the assignee in free format"
      Top             =   2400
      Width           =   4335
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   15
      ToolTipText     =   "Enter the date the lien is released."
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Assignment Ends:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Shareholder Name"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   13
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lbl 
      Caption         =   "Cert Number"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   12
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Assigned To:"
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
      TabIndex        =   11
      Top             =   2400
      Width           =   1740
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   9480
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6960
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
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6960
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Shareholder Name:"
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
      TabIndex        =   8
      Top             =   960
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
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cert Number:"
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
      TabIndex        =   6
      Top             =   600
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Assignment Starts:"
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
      Left            =   45
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
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
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim rsMain As ADODB.Recordset
Dim iOpenMain As Integer
Dim iOpenAdt As Integer
Dim OpenErr As Integer
Dim SpCon As ADODB.Connection
Dim strTable As String
Dim strRecNO As String
Dim iNewAcct As Long
Private Sub Shutdown()
If iOpenMain = True Then rsMain.Close
Set rsMain = Nothing
SpCon.Close
End Sub
Function IsValid() As Integer
Dim iErr As Integer
IsValid = True
'--
If meb(0).Text = "" Then 'Assignment start not entered
   iErr = 36
   meb(0).SetFocus
   GoTo Validate_Err
 End If
 '--
 If Not IsDate(meb(0).Text) Then
    iErr = 14
    meb(0).SetFocus
   GoTo Validate_Err
 End If
 '--
If tbfld(0) = "" Then  ' Assignee
   iErr = 163
   tbfld(0).SetFocus
   GoTo Validate_Err
 End If
 '--
 If gblOptions = 2 Then
   If meb(1).Text = "" Then 'Assignment end not entered
      iErr = 37
      meb(1).SetFocus
      GoTo Validate_Err
   End If
   '--
   If Not IsDate(meb(1).Text) Then
      iErr = 14
      meb(1).SetFocus
      GoTo Validate_Err
    End If
    '--
    If DateValue(meb(1).Text) < DateValue(meb(0).Text) Then
       iErr = 38
       meb(1).SetFocus
       GoTo Validate_Err
    End If
 
 End If
 '--
Validate_Exit:
   Exit Function
'--
Validate_Err:
  csvShowUsrErr iErr, "Client Account"
  IsValid = False
  GoTo Validate_Exit
'--
End Function

Private Sub cmdCancel_Click()
Shutdown
Set frmSIS012 = Nothing
Unload Me
End Sub

Private Sub cmdClear_Click()

If gblOptions = 1 Then
   ClearScreen
   tbfld(0).Text = iNewAcct
   tbfld(0).SetFocus
Else
   ClearScreen
End If
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer, X As Integer
Dim i As Integer, imsg As Integer, sField As String
On Error GoTo cmdUpdate_Err
If IsValid Then
  '--
If gblOptions = 1 Then ' we are adding a new lien
   i = RunSP(SpCon, "usp_AssignUpdate", 0, 1, CInt(gblFileKey), Format(meb(0).Text, "dd-mmm-yyyy"), gblLoginName, tbfld(0).Text)
Else
   i = RunSP(SpCon, "usp_AssignUpdate", 0, 2, CInt(gblFileKey), Format(meb(0).Text, "dd-mmm-yyyy"), gblLoginName, tbfld(0).Text, Format(meb(1).Text, "dd-mmm-yyyy"))
End If

If i <> 1 Then
   MsgBox "Update was unsucessfull"
Else
   MsgBox "Record successfully updated"
End If
End If
'---

Done:
 Exit Sub
'--
cmdUpdate_Err:
  MsgBox Err & " " & Err.Description, vbOKOnly, "SIS012/cmdUpdate"
  cmdCancel_Click
  
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
  GoTo Form_Activate_Exit
End If
'--
If gblOptions = 2 Then
  If rsMain.RecordCount > 0 Then UpdateScreen
  Me.Caption = "Edit Lien on Certification"
Else
    Me.Caption = "Add Lien on Certification"
End If

'--
Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
 If Err = -2147168242 Then ' no current transactions
   Resume 0
 Else
   MsgBox "SIS012/Activate"
   Shutdown
   Unload Me
   Exit Sub
 End If
End Sub

Private Sub Form_Load()
Dim iDay As Integer
Dim qSQL As String
Dim indx As Integer
Dim strTmp As String
'On Error GoTo FL_ERR
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
 lbl(0).Caption = gblFileKey
 lbl(1).Caption = frmSIS011.tbfld(1).Text
'--
csvCenterForm Me, gblMDIFORM
'-----
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
iOpenMain = False
iOpenAdt = False
Set rsMain = New ADODB.Recordset

'----------------------------
'---- open recordsets -----
'----------------------------
iOpenAdt = True

Set rsMain = RunSP(SpCon, "usp_AssignSelect", 1, CLng(gblFileKey))
iOpenMain = True
If Not rsMain.EOF Then UpdateScreen
'--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS012/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
End Sub
Private Sub UpdateScreen()
 With rsMain
    meb(0).Text = !ASSDAT
    tbfld(0).Text = !ASSIGNEE
    meb(0).Enabled = False
    tbfld(0).Enabled = False
 End With
End Sub


Private Sub meb_GotFocus(Index As Integer)
Select Case Index
Case 0
If gblOptions = 1 Then meb(0).Mask = "##-???-####"
Case Else
End Select
End Sub

Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
   KeyCode = 0
   Select Case Index
   Case 0
      cmdUpdate.SetFocus
   Case Else
   End Select
Case vbKeyUp
   KeyCode = 0
   Select Case Index
   Case 1
      meb(0).SetFocus
   Case Else
   End Select
Case Else
End Select
End Sub

Private Sub ClearScreen()
 meb(0).Text = ""
 tbfld(0).Text = ""
  If gblOptions = 2 Then
     meb(1).Text = "'"
     Set rsMain = RunSP(SpCon, "usp_AssignSelect", 1, CLng(gblFileKey))
     If rsMain.RecordCount > 0 Then UpdateScreen
  End If
End Sub




