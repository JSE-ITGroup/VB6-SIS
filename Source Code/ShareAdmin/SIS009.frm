VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSIS009 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Joint Account"
   ClientHeight    =   4245
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "SIS009.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6750
   Begin VB.TextBox tbFld 
      Height          =   285
      Index           =   4
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   3
      ToolTipText     =   "Enter the name of a third joint account holder."
      Top             =   3120
      Width           =   4335
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Enter the date this joint account begins."
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin VB.TextBox tbFld 
      Height          =   285
      Index           =   3
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   2
      ToolTipText     =   "Enter the name of a second joint account holder."
      Top             =   2640
      Width           =   4335
   End
   Begin VB.TextBox tbFld 
      Height          =   285
      Index           =   2
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Enter the name of the firat joint account holder. Not the lead."
      Top             =   2160
      Width           =   4335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3480
      TabIndex        =   6
      ToolTipText     =   "Clears the screen and resets it if in edit mode."
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5640
      TabIndex        =   5
      ToolTipText     =   "Cancels changes and returns to Account maintenance."
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   4560
      TabIndex        =   4
      ToolTipText     =   "Update Joint Table for saving to disk by Accounts Maintainace"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox tbFld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   2160
      MaxLength       =   11
      TabIndex        =   8
      ToolTipText     =   "Use generate number or enter your own unique client Number"
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox tbFld 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   9
      ToolTipText     =   "Enter Address line 2"
      Top             =   1080
      Width           =   4335
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   7
      ToolTipText     =   "Enter the date this joint account ceases."
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Name 3:"
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
      Index           =   8
      Left            =   600
      TabIndex        =   19
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "End Date:"
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
      Left            =   4080
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Name 2:"
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
      Left            =   600
      TabIndex        =   17
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Name 1:"
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
      Index           =   7
      Left            =   360
      TabIndex        =   16
      Top             =   2160
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Date:"
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
      Left            =   360
      TabIndex        =   15
      Top             =   1680
      Width           =   1740
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
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
      TabIndex        =   13
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Name:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1080
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
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Lead Stockholder No:"
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
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   1980
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
      TabIndex        =   14
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iErr As Integer
Dim X As Integer
Dim rsJoint As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim OpenErr As Integer
Dim strTable As String
Dim iMode As Integer  ' 0 = New; 1 = Active; 2 = inactive joint
Function IsValid() As Integer
On Error GoTo IsValid_Err
Dim dtefld As Date
IsValid = False
'--
If iMode = 0 Or iMode = 2 Then
  '--
  If meb(0).Text = "" Then  ' Start Date
     iErr = 36
     MsgBox "Invalid Start Date", vbCritical + vbOKOnly, "Joint Accounts"
     meb(0).SetFocus
     GoTo Validate_Exit
  End If
  '--
  If Not IsDate(meb(0).Text) Then
     iErr = 14
     MsgBox "Invalid Start Date", vbCritical + vbOKOnly, "Joint Accounts"
     meb(0).SetFocus
     GoTo Validate_Err
  End If
  '--
  '--
  If tbfld(2) = "" Then 'Joint Name 1
   iErr = 105
   tbfld(2).SetFocus
   GoTo Validate_Err
  End If
  '--
  tbfld(2) = Trim(tbfld(2))
  If Not IsNull(tbfld(3)) Then tbfld(3) = Trim(tbfld(3))
Else
  If meb(1).Text = "" Then
      iErr = 37
      MsgBox "End Date was not entered", vbCritical + vbOKOnly, "End Date"
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
   '--
   If Not IsDate(meb(1).Text) Then
     iErr = 14
     MsgBox "End Date is Invalid", vbCritical + vbOKOnly, "End Date"
     meb(1).SetFocus
     GoTo Validate_Exit
   End If
   '--
   dtefld = meb(1).Text
   If dtefld < Format(meb(0).Text, "dd-mmm-yyyy") Then
      iErr = 38
      MsgBox "End Date Less than or Equal Start Date", vbCritical + vbOKOnly, "Joint Accounts"
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
  '--
End If
'--
IsValid = True
Validate_Exit:
   Exit Function
'--
Validate_Err:
  GoTo Validate_Exit
'--
IsValid_Err:
  MsgBox "SIS009/IsValid " & Err.Number & " " & Err.Description
  Shutdown
  Unload Me
End Function

Private Sub cmdCancel_Click()
  'If frmSIS002.tbJntMode = "2" Then rsJoint.Update
  Shutdown
  frmSIS002.cmdUpdate.SetFocus
  Unload Me
  Set frmSIS009 = Nothing
End Sub
Private Sub cmdClear_Click()
If iMode = 1 Then
   ClearScreen
   tbfld(2).SetFocus
Else
   ClearScreen
   UpdateScreen
End If
End Sub
Private Sub cmdUpdate_Click()
Dim strChg As Integer, iAcct As Long
Dim i As Integer
Dim newval As Integer
On Error GoTo cmdUpdate_Err
If IsValid Then
  '--
  iAcct = Val(tbfld(0).Text)
  i = RunSP(SpCon, "usp_Sis009Update", 0, iMode, iAcct, Format(meb(0).Text, "dd-mmm-yyyy"), tbfld(2), tbfld(3), tbfld(4), Format(meb(1).Text, "dd-mmm-yyyy"), gblLoginName)
  If iMode = 1 Then
     frmSIS002.Optbtn.IndexSelected = 1
  Else
     frmSIS002.Optbtn.IndexSelected = 0
  End If
  If gblOptions = 2 Then
     frmSIS002.tbJntMode = "1" ' Update done
  End If
  If iMode = 1 Then
     EnableData
     iMode = 2
     meb(0).Text = DateValue(meb(1).Text) + 1
     ClearScreen
     Exit Sub
  End If
  If gblOptions = 2 Then
     cmdCancel_Click
  Else
     frmSIS002.tbJntMode = "2" ' Update done but rec not saved
     frmSIS009.Visible = False
  End If
End If
'---

Done:
 Exit Sub
'--
cmdUpdate_Err:
  MsgBox "SIS009/cmdUpdate"
  Shutdown
  Unload Me
  frmSIS001.Show
  frmSIS002.Show
End Sub

Private Sub Form_Activate()

'--
' ready message
'---
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 '--
 If OpenErr = True Then
    Unload Me
 Else
    UpdateScreen
 End If
End Sub

Private Sub Form_Load()
Dim strTmp As String
On Error GoTo FL_ERR
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
 tbfld(0).BackColor = &HC0C0C0
 tbfld(1).BackColor = &HC0C0C0
'--
csvCenterForm Me, gblMDIFORM
'-----------------------------------
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
'''cnn.Errors.Clear
OpenErr = False
'----------------------------
'---- open recordsets -----
'-- create SQL for selecting record to edit
'----------------------------------------
'---
Set rsJoint = RunSP(SpCon, "usp_Sis009Select", 1, CDbl(gblFileKey))
'--------------------
If rsJoint.EOF = True Then
    iMode = 0
    Me.Caption = "New Joint Account"
Else
 iMode = 1
  Me.Caption = "Edit Joint Account"
End If
'--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS009/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
  
End Sub
Private Sub UpdateScreen()
tbfld(0).Text = frmSIS002.tbfld(0)
tbfld(1).Text = Trim(frmSIS002.tbfld(2))
If frmSIS002.dbc(0) = "Person" Then
   tbfld(1).Text = tbfld(1).Text & "," & Trim(frmSIS002.tbfld(1))
End If
'--
If gblOptions = 1 Then meb(0).Text = Format(Now, "dd-mmm-yyyy")
With rsJoint
If iMode = 1 Then
    meb(0).Text = !JNTSTADTE
    tbfld(2).Text = !JNTNAME1
    If Not IsNull(!JNTNAME2) Then tbfld(3).Text = !JNTNAME2
    If Not IsNull(!jntname3) Then tbfld(4).Text = !jntname3
    DisableData
Else
    iMode = 2 ' no active joint account
    meb(0).Enabled = True
End If
End With
End Sub
Private Sub ClearScreen()
  For X = 2 To 3
    tbfld(X).Text = ""
    meb(X - 2).Mask = ""
  Next
  tbfld(4) = ""
  '--
  If iMode = 1 Then
     UpdateScreen
     tbfld(2).SetFocus
  Else
     meb(0).SetFocus
  End If
End Sub


Private Sub Shutdown()
rsJoint.Close
Set rsJoint = Nothing
SpCon.Close
End Sub
Private Sub EnableEndDte()
meb(1).Visible = True
lblLabels(3).Visible = True
End Sub

Private Sub DisableData()
meb(0).Enabled = False
tbfld(2).Enabled = False
tbfld(3).Enabled = False
tbfld(4).Enabled = False
EnableEndDte
End Sub

Private Sub DisableEndDte()
meb(1).Visible = False
lblLabels(3).Visible = False
End Sub

Private Sub EnableData()
meb(0).Enabled = True
tbfld(2).Enabled = True
tbfld(3).Enabled = True
tbfld(4).Enabled = True
DisableEndDte
End Sub

Private Sub meb_GotFocus(Index As Integer)
If gblOptions = 1 Then
  meb(Index).Mask = "##-???-####"
End If
End Sub
Private Sub meb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index  'Only check for Start date
 Case 0
    Select Case KeyCode
    Case vbKeyReturn, vbKeyDown
      KeyCode = 0
      tbfld(2).SetFocus
    End Select
 Case Else
End Select
End Sub

Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn, vbKeyDown
    KeyCode = 0
    Select Case Index
      Case 2, 3
        tbfld(Index + 1).SetFocus
      Case 4
        cmdUpdate.SetFocus
      Case Else
    End Select
 Case vbKeyUp
    KeyCode = 0
    Select Case Index
     Case 3, 4
       tbfld(Index - 1).SetFocus
     Case 2
       meb(0).SetFocus
     Case Else
   End Select
Case Else
End Select
End Sub

