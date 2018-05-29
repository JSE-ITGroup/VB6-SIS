VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSIS019 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post Payments"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS019.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6765
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Begin"
      Default         =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   4320
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      ToolTipText     =   "Returns to main menu"
      Top             =   3360
      Width           =   975
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   6132
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      Height          =   300
      Index           =   0
      Left            =   5400
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      ToolTipText     =   "Returns to main menu"
      Top             =   3360
      Width           =   975
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
      TabIndex        =   5
      Top             =   0
      Width           =   375
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
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Display program Information"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   6135
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
      TabIndex        =   6
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS019"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpCon As ADODB.Connection

Private Sub cmdBtn_Click(Index As Integer)
Dim msgno As Integer
Dim X As Integer
On Error GoTo cmdBtn_Click_Err
Select Case Index
Case 0 'Cancel
    Unload Me
    Set frmSIS019 = Nothing
    frmSIS013.Visible = True
     '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
Case 1 'Perform Post
msgno = 120
'X = csvYesNo(msgno, "Post Payments")
X = MsgBox("Post Payments?", vbYesNo, "Confirm Post Payments")
X = vbYes
If X = vbNo Then
  Unload Me
Else
  '--
  ' wait & hourglass message
  '--------------
  frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
  Screen.MousePointer = vbHourglass
  frmMDI.txtStatusMsg.Refresh
  '----------------------
  '--post payment cheques
  '----------------------
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

  X = RunSP(SpCon, "usp_PostPayments", 0)
  If X = 0 Then
     MsgBox "Payments successfully posted"
     frmSIS013.CmdBtn(7).Enabled = False
  Else
     If X = 2 Then
        MsgBox "Duplicate cheque numbers found", vbCritical + vbOKOnly
     Else
        MsgBox "No active dividend found to post", vbCritical + vbOKOnly
     End If
  End If
 SpCon.Close
End If
Case Else
End Select
cmdBtn_Click_Exit:
Exit Sub
cmdBtn_Click_Err:
  MsgBox Err & " " & Err.Description
  cmdBtn_Click (0)
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim strTmp As String
On Error GoTo FL_ERR
'--
   csvCenterForm Me, gblMDIFORM
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '----------------------
   '--  disable menu items
   '----------------------
   frmMDI.mnuFile.Enabled = False
   frmMDI.btnClose.Enabled = False
   frmMDI.mnuLists.Enabled = False
   frmMDI.mnuAct.Enabled = False
   frmMDI.mnuAdm.Enabled = False
   '--
   ProgressBar1.Visible = False
   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS019/Load"
  Unload Me
End Sub

Private Sub InitProgressBar(max As Long)
    ProgressBar1.Min = 0
    ProgressBar1.max = max
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.Min

End Sub

