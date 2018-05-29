VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSIS058 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Period End Reset"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS058.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6750
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Begin"
      Default         =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   4320
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      ToolTipText     =   "Creates Certificate Print file from Register Work File."
      Top             =   3480
      Width           =   975
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   1680
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
      Top             =   3480
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   3400
      Y2              =   3400
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   480
      Y2              =   480
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
      TabIndex        =   5
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS058"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpCon As ADODB.Connection

Private Sub cmdBtn_Click(Index As Integer)
Dim X As Integer
On Error GoTo cmdBtn_Click_Err
Select Case Index
Case 0 'Cancel
    Unload Me
    Set frmSIS058 = Nothing
   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
Case 1 'Start Processing
   OpenFiles
   X = RunSP(SpCon, "usp_CloseCycle", 0)
   If X = 0 Then
      MsgBox "Cycle successfully closed"
      GoTo cmdBtn_Click_Exit
   End If
   If X = 1 Then
      MsgBox "A problem was encountered in trying to close the cycle"
   GoTo cmdBtn_Click_Exit
   End If
   If X = 3 Then
      MsgBox "The Audit Report was not printed!"
      GoTo cmdBtn_Click_Exit
   End If
   If X = 4 Then
      MsgBox "The Certificate was not printed!"
      GoTo cmdBtn_Click_Exit
   End If
   If X = 5 Then
      MsgBox "The Certficate Register was not printed!"
      GoTo cmdBtn_Click_Exit
   End If
   If X = 6 Then
      MsgBox "Please run the Verification report!"
      GoTo cmdBtn_Click_Exit
   End If
   
Done:
    cmdBtn_Click (0)
Case Else
End Select

cmdBtn_Click_Exit:
  
  Exit Sub
cmdBtn_Click_Err:
  MsgBox "SIS058/cmdBTN"
  Unload Me
End Sub

Private Sub Form_Load()
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
   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
   
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS058/Load"
  Unload Me
End Sub

Private Sub InitProgressBar(max As Long)
    ProgressBar1.Min = 0
    ProgressBar1.max = max
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.Min

End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub OpenFiles()
On Error GoTo OpenFiles_Err
'--
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

OpenFiles_Exit:
Exit Sub
OpenFiles_Err:
  MsgBox "SIS058/OpenFiles"
  cmdBtn_Click (0)
  GoTo OpenFiles_Exit
End Sub

