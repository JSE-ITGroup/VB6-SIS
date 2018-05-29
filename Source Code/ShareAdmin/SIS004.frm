VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSIS004 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archive Audit Data"
   ClientHeight    =   2190
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "SIS004.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6735
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      ToolTipText     =   "Audit records created on or before the effective date will be purged."
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy "
      Mask            =   "##-???-####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3480
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5640
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4560
      TabIndex        =   1
      Top             =   1800
      Width           =   975
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
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   1680
      Y2              =   1680
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
      Alignment       =   1  'Right Justify
      Caption         =   "Effective Date:"
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
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   1740
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
      Left            =   720
      TabIndex        =   6
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmSIS004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, OpenErr As Integer
Dim rsARC As ADODB.Recordset
Dim rsAUD As ADODB.Recordset
Dim iARCOpen As Integer, iAUDOpen As Integer
Private Sub cmdCancel_Click()
cnn.RollbackTrans
'''cnn.close
Unload Me
End Sub

Private Sub cmdClear_Click()
meb.Mask = "##-???-####"
meb.SetFocus
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer
Dim qSQL As String
Dim iErr As Integer
On Error GoTo cmdUpdate_Err
'--
If IsNull(meb.Text) Then
 iErr = 90
 meb.SetFocus
 GoTo entry_err
End If
'--
If Not IsDate(meb.Text) Then
     iErr = 14
     meb.SetFocus
     GoTo entry_err
End If
'--

Set rsAUD = New ADODB.Recordset
Set rsARC = New ADODB.Recordset
'--
qSQL = "SELECT * FROM AUDTRN where CHGDATE = " & meb.Text & ""
qSQL = qSQL & " or CHGDATE < " & meb.Text & ""

'--
rsAUD.Open qSQL, cnn, adOpenKeyset, adLockBatchOptimistic, adCmdText
iAUDOpen = True
rsARC.Open "ArchAudit", cnn, adOpenKeyset, adLockOptimistic, adCmdTable
iARCOpen = True
'--

With rsAUD
'---
 
 If .EOF And .BOF Then
    iErr = 92
    Shutdown
    meb.Mask = "##-???-####"
    meb.SetFocus
    GoTo entry_err
  End If
  .MoveFirst
  '--
  Do While Not .EOF
     rsARC.AddNew
     rsARC!tableno = !tableno
     rsARC!fieldid = !fieldid
     rsARC!UserId = !UserId
     rsARC!chgdate = !chgdate
     rsARC!newinfo = !newinfo
     rsARC!oldinfo = !oldinfo
     rsARC.Update
     .MoveNext
  Loop
  '.Requery
 ' .Filter = "chgdate = #" & meb.Text & "# or chgdate < #" & meb.Text & "#"
  .MoveFirst
  Do While Not .EOF
    .Delete
    .MoveNext
  Loop
  .UpdateBatch

End With
Shutdown
Set rsARC = Nothing
Set rsAUD = Nothing
cnn.CommitTrans
'''cnn.close
Unload Me
'---
Done:
 Exit Sub
'--
cmdUpdate_Err:
  MsgBox "SIS004/cmdUpdate"
  Shutdown
  Set rsARC = Nothing
  Set rsAUD = Nothing
  cmdCancel_Click
  Resume Done
entry_err:
  csvShowUsrErr iErr, "SIS004"
  GoTo Done
End Sub

Private Sub Form_Activate()
 ' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 '--
If OpenErr = True Then
  Unload Me
End If
End Sub

Private Sub Form_Load()
On Error GoTo FL_ERR
'--
   csvCenterForm Me, gblMDIFORM
   '-------------------------------------
   '-- Initialize License Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '--
   '''Set cnn = New ADODB.Connection
   cnn.Open
   '''cnn.Errors.Clear
   OpenErr = False
   iARCOpen = False
   iAUDOpen = False
   cnn.BeginTrans
   '---
 
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS004/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
End Sub

Public Sub Shutdown()
If iAUDOpen = True Then rsAUD.Close
If iARCOpen = True Then rsARC.Close
cnn.Close
End Sub
