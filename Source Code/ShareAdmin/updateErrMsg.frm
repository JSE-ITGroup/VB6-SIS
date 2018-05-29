VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form UpdateErrMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Error Messages"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   300
      Left            =   2520
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   3600
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "display Key Field"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
End
Attribute VB_Name = "UpdateErrMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As ADODB.Recordset
Dim fs, F, iRecs As Long, iRecsAdded As Long
Dim sInRec As String, txtfile As String, sTab As String
Dim x As Integer, Y As Integer
Dim qSQL As String

Private Sub cmdCancel_Click()
Unload Me
Set UpdateErrMsg = Nothing
End Sub

Private Sub cmdStart_Click()
Dim iResp As Integer
Dim sMsg As String, sTitle As String
Dim sErrMsgL1 As String, sErrMsgL2 As String, sErrMsg As String
sErrMsgL1 = "Update failed during Purge of "
sErrMsgL2 = " Note this error."
Set rsMain = New ADODB.Recordset
'--
sMsg = "WARNING: This update will will erase your existing"
sMsg = sMsg & "  data then copy its replacement from text. Select No if "
sMsg = sMsg & " you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Updating Error Messages"
iResp = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If iResp = vbNo Then
  cmdCancel_Click
  Exit Sub
End If
'--
lbl.Caption = "Clearing existing file"
lbl.Visible = True
qSQL = "DELETE * FROM ERRMSG"
If csvADODML(qSQL) = False Then
  sErrMsg = sErrMsgL1 & "USERS." & sErrMsgL2
  GoTo Updateerrmsg_Fail
  Exit Sub
End If
'--
Set fs = CreateObject("Scripting.FileSystemObject")
lbl.Caption = "Recreating error messages"
txtfile = App.Path & "\errmsg.txt"
lbl.Refresh
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Recreating Error messages..."
frmMDI.txtStatusMsg.Refresh
'--
lbl.Caption = "Recreating Error Messages"
lbl.Refresh
txtfile = App.Path & "\errmsg.txt"
Set F = fs.opentextfile(txtfile)
sInRec = F.readline
iRecs = 0
Do Until F.AtEndofStream = True
    iRecs = iRecs + 1
    sInRec = F.readline
Loop
InitProgressBar (iRecs)
iRecs = 0: iRecsAdded = 0
ProgressBar1.Visible = True
F.Close
Set F = fs.opentextfile(txtfile)
'--Updating errmsg routine
'On Error GoTo Open_err
'--
qSQL = "Select * from ERRMSG"
rsMain.Open qSQL, gblFileName, adOpenDynamic, adLockOptimistic, adCmdText
sInRec = F.readline
If F.AtEndofStream = True Then GoTo Open_err
cnn.BeginTrans
Do Until F.AtEndofStream = True
  UpdateErr
  iRecs = iRecs + 1
  ProgressBar1.Value = iRecs
  sInRec = F.readline
Loop
If iRecsAdded > 0 Then
       cnn.CommitTrans
       iRecsAdded = 0
End If
F.Close
rsMain.Close
'--
Set rsMain = Nothing
cmdStart.Enabled = False
'-- display success message
lbl.Caption = ""
ProgressBar1.Visible = False
MsgBox "Update successfull. Select Ok to clear this message, then Cancel to end."
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Exit Sub
Open_err:
  MsgBox "Input Text File " & txtfile & " is blank; update aborting. "
 
  Exit Sub
Updateerrmsg_Fail:
   MsgBox sErrMsg
   cmdCancel_Click
   Exit Sub
End Sub
Private Sub InitProgressBar(max As Long)
  If max = 0 Then Exit Sub
    ProgressBar1.Min = 0
    ProgressBar1.max = max
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.Min

End Sub
Private Sub Form_Load()
csvCenterForm Me, gblMDIFORM
'ready Message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
ProgressBar1.Visible = False
lbl.Caption = ""
lbl.Visible = False
sTab = Chr(9)
'--
Set cnn = New ADODB.Connection
cnn.Open gblFileName
End Sub


Private Sub UpdateErr()
Dim ilen As Integer
'--
With rsMain
  .AddNew
  ilen = Len(RTrim(sInRec))
  '-- extract error code
  Y = InStr(1, sInRec, sTab, vbTextCompare)
  !errcde = left(sInRec, Y - 1)
  '-- extract error message 1 description
  x = Y + 1
  Y = InStr(x, sInRec, sTab, vbTextCompare)
  !errdes = Mid(sInRec, x, Y - x)
  '-- extract alert code
  x = Y + 1
  Y = InStr(x, sInRec, sTab, vbTextCompare)
  If Y = 0 Then
    !alert = Mid(sInRec, x, 1)
    GoTo endd
  Else
    !alert = Mid(sInRec, x, Y - x)
   End If
  '-- extract error message 2
  x = Y + 1
  If x < ilen Then !errdes2 = Mid(sInRec, x, ilen - x + 1)
  
endd:
 .Update
 iRecsAdded = iRecsAdded + 1  'count records added
 If iRecsAdded = 500 Then
    cnn.CommitTrans
    cnn.BeginTrans
    iRecsAdded = 0
 End If
End With
End Sub




