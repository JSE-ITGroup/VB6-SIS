VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSIS047 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prepare Annual Government Return "
   ClientHeight    =   2760
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "SIS047.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5205
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   1920
      TabIndex        =   9
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Cancels all processing and exits program."
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Print"
      Height          =   300
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "Saves the screen information and prints the report"
      Top             =   2400
      Width           =   975
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   0
      ToolTipText     =   "Enter the date of the last return"
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      ToolTipText     =   "Enter the date of  the current AGM."
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Enter the record date for the  current return."
      Top             =   1680
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
      Caption         =   "Current Annual  Return Date"
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
      Left            =   600
      TabIndex        =   11
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Annual General Meeting  Date"
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
      Left            =   600
      TabIndex        =   10
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6960
      Y1              =   360
      Y2              =   360
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
      TabIndex        =   7
      Top             =   0
      Width           =   735
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
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Annual Return Date"
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
      Left            =   720
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
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
      Height          =   252
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   0
      Width           =   5892
   End
End
Attribute VB_Name = "frmSIS047"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim i As Integer
Dim rptSISRept As New SISRepts
Dim fldLoop As ADODB.Field
Dim sql As String
Dim iErr As Integer
Dim rsComp As ADODB.Recordset
Dim SpCon As ADODB.Connection

Function IsValid() As Integer

Dim sElable As String
sElable = "Prepare Annual Government Return"
IsValid = False
iErr = 0
'--
If Not IsDate(meb(0)) Then
      iErr = 14
      csvShowUsrErr iErr, sElable
      meb(0).Enabled = True
      meb(0).SetFocus
      GoTo Validate_Exit
End If
'--
If Not IsDate(meb(1)) Then
      iErr = 14
      csvShowUsrErr iErr, sElable
      meb(1).SetFocus
      GoTo Validate_Exit
End If
'--
If Not IsDate(meb(2)) Then
      iErr = 14
      csvShowUsrErr iErr, sElable
      meb(2).SetFocus
      GoTo Validate_Exit
End If
'--
If DateValue(meb(1)) <= DateValue(meb(0)) Then '
    iErr = 187
    csvShowUsrErr iErr, sElable
    meb(1).SetFocus
    GoTo Validate_Exit
 End If
 '--
 If DateValue(meb(2)) <= DateValue(meb(1)) Then '
    iErr = 188
    csvShowUsrErr iErr, sElable
    meb(2).SetFocus
    GoTo Validate_Exit
 End If
 '--
 IsValid = True
Validate_Exit:
   
   Exit Function
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClear_Click()
ClearScreen
If Not IsNull(rsComp!lstagm) Then
    meb(0) = Format(rsComp!lstagm, "dd-mmm-yyyy")
End If
If Not IsNull(rsComp!curAGM) Then
    meb(1) = Format(rsComp!curAGM, "dd-mmm-yyyy")
End If

End Sub

Private Sub cmdUpdate_Click()
i = RunSP(SpCon, "usp_CreateAnnualReturn", 0, meb(0), meb(1), meb(2))
Set rptSISRept = New SISRepts
rptSISRept.ClientID = 0
rptSISRept.LoginId = gblFileName
rptSISRept.ReportNumber = 47
rptSISRept.ReportType = 9
rptSISRept.siteid = gblSiteId
rptSISRept.RunShareHolderReport
cmdCancel_Click

cmdUpdate_exit:
 Exit Sub
Create_STKRET:
  sql = "Create Table STKRETRN (" _
        & "CLIENTID  long not null CONSTRAINT pkSTKRETRN PRIMARY KEY, " _
        & "SHARES  long not null, " _
        & "LIVESHRS long not null, " _
        & "SOLDSHRS long not null)"
  i = csvADODML(sql, SpCon)
  If i = 0 Then
    GoTo cmdUpdate_Open_Err
  Else
    Resume
  End If
  
cmdUpdate_Open_Err:

Unload Me
GoTo cmdUpdate_exit
Update_err:
MsgBox "SIS047/Activate"
Unload Me
GoTo cmdUpdate_exit:
End Sub

Private Sub Form_Activate()

If Not IsNull(rsComp!lstret) Then
    meb(0) = Format(rsComp!lstret, "dd-mmm-yyyy")
End If
If Not IsNull(rsComp!curAGM) Then
    meb(1) = Format(rsComp!curAGM, "dd-mmm-yyyy")
End If
Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
   MsgBox "SIS047/Activate"
   Exit Sub
 
End Sub
Private Sub Form_Load()
On Error GoTo FL_ERR
Dim iEOF As Boolean
iEOF = False
'--
csvCenterForm Me, gblMDIFORM
Set rsComp = New ADODB.Recordset
'iOpenComp = 0: iOpenMain = 0: iOpenRet = 0
'-----------------------
'-- open tables --------
'-----------------------
'''Set cnn = New ADODB.Connection
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
rsComp.CursorLocation = adUseClient
rsComp.Open "Company", SpCon, adOpenKeyset, adLockOptimistic, adCmdTable
'''rsComp.ActiveConnection = Nothing

'iOpenComp = True
'-------------------------------------
'-- Initialize Company Details -------
'-------------------------------------
lblLabels(0).Caption = gblCompName
lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
'--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS047/Load"
  'iOpenErr = True
  GoTo FL_Exit
  
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub meb_GotFocus(Index As Integer)

Select Case Index
Case 0
  If meb(Index) = "" Then meb(Index).Mask = "##-???-####"
Case 1, 2
  meb(Index).Mask = "##-???-####"

Case Else
End Select
End Sub

Private Sub meb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  Select Case Index
  Case 0
   meb(1).SetFocus
  Case 1
    meb(2).SetFocus
     Case 3
    cmdUpdate.SetFocus
  Case Else
  End Select
Case vbKeyUp
KeyCode = 0
  Select Case Index
  
  Case 1
    meb(0).SetFocus
  Case 3
    meb(2).SetFocus
    
  Case Else
  End Select
Case Else
End Select
End Sub
Private Sub ClearScreen()
Dim X As Integer
For X = 0 To 1
    meb(X).Mask = ""
    meb(X).Text = ""
 Next

End Sub
Private Sub meb_Validate(Index As Integer, Cancel As Boolean)
If Not IsDate(meb(Index)) Then
  iErr = 14
  csvShowUsrErr iErr, "Prepare Annual Government Return"
  Cancel = True
End If

End Sub
