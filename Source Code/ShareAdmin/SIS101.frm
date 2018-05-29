VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSIS101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rights Issue Information Details"
   ClientHeight    =   2715
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "SIS101.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6735
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3480
      TabIndex        =   5
      ToolTipText     =   "Clears the screen"
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5640
      TabIndex        =   4
      ToolTipText     =   "Terminates the process"
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4560
      TabIndex        =   3
      ToolTipText     =   "Saves the entries"
      Top             =   2280
      Width           =   975
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Enter date on record in the format dd-mmm-yyyy"
      Top             =   600
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
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Enter the number of stocks to be offered"
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   6
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Enter the base amount of stocks for the stock offer"
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   13
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   14
      ToolTipText     =   "Enter the Issue Price for the offered Stocks"
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   13
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Issue Price"
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
      Index           =   6
      Left            =   480
      TabIndex        =   15
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Stock(s) currently owned"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Rights Issue of "
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
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "for Every"
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
      Left            =   480
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Stock(s)"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   2160
      Y2              =   2160
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Record Date:"
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
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   1575
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
      TabIndex        =   9
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, iEOF As Integer
Dim rsRiWrk As ADODB.Recordset
Dim rsRights As ADODB.Recordset
Dim iOpenMain As Integer, iOpenWrk As Integer
Dim OpenErr As Integer
Dim iMode As Integer ' 0 = add

Function IsValid() As Integer
Dim iErr As Integer, dtefld As Date
IsValid = False
iErr = 0
'--
If meb(0) = "" Then ' date on record
   iErr = 111
   csvShowUsrErr iErr, "Rights Issue Information Entry"
   meb(0).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(0)) Then
      iErr = 14
      csvShowUsrErr iErr, "Rights Issue Information Entry"
      meb(0).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
  If IsNothing(meb(1)) Or Not IsNumeric(meb(1)) Then
      iErr = 28
      meb(1).SetFocus
      GoTo Validate_Err
 End If
 If IsNothing(meb(2)) Or Not IsNumeric(meb(2)) Then
      iErr = 28
      meb(2).SetFocus
      GoTo Validate_Err
 End If
 If IsNothing(meb(3)) Or Not IsNumeric(meb(3)) Then
    iErr = 28
    meb(3).SetFocus
    GoTo Validate_Err
 End If
 '--
 IsValid = True
Validate_Exit:
   Exit Function
'--
Validate_Err:
  'MsgBox msg, vbInformation, "Users"
  csvShowUsrErr iErr, "Rights Issue Entry"
  IsValid = False
  GoTo Validate_Exit
'--
End Function

Private Sub cmdCancel_Click()
If iOpenMain = True Then rsRights.Close
If iOpenWrk = True Then rsRiWrk.Close
Set rsRights = Nothing
Set rsRiWrk = Nothing
'''set cnn = nothing
iEOF = True
Unload Me
Set frmSIS101 = Nothing
End Sub

Private Sub cmdClear_Click()
ClearScreen
If iMode = 1 Then UpdateScreen
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer, i As Integer
Dim strMth As String * 2, strDay As String * 2
Dim iPayamt As Integer
Dim iPayPct As Integer
Dim str As String * 1
On Error GoTo cmdUpdate_Err
If IsValid Then
  '--
  If iMode = 0 Then
     rsRights.AddNew
  End If
  rsRights!RECDAT = DateValue(meb(0).Text)
  rsRights!STKSTO = Val(meb(1))
  rsRights!STKBASE = Val(meb(2))
  rsRights!RIPRICE = Val(meb(3))
  rsRights.Update
  ClearScreen
  UpdateScreen
End If
'---
Done:
 Exit Sub
'--
cmdUpdate_Err:
  MsgBox "SIS101/cmdUpdate"
  cmdCancel_Click
End Sub
Private Sub Form_Activate()
If OpenErr = True Then
  If iOpenMain = True Then
    rsRights.Close
  End If
  If iOpenWrk = True Then rsRiWrk.Close
  '''set cnn = nothing
  Set frmSIS101 = Nothing
  iEOF = True
  Unload Me
Else

UpdateScreen
End If

 ' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 '--
End Sub

Private Sub Form_Load()
Dim iDay As Integer
Dim qSQL As String
Dim i As Integer
Dim strTmp As String
On Error GoTo FL_ERR
iEOF = False
'--
   csvCenterForm Me, gblMDIFORM
   '''Set cnn = New ADODB.Connection
   cnn.Open
   OpenErr = False
   iOpenMain = False
   iOpenWrk = False
   Set rsRiWrk = New ADODB.Recordset
   Set rsRights = New ADODB.Recordset
   On Error GoTo Create_RITable
   rsRiWrk.Open "STKRIWRK", cnn, , , adCmdTable
   iOpenWrk = True
   On Error GoTo FL_ERR
   rsRights.Open "BONUSREF", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
   iOpenMain = True
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '--
   If rsRights.EOF = True Then
      iMode = 0
   Else
      iMode = 1
   End If
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS101/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
Create_RITable:
qSQL = "Create Table STKRIWRK (" _
        & "CLIENTID  long not null CONSTRAINT pkCLIENTID PRIMARY KEY, " _
        & "SHARES  long not null, " _
        & "OFFER   long not null, " _
        & "UNUSED  long not null, " _
        & "COST    currency not null, " _
        & "LEDGER  text(1) not null)"
 X = csvADODML(qSQL, cnn)
 '--
 qSQL = "ALTER TABLE BONUSREF ADD RIPRICE Currency"
 X = csvADODML(qSQL, cnn)
Resume 0
End Sub
Private Sub UpdateScreen()
Dim i As Integer

With rsRights
  If Not .EOF Then
      meb(0).Text = !RECDAT
      meb(1).Text = !STKSTO
      meb(2).Text = !STKBASE
      If Not IsNull(!RIPRICE) Then meb(3).Text = !RIPRICE
  End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
cnnClose
If iEOF = False Then
  Cancel = -1
End If
End Sub

Private Sub meb_GotFocus(Index As Integer)
Select Case Index
Case 0
  If iMode = 0 Then meb(Index).Mask = "##-???-####"
Case Else
End Select
End Sub

Private Sub meb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  If Index = 3 Then
     cmdUpdate.SetFocus
  Else
     meb(Index + 1).SetFocus
  End If
Case vbKeyUp
KeyCode = 0
  If Index <> 0 Then meb(Index - 1).SetFocus
Case Else
End Select
End Sub



Private Sub ClearScreen()
Dim i As Integer
For i = 0 To 3
  meb(i).Mask = ""
  meb(i).Text = ""
Next
End Sub





