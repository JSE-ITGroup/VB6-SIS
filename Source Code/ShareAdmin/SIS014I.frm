VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmSIS014I 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5280
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "SIS014I.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6795
   Begin VB.ComboBox CmbFreq 
      Height          =   315
      ItemData        =   "SIS014I.frx":030A
      Left            =   1920
      List            =   "SIS014I.frx":031D
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox tb 
      Height          =   375
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   6
      Top             =   4080
      Width           =   4575
   End
   Begin SSDataWidgets_A.SSDBOptSet optBtn 
      Height          =   495
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   600
      Width           =   1365
      _Version        =   196611
      _ExtentX        =   2408
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "&No"
      BackColor       =   -2147483643
      IndexSelected   =   1
      NumberOfButtons =   2
      Buttons.Button(0).OptionValue=   "-1"
      Buttons.Button(0).Caption=   "&Yes"
      Buttons.Button(0).Mnemonic=   89
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   33
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   35
      Buttons.Button(0).PictureRight=   34
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   90
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(1).OptionValue=   "0"
      Buttons.Button(1).Caption=   "&No"
      Buttons.Button(1).Mnemonic=   78
      Buttons.Button(1).Value=   -1  'True
      Buttons.Button(1).TextLeft=   15
      Buttons.Button(1).TextTop=   16
      Buttons.Button(1).TextRight=   29
      Buttons.Button(1).TextBottom=   30
      Buttons.Button(1).ButtonTop=   16
      Buttons.Button(1).ButtonRight=   13
      Buttons.Button(1).ButtonBottom=   29
      Buttons.Button(1).PictureLeft=   31
      Buttons.Button(1).PictureTop=   16
      Buttons.Button(1).PictureRight=   30
      Buttons.Button(1).PictureBottom=   30
      Buttons.Button(1).ButtonToColTop=   16
      Buttons.Button(1).ButtonToColRight=   90
      Buttons.Button(1).ButtonToColBottom=   30
      Buttons.Button(1).ButtonBitmapID=   2
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3600
      TabIndex        =   7
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5760
      TabIndex        =   10
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4680
      TabIndex        =   8
      Top             =   4920
      Width           =   975
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "Enter date on record in the format dd-mmm-yyyy"
      Top             =   1320
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
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
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
      Index           =   3
      Left            =   5640
      TabIndex        =   5
      Top             =   2640
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   13
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Enter date on record in the format dd-mmm-yyyy"
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblLabels 
      Caption         =   "Interest Calculation Date:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Close:"
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
      Index           =   13
      Left            =   4320
      TabIndex        =   20
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cheque Remarks:"
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
      Index           =   12
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Interest Rate"
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
      Index           =   10
      Left            =   3960
      TabIndex        =   18
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Interest Fequency"
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
      Left            =   0
      TabIndex        =   17
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Payable:"
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
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Tax Free Limit:"
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
      TabIndex        =   15
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label lblTaxFree 
      Caption         =   "Tax Free"
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
      Left            =   1920
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
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
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Last Interest Calculation Date:"
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
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   795
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
      TabIndex        =   14
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS014I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, iEOF As Integer
Dim Conn As ADODB.Connection
Dim rsCmp As ADODB.Recordset
Dim rsInt As ADODB.Recordset
Dim iOpenMain As Integer
Dim iOpenCmp As Integer
Dim OpenErr As Integer
Dim strTable As String
Dim strRecNO As String
Dim iMode As Integer ' 0 = new; 1 = active

Function IsValid() As Integer
Dim iErr As String, dtefld As Date
IsValid = False
iErr = 0
'--
 '--
 If meb(0) = "" Then 'last interest calculation date
   iErr = "Please enter last interest claculation date"
   MsgBox iErr, vbOKOnly, "Last Interest Calculation date Information Entry"
   meb(0).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(0)) Then
      iErr = "Invalid last interest calculation date"
      MsgBox iErr, vbOKOnly, "Last Interest Calculation date Information Entry"
      meb(0).SetFocus
      GoTo Validate_Exit
    End If
 End If
 '--
 If meb(1) = "" Then ' date on record
   iErr = "Please enter the Interest Calculation Date"
   MsgBox iErr, vbOKOnly, "Payment Information Entry"
   meb(1).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(1)) Then
      iErr = "The calculation date enter is invalid"
      MsgBox iErr, vbOKOnly, "Payment Information Entry"
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If meb(2) = "" Then 'Date Payable
   iErr = "Please enter the Date Payable"
   MsgBox iErr, vbOKOnly, "Payment Information Entry"
   meb(2).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(2)) Then
      iErr = "The Payable date entered is invalid"
      MsgBox iErr, vbOKOnly, "Payment Information Entry"
      meb(2).SetFocus
      GoTo Validate_Exit
   Else
      If DateValue(meb(2).Text) < DateValue(meb(0).Text) Then
         iErr = "The date payable is before the last interest calculation date"
         MsgBox iErr, vbOKOnly, "Payment Information Entry"
         meb(2).SetFocus
         GoTo Validate_Exit
      End If
   End If
 End If
 '--
 'If meb(3) = "" Then 'payment Data
 '  iErr = 113
 '  csvShowUsrErr iErr, "Payment Information Entry"
 '  meb(3).SetFocus
 '  GoTo Validate_Exit
 'End If
 If meb(3) <> "" And Not IsNumeric(meb(3)) Then
      iErr = "Invalid Interest Rate. Please correct"
      meb(3).SetFocus
      GoTo Validate_Err
 End If
 '--
IsValid = True
Validate_Exit:
   Exit Function
'--
Validate_Err:
  MsgBox iErr, vbOKOnly, "Company"
  IsValid = False
  GoTo Validate_Exit
'--
End Function

Private Sub cmdCancel_Click()
If iOpenMain = True Then rsInt.Close
If iOpenCmp = True Then rsCmp.Close

Set rsInt = Nothing
Set rsCmp = Nothing
iEOF = True
Unload Me
Set frmSIS014I = Nothing
frmSIS013I.Visible = True
End Sub

Private Sub cmdClear_Click()
ClearScreen
If iMode = 1 Then UpdateScreen
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer, i As Integer
Dim str As String * 1
Dim strMth As String * 2, strDay As String * 2
Dim strRecNO As String

On Error GoTo cmdUpdate_Err
strMth = CStr(Month(meb(0)))
str = Mid(strMth, 2, 1)
If str = " " Then strMth = "0" & strMth
strDay = CStr(Day(meb(0)))
str = Mid(strDay, 2, 1)
If str = " " Then strDay = "0" & strDay
  strRecNO = UCase(Year(meb(0)) & strMth _
             & strDay)
             
If IsValid Then
  '--
  i = RunSP(Conn, "usp_IntRefUpdate", 0, meb(0), meb(1), meb(2), CmbFreq.Text, _
      CCur(meb(3)), _
  CInt(optBtn(0).OptionValue), tb, gblLoginName, strRecNO)
  If i = 0 Then
     MsgBox "Record Updated"
  Else
     MsgBox "Update Failed"
  End If
  ClearScreen
  'UpdateScreen
End If
'---
Done:
 Exit Sub
'--
cmdUpdate_Err:
  MsgBox Err & " " & Err.Description, vbOKOnly, "SIS014I/cmdUpdate"
  cmdCancel_Click
End Sub

Private Sub Form_Activate()
If OpenErr = True Then
  If iOpenMain = True Then
    rsInt.Close
  End If
  If iOpenCmp = True Then
    rsCmp.Close
  End If
  Set rsCmp = Nothing
  '''set cnn = nothing
  Set frmSIS014I = Nothing
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
Set Conn = New ADODB.Connection
'On Error GoTo FL_ERR
iEOF = False
'--
   csvCenterForm Me, gblMDIFORM
   Set Conn = New ADODB.Connection
   With Conn
         .ConnectionString = gblFileName
         .CursorLocation = adUseServer
         .ConnectionTimeout = 0
         '.Provider = "SQLOLEDB.1"
    End With
    Conn.Open , , , adAsyncConnect
    Do While Conn.State = adStateConnecting
       Screen.MousePointer = vbHourglass
       frmMDI.txtStatusMsg.SimpleText = "Connecting, Please wait......"
       frmMDI.txtStatusMsg.Refresh
    Loop
    Screen.MousePointer = vbDefault
'''cnn.Errors.Clear
   OpenErr = False
   iOpenMain = False
   iOpenCmp = False
   '-----------------------
   '-- open tables --------
   '-----------------------
   Set rsInt = RunSP(Conn, "usp_IntRef", 1)
   iOpenMain = True
   Set rsCmp = RunSP(Conn, "usp_Company", 1)
   iOpenCmp = True
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   lblTaxFree.Caption = rsCmp!TAXFREELIMIT
   '--
   If rsInt.EOF = True Then
      iMode = 0
   Else
      iMode = 1
   End If
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS014I/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
End Sub
Private Sub UpdateScreen()
Dim i As Integer, bm As Variant
With rsInt
  If Not .EOF Then
     meb(0).Text = !licdate
     meb(0).Enabled = False
     meb(1).Text = Format(!InterestDate, "dd-mmm-yyyy")
     meb(2).Text = !ChqDate
     CmbFreq.Text = !IntFreq
     meb(3).Text = Format(!InterestRate, "##.###")
     'meb(3).Text = !InterestRate / 100
     If !CLOSED = True Then
        optBtn(0).IndexSelected = 0
     Else
        optBtn(0).IndexSelected = 1
     End If
     If Not IsNull(!Remarks) Then tb = !Remarks
   Else
     iMode = 2
     meb(0).Enabled = True
    ' optBtn(1).IndexSelected = 1
   End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Conn.Close
If iEOF = False Then
  Cancel = -1
End If
End Sub

Private Sub meb_GotFocus(Index As Integer)
Select Case Index
Case 0
  If iMode = 0 Then meb(Index).Mask = "##-???-####"
Case 1 To 2
  meb(Index).Mask = "##-???-####"
Case Else
End Select
End Sub

Private Sub ClearScreen()
Dim i As Integer
For i = 0 To 3
  meb(i).Mask = ""
  meb(i).Text = ""
Next
tb = ""
CmbFreq = "Quarterly"
optBtn(0).IndexSelected = 1
End Sub

