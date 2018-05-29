VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmSIS000 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control File Maintenance"
   ClientHeight    =   4950
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "SIS000.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6750
   Begin VB.Frame Frame1 
      Caption         =   "Archive Data"
      Height          =   975
      Left            =   4440
      TabIndex        =   22
      Top             =   2520
      Width           =   2055
      Begin SSDataWidgets_A.SSDBOptSet optBtn 
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1710
         _Version        =   196611
         _ExtentX        =   3069
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "&Yes"
         BackColor       =   -2147483643
         Cols            =   2
         IndexSelected   =   0
         NumberOfButtons =   2
         Buttons.Button(0).OptionValue=   "-1"
         Buttons.Button(0).Caption=   "&Yes"
         Buttons.Button(0).Mnemonic=   89
         Buttons.Button(0).Value=   -1  'True
         Buttons.Button(0).TextLeft=   15
         Buttons.Button(0).TextRight=   33
         Buttons.Button(0).TextBottom=   14
         Buttons.Button(0).ButtonRight=   13
         Buttons.Button(0).ButtonBottom=   13
         Buttons.Button(0).PictureLeft=   35
         Buttons.Button(0).PictureRight=   34
         Buttons.Button(0).PictureBottom=   14
         Buttons.Button(0).ButtonToColRight=   57
         Buttons.Button(0).ButtonToColBottom=   14
         Buttons.Button(0).ButtonBitmapID=   2
         Buttons.Button(1).OptionValue=   "0"
         Buttons.Button(1).Caption=   "&No"
         Buttons.Button(1).Mnemonic=   78
         Buttons.Button(1).TextLeft=   73
         Buttons.Button(1).TextRight=   87
         Buttons.Button(1).TextBottom=   14
         Buttons.Button(1).ButtonLeft=   58
         Buttons.Button(1).ButtonRight=   71
         Buttons.Button(1).ButtonBottom=   13
         Buttons.Button(1).PictureLeft=   89
         Buttons.Button(1).PictureRight=   88
         Buttons.Button(1).PictureBottom=   14
         Buttons.Button(1).ButtonToColLeft=   58
         Buttons.Button(1).ButtonToColRight=   115
         Buttons.Button(1).ButtonToColBottom=   14
         Buttons.Button(1).Column=   1
      End
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   2400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   13
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5640
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4560
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   1
      Left            =   4920
      MaxLength       =   12
      TabIndex        =   1
      ToolTipText     =   "Enter the Company's Tax Reference Number"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   0
      Left            =   960
      MaxLength       =   11
      TabIndex        =   0
      ToolTipText     =   "Enter the Company's TRN Reference Number like 000-000-000"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   5
      Left            =   1920
      MaxLength       =   16
      TabIndex        =   5
      ToolTipText     =   "Enter the Company's Address Line 4 "
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   4
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   4
      ToolTipText     =   "Enter the Company's Address line 3 here "
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   3
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   3
      ToolTipText     =   "Enter the Company's Address Line 2 here"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   2
      Left            =   1920
      MaxLength       =   35
      TabIndex        =   2
      ToolTipText     =   "Enter the Company's Address line 1 here"
      Top             =   1200
      Width           =   3375
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   13
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   18
      Top             =   3120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   5
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   20
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   13
      Format          =   "$#,##0;($#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   25
      Top             =   3840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Ownership %:"
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
      Left            =   360
      TabIndex        =   24
      Top             =   3840
      Width           =   1455
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
      TabIndex        =   21
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Par Value:"
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
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Stocks on Record:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Issuable Stocks:"
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
      TabIndex        =   15
      Top             =   2400
      Width           =   1455
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
      TabIndex        =   14
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "CUSIP  No:"
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
      Left            =   3480
      TabIndex        =   13
      Top             =   600
      Width           =   1380
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
      TabIndex        =   12
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "TRN No:"
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
      TabIndex        =   11
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   1200
      Width           =   735
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
      TabIndex        =   17
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, iADTYesNo As Integer
Dim rsCmp As ADODB.Recordset
Dim OpenErr As Integer
Dim SpCon As ADODB.Connection

Function IsValid() As Integer
Dim iErr As Integer
IsValid = True
'--
If IsNothing(tbfld(0)) Then  ' TRN Number
   iErr = 8
   tbfld(0).SetFocus
   GoTo Validate_Err
 End If
 If tbfld(1) = "" Then 'CUSIP Number
   iErr = 88
   tbfld(1).SetFocus
   GoTo Validate_Err
 End If
 '--
 If tbfld(2) = "" Then ' Address Line 1
   iErr = 9
   tbfld(2).SetFocus
   GoTo Validate_Err
 End If
 '--
 If tbfld(3) = "" Then  ' Address Line 2
   iErr = 9
   tbfld(3).SetFocus
   GoTo Validate_Err
 End If
 tbfld(4) = Trim(tbfld(4))
 tbfld(5) = Trim(tbfld(5))
 '--
 If meb(0) = "" Then
   iErr = 89
   meb(0).SetFocus
   GoTo Validate_Err
 Else
   If Not IsNumeric(meb(0)) Then
      iErr = 28
      meb(0).SetFocus
      GoTo Validate_Err
   End If
 End If
 meb(0) = Val(meb(0))
 '--
 If meb(1) = "" Then
   iErr = 90
   meb(1).SetFocus
   GoTo Validate_Err
 Else
   If Not IsNumeric(meb(1)) Then
      iErr = 28
      meb(1).SetFocus
      GoTo Validate_Err
   End If
 End If
 meb(1) = Val(meb(1))
 '--
 If meb(2) = "" Then
   iErr = 107
   meb(2).SetFocus
   GoTo Validate_Err
 Else
   If Not IsNumeric(meb(2)) Then
      iErr = 28
      meb(2).SetFocus
      GoTo Validate_Err
   End If
 End If
 meb(2) = Val(meb(2))
 '--
 If meb(3) = "" Then
   iErr = 108
   meb(3).SetFocus
   GoTo Validate_Err
 Else
   If Not IsNumeric(meb(3)) Then
      iErr = 28
      meb(3).SetFocus
      GoTo Validate_Err
   End If
 End If
 meb(2) = Val(meb(2))
 '--
 If meb(4) <> "" Then
    If Not IsNumeric(meb(4)) Then
       iErr = 28
       meb(4).SetFocus
       GoTo Validate_Err
    End If
 End If
Validate_Exit:
   Exit Function
'--
Validate_Err:
  'MsgBox msg, vbInformation, "Users"
  csvShowUsrErr iErr, "Company"
  IsValid = False
  GoTo Validate_Exit
'--
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer
On Error GoTo cmdUpdate_Err
If IsValid Then
  '--
  strChg = RunSP(SpCon, "usp_CompanyUpdate", 0, gblLoginName, tbfld(2), tbfld(3), tbfld(4), tbfld(5), tbfld(0), tbfld(1), CDbl(meb(0)), CDbl(meb(1)), CCur(meb(2)), CCur(meb(3)), CLng(meb(4)), Optbtn.OptionValue)
If strChg = 0 Then
   MsgBox "Control File was successfully updated"
Else
   MsgBox "Update failed!"
End If
End If

Done:
 Exit Sub
'--
cmdUpdate_Err:
  MsgBox "SIS000/cmdUpdate, Error on Update"
 
End Sub

Private Sub Form_Activate()
' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
If OpenErr = True Then
  Unload Me
End If
End Sub

Private Sub Form_Load()
On Error GoTo FL_ERR
'--
   csvCenterForm Me, gblMDIFORM
   Set SpCon = New ADODB.Connection
   With SpCon
        .ConnectionString = gblFileName
        .CursorLocation = adUseClient
        .ConnectionTimeout = 0
        '.Provider = "SQLOLEDB.1"
   End With
   SpCon.Open , , , adAsyncConnect
   Do While SpCon.State = adStateConnecting
      Screen.MousePointer = vbHourglass
      frmMDI.txtStatusMsg.SimpleText = "Connecting, Please wait......"
   Loop
   Screen.MousePointer = vbDefault

   Set rsCmp = RunSP(SpCon, "usp_Company", 1)
    
   '-------------------------------------
   '-- Initialize License Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '--
   UpdateScreen
 
 '--
 OpenErr = False
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS000/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
   
End Sub
Private Sub UpdateScreen()
With rsCmp
  If Not .EOF Then
    tbfld(0).Text = !TRNUMBER
    tbfld(1).Text = !CUSIP
    tbfld(2).Text = !COMPSTREET
    tbfld(3).Text = !COMPPOB
    tbfld(4).Text = !COMPCITY
    tbfld(5).Text = !COMPCOUNTRY
    meb(0) = !Totstocks
    meb(1) = !issStocks
    meb(2) = !PARVALUE
    meb(3) = !TAXFREELIMIT
    meb(4) = !OWNERSHIPPER
    If !archivedata = True Then
      Optbtn.IndexSelected = 0
    Else
      Optbtn.IndexSelected = 1
    End If
  End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsCmp.Close
Set rsCmp = Nothing
SpCon.Close
End Sub

Private Sub meb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  If Index = 4 Then
     cmdUpdate.SetFocus
  Else
     meb(Index + 1).SetFocus
  End If
Case vbKeyUp
KeyCode = 0
  If Index = 0 Then
    tbfld(5).SetFocus
  Else
    meb(Index - 1).SetFocus
  End If
Case Else
End Select
End Sub

Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
KeyCode = 0
   If Index = 5 Then
      meb(0).SetFocus
   Else
      tbfld(Index + 1).SetFocus
   End If
Case vbKeyUp
   KeyCode = 0
   If Index <> 0 Then
      tbfld(Index - 1).SetFocus
   End If
Case Else
End Select
End Sub

