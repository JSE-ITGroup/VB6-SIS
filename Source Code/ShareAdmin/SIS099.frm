VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS099 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferences"
   ClientHeight    =   3795
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "SIS099.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6810
   Begin VB.TextBox TxtFinacle 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   7
      ToolTipText     =   "Enter the period ending date for Transfer & Issue Processing."
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5760
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4680
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      ToolTipText     =   "Enter the next number to assign to new clients."
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "###0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   9
      ToolTipText     =   "Enter the next number to assign to printed cheques. Required only if processing payments."
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "###0"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBChqFormat 
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   2280
      Width           =   2415
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   503
      Columns.Count   =   2
      Columns(0).Width=   6456
      Columns(0).Caption=   "Cheque Format"
      Columns(0).Name =   "Cheque Format"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "FormatID"
      Columns(1).Name =   "FormatID"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   4260
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Label Label2 
      Caption         =   "Cheque Format:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Finacle Account No:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Next Certificate No:"
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
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
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
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Next Client No:"
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
      Left            =   600
      TabIndex        =   4
      Top             =   1080
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
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period Ending Date:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS099"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim rsCmp As ADODB.Recordset
Dim rsFinacle As ADODB.Recordset
Dim OpenErr As Integer
Dim SpCon As ADODB.Connection

Function IsValid() As Integer
Dim iErr As Integer
IsValid = False
'--
If meb(0) <> "" Then  ' Period date
   If Not IsDate(meb(0)) Then
      MsgBox "The Period Ending date is not contain a valid date", vbOKOnly, "Period Ending Date"
      meb(0).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If meb(1) <> "" Then
   If Not IsNumeric(meb(1)) Then
      MsgBox "Client ID MUST be NUMERIC", vbOKOnly, "Client ID"
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
 End If
 meb(1) = Val(meb(1))
 '--
 '--
 If meb(3) <> "" Then
   If Not IsNumeric(meb(3)) Then
      MsgBox "Next Cert No MUST be NUMERIC", vbOKOnly, "Next Certificate Number"
      meb(3).SetFocus
      GoTo Validate_Exit
   End If
 End If
 meb(3) = Val(meb(3))
 '--
 If Not IsNumeric(TxtFinacle) Then
    MsgBox "The Finacle Account No MUST contain ONLY NUMBERS", vbOKOnly, "Finacle Account Number"
    TxtFinacle.SetFocus
    GoTo Validate_Exit
 End If
 If Len(TxtFinacle) <> 9 Then
    MsgBox "The Finacle Account Number must be Nine (9) digits", vbOKOnly, "Finacle Account Number"
    TxtFinacle.SetFocus
    GoTo Validate_Exit
 End If
 
 meb(3) = Val(meb(3))
 IsValid = True
Validate_Exit:
   Exit Function
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer
On Error GoTo Err_cmdUpdate_Click

If IsValid Then
  strChg = RunSP(SpCon, "usp_PreferenceUpdate", 0, gblLoginName, Format(meb(0), "dd-mmm-yyyy"), CLng(meb(1)), CLng(meb(3)), TxtFinacle, CInt(SSDBChqFormat.Columns(1).Text))
  If strChg = 0 Then
     MsgBox "Preference information successfully updated"
  Else
     MsgBox "Update unsuccesfull"
  End If
End If

'---

Exit_cmdUpdate_Click:
 Exit Sub
'--
Err_cmdUpdate_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on preferences update"
Resume Exit_cmdUpdate_Click
End Sub

Private Sub Form_Activate()
' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 If OpenErr = True Then
  Unload Me
 Else
  UpdateScreen
End If

End Sub

Private Sub Form_Load()
On Error GoTo FL_ERR
'--
   csvCenterForm Me, gblMDIFORM
   OpenErr = False
   '--
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

   Set rsCmp = RunSP(SpCon, "usp_PreferenceSelect", 1)
   Set rsFinacle = rsCmp.NextRecordset
   
   '-------------------------------------
   '-- Initialize License Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS073/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
   
End Sub
Private Sub UpdateScreen()
Dim i As Integer
Dim bm As Variant

With rsCmp
  If Not .EOF Then
   If Not IsNull(!PeriodDte) Then meb(0).Text = !PeriodDte
   If Not IsNull(!NEXTACCT) Then meb(1).Text = !NEXTACCT
   If Not IsNull(!nextcert) Then meb(3).Text = !nextcert
   If Not IsNull(rsFinacle!AccountNo) Then TxtFinacle = rsFinacle!AccountNo
  End If
End With
With SSDBChqFormat
     .MoveFirst
     For i = 0 To .Rows - 1
         bm = .GetBookmark(i)
         If .Columns(1).CellText(bm) = Trim(rsCmp!ChqFormat) Then
            .Bookmark = .GetBookmark(i)
            SSDBChqFormat = .Columns(0).CellText(bm)
            Exit For
         End If
     Next i
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
rsCmp.Close
Set rsCmp = Nothing
rsFinacle.Close
Set rsFinacle = Nothing
SpCon.Close
End Sub

Private Sub meb_GotFocus(Index As Integer)
If Index = 0 Then
   meb(0).Mask = "##-???-####"
End If
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
  If Index <> 0 Then
    meb(Index - 1).SetFocus
  End If
Case Else
End Select
End Sub
Private Sub SSDBChqFormat_InitColumnProps()
On Error GoTo Err_SSDBChqFormat_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset
Dim i As Integer

Set adoRst = RunSP(SpCon, "usp_ListChqFormats", 1)
If adoRst.EOF Then
   MsgBox "Cheque formats are missing" & vbCrLf & "Please have IT set them up now", vbCritical, "Cheque formats Error"
   GoTo Exit_SSDBChqFormat_InitColumnProps
End If

With SSDBChqFormat
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst!ChqFormatName & vbTab & adoRst!ChqFormatID
     .AddItem StrSql
     adoRst.MoveNext
     StrSql = ""
     Loop
End With

adoRst.Close
Set adoRst = Nothing
Exit_SSDBChqFormat_InitColumnProps:
Exit Sub

Err_SSDBChqFormat_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing cheque formats"
Resume Exit_SSDBChqFormat_InitColumnProps
End Sub
