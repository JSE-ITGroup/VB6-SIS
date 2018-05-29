VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS008 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categories"
   ClientHeight    =   3225
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "SIS008.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6735
   Begin SSDataWidgets_A.SSDBOptSet optBtn 
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Select the required tax option for the category."
      Top             =   1680
      Width           =   1950
      _Version        =   196611
      _ExtentX        =   3440
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
      Buttons.Button(0).ButtonToColRight=   64
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
      Buttons.Button(1).OptionValue=   "0"
      Buttons.Button(1).Caption=   "&No"
      Buttons.Button(1).Mnemonic=   78
      Buttons.Button(1).TextLeft=   80
      Buttons.Button(1).TextRight=   94
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   65
      Buttons.Button(1).ButtonRight=   78
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   96
      Buttons.Button(1).PictureRight=   95
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   65
      Buttons.Button(1).ButtonToColRight=   129
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).Column=   1
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3480
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5640
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4560
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   1
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   1
      ToolTipText     =   "Enter a description for the category"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox tbfld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   0
      ToolTipText     =   "Enter a unique code for the Category"
      Top             =   720
      Width           =   495
   End
   Begin SSDataWidgets_A.SSDBOptSet HldBtn 
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      ToolTipText     =   "Select the required tax option for the category."
      Top             =   2040
      Width           =   1950
      _Version        =   196611
      _ExtentX        =   3440
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
      Buttons.Button(0).ButtonToColRight=   64
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
      Buttons.Button(1).OptionValue=   "0"
      Buttons.Button(1).Caption=   "&No"
      Buttons.Button(1).Mnemonic=   78
      Buttons.Button(1).TextLeft=   80
      Buttons.Button(1).TextRight=   94
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   65
      Buttons.Button(1).ButtonRight=   78
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   96
      Buttons.Button(1).PictureRight=   95
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   65
      Buttons.Button(1).ButtonToColRight=   129
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).Column=   1
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBAccount 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   2280
      Width           =   1335
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
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Account Number"
      Columns(0).Name =   "Account Number"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2275
      Columns(1).Caption=   "Currency"
      Columns(1).Name =   "Currency"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2355
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
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Dividend Currency:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Hold:"
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
      TabIndex        =   12
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Taxable:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1740
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   2760
      Y2              =   2760
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
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category Description:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1860
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
      Caption         =   "Category Code:"
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
      TabIndex        =   5
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
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim rsMain As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim strTable As String
Dim strRecNO As String
Dim OpenErr As Integer
Dim iOpenMain As Integer
Dim iOpenAdt As Integer
Function IsValid() As Integer
Dim iErr As Integer
IsValid = True
'--
If tbfld(0) = "" Then  ' Category Code
   iErr = 100
   tbfld(0).SetFocus
   GoTo Validate_Err
 End If
 tbfld(0) = UCase(tbfld(0))
 '--
 If tbfld(1) = "" Then ' Category Description
   iErr = 101
   tbfld(1).SetFocus
   GoTo Validate_Err
 End If
 tbfld(1) = Trim(tbfld(1))
 '--
Validate_Exit:
   Exit Function
'--
Validate_Err:
  'MsgBox msg, vbInformation, "Categorys"
  MsgBox iErr, vbOKOnly, "Category"
  IsValid = False
  GoTo Validate_Exit
'--
End Function

Private Sub cmdCancel_Click()
Shutdown
Unload Me
End Sub

Private Sub cmdClear_Click()

If gblOptions = 1 Then
   ClearScreen
   tbfld(0).SetFocus
Else
   ClearScreen
   tbfld(1).SetFocus
End If
End Sub

Private Sub cmdUpdate_Click()
Dim DivCurr As String
Dim i As Integer
Dim h As Boolean
Dim t As Boolean
On Error GoTo cmdUpdate_Err
If IsValid Then
   DivCurr = SSDBAccount.Columns(1).Text
  
  If HldBtn.IndexSelected = 0 Then
     h = True
  Else
     h = False
  End If
  If optBtn.IndexSelected = 0 Then
     t = True
  Else
     t = False
  End If
  
  i = RunSP(SpCon, "usp_CategoryCodeUpdate", 0, tbfld(0), tbfld(1), t, _
      h, DivCurr, gblLoginName)
  If i = 0 Then
     MsgBox "Category successfully saved"
  Else
     MsgBox "Category save failed"
  End If
    '--
  If gblOptions = 1 Then
     ClearScreen
  Else
     Shutdown
     Unload Me
  End If
End If
'---

Done:
 Exit Sub
'--
cmdUpdate_Err:
 
  MsgBox "SIS008/cmdUpdate", vbOKOnly, "Category Update"
  Shutdown
  Unload Me
End Sub
Private Sub Form_Activate()
' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 '--
 If gblOptions = 1 Then
  Me.Caption = "New Category"
End If
'--
If gblOptions = 2 Then
   Me.Caption = "Edit Category"
   tbfld(0).Enabled = False
   Set rsMain = RunSP(SpCon, "usp_CategoryFind", 1, gblFileKey)
   iOpenMain = True
   UpdateScreen
End If
 
 If OpenErr = True Then
  Shutdown
  Unload Me
End If
End Sub

Private Sub Form_Load()
Dim iDay As Integer
Dim qSQL As String
Dim indx As Integer
Dim strTmp As String
On Error GoTo FL_ERR
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
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
      frmMDI.txtStatusMsg.Refresh
   Loop
   Screen.MousePointer = vbDefault
OpenErr = False


 '--
 
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS008/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
End Sub
Private Sub UpdateScreen()
Dim i As Integer, bm As Variant
 With rsMain
    tbfld(0).Text = !CatCode
    tbfld(1).Text = !catdesc
    If !cattax = True Then
       optBtn.IndexSelected = 0
    Else
       optBtn.IndexSelected = 1
    End If
    If !Hold = True Then
       HldBtn.IndexSelected = 0
    Else
       HldBtn.IndexSelected = 1
    End If
    With SSDBAccount
         .MoveFirst
         For i = 0 To .Rows - 1
             bm = .GetBookmark(i)
             If .Columns(1).CellText(bm) = Trim(rsMain!DivCurrency) Then
                .Bookmark = .GetBookmark(i)
                 SSDBAccount = .Columns(1).CellText(bm)
                 Exit For
             End If
         Next i
    End With
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If iOpenMain = True Then rsMain.Close
Set rsMain = Nothing
SpCon.Close
End Sub

Private Sub optBtn_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  cmdUpdate.SetFocus
Case vbKeyUp
  tbfld(1).SetFocus
End Select
End Sub

Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
   Select Case Index
   Case 0
      tbfld(1).SetFocus
   Case 1
      optBtn.SetFocus
   End Select
Case vbKeyUp
   Select Case Index
   Case 1
     If gblOptions = 1 Then tbfld(0).SetFocus
   Case Else
   End Select
Case Else
End Select
End Sub

Private Sub ClearScreen()
Dim qSQL As String

  For X = 0 To 1
    tbfld(X).Text = ""
  Next
  If gblOptions = 2 Then
     Set rsMain = RunSP(SpCon, "usp_CategoryFind", 1, gblFileKey)
     UpdateScreen
     tbfld(1).SetFocus
  End If
End Sub
Private Sub Shutdown()
Unload Me
End Sub
Private Sub SSDBAccount_InitColumnProps()
On Error GoTo Err_SSDBAccount_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset
Dim i As Integer

Set adoRst = RunSP(SpCon, "usp_ListActiveAccounts", 1)
If adoRst.EOF Then
   MsgBox "Accounts were not setup" & vbCrLf & "Please do so now", vbCritical, "Account Error"
   GoTo Exit_SSDBAccount_InitColumnProps
End If

With SSDBAccount
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst!AccountNo & vbTab & adoRst!Currency & vbTab
     .AddItem StrSql
     adoRst.MoveNext
     StrSql = ""
     Loop
End With

adoRst.Close
Set adoRst = Nothing
Exit_SSDBAccount_InitColumnProps:
Exit Sub

Err_SSDBAccount_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing active accounts"
Resume Exit_SSDBAccount_InitColumnProps
End Sub

