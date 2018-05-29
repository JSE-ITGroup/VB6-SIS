VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS014 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Information Details"
   ClientHeight    =   5280
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "SIS014.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6795
   Begin VB.TextBox tb 
      Height          =   375
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   29
      Top             =   4080
      Width           =   4575
   End
   Begin SSDataWidgets_A.SSDBOptSet optBtn 
      Height          =   495
      Index           =   0
      Left            =   5160
      TabIndex        =   27
      Top             =   600
      Width           =   1335
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
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Enter Payment Type from list"
      Top             =   600
      Width           =   2055
      DataFieldList   =   "Column 1"
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
      FieldSeparator  =   ","
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2884
      Columns(0).Caption=   "Payment Types"
      Columns(0).Name =   "Payment Types"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   25
      Columns(1).Width=   1482
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   1
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3600
      TabIndex        =   20
      Top             =   4920
      Width           =   975
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Enter Declaration date in format dd-mm-yyyy"
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5760
      TabIndex        =   9
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
      Top             =   1560
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
      Left            =   1920
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Format          =   "$#,##0.0000000;($#,##0.0000000)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   5
      Top             =   2640
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   13
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
      DataFieldList   =   "Column 1"
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
      FieldSeparator  =   ","
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2884
      Columns(0).Caption=   "Income Types"
      Columns(0).Name =   "Income Types"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   18
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   1
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_A.SSDBOptSet optBtn 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   3600
      Width           =   1320
      _Version        =   196611
      _ExtentX        =   2328
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "&No"
      BackColor       =   -2147483643
      Cols            =   2
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
      Buttons.Button(0).ButtonToColRight=   43
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(1).OptionValue=   "0"
      Buttons.Button(1).Caption=   "&No"
      Buttons.Button(1).Mnemonic=   78
      Buttons.Button(1).Value=   -1  'True
      Buttons.Button(1).TextLeft=   59
      Buttons.Button(1).TextRight=   73
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   44
      Buttons.Button(1).ButtonRight=   57
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   75
      Buttons.Button(1).PictureRight=   74
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   44
      Buttons.Button(1).ButtonToColRight=   87
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).ButtonBitmapID=   2
      Buttons.Button(1).Column=   1
   End
   Begin VB.Label lblLabels 
      Caption         =   "First Payment  for the Tax Year:"
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
      Index           =   14
      Left            =   360
      TabIndex        =   28
      Top             =   3480
      Width           =   1455
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Payment From:"
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
      Index           =   11
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Payment % of Par:"
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
      TabIndex        =   23
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Payment Per Share:"
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      Left            =   3720
      TabIndex        =   19
      Top             =   1920
      Width           =   1335
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
      Left            =   3960
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label lblParValue 
      Caption         =   "Par Value"
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
      Left            =   5160
      TabIndex        =   16
      Top             =   1560
      Width           =   1335
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
      Left            =   5160
      TabIndex        =   15
      Top             =   1920
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
      TabIndex        =   14
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Declaration Date:"
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
      TabIndex        =   13
      Top             =   1080
      Width           =   1740
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
      Caption         =   "Type of Payment:"
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
      Width           =   1740
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
      Left            =   240
      TabIndex        =   10
      Top             =   1560
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
      TabIndex        =   17
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, iEOF As Integer
Dim SpCon As ADODB.Connection
Dim rsCmp As ADODB.Recordset
Dim rsDiv As ADODB.Recordset
Dim rsAdt As ADODB.Recordset
Dim iOpenMain As Integer
Dim iOpenAdt As Integer
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
If dbc(0) = "" Then   ' Type of payment
   iErr = "Please select type of payment"
   MsgBox iErr, , "Payment Information Entry"
   dbc(0).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If meb(0) = "" Then 'declaration date
   iErr = "Declaration Date is empty"
   MsgBox iErr, , "Payment Information Entry"
   meb(0).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(0)) Then
      iErr = "Invalid Declaration Date"
      MsgBox iErr, , "Payment Information Entry"
      meb(0).SetFocus
      GoTo Validate_Exit
   Else
      If iMode = 1 Or iMode = 2 Then
        dtefld = meb(0).Text
        With rsDiv
          '.MoveFirst
          Do While Not .EOF
             If dtefld <= !DecDate And !CLOSED = True Then
               If dbc(0).Columns(1).Text = !PayTyp Then
                  iErr = "Invalid Payment Type"
                  MsgBox iErr, , "Dividend Information Entry"
                  dbc(0).SetFocus
                  Exit Do
                End If
             End If
             .MoveNext
          Loop
          If iErr = "Invalid Payment Type" Then GoTo Validate_Exit
        End With
      End If
   End If
 End If
 '--
 If meb(1) = "" Then ' date on record
   iErr = "Record Date is empty"
   MsgBox iErr, , "Payment Information Entry"
   meb(1).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(1)) Then
      iErr = "Invalid Record Date"
      MsgBox iErr, , "Payment Information Entry"
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If meb(2) = "" Then 'Date Payable
   iErr = "Date payable is empty"
   MsgBox iErr, , "Payment Information Entry"
   meb(2).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(2)) Then
      iErr = "Payable Date not a valid date"
      MsgBox iErr, , "Payment Information Entry"
      meb(2).SetFocus
      GoTo Validate_Exit
   Else
      If DateValue(meb(2).Text) < DateValue(meb(0).Text) Then
         iErr = "Date Payable before Declaration Date"
         MsgBox iErr, , "Payment Information Entry"
         meb(2).SetFocus
         GoTo Validate_Exit
      End If
   End If
 End If
 '--
 If meb(3) = "" And meb(4) = "" Then 'payment Data
   iErr = "Payment data is empty"
   MsgBox iErr, , "Payment Information Entry"
   meb(3).SetFocus
   GoTo Validate_Exit
 End If
 If meb(3) <> "" And Not IsNumeric(meb(3)) Then
      iErr = "Payment per share is not numeric"
      MsgBox iErr, , "Payment per share"
      meb(3).SetFocus
      GoTo Validate_Exit
 End If
 If meb(4) <> "" And Not IsNumeric(meb(4)) Then
      iErr = "% per share is not numeric"
      meb(4).SetFocus
      GoTo Validate_Exit
 End If
 '--
 If dbc(1) = "" Then ' Income type
    iErr = "Invalid Income Type"
    MsgBox iErr, , "Payment Information Entry"
    dbc(1).SetFocus
    GoTo Validate_Exit
 End If
 IsValid = True
Validate_Exit:
   Exit Function
'--
Validate_Err:
  'MsgBox msg, vbInformation, "Users"
  MsgBox iErr, "Company"
  IsValid = False
  GoTo Validate_Exit
'--
End Function

Private Sub cmdCancel_Click()
If rsDiv Is Nothing Then Else rsDiv.Close
If rsCmp Is Nothing Then Else rsCmp.Close
If rsAdt Is Nothing Then Else rsAdt.Close

Set rsDiv = Nothing
Set rsCmp = Nothing
Set rsAdt = Nothing
'''set cnn = nothing
iEOF = True
Unload Me
Set frmSIS014 = Nothing
frmSIS013.Visible = True
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
Dim Resp As Integer

On Error GoTo cmdUpdate_Err

If IsValid Then
  '--
  strTable = 5  'DivRef
  '---------------------------------------
  '-- convert dec date to YYYYMMDD format ----
  '-- and store with payment type in strRecNo
  '--------------------------------------------
  strMth = CStr(Month(meb(0)))
  str = Mid(strMth, 2, 1)
  If str = " " Then strMth = "0" & strMth
  strDay = CStr(Day(meb(0)))
  str = Mid(strDay, 2, 1)
  If str = " " Then strDay = "0" & strDay
  strRecNO = UCase(Year(meb(0)) & strMth _
             & strDay)
  '---
  strChg = 0
  
  Resp = RunSP(SpCon, "usp_DivRefUpdate", 0, CInt(strTable), strRecNO, gblLoginName, _
         Format(meb(0).Text, "dd-mmm-yyyy"), dbc(0).Columns(1).Text, _
         Format(meb(1).Text, "dd-mmm-yyyy"), Format(meb(2).Text, "dd-mmm-yyyy"), _
         CCur(meb(3).Text), CDbl(meb(4).Text), dbc(1).Columns(1).Text, tb, _
         CInt(optBtn(1).OptionValue), CInt(optBtn(0).OptionValue))
  
  If Resp = 0 Then
     MsgBox "Dividend record successfully updated"
     ClearScreen
     UpdateScreen
  Else
     MsgBox "Update failed"
  End If
End If
'---
Done:
 Exit Sub
'--
cmdUpdate_Err:
  
  MsgBox Err.Description, vbOKOnly, "SIS014/cmdUpdate"
  'SpCon.RollbackTrans
  cmdCancel_Click
End Sub

Private Sub dbc_InitColumnProps(Index As Integer)
Select Case Index
Case 0 'Init Payment types
With dbc(0)
  .RemoveAll
  .AddItem "Dividend,D"
  .AddItem "Capital Distrubition,C"
End With
Case 1 'Init Income types
With dbc(1)
  .RemoveAll
  .AddItem "Franked Income,F"
  .AddItem "Unfranked Income,U"
End With
Case Else
End Select
End Sub

Private Sub dbc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  If Index = 0 Then
   meb(0).SetFocus
  Else
    tb.SetFocus
  End If
 Case vbKeyUp
    KeyCode = 0
   If Index = 1 Then meb(4).SetFocus
 End Select
End Sub

Private Sub Form_Activate()
If OpenErr = True Then
  If iOpenMain = True Then
    rsDiv.Close
  End If
  If iOpenCmp = True Then
    rsCmp.Close
  End If
  Set rsCmp = Nothing
  Set rsAdt = Nothing
  '''set cnn = nothing
  Set frmSIS014 = Nothing
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
Set SpCon = New ADODB.Connection
'On Error GoTo FL_ERR
iEOF = False
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

   csvCenterForm Me, gblMDIFORM
   OpenErr = False
   iOpenMain = False
   iOpenAdt = False
   iOpenCmp = False
   '-----------------------
   '-- open tables --------
   '-----------------------
   Set rsDiv = RunSP(SpCon, "usp_DivRef", 1)
   iOpenMain = True
   Set rsCmp = RunSP(SpCon, "usp_Company", 1)
   iOpenCmp = True
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   lblParValue.Caption = rsCmp!PARVALUE
   lblTaxFree.Caption = rsCmp!TAXFREELIMIT
   '--
   If rsDiv.EOF = True Then
      iMode = 0
   Else
      iMode = 1
   End If
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS014/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
End Sub
Private Sub UpdateScreen()
Dim i As Integer, bm As Variant
With rsDiv
  If Not .EOF Then
    If CurRec Then
      dbc(0).MoveFirst
      For i = 0 To dbc(0).Rows - 1
        bm = dbc(0).GetBookmark(i)
        If dbc(0).Columns(1).CellText(bm) = !PayTyp Then
          dbc(0).Bookmark = dbc(0).GetBookmark(i)
          dbc(0) = dbc(0).Columns(0).CellText(bm)
          Exit For
        End If
      Next i
      dbc(0).Enabled = False
      meb(0).Text = !DecDate
      meb(0).Enabled = False
      meb(1).Text = !RecDate
      meb(2).Text = !ChqDate
      If Not IsNull(!PAYAMT) Then meb(3).Text = !PAYAMT
      If Not IsNull(!PAYPER) Then meb(4).Text = !PAYPER
      dbc(1).MoveFirst
      For i = 0 To dbc(1).Rows - 1
        bm = dbc(1).GetBookmark(i)
        If dbc(1).Columns(1).CellText(bm) = !INCTYP Then
          dbc(1).Bookmark = dbc(1).GetBookmark(i)
          dbc(1) = dbc(1).Columns(0).CellText(bm)
          Exit For
        End If
      Next i
      If !FSTRUN = True Then
          optBtn(1).IndexSelected = 0
      Else
          optBtn(1).IndexSelected = 1
      End If
     ' If IsNull(!Remarks) Then
      '   tb = " "
      'Else
         tb = IsNullMove(!Remarks)
      'End If
      
   Else
     iMode = 2
     meb(0).Enabled = True
     dbc(0).Enabled = True
     optBtn(1).IndexSelected = 1
   End If
Else
   meb(0).Enabled = True
   dbc(0).Enabled = True
   optBtn(1).IndexSelected = 1
End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
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

Private Sub meb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  If Index = 4 Then
     dbc(1).SetFocus
  Else
     meb(Index + 1).SetFocus
  End If
Case vbKeyUp
KeyCode = 0
  If Index = 0 Then
     If gblOptions = 0 Then dbc(0).SetFocus
  Else
    If Index <> 0 Then meb(Index - 1).SetFocus
  End If
Case Else
End Select
End Sub

Private Sub ClearScreen()
Dim i As Integer
For i = 0 To 1
  dbc(i) = ""
Next
For i = 0 To 4
  meb(i).Mask = ""
  meb(i).Text = ""
Next
tb = ""
optBtn(0).IndexSelected = 1
End Sub

Private Function CurRec()
CurRec = False
'rsDiv.Requery
With rsDiv
   If .EOF Then Exit Function
   '.MoveFirst
   While Not .EOF
     If !CLOSED = False Then
        CurRec = True
        Exit Function
     Else
        .MoveNext
     End If
   Wend
End With
End Function

Private Sub meb_LostFocus(Index As Integer)
Select Case Index
Case 3
   If meb(3).Text <> "" Then
      meb(4).Text = Val(meb(3).Text) / rsCmp!PARVALUE
   End If
Case 4
    If meb(4).Text <> "" Then
      meb(3).Text = Val(meb(4).Text) * rsCmp!PARVALUE
   End If
Case Else
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Or UnloadMode = 1 Then
        'the X has been clicked or the user has pressed Alt+F4
        cmdCancel = True
    End If
End Sub
