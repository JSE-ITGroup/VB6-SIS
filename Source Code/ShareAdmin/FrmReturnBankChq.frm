VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmReturnBankChq 
   BackColor       =   &H00404080&
   Caption         =   "Returned Cheque Posting "
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   Icon            =   "FrmReturnBankChq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptSearch 
      BackColor       =   &H00004080&
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   7320
      TabIndex        =   27
      Top             =   6120
      Width           =   255
   End
   Begin VB.OptionButton OptSearch 
      BackColor       =   &H00004080&
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   7320
      TabIndex        =   26
      Top             =   5640
      Width           =   255
   End
   Begin VB.TextBox TxtAccount 
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      ToolTipText     =   "Enter the Bank Account number"
      Top             =   6120
      Width           =   4335
   End
   Begin VB.Frame FmeNameSearch 
      BackColor       =   &H00FF8080&
      Caption         =   "Search By Name Options"
      Height          =   2655
      Left            =   0
      TabIndex        =   20
      Top             =   2280
      Width           =   7575
      Begin SSDataWidgets_B.SSDBGrid SSDBShareholders 
         Height          =   2415
         Left            =   0
         TabIndex        =   21
         Top             =   240
         Width           =   7455
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   3
         BackColorEven   =   16761024
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   3281
         Columns(0).Caption=   "Client Name"
         Columns(0).Name =   "Client Name"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Client ID"
         Columns(1).Name =   "Client ID"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Caption=   "Shares"
         Columns(2).Name =   "Shares"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   13150
         _ExtentY        =   4260
         _StockProps     =   79
         Caption         =   "List of Shareholders with Names like "
         BackColor       =   -2147483633
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox TxtName 
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      ToolTipText     =   "Names of Individuals are to be entered Last name, then first name e.g. ""Williams David"""
      Top             =   5640
      Width           =   4335
   End
   Begin VB.Frame FmeDetails 
      BackColor       =   &H8000000B&
      Caption         =   "Details"
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7575
      Begin VB.TextBox TxtBankName 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   7335
      End
      Begin VB.TextBox TxtCurrency 
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox TxtType 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox TxtAmount 
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox TxtChequeNo 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox TxtChqDate 
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox TxtDecDate 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtShareholder 
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2640
         Width           =   7335
      End
      Begin VB.Label LblCurrency 
         BackColor       =   &H8000000B&
         Caption         =   "Currency"
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
         Left            =   3960
         TabIndex        =   16
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label LblType 
         BackColor       =   &H8000000B&
         Caption         =   "Type:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label LblAmount 
         BackColor       =   &H8000000B&
         Caption         =   "Amount"
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
         Left            =   3960
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label LblChequeNo 
         BackColor       =   &H8000000B&
         Caption         =   "Cheque Number:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label LblChqDate 
         BackColor       =   &H8000000B&
         Caption         =   "Cheque Date:"
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
         Left            =   3960
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label LblDecDate 
         BackColor       =   &H8000000B&
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
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame FmeCommands 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   7455
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Batang"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5640
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton CmdPost 
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Batang"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton CmdFind 
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Batang"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBChqList 
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   5040
      Width           =   7575
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
      BackColorEven   =   12648384
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   3200
      Columns(0).Caption=   "Account"
      Columns(0).Name =   "Account"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Cheque Number"
      Columns(1).Name =   "Cheque Number"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Chq Date"
      Columns(2).Name =   "Chq Date"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   7
      Columns(2).NumberFormat=   "dd-mmm-yyyy"
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Amount"
      Columns(3).Name =   "Amount"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   6
      Columns(3).NumberFormat=   "CURRENCY"
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "Type"
      Columns(4).Name =   "Type"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "Payee Name"
      Columns(5).Name =   "Payee Name"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   13361
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Column 1"
   End
   Begin VB.Label LblName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter Account Number:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label LblName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter Shareholder's Name:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   2535
   End
End
Attribute VB_Name = "FrmReturnBankChq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim WrkTitle As String

Private Sub CmdExit_Click()
On Error GoTo Err_CmdExit_Click

Unload Me

Exit_CmdExit_Click:
Exit Sub

Err_CmdExit_Click:
MsgBox Err.Description, vbOKOnly, "Returned Cheques Exit"
GoTo Exit_CmdExit_Click

End Sub

Private Sub cmdFind_Click()
On Error GoTo Err_CmdFind_Click

If OptSearch(0).Value = True Then
   If TxtName = vbNullString Then
      MsgBox "To search by Name, please enter the search criteria", vbOKOnly, "Name field was left Blank"
      GoTo Exit_CmdFind_Click
   End If
   LoadShareholders
   GoTo Exit_CmdFind_Click
Else
   If TxtAccount = vbNullString Then
      MsgBox "To search by Account, please enter the search criteria", vbOKOnly, "Account No field was left Blank"
      GoTo Exit_CmdFind_Click
   End If
   LoadShareholdersA
   GoTo Exit_CmdFind_Click
End If

Exit_CmdFind_Click:
Exit Sub

Err_CmdFind_Click:
MsgBox Err.Description, vbOKOnly, "CmdFind Error"
GoTo Exit_CmdFind_Click

End Sub

Private Sub CmdPost_Click()
On Error GoTo Exit_CmdPost_Click
Dim i As Integer

i = RunSP(SpCon, "usp_PostReturnedCheque", 0, CLng(SSDBChqList.Columns(0).Text), SSDBChqList.Columns(1).Text, Format(SSDBChqList.Columns(2).Text, "dd-mmm-yyyy"), gblLoginName)
If i <> 0 Then
   MsgBox "There was an error and the posting was not completed"
Else
   MsgBox "Posting was successful"
End If

Exit_CmdPost_Click:
Exit Sub

End Sub

Private Sub Form_Load()
On Error GoTo FL_ERR
'--
   csvCenterForm Me, gblMDIFORM
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
   Loop
   Screen.MousePointer = vbDefault
  
   '-------------------------------------
   '-- Initialize License Details -------
   '-------------------------------------
   '--
 WrkTitle = "List of Shareholders with Name like"
 '--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox Err.Description, vbOKOnly, "Returned Cheques Form Load"
  GoTo FL_Exit
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close

End Sub

Private Sub SSDBChqList_Click()
LoadDetails
End Sub


Private Sub SSDBShareholders_Click()
Dim adoRst As ADODB.Recordset
Dim StrSql As String

Set adoRst = RunSP(SpCon, "usp_FindCheques", 1, CLng(SSDBShareholders.Columns(1).Text))
If adoRst.EOF Then
   MsgBox "No Dividend payments were found for the selected shareholder"
   GoTo Exit_SSDBShareholders_Click
End If

SSDBChqList = vbNullString

With SSDBChqList
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst!AccountNo & vbTab & adoRst!ChqNum & vbTab
     StrSql = StrSql & adoRst!ChqDat & vbTab & adoRst!ChqAmt & vbTab
     If adoRst!PayTyp = "D" Then
        StrSql = StrSql & "Dividend" & vbTab
     Else
        StrSql = StrSql & "Capital Distribution" & vbTab
     End If
     StrSql = StrSql & adoRst!PAYEENAME & vbTab
     .AddItem StrSql
     adoRst.MoveNext
     Loop
End With

adoRst.Close

Exit_SSDBShareholders_Click:
Set adoRst = Nothing
Exit Sub

End Sub

Private Sub LoadDetails()
Dim adoRst As ADODB.Recordset
Dim adoRst1 As ADODB.Recordset

TxtBank = vbNullString
TxtDecDate = vbNullString
TxtChqDate = vbNullString
TxtChequeNo = vbNullString
TxtAmount = vbNullString
TxtType = vbNullString

Set adoRst = RunSP(SpCon, "usp_ReturnChqDetails", 1, CLng(SSDBChqList.Columns(0).Text), SSDBChqList.Columns(1).Text, Format(SSDBChqList.Columns(2).Text, "dd-mmm-yyyy"))
If adoRst.State = adStateClosed Then
   MsgBox "The Search criteria used returned no records"
   GoTo Exit_LoadDetails
End If
Set adoRst1 = RunSP(SpCon, "usp_ReturnChqDetails2", 1, SSDBChqList.Columns(1).Text)
With adoRst
TxtShareholder = !PAYEENAME
TxtDecDate = Format(!DecDate, "dd-mmm-yyyy")
TxtChqDate = Format(!ChqDat, "dd-mmm-yyyy")
TxtChequeNo = !ChqNum
TxtAmount = !ChqAmt
TxtCurrency = !Currency
If !PayTyp = "D" Then
   TxtType = "Dividend"
Else
   TxtType = "Capital Distribution"
End If
.Close
End With

TxtBankName = Trim(adoRst1!BnkName) & " " & Trim(adoRst1!BNKADDR1)
adoRst1.Close

Exit_LoadDetails:
Set adoRst = Nothing
Set adoRst1 = Nothing
Exit Sub

End Sub
Private Sub LoadShareholders()
Dim adoRst As ADODB.Recordset
Set adoRst = RunSP(SpCon, "usp_FindShareholders", 1, TxtName)
WrkTitle = "List of Shareholders with Names like:"
SSDBShareholders.RemoveAll
SSDBShareholders.Caption = WrkTitle & " " & TxtName

If adoRst.EOF Then
   MsgBox "The Search criteria specified returned no records"
   GoTo Exit_LoadShareholders
End If
With SSDBShareholders
     Do While Not adoRst.EOF
        StrSql = adoRst!CliName & vbTab & adoRst!ClientID & vbTab & adoRst!shares
        .AddItem StrSql
        adoRst.MoveNext
     Loop
End With

adoRst.Close
Exit_LoadShareholders:
Set adoRst = Nothing
Exit Sub

End Sub

Private Sub TxtName_GotFocus()
OptSearch(0).Value = True
OptSearch(1).Value = False
End Sub
Private Sub TxtAccount_GotFocus()
OptSearch(0).Value = False
OptSearch(1).Value = True
End Sub
Private Sub LoadShareholdersA()
Dim adoRst As ADODB.Recordset
Set adoRst = RunSP(SpCon, "usp_FindShareholdersA", 1, TxtAccount)
WrkTitle = "List of Shareholders with Mandate Account number:"
SSDBShareholders.RemoveAll
SSDBShareholders.Caption = WrkTitle & " " & TxtAccount
If adoRst.EOF Then
   MsgBox "The Search criteria specified returned no records"
   GoTo Exit_LoadShareholdersA
End If
With SSDBShareholders
     Do While Not adoRst.EOF
        StrSql = adoRst!CliName & vbTab & adoRst!ClientID & vbTab & adoRst!shares
        .AddItem StrSql
        adoRst.MoveNext
     Loop
End With

adoRst.Close
Exit_LoadShareholdersA:
Set adoRst = Nothing
Exit Sub

End Sub

