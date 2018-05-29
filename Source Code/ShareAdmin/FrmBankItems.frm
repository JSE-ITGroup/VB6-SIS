VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBankItems 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Unmatched Bank Items"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   Icon            =   "FrmBankItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8760
      MouseIcon       =   "FrmBankItems.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Frame FmeOther 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search by Other Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   7815
      Begin VB.OptionButton OptOther 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   21
         Top             =   1200
         Width           =   495
      End
      Begin VB.OptionButton OptOther 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.OptionButton OptOther 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   19
         Top             =   120
         Width           =   375
      End
      Begin MSComCtl2.DTPicker DTPTransDate 
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16580611
         CurrentDate     =   39649
      End
      Begin VB.CommandButton CmdSearch1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6000
         MouseIcon       =   "FrmBankItems.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBTransType 
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   1200
         Width           =   2295
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
         Columns(0).Width=   3200
         Columns(0).Caption=   "Trans Type"
         Columns(0).Name =   "Trans Type"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   4048
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
      End
      Begin VB.TextBox TxtChqNo 
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Only one of the options to the left can be selected at a time"
         Height          =   495
         Left            =   5160
         TabIndex        =   22
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Select Trans type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enter Cheque No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Select Trans Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame FmeAmount 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search By Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4455
      Left            =   8160
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         MouseIcon       =   "FrmBankItems.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   3480
         Width           =   1815
      End
      Begin VB.OptionButton OptAmount 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Amounts Greater than"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   2655
      End
      Begin VB.OptionButton OptAmount 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Amounts Less than"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   2655
      End
      Begin VB.OptionButton OptAmount 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Exact Amount"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.TextBox TxtAmount 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Enter $ amount to be searched for"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter amount to search for:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBPostings 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   6
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   2064
      Columns(0).Caption=   "Trans Date"
      Columns(0).Name =   "Trans Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).NumberFormat=   "dd-mmm-yyyy"
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2275
      Columns(2).Caption=   "Chq No"
      Columns(2).Name =   "Chq No"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Amount"
      Columns(3).Name =   "Narration"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1138
      Columns(4).Caption=   "DBCR"
      Columns(4).Name =   "DBCR"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   256
      Columns(5).Width=   1482
      Columns(5).Caption=   "ItemID"
      Columns(5).Name =   "ItemID"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   13996
      _ExtentY        =   8070
      _StockProps     =   79
      Caption         =   "Click on a Bank  Item  below to begin the matching process"
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
   Begin SSDataWidgets_B.SSDBCombo SSDBAccount 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2295
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
      Columns(0).Width=   2646
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
      _ExtentX        =   4048
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
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select Account:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FrmBankItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoRst As ADODB.Recordset
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdSearch_Click()
On Error GoTo Err_CmdSearch_Click

Dim StrSql As String
Dim i As Integer
Dim Opt As Integer

If IsEmpty(SSDBAccount.SelBookmarks(0)) Then
   Beep
   MsgBox "Select an account "
   SSDBAccount.SetFocus
   GoTo Exit_CmdSearch_Click
End If
If Not IsNumber(TxtAmount) Then
   MsgBox "You are required to enter a number before clicking search", vbOKOnly, "Invalid Number"
   GoTo Exit_CmdSearch_Click
End If

For i = 0 To 2
    If OptAmount(i).Value = True Then
       Opt = i
       Exit For
    End If
Next i

SSDBPostings.RemoveAll

Set adoRst = RunSP(SpCon, "usp_SelectUnReconciledItems", 1, Opt, CDbl(SSDBAccount.Columns(0).Text), CCur(TxtAmount))
If adoRst.EOF Then
   MsgBox "Sorry, no unmatched Bank Items were found"
   GoTo Exit_CmdSearch_Click
End If
With SSDBPostings
     Do While Not adoRst.EOF
     StrSql = ""
     For i = 0 To 5
         StrSql = StrSql & adoRst(i) & vbTab
         Next i
     .AddItem StrSql
     adoRst.MoveNext
Loop
End With

Exit_CmdSearch_Click:
Exit Sub
Err_CmdSearch_Click:
MsgBox Err.Description, vbOKOnly, "Error on Search"
GoTo Exit_CmdSearch_Click

End Sub

Private Sub CmdSearch1_Click()
On Error GoTo Err_CmdSearch1_Click
Dim StrSql As String
Dim i As Integer
Dim Opt As Integer

If IsEmpty(SSDBAccount.SelBookmarks(0)) Then
   Beep
   MsgBox "Select an account "
   SSDBAccount.SetFocus
   GoTo Exit_CmdSearch1_Click
End If
For i = 0 To 2
    If OptOther(i).Value = True Then
       Opt = i
       Exit For
    End If
Next i
StrSql = Mid(SSDBTransType.Columns(0).Text, 1, 1)
SSDBPostings.RemoveAll

Set adoRst = RunSP(SpCon, "usp_SelectUnReconciledItemsO", 1, Opt, CDbl(SSDBAccount.Columns(0).Text), Format(DTPTransDate.Value, "dd-mmm-yyyy"), TxtChqno, StrSql)
If adoRst.EOF Then
   MsgBox "Sorry, no unmatched Bank Items were found"
   GoTo Exit_CmdSearch1_Click
End If
With SSDBPostings
     Do While Not adoRst.EOF
     StrSql = ""
     For i = 0 To 5
         StrSql = StrSql & adoRst(i) & vbTab
         Next i
     .AddItem StrSql
     adoRst.MoveNext
Loop
End With


Exit_CmdSearch1_Click:
Exit Sub
Err_CmdSearch1_Click:
MsgBox Err.Description, vbOKOnly, "Error on Search"
GoTo Exit_CmdSearch1_Click

End Sub

Private Sub DTPTransDate_Change()
OptOther(0).Value = True
OptOther(1).Value = False
OptOther(2).Value = False
'SSDBTransType = ""
'TxtChqNo = ""
End Sub

Private Sub Form_Activate()
DTPTransDate = Date
End Sub

Private Sub Form_Load()
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
   frmMDI.txtStatusMsg.Refresh
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
frmMDI.txtStatusMsg.Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Exit_Form_Unload

Set FrmBankItems = Nothing
If adoRst.State <> 0 Then
   adoRst.Close
End If
Set adoRst = Nothing
Exit_Form_Unload:
SpCon.Close
End Sub

Private Sub SSDBAccount_InitColumnProps()
On Error GoTo Err_SSDBAccount_InitColumnProps
Dim StrSql As String

Set adoRst = RunSP(SpCon, "usp_SelectAccounts", 1, 1)
If adoRst.EOF Then
   MsgBox "Accounts were not setup" & vbCrLf & "Please do so now", vbCritical, "Account Error"
   GoTo Exit_SSDBAccount_InitColumnProps
End If

adoRst.MoveFirst
With SSDBAccount
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst(0) & vbTab & adoRst(1) & vbTab
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
MsgBox Err & " " & Err.Description, vbOKOnly, "SSDB Combo Box load Error"
Resume Exit_SSDBAccount_InitColumnProps

End Sub
Private Sub SSDBAccount_Click()
On Error GoTo Err_SSDBAccount_Click
FmeAmount.Visible = True
FmeOther.Visible = True

Exit_SSDBAccount_Click:
Exit Sub
Err_SSDBAccount_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Unmatched Bank Items Error"
Resume Exit_SSDBAccount_Click

End Sub

Private Sub SSDBPostings_DblClick()
On Error GoTo Err_SSDBPostings_DblClick

gblFileKey = SSDBPostings.Columns(5).Text
SSDBPostings.RemoveAll
FrmBankItems.Hide
FrmBankItemsMatch.Show 0

Exit_SSDBPostings_DblClick:
Exit Sub

Err_SSDBPostings_DblClick:
MsgBox Err.Description, vbOKOnly, "Error on Selecting Item for reconciling"
GoTo Exit_SSDBPostings_DblClick
End Sub

Private Sub SSDBTransType_Change()
   'TxtChqNo = ""
   'DTPTransDate.CheckBox = False

End Sub

Private Sub SSDBTransType_Click()
   OptOther(0).Value = False
   OptOther(1).Value = False
   OptOther(2).Value = True
End Sub

Private Sub SSDBTransType_InitColumnProps()
On Error GoTo Err_SSDBTransType_InitColumnProps
Dim StrSql As String
With SSDBTransType
     StrSql = "Debit"
     .AddItem StrSql
     StrSql = "Credit"
     .AddItem StrSql
End With

Exit_SSDBTransType_InitColumnProps:
Exit Sub
Err_SSDBTransType_InitColumnProps:
MsgBox Err.Description, vbOKOnly, "Error on populating Trans Type box"
GoTo Exit_SSDBTransType_InitColumnProps

End Sub

Private Sub TxtChqNo_Change()
If Len(TxtChqno) > 0 Then
   OptOther(0).Value = False
   OptOther(1).Value = True
   OptOther(2).Value = False
   'SSDBTransType = ""
   'DTPTransDate.CheckBox = False
Else
   OptOther(1).Value = False
End If

End Sub
