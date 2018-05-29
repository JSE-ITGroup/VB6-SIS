VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmExtractDividend 
   BackColor       =   &H8000000A&
   Caption         =   "Extract Dividends"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   Icon            =   "FrmExtractDividend.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExcel 
      Caption         =   "Generate Excel File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFC0&
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
      Height          =   735
      Left            =   7920
      TabIndex        =   17
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Frame FmeFields 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fields to Show"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3735
      Left            =   7560
      TabIndex        =   5
      Top             =   0
      Width           =   2895
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Payee Name"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Client Name"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Client ID"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cheque Number"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Declaration Date"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cheque Date"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Replaced Number"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   2535
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Reconcile Status"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   2535
      End
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBDivDate 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   3015
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
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).Caption=   "Dividend Date"
      Columns(0).Name =   "Dividend Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   5318
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Frame FmeCriteria 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Selection Criteria"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7455
      Begin SSDataWidgets_B.SSDBCombo SSDBStatus 
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   2040
         Width           =   2055
         DataFieldList   =   "Column 0"
         _Version        =   196617
         DataMode        =   2
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   979
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 1"
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBReturn 
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   840
         Width           =   4815
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
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   794
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   6747
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   8493
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 1"
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBBanks 
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   4815
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
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   7461
         Columns(0).Caption=   "Branch Name"
         Columns(0).Name =   "Branch Name"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1535
         Columns(1).Caption=   "Bank Id"
         Columns(1).Name =   "Bank Id"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   8493
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 0"
      End
      Begin VB.Label LblBank 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label LblBank 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Select Bank"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Data to Return"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Select Payment Date:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmExtractDividend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExcel_Click()
On Error GoTo Err_CmdExcel_Click
Dim adoRst As ADODB.Recordset
Dim i As Integer
Dim ListOfFields As String
Dim NoOfFields As Integer
Dim DateString As String
Dim BankID As String

ListOfFields = ""
NoOfFields = 0
For i = 0 To ChkField.Count - 1
    If ChkField(i).Value = 1 Then
       ListOfFields = ListOfFields & i & ";"
       NoOfFields = NoOfFields + 1
    End If
Next i
If NoOfFields < 1 Then
   MsgBox "Please select fields to display in the spreadsheet"
   GoTo Exit_CmdExcel_Click
End If

If SSDBDivDate = vbNullString Then
   MsgBox "Please Select a dividend date first"
   GoTo Exit_CmdExcel_Click
End If

If SSDBReturn = vbNullString Then
   MsgBox "Please let me know the data to extract"
   GoTo Exit_CmdExcel_Click
End If

BankID = "0"
If SSDBBanks.Visible = True Then
   If SSDBBanks = vbNullString Then
      MsgBox "Please let me know the data to extract"
      GoTo Exit_CmdExcel_Click
   End If
   BankID = SSDBBanks.Columns(0).Text
End If

Set adoRst = RunSP(SpCon, "usp_DividendDataExtract", 1, ListOfFields, NoOfFields, SSDBDivDate.Columns(0).Text, _
             SSDBReturn.Columns(0).Text, BankID, SSDBStatus.Columns(0).Text)
             
If adoRst.EOF Then
   MsgBox "Sorry, No records were found"
Else
   Call ExportToExcel(adoRst)
End If
adoRst.Close
Set adoRst = Nothing
Exit_CmdExcel_Click:
Exit Sub

Err_CmdExcel_Click:
MsgBox Err.Description, vbOKOnly, "Dividend Data Excel File Creation"
GoTo Exit_CmdExcel_Click


End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load

frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
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
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

Exit_Form_Load:
Exit Sub
Err_Form_Load:
MsgBox Err.Description, vbOKOnly, "Dividend Data Etract Form load"
GoTo Exit_Form_Load

End Sub

Private Sub SSDBBanks_InitColumnProps()
Dim rsBanks As ADODB.Recordset
Dim sRowinfo As String

SSDBBanks.RemoveAll
  Set rsBank = RunSP(SpCon, "usp_Banks", 1)
  If rsBank.EOF Then
     GoTo Exit_SSDBBanks_InitColumnProps
  End If

  With rsBank
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        sRowinfo = !BnkName & vbTab & !BankID
        SSDBBanks.AddItem sRowinfo
       .MoveNext
      Loop
    End If
  End With
  '--

Exit_SSDBBanks_InitColumnProps:
rsBank.Close
Exit Sub

End Sub

Private Sub SSDBDivDate_InitColumnProps()
Dim adoDivDate As ADODB.Recordset

Set adoDivDate = RunSP(SpCon, "usp_DivDateData", 1)

Do While Not adoDivDate.EOF
   With SSDBDivDate
        .AddItem Format(adoDivDate!ChqDate, "dd-mmm-yyyy")
   End With
   adoDivDate.MoveNext
Loop

adoDivDate.Close
Set adoDivDate = Nothing

End Sub

Private Sub SSDBReturn_Click()
With SSDBReturn
     If .Columns(0).Text = "E" Then
        LblBank(0).Visible = True
        SSDBBanks.Visible = True
     Else
       LblBank(0).Visible = False
        SSDBBanks.Visible = False
     End If
End With
End Sub

Private Sub SSDBReturn_InitColumnProps()
With SSDBReturn
     .RemoveAll
     StrSql = "A" & vbTab & "All Payments" & vbTab
     .AddItem StrSql
     StrSql = "B" & vbTab & "All Bank Payments"
     .AddItem StrSql
     StrSql = "C" & vbTab & "All Cheque Payments"
     .AddItem StrSql
     StrSql = "D" & vbTab & "All Other Banks Payments"
     .AddItem StrSql
     StrSql = "E" & vbTab & "Payments with a Specific Bank"
     .AddItem StrSql
     StrSql = "F" & vbTab & "Non Payments"
      .AddItem StrSql
    .MoveFirst
    .SelBookmarks.RemoveAll
     For X = 0 To .Rows - 1
     If .Columns(0).Text = "A" Then
        .Text = .Columns(1).Text
        .SelBookmarks.Add .Bookmark
        GoTo CloseRecordset
     End If
     .MoveNext
     Next X
End With

CloseRecordset:

End Sub

Private Sub SSDBStatus_InitColumnProps()
With SSDBStatus
     .RemoveAll
     StrSql = "A" & vbTab & "All Statuses"
     .AddItem StrSql
     StrSql = "B" & vbTab & "All Paid Cheques"
     .AddItem StrSql
     StrSql = "C" & vbTab & "All UnPaid Cheques"
     .AddItem StrSql

    .MoveFirst
    .SelBookmarks.RemoveAll
     For X = 0 To .Rows - 1
     If .Columns(0).Text = "A" Then
        .Text = .Columns(1).Text
        .SelBookmarks.Add .Bookmark
        GoTo CloseRecordset
     End If
     .MoveNext
     Next X
End With

CloseRecordset:

End Sub
