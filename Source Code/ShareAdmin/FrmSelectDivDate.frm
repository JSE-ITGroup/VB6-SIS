VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmSelectDivDate 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Dividend Reconciliation Report"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   Icon            =   "FrmSelectDivDate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FmeReconciliation 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Report on "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   3375
      Begin VB.OptionButton OptRecon 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Unreconciled Items"
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton OptRecon 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Reconciled Items"
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
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.OptionButton OptDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select All Dividend dates"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBList 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
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
      Columns(0).Width=   3200
      Columns(0).Caption=   "Chq Date"
      Columns(0).Name =   "Chq Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Dividend Type"
      Columns(1).Name =   "Dividend Type"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   6588
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBAccount 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   600
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
      Columns(0).Width=   5741
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select Account"
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
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select Dividend to Report"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSelectDivDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdStart_Click()
On Error GoTo Exit_CmdStart_Click

If SSDBList = vbNullString Then
   MsgBox "Please select a dividend date to reconcile first"
   SSDBList.SetFocus
   GoTo Exit_CmdStart_Click
End If

If SSDBAccount = vbNullString Then
   MsgBox "Please select an account first"
   SSDBAccount.SetFocus
   GoTo Exit_CmdStart_Click
End If
If FmeReconciliation.Visible = True Then
   Call DetailedReconReports
   GoTo Exit_CmdStart_Click
Else
   gblOptions = 9
   gblDate = SSDBList.Columns(0).Text
   gblFileKey = SSDBAccount.Columns(0).Text
   gblHold = SSDBList.Columns(1).Text
End If

FrmReportView.Show 0

Exit_CmdStart_Click:
Exit Sub
Err_CmdStart_Click:
MsgBox Err.Description, vbOKOnly, "Dividend Reconciliation report Error"
Resume Exit_CmdStart_Click
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load

csvCenterForm Me, gblMDIFORM
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
MsgBox Err.Description, vbOKOnly, "Reverse Posted Dividend Form Load"
GoTo Exit_Form_Load

End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
Set FrmSelectDivDate = Nothing
End Sub

Private Sub SSDBList_Click()
OptDate.Value = False
End Sub

Private Sub SSDBList_InitColumnProps()
Dim adoList As ADODB.Recordset
Dim StrSql As String

Set adoList = RunSP(SpCon, "usp_SelectPostedDividend", 1)

Do While Not adoList.EOF
   With SSDBList
        StrSql = Format(adoList!ChqDate, "dd-mmm-yyyy") & vbTab & adoList!PayTyp
        .AddItem StrSql
   End With
   adoList.MoveNext
Loop

adoList.Close
Set adoList = Nothing

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
Private Sub DetailedReconReports()
Dim repSISRept As New SISRepts

Set repSISRept = New SISRepts
If OptDate.Value = True Then
      repSISRept.Date1 = "01-Jan-1900"
   Else
      repSISRept.Date1 = SSDBList.Columns(0).Text
End If
repSISRept.siteid = SSDBAccount.Columns(0).Text
If OptRecon(0).Value = True Then
   repSISRept.OptNo = 1
Else
   repSISRept.OptNo = 0
End If

repSISRept.LoginId = gblFileName
repSISRept.ReportType = 9
repSISRept.ReportNumber = 85
repSISRept.RunShareHolderReport

End Sub
