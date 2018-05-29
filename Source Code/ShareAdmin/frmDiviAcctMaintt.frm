VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmDiviAcctMaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dividend Account Update"
   ClientHeight    =   2625
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6660
   Icon            =   "frmDiviAcctMaintt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6660
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5640
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4560
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtAccountNo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Enter a valid Nine Digit Account No."
      Top             =   720
      Width           =   1935
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBCurrency 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   1080
      Width           =   2175
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
      Columns(0).Width=   2143
      Columns(0).Caption=   "Currency Code"
      Columns(0).Name =   "Account Number"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4022
      Columns(1).Caption=   "Currency Description"
      Columns(1).Name =   "Currency"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   3836
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
   Begin SSDataWidgets_B.SSDBCombo SSDBStatus 
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   1560
      Width           =   2175
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
      Columns(0).Width=   2143
      Columns(0).Caption=   "Status"
      Columns(0).Name =   "Account Number"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   3836
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
      Caption         =   "Status:"
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
      TabIndex        =   8
      Top             =   1680
      Width           =   1740
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
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
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Currency:"
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
      TabIndex        =   5
      Top             =   1200
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
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account No:"
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
      TabIndex        =   3
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
      Left            =   600
      TabIndex        =   7
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmDiviAcctMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As ADODB.Recordset
Dim SpCon As ADODB.Connection
Function IsValid() As Integer
Dim msg As String

IsValid = True
'--
If TxtAccountNo = "" Then  ' AccountNo
   TxtAccountNo.SetFocus
   msg = "Account No is missing"
   GoTo Validate_Err
 End If

 '--
 If SSDBCurrency = "" Then ' Currency
   SSDBCurrency.SetFocus
   msg = "Currency Code is Missing"
   GoTo Validate_Err
 End If

If SSDBStatus = "" Then ' Currency
   SSDBStatus.SetFocus
   msg = "Status is Missing"
   GoTo Validate_Err
 End If
 
Validate_Exit:
   Exit Function
'--
Validate_Err:
  MsgBox msg, vbInformation, "Bank Account"
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
Else
   UpdateScreen
End If

TxtAccountNo.SetFocus
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer
Dim i As Integer
Dim iStatus As String

On Error GoTo cmdUpdate_Err
If IsValid Then
  '--
  If SSDBStatus.Columns(0).Text = "Active" Then
     iStatus = "A"
  Else
     iStatus = "C"
  End If
  
  i = RunSP(SpCon, "usp_AccountsUpdate", 0, TxtAccountNo, SSDBCurrency.Columns(0).Text, iStatus, gblLoginName)
  If i <> 0 Then
      MsgBox "There was an error saving the changes. Please re-try"
      GoTo Done
  Else
      MsgBox "Update was successful"
  End If
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
  Shutdown
  Unload Me
End Sub
Private Sub Form_Activate()
' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 '--

Call LoadSSDBs

If gblOptions = 1 Then
  Me.Caption = "New Bank Account"
Else
  Me.Caption = "Edit Bank Account"
  UpdateScreen
End If

End Sub

Private Sub Form_Load()
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

'--
   
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "Error on loading frmSisAcctMaint"
  On Error Resume Next
  Resume FL_Exit
   
End Sub
Private Sub UpdateScreen()
Dim i As Integer, bm As Variant
Dim iStatus As String
Set rsMain = RunSP(SpCon, "usp_AccountsEdit", 1, CDbl(gblFileKey))
 With rsMain
    TxtAccountNo.Text = !AccountNo
     SSDBCurrency.MoveFirst
     For i = 0 To SSDBCurrency.Rows - 1
         bm = SSDBCurrency.GetBookmark(i)
         If SSDBCurrency.Columns(0).CellText(bm) = !Currency Then
            SSDBCurrency.Bookmark = SSDBCurrency.GetBookmark(i)
             SSDBCurrency = SSDBCurrency.Columns(0).CellText(bm)
         Exit For
         End If
     Next i
     If !Status = "A" Then
        iStatus = "Active"
     Else
        iStatus = "Closed"
     End If
    With SSDBStatus
     .MoveFirst
     For i = 0 To .Rows - 1
         bm = .GetBookmark(i)
         If .Columns(0).CellText(bm) = iStatus Then
            .Bookmark = .GetBookmark(i)
             SSDBStatus = .Columns(0).CellText(bm)
         Exit For
         End If
     Next i
     End With
 End With
End Sub

Private Sub ClearScreen()
TxtAccountNo = ""
SSDBCurrency = ""
SSDBStatus = ""
End Sub

Private Sub Shutdown()
SpCon.Close
End Sub

Private Sub LoadSSDBs()
Dim adoRst As New ADODB.Recordset
Dim sRowinfo As String

Set adoRst = RunSP(SpCon, "usp_CurrencyList", 1)

With adoRst
      SSDBCurrency.RemoveAll
      If Not .EOF Then
        Do While Not .EOF
          sRowinfo = !CurrencyCode & vbTab & !CurrencyDesc & vbTab
          SSDBCurrency.AddItem sRowinfo
         .MoveNext
        Loop
      End If
End With
adoRst.Close
Set adoRst = Nothing


With SSDBStatus
     .RemoveAll
     sRowinfo = "Active"
     .AddItem sRowinfo
     sRowinfo = "Closed"
     .AddItem sRowinfo
End With
End Sub
