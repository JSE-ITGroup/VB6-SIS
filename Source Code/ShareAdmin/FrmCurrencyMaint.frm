VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmCurrencyMaint 
   Caption         =   "Currency Maintenance"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FrmCurrencyMaint.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3000
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox TxtDescription 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox TxtCurrencyCode 
      Height          =   375
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBCurrencyType 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
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
   Begin VB.Label Label1 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label LblDescription 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label LblCurrencyCode 
      Caption         =   "Currency Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "FrmCurrencyMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim rsMain As New ADODB.Recordset

Private Sub cmdCancel_Click()
Shutdown
Unload Me
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer
Dim i As Integer

On Error GoTo cmdUpdate_Err

If Len(TxtCurrencyCode) < 3 Then
   MsgBox "Please correct currency code. This should be 3 characters"
   TxtCurrencyCode.SetFocus
   GoTo Done
End If

If Len(TxtDescription) < 9 Then
   MsgBox "Please correct description. This should be at least 9 characters"
   TxtDescription.SetFocus
   GoTo Done
End If
If Len(SSDBCurrencyType) < 1 Then
   MsgBox "Please select a currency type before saving"
   SSDBCurrencyType.SetFocus
   GoTo Done
End If

  '--
i = RunSP(SpCon, "usp_CurrencyUpdate", 0, UCase(TxtCurrencyCode), TxtDescription, SSDBCurrencyType.Columns(1).Text, gblLoginName)
If i = 0 Then
   MsgBox "Currency was successfully created"
   GoTo Done
End If
If i = 1 Then
   MsgBox "Currency was successfully updated:"
   GoTo Done
End If
If i = 99 Then
    MsgBox "Currency was not created. Please contact your system admistrator"
    GoTo Done
End If
If i = 9 Then
   MsgBox "Currency was not amended. Please contact your system admistrator"
   GoTo Done
End If
If i = 22 Then
   MsgBox "Another local currency already exists. Please make the necessary changes and resubmit"
   GoTo Done
End If
  
If gblOptions = 1 Then
   ClearScreen
Else
   Shutdown
   Unload Me
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
End Sub

Private Sub Form_Load()
Dim indx As Integer
Dim strTmp As String
On Error GoTo FL_ERR
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
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

If gblOptions = 1 Then
  Me.Caption = "Add a currency"
Else
  Me.Caption = "Edit a currency"
  Call SSDBCurrencyType_InitColumnProps
  UpdateScreen
End If
'--
   
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox Err & " " & Err.Description, , "Error on loading FrmCurrencyMaint"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
   
End Sub
Private Sub UpdateScreen()
Dim CurrencyType As String
Dim i As Integer
Dim bm As Variant

Set rsMain = RunSP(SpCon, "usp_CurrencyDetails", 1, gblFileKey)
 With rsMain
    TxtCurrencyCode.Locked = True
    TxtCurrencyCode = !CurrencyCode
    TxtDescription = !CurrencyDesc
    If !CurrencyType = "L" Then
       CurrencyType = "Local"
    Else
       CurrencyType = "Foreign"
    End If
    
End With

With SSDBCurrencyType
     .MoveFirst
     For i = 0 To .Rows - 1
         bm = .GetBookmark(i)
         If .Columns(0).CellText(bm) = CurrencyType Then
            .Bookmark = .GetBookmark(i)
             SSDBCurrencyType = .Columns(0).CellText(bm)
         Exit For
         End If
     Next i
 End With
End Sub

Private Sub ClearScreen()
  TxtCurrencyCode = ""
  TxtDescription = ""
End Sub

Private Sub Shutdown()
SpCon.Close
End Sub

Private Sub SSDBCurrencyType_InitColumnProps()
On Error GoTo Err_SSDBCurrencyType_InitColumnProps
Dim Strsql As String
Dim i As Integer

With SSDBCurrencyType
     .RemoveAll
     Strsql = ""
     Strsql = "Local" & vbTab & "L" & vbTab
     .AddItem Strsql
     Strsql = "Foreign" & vbTab & "F" & vbTab
     .AddItem Strsql
End With

Exit_SSDBCurrencyType_InitColumnProps:
Exit Sub

Err_SSDBCurrencyType_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error while populating Currency Type"
Resume Exit_SSDBCurrencyType_InitColumnProps

End Sub
