VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS034 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Broker to Stockholder Transfer"
   ClientHeight    =   4545
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "SIS034.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7320
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   300
      Left            =   1920
      TabIndex        =   23
      ToolTipText     =   "Pressing this button will activate the search program to locate a shareholder."
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox tbFld 
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   2
      ToolTipText     =   "Assign a document number."
      Top             =   960
      Width           =   1500
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   4080
      TabIndex        =   20
      Top             =   4080
      Width           =   975
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Select from existing batches or enter a new batch number"
      Top             =   480
      Width           =   1695
      DataFieldList   =   "Column 0"
      AllowNull       =   0   'False
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
      Columns(0).Width=   2090
      Columns(0).Caption=   "Batch #"
      Columns(0).Name =   "Batch #"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   8
      Columns(1).Width=   2170
      Columns(1).Caption=   "Date"
      Columns(1).Name =   "Date"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd-mmm-yyyy"
      Columns(1).FieldLen=   11
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3000
      TabIndex        =   14
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "Cancels all processing and exits program."
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   5160
      TabIndex        =   7
      ToolTipText     =   "Saves the screen information to the database."
      Top             =   4080
      Width           =   975
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   1
      ToolTipText     =   "Enter the date of the new batch"
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   3
      ToolTipText     =   "Enter the date of  issue you want to appear on the certificate.."
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      ToolTipText     =   "Select a share holder from the list."
      Top             =   2520
      Width           =   2895
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
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   5636
      Columns(0).Caption=   "Client Name"
      Columns(0).Name =   "Client Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   50
      Columns(1).Width=   3200
      Columns(1).Caption=   "Client Id"
      Columns(1).Name =   "Client Id"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   10
      Columns(2).Width=   2487
      Columns(2).Caption=   "Shares"
      Columns(2).Name =   "Shares"
      Columns(2).Alignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   12
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      Enabled         =   0   'False
      DataFieldToDisplay=   "Column 0"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   24
      ToolTipText     =   "Enter the date of  issue you want to appear on the certificate.."
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   15
      Format          =   "#,###"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "Select a share holder from the list."
      Top             =   1800
      Width           =   2895
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
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   5636
      Columns(0).Caption=   "Client Name"
      Columns(0).Name =   "Client Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   50
      Columns(1).Width=   3200
      Columns(1).Caption=   "Client Id"
      Columns(1).Name =   "Client Id"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   10
      Columns(2).Width=   2487
      Columns(2).Caption=   "Shares"
      Columns(2).Name =   "Shares"
      Columns(2).Alignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   12
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      Enabled         =   0   'False
      DataFieldToDisplay=   "Column 0"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   6
      ToolTipText     =   "Enter the date of  issue you want to appear on the certificate.."
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   15
      Format          =   "#,###"
      PromptChar      =   "_"
   End
   Begin VB.Label lbl 
      Caption         =   "Shares:"
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
      Left            =   4680
      TabIndex        =   27
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "From Broker:"
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
      TabIndex        =   26
      Top             =   1440
      Width           =   1740
   End
   Begin VB.Label lbl 
      Caption         =   "Available Shares"
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
      Left            =   5280
      TabIndex        =   25
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Form No:"
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
      Top             =   960
      Width           =   1380
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Batch Date:"
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
      Left            =   4320
      TabIndex        =   21
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Joint Holder #1 Name:"
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
      Left            =   1800
      TabIndex        =   19
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label lblLabels 
      Caption         =   "Joint Holder #2 Name:"
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
      Left            =   1800
      TabIndex        =   18
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Holder #2:"
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
      Left            =   0
      TabIndex        =   17
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Batch No:"
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
      Index           =   14
      Left            =   360
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Holder #1:"
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
      Left            =   0
      TabIndex        =   15
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   7320
      Y1              =   360
      Y2              =   360
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
      TabIndex        =   12
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   0
      X2              =   7320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "To Share Holder:"
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
      Left            =   0
      TabIndex        =   11
      Top             =   2520
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
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Transfer Date"
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
      Left            =   3840
      TabIndex        =   9
      Top             =   960
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
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmSIS034"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim X, i, iNewBroker As Integer, iEOF As Integer
Dim iNewApp As Integer, iNew As Integer
Dim iAvailShares, iBuyShares As Double
Dim rsClient As ADODB.Recordset
Dim rsMain As ADODB.Recordset, rsBroker As ADODB.Recordset
Dim rsJoint As ADODB.Recordset
Dim rsBat As ADODB.Recordset
Dim rsVerBat As ADODB.Recordset
Dim rsVerFrm As ADODB.Recordset
Dim rsPool As ADODB.Recordset
Dim cmdChange As ADODB.Command
Dim cmdDel As ADODB.Command
Dim SpCon As ADODB.Connection
Dim errLoop As Error
Dim errs1 As Error
Dim strTable As String
Dim strRecNO As String
Dim iStocks As Double, icert As Long
Dim bm As Variant
Dim iBrokerId As Long, iClient As Long, iBrkCert As Long
Dim sBatch As String
Dim sForm As String

Private Sub FillCombo(i As Integer)
Dim sRowinfo As String

'Set rsClient = RunSP(SpCon, "usp_ClientFind", 1, qCli, cChk, X)
With rsClient
     If Not .EOF And Not .BOF Then
      dbc(i).RemoveAll
      Do While Not .EOF
         sRowinfo = !CliName & Chr(9) & !ClientID
         sRowinfo = sRowinfo & Chr(9) & !shares
         dbc(i).AddItem sRowinfo
         If dbc(i).Row = 0 Then dbc(i) = !CliName
         .MoveNext
      Loop
    End If
End With
End Sub
Function IsValid() As Integer
Dim iErr As Integer, dtefld As Date, qSQL
Dim sElable As String
sElable = "Broker to Stockholder Entry"
IsValid = False
iErr = 0
'--
If dbc(0) = "" Then ' batch
    iErr = 132
    MsgBox "No Batch Number found", vbOKOnly, sElable
    tbfld(0).SetFocus
    GoTo Validate_Exit
End If
dbc(0) = UCase(dbc(0))
'--
If meb(0) = "" Then 'batch date
   iErr = 139
   MsgBox "Batch Date Missing", vbOKOnly, sElable
   meb(0).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(0)) Then
      iErr = 14
      MsgBox "Batch Date not a valid date", vbOKOnly, sElable
      meb(0).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If tbfld(1).Text = "" Then 'Invalid form
   iErr = 140
   MsgBox "Form Number missing", vbOKOnly, sElable
   tbfld(1).SetFocus
   GoTo Validate_Exit
End If
tbfld(1).Text = UCase(tbfld(1).Text)
'--
If gblOptions = 1 Then 'check for duplicate form
   Set rsVerFrm = RunSP(SpCon, "usp_Sis034a", 1, dbc(0), tbfld(1))
  If Not rsVerFrm.EOF Then
     iErr = 141
     MsgBox "Duplicate form number found", vbOKOnly, sElable
     tbfld(1).SetFocus
     rsVerFrm.Close
     GoTo Validate_Exit
  End If
  rsVerFrm.Close
End If
'--
If meb(1) = "" Then 'Transfer date
   iErr = 129
   MsgBox "Transfer date is missing", vbOKOnly, sElable
   meb(1).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(1)) Then
      iErr = 14
      MsgBox "Transfer date is not a valid date", vbOKOnly, sElable
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If dbc(2) = "" Then 'Sell Broker
   iErr = 153
   MsgBox "The Broker is missing", vbOKOnly, sElable
   dbc(2).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If dbc(1) = "" Then   ' Buy Client name
   iErr = 130
   MsgBox "Client is missing", vbOKOnly, sElable
   dbc(1).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If meb(3) = "" Then ' Shares bought by shareholder
    iErr = 156
    MsgBox "The number of shares bought is missing", vbOKOnly, sElable
    meb(3).SetFocus
    GoTo Validate_Exit
 End If
 '--
 If Val(meb(3)) = 0 Then
    iErr = 157
    MsgBox "The number of shares bought cannot be zero", vbOKOnly, sElable
   meb(3).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If Val(meb(3)) > Val(meb(2)) Then
    iErr = 162
    MsgBox "Shares bought is greater than shares available", vbOKOnly, sElable
   meb(3).SetFocus
   GoTo Validate_Exit
 End If
 '--
 IsValid = True
Validate_Exit:
   
   Exit Function
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClear_Click()
ClearScreen
If gblOptions = 2 Then UpdateScreen
End Sub

Private Sub cmdDelete_Click()
On Error GoTo cmdDelete_Err
Dim X As Integer
Dim imsg As String

imsg = "Are You Sure?"
X = MsgBox("are you sure", vbYesNo, "Broker to Shareholder")

If X = vbYes Then
   cmdDelete.Enabled = False
   X = RunSP(SpCon, "usp_Sis034Delete", 0, iBuyShares, iBrokerId, iBrkCert, sForm, sBatch, iClient)
   cmdCancel_Click
End If
cmdDelete_Exit:
Exit Sub

cmdDelete_Err:
 MsgBox Err.Description, vbOKOnly, "cmdDelete"
 GoTo cmdDelete_Exit

End Sub

Private Sub cmdFind_Click()
Dim X As Integer
Dim cChk As Integer, qCli As String
Dim sWhere As String, sRowinfo As String

Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
     .cbWhere.Clear
    .cbWhere.AddItem "CliName"
    .cbWhere.AddItem "ClientId"
    .cbWhere.ListIndex = 0
    .cbOptions.ListIndex = 0
    .lbl(3).Caption = "Find Broker/Client"
    .optBtn.Buttons(0).Caption = "&Selling"
    .optBtn.Buttons(1).Caption = "&Buying"
    .Show vbModal
    '----------------------------
    '-------- main line ---------
    '----------------------------
    If .tbFind.Text = vbNullString Then
    Else
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      sWhere = Trim(.tbFind.Text)
      X = .cbWhere.ListIndex
      cChk = .cbOptions.ListIndex
      qCli = sWhere
      '---
      If .optBtn.IndexSelected = 0 Then
         i = 2
      Else
         i = 1
      End If
         Set rsClient = RunSP(SpCon, "usp_ClientFind", 1, qCli, cChk, X, i)
      If Not rsClient.EOF Then
         If i = 1 Then cmdUpdate.Enabled = True
         FillCombo (i)
         dbc(i).Enabled = True
         dbc(i).SetFocus
      End If
      rsClient.Close
    End If
  End With
Unload frmFind
Set frmFind = Nothing
Exit Sub
cmdFind_Click_err:
  MsgBox Err.Description, vbOKOnly, "SIS034/CmdFind"
End Sub
Private Sub cmdUpdate_Click()
Dim qBat As String, iLines As Integer, iCrt As Long
Dim iss As Date, iShares As Double
Dim strncode As String, strnbatch As String, StrnDate As Date
Dim i, iUpd, X As Double, bm As Variant, iabort As Integer

On Error GoTo cmdUpdate_Err
' wait message
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.Refresh
iabort = False: iUpd = False
If Not IsValid Then Exit Sub
'--
iLines = 0
iss = DateValue(meb(1).Text)
strncode = "T"
strnbatch = dbc(0)
StrnDate = DateValue(meb(0).Text)
iShares = Val(meb(3))
If gblOptions = 2 Then
    iStocks = iShares - iBuyShares
Else
    iStocks = iShares
End If
X = RunSP(SpCon, "usp_Sis034Update", 0, iBrokerId, iShares, CStr(strnbatch), Format(StrnDate, "dd-mmm-yyyy"), _
tbfld(1), Format(meb(1).Text, "dd-mmm-yyyy"), strncode, Format(iss, "dd-mmm-yyyy"), iClient)

 If gblOptions = 1 Then
    dbc_InitColumnProps (0)
    dbc(0).Enabled = False
    meb(0).Enabled = False
    InitAddNew
 Else
    cmdCancel_Click
 End If

'---
Done:
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Exit Sub
'--
cmdUpdate_Err:
  MsgBox Err.Description, vbOKOnly, "Sis034/Update"
  cmdCancel_Click
End Sub

Private Sub dbc_InitColumnProps(Index As Integer)
Dim sRowinfo As String
Select Case Index
Case 0 ' Load Open Batches
With rsBat
  If Not .EOF And Not .BOF Then
     dbc(0).RemoveAll
     Do While Not .EOF
       sRowinfo = !BATCHNO & vbTab & !BATDATE
       dbc(0).AddItem sRowinfo
       .MoveNext
     Loop
  End If
End With
Case Else
End Select
End Sub

Private Sub dbc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  Select Case Index
    Case 0
      If meb(0).Enabled = True Then
         meb(0).SetFocus
      Else
         tbfld(1).SetFocus
      End If
    Case 1
      If meb(3).Enabled = True Then
        meb(3).SetFocus
      Else
        CmdFind.SetFocus
      End If
    Case 2
      dbc(1).SetFocus
    Case Else
  End Select
     
 Case vbKeyUp
  KeyCode = 0
  Select Case Index
    Case 1
       dbc(2).SetFocus
    Case 2
       meb(1).SetFocus
    Case Else
  End Select
End Select
End Sub

Private Sub dbc_LostFocus(Index As Integer)
Dim i As Integer
Dim sRowinfo As String
Select Case Index
Case 0
  '-----------------------------------------
  '-- get batch date if existing batch keyed
  '-- if not set focus to get date
  '-----------------------------------------
  If dbc(0) = "" Then
     dbc(0).SetFocus
  Else
   iNew = True
   For i = 0 To dbc(0).Rows - 1
     bm = dbc(0).GetBookmark(i)
     If dbc(0).Columns(0).CellText(bm) = dbc(0) Then
       meb(0).Mask = ""
       meb(0).Text = dbc(0).Columns(1).CellText(bm)
       If Not IsDate(meb(1)) Then meb(1) = meb(0)
       iNew = False
       Exit For
     End If
   Next
   If iNew Then
    meb(0).Enabled = True
    meb(0).SetFocus
   End If
  End If
Case 1
  '---------------------------------------
  '-- get corresponding joint record for
  '-- displaying
  '----------------------------------------
  For i = 0 To dbc(1).Rows - 1
     bm = dbc(1).GetBookmark(i)
     If dbc(1).Columns(0).CellText(bm) = dbc(1) Then
       iClient = dbc(1).Columns(1).CellText(bm)
       Exit For
     End If
   Next
  Set rsJoint = RunSP(SpCon, "usp_FindJoint", 1, iClient)
  If Not rsJoint.EOF Then
    lblLabels(12).Caption = rsJoint!JNTNAME1
    If Not IsNothing(rsJoint!JNTNAME2) Then
       lblLabels(13).Caption = rsJoint!JNTNAME2
    End If
  Else
    lblLabels(12).Caption = " "
    lblLabels(13).Caption = " "
  End If
  rsJoint.Close
  meb(3).Enabled = True
Case 2
   For i = 0 To dbc(2).Rows - 1
     bm = dbc(2).GetBookmark(i)
     If dbc(2).Columns(0).CellText(bm) = dbc(2) Then
       iBrokerId = Val(dbc(2).Columns(1).CellText(bm))
       Exit For
     End If
   Next
  Set rsPool = RunSP(SpCon, "usp_FindBrokerPl", 1, 3, iBrokerId)
  If Not rsPool.EOF Then
     meb(2) = dbc(2).Columns(2).CellText(bm) - rsPool!SHRHELD
  Else
     meb(2) = dbc(2).Columns(2).CellText(bm)
  End If
  rsPool.Close
Case Else
End Select
End Sub

Private Sub Form_Activate()
 On Error GoTo Form_Activate_Err
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
'--
If gblOptions = 2 Then
   UpdateScreen
   Me.Caption = "Edit Broker to Stockholder"
   meb(0).Enabled = False   ' Batch Date
   dbc(0).Enabled = False   ' Batch No
   cmdClear.Enabled = False
   CmdFind.Enabled = False
   cmdUpdate.Enabled = False
End If
'--
Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
 If Err = -2147168242 Then ' no current transactions
   Resume 0
 Else
   MsgBox "SIS034/Activate"
   Exit Sub
 End If
End Sub
Private Sub Form_Load()
Dim iDay, ipos, i As Integer
Dim qMain As String
Dim qSql1, strTmp As String
'On Error GoTo FL_ERR
iEOF = 0
'--
   csvCenterForm Me, gblMDIFORM
   '''Set spcon = New ADODB.Connection
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

   Set rsClient = New ADODB.Recordset
   Set rsVerBat = New ADODB.Recordset
   Set rsVerFrm = New ADODB.Recordset
   Set rsBroker = New ADODB.Recordset
   Set rsPool = New ADODB.Recordset

   '-----------------------
   '-- open tables --------
   '-----------------------
   If gblOptions = 2 Then
       ipos = InStr(1, gblFileKey, ";", 1)
       sBatch = Mid(gblFileKey, 1, ipos - 1)
       sForm = Mid(gblFileKey, ipos + 1, (Len(gblFileKey) - ipos))
   Else
       sBatch = ""
       sForm = ""
   End If
   '--
   Set rsMain = RunSP(SpCon, "usp_Sis034C", 1, CStr(sBatch), sForm)
   Set rsBat = RunSP(SpCon, "usp_SIS034Batch", 1)
   If rsMain.State = 0 Then
      iEOF = 1
   End If
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '--
    If gblOptions = 1 Then
      InitAddNew
    End If
   '--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS034/Load"
  Unload Me
   
End Sub
Private Sub UpdateScreen()
Dim i As Integer
Dim qSQL As String
With rsMain
  Do While Not .EOF
    If !stklineno = 1 Then
      dbc(2).RemoveAll
      '--
      dbc(0) = !TrnBatch
      meb(0).Text = !TRNDATE
      tbfld(1).Text = !Form
      meb(1).Text = !CANDATE
      iBrkCert = !FRCERT
      '--
      Set rsClient = RunSP(SpCon, "usp_ClientFind", 1, !ClientID, 0, 1, 2)
      i = 2
      FillCombo (2)
      dbc_LostFocus (2)
      rsClient.Close
    End If
    '--
    If !certno > 0 Then
       Set rsClient = RunSP(SpCon, "usp_ClientFind", 1, !ClientID, 0, 1, 1)
       i = 1
       FillCombo (1)
       dbc_LostFocus (1)
       meb(3) = !shares
       iBuyShares = !shares
       iClient = !ClientID
       rsClient.Close
    End If
    .MoveNext
  Loop
  '--
  dbc(0).Enabled = False
  dbc(1).Enabled = False
  dbc(2).Enabled = False
  meb(0).Enabled = False
  meb(1).Enabled = False
  meb(2).Enabled = False
  meb(3).Enabled = True
  tbfld(1).Enabled = False
  cmdUpdate.Enabled = True
  CmdFind.Enabled = False
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim iSeqKey As Integer
iSeqKey = 0
If gblOptions = 1 And icert <> 0 Then
   iSeqKey = RunSP(SpCon, "usp_Sis034B", 0, icert)
End If
If iSeqKey <> 0 Then
   MsgBox "Unsed numbers not updated"
End If

If iEOF = 0 Then
   rsMain.Close
End If

Set rsBroker = Nothing
Set rsMain = Nothing
Set rsClient = Nothing
Set rsJoint = Nothing
Set rsVerBat = Nothing
Set rsVerFrm = Nothing
Set frmSIS034 = Nothing

SpCon.Close
End Sub

Private Sub meb_GotFocus(Index As Integer)

Select Case Index
Case 0
  meb(Index).Mask = "##-???-####"
Case 1
  If meb(1) = "" Then meb(Index).Mask = "##-???-####"
Case Else
End Select
End Sub

Private Sub meb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  Select Case Index
  Case 0
   tbfld(1).SetFocus
  Case 1
    If dbc(1).Enabled = True Then
       dbc(1).SetFocus
    Else
       CmdFind.SetFocus
    End If
  Case 3
    cmdUpdate.SetFocus
  Case Else
  End Select
Case vbKeyUp
KeyCode = 0
  Select Case Index
  Case 0
    dbc(0).SetFocus
  Case 1
    tbfld(1).SetFocus
  Case 3
    If dbc(2).Enabled Then
          dbc(2).SetFocus
    End If
  Case Else
  End Select
Case Else
End Select
End Sub
Private Sub ClearScreen()
Dim X As Integer
If gblOptions = 1 Then
  For X = 1 To 2
    meb(X).Mask = ""
    meb(X).Text = ""
    dbc(X) = ""
    dbc(X).RemoveAll
    dbc(X).Enabled = False
  Next
    meb(3).Text = ""
    tbfld(1) = ""
    lblLabels(12).Caption = ""
    lblLabels(13).Caption = ""
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
End If
End Sub
Private Sub InitAddNew()
ClearScreen
  Me.Caption = "New Broker to Stockholder"
  icert = 0
  meb(0).Enabled = False
End Sub

Private Sub meb_LostFocus(Index As Integer)
Select Case Index
  Case 0
      If IsDate(meb(0)) Then
        meb(1) = meb(0)
      End If
End Select
End Sub

Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
 KeyCode = 0
 If Index = 1 Then
    meb(1).SetFocus
 End If
Case vbKeyUp
 KeyCode = 0
 If gblOptions = 1 Then
  If Index = 1 Then
    If iNew Then
       meb(0).SetFocus
    Else
       dbc(0).SetFocus
    End If
  End If
 End If
Case Else
End Select
End Sub

