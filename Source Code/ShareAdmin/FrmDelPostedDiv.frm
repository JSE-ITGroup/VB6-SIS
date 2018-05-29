VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmDelPostedDiv 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Reverse Posted Dividend"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   Icon            =   "FrmDelPostedDiv.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton CmdReverse 
      Caption         =   "Reverse"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1080
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Select Dividend to Reverse"
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
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "FrmDelPostedDiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
Unload Me
End Sub


Private Sub CmdReverse_Click()
On Error GoTo Exit_CmdReverse_Click
Dim adoRst As ADODB.Recordset
Dim StrSql As String
Dim i As Integer

If SSDBList = vbNullString Then
   MsgBox "Please select a date to reverse first"
   SSDBList.SetFocus
   GoTo Exit_CmdReverse_Click
End If
StrSql = "You are about to Delete the Dividends posted for "
StrSql = StrSql & SSDBList.Columns(0).Text & "." & vbCrLf
StrSql = StrSql & "Are you sure you want to do this?"
i = MsgBox(StrSql, vbYesNo)
If i = vbYes Then
i = RunSP(SpCon, "usp_DeletePostedDividend", 1, Format(SSDBList.Columns(0).Text, "dd-mmm-yyyy"), SSDBList.Columns(1).Text)
If i <> 0 Then
   MsgBox "An error occurred. Dividend was not reversed"
   GoTo Exit_CmdReverse_Click
Else
   MsgBox "Posted Divdend was successfully reversed"
End If
Else
    MsgBox "Reversal abondoned"
End If

Exit_CmdReverse_Click:
Exit Sub
Err_CmdReverse_Click:
MsgBox Err.Description, vbOKOnly, "Reverse Posted Dividend Error"
Resume Exit_CmdReverse_Click

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
MsgBox Err.Description, vbOKOnly, "Reverse Posted Dividend Form Load"
GoTo Exit_Form_Load

End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
Set FrmDelPostedDiv = Nothing
End Sub

Private Sub SSDBList_InitColumnProps()
Dim adoList As ADODB.Recordset
Dim StrSql As String

Set adoList = RunSP(SpCon, "usp_SelectPostedDividend", 1)

Do While Not adoList.EOF
   With SSDBList
        StrSql = Format(adoList!ChqDat, "dd-mmm-yyyy") & vbTab & adoList!PayTyp
        .AddItem StrSql
   End With
   adoList.MoveNext
Loop

adoList.Close
Set adoList = Nothing

End Sub
