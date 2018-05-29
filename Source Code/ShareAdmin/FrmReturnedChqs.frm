VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmReturnedChqs 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Process Returned Cheque(s)"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15150
   Icon            =   "FrmReturnedChqs.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   15150
   StartUpPosition =   3  'Windows Default
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
      Left            =   12480
      TabIndex        =   6
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton CmdProcess 
      Caption         =   "Process Selected items"
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
      Left            =   8040
      TabIndex        =   5
      Top             =   7440
      Width           =   2655
   End
   Begin VB.Frame FmeFindCheque 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Find Cheque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   -120
      TabIndex        =   0
      Top             =   7080
      Width           =   7215
      Begin VB.CommandButton CmdFind 
         Caption         =   "Find"
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
         Left            =   5640
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtChqno 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enter Cheque number:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBResults 
      Height          =   6975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15135
      _Version        =   196617
      DataMode        =   2
      BackColorEven   =   12648384
      BackColorOdd    =   12632319
      RowHeight       =   503
      ExtraHeight     =   238
      Columns.Count   =   9
      Columns(0).Width=   2752
      Columns(0).Caption=   "Account No"
      Columns(0).Name =   "Account No"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   8176
      Columns(1).Caption=   "Payee"
      Columns(1).Name =   "Payee"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3387
      Columns(2).Caption=   "Cheque Number"
      Columns(2).Name =   "Cheque Number"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Cheque Date"
      Columns(3).Name =   "Cheque Date"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3598
      Columns(4).Caption=   "Cheque Amount"
      Columns(4).Name =   "Cheque Amount"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1085
      Columns(5).Caption=   "Select"
      Columns(5).Name =   "Select"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Style=   2
      Columns(6).Width=   1323
      Columns(6).Caption=   "Relodge"
      Columns(6).Name =   "Relodge"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "PayTyp"
      Columns(7).Name =   "PayTyp"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "Bank Chq No"
      Columns(8).Name =   "Bank Chq No"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      _ExtentX        =   26696
      _ExtentY        =   12303
      _StockProps     =   79
      Caption         =   "List of Cheques "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Attribute VB_Name = "FrmReturnedChqs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim iChqAmt As Currency

Private Sub CmdExit_Click()
On Error GoTo Err_CmdExit_Click

Unload Me

Exit_cmdExit_Click:
Exit Sub

Err_CmdExit_Click:
MsgBox Err.Description, vbOKOnly, "Returned Cheques Exit"
GoTo Exit_cmdExit_Click
End Sub

Private Sub cmdFind_Click()
On Error GoTo Err_CmdFind_Click
Dim adoRst As ADODB.Recordset

If Len(TxtChqno) < 1 Then
   MsgBox "Please enter a cheque number to find", vbOKOnly
   GoTo Exit_CmdFind_Click
End If

Set adoRst = RunSP(SpCon, "usp_ListChqsToBeReplaced", 1, TxtChqno)
If adoRst.EOF Then
   MsgBox "No cheques matching your criteria was found", vbOKOnly
   GoTo Exit_CmdFind_Click
End If

With SSDBResults
     .RemoveAll
     Do While Not adoRst.EOF
        StrSql = adoRst!AccountNo & vbTab & adoRst!PayeeName & vbTab & adoRst!ChqNum & vbTab
        StrSql = StrSql & Format(adoRst!ChqDat, "dd-mmm-yyyy") & vbTab & Format(adoRst!ChqAmt, "#,###.00") & vbTab & 0 & vbTab & adoRst!PayType & vbTab & adoRst!PayTyp
        .AddItem StrSql
        adoRst.MoveNext
    Loop
End With
Exit_CmdFind_Click:
Exit Sub

Err_CmdFind_Click:
MsgBox Err.Description, vbOKOnly, "Returned Cheques Find Error"
GoTo Exit_CmdFind_Click
End Sub

Private Sub CmdProcess_Click()
On Error GoTo Err_CmdProcess_Click

Dim i As Integer
Dim iNoofRanges As Integer
Dim ListofChqs As String
Dim ListofAccts As String
Dim ListofRelodge As String
Dim ListofBnkChqs As String
Dim MsgStr As String

iNoofRanges = 0
ListofChqs = ""
ListofAccts = ""

With SSDBResults
     If .Rows = 0 Then
        MsgBox "There are no cheques listed. This option cannot be carried out at this time"
        GoTo Exit_CmdProcess_Click
     End If
     .MoveFirst
     .Redraw = False
     For i = 1 To .Rows
        If .Columns(5).Text = True Then
           iNoofRanges = iNoofRanges + 1
           ListofChqs = ListofChqs & .Columns(2).Text & ";"
           ListofAccts = ListofAccts & .Columns(0).Text & ";"
           ListofRelodge = ListofRelodge & .Columns(6).Value & ";"
           If Len(.Columns(8).Value) < 1 Then
              ListofBnkChqs = ListofBnkChqs & "0;"
           Else
              ListofBnkChqs = ListofBnkChqs & .Columns(7).Value & ";"
           End If
        End If
        .MoveNext
     Next i
     .Redraw = True
     If .Rows - iNoofRanges = 1 Then
        MsgStr = "One Item has not been selected" & vbCrLf
        MsgStr = MsgStr & "Remember the original bank cheque must be selected if it was actually returned" & vbCrLf
        MsgStr = MsgStr & "Do you still want to proceed?"
        i = MsgBox(MsgStr, vbYesNo, "Confirm Process")
        If i = vbNo Then
           GoTo Exit_CmdProcess_Click
        End If
     End If
End With


If iNoofRanges > 0 Then
   i = RunSP(SpCon, "usp_ActionReturns", 0, ListofChqs, iNoofRanges, ListofAccts, ListofRelodge, ListofBnkChqs, gblLoginName)
   If i = 0 Then
      MsgBox "Returns saved to Pending list"
      GoTo Exit_CmdProcess_Click
   Else
      If i = 2 Then
         MsgBox "This was previosly posted"
         GoTo Exit_CmdProcess_Click
      End If
   End If
Else
    MsgBox "No items were selected. Please correct", vbOKOnly
End If

Exit_CmdProcess_Click:
Exit Sub

Err_CmdProcess_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on processing selected items"
Resume Exit_CmdProcess_Click

End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load

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
 '--
Exit_Form_Load::
Exit Sub

Err_Form_Load:
MsgBox Err.Description, vbOKOnly, "Returned Cheques Form Load"

End Sub
Private Sub SSDBResults_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
With SSDBResults
     If ColIndex < 5 Then
        If ColIndex = 4 Then
           .Columns(4).Text = Format(iChqAmt, "#,###.00")
        Else
        .Columns(ColIndex).Text = OldValue
        End If
        GoTo Exit_Sub
     End If
End With
Exit_Sub:
Exit Sub
End Sub

Private Sub SSDBResults_Click()
Dim iChq As String
Dim i As Integer
Dim pos As Integer
Dim iStatus As Integer
Dim BankChq As Boolean


With SSDBResults
     iChqAmt = CCur(.Columns(4).Text)
     If .Col = 5 Then
        SelectEntry
     Else
        ReLodge
     End If
End With

End Sub
Private Sub SelectEntry()
With SSDBResults
     iChq = .Columns(2).Text
     .Redraw = False
     If .Columns(5).Value = 0 Then
        .Columns(5).Value = 1
     Else
        .Columns(5).Value = 0
        .Columns(6).Value = 0
     End If
     iStatus = .Columns(5).Value
     If .Columns(7).Value = "B" Then
        iChq = iChq & "-"
        BankChq = True
        GoTo CheckGrid
     Else
        If pos > 0 Then
           iChq = Mid(iChq, 1, pos - 1)
           BankChq = False
        End If
     End If
CheckGrid:
     .MoveFirst
     For i = 1 To .Rows
        pos = InStr(1, .Columns(2).Text, iChq)
        If BankChq Then
           If pos <> 0 Then
              .Columns(5).Value = iStatus
           End If
        Else
            If .Columns(3).Text = "" And .Columns(5).Value = 1 Then
               .Columns(5).Value = 0
            End If
        End If
        .MoveNext
     Next i
     .Redraw = True
End With
End Sub

Private Sub ReLodge()
Dim i As Integer
With SSDBResults
     If .Columns(6).Value = 0 Then
        .Columns(6).Value = 1
        .Columns(5).Value = 1
     Else
        .Columns(6).Value = 0
        .Columns(5).Value = 0
     End If
End With

End Sub
