VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmApproveReturns 
   BackColor       =   &H008080FF&
   Caption         =   "Action Pending Returns"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15315
   Icon            =   "FrmApproveReturns.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   15315
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
      Left            =   11160
      TabIndex        =   3
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
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
      Left            =   6120
      TabIndex        =   2
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton CmdApprove 
      Caption         =   "Approve"
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
      Left            =   360
      TabIndex        =   1
      Top             =   6960
      Width           =   2055
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBAction 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   10
      BackColorEven   =   12640511
      BackColorOdd    =   16761087
      RowHeight       =   423
      ExtraHeight     =   185
      Columns.Count   =   10
      Columns(0).Width=   2910
      Columns(0).Caption=   "Bank Account"
      Columns(0).Name =   "Bank Account"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Chq No/Ref No"
      Columns(1).Name =   "Chq No/Ref No"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   7435
      Columns(2).Caption=   "Payee"
      Columns(2).Name =   "Payee"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2249
      Columns(3).Caption=   "Cheque Date"
      Columns(3).Name =   "Cheque date"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   9
      Columns(3).NumberFormat=   "9,999.99"
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "Chq Amount"
      Columns(4).Name =   "Chq Amount"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2487
      Columns(5).Caption=   "User ID"
      Columns(5).Name =   "User ID"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1323
      Columns(6).Caption=   "Select"
      Columns(6).Name =   "Select"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   11
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "PendingID"
      Columns(7).Name =   "PendingID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1296
      Columns(8).Caption=   "Relodge"
      Columns(8).Name =   "Relodge"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "Chq Rec'vd"
      Columns(9).Name =   "Chq Rec'vd"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   27120
      _ExtentY        =   11880
      _StockProps     =   79
      Caption         =   "Select the Items to be approved or revoked"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
Attribute VB_Name = "FrmApproveReturns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdApprove_Click()
On Error GoTo Err_CmdApprove_Click

Dim i As Integer
Dim iNoofRanges As Integer
Dim iPendingID As String
Dim MsgStr As String

iNoofRanges = 0
iPendingID = ""

With SSDBAction
     If .Rows = 0 Then
        MsgBox "There are no pending replacements. This option cannot be selected at this time"
        GoTo Exit_CmdApprove_Click
     End If
     .MoveFirst
     .Redraw = False
     For i = 1 To .Rows
        If .Columns(6).Text = True Then
           iNoofRanges = iNoofRanges + 1
           iPendingID = iPendingID & .Columns(7).Text & ";"
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
           GoTo Exit_CmdApprove_Click
        End If
     End If
End With

If iNoofRanges > 0 Then
   i = RunSP(SpCon, "usp_ApproveReturns", 0, iPendingID, iNoofRanges, gblLoginName)
       If i = 0 Then
          MsgBox "Replacement approved"
          SSDBAction_InitColumnProps
          GoTo Exit_CmdApprove_Click
       End If
Else
   MsgBox "No items were selected. Please correct", vbOKOnly
End If

Exit_CmdApprove_Click:
Exit Sub

Err_CmdApprove_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on actioning pending replacements"
Resume Exit_CmdApprove_Click

End Sub

Private Sub CmdDelete_Click()
On Error GoTo Err_CmdDelete_Click

Dim i As Integer
Dim iNoofRanges As Integer
Dim iPendingID As String

iNoofRanges = 0
iPendingID = ""

With SSDBAction
     If .Rows = 0 Then
        MsgBox "There are no pending returns. This option cannot be selected at this time"
        GoTo Exit_CmdDelete_Click
     End If
     .MoveFirst
     .Redraw = False
     For i = 1 To .Rows
        If .Columns(6).Text = True Then
           iNoofRanges = iNoofRanges + 1
           iPendingID = iPendingID & .Columns(7).Text & ";"
        End If
        .MoveNext
     Next i
     .Redraw = True
End With

If iNoofRanges > 0 Then
   i = RunSP(SpCon, "usp_DeleteReturns", 0, iPendingID, iNoofRanges, gblLoginName)
       If i = 0 Then
          MsgBox "Return deleted"
          GoTo Exit_CmdDelete_Click
       End If
Else
   MsgBox "No items were selected. Please correct", vbOKOnly
End If

Exit_CmdDelete_Click:
Exit Sub

Err_CmdDelete_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on deleting pending returns"
Resume Exit_CmdDelete_Click

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
Private Sub CmdExit_Click()
SpCon.Close
Unload Me
End Sub

Private Sub SSDBAction_Click()
Dim iChq As String
Dim i As Integer
Dim pos As Integer
Dim iStatus As Integer
Dim BankChq As Boolean

With SSDBAction
     iChq = .Columns(1).Text
     If .Columns(6).Value = 0 Then
        .Columns(6).Value = -1
     Else
         .Columns(6).Value = 0
     End If
     iStatus = .Columns(6).Value
     .Redraw = False
     If .Columns(3).Text = "" Then 'this is a bank cheque
        iChq = iChq & "-"
        BankChq = True
        GoTo CheckGrid
     Else
         pos = InStr(1, iChq, "-")
         If pos = 0 Then
            iChq = iChq & "-"
            BankChq = True
            GoTo CheckGrid
         Else
            iChq = Mid(iChq, 1, pos - 1)
            BankChq = False
         End If
     End If
CheckGrid:
     .MoveFirst
     For i = 1 To .Rows
        pos = InStr(1, .Columns(1).Text, iChq)
        If BankChq Then
           If pos <> 0 Then
              .Columns(6).Value = iStatus
           End If
        Else
            If .Columns(3).Text = "" And .Columns(6).Value = -1 Then
               .Columns(6).Value = 0
            End If
        End If
        .MoveNext
     Next i
     .Redraw = True
End With

End Sub

Private Sub SSDBAction_InitColumnProps()
Dim adoRst As ADODB.Recordset
Dim sRowinfo As String
Dim qSQL As String

Set adoRst = RunSP(SpCon, "usp_ListPendingReturns", 1, gblLoginName)

With adoRst
      SSDBAction.RemoveAll
      If Not .EOF Then
        Do While Not .EOF
          sRowinfo = !AccountNo & vbTab & !ChqNum & vbTab & !Payee & vbTab & Format(!ChqDate, "dd-mmm-yyyy") & vbTab
          sRowinfo = sRowinfo & Format(!ChqAmt, "#,###.00") & vbTab & !UserID & vbTab & 0 & vbTab & !PendingID & vbTab
          If !ReLodge = 0 Then
              sRowinfo = sRowinfo & "No" & vbTab & !BnkChqNo
          Else
              sRowinfo = sRowinfo & "Yes" & vbTab & !BnkChqNo
          End If
          SSDBAction.AddItem sRowinfo
         .MoveNext
        Loop
      End If
End With
End Sub
