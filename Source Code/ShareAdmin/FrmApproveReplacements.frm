VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmApproveReplacements 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Action Pending Replacements"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "FrmApproveReplacements.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   15105
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
      Left            =   12600
      TabIndex        =   3
      Top             =   4680
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
      Left            =   5880
      TabIndex        =   2
      Top             =   4680
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
      Top             =   4680
      Width           =   2055
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBAction 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double-click on an item to view all connect cheques"
      Top             =   0
      Width           =   15015
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   9
      BackColorEven   =   16761024
      BackColorOdd    =   12648384
      RowHeight       =   423
      ExtraHeight     =   185
      Columns.Count   =   9
      Columns(0).Width=   3200
      Columns(0).Caption=   "Bank Account"
      Columns(0).Name =   "Bank Account"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "ClientID"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   7435
      Columns(2).Caption=   "Shareholder"
      Columns(2).Name =   "Chq No"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3836
      Columns(3).Caption=   "Replacement value"
      Columns(3).Name =   "Narration"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   9
      Columns(3).NumberFormat=   "9,999.99"
      Columns(3).FieldLen=   256
      Columns(4).Width=   3281
      Columns(4).Caption=   "Replacement method"
      Columns(4).Name =   "DBCR"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3254
      Columns(5).Caption=   "Number of cheques "
      Columns(5).Name =   "Number of cheques "
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1323
      Columns(6).Caption=   "Select"
      Columns(6).Name =   "Authorise"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   11
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      Columns(7).Width=   370
      Columns(7).Caption=   "PendingID"
      Columns(7).Name =   "ItemID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "UserID"
      Columns(8).Name =   "UserID"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      _ExtentX        =   26485
      _ExtentY        =   8070
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
Attribute VB_Name = "FrmApproveReplacements"
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
End With

If iNoofRanges > 0 Then
   i = RunSP(SpCon, "usp_ApproveReplacements", 0, iPendingID, iNoofRanges, gblLoginName)
       If i = 0 Then
          MsgBox "Replacement approved"
          GoTo Exit_CmdApprove_Click
       Else
          MsgBox "No Cheques are avalable in Working Stock. Please have a supervisor correct"
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

Private Sub cmdDelete_Click()
On Error GoTo Err_CmdDelete_Click

Dim i As Integer
Dim iNoofRanges As Integer
Dim iPendingID As String

iNoofRanges = 0
iTransferID = ""

With SSDBAction
     If .Rows = 0 Then
        MsgBox "There are no pending replacements. This option cannot be selected at this time"
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
   i = RunSP(SpCon, "usp_DeleteReplacements", 0, iPendingID, iNoofRanges, gblLoginName)
       If i = 0 Then
          MsgBox "Replacement deleted"
          GoTo Exit_CmdDelete_Click
       End If
Else
   MsgBox "No items were selected. Please correct", vbOKOnly
End If

Exit_CmdDelete_Click:
Exit Sub

Err_CmdDelete_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on deleting pending replacements"
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

Private Sub SSDBAction_DblClick()
gblFileKey = SSDBAction.Columns(7).Text
FrmReplacementDetails.Show 0
End Sub

Private Sub SSDBAction_InitColumnProps()
Dim adoRst As ADODB.Recordset
Dim sRowinfo As String
Dim qSQL As String

Set adoRst = RunSP(SpCon, "usp_ListPendingReplacements", 1, 1)

With adoRst
      SSDBAction.RemoveAll
      If Not .EOF Then
        Do While Not .EOF
          sRowinfo = !AccountNo & vbTab & !ClientID & vbTab & !CliName & vbTab & Format(!PayTotal, "#,###.00") & vbTab
          sRowinfo = sRowinfo & !DistDesc & vbTab & !PendingCount & vbTab & 0 & vbTab & !PendingID & vbTab & !UserID
          SSDBAction.AddItem sRowinfo
         .MoveNext
        Loop
      End If
End With
End Sub
Private Sub SSDBAction_Click()
With SSDBAction
     If .Columns(6).Value = 0 Then
        .Columns(6).Value = -1
        If .Columns(8).Value = gblLoginName Then
           MsgBox "You cannot approve a replacement that you initiated"
           .Columns(6).Value = 0
        End If
     Else
         .Columns(6).Value = 0
     End If
End With

End Sub

