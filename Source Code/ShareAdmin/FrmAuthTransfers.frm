VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmAuthTransfers 
   BackColor       =   &H00400040&
   Caption         =   "Transfer Authorization"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16650
   Icon            =   "FrmAuthTransfers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   16650
   StartUpPosition =   3  'Windows Default
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
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton CmdRevoke 
      Caption         =   "Revoke"
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
      Left            =   5640
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
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
      Left            =   12360
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBTransfers 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16575
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   11
      RowHeight       =   423
      Columns.Count   =   11
      Columns(0).Width=   2064
      Columns(0).Caption=   "AccountNo"
      Columns(0).Name =   "AccountNo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "From"
      Columns(1).Name =   "From"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2275
      Columns(2).Caption=   "To"
      Columns(2).Name =   "To"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Start No"
      Columns(3).Name =   "Start No"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3043
      Columns(4).Caption=   "End No"
      Columns(4).Name =   "End No"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   9
      Columns(4).FieldLen=   256
      Columns(5).Width=   3043
      Columns(5).Caption=   "No of Cheques"
      Columns(5).Name =   "No of Cheques"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   7
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "Effective Date"
      Columns(6).Name =   "Effective Date"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Caption=   "User ID"
      Columns(7).Name =   "User ID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "Date Done"
      Columns(8).Name =   "Date Done"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1746
      Columns(9).Caption=   "Authorise"
      Columns(9).Name =   "Authorise"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   11
      Columns(9).FieldLen=   256
      Columns(9).Style=   2
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "ItemID"
      Columns(10).Name=   "ItemID"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      _ExtentX        =   29236
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Pending Transfers Awaiting Authorisation"
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
Attribute VB_Name = "FrmAuthTransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdApprove_Click()
On Error GoTo Err_CmdApprove_Click
Dim i As Integer
Dim iNoofRanges As Integer
Dim iTransferID As String

iNoofRanges = 0
iTransferID = ""

With SSDBTransfers
     If .Rows = 0 Then
        MsgBox "There are no pending transfers. This option cannot be selected at this time"
        GoTo Exit_CmdApprove_Click
     End If
     .MoveFirst
     .Redraw = False
     For i = 1 To .Rows
        If .Columns(9).Text = True Then
           iNoofRanges = iNoofRanges + 1
           iTransferID = iTransferID & .Columns(10).Text & ";"
        End If
        .MoveNext
     Next i
     .Redraw = True
End With

If iNoofRanges > 0 Then
   i = RunSP(SpCon, "usp_ActionTransfers", 0, "A", iTransferID, iNoofRanges, gblLoginName)
       If i = 0 Then
          MsgBox "Transfer approved"
          GoTo Exit_CmdApprove_Click
       End If
Else
   MsgBox "No items were selected. Please correct", vbOKOnly
End If

Exit_CmdApprove_Click:
Exit Sub

Err_CmdApprove_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on completing the APPROVE option"
Resume Exit_CmdApprove_Click

End Sub

Private Sub CmdRevoke_Click()
On Error GoTo Err_CmdRevoke_Click
Dim i As Integer
Dim iNoofRanges As Integer
Dim iTransferID As String

iNoofRanges = 0
iTransferID = ""

With SSDBTransfers
     If .Rows = 0 Then
        MsgBox "There are no pending transfers. This option cannot be selected at this time"
        GoTo Exit_CmdRevoke_Click
     End If
     .MoveFirst
     .Redraw = False
     For i = 1 To .Rows
        If .Columns(9).Text = True Then
           iNoofRanges = iNoofRanges + 1
           iTransferID = iTransferID & .Columns(10).Text & ";"
        End If
        .MoveNext
     Next i
     .Redraw = True
End With

If iNoofRanges > 0 Then
   i = RunSP(SpCon, "usp_ActionTransfers", 0, "D", iTransferID, iNoofRanges, gblLoginName)
       If i = 0 Then
          MsgBox "Revocation completed"
          GoTo Exit_CmdRevoke_Click
       End If
Else
   MsgBox "No items were selected. Please correct", vbOKOnly
End If


Exit_CmdRevoke_Click:
Exit Sub

Err_CmdRevoke_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on completing the REVOKE option"
Resume Exit_CmdRevoke_Click

End Sub

Private Sub Form_Load()
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

Private Sub SSDBTransfers_Click()
With SSDBTransfers
     If .Columns(9).Value = 0 Then
        .Columns(9).Value = -1
        If .Columns(7).Value = gblLoginName Then
           MsgBox "You cannot approve a transfer that you initiated"
           .Columns(9).Value = 0
        End If
     Else
         .Columns(9).Value = 0
     End If
End With

End Sub

Private Sub SSDBTransfers_InitColumnProps()
Dim adoRst As ADODB.Recordset
Dim sRowinfo As String
Dim qSQL As String

Set adoRst = RunSP(SpCon, "usp_ListPendingTransfers", 1)

With adoRst
      SSDBTransfers.RemoveAll
      If Not .EOF Then
        Do While Not .EOF
          sRowinfo = !AccountNo & vbTab & !FromLoc & vbTab & !ToLoc & vbTab & !StartNo & vbTab
          sRowinfo = sRowinfo & !EndNo & vbTab & !NoofChqs & vbTab & Format(!TransferDate, "dd-mmm-yyyy") & vbTab & !UserID & vbTab
          sRowinfo = sRowinfo & Format(!PostDate, "dd-mmm-yyyy") & vbTab & 0 & vbTab & !TransferID
          SSDBTransfers.AddItem sRowinfo
         .MoveNext
        Loop
      End If
End With

End Sub
