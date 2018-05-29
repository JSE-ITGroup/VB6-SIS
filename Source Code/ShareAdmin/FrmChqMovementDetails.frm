VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmChqMovementDetails 
   Caption         =   "Cheque Movement Details"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16065
   Icon            =   "FrmChqMovementDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   16065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12840
      TabIndex        =   1
      Top             =   7080
      Width           =   2175
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBMovements 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16095
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   10
      RowHeight       =   423
      Columns.Count   =   10
      Columns(0).Width=   3200
      Columns(0).Caption=   "Account No"
      Columns(0).Name =   "Account No"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Location"
      Columns(1).Name =   "Location"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2064
      Columns(2).Caption=   "Trans Date"
      Columns(2).Name =   "Trans Date"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).NumberFormat=   "dd-mmm-yyyy"
      Columns(2).FieldLen=   256
      Columns(3).Width=   2223
      Columns(3).Caption=   "Stating Chq No"
      Columns(3).Name =   "Stating Chq No"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2328
      Columns(4).Caption=   "Ending Chq No"
      Columns(4).Name =   "Ending Chq No"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "No of Chqs in Range"
      Columns(5).Name =   "No of Chqs in Range"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "Action"
      Columns(6).Name =   "Action"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Caption=   "Status"
      Columns(7).Name =   "Status"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "UserID"
      Columns(8).Name =   "UserID"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "Post Date"
      Columns(9).Name =   "Post Date"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   7
      Columns(9).NumberFormat=   "dd-mmm-yyyy"
      Columns(9).FieldLen=   256
      _ExtentX        =   28390
      _ExtentY        =   12303
      _StockProps     =   79
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
Attribute VB_Name = "FrmChqMovementDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdClose_Click()
SpCon.Close
Unload Me
End Sub

Private Sub Form_Load()
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

End Sub

Private Sub SSDBMovements_InitColumnProps()
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_ListChqMovementDetails", 1, gblFileKey, gblHold, gblreply, gblOptions, Format(gblDate, "dd-mmm-yyyy"), Format(gblDate1, "dd-mmm-yyyy"))
   If adoRst.EOF Then
      MsgBox "Sorry, no records matching your criteria were found"
      GoTo Exit_SSDBMovements_InitColumnProps
   End If
   With SSDBMovements
        .RemoveAll
        Do While Not adoRst.EOF
           StrSql = adoRst(0) & vbTab & adoRst(1) & vbTab & adoRst(2) & vbTab & adoRst(3) & vbTab & adoRst(4) & vbTab
           StrSql = StrSql & adoRst(5) & vbTab & adoRst(6) & vbTab & adoRst(8) & vbTab & adoRst(7) & vbTab & adoRst(9)
           .AddItem StrSql
           adoRst.MoveNext
        Loop
   End With
Exit_SSDBMovements_InitColumnProps:
Exit Sub

End Sub
