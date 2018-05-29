VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmJCSD 
   Caption         =   "Stock Exchange Imported File"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13515
   Icon            =   "FrmJCSD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   13515
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11880
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "List of Imported Shareholders"
      TabPicture(0)   =   "FrmJCSD.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSDBJCSD"
      Tab(0).Control(1)=   "CmdUpdate"
      Tab(0).Control(2)=   "CmdExit(0)"
      Tab(0).Control(3)=   "CmdSearch(1)"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "List of imported Mandates"
      TabPicture(1)   =   "FrmJCSD.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSDBMandate"
      Tab(1).Control(1)=   "CmdExit(1)"
      Tab(1).Control(2)=   "CmdSearch(0)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "List of Rejected JCSD Mandates"
      TabPicture(2)   =   "FrmJCSD.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "SSDBRejects"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "CmdExport"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CmnDialog"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   3120
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export To Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   6120
         Width           =   2055
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   -72840
         TabIndex        =   7
         Top             =   6000
         Width           =   2055
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   -74640
         TabIndex        =   6
         Top             =   6000
         Width           =   1935
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   -72000
         TabIndex        =   5
         Top             =   6000
         Width           =   1575
      End
      Begin SSDataWidgets_B.SSDBGrid SSDBMandate 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   13215
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   10
         RowHeight       =   423
         Columns.Count   =   10
         Columns(0).Width=   3200
         Columns(0).Caption=   "ClientID"
         Columns(0).Name =   "ClientID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Client Name"
         Columns(1).Name =   "Client Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Caption=   "Mnd Acnt"
         Columns(2).Name =   "Mnd Acnt"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Caption=   "Account Name"
         Columns(3).Name =   "Account Name"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Caption=   "Bank Id"
         Columns(4).Name =   "Bank Id"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Caption=   "BnkName"
         Columns(5).Name =   "BnkName"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Caption=   "Mnd Name"
         Columns(6).Name =   "Mnd Name"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3200
         Columns(7).Caption=   "Mnd Addr1"
         Columns(7).Name =   "Mnd Addr1"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   3200
         Columns(8).Caption=   "Mnd Addr2"
         Columns(8).Name =   "Mnd Addr2"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3200
         Columns(9).Caption=   "Mnd Addr3"
         Columns(9).Name =   "Mnd Addr3"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         _ExtentX        =   23310
         _ExtentY        =   9763
         _StockProps     =   79
         Caption         =   "JCSD Mandates"
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
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   -70680
         TabIndex        =   3
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74640
         TabIndex        =   2
         Top             =   6000
         Width           =   1575
      End
      Begin SSDataWidgets_B.SSDBGrid SSDBJCSD 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   1
         Top             =   600
         Width           =   13215
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   9
         BackColorOdd    =   16761024
         RowHeight       =   423
         Columns.Count   =   9
         Columns(0).Width=   3200
         Columns(0).Caption=   "GR8NIN"
         Columns(0).Name =   "GR8NIN"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "GR8NAM"
         Columns(1).Name =   "GR8NAM"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Caption=   "GR8AD1"
         Columns(2).Name =   "GR8AD1"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Caption=   "GR8AD2"
         Columns(3).Name =   "GR8AD2"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Caption=   "GR8AD3"
         Columns(4).Name =   "GR8AD3"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Caption=   "Shares/GR8CBL"
         Columns(5).Name =   "Shares/GR8CBL"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Caption=   "CAT"
         Columns(6).Name =   "CAT"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3200
         Columns(7).Caption=   "TAX"
         Columns(7).Name =   "TAX"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   3200
         Columns(8).Caption=   "GR8RATE"
         Columns(8).Name =   "GR8RATE"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         _ExtentX        =   23310
         _ExtentY        =   9340
         _StockProps     =   79
         Caption         =   "JCSD Shareholders && Holdings"
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
      Begin SSDataWidgets_B.SSDBGrid SSDBRejects 
         Height          =   5535
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   13215
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   10
         BackColorOdd    =   12632319
         RowHeight       =   423
         Columns.Count   =   10
         Columns(0).Width=   2143
         Columns(0).Caption=   "ClientID"
         Columns(0).Name =   "ClientID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3969
         Columns(1).Caption=   "Client Name"
         Columns(1).Name =   "Client Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1958
         Columns(2).Caption=   "Mnd Acnt"
         Columns(2).Name =   "Mnd Acnt"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3995
         Columns(3).Caption=   "Account Name"
         Columns(3).Name =   "Account Name"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   847
         Columns(4).Caption=   "Bank Id"
         Columns(4).Name =   "Bank Id"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Caption=   "BnkName"
         Columns(5).Name =   "BnkName"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   2540
         Columns(6).Caption=   "Mnd Name"
         Columns(6).Name =   "Mnd Name"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3200
         Columns(7).Caption=   "Mnd Addr1"
         Columns(7).Name =   "Mnd Addr1"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   3200
         Columns(8).Caption=   "Mnd Addr2"
         Columns(8).Name =   "Mnd Addr2"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3200
         Columns(9).Caption=   "Mnd Addr3"
         Columns(9).Name =   "Mnd Addr3"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         _ExtentX        =   23310
         _ExtentY        =   9763
         _StockProps     =   79
         Caption         =   "These JCSD Mandates will generate a cheque"
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
   Begin SSDataWidgets_B.SSDBCombo SSDBExchange 
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   6840
      Width           =   3615
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
      Columns(0).Width=   3200
      Columns(0).Caption=   "Stock Exchange"
      Columns(0).Name =   "Stock Exchange"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   635
      Columns(1).Caption=   "ID"
      Columns(1).Name =   "ID"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   6376
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
      Caption         =   "Select Stock Exchange"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   6840
      Width           =   3735
   End
End
Attribute VB_Name = "FrmJCSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim SearchCriteria As String

Private Sub CmdExit_Click(Index As Integer)
Unload Me
End Sub

Private Sub CmdExport_Click()
CmnDialog.DialogTitle = "Export Rejected Mandates Listing"
CmnDialog.Filter = "XLS(*.xls)|*.xls"
CmnDialog.DefaultExt = "xls"
CmnDialog.FileName = SSDBExchange.Columns(0).Text & " Mandates Reject List"
CmnDialog.ShowSave
If CmnDialog.FileName = SSDBExchange.Columns(0).Text & " Mandates Reject List" Then
   MsgBox "Save Abondoned"
Else

SSDBRejects.Export ssExportTypeExcel, ssExportAllRows + _
    ssExportColumnHeaders + ssExportOverwriteExisting, CmnDialog.FileName
End If
End Sub

Private Sub CmdSearch_Click(Index As Integer)
Dim StrSql As String
If Index = 0 Then
   StrSql = SSDBExchange.Columns(0).Text & " Mandate Search"
   SearchCriteria = InputBox("Enter Search criteria, that is, part of the shareholder's name", StrSql)
   SEMandates
Else
    StrSql = SSDBExchange.Columns(0).Text & " Search"
    SearchCriteria = InputBox("Enter Search criteria, that is, part of the shareholder's name", StrSql)
    SESearch
End If
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Exit_CmdUpdate_Click
Dim i As Integer

With SSDBJCSD
     i = RunSP(SpCon, "usp_UpdateSE", 0, CDbl(.Columns(0).Text), .Columns(6).Text, .Columns(7).Text, CInt(SSDBExchange.Columns(1).Text))
End With
If i = 0 Then
   MsgBox "Update completed"
Else
   MsgBox "Error on Update"
End If

Exit_CmdUpdate_Click:
Exit Sub

Err_CmdUpdate_Click:
MsgBox Err & Err.Description, vbOKOnly, "Updating JCSD Shareholder Record"
GoTo Exit_CmdUpdate_Click

End Sub

Private Sub Form_Load()
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

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
Set FrmJCSD = Nothing

End Sub

Private Sub SESearch()
Dim adoRst As ADODB.Recordset
Dim StrSql As String

Set adoRst = RunSP(SpCon, "usp_SEQry", 1, SearchCriteria, CInt(SSDBExchange.Columns(1).Text))
If adoRst.EOF Then
   GoTo Exit_SESearch
End If

adoRst.MoveFirst
With SSDBJCSD
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = ""
     For i = 0 To 8
         StrSql = StrSql & adoRst(i) & vbTab
         Next i
     .AddItem StrSql
     adoRst.MoveNext
Loop
End With
adoRst.Close

Exit_SESearch:
Exit Sub

Err_SESearch:
MsgBox Err & Err.Description, vbOKOnly, "Loading Stock Exchange List"
GoTo Exit_SESearch

End Sub
Private Sub SEMandates()
Dim adoRst As ADODB.Recordset
Dim StrSql As String

'If SSTab1.Tab = 1 Then
   Set adoRst = RunSP(SpCon, "usp_SEMandateQry", 1, SearchCriteria, CInt(SSDBExchange.Columns(1).Text))
   If adoRst.EOF Then
      GoTo Exit_SEMandates
   End If

   adoRst.MoveFirst
   With SSDBMandate
        .RemoveAll
        Do While Not adoRst.EOF
        StrSql = ""
        For i = 0 To 9
            StrSql = StrSql & adoRst(i) & vbTab
        Next i
       .AddItem StrSql
       adoRst.MoveNext
       Loop
   End With
   adoRst.Close
'End If

Exit_SEMandates:
Exit Sub

Err_SEMandates:
MsgBox Err & Err.Description, vbOKOnly, "Loading Stoc Exchange Mandates"
GoTo Exit_SEMandates

End Sub

Private Sub SSDBExchange_Click()
On Error GoTo Err_SSDBExchange_Click
Dim adoRst As ADODB.Recordset
Dim StrSql As String

Me.Caption = SSDBExchange.Columns(0).Text & " Imported File"

If SSTab1.Tab = 2 Then
   Set adoRst = RunSP(SpCon, "usp_SEMandateErrQry", 1, CInt(SSDBExchange.Columns(1).Text))
   If adoRst.EOF Then
      GoTo Exit_SSDBExchange_Click
   End If

   adoRst.MoveFirst
   With SSDBRejects
        .RemoveAll
        Do While Not adoRst.EOF
        StrSql = ""
        For i = 0 To 8
            StrSql = StrSql & adoRst(i) & vbTab
        Next i
        .AddItem StrSql
        adoRst.MoveNext
   Loop
   End With
   adoRst.Close
End If

Exit_SSDBExchange_Click:
Exit Sub

Err_SSDBExchange_Click:
MsgBox Err & Err.Description, vbOKOnly, "Loading Rejects List"
GoTo Exit_SSDBExchange_Click
End Sub

Private Sub SSDBExchange_InitColumnProps()
Dim adoRst As ADODB.Recordset
Dim i As Integer
Dim StrSql As String

Set adoRst = RunSP(SpCon, "usp_ListStockExchanges", 1)
SSDBExchange.RemoveAll

With adoRst
     Do While Not .EOF
        StrSql = !ExchangeABBR & vbTab & !StockExchangeID
        SSDBExchange.AddItem StrSql
        .MoveNext
     Loop
End With
adoRst.Close
Set adoRst = Nothing
End Sub


