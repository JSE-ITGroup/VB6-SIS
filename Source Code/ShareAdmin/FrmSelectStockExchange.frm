VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form FrmSelectStockExchange 
   Caption         =   "Select Stock Exchange to upload"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   Icon            =   "FrmSelectStockExchange.frx":0000
   ScaleHeight     =   2400
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   5040
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      ToolTipText     =   "Returns to main menu"
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   120
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      ToolTipText     =   "Start the import"
      Top             =   1920
      Width           =   975
   End
   Begin SSDataWidgets_A.SSDBOptSet Opt 
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      _Version        =   196611
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Main Ledger"
      BackColor       =   -2147483643
      IndexSelected   =   0
      Buttons.Button(0).OptionValue=   "0"
      Buttons.Button(0).Caption=   "Main Ledger"
      Buttons.Button(0).Mnemonic=   77
      Buttons.Button(0).Value=   -1  'True
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   74
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   76
      Buttons.Button(0).PictureRight=   75
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   168
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
   End
End
Attribute VB_Name = "FrmSelectStockExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim ReportNumber As Integer
Dim SubOption As Integer

Private Sub cmdBtn_Click(Index As Integer)
Select Case Index
Case 0
     Unload Me
Case 1
     If SubOption = 1 Then
        ProcessImports
     Else
        If SubOption = 2 Then
           GenerateReports
        Else
           DeleteCheques
        End If
     End If
End Select
Exit_CmdBtn_Click:
Exit Sub

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

   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
   LoadStockExchanges
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "Selection Stock Exchange Error/Load"
  Unload Me
End Sub
Private Function LoadStockExchanges()
Dim adoRst As ADODB.Recordset
Dim i As Integer

Set adoRst = RunSP(SpCon, "usp_ListStockExchanges", 1)
i = 1
With adoRst
     Do While Not .EOF
        Opt.Buttons.Add (1)
        Opt.Buttons.Item(i).Caption = !ExchangeABBR
        Opt.Buttons.Item(i).OptionValue = !StockExchangeID
        i = i + 1
        .MoveNext
     Loop
End With
adoRst.Close
Set adoRst = Nothing

Select Case gblOptions
       Case 1
            Opt.Buttons(0).Visible = False
            Me.Caption = "Import Stock Exchange Members"
            SubOption = 1
            ReportNumber = 1
       Case 2
            Opt.Buttons(0).Visible = False
            Me.Caption = "Import Stock Exchange Categories and Tax groupings"
            SubOption = 1
            ReportNumber = 2
       Case 3
            Opt.Buttons(0).Visible = True
            Me.Caption = "Top N Largest Shareholder's Listing"
            SubOption = 2
            ReportNumber = 3
       Case 4
            Opt.Buttons(0).Visible = True
            Me.Caption = "Top N Largest Shareholder's and Address Listing"
            SubOption = 2
            ReportNumber = 31
       Case 6
            Opt.Buttons(0).Visible = True
            Me.Caption = "Shareholder's Multi line listing by name"
            SubOption = 2
            ReportNumber = 6
       Case 7
            Opt.Buttons(0).Visible = True
            Me.Caption = "Shareholder's Single line listing by Category"
            SubOption = 2
            ReportNumber = 7
       Case 8
            Opt.Buttons(0).Visible = True
            Me.Caption = "Alpha Name and Address by Country and Category"
            SubOption = 2
            ReportNumber = 8
       Case 9
            Opt.Buttons(0).Visible = True
            Me.Caption = "Delete Cheques assigned during current dividend run"
            SubOption = 3
End Select

End Function
Private Sub ProcessImports()
gblReply = Opt.OptionValue
gblHold = Opt.Caption
frmMDI.CmnDialog.DialogTitle = "Import " & gblHold & " XL File"
frmMDI.CmnDialog.Filter = "XLS(*.xls)|*.xls"
frmMDI.CmnDialog.DefaultExt = "XLS"
frmMDI.CmnDialog.ShowOpen
If Len(frmMDI.CmnDialog.FileName) > 0 Then
   If ReportNumber = 1 Then
      ImpSE.Show 0
      'ImpJCSD.Show 0
   Else
      ImpJCSDCat.Show 0
   End If
End If
End Sub

Private Sub GenerateReports()
Dim repSISRept As New SISRepts
repSISRept.ReportType = 9
repSISRept.OptNo = Opt.OptionValue 'Stock Exchange ID
repSISRept.PassData = Opt.Caption 'Stock Exchange Abbr
repSISRept.LoginId = gblFileName   'Database name

repSISRept.ReportNumber = ReportNumber
repSISRept.RunShareHolderReport
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SpCon.Close
Set FrmSelectStockExchange = Nothing
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
End Sub
Private Sub DeleteCheques()
Dim i As Integer
Dim StrSql As String

StrSql = "Note that cheque number assigned to other registers" & vbCrLf
StrSql = StrSql & "below the current register will also be deleted." & vbCrLf
StrSql = StrSql & "These will have to be re-assigned before printing"
MsgBox StrSql, vbOKOnly, "Delete cheque number notice"
i = RunSP(SpCon, "usp_ResetChequeNumbers", 0, Opt.OptionValue)
If i = 0 Then
   MsgBox "Numbers deleted. You may now assign the cheque numbers"
Else
   MsgBox "Deletion failed. If this continues contact IT"
End If


End Sub
