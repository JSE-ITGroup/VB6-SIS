VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSIS018 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Cheque Switchboard"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9030
   Icon            =   "SIS018.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   9030
   Begin MSComDlg.CommonDialog CmnDialog 
      Left            =   5520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   9015
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   1
      Left            =   7920
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Menu MnuNumbering 
      Caption         =   "Assign Number to Payments"
      Begin VB.Menu MnuNumberingItem 
         Caption         =   "Assign Cheque Numbers"
         Index           =   1
      End
      Begin VB.Menu MnuNumberingItem 
         Caption         =   "Assign Ref/ChqNo to Bank payments"
         Index           =   2
      End
      Begin VB.Menu MnuNumberingItem 
         Caption         =   "ReNumber Cheques"
         Index           =   3
      End
   End
   Begin VB.Menu MnuPrintCheques 
      Caption         =   "Print Cheques"
      Begin VB.Menu MnuPrintChequesItems 
         Caption         =   "Print Cheques"
         Index           =   1
      End
      Begin VB.Menu MnuPrintChequesItems 
         Caption         =   "Print Bank Cheques"
         Index           =   2
      End
   End
   Begin VB.Menu MnuElectronicPay 
      Caption         =   "Electronic Payments"
      Begin VB.Menu MnuElectronicPayItem 
         Caption         =   "Finacle Payments"
         Index           =   1
      End
      Begin VB.Menu MnuElectronicPayItem 
         Caption         =   "ACH Payments"
         Index           =   2
      End
      Begin VB.Menu MnuElectronicPayItem 
         Caption         =   "RTGS Payments"
         Index           =   3
      End
   End
   Begin VB.Menu MnuPaymentReports 
      Caption         =   "Reports"
      Begin VB.Menu MnuPaymentReportsItem 
         Caption         =   "Finacle List"
         Index           =   1
      End
      Begin VB.Menu MnuPaymentReportsItem 
         Caption         =   "Other Bank List (Internal)"
         Index           =   2
      End
      Begin VB.Menu MnuPaymentReportsItem 
         Caption         =   "Other Bank List (External)"
         Index           =   3
      End
      Begin VB.Menu MnuPaymentReportsItem 
         Caption         =   "Bank Letters"
         Index           =   4
      End
      Begin VB.Menu MnuPaymentReportsItem 
         Caption         =   "Foreign Currency List"
         Index           =   5
      End
      Begin VB.Menu MnuPaymentReportsItem 
         Caption         =   "Print Foreign Currency Letters"
         Index           =   6
      End
      Begin VB.Menu MnuPaymentReportsItem 
         Caption         =   "Export Bank Payments to Excel"
         Index           =   7
      End
      Begin VB.Menu MnuPaymentReportsItem 
         Caption         =   "Shareholders Advice"
         Index           =   8
      End
      Begin VB.Menu MnuPaymentReportsItem 
         Caption         =   "Local Cheques Only Report"
         Index           =   9
      End
   End
   Begin VB.Menu MnuConvert 
      Caption         =   "Convert Foreign Payments"
      Begin VB.Menu MnuConvertItems 
         Caption         =   "Run Conversion"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmSIS018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpCon As ADODB.Connection
Dim iLedger As String

Private Sub cmdBtn_Click(Index As Integer)
Dim sql As String
Select Case Index
   Case 1 ' End Job
     Unload Me
   Case Else
End Select
End Sub

Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_SelectLedger", 1)
If adoRst.EOF Then
   MsgBox "No ledger found. Unable to proceed"
   GoTo Exit_Form_Activate
Else
   iLedger = adoRst!StockExchange
End If
adoRst.Close
Set adoRst = Nothing

Exit_Form_Activate:
Exit Sub

Err_Form_Activate:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on retrieving current ledger"
Resume Exit_Form_Activate
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

 csvCenterForm Me, gblMDIFORM
End Sub

Private Sub Form_Resize()
crv.Width = Me.ScaleWidth
crv.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set cr = Nothing
Set frmSIS018 = Nothing
SpCon.Close
frmSIS013.Visible = True
End Sub

Private Sub MnuConvertItems_Click(Index As Integer)
FrmConvert.Show
End Sub

Private Sub MnuElectronicPayItem_Click(Index As Integer)
On Error GoTo Err_MnuElectronicPayItem_Click
Dim dFileName As String
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_CurrentLedger", 1)
If adoRst.EOF Then
   MsgBox "Unable to determine current ledger"
   GoTo Exit_MnuElectronicPayItem_Click
End If
dFileName = adoRst!CompanyABBR & "_" & adoRst!ExchangeABBR & "_"
adoRst.Close
Set adoRst = Nothing

Select Case Index
       Case 1
            With CmnDialog
                 .DialogTitle = "Create Finacle File"
                 .Filter = "TXT(*.txt)|*.txt"
                 .DefaultExt = "txt"
                 dFileName = dFileName & "Finacle"
                 .FileName = dFileName
                 .CancelError = True
                 .ShowSave
                 CreateFinacleFile ("D")
            End With
       Case 2
            With CmnDialog
                 .DialogTitle = "Create ACH File"
                 .Filter = "TXT(*.txt)|*.txt"
                 .DefaultExt = "txt"
                 dFileName = dFileName & "ACH"
                 .FileName = dFileName
                 .CancelError = True
                 .ShowSave
                 CreateACHFile ("D")
            End With
       Case 3
            Set adoRst = RunSP(SpCon, "usp_ExportRTGSPayments", 1)
            Call ExportToExcel(adoRst)
            adoRst.Close
            Set adoRst = Nothing
       Case Else
            GoTo Exit_MnuElectronicPayItem_Click
End Select

Exit_MnuElectronicPayItem_Click:
Exit Sub
Err_MnuElectronicPayItem_Click:
If Err.Number <> cdlCancel Then
   MsgBox Err & " " & Err.Number, vbOKOnly, "Electronic File Creation"
   GoTo Exit_MnuElectronicPayItem_Click
End If

End Sub

Private Sub MnuNumberingItem_Click(Index As Integer)
Select Case Index
       Case 1
            FrmCreateChqPrintFile.Show vbModal
       Case 2
            FrmCreateBnkPrintFile.Show vbModal
       Case 3
            FrmReNumberChqs.Show vbModal
       Case Else
            GoTo Exit_MnuNumberingItem_Click
End Select
Exit_MnuNumberingItem_Click:
Exit Sub

End Sub

Private Sub MnuPaymentReportsItem_Click(Index As Integer)
Dim adoRs As ADODB.Recordset

Select Case Index
      Case 1
            Set adoRs = RunSP(SpCon, "usp_DividendList", 1, 1, "JMD")
            Me.Caption = iLedger & " Finacle Payment Register Viewer"
            Set cr = New crSIS017
      Case 2
            Set adoRs = RunSP(SpCon, "usp_DividendList", 1, 2, "JMD")
            Me.Caption = iLedger & " Other Banks' Payment Register Viewer"
            Set cr = New crSIS017B
      Case 3
            Set adoRs = RunSP(SpCon, "usp_BankPaymentList", 1)
            Set cr = New CRBnkList
      Case 4
            Set adoRs = RunSP(SpCon, "usp_printBnkChq", 1)
            Me.Caption = iLedger & " Bank Cheques Viewer"
            Set cr = New CRBankLetter
      Case 5
            FrmSelectCurrency.Show vbModal
            Set adoRs = RunSP(SpCon, "usp_DividendList", 1, 3, gblFileKey)
            Me.Caption = iLedger & " " & gblFileKey & " Dividend Payment Register Viewer"
            Set cr = New crSIS017
      Case 6
            FrmSelectCurrency.Show vbModal
            Set adoRs = RunSP(SpCon, "usp_ForeignChqLetters", 1, gblFileKey)
            Me.Caption = iLedger & " " & gblFileKey & " Dividend Letters Viewer"
            Set cr = New CRUSLetter
      Case 7
            FrmExportBank.Show vbModal
            GoTo Exit_MnuPaymentReportsItem_Click
      Case 8
            Set adoRs = RunSP(SpCon, "usp_CustomerAdv", 1)
            Set cr = New crNCB018C
            Me.Caption = iLedger & " Shareholder Advice Viewer"
      Case 9
            Set adoRs = RunSP(SpCon, "usp_DividendList", 1, 4, "JMD")
            Me.Caption = iLedger & " Local Cheques Only Payment Viewer"
            Set cr = New crSIS017
      Case Else
           GoTo Exit_MnuPaymentReportsItem_Click
End Select
ViewPrint:
cr.Database.SetDataSource adoRs
cr.PrinterSetup Me.hwnd
crv.ReportSource = cr
crv.ViewReport

Exit_MnuPaymentReportsItem_Click:
Exit Sub

End Sub
Private Sub MnuPrintChequesItems_Click(Index As Integer)
Dim adoRs As ADODB.Recordset
Select Case Index
       Case 1
            FrmPrintCheques.Show vbModal
            If Len(gblFileKey) > 1 Then
               SelectChqFormat
               Set adoRs = RunSP(SpCon, "usp_printChq", 1, gblFileKey)
            Else
               GoTo Exit_MnuPrintChequesItems_Click
            End If
       Case 2
            Set adoRs = RunSP(SpCon, "usp_printBnkChq", 1)
            Set cr = New crNCB018B
            cr.ParameterFields.Item(1).AddCurrentValue gblCompName
       Case Else
            GoTo Exit_MnuPrintChequesItems_Click
End Select

ViewReport:
Screen.MousePointer = vbDefault
If Not adoRs.EOF Then
   cr.PrinterSetup Me.hwnd
   cr.Database.SetDataSource adoRs
   crv.ReportSource = cr
   crv.ViewReport
End If

Exit_MnuPrintChequesItems_Click:
Exit Sub

End Sub
Private Sub SelectChqFormat()
Dim adoRs As ADODB.Recordset
Dim ChqFormat As Integer

Set adoRs = RunSP(SpCon, "usp_GetChequeFormat", 1)
If IsNull(adoRs!ChqFormat) Then
   ChqFormat = 4
Else
   ChqFormat = adoRs!ChqFormat
End If
adoRs.Close
Set adoRs = Nothing
Select Case ChqFormat
Case 1 ' Courts Cheques - crCTS018
 Set cr = New crCTS018
Case 2 ' D&G - crDG018
 Set cr = New crDG018
Case 3 ' Dyoll Cheques - crDYL018
 Set cr = New crDG018
Case 4 ' First Life - crFL018
 Set cr = New crNCB018
Case 5 ' Goodyear
 Set cr = New crDG018
Case 6 ' NCB - crNCB018
 Set cr = New crNCB018
Case 7 ' Pegasus
 Set cr = New crNCB018
Case 8 ' Producer - crJBP018
 Set cr = New crJBP018
Case Else ' Default
 Set cr = New crNCB018
End Select
End Sub

