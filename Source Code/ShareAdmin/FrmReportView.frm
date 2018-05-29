VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmReportView 
   ClientHeight    =   6120
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7815
   Icon            =   "FrmReportView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CRV 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
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
   Begin VB.Menu MnuPrintPreview 
      Caption         =   "Print Preview"
   End
End
Attribute VB_Name = "FrmReportView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim adoRs As ADODB.Recordset

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

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub MnuPrintPreview_Click()
Select Case gblOptions
       Case 0
            Set adoRs = RunSP(SpCon, "usp_StatementPrint", 1, gblFileKey, gblOptions, Format(gblDate, "dd-mmm-yyyy"))
            Me.Caption = "Bank Reconciliation Viewer"
            If adoRs.EOF Then
               GoTo Exit_MnuPrintPreview_Click
            End If
            Set cr = New crReconReport
            cr.Database.SetDataSource adoRs
            cr.PrinterSetup Me.hwnd
            cr.ParameterFields.Item(1).AddCurrentValue gblCompName
            cr.ParameterFields.Item(2).AddCurrentValue Format(gblDate, "dd-mmm-yyyy")
            cr.ParameterFields.Item(3).AddCurrentValue gblFileKey
            cr.ParameterFields.Item(4).AddCurrentValue CStr(gblOptions)
       Case 1
            Me.Caption = "Replaced Cheques Viewer"
            Set adoRs = RunSP(SpCon, "usp_ReplacedQry", 1, gblReply, Format(gblDate, "dd-mmm-yyyy"), Format(gblDate1, "dd-mmm-yyyy"))
            If adoRs.EOF Then
               GoTo Exit_MnuPrintPreview_Click
            End If
            Set cr = New CrSIS084
            cr.PrinterSetup Me.hwnd
            cr.Database.SetDataSource adoRs
       Case 3 'Cheque Inventory report called from FrmChqMovement
            Me.Caption = gblDSN & " Cheque Movement Viewer"
            'Set cr = New CrChqInventory
            If gblReply = 1 Then
               gblDate = "30-Sep-2011"
               gblDate1 = Date
            End If
            Set adoRs = RunSP(SpCon, "usp_ChqMovementQry", 1, gblFileKey, gblHold, Format(gblDate, "dd-mmm-yyyy"), Format(gblDate1, "dd-mmm-yyyy"))
            If adoRs.EOF Then
               GoTo Exit_MnuPrintPreview_Click
            End If
            Set cr = New CrChqMovement
            cr.Database.SetDataSource adoRs
            cr.ParameterFields.Item(1).AddCurrentValue gblDSN
            cr.ParameterFields.Item(2).AddCurrentValue gblCompName
       Case 4 'Dividend Cheque Reconciliation Report
            Me.Caption = "Dividend Cheque Reconciliation Viewer"
            Set adoRs = RunSP(SpCon, "usp_DividendChqRecon", 1, gblLoginName, gblFileKey)
            If adoRs.EOF Then
               GoTo Exit_MnuPrintPreview_Click
            End If
            Set cr = New crDividendChqsRecon
            cr.Database.SetDataSource adoRs
       Case 5
            Me.Caption = "Finacle Exceptions Processed Report"
            Set adoRs = RunSP(SpCon, "usp_FinacleExceptionsQry", 1, Format(gblDate, "dd-mmm-yyyy"), Format(gblDate1, "dd-mmm-yyyy"))
            'If adoRs.EOF Then
            '   GoTo Exit_MnuPrintPreview_Click
            'End If
            Set cr = New CrElectronicExceptions
            cr.Database.SetDataSource adoRs
            cr.ParameterFields.Item(1).AddCurrentValue "Finacle Exceptions Processed Report"
       Case 6
            Me.Caption = "ACH Exceptions Processed Report"
            Set adoRs = RunSP(SpCon, "usp_ACHExceptionsQry", 1, Format(gblDate, "dd-mmm-yyyy"), Format(gblDate1, "dd-mmm-yyyy"))
            'If adoRs.EOF Then
            '   GoTo Exit_MnuPrintPreview_Click
            'End If
            Set cr = New CrElectronicExceptions
            cr.Database.SetDataSource adoRs
            cr.ParameterFields.Item(1).AddCurrentValue "ACH Exceptions Processed Report"
       Case 7 'called from FrmDates - gblOptons 7
            Me.Caption = "Replaced Cheques Report Viewer"
            Set adoRs = RunSP(SpCon, "usp_ReplacementReport", 1, Format(gblDate, "dd-mmm-yyyy"), Format(gblDate1, "dd-mmm-yyyy"))
            'If adoRs.EOF Then
            '   GoTo Exit_MnuPrintPreview_Click
            'End If
            Set cr = New CrReplacementReport
            cr.Database.SetDataSource adoRs
            cr.ParameterFields.Item(1).AddCurrentValue gblCompName
            cr.ParameterFields.Item(2).AddCurrentValue Format(gblDate, "dd-mmm-yyyy")
            cr.ParameterFields.Item(3).AddCurrentValue Format(gblDate1, "dd-mmm-yyyy")
       Case 8 ' Called from FrmDates - gblOptions 8
            Me.Caption = "Returned Cheques Report Viewer"
            Set adoRs = RunSP(SpCon, "usp_ReturnsReport", 1, Format(gblDate, "dd-mmm-yyyy"), Format(gblDate1, "dd-mmm-yyyy"))
            'If adoRs.EOF Then
            '   GoTo Exit_MnuPrintPreview_Click
            'End If
            Set cr = New CrReturnsReport
            cr.Database.SetDataSource adoRs
            cr.ParameterFields.Item(1).AddCurrentValue gblCompName
            cr.ParameterFields.Item(2).AddCurrentValue Format(gblDate, "dd-mmm-yyyy")
            cr.ParameterFields.Item(3).AddCurrentValue Format(gblDate1, "dd-mmm-yyyy")
      Case 9 'Called from FrmSelectDivDate
           Me.Caption = "Dividend Reconciliation Report Viewer"
           Set adoRs = RunSP(SpCon, "usp_DividendReconReport", 1, gblFileKey, Format(gblDate, "dd-mmm-yyyy"), gblHold)
           If adoRs.EOF Then
              GoTo Exit_MnuPrintPreview_Click
           End If
           Set cr = New CrDividendReconReport
           cr.Database.SetDataSource adoRs
           cr.ParameterFields.Item(1).AddCurrentValue gblCompName
           If gblHold = "D" Then
              cr.ParameterFields.Item(2).AddCurrentValue "Dividend"
           Else
              cr.ParameterFields.Item(2).AddCurrentValue "Capital Distribution"
           End If
           
           cr.ParameterFields.Item(3).AddCurrentValue Format(gblDate, "dd-mmm-yyyy")
    Case 10
          Me.Caption = "Divdend History Viewer"
          Set adoRs = RunSP(SpCon, "usp_DivHistory", 1, gblFileKey, gblReply)
          Set cr = New CRDivHist
          cr.Database.SetDataSource adoRs
          cr.ParameterFields.Item(1).AddCurrentValue gblCompName
     Case 11
          Me.Caption = "Unpaid Divdend History Viewer"
          Set adoRs = RunSP(SpCon, "usp_UnPaidDivHistory", 1, gblFileKey, gblReply)
          Set cr = New CRDivHist
          cr.Database.SetDataSource adoRs
          cr.ParameterFields.Item(1).AddCurrentValue gblCompName
    Case 12
         Me.Caption = "Unclaimed Balances Report Viewer"
         Set adoRs = RunSP(SpCon, "usp_UnclaimedBalances", 1, Format(gblDate, "dd-mmm-yyyy"))
         Set cr = New CRUnclaimedBalances
         cr.Database.SetDataSource adoRs
         cr.ParameterFields(1).AddCurrentValue gblCompName
         cr.ParameterFields.Item(2).AddCurrentValue Format(gblDate, "dd-mmm-yyyy")
    Case 13
         Me.Caption = "Unclaimed Balances Full Report Viewer"
         Set adoRs = RunSP(SpCon, "usp_UnclaimedBalancesFull", 1, Format(gblDate, "dd-mmm-yyyy"))
         Set cr = New CRUnclaimedBalances
         cr.Database.SetDataSource adoRs
         cr.ParameterFields(1).AddCurrentValue gblCompName
         cr.ParameterFields.Item(2).AddCurrentValue Format(gblDate, "dd-mmm-yyyy")
End Select
 
CRV.ReportSource = cr
CRV.ViewReport

Exit_MnuPrintPreview_Click:
Exit Sub

End Sub
Private Sub Form_Resize()
CRV.Top = 0
CRV.Left = 0
CRV.Height = ScaleHeight
CRV.Width = ScaleWidth
End Sub

