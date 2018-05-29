VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form frmSIS054 
   Caption         =   "Report Viewer"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   5430
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControl=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertControl=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "Preview (&Application Window)"
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Printer &Setup"
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Print..."
         Index           =   3
         Begin VB.Menu mnuFileItemA 
            Caption         =   "Expor&t"
            Index           =   0
         End
         Begin VB.Menu mnuFileItemA 
            Caption         =   "Pri&nter"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Longon/Logoff Server"
         Index           =   4
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   6
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "A&bout Sis"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmSIS054"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Report As New crSIS054
' This program controls the printing of Cert registers and Issued Certs
Dim rsCmp As New ADODB.Recordset
Dim repSISRept As New SISRepts
Dim sOpt As String
Dim iRepNo As Integer ' used to indicate when the issued report is to be printed

Private Sub Form_Load()
Dim sql As String
sql = ""
sql = "Select CompName, PeriodDte, TrnBatch, "
sql = sql & "Form, LineNo, STKACTIV.ClientId, TrnCode, "
sql = sql & "Status, FrCert, CanDate, FrShares, IssDate, "
sql = sql & "STKACTIV.Shares, CertNo, CliName, CliAddr1, "
sql = sql & "CliAddr2, CliAddr3, CliAddr4 "
sql = sql & " from (Company inner join STKACTIV"
sql = sql & " on company.nextacct <> Stkactiv.ClientId)"
sql = sql & " inner join STKName on STKACTIV.CLIENTID = STKNAME.CLIENTID "
sql = sql & "WHERE STKACTIV.status = 'O' "
'--
iRepNo = 0
Set adoRs = New ADODB.Recordset
Set cr = New crSIS054
If Isloaded("frmSIS053") Then
   frmSIS054.Caption = "Certificate Register Report"
   sOpt = "R"
Else
   If Isloaded("frmSIS022") Then
      frmSIS054.Caption = "Pending Stock Issues List"
      sql = sql & "AND TrnCode = 'I' "
      sOpt = "I"
   Else
      If Isloaded("frmSIS024") Then
        frmSIS054.Caption = "Pending Stockholder to Stockholder List"
        sql = sql & "AND TrnCode = 'S' "
        sOpt = "S"
      Else
        If Isloaded("frmSIS031") Then
          frmSIS054.Caption = "Pending Shareholder to Broker List"
          sql = sql & "AND TrnCode = 'C' "
          sOpt = "C"
        Else
          If Isloaded("frmSIS033") Then
            frmSIS054.Caption = "Pending Broker to Shareholder Listing"
            sql = sql & "AND TrnCode = 'T' "
            sOpt = "T"
          Else
            If Isloaded("frmSIS027") Then
             frmSIS054.Caption = "Pending Broker to Broker"
             sql = sql & "AND TrnCode = 'B' "
             sOpt = "B"
            End If
          End If
        End If
      End If
   End If
End If

sql = sql & " order by TRNBATCH, FORM, STKACTIV.CLIENTID, STKACTIV.LINENO"
adoRs.Open sql, gblFileName, adOpenDynamic, adLockReadOnly
'Report.Database.SetDataSource adors
'crv.ReportSource = Report
cr.Database.Tables.Item(1).SetPrivateData 3, adoRs
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
crv.top = 0
crv.left = 0
crv.Height = ScaleHeight
crv.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_unload_Err
Set rsCmp = New ADODB.Recordset
Select Case sOpt
Case "R"
  rsCmp.Open "Company", gblFileName, adOpenDynamic, adLockOptimistic, adCmdTable
  With rsCmp
    If Not .EOF Then
       !REGISTERIND = True
       .Update
    End If
    
  End With
Set cr = Nothing
End Select
Form_Unload_Exit:
Set cr = Nothing
adoRs.Close
Set adoRs = Nothing
 
If sOpt = "R" Then
   'gblOptions = 54
   'frmReportEngine.Show 0
   'Setup Call to SISRepts
  '----------------------
 Set repSISRept = New SISRepts
 repSISRept.ReportType = 9
 repSISRept.LoginId = gblFileName
 repSISRept.siteid = gblSiteId
 repSISRept.ReportNumber = 54
 repSISRept.RunShareHolderReport
End If
rsCmp.Close
Set rsCmp = Nothing

Set frmSIS054 = Nothing
 Exit Sub
Form_unload_Err:
  csvShowError "SIS054"
  GoTo Form_Unload_Exit
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
Select Case Index
Case 0
   'Pass the report to the viewer to display it
    Me.crv.ReportSource = cr
    
    'Get the PrintingStatus object
    Set CrystalPrintingStatus = cr.PrintingStatus
    
    
    'Want to load the form, but not show it, bcs want to pass the
    'pre-preview values first
    Load frmPrintingStatus
    
    With CrystalPrintingStatus
    
        'Display the info before the report is previewed
        frmPrintingStatus.txtPSPrintedBefore = .NumberOfRecordPrinted
        frmPrintingStatus.txtPSReadBefore = .NumberOfRecordRead
        frmPrintingStatus.txtPSSelectedBefore = .NumberOfRecordSelected
        frmPrintingStatus.txtPSProgressBefore = .Progress
        
        'Preview the report
        Me.crv.ViewReport
            
        'Display the info after the report is previewed
        frmPrintingStatus.txtPSPrintedAfter = .NumberOfRecordPrinted
        frmPrintingStatus.txtPSReadAfter = .NumberOfRecordRead
        frmPrintingStatus.txtPSSelectedAfter = .NumberOfRecordSelected
        frmPrintingStatus.txtPSProgressAfter = .Progress
    
    End With
    
    
    Me.crv.Visible = True
    Me.crv.Width = frmSIS054.ScaleWidth

    Me.crv.Height = (Me.ScaleHeight - Me.crv.top)
 
 
    'Bring the Printing Status form to the front
    frmPrintingStatus.Show
    frmPrintingStatus.SetFocus
 Case 1 'Printersetup
   frmPrinterSetup.Show vbModal
 Case 6
  Unload Me
 End Select
End Sub

Private Sub mnuFileItemA_Click(Index As Integer)
Select Case Index
Case 0 ' Export
 frmExport.Show vbModal
 Case 1 'Print it
  frmPrintOut.Show vbModal
End Select
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
SISAbout.Show vbModal
End Sub


