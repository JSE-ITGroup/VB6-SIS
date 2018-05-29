VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form frmJBP056 
   Caption         =   "Print Certificate Switchboard"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8415
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControl=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertControl=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Enabled         =   0   'False
      Begin VB.Menu mnuPreview 
         Caption         =   "Preview (&Application Window)"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuFilePrintExport 
            Caption         =   "Expo&rt"
         End
         Begin VB.Menu mnuFilePrintPrinter 
            Caption         =   "Pri&nter"
         End
      End
      Begin VB.Menu mnuFileSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "A&bout SIS"
      End
   End
End
Attribute VB_Name = "frmJBP056"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCmp As New ADODB.Recordset
Dim iPrint As Integer
Private Sub EndJob()
Set cr = Nothing
Set rsCmp = Nothing
Set frmJBP056 = Nothing
Unload Me
End Sub

Private Sub Form_Load()
Set adoRs = New ADODB.Recordset
iPrint = False
Set cr = New crJBP056
Dim iErr As Integer, qSQL As String
'frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
Screen.MousePointer = vbHourglass
'frmMDI.txtStatusMsg.Refresh
'select goes here
qSQL = "SELECT d.Cusip, d.ParValue, a.ClientId, " _
       & "a.IssDate, a.Shares, a.CertNo, " _
       & "a.BrokerBuy, b.CliName, b.CliAddr1, " _
       & "b.CliAddr2, b.CliAddr3, b.CliAddr4, b.CliAddr5, " _
       & "(Select c.JNTNAME1 from STKJOINT c where b.CLIENTID = c.CLIENTID and JNTENDDTE is null) as JNT1, " _
       & "(Select c.JNTNAME2 from STKJOINT c where b.CLIENTID = c.CLIENTID and JNTENDDTE is null) as JNT2, " _
       & "(Select c.JNTNAME3 from STKJOINT c where b.CLIENTID = c.CLIENTID and JNTENDDTE is null) as JNT3, " _
       & " compname " _
       & "From " _
       & "Company d INNER JOIN (STKACTIV a INNER JOIN STKNAME b ON " _
       & " a.ClientId = b.ClientId) ON d.NextCert <> a.CertNo " _
       & " Where " _
       & " a.Shares > 0 AND a.CertNo > 0" _
       & " AND a.status = 'O' " _
       & " AND a.BrokerBuy = 0 " _
       & " AND a.IssDate is not null  " _
       & " Order By a.CertNo"
  adoRs.Open qSQL, gblFileName, adOpenDynamic, adLockReadOnly
  cr.Database.Tables.Item(1).SetPrivateData 3, adoRs
  iPrint = True
  'frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
  Screen.MousePointer = vbDefault
  'frmMDI.txtStatusMsg.Refresh
  mnuFile.Enabled = True
End Sub

Private Sub Form_Resize()
Me.crv.Width = Me.ScaleWidth
Me.crv.Height = (Me.ScaleHeight - Me.crv.top)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set cr = Nothing
rsCmp.Open "Company", gblFileName, adOpenDynamic, adLockOptimistic, adCmdTable
With rsCmp
  If iPrint = True Then
    If Not .EOF Then
       !CERTIND = True
       .Update
    End If
  End If
  .Close
End With
Set rsCmp = Nothing
End Sub

Private Sub mnuFilePrintExport_Click()
 frmExport.Show vbModal
End Sub

Private Sub mnuFilePrintPrinter_Click()
  frmPrintOut.Show vbModal
End Sub
Private Sub mnuHelpAbout_Click()
  SISAbout.Show vbModal
End Sub

Private Sub mnuPreview_Click()
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
    Me.crv.Width = frmJBP056.ScaleWidth

    Me.crv.Height = (Me.ScaleHeight - Me.crv.top)
 
 
    'Bring the Printing Status form to the front
    frmPrintingStatus.Show
    frmPrintingStatus.SetFocus
End Sub

Private Sub mnuPrinterSetup_Click()
    frmPrinterSetup.Show vbModal
End Sub

