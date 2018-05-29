VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form frmSIS059 
   Caption         =   "Trustee Report Viewer"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
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
Attribute VB_Name = "frmSIS059"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim sql As String
sql = ""
sql = "SELECT CompName, ClientId, CliName, CliAddr1," _
      & " DteOpened, Joint, Shares, a.CatCode, catDesc" _
      & " from (Company inner join STKNAME a" _
      & " on company.nextacct <> a.clientid)" _
      & " inner join STKCAT b" _
      & " on b.catcode = a.catcode" _
      & " where shares > 0 " _
      & " and a.catcode between 'CA' and 'CZ'" _
      & " order by a.catcode, cliName "
Set adoRs = New ADODB.Recordset
Set cr = New crSIS059
adoRs.Open sql, gblFileName, adOpenDynamic, adLockReadOnly
'crv.ReportSource = Report
cr.Database.Tables.Item(1).SetPrivateData 3, adoRs
'frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
crv.top = 0
crv.left = 0
crv.Height = ScaleHeight
crv.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set cr = Nothing
adoRs.Close
Set adoRs = Nothing
gblReply = True
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
    Me.crv.Width = frmSIS059.ScaleWidth

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


