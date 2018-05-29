VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmSIS056 
   Caption         =   "Print Certificates "
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8415
   Icon            =   "SIS056.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8415
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      lastProp        =   500
      _cx             =   5080
      _cy             =   5080
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
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
Attribute VB_Name = "frmSIS056"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCmp As New ADODB.Recordset
Dim iPrint As Integer
Dim SPCon As ADODB.Connection

Private Sub EndJob()
Set cr = Nothing
Set rsCmp = Nothing
Set frmSIS056 = Nothing
Unload Me
End Sub

Private Sub Form_Load()
Dim msg As String, i As Integer
msg = "Are you Using Laser Printer?"
Set adoRs = New ADODB.Recordset
Screen.MousePointer = vbArrowHourglass
Set SPCon = New ADODB.Connection
With SPCon
     .ConnectionString = gblFileName
     .CursorLocation = adUseClient
     '.Provider = "SQLOLEDB.1"
End With
SPCon.Open , , , adAsyncConnect
Do While SPCon.State = adStateConnecting
   Screen.MousePointer = vbHourglass
Loop
'''MsgBox "Connected"
Screen.MousePointer = vbDefault

iPrint = False
i = MsgBox(msg, vbQuestion + vbYesNo, "Certificate Print")
Select Case gblSiteId
  Case "JBPA"  'Print Producers Certificates
    Set cr = New crJBP056
    Me.Caption = "Producer's Certificate Viewer"
  Case "NCB"   'Print NCB Certificates
    If i = vbYes Then
       Set cr = New crNCB056
    Else
       Set cr = New crNCB056G
    End If
    Me.Caption = "NCB's Certificate Viewer"
  Case "DYOLL"  'Print DYOLL Certifictes
    If i = vbYes Then
       Set cr = New crDYL056
    Else
       Set cr = New crDYL056G
    End If
    Me.Caption = "Dyoll's Certificate Viewer"
  Case "FLIFE"  'Print First life Certificates
    Set cr = New crFL056
    Me.Caption = "First Life's Certificate Viewer"
  Case "COURTS"  'Print Courts Certificates
    If i = vbYes Then
       Set cr = New crCTS056
    Else
       Set cr = New crCTS056G
    End If
    Me.Caption = "Court's Certificate Viewer"
  Case "GYEAR"  'Print Goodyear Certificates
    Set cr = New crGYR056
    Me.Caption = "GoodYear's Certificate Viewer"
  Case "LOJ"   'Print Life of Jamaica Certificates
    Me.Caption = "LOJ's Certificate Viewer"
    Set cr = New crLOJ056
  Case "D&G"   'Print D& G Certificates
    If i = vbYes Then
       Set cr = New crDG056
    Else
       Set cr = New crDG056G
    End If
    Me.Caption = "D&G's Certificate Viewer"
  Case "PEG"   'Print Pegasus Certificates
    Set cr = New crPEG056
    Me.Caption = "Pegasus's Certificate Viewer"
  Case "H&L"   'Print Hardware & Lumber Certs
    Set cr = New crFL056
    Me.Caption = "Hardware & Lumber's Certificate Viewer"
  Case "PANJAM"   'Print Pan Carib Certificates
    Set cr = New crFL056
    Me.Caption = "Pan Carib's Certificate Viewer"
  Case "JPS"   'Print JPS Certificates
  Case "CARRERAS"   'Print NCB Certificates
    If i = vbYes Then
       Set cr = New crCAR056
    Else
       Set cr = New crCAR056G
    End If
    Me.Caption = "CARRERAS's Certificate Viewer"
  
  Case Else
    MsgBox ("No Certificate Template exist for SiteId " & gblSiteId)
    Exit Sub
  End Select

Dim iErr As Integer, qSQL As String
'frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
Screen.MousePointer = vbHourglass
'frmMDI.txtStatusMsg.Refresh
'select goes here
Set adoRs = RunSP(SPCon, "usp_PrintCert", 1)
  If adoRs.EOF Then
     MsgBox "No records were found"
  End If
  'cr.Database.Tables.Item(1).SetPrivateData 3, adoRs
  cr.Database.SetDataSource adoRs
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
Dim i As Integer
Set cr = Nothing

If iPrint = True Then
   i = RunSP(SPCon, "usp_IndicatorUpd", 0, 1)
End If
gblReply = True
SPCon.Close
Set frmSIS056 = Nothing
End Sub

Private Sub mnuExit_Click()
Unload Me
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
cr.PrinterSetup Me.hwnd
ContinueProc:
    
    crv.ReportSource = cr
    
    'Get the PrintingStatus object
    '''Set CrystalPrintingStatus = cr.PrintingStatus
    
    
    'Want to load the form, but not show it, bcs want to pass the
    'pre-preview values first
    '''Load frmPrintingStatus
    
    '''With CrystalPrintingStatus
    
        'Display the info before the report is previewed
       ''' frmPrintingStatus.txtPSPrintedBefore = .NumberOfRecordPrinted
       ''' frmPrintingStatus.txtPSReadBefore = .NumberOfRecordRead
       ''' frmPrintingStatus.txtPSSelectedBefore = .NumberOfRecordSelected
       ''' frmPrintingStatus.txtPSProgressBefore = .Progress
        
        'Preview the report
        crv.ViewReport
        
        'Display the info after the report is previewed
        '''frmPrintingStatus.txtPSPrintedAfter = .NumberOfRecordPrinted
        '''frmPrintingStatus.txtPSReadAfter = .NumberOfRecordRead
        '''frmPrintingStatus.txtPSSelectedAfter = .NumberOfRecordSelected
        '''frmPrintingStatus.txtPSProgressAfter = .Progress
    
    '''End With
        
    Me.crv.Visible = True
    Me.crv.Width = frmSIS056.ScaleWidth

    Me.crv.Height = (Me.ScaleHeight - Me.crv.top)
 
 
    'Bring the Printing Status form to the front
    '''frmPrintingStatus.Show
    '''frmPrintingStatus.SetFocus
End Sub

Private Sub mnuPrinterSetup_Click()
   ' frmPrinterSetup.Show vbModal
End Sub

