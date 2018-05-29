VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmReportEngine 
   Caption         =   "SIS Reports Generator"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8475
   Icon            =   "frmReportEngine.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8475
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   8475
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   8535
   End
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   5415
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8535
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPreviewApplication 
         Caption         =   "Preview (&Application Window)"
      End
      Begin VB.Menu mnuPrinterSetup 
         Caption         =   "Printer &Setup..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuFilePrintExport 
            Caption         =   "Expor&t..."
         End
         Begin VB.Menu mnuFilePrintPrinter 
            Caption         =   "Pri&nter"
         End
      End
      Begin VB.Menu mnuFileLogonServer 
         Caption         =   "&Logon\Logoff Server"
      End
      Begin VB.Menu mnuFileSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "A&bout SISl..."
      End
   End
End
Attribute VB_Name = "frmReportEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCmp As New ADODB.Recordset
Dim sQuote As String, iEmpty As Integer
Dim SPCon As ADODB.Connection
Dim qSQL  As String, sql As String, sOpt As String
Private Sub Form_Activate()
Select Case gblOptions
Case 0   'Brokers summary report
  Call SIS040
Case 1 'Brokers Detail Report
  Call SIS041
Case 2  'Percentage Ownership Report
  Call SIS042
Case 3 'Top n shareholder report
  Call SIS043
Case 4 'Bank Mandate List
  Call SIS044
Case 5 'Shareholder Profile
  Call SIS045
Case 6 ' Stockholder Register
   Call SIS046
Case 7 ' Stockholder Register - Category
   Call SIS046C
Case 8 ' Name & Address by country & Category
   Call SIS046D
Case 9 ' Certificate register - All Transactions
   sOpt = "R"
   Call SIS054
Case 10 ' Certificate register - Pending stock Issues
    sOpt = "I"
   Call SIS054
Case 11 ' Certificate register - Pending Stockholder to Stockholder
   sOpt = "S"
   Call SIS054
Case 12 ' Certificate register - Pending Stockholder to Broker
    sOpt = "C"
   Call SIS054
Case 13 ' Certificate register - Pending Broker to Stockholder
    sOpt = "T"
   Call SIS054
Case 14 ' Certificate register - Pending Broker to Broker
    sOpt = "B"
   Call SIS054
Case 15 ' Trustee Report for Register = 'JBPA'
   Call SIS059
Case 16 ' Payment Register Summary
   Call SIS016
Case 17 'Payment Register
   Call SIS017
Case 18
   Call SIS104
Case 31 'Top n shareholder report
  Call SIS043A
Case 47 'Annual Government Return
   Call SIS047
Case 54 'Issued Certificates
   Call SIS054A
Case 55 'Audit Trail Report
   Call SIS055
Case 63 ' Working allotment report Viewer
   Call SIS063
Case 64 ' Bonus allotment report Viewer
   Call SIS064
Case 85 'Outstanding Cheque report
   Call SIS085
Case 161 'Interest Paymnet Summary
   Call SIS0161
Case 174 'Interest Payment Details Report
   Call SIS0174
Case 851 'Reconciled Cheque Report
   Call SIS085A
Case 852 'List all replaced cheques
   Call SIS085B
Case 103 ' RI Allotment Report
   Call SIS103
Case 856 ' Batch Reconciled Cheque Report
   Call SIS086
Case 857 ' RI Allotment Report
   Call SIS087
Case Else
End Select
End Sub

Private Sub Form_Load()
On Error GoTo FormLoad_Err
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
sql = ""
iEmpty = False

FormLoad_Exit:
 Exit Sub
FormLoad_Err:
  
  csvShowError "ReportEngine/Load"
 GoTo FormLoad_Exit
End Sub

Private Sub Form_Resize()
    Me.crv.Width = Me.ScaleWidth
    Me.crv.Height = (Me.ScaleHeight - Me.crv.top - Me.Picture1.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim iRep As Integer
Dim i As Integer
  
'Clean up
Set cr = Nothing
If Not iEmpty Then adoRs.Close
Set adoRs = Nothing
Set frmReportEngine = Nothing
'frmMDI.btnClose.Enabled = True
'--
Select Case gblOptions
Case 9
  i = RunSP(SPCon, "usp_RegisterInd", 0)
  gblOptions = 54
  frmReportEngine.Show 0  ' program reloads itself with option set to 54
  
Case 55
  iRep = MsgBox("Did you print the report & was it okay?", _
        vbQuestion + vbYesNo, "Print Audit Trail")
  If iRep = vbNo Then Exit Sub
  '-- update company!audit report printed
  i = RunSP(SPCon, "usp_AuditInd", 0)
End Select
SPCon.Close
End Sub

Private Sub mnuExit_Click()
    Unload Me
    '''Set cnn = Nothing
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
Private Sub mnuPreviewApplication_Click()
cr.PrinterSetup Me.hwnd
ContinueProc:
    Me.crv.ReportSource = cr
    
    Set CrystalPrintingStatus = cr.PrintingStatus
    
    Me.crv.ViewReport
    
    Me.crv.Visible = True
    Me.crv.Width = frmReportEngine.ScaleWidth

    Me.crv.Height = (Me.ScaleHeight - Me.crv.top - Me.Picture1.Height)
 
End Sub

Private Sub mnuPrinterSetup_Click()
    'frmPrinterSetup.Show vbModal
End Sub
Private Sub SIS041()
Dim iRep As Integer
Set cr = New crSIS041

Set adoRs = RunSP(SPCon, "usp_BrokerDetail", 1)
cr.Database.SetDataSource adoRs

Me.Caption = "Brokers Detail Register Report "
''  iEmpty = True
''  MsgBox "No Register to Print"
End Sub

Private Sub SIS040()
Set adoRs = RunSP(SPCon, "usp_BrokerSummary", 1)
Set cr = New crSIS040
cr.Database.SetDataSource adoRs
Me.Caption = "Brokers Summary Report "
  
End Sub

Private Sub SIS042()
Set adoRs = RunSP(SPCon, "usp_PercentOwn", 1)
Set cr = New crSIS042
cr.Database.SetDataSource adoRs
Me.Caption = "Percentage Ownership Report"

End Sub

Private Sub SIS043()

Dim sMsg, sTitle, sDefault, sValue As String
Dim iTopN As Integer
'--
sMsg = "Enter a number representing the 'N'" _
& " largest shareholders required?"
sTitle = "Top Shareholders Input Box"
sDefault = "10"
sValue = InputBox(sMsg, sTitle, sDefault)
Me.Caption = "Top " & sValue & " Stockholders Report " & gblPassData

If gblOptNo = 0 Then
   Set adoRs = RunSP(SPCon, "usp_TopN", 1, sValue)
Else
   Set adoRs = RunSP(SPCon, "usp_TopNSE", 1, sValue, gblOptNo)
End If

Set cr = New crSIS043
cr.Database.SetDataSource adoRs

Frm_Load_exit:
 Exit Sub
Frm_load_err:
 csvShowError "Report Engine/SIS043"
 Unload Me
  Resume Frm_Load_exit

End Sub
Private Sub SIS043A()

Dim sMsg, sTitle, sDefault, sValue As String
Dim iTopN As Integer
'--
sMsg = "Enter a number representing the 'N'" _
& " largest shareholders required?"
sTitle = "Top Shareholders Input Box"
sDefault = "10"
sValue = InputBox(sMsg, sTitle, sDefault)
Me.Caption = "Top " & sValue & " Stockholders Report"

If gblOptNo = 0 Then
   Set adoRs = RunSP(SPCon, "usp_TopNDetails", 1, sValue)
Else
   Me.Caption = "Top " & sValue & " Stockholders Report " & gblPassData
   Set adoRs = RunSP(SPCon, "usp_TopNSEDetails", 1, sValue, gblOptNo)
End If

Set cr = New crSIS043A
cr.Database.SetDataSource adoRs

Frm_Load_exit:
 Exit Sub
Frm_load_err:
 csvShowError "Report Engine/SIS043"
 Unload Me
  Resume Frm_Load_exit

End Sub

Private Sub SIS044()
Set cr = New crSIS044
Set adoRs = RunSP(SPCon, "usp_MandateList", 1)
cr.Database.SetDataSource adoRs
Me.Caption = "Bank Mandate Listing"
End Sub

Private Sub SIS045()
Set cr = New crSIS045
Set adoRs = RunSP(SPCon, "usp_StockHolderProfile", 1, gblFileKey)
cr.Database.SetDataSource adoRs
Me.Caption = "Stockholder Profile"
End Sub

Private Sub SIS046()
Set cr = New crSIS046
Me.Caption = gblPassData & " - Stockholder Register"

Set adoRs = RunSP(SPCon, "usp_StockholderRegister", 1, gblOptNo)
cr.Database.SetDataSource adoRs
cr.ParameterFields.Item(1).AddCurrentValue gblPassData
Screen.MousePointer = vbDefault

End Sub

Private Sub SIS046C()
Set cr = New crsis046C
Me.Caption = gblPassData & " - Stockholder Register By Category"

Set adoRs = RunSP(SPCon, "usp_RegisterCategory", 1, gblOptNo)
cr.Database.SetDataSource adoRs
cr.ParameterFields.Item(1).AddCurrentValue gblPassData
Screen.MousePointer = vbDefault

End Sub

Private Sub SIS085()
Dim AccountNo As String
Dim StrSql As String
Set cr = New crSIS085

If gblOptNo = 1 Then
   StrSql = "Reconciled Cheque Report"
Else
   StrSql = "Outstanding Cheque Report"
End If
If gblDate1 = "01-Jan-1900" Then
   StrSql = StrSql & " for all payments"
Else
   StrSql = StrSql & " for payments dated " & Format(gblDate1, "dd-mmm-yyyy")
End If
Me.Caption = StrSql
AccountNo = gblSiteId
Set adoRs = RunSP(SPCon, "usp_SIS085", 1, Format(gblDate1, "dd-mmm-yyyy"), AccountNo, gblOptNo)
If adoRs.EOF Then
End If

cr.Database.SetDataSource adoRs
cr.ParameterFields.Item(1).AddCurrentValue StrSql
cr.FormulaFields.Item(2).Text = sQuote & " " & sQuote



Screen.MousePointer = vbDefault
End Sub

Private Sub SIS085A()
Dim sPeriodEnd As String
Set rsCmp = RunSP(SPCon, "usp_Company", 1)
sPeriodEnd = Format(rsCmp!PERIODDTE, "dd-mmm-yyyy")
sQuote = Chr(34)

'--Rene Tomlinson 01/03/2004
Set cr = New crSIS085
Set adoRs = RunSP(SPCon, "usp_SIS085A", 1)
cr.Database.SetDataSource adoRs
cr.FormulaFields.Item(1).Text = sQuote & "Reconciled Cheques for Period Ending " _
                    + sPeriodEnd & sQuote
Me.Caption = "Reconciled Cheque Report Viewer"
rsCmp.Close
End Sub

Private Sub SIS085B()
Dim sYYYYMM As String, sPeriodEnd As String
Set rsCmp = RunSP(SPCon, "usp_Company", 1)
sYYYYMM = Year(rsCmp!PERIODDTE) & Month(rsCmp!PERIODDTE)
sPeriodEnd = Format(rsCmp!PERIODDTE, "dd-mmm-yyyy")
sQuote = Chr(34)
'--
Set cr = New crSIS085
'''cr.Database.Tables.Item(1).SetLogOnInfo gblDSN

'''adoRs.CursorLocation = adUseClient
Set adoRs = RunSP(SPCon, "usp_SIS085B", 1)
'''adoRs.ActiveConnection = Nothing

'''cr.Database.Tables.Item(1).SetPrivateData 3, adoRs
cr.Database.SetDataSource adoRs
cr.FormulaFields.Item(1).Text = sQuote & "Replaced Cheques at " _
                    + sPeriodEnd & sQuote
Me.Caption = "Replaced Cheque Report Viewer"
rsCmp.Close
End Sub

Private Sub SIS086()
Dim sYYYYMM As String, sPeriodEnd As String
'Set rsCmp = New ADODB.Recordset
'rsCmp.Open "COMPANY", gblFileName, , , adCmdTable
'sYYYYMM = Year(rsCmp!PERIODDTE) & Month(rsCmp!PERIODDTE)
'sPeriodEnd = Format(rsCmp!PERIODDTE, "dd-mmm-yyyy")
'sQuote = Chr(34)
'--
    
Set cr = New crSIS086
'''cr.Database.Tables.Item(1).SetLogOnInfo gblDSN

'''adoRs.CursorLocation = adUseClient
Set adoRs = RunSP(SPCon, "usp_SIS086", 1)
'''adoRs.ActiveConnection = Nothing

'''cr.Database.Tables.Item(1).SetPrivateData 3, adoRs
'cr.FormulaFields.Item(1).Text = sQuote & "Replaced Cheques at " _
                    + sPeriodEnd & sQuote
cr.Database.SetDataSource adoRs
Me.Caption = "Batch Reconciled Cheques Report"
'rsCmp.Close

End Sub
Private Sub SIS087()
Dim sYYYYMM As String, sPeriodEnd As String
'Set rsCmp = New ADODB.Recordset
'rsCmp.Open "COMPANY", gblFileName, , , adCmdTable
'sYYYYMM = Year(rsCmp!PERIODDTE) & Month(rsCmp!PERIODDTE)
'sPeriodEnd = Format(rsCmp!PERIODDTE, "dd-mmm-yyyy")
'sQuote = Chr(34)
'--
    
Set cr = New crSIS087
'''cr.Database.Tables.Item(1).SetLogOnInfo gblDSN

'''adoRs.CursorLocation = adUseClient
Set adoRs = RunSP(SPCon, "usp_SIS089", 1)
'''adoRs.ActiveConnection = Nothing

'''cr.Database.Tables.Item(1).SetPrivateData 3, adoRs
'cr.FormulaFields.Item(1).Text = sQuote & "Replaced Cheques at " _
                    + sPeriodEnd & sQuote
cr.Database.SetDataSource adoRs
Me.Caption = "Batch Outstanding Cheques Report"
'rsCmp.Close

End Sub
Private Sub SIS055()
Dim X As Integer
Dim Y As Integer
Dim CrDatabase As craxdrt.Database
Dim CrDatabaseTables As craxdrt.DatabaseTables
Dim CrDatabaseTable As craxdrt.DatabaseTable
Dim CrSections As craxdrt.Sections
Dim CrSection As craxdrt.Section
Dim CrReportObj As craxdrt.ReportObjects
Dim CrSubreportObj As craxdrt.SubreportObject
Dim CrSubreport As craxdrt.Report
Dim adoRs1 As ADODB.Recordset

Set cr = New crSIS055
Set adoRs = RunSP(SPCon, "usp_AuditTrail", 1)
Set adoRs1 = adoRs.NextRecordset

Set CrSections = cr.Sections
cr.Database.SetDataSource adoRs
For X = 1 To CrSections.Count
     Set CrSection = CrSections.Item(X)
     Set CrReportObj = CrSection.ReportObjects
     For Y = 1 To CrReportObj.Count
         If CrReportObj.Item(Y).Kind = crSubreportObject Then
             Set CrSubreportObj = CrReportObj.Item(Y)
             Set CrSubreport = CrSubreportObj.OpenSubreport

             Set CrDatabase = CrSubreport.Database
             Set CrDatabaseTables = CrDatabase.Tables
             Set CrDatabaseTable = CrDatabaseTables.Item(1)
             'MsgBox CrDatabaseTables.Item(1)
             'CrDatabaseTable.SetDataSource adoRs1
             CrDatabaseTable.SetPrivateData 3, adoRs1
         End If
     Next
Next

Me.Caption = "Audit Trail Report Viewer"
End Sub

Private Sub SIS016()
Set cr = New crSIS016
Set adoRs = RunSP(SPCon, "usp_PaymentSummary", 1)
cr.Database.SetDataSource adoRs
Me.Caption = "Payment summary Viewer"
End Sub
Private Sub SIS0161()
Set cr = New crSIS016I
Set adoRs = RunSP(SPCon, "usp_InterestSummary", 1)
cr.Database.SetDataSource adoRs
Me.Caption = "Interest Payment summary Viewer"

End Sub
Private Sub SIS017()
Dim iCurr As String

If gblOptNo = 3 Then
   iCurr = gblSiteId
Else
   iCurr = "OOO"
End If

Set adoRs = RunSP(SPCon, "usp_DividendList", 1, gblOptNo, iCurr)
If adoRs.EOF Then
   MsgBox "No data found"
   Exit Sub
End If

Set cr = New crSIS017
Screen.MousePointer = vbDefault
cr.Database.SetDataSource adoRs
   
cr.Database.SetDataSource adoRs
Me.Caption = adoRs!CompName & " Payment Register Viewer"
End Sub

Private Sub SIS0174()
Set adoRs = RunSP(SPCon, "usp_InterestDetails", 1)
Set cr = New crSIS017I

cr.Database.SetDataSource adoRs
Me.Caption = "Interest Payment Register Viewer"
End Sub

Private Sub SIS063()
Set adoRs = RunSP(SPCon, "usp_BonusAllotment", 1)
Set cr = New crSIS063
cr.Database.SetDataSource adoRs
Me.Caption = "Working Allotment Report Viewer"

End Sub

Private Sub SIS064()
Set adoRs = RunSP(SPCon, "usp_BonusAllotment", 1)
Set cr = New crSIS064
cr.Database.SetDataSource adoRs
cr.ReportTitle = "Bonus Allotment Report"
Me.Caption = "Bonus Allotment Report Viewer"

End Sub

Private Sub SIS054()
' This module controls the printing of Certificate Registers and is driven
' by an Option that indicates which program the routine was called from.
' If sOpt = "R" the module was called from SIS053
' If sOpt = "I" the module was called from SIS022
' If sOpt = "S" the module was called from SIS024
' If sOpt = "C" the module was called from SIS031
' If sOpt = "T" the module was called from SIS033
' If sOpt = "B" the module was called from SIS027
'--
'--
Set cr = New crSIS054
Select Case sOpt
 Case "R"
   Me.Caption = "Certificate Register Report"
 Case "I"
   Me.Caption = "Pending Stock Issues List"
 Case "S"
   Me.Caption = "Pending Stockholder to Stockholder List"
 Case "C"
   Me.Caption = "Pending Shareholder to Broker List"
 Case "T"
   Me.Caption = "Pending Broker to Shareholder Listing"
 Case "B"
   Me.Caption = "Pending Broker to Broker"
 Case Else
 End Select
Set adoRs = RunSP(SPCon, "usp_CertRegister", 1, sOpt)
cr.Database.SetDataSource adoRs
End Sub

Private Sub SIS047()
Set adoRs = RunSP(SPCon, "usp_AnnualReport", 1)
Set cr = New crSIS047
cr.Database.SetDataSource adoRs
Me.Caption = "Annual Return Viewer"
End Sub

Private Sub SIS046D()
Me.Caption = gblPassData & " Alpha Stockholder Register"
Set adoRs = RunSP(SPCon, "usp_AlphaList", 1, gblOptNo)
Set cr = New CrSIS046D
cr.Database.SetDataSource adoRs
cr.ParameterFields.Item(1).AddCurrentValue gblPassData
Screen.MousePointer = vbDefault

End Sub

Private Sub SIS103()
Set adoRs = RunSP(SPCon, "usp_WorkingAllotment", 1)
Set cr = New crSIS103
cr.Database.SetDataSource adoRs
Me.Caption = "Working Allotment Report Viewer"
End Sub

Private Sub SIS054A()
Dim iErr As Integer, qSQL As String
Set cr = New crSIS054A
Set adoRs = RunSP(SPCon, "usp_IssuedCerts", 1)
cr.Database.SetDataSource adoRs
cr.ReportTitle = "Issued Certificates Report"
Me.Caption = "Issued Certificates Report Viewer"
End Sub

Private Sub SIS059()
Dim sql As String
sql = ""
Set cr = New crSIS059

'''adoRs.CursorLocation = adUseClient
Set adoRs = RunSP(SPCon, "usp_SIS059", 1)
'''adoRs.ActiveConnection = Nothing

'''cr.Database.Tables.Item(1).SetPrivateData 3, adoRs
cr.Database.SetDataSource adoRs
Me.Caption = "Trustee Report Viewer"
End Sub
Private Sub SIS104()
Set cr = New crSIS054

Me.Caption = "Certificate Register Report"
Set adoRs = RunSP(SPCon, "usp_PrintCloseCertRegister", 1, gblSiteId)
cr.Database.SetDataSource adoRs
End Sub
