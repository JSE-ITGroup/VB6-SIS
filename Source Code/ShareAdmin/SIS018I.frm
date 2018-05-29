VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS018I 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Interest Cheque Switchboard"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9030
   Icon            =   "SIS018I.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9030
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   4935
      Left            =   0
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Ok"
      Height          =   315
      Index           =   0
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox tb 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   7680
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox tb 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2415
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
      FieldSeparator  =   ","
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   6085
      Columns(0).Caption=   "Cheque Name"
      Columns(0).Name =   "ChequeName"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "Report Name"
      Columns(1).Name =   "ReportName"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   4260
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "To Replace Cheque No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Starting Cheque No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Cheque Format:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "Create &Main Cheque Print File"
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Create &Sub-Ledger Print File"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Create TTSE Sub-Ledger Print File"
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Print &All Cheques"
         Index           =   5
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Print ALL &JCSD Cheques"
         Index           =   6
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Enter TT Exchange Rate"
         Index           =   8
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Print All TTSE Cheques"
         Index           =   9
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Restart Print"
         Index           =   11
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Restart JCSD &Print"
         Index           =   12
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Restart TTSE Print"
         Index           =   13
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Extract Financle File"
         Index           =   15
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Print Financle List"
         Index           =   16
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Generate Bank Cheques"
         Index           =   18
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Print Bank Cheques"
         Index           =   19
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Print Bank List (Internal)"
         Index           =   20
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Print Bank Letters"
         Index           =   21
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Print Bank List (External)"
         Index           =   22
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   23
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Shareholders Advice "
         Index           =   24
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   25
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   26
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmSIS018I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iChq As Long
Dim mytest As String
Dim iErr As Integer
Dim iOpen As Integer
Dim iRst As Integer
Dim iLedger As String
Dim pname As String
Dim pport As String
Dim pdriver As String, sLedg As String
Dim psize As Integer
Dim porient As String
Dim SpCon As ADODB.Connection

Private Sub cmdBtn_Click(Index As Integer)
Dim sql As String
Select Case Index
   Case 0
    cmdBtn(0).Visible = False
    If IsNothing(tb(0)) Then
      iErr = 167
      csvShowUsrErr iErr, "Print Cheques"
      tb(0).SetFocus
      Exit Sub
    End If
    mnuFile.Enabled = True
    '--- test for restart
    If iRst = True Then
      If IsNothing(tb(1)) Then
        iErr = 167
        csvShowUsrErr iErr, "Print Cheques"
        tb(1).SetFocus
        Exit Sub
      End If
     frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
     Screen.MousePointer = vbHourglass
     frmMDI.txtStatusMsg.Refresh
    ''' If MakChq(2, Val(tb(1)), Val(tb(0)), sLedg) Then
    '''   sql = ""
    '''   sql = "SELECT " _
   '''      & "StkPymnts.PayTyp, StkPymnts.ClientId, StkPymnts.GrossPymnt, " _
     '''    & "StkPymnts.WhldTax, StkPymnts.Shares, StkPymnts.ChqNum, " _
    '''     & "DivRef.RecDate, DivRef.ChqDate, DivRef.PayPer, DivRef.IncTyp, " _
    '''     & "StkTaxr.TaxRate, " _
    '''     & "CHQTRN.PayeeName, CHQTRN.MndAddr1, CHQTRN.MndAddr2, CHQTRN.mndAcntNme, CHQTRN.mndAccnt, "
    '''   If sLedg = "M" Then
    '''     sql = sql & "a.CliName, a.CliAddr1, a.CliAddr2, a.CliAddr3, a.CliAddr4, a.CliAddr5 "
    '''   Else
    '''     sql = sql & "a.GR8NAM, a.GR8Ad1, a.GR8AD2, a.GR8AD3, ' ', ' ' "
    '''   End If
    '''  sql = sql & ", company.ParValue, DivRef.Remarks From " _
         & "Company, (((StkPymnts INNER JOIN StkTaxr ON " _
         & "StkPymnts.ResCode = StkTaxr.ResCode) " _
         & "INNER JOIN CHQTRN ON " _
         & "StkPymnts.ChqNum = CHQTRN.ChqNum) "
     '''If sLedg = "M" Then
         sql = sql & "INNER JOIN StkName a ON StkPymnts.ClientId = a.ClientId ) "
    '''Else
       '''  If sLedg = "S" Then
          '''  sql = sql & "INNER JOIN JCSDSub a on stkPymnts.ClientId = a.GR8NIN )"
       '''  Else
          '''  sql = sql & "INNER JOIN TTSESub a on stkPymnts.ClientId = a.GR8NIN )"
       '''  End If
   ''' End If
      '''sql = sql & "INNER JOIN DivRef ON " _
         '''& "StkPymnts.DecDate = DivRef.DecDate AND " _
         & "StkPymnts.PayTyp = DivRef.PayTyp " _
         & "Where StkPymnts.ClientId <> Company.NextAcct and " _
         & "StkPymnts.GrossPymnt > 0. " _
         & " and StkPymnts.ChqNum >= " & Val(tb(0)) _
         & " and StkPymnts.CatCode in (select catcode from stkcat where hold = 0) " _
         & " Order By StkPymnts.ChqNum "
     ''' Set adoRs = New ADODB.Recordset
      '''cr.Database.Tables.Item(1).SetLogOnInfo gblDSN
      '''adoRs.Open sql, cnn, adOpenDynamic, adLockReadOnly
     ''' iOpen = True
     ' cr.Database.SetDataSource adoRs
     '''cr.Database.Tables.Item(1).SetPrivateData 3, adoRs
     '''cr.Database.SetDataSource adoRs
     '''CRV.ReportSource = cr
    ''' CRV.ViewReport

     '''crv.ReportSource = cr
     '''cr.DiscardSavedData
     '''crv.ViewReport
     Else
      'MsgBox "Create Print procedure failed"
     End If
     frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
     Screen.MousePointer = vbDefault
     frmMDI.txtStatusMsg.Refresh
    '''End If
   Case 1 ' End Job
     EndJob
   Case Else
End Select
End Sub

Private Sub dbc_InitColumnProps()
Dim i As Integer
With dbc
     .RemoveAll
     For i = 1 To 9
      .AddItem Choose(i, "Courts,1", "D&G,2", "Dyoll,3", _
         "First Life,4", "Goodyear,5", "NCB,6", _
         "Pegasus,7", "Producers,8", "AIC Bonds,9")
    Next i
End With
End Sub

Private Sub dbc_LostFocus()
Select Case dbc.Columns(1).Text
Case 1 ' Courts Cheques - crCTS018
 Set cr = New crNCB018
Case 2 ' D&G - crDG018
 Set cr = New crDG018
Case 3 ' Dyoll Cheques - crDYL018
 Set cr = New crFL018
Case 4 ' First Life - crFL018
 Set cr = New crFL018
Case 5 ' Goodyear
 Set cr = New crDG018
Case 6 ' NCB - crNCB018
 Set cr = New crNCB018
Case 7 ' Pegasus
 Set cr = New crNCB018
Case 8 ' Producer - crJBP018
 Set cr = New crJBP018
Case 9 'AIC BONDS
 Set cr = New crAIC018
Case Else ' Default
 Set cr = New crNCB018
End Select
   'assign default system printer to crystal
  'porient = cr.PaperOrientation
  'psize = cr.PaperSize
  'pname = Printer.DeviceName
  pdriver = Printer.DriverName
  'pport = Printer.Port
  'Call cr.SelectPrinter(pdriver, pname, pport)
  'cr.PaperOrientation = porient
  'cr.PaperSize = psize
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
 iRst = False
 'Set cr = New crSIS018
 'dbc.Columns(1).Text = "crSIS018"
 GetDefChqNo
 iOpen = False
End Sub

Private Sub Form_Resize()
'CRV.Top = ScaleTop
crv.Left = 0
'CRV.Height = ScaleHeight
crv.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)

If iOpen = True Then adoRs.Close
Set adoRs = Nothing
Set cr = Nothing
Set frmSIS018I = Nothing
SpCon.Close
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
Dim iRep As Integer, iErr As Integer
Dim sql As String

Select Case Index
   Case 1  ' Create Cheque Print file
   sLedg = "M"
    iRep = MsgBox("You are about to create the cheque print file," _
            & vbCrLf & "This will overwrite the previous cheque file data." _
             , vbInformation + vbOKCancel, "Make Cheques")
    If iRep = vbCancel Then
         EndJob
    Else
     If IsNothing(Trim(tb(0).Text)) Then
       iErr = 167
       csvShowUsrErr iErr, "Print Cheques"
       tb(0).SetFocus
     Else
       frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
       Screen.MousePointer = vbHourglass
       frmMDI.txtStatusMsg.Refresh
       '--
       iChq = Val(tb(0))
       iRep = RunSP(SpCon, "usp_MakChq", 0, 1, iChq, 0, sLedg)
       '--
       frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
       Screen.MousePointer = vbDefault
       frmMDI.txtStatusMsg.Refresh
       If iRep = 2 Then
          MsgBox "Create Print procedure failed"
          EndJob
      End If
     End If
   End If
   Case 2 ' Make JCSD Cheques
   sLedg = "S"
    iRep = MsgBox("You are about to create the JCSD cheque print file," _
            & vbCrLf & "This will overwrite the previous cheque file data." _
             , vbInformation + vbOKCancel, "Make Cheques")
    If iRep = vbCancel Then
         EndJob
    Else
     If IsNothing(Trim(tb(0).Text)) Then
       iErr = 167
       csvShowUsrErr iErr, "Print Cheques"
       tb(0).SetFocus
     Else
       frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
       Screen.MousePointer = vbHourglass
       frmMDI.txtStatusMsg.Refresh
       '--
       iChq = Val(tb(0))
       iRep = RunSP(SpCon, "usp_MakChq", 0, 1, iChq, 0, sLedg)
       '--
       frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
       Screen.MousePointer = vbDefault
       frmMDI.txtStatusMsg.Refresh
       If iRep = 2 Then
        MsgBox "Create Print procedure failed"
        EndJob
      End If
     End If
   End If
   '--
  Case 3 ' Make TTSE Cheques
   sLedg = "T"
    iRep = MsgBox("You are about to create the TTSE cheque print file," _
            & vbCrLf & "This will overwrite the previous cheque file data." _
             , vbInformation + vbOKCancel, "Make Cheques")
    If iRep = vbCancel Then
         EndJob
    Else
     If IsNothing(Trim(tb(0).Text)) Then
       iErr = 167
       csvShowUsrErr iErr, "Print Cheques"
       tb(0).SetFocus
     Else
       frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
       Screen.MousePointer = vbHourglass
       frmMDI.txtStatusMsg.Refresh
       '--
       iChq = Val(tb(0))
       iRep = RunSP(SpCon, "usp_MakChq", 0, 1, iChq, 0, sLedg)
       '--
       frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
       Screen.MousePointer = vbDefault
       frmMDI.txtStatusMsg.Refresh
       If iRep = 2 Then
        MsgBox "Create Print procedure failed"
        EndJob
      End If
     End If
   End If
    
   Case 5, 6, 9 'Print All Cheques
     Set adoRs = RunSP(SpCon, "usp_printIntChq", 1, Index)
     '''*** Another approach to using stored procedure and this time
     '''*** getting a dynamic recordset
     '''mytest = "usp_printChq(" & Index & ")"
     '''Set adoRs = New ADODB.Recordset
     '''adoRs.Open mytest, SpCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc
     '''***End
     tb(0).Enabled = False
     iOpen = True
     Screen.MousePointer = vbDefault
     
     If Not adoRs.EOF Then
        cr.PrinterSetup Me.hwnd
        cr.Database.SetDataSource adoRs
        crv.ReportSource = cr
        crv.ViewReport
     End If
   Case 8  'Accept T&T Exchange Rate
       Dim sMsg As String
       Dim sTitle As String
       Dim sRate As String
       
       sMsg = "Make sure the rate you are entering is correct." & vbCrLf
       sMsg = sMsg & "If Not the calculation process will have to be re-run to correct the error!!!"
       sTitle = "T&T Exchange Rate Setup & Calculation"
       sRate = 10
       sRate = InputBox(sMsg, sTitle, sRate)
       If IsEmpty(sRate) Or sRate = "" Then
          EndJob
       End If
       
       iRep = RunSP(SpCon, "usp_TTRate", 0, CCur(sRate))
       If iRep = 2 Then
          MsgBox "An error occurred while updating the rate. Please contact Systems Department"
          EndJob
       End If
       
   Case 11, 12, 13 'Restart prining of all cheques
     If Index = 11 Then
       sLedg = "M"
     Else
       If Index = 12 Then
          sLedg = "S"
       Else
          sLedg = "T"
       End If
     End If
     tb(1).Visible = True
     lbl(2).Visible = True
     cmdBtn(0).Visible = True
     iRep = MsgBox("Enter spoilt cheque number in second box." _
            & vbCrLf & "Replacement no in first box." _
            & vbCrLf & "All remaining cheques will be renumbered." _
         , vbInformation + vbOKCancel, "Make Cheques")
    If iRep = vbCancel Then
         EndJob
    Else
        iRst = True
    End If
   Case 15 ' Create Finacle Export File
         CreateFinacleFile ("D")
   Case 16 ' Print Finacle Listing
           If iLedger = "M" Then
              Set adoRs = RunSP(SpCon, "usp_DividendList", 1, 3)
              Me.Caption = "Finacle Main Payment Register Viewer"
           End If
           If iLedger = "J" Then
              Set adoRs = RunSP(SpCon, "usp_DividendList", 1, 6)
              Me.Caption = "Finacle JCSD Payment Register Viewer"
           End If

           Set cr = New crSIS017
           cr.Database.SetDataSource adoRs
           crv.ReportSource = cr
           crv.ViewReport

   Case 18 ' Generate Other Banks File
            iRep = RunSP(SpCon, "usp_BnkPymnt", 0, iChq)
            If iRep = 2 Then
               MsgBox "Generation of Other Banks File Failed"
            End If
            
   Case 19 'Print Cheques for accounts at other Banks
            Set adoRs = RunSP(SpCon, "usp_printBnkChq", 1)
            Set cr = New crNCB018B
            cr.Database.SetDataSource adoRs
            crv.ReportSource = cr
            crv.ViewReport

   Case 20 'Print Bank Listing of Shareholders
           If iLedger = "M" Then
              Set adoRs = RunSP(SpCon, "usp_DividendList", 1, 4)
              Me.Caption = "Other Banks' Main Payment Register Viewer"
           End If
           If iLedger = "J" Then
              Set adoRs = RunSP(SpCon, "usp_DividendList", 1, 5)
              Me.Caption = "Other Banks' JCSD Payment Register Viewer"
           End If
           
           Set cr = New crSIS017B
           cr.Database.SetDataSource adoRs
           crv.ReportSource = cr
           crv.ViewReport

   Case 21 'Print Letters to Other Banks
           Set adoRs = RunSP(SpCon, "usp_printBnkChq", 1)
           Set cr = New CRBankLetter
           Set cr = New CRBankLetter
           cr.Database.SetDataSource adoRs
           crv.ReportSource = cr
           crv.ViewReport

   Case 22 ' Print list to be sent to The Banks
           Set adoRs = RunSP(SpCon, "usp_BnkLetter", 1)
           Set cr = New CRBnkList
           cr.Database.SetDataSource adoRs
           crv.ReportSource = cr
           crv.ViewReport

   Case 24
           Set adoRs = RunSP(SpCon, "usp_CustomerAdvJCSD", 1)
           Set cr = New crNCB018C
           cr.Database.SetDataSource adoRs
           Me.Caption = "Shareholder Advice Viewer"
           crv.ReportSource = cr
           crv.ViewReport

   Case 26
     EndJob
   Case Else
End Select

End Sub

Private Sub GetDefChqNo()
Dim rsComp As New ADODB.Recordset
Dim iOpenComp As Integer
Dim sql As String
On Error GoTo GetDefChqNo_Err
'--
'Get default starting cheque number
'----------------------------------
'''sql = "SELECT nextchq from COMPANY"
'''Set rsComp = New ADODB.Recordset
'''iOpenComp = False
Set rsComp = RunSP(SpCon, "usp_GetDefChqNo", 1, 0)
iOpenComp = True
If Not rsComp.EOF Then
   iChq = rsComp!NEXTCHQ
   iLedger = rsComp!Ledger
Else
   iChq = 0
End If
'--
'lbl(1).Visible = True
tb(0).Text = iChq
'tb(0).Visible = True
GetDefChqNo_Exit:
If iOpenComp = True Then rsComp.Close
Set rsComp = Nothing
Exit Sub
GetDefChqNo_Err:
 MsgBox "SIS018I/GetDefChqNo"
  GoTo GetDefChqNo_Exit
End Sub

Private Sub EndJob()
Set cr = Nothing
Set frmSIS018I = Nothing
frmSIS013I.Visible = True
Unload Me
End Sub

Private Sub tb_Validate(Index As Integer, Cancel As Boolean)
Dim iErr As Integer
Select Case Index
Case 0 'Validate Chq number
  If IsNothing(tb(Index)) Then
     iErr = 169
     csvShowUsrErr iErr, "Print Cheques"
    Cancel = True
  Else
     If Not IsNumeric(tb(Index)) Then
       iErr = 28
       csvShowUsrErr iErr, "Print Cheques"
       Cancel = True
     End If
  End If
Case Else
End Select
End Sub
