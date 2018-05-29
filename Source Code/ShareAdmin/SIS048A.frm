VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRVIEWER.DLL"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3A32.OCX"
Begin VB.Form frmSIS048 
   Caption         =   "Name & Address Labels"
   ClientHeight    =   6510
   ClientLeft      =   2715
   ClientTop       =   1320
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9435
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Find"
      Default         =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   5160
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Index           =   1
      Left            =   6240
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cbOptions 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Select the search argument."
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ComboBox cbWhere 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select search field which will also be used to sequence the report."
      Top             =   600
      Width           =   1725
   End
   Begin VB.TextBox tbFind 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Enter selection value. Use ""ALL"" to select entire file."
      Top             =   240
      Width           =   3345
   End
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   9255
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
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
   End
   Begin SSDataWidgets_A.SSDBOptSet Opt 
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   1560
      Width           =   1470
      _Version        =   196611
      _ExtentX        =   2593
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "3-Up"
      BackColor       =   -2147483643
      Cols            =   2
      IndexSelected   =   1
      NumberOfButtons =   2
      Buttons.Button(0).OptionValue=   "1"
      Buttons.Button(0).Caption=   "1-Up"
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   38
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   40
      Buttons.Button(0).PictureRight=   39
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   48
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(1).OptionValue=   "3"
      Buttons.Button(1).Caption=   "3-Up"
      Buttons.Button(1).Mnemonic=   83
      Buttons.Button(1).Value=   -1  'True
      Buttons.Button(1).TextLeft=   64
      Buttons.Button(1).TextRight=   87
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   49
      Buttons.Button(1).ButtonRight=   62
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   89
      Buttons.Button(1).PictureRight=   88
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   49
      Buttons.Button(1).ButtonToColRight=   97
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).ButtonBitmapID=   2
      Buttons.Button(1).Column=   1
   End
   Begin VB.Line Line4 
      X1              =   7320
      X2              =   7320
      Y1              =   0
      Y2              =   2040
   End
   Begin VB.Line Line3 
      X1              =   2520
      X2              =   7320
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      X1              =   2520
      X2              =   2520
      Y1              =   0
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   7320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   300
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Labels"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   2640
      TabIndex        =   8
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   6
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Where"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPreviewApplication 
         Caption         =   "Preview (&Application Window)"
      End
      Begin VB.Menu mnuPrinterSetup 
         Caption         =   "Printer &Setup..."
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
Attribute VB_Name = "frmSIS048"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iOpen As Integer

Private Sub cmdFind_Click(Index As Integer)
Dim i As Integer, X As Integer
Dim sWhere As String, sql As String, sFldName As String
On Error GoTo cmdFind_Err
Select Case Index
Case 0
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Selecting Records for Labels..."
'----------------------------
'-------- main line ---------
'----------------------------
sql = "SELECT a.CLIENTID, CLINAME, CLIADDR1,"
sql = sql & " CLIADDR2, CLIADDR3, CLIADDR4, CLIADDR5,"
sql = sql & " 'A' , 'B', 'C', 'D', 'E'"
sql = sql & " FROM STKNAME a, STKCAT b"
sql = sql & " WHERE a.CATCODE = b.CATCODE"
sql = sql & " and SHARES > 0 and CLIADDR1 <> 'ADDRESS UNKNOWN' "
'--
If tbFind.Text = vbNullString Then
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Exit Sub
End If
sWhere = UCase(Trim(tbFind.Text))
If sWhere = "" Then
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Exit Sub
End If
'---
X = cbWhere.ListIndex
Select Case X
Case 0
  sFldName = "CATDESC"
Case 1
  sFldName = "a.CATCODE"
Case 2
  sFldName = "CLINAME"
Case 3
  sFldName = "CLIENTID"
End Select
'---
If cbOptions.Enabled = False Then 'create labels for all
  GoTo SET_ORDER_BY
End If
'---
If cbOptions.ListIndex > 6 Then cbOptions.ListIndex = 0
Select Case cbOptions.ListIndex
Case 0 ' Exact Match
   sql = sql & " and " & sFldName
   If X = 3 Then  'numeric key for client id
      sql = sql & " = " & Val(sWhere)
   Else
      sql = sql & " = '" & sWhere & "' "
   End If
Case 1 ' Starts With
   sWhere = Trim(tbFind.Text) & "%"
Case 2 ' Ends With
   sWhere = "%" & Trim(tbFind.Text)
Case 3 ' AnyWhere
   sWhere = "%" & Trim(tbFind.Text) & "%"
End Select
If cbOptions.ListIndex <> 0 Then
 sql = sql & " and " & sFldName & " like '" & sWhere & "' "
End If
'---
SET_ORDER_BY:
 sql = sql & " order by " & sFldName
 If sFldName = "CATDESC" Then
    sql = sql & ", CLINAME"
 Else
    If sFldName = "CLINAME" Then
      sql = sql & ", CLIENTID"
    End If
 End If
'---
If Opt.IndexSelected = 1 Then
  Set cr = New crSIS048T
Else
  Set cr = New crSIS048O
End If
cr.Database.Tables.Item(1).SetLogOnInfo gblDSN
If iOpen = True Then
  iOpen = 0
  adoRs.Close
End If
adoRs.Open sql, gblFileName, adOpenDynamic, adLockReadOnly
iOpen = -1
cr.Database.Tables.Item(1).SetPrivateData 3, adoRs
Me.Caption = "Name & Address Labels"
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Case 1
  Unload Me
  
End Select
cmdFind_Exit:
  Exit Sub
cmdFind_Err:
  csvShowError "SIS048/cmdFind"
   Unload Me
  
End Sub



Private Sub crv_PrintButtonClicked(UseDefault As Boolean)
frmPrintOut.Show vbModal
End Sub

Private Sub Form_Load()
On Error GoTo FormLoad_Err
Set adoRs = New ADODB.Recordset
csvCenterForm Me, gblMDIFORM
frmMDI.btnClose.Enabled = False
Set cnn = New ADODB.Connection
cnn.Open gblFileName
iOpen = 0
' readymsg
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
csvCenterForm Me, gblMDIFORM
'- load comparison key fields and show frmFind
'---------------------------------------------
 cbWhere.AddItem "Category Name"
 cbWhere.AddItem "Category Code"
 cbWhere.AddItem "Client Name"
 cbWhere.AddItem "Account No"
 cbWhere.ListIndex = 0
 'cbOptions.ListIndex = 0
 cbOptions.Enabled = False
 tbFind.Text = vbNullString
 '--
FormLoad_Exit:
 Exit Sub
FormLoad_Err:
  
  csvShowError "SIS048/Load"
 GoTo FormLoad_Exit
End Sub

Private Sub Form_Resize()
    Me.crv.Width = Me.ScaleWidth
    Me.crv.Height = (Me.ScaleHeight - Me.crv.top)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Clean up
Set cr = Nothing
If iOpen = -1 Then adoRs.Close
Set adoRs = Nothing
Set frmSIS048 = Nothing
frmMDI.btnClose.Enabled = True
End Sub

Private Sub mnuExit_Click()
    Unload Me
    Set cnn = Nothing
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
    Me.crv.Width = frmSIS048.ScaleWidth

    Me.crv.Height = (Me.ScaleHeight - Me.crv.top)
 
 
    'Bring the Printing Status form to the front
    frmPrintingStatus.Show
    frmPrintingStatus.SetFocus
 
End Sub

Private Sub mnuPrinterSetup_Click()
    frmPrinterSetup.Show vbModal
End Sub


Private Sub tbFind_LostFocus()
If tbFind.Text <> vbNullString Then
   If UCase(Trim(tbFind.Text)) <> "ALL" Then
      cbOptions.Enabled = True
   Else
      cbOptions.Enabled = False
   End If
End If
End Sub
