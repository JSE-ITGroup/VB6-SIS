VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00C0C0C0&
   Caption         =   "NCB Jamaica Nominees  - Shareholder Information System"
   ClientHeight    =   4515
   ClientLeft      =   315
   ClientTop       =   3165
   ClientWidth     =   9465
   Icon            =   "SISMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "SISMDI.frx":030A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   6000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "Dat"
      DialogTitle     =   "Import Bank Recon Ascii File"
      Filter          =   "Text(*.txt;*.dat)|*.txt;*.dat"
      InitDir         =   "C:\"
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   372
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9405
      TabIndex        =   0
      Top             =   0
      Width           =   9465
      Begin VB.CommandButton btnClose 
         Caption         =   "E&xit"
         Height          =   315
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   615
      End
   End
   Begin ComctlLib.StatusBar txtStatusMsg 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4140
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Ready"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Select Printer"
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Switch &Registers"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Utilities"
         Index           =   2
         Begin VB.Menu mnuUtlItem 
            Caption         =   "Stock Exchange Import Options"
            Index           =   1
            Begin VB.Menu mnuUtlJCSD 
               Caption         =   "Import &Stock Exchange XLS Payments"
               Index           =   1
            End
            Begin VB.Menu mnuUtlJCSD 
               Caption         =   "View Stock Exchange File"
               Index           =   2
            End
            Begin VB.Menu mnuUtlJCSD 
               Caption         =   "Update Stock Exchange Categories Using a File"
               Index           =   3
            End
         End
         Begin VB.Menu mnuUtlItem 
            Caption         =   "&Export Reconciliation Data"
            Index           =   2
         End
         Begin VB.Menu mnuUtlItem 
            Caption         =   "Import Data for Dividend"
            Index           =   3
         End
         Begin VB.Menu mnuUtlItem 
            Caption         =   "Import BNS Recon File"
            Index           =   4
         End
         Begin VB.Menu mnuUtlItem 
            Caption         =   "Import New Register Data (1)"
            Index           =   5
         End
         Begin VB.Menu MnuSpecialBulk 
            Caption         =   "Special Bulk Postings"
            Index           =   0
            Begin VB.Menu MnuSpecialBulkItem 
               Caption         =   "Import Bulk Returns"
               Index           =   0
            End
            Begin VB.Menu MnuSpecialBulkItem 
               Caption         =   "Post Bulk Returns"
               Index           =   1
            End
            Begin VB.Menu MnuSpecialBulkItem 
               Caption         =   "Import Bulk Replacements"
               Index           =   2
            End
            Begin VB.Menu MnuSpecialBulkItem 
               Caption         =   "Post Bulk Replacements"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Change &User Password"
         Index           =   5
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Dividend Utilities"
         Index           =   7
         Begin VB.Menu MnuDividendItem 
            Caption         =   "Backup Dividend Files"
            Index           =   0
         End
         Begin VB.Menu MnuDividendItem 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuDividendItem 
            Caption         =   "Restore Dividend Files"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   9
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuedititem 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuedititem 
         Caption         =   "C&ut"
         Index           =   1
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuedititem 
         Caption         =   "&Copy"
         Index           =   2
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuedititem 
         Caption         =   "&Paste"
         Index           =   3
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuedititem 
         Caption         =   "&Delete"
         Index           =   4
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuLists 
      Caption         =   "&Lists"
      Begin VB.Menu mnuLstItem 
         Caption         =   "&Client Account"
         Index           =   0
      End
      Begin VB.Menu mnuLstItem 
         Caption         =   "&Tax Rates"
         Index           =   1
      End
      Begin VB.Menu mnuLstItem 
         Caption         =   "&Stockholder Categories"
         Index           =   2
      End
      Begin VB.Menu mnuLstItem 
         Caption         =   "&Bank List"
         Index           =   3
      End
      Begin VB.Menu mnuLstItem 
         Caption         =   "&JCSD Bank List"
         Index           =   4
      End
      Begin VB.Menu mnuLstItem 
         Caption         =   "&Finacle Accounts List"
         Index           =   5
      End
      Begin VB.Menu mnuLstItem 
         Caption         =   "C&urrency List"
         Index           =   6
      End
   End
   Begin VB.Menu mnuAct 
      Caption         =   "&Actions"
      Begin VB.Menu mnuActItem 
         Caption         =   "Stock &Transfers"
         Index           =   0
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "Stock &Issue"
         Index           =   1
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "B&onus Issues"
         Index           =   2
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "&Stock Splits"
         Index           =   3
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "&Certificate Production"
         Index           =   4
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "&Payment Processing"
         Index           =   6
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "&Interest Processing"
         Index           =   7
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "&Returned Cheques Processing"
         Index           =   8
         Begin VB.Menu MnuReturns 
            Caption         =   "Returned Cheques"
            Index           =   1
         End
         Begin VB.Menu MnuReturns 
            Caption         =   "Approve returned Cheques"
            Index           =   2
         End
         Begin VB.Menu MnuReturns 
            Caption         =   "Returned Cheques Report"
            Index           =   3
         End
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "Replacement &Cheque Processing"
         Index           =   9
         Begin VB.Menu MnuReplacements 
            Caption         =   "Replacement  &Cheque Entry"
            Index           =   1
         End
         Begin VB.Menu MnuReplacements 
            Caption         =   "Approve Pending Replacements"
            Index           =   2
         End
         Begin VB.Menu MnuReplacements 
            Caption         =   "Replaced Cheques Report"
            Index           =   3
         End
         Begin VB.Menu MnuReplacements 
            Caption         =   "Make Payments"
            Index           =   4
            Begin VB.Menu PaymentItem 
               Caption         =   "Print Replacement Cheques"
               Index           =   1
            End
            Begin VB.Menu PaymentItem 
               Caption         =   "Create Finacle File"
               Index           =   2
            End
            Begin VB.Menu PaymentItem 
               Caption         =   "Create ACH File"
               Index           =   3
            End
            Begin VB.Menu PaymentItem 
               Caption         =   "Print RTGS List"
               Index           =   4
            End
         End
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "&Bank Reconciliation"
         Index           =   10
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Import &Bank Statement"
            Index           =   0
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "&Automatic Reconciliation "
            Index           =   1
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Mark &Unmatched Bank Items"
            Index           =   2
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Mark Full Payments"
            Index           =   3
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Import Finacle Exception Report"
            Index           =   5
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Finacle Exceptions Update"
            Index           =   6
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Print Processed Exceptions"
            Index           =   7
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Import ACH Exception File"
            Index           =   9
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "ACH Exceptions Update"
            Index           =   10
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Print Processed Exceptions"
            Index           =   11
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "-"
            Index           =   12
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Dividend Reconciliaton Summary"
            Index           =   13
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "&Reconciled/UnReconciled Report by Dividend date"
            Index           =   14
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Reconciled/Unreconciled Bank items"
            Index           =   15
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "-"
            Index           =   16
         End
         Begin VB.Menu MnuBnkITm 
            Caption         =   "Reconciliation &Statement"
            Index           =   17
         End
      End
      Begin VB.Menu mnuActItem 
         Caption         =   "&Rights Issue Offer"
         Index           =   11
      End
   End
   Begin VB.Menu mnuAdm 
      Caption         =   "A&dministration"
      Begin VB.Menu mnuAdmItem 
         Caption         =   "&Users"
         Index           =   0
      End
      Begin VB.Menu mnuAdmItem 
         Caption         =   "&Company"
         Index           =   1
      End
      Begin VB.Menu mnuAdmItem 
         Caption         =   "&Preferences"
         Index           =   2
      End
      Begin VB.Menu mnuAdmItem 
         Caption         =   "&Archive Audit"
         Index           =   3
      End
      Begin VB.Menu mnuAdmItem 
         Caption         =   "&Who is Logged On"
         Index           =   4
      End
      Begin VB.Menu mnuAdmItem 
         Caption         =   "&Merge Accounts"
         Index           =   5
      End
      Begin VB.Menu mnuAdmItem 
         Caption         =   "&Update Missing Form Nos on Certificates Master"
         Index           =   6
      End
      Begin VB.Menu mnuAdmItem 
         Caption         =   "Approve Reconciling &Items"
         Index           =   7
      End
      Begin VB.Menu mnuAdmItem 
         Caption         =   "Maintain RTGS Limit"
         Index           =   8
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu RepItem 
         Caption         =   "Brokers &Summary"
         Index           =   0
      End
      Begin VB.Menu RepItem 
         Caption         =   "Brokers &Detail"
         Index           =   1
      End
      Begin VB.Menu RepItem 
         Caption         =   "&Percentage Ownership Exception Report"
         Index           =   2
      End
      Begin VB.Menu RepItem 
         Caption         =   "&Top Largest Shareholders List"
         Index           =   3
      End
      Begin VB.Menu RepItem 
         Caption         =   "Top N Largest Shareholders and Address"
         Index           =   4
      End
      Begin VB.Menu RepItem 
         Caption         =   "&Bank Mandate Listing"
         Index           =   5
      End
      Begin VB.Menu RepItem 
         Caption         =   "Stockholders' &Profile"
         Index           =   6
      End
      Begin VB.Menu RepItem 
         Caption         =   "Stockholders' &Register"
         Index           =   7
         Begin VB.Menu MnuRegItm 
            Caption         =   "&Multi-line register by Name"
            Index           =   0
         End
         Begin VB.Menu MnuRegItm 
            Caption         =   "&Single-line register by Category"
            Index           =   1
         End
         Begin VB.Menu MnuRegItm 
            Caption         =   "&Alpha Name && Address by Country && Category"
            Index           =   2
         End
      End
      Begin VB.Menu RepItem 
         Caption         =   "&Name && Address Labels"
         Index           =   8
         Begin VB.Menu LabItem 
            Caption         =   "&Create Mail Merge Datafile for Word Labels"
            Index           =   0
         End
         Begin VB.Menu LabItem 
            Caption         =   "&Dot Matrix Labels"
            Index           =   1
         End
      End
      Begin VB.Menu RepItem 
         Caption         =   "&Category Code Reference Lists"
         Index           =   9
      End
      Begin VB.Menu RepItem 
         Caption         =   "Ta&x Code Reference Lists"
         Index           =   10
      End
      Begin VB.Menu RepItem 
         Caption         =   "&Annual Returns"
         Index           =   11
      End
      Begin VB.Menu RepItem 
         Caption         =   "Print Closed Certificate Register"
         Index           =   12
      End
      Begin VB.Menu RepItem 
         Caption         =   "BOJ Return (Single)"
         Index           =   13
      End
      Begin VB.Menu RepItem 
         Caption         =   "BOJ Return (Full)"
         Index           =   14
      End
   End
   Begin VB.Menu mnuEnq 
      Caption         =   "E&nquiry"
      Begin VB.Menu mnuEnqItm 
         Caption         =   "&Stockholder"
         Index           =   0
      End
      Begin VB.Menu mnuEnqItm 
         Caption         =   "Stock E&xchange Accounts"
         Index           =   1
      End
      Begin VB.Menu mnuEnqItm 
         Caption         =   "&Broker"
         Index           =   2
      End
      Begin VB.Menu mnuEnqItm 
         Caption         =   "&Payments Rates"
         Index           =   3
      End
      Begin VB.Menu mnuEnqItm 
         Caption         =   "Bank Account No Search"
         Index           =   4
      End
      Begin VB.Menu mnuEnqItm 
         Caption         =   "Exchange Rates"
         Index           =   5
      End
      Begin VB.Menu mnuEnqItm 
         Caption         =   "Payment Summary"
         Index           =   6
      End
   End
   Begin VB.Menu MnuDataExtract 
      Caption         =   "Da&ta Extract"
      Begin VB.Menu MnuDataExtractItem 
         Caption         =   "Shareholder Data Extract"
         Index           =   0
      End
      Begin VB.Menu MnuDataExtractItem 
         Caption         =   "Extract Dividend"
         Index           =   1
      End
      Begin VB.Menu MnuDataExtractItem 
         Caption         =   "Shareholders As At"
         Index           =   2
      End
   End
   Begin VB.Menu MnuSysFunctions 
      Caption         =   "System Functions"
      Begin VB.Menu MnuSysFunctionsItems 
         Caption         =   "Restore Entire Database"
         Index           =   0
      End
      Begin VB.Menu MnuSysFunctionsItems 
         Caption         =   "Re-Open a Closed Batch"
         Index           =   1
      End
      Begin VB.Menu MnuSysFunctionsItems 
         Caption         =   "Reverse Posted Dividend"
         Index           =   2
      End
      Begin VB.Menu MnuSysFunctionsItems 
         Caption         =   "Password Parameters"
         Index           =   3
      End
      Begin VB.Menu MnuSysFunctionsItems 
         Caption         =   "Stock Exchange Maintenance"
         Index           =   4
      End
      Begin VB.Menu MnuSysFunctionsItems 
         Caption         =   "Delete Dividend Cheque Numbers"
         Index           =   5
      End
   End
   Begin VB.Menu MnuChqInventory 
      Caption         =   "&Cheque Inventory"
      Begin VB.Menu MnuChequeInventoryItems 
         Caption         =   "Cheque Inventory Management"
         Index           =   1
      End
      Begin VB.Menu MnuChequeInventoryItems 
         Caption         =   "Pending Cheque Transfers"
         Index           =   2
      End
      Begin VB.Menu MnuChequeInventoryItems 
         Caption         =   "Cheque Movement Enquiry"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Contents"
         Index           =   0
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Search for Help On..."
         Index           =   1
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About SIS"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msg As String
Dim repSISRept As New SISRepts
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Sub btnClose_Click()
Unload Me
End
End Sub

Private Sub LabItem_Click(Index As Integer)
Select Case Index
 Case 0
   gblOptions = 0
   frmSelLabel.Show 0
  Exit Sub
 Case 1
    gblOptions = 1
    frmSIS048.Show 0
    Exit Sub
End Select
End Sub

Private Sub MDIForm_Activate()
If Forms.Count = 1 Then frmMain.Show
Me.Caption = Trim(gblCompName) & "'s Register:- Shareholder Information System"

'Formally included in Load
Screen.MousePointer = vbHourglass
txtStatusMsg.SimpleText = gblWaitMsg
txtStatusMsg.Refresh
' enable menu & tb_buttons
'--
mnuFileItem(1).Enabled = True ' activate Close company
If gblUserLevel = 1 Then ' activate import option
 mnuFileItem(2).Enabled = True
End If
'mnuFileItem(4).Enabled = True ' activate Backup item
mnuedit.Enabled = True
If gblUserLevel <> gblViewOnly Then
  mnuLists.Enabled = True
  mnuAct.Enabled = True
  mnuReports.Enabled = True
End If
mnuEnq.Enabled = True
'--

If gblUserLevel = 1 Then
    mnuAdm.Enabled = True
    MnuSysFunctions.Enabled = True
Else
    MnuSysFunctions = False
    mnuAdm.Enabled = False
    MnuChequeInventoryItems(1).Enabled = False
    MnuChequeInventoryItems(2).Enabled = False
    MnuReturns(2).Enabled = False
    MnuReplacements(2).Enabled = False
End If

   
Screen.MousePointer = vbNormal
txtStatusMsg.SimpleText = gblReadyMsg
txtStatusMsg.Refresh

End Sub

Private Sub MDIForm_Load()

' wait while initialization completes
' set status msg to wait...
'--


End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatusMsg.SimpleText = gblReadyMsg
End Sub

Private Sub mnuInstall_Click()
msg = "Setup & install procedures for system owner"
txtStatusMsg.SimpleText = msg
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim n As Integer
Set frmMDI = Nothing
If gblUserLevel <> 0 Then
          n = LogOff()
        End If
End Sub

Private Sub mnuAct_Click()
msg = "Shareholder Main Activities"
txtStatusMsg.SimpleText = msg
End Sub

Private Sub mnuActItem_Click(Index As Integer)
Dim n As Integer
On Error GoTo mnuActItem_Err
' set status msg to wait...
' Screen.MousePointer = vbHourglass
' txtStatusMsg.SimpleText = gblWaitMsg
' txtStatusMsg.Refresh
Select Case Index
   Case 0 ' Transfers
     frmSIS025.Show 0
     Set frmSIS025 = Nothing
   Case 1  ' Stock Issue
     frmSIS022.Show 0
     Set frmSIS022 = Nothing
   Case 2 ' Bonus Issue
     frmSIS060.Show 0
     Set frmSIS060 = Nothing
   Case 3 'Stock Split Issue
     frmSIS090.Show 0
     Set frmSIS090 = Nothing
   Case 4   'Certificate Production
     frmSIS053.Show 0
     Set frmSIS053 = Nothing
   Case 6
     frmSIS013.Show 0 'Payment processing menu
     Set frmSIS013 = Nothing
   Case 7
     frmSIS013I.Show 0 'Interest Payment processing menu
     Set frmSIS013I = Nothing
   Case 11 ' Rights Issue Offer
    frmSIS100.Show 0  ' Rights Issue Switch Board
    Set frmSIS100 = Nothing
   Case Else
End Select
mnuActItem_Exit:
Exit Sub
mnuActItem_Err:
Select Case Err.Number
    Case 3021   'No Current Record
       Resume Next
    Case Else
       Call csvLogError("mnuAdminItem " & Index, Err.Number, Err.Description)
       MsgBox Err.Number & " " & Err.Description
       GoTo mnuActItem_Exit
    End Select
End Sub

Private Sub mnuAdm_Click()
msg = "Setup & maintain administrators data"
txtStatusMsg.SimpleText = msg
End Sub

Private Sub mnuAdmItem_Click(Index As Integer)
Dim n As Integer
On Error GoTo mnuAdmItem_err
' set status msg to wait...
 Screen.MousePointer = vbHourglass
 txtStatusMsg.SimpleText = gblWaitMsg
 txtStatusMsg.Refresh
 
Select Case Index
   Case 0
      frmSDI014.Show 0  ' Users
      Exit Sub
   Case 1
      frmSIS000.Show 0  ' Company Control
      Exit Sub
   Case 2
      frmSIS099.Show 0  ' Preferences
      Exit Sub
   Case 3
      frmSIS004.Show 0  ' Archive Audit
      Exit Sub
   Case 4
      frmSIS003.Show 0  ' Who is logged on
      Exit Sub
   Case 5
      frmSIS036.Show 0  ' Merge Accounts
      Exit Sub
   Case 6
      'frmFixCerts.Show 0 ' One off fix to certmaster
      Exit Sub
   Case 7
      FrmAuthRecon.Show 0 'Approve Reconciling Items
   Case 8
      FrmRTGSThreshold.Show 0 'RTGS Threshold maintenance
      Exit Sub
   Case Else
      
End Select
mnuAdmItem_Exit:
 On Error Resume Next
 Exit Sub

mnuAdmItem_err:
  
  Select Case Err.Number
    Case 3021   'No Current Record
       Resume Next
    Case Else
       Call csvLogError("mnuAdminItem " & Index, Err.Number, Err.Description)
       MsgBox Err.Number & " " & Err.Description
       GoTo mnuAdmItem_Exit
    End Select

End Sub

Private Sub mnuBnkAuto_Click()
Dim CountRec, Countrec2, TotCount
Dim rsAutoRec As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim i As Integer
Set rsAutoRec = New ADODB.Recordset
Set rsAutoRecSub = New ADODB.Recordset

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

If MsgBox("Confirm Automatic Reconciliation", vbYesNo) = vbYes Then
    frmMDI.txtStatusMsg.SimpleText = "Processing reconciliation file, Please wait......"
    frmMDI.txtStatusMsg.Refresh
    'Set rsAutoRec = RunSP(SpCon, "usp_mnuBnkAutoQry", 1)
    i = RunSP(SpCon, "usp_mnuBnkAutoUpdate", 0, Format(Now, "yyyymm"), _
        Format(Now, "yyyy/mm/dd"), gblLoginName)
    If i = 0 Then
       MsgBox "Automatic reconciliation completed"
    Else
       MsgBox "Reconciliation failed"
    End If
End If
    Set rsAutoRec = Nothing
    SpCon.Close
    MsgBox ("" & CountRec & " OF " & Count2 & " Record[s] have been reconciled"), vbInformation
        
End Sub

Private Sub mnuBnkItm_Click(Index As Integer)
Set repSISRept = New SISRepts
'repSISRept.DSN = gblDSN
repSISRept.LoginId = gblFileName
repSISRept.ReportType = 9
 '--
Select Case Index
Case 0 'Import Bank Data
 ' set status msg to wait...
 Screen.MousePointer = vbHourglass
 CmnDialog.ShowOpen
 If Len(CmnDialog.FileName) > 0 Then
   X = ImpBankRecon(CmnDialog.FileName)
 End If
Case 1 ' Automatic Reconciliation
  FrmAutoRecon.Show 0
Case 2 ' Mark Unreconciled Bank Items
  FrmBankItems.Show 0
Case 3 ' Mark full payments
  FrmSISItemsMatch.Show 0
Case 5
     CmnDialog.DialogTitle = "Import Finacle Exception Report"
     CmnDialog.Filter = "All files (*.*)*.*"
     CmnDialog.DefaultExt = ""
     CmnDialog.ShowOpen
     ImportFinacleExceptions
Case 6
     ProcessFinacleExceptions
Case 7
     gblFileKey = "2"
     FrmDates.Show 0
Case 9
     CmnDialog.DialogTitle = "Import ACH Exception File"
     CmnDialog.Filter = "All files (*.*)*.*"
     CmnDialog.DefaultExt = ""
     CmnDialog.ShowOpen
     ImportACHExceptions
Case 10
     ProcessACHExceptions
Case 11
     gblFileKey = "3"
     FrmDates.Show 0
Case 13 'Divdend Reconciliation Summary
       FrmSelectDivDate.Show 0
Case 14  'Reconciled/Unreconciled Cheque Report by Dividend Date
  FrmSelectDivDate.FmeReconciliation.Visible = True
  FrmSelectDivDate.OptDate.Visible = True
  FrmSelectDivDate.Show 0
  'FrmReportOption.Show 0
  
Case 15 'Reconciled/Unreconciled Bank Items
  FrmSelectBatchDate.Show 0
Case 17 ' Reconciliation Statement
  FrmReconPrint.Show 0
  
End Select
Set repSISRept = Nothing
End Sub

Private Sub MnuChequeInventoryItems_Click(Index As Integer)
Select Case Index
       Case 1
            FrmCIM.Show 0
       Case 2
            FrmAuthTransfers.Show 0
       Case 3
            FrmChqMovement.Show 0
       Case 5
            
End Select

End Sub

Private Sub MnuDataExtractItem_Click(Index As Integer)

Select Case Index
       Case 0
            FrmShareholderData.Show 0
       Case 1
            FrmExtractDividend.Show 0
       Case 2
            Dim SpCon As ADODB.Connection
            Dim adoRst As ADODB.Recordset
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
            gblFileKey = 4
            FrmDates.Show vbModal
            Set adoRst = RunSP(SpCon, "usp_ShareholdersAsAt", 1, Format(gblDate, "dd-mmm-yyyy"))
            Call ExportToExcel(adoRst)
            adoRst.Close
            Set adoRst = Nothing
            SpCon.Close
End Select
End Sub

Private Sub MnuDividendItem_Click(Index As Integer)
On Error GoTo Err_MnuDividendItem_Click
Dim i As Integer
Dim SpCon As ADODB.Connection

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
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

Select Case Index
      Case 0
           i = RunSP(SpCon, "usp_BackupTables", 0) ' Backup Tables
           If i <> 0 Then
              MsgBox "Backup Failed"
           Else
              MsgBox "Backup was successfull"
           End If
           SpCon.Close
      Case 2
           frmSISRestore.Show 0 'Restore tables
End Select

Exit_MnuDividendItem_Click:
Exit Sub
Err_MnuDividendItem_Click:
MsgBox "Error On Dividend Utilities"
GoTo Exit_MnuDividendItem_Click
End Sub

Private Sub mnuEdit_Click()
msg = "Copy, move, clear selections... etc."
txtStatusMsg.SimpleText = msg
End Sub

Private Sub mnuEditItem_Click(Index As Integer)
On Error GoTo ErrHandler
  Select Case Index
    Case 0 '            choose undo
      Screen.ActiveControl.SelText = gblHold
      mnuedititem(0).Enabled = False  ' Undo
      mnuedititem(1).Enabled = True     ' cut
      mnuedititem(2).Enabled = True     ' copy
      mnuedititem(3).Enabled = False    ' paste
      mnuedititem(4).Enabled = True     ' delete
    Case 1      ' choose cut
      Clipboard.Clear
      Clipboard.SetText Screen.ActiveControl.SelText
      gblHold = Screen.ActiveControl.SelText
      Screen.ActiveControl.SelText = ""
      mnuedititem(0).Enabled = True     ' undo
      mnuedititem(1).Enabled = False    ' cut
      mnuedititem(2).Enabled = False    ' copy
      mnuedititem(3).Enabled = True     ' paste
      mnuedititem(4).Enabled = False    ' delete
    Case 2      ' choose copy
      Clipboard.Clear
      Clipboard.SetText Screen.ActiveControl.SelText
      gblHold = Screen.ActiveControl.SelText
      mnuedititem(0).Enabled = True     ' undo
      mnuedititem(1).Enabled = False    ' cut
      mnuedititem(2).Enabled = False    ' copy
      mnuedititem(3).Enabled = True     ' paste
      mnuedititem(4).Enabled = False    ' delete
    Case 3      ' choose paste
      glbhold = Screen.ActiveControl.SelText
      Screen.ActiveControl.SelText = Clipboard.GetText
      mnuedititem(0).Enabled = True     ' undo
      mnuedititem(1).Enabled = True    ' cut
      mnuedititem(2).Enabled = True    ' copy
      mnuedititem(3).Enabled = True     ' paste
      mnuedititem(4).Enabled = True    ' delete
    Case 4      ' choose delete
      gblHold = Screen.ActiveControl.SelText
      Screen.ActiveControl.SelText = ""
      mnuedititem(0).Enabled = True     ' undo
      mnuedititem(1).Enabled = True   ' cut
      mnuedititem(2).Enabled = True    ' copy
      mnuedititem(3).Enabled = True     ' paste
      mnuedititem(4).Enabled = True    ' delete
  End Select

Exit Sub

ErrHandler:
If Err.Number = 91 Or Err.Number = 438 Then
 MsgBox "No text selected"
Else
 MsgBox Error
End If
Exit Sub




End Sub


Private Sub mnuEnq_Click()
msg = "Enquiry Options"
txtStatusMsg.SimpleText = msg
End Sub

Private Sub mnuEnqItm_Click(Index As Integer)
On Error GoTo mnuEnqItm_Err
Select Case Index
Case 0 'Stockholder
    frmSIS070.Show 0
Case 1 'JCSD Accounts
    frmSIS070J.Show 0
Case 2 'Brokers
    frmSIS027.Show 0
    frmSIS027.cmdEdit.Enabled = False
    frmSIS027.cmdPrint.Enabled = False
Case 3 'Payment Rates
    frmSIS077.Show 0
Case 4 ' Bank Account No Search
    FrmMandate.Show 0
Case 5 ' Exchange rates enquiry
    frmExchangeRateList.Show 0
Case 6
    FrmDividendDetails.Show 0
Case Else
End Select
mnuEnqItm_Exit:
Exit Sub
mnuEnqItm_Err:
 Call MsgBox("mnuEnqItm" & Index)
 Call csvLogError("mnuEnqItm" & Index, Err.Number, Err.Description)
  GoTo mnuEnqItm_Exit
End Sub

Private Sub mnuFile_Click()
  msg = "Open, Backup, Restore and Close files...etc"
  txtStatusMsg.SimpleText = msg
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
Dim n As Integer, sTmp As String
'Dim cnn As New cADOAccess
Dim WrkDSN As String
On Error GoTo Cancel_Click
Select Case Index
   Case 0 'Set Printer
   CmnDialog.ShowPrinter
   
   Case 1 ' Close database and disable all options
     LogOff
     n = CloseAllForms("frmMDI")
     ' Clear global variables set by open Open Company
     gblOpenComp = ""
     '''cnn = ""
     '--
     WrkDSN = InputBox("Enter Register's DSN to open:", "File Open")
     
     If Len(WrkDSN) = 0 Then End
     SDILogin.Show
     
     SDILogin.txtField(0) = Trim(gblLoginName)
     SDILogin.txtField(1) = Trim(gblPassword)
     SDILogin.txtField(2) = Trim(WrkDSN)
     
     SDILogin.btnOk = True
     
ReTrySwitch:
     SDILogin.txtField(0) = Trim(gblLoginName)
     SDILogin.txtField(2) = Trim(WrkDSN)
        
     If gblOpenComp <> "O" Then
        SDILogin.btnOk = False
        GoTo ReTrySwitch
     End If
     frmMain.Show
   Case 4 'Compact Database
     Dim sNewfile As String, sNewName2 As String
    
     msg = "Compacting make take awhile. Please wait..."
     txtStatusMsg.SimpleText = msg
     Screen.MousePointer = vbHourglass
    
    ' Make sure there isn't already a file with the
    ' name of the compacted database.
    sNewfile = App.Path & "\payback.mdb"
    If Dir(sNewfile) <> "" Then Kill sNewfile
     
    'Compact the database
    'db.Close
    DBEngine.CompactDatabase cnn, sNewfile
    ' overwrite old file
    Kill cnn ' delete the old file
    Name sNewfile As cnn ' rename the file
    txtStatusMsg.SimpleText = gblReadyMsg
    Screen.MousePointer = vbDefault
   Case 5   ' Maintain Passwords
     frmChangePassword.txtUserName = gblLoginName
     frmChangePassword.Show vbModal
  Exit Sub
   Case 9  'Exit
     Unload Me
     End
   End Select

FileItem_exit:
   On Error Resume Next
   Exit Sub
   
Cancel_Click:
   If Index = 0 Then     'Cancel from open
      Exit Sub
   Else
      GoTo mnuFileItem_err
   End If

mnuFileItem_err:
 Select Case Err.Number
    Case 3021   'No Current Record
       Resume Next
    Case Else
      Call MsgBox("mnuFileItem" & Index)
      Call csvLogError("mnuFileItem " & Index, Err.Number, Err.Description)
      GoTo FileItem_exit
       
End Select
End Sub

  
Private Sub mnuHelp_Click()
 msg = "Access information for using VBSIS"
 txtStatusMsg.SimpleText = msg
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
Dim nRet As Integer
'--
Select Case Index
Case 0  'Contents
  On Error Resume Next
   nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
   CmnDialog.HelpFile = "sishelp.hlp"
   CmnDialog.HelpCommand = sisHelp
   CmnDialog.ShowHelp   ' Display Visual Basic Help contents topic.
  If Err Then
    MsgBox Err.Description
  End If
Case 1   'Search
On Error Resume Next
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
  If Err Then
    MsgBox Err.Description
  End If
Case 3 'about
     SISAbout.Show vbModal
Case Else
End Select
End Sub

Private Sub mnuLists_Click()
msg = "Setup & maintain shareholder's reference information"
txtStatusMsg.SimpleText = msg
End Sub

Private Sub mnuLstItem_Click(Index As Integer)
On Error GoTo mnuLstItem_Click_Err
 '--
 ' set status msg to wait...
 Screen.MousePointer = vbHourglass
 txtStatusMsg.SimpleText = gblWaitMsg
 txtStatusMsg.Refresh
 '--
 Select Case Index
 Case 0
   frmSIS001.Show 0  ' Clients list
  Exit Sub
 Case 1
   frmSIS005.Show 0   ' Tax Listing
 Case 2
   frmSIS007.Show 0   ' Shareholder Category
 Case 3
   frmSIS020.Show 0    ' Bank mandate Reference Lists
 Case 4
   frmSIS020J.Show 0   'JCSD Bank Reference List
 Case 5
   frmAccounts.Show 0  'Finacle Acounts Listing
 Case 6
   frmCurrencyList.Show 0 'Currency listing
 Case Else
 End Select
mnuLstItem_Click_Exit:
On Error Resume Next
 Exit Sub
mnuLstItem_Click_Err:
Select Case Err.Number
    Case 3021   'No Current Record
       Resume Next
    Case Else
        MsgBox "mnuLstItem"
       Call csvLogError("mnuLstItem " & Index, Err.Number, Err.Description)
       MsgBox Err.Number & " " & Err.Description
       GoTo mnuLstItem_Click_Exit
    End Select
End Sub

Private Sub mnuRegItm_Click(Index As Integer)
Set repSISRept = New SISRepts
'repSISRept.DSN = gblDSN
repSISRept.LoginId = gblFileName
repSISRept.ReportType = 9

Select Case Index
Case 0 ' Alphabetic register listing
    gblOptions = 6
    FrmSelectStockExchange.Show 0
    Exit Sub
Case 1 ' Category listing
    gblOptions = 7
    FrmSelectStockExchange.Show 0
    Exit Sub
Case 2 ' Alpha Name & Address by Country & Category
    gblOptions = 8
    FrmSelectStockExchange.Show 0
    Exit Sub
End Select
Set repSISRept = Nothing
End Sub

Private Sub MnuReplacements_Click(Index As Integer)
On Error GoTo Err_MnuReplacements_Click

Select Case Index
Case 1 'Replace Cheque
 ' set status msg to wait...
 Screen.MousePointer = vbHourglass
 txtStatusMsg.SimpleText = gblWaitMsg
 txtStatusMsg.Refresh
 frmSIS084.Show 0
Case 2 ' Approve pending replacements
     FrmApproveReplacements.Show 0
Case 3 ' Replacement cheque report
     gblFileKey = "0"
     FrmDates.Show 0
End Select

Exit_MnuReplacements_Click:
Exit Sub

Err_MnuReplacements_Click:
MsgBox Err.Description, vbOKOnly, "Replacement cheques menu Error"
GoTo Exit_MnuReplacements_Click
End Sub

Private Sub mnuReports_Click()
msg = "Displays or prints main system reports"
txtStatusMsg.SimpleText = msg
End Sub


Private Sub MnuReturns_Click(Index As Integer)
On Error GoTo Err_MnuReturns_Click

Select Case Index
Case 1 ' Record Returned Bank Cheques
 FrmReturnedChqs.Show 0
Case 2 'Approve Returned Cheques
    FrmApproveReturns.Show 0
Case 3 ' Returned Cheques report
     gblFileKey = "1"
     FrmDates.Show 0
End Select

Exit_MnuReturns_Click:
Exit Sub

Err_MnuReturns_Click:
MsgBox Err.Description, vbOKOnly, "Returns cheques menu Error"
GoTo Exit_MnuReturns_Click
End Sub

Private Sub MnuSpecialBulkItem_Click(Index As Integer)
On Error GoTo Err_MnuSpecialBulkItem_Click
Dim StrSql As String
Dim lngRecsAff As Long

Select Case Index
       Case 0
            CmnDialog.DialogTitle = "Import Returns XL File"
            CmnDialog.Filter = "XLS(*.xls)|*.xls"
            CmnDialog.DefaultExt = "XLS"
            CmnDialog.ShowOpen
            If Len(CmnDialog.FileName) > 0 Then
               lngRecsAff = ImportExcel2(CmnDialog.FileName, "ReturnListing")
               StrSql = "Import completed successfully" & vbCrLf
               StrSql = StrSql & lngRecsAff & " records imported"
               MsgBox StrSql
            End If
       Case 1
            gblOptions = 1
            FrmSelectAccount.Show 0
       Case 2
            CmnDialog.DialogTitle = "Import Replacements XL File"
            CmnDialog.Filter = "XLS(*.xls)|*.xls"
            CmnDialog.DefaultExt = "XLS"
            CmnDialog.ShowOpen
            If Len(CmnDialog.FileName) > 0 Then
              lngRecsAff = ImportExcel(CmnDialog.FileName, "BulkReplacements")
              StrSql = "Import completed successfully" & vbCrLf
              StrSql = StrSql & lngRecsAff & " records imported"
              MsgBox StrSql
            End If
      Case 3
           gblOptions = 2
           FrmSelectAccount.Show 0
End Select

Exit_MnuSpecialBulkItem_Click:
Exit Sub

Err_MnuSpecialBulkItem_Click:
MsgBox Err.Description, vbOKOnly, "Bulk Processing Error"
GoTo Exit_MnuSpecialBulkItem_Click
End Sub

Private Sub MnuSysFunctionsItems_Click(Index As Integer)
On Error GoTo Err_MnuSysFunctionsItems_Click

Select Case Index
      Case 0
           FrmDBRestore.Show 0
      Case 1
           FrmReOpenBatch.Show 0
      Case 2
           FrmDelPostedDiv.Show 0
      Case 3
           FrmPasswordParameters.Show 0
      Case 4
           FrmStockExchange.Show 0
      Case 5
           gblOptions = 9
           FrmSelectStockExchange.Show 0
End Select

Exit_MnuSysFunctionsItems_Click:
Exit Sub

Err_MnuSysFunctionsItems_Click:
MsgBox Err.Description, vbOKOnly, "System Functions Error"
GoTo Exit_MnuSysFunctionsItems_Click
End Sub

Private Sub mnuUtlItem_Click(Index As Integer)
Dim X As Integer
'--
' set status msg to wait...
Screen.MousePointer = vbHourglass
txtStatusMsg.SimpleText = gblWaitMsg
txtStatusMsg.Refresh
Select Case Index
'Case 0  'Dataflex Import
'--
'    Import.Show 0

Case 2 ' Export Reconciliation Data
        exportrecdata.Show
Case 3 ' Import Excel payments in XL format
   CmnDialog.DialogTitle = "Import Client XL File"
   CmnDialog.Filter = "XLS(*.xls)|*.xls"
   CmnDialog.DefaultExt = "XLS"
   CmnDialog.ShowOpen
   If Len(CmnDialog.FileName) > 0 Then
     ImpExcel.Show 0
   End If
Case 4 ' Import Bank recon Excel File
  CmnDialog.DialogTitle = "Select Recon XL File"
  CmnDialog.Filter = "XLS(*.xls)|*.xls"
  CmnDialog.DefaultExt = "XLS"
  CmnDialog.ShowOpen
 If Len(CmnDialog.FileName) > 0 Then
   X = ImpBankReconXL(CmnDialog.FileName)
 End If
Case 5 ' Import New Register Data Format 1 (BNS)
   CmnDialog.DialogTitle = "Import Register Data XL File"
   CmnDialog.Filter = "XLS(*.xls)|*.xls"
   CmnDialog.DefaultExt = "XLS"
   CmnDialog.ShowOpen
   If Len(CmnDialog.FileName) > 0 Then
     ImpRegisterExcel.Show 0
   End If
End Select
 Screen.MousePointer = vbDefault
txtStatusMsg.SimpleText = gblReadyMsg
txtStatusMsg.Refresh
End Sub

Private Sub mnuUtlItemrec_Click(Index As Integer)
exportrecdata.Show
End Sub

Private Sub mnuUtlJCSD_Click(Index As Integer)
Dim X As Integer
'--
' set status msg to wait...
Screen.MousePointer = vbHourglass
txtStatusMsg.SimpleText = gblWaitMsg
txtStatusMsg.Refresh
Select Case Index
'Case 0  'Dataflex Import
'--
'    Import.Show 0
Case 1 ' Import SE payments in XL format
   gblOptions = 1
   FrmSelectStockExchange.Show 0
Case 2
    FrmJCSD.Show 0

Case 3 ' Import SE Categories in XL format
   gblOptions = 2
   FrmSelectStockExchange.Show 0
End Select

Screen.MousePointer = vbDefault
txtStatusMsg.SimpleText = gblReadyMsg
txtStatusMsg.Refresh
    
End Sub

Private Sub mnuWindow_Click() '
msg = "Rearrange windows or activate specified window"
txtStatusMsg.SimpleText = msg
End Sub


Private Sub PaymentItem_Click(Index As Integer)
On Error GoTo Err_PaymentItem_Click
Dim dFileName As String
Dim adoRst As ADODB.Recordset
Dim SpCon As New ADODB.Connection

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
    
Set adoRst = RunSP(SpCon, "usp_CurrentLedger", 1)
If adoRst.EOF Then
   MsgBox "Unable to determine current ledger"
   GoTo Exit_PaymentItem_Click
End If
dFileName = adoRst!CompanyABBR & "_"
adoRst.Close
Set adoRst = Nothing

Select Case Index
       Case 1
            FrmPrintRepChq.Show 0
       Case 2
            With frmMDI.CmnDialog
                 .DialogTitle = "Create Finacle File"
                 .Filter = "TXT(*.txt)|*.txt"
                 .DefaultExt = "txt"
                 dFileName = dFileName & "Finacle"
                 .FileName = dFileName
                 .CancelError = True
                 .ShowSave
                 CreateFinacleFile ("R")
            End With
       Case 3
            With frmMDI.CmnDialog
                 .DialogTitle = "Create ACH File"
                 .Filter = "TXT(*.txt)|*.txt"
                 .DefaultExt = "txt"
                 dFileName = dFileName & "ACH"
                 .FileName = dFileName
                 .CancelError = True
                 .ShowSave
                 CreateACHFile ("R")
            End With
       Case 4
            Set adoRst = RunSP(SpCon, "usp_RTGSDetailsR", 1)
            If Not adoRst.EOF Then
               Call ExportToExcel(adoRst)
            End If
            adoRst.Close
            Set adoRst = Nothing
End Select
SpCon.Close

Exit_PaymentItem_Click:
Exit Sub
Err_PaymentItem_Click:
If Err.Number <> cdlCancel Then
   MsgBox Err & " " & Err.Number, vbOKOnly, "Replacement Cheques Electronic File Creation"
   GoTo Exit_PaymentItem_Click
Else
   MsgBox "You cancelled this option"
   GoTo Exit_PaymentItem_Click
End If
End Sub

Private Sub RepItem_Click(Index As Integer)
Dim rs As New ADODB.Recordset
Dim repSISRept As New SISRepts
Dim sql As String
Dim SpCon As ADODB.Connection
'--
Set repSISRept = New SISRepts
repSISRept.LoginId = gblFileName
repSISRept.ReportType = 9
repSISRept.siteid = gblSiteId

Select Case Index
 Case 0         ' Brokers Summary list
    repSISRept.ReportNumber = 0
    repSISRept.RunShareHolderReport
    Exit Sub
 Case 1     ' Brokers Detail
    repSISRept.ReportNumber = 1
    repSISRept.RunShareHolderReport
    Exit Sub
 Case 2   'Percentage Ownership Exception Report
    repSISRept.ReportNumber = 2
    repSISRept.RunShareHolderReport
    Exit Sub
 Case 3   ' List n largest shareholders
    gblOptions = 3
    FrmSelectStockExchange.Show 0
    Exit Sub
 Case 4     ' List n largest shareholders and addresses
    gblOptions = 4
    FrmSelectStockExchange.Show 0
    Exit Sub
 Case 5     ' Bank Mandate Listing
    repSISRept.ReportNumber = 4
    repSISRept.RunShareHolderReport
    Exit Sub
 Case 6
    frmSIS045.Show 0 ' Stockholders Profile
    Exit Sub
 Case 9 'Category Reference Lists
    Dim Rep8 As New crSIS050
    Dim sMsg8, sTitle8, sDefault8, sValue8 As String
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

    '--
    sMsg8 = "Should the Category Code Viewer be Activated"
    sTitle8 = "List Category Codes"
    sDefault8 = "N"
    sValue8 = InputBox(sMsg8, sTitle8, sDefault8)
    If IsNothing(sValue8) Then
      Set rs = RunSP(SpCon, "usp_CategoryList", 1)
      Rep8.Database.SetDataSource rs
      Rep8.DiscardSavedData
      Rep8.ReadRecords
      Rep8.PrintOut
    Else
      If UCase(sValue8) = "N" Then
        Set rs = RunSP(SpCon, "usp_CategoryList", 1)
        Rep8.Database.SetDataSource rs
        Rep8.ReadRecords
        Rep8.DiscardSavedData
        Rep8.PrintOut
     Else
        frmSIS050.Show 0
     End If
   End If
   SpCon.Close
   Exit Sub
   Case 10  'Tax code Lists
    Dim Rep9 As New crSIS049
    Dim sMsg9, sTitle9, sDefault9, sValue9 As String
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

    '--
    sMsg9 = "Should the Tax Code Viewer be Activated"
    sTitle9 = "List tax Codes"
    sDefault9 = "N"
    sValue9 = InputBox(sMsg9, sTitle9, sDefault9)
    If IsNothing(sValue9) Then
        Set rs = RunSP(SpCon, "usp_TaxList", 1)
        Rep9.Database.SetDataSource rs
        Rep9.ReadRecords
        Rep9.DiscardSavedData
        Rep9.PrintOut
    Else
        If UCase(sValue9) = "N" Then
          Set rs = RunSP(SpCon, "usp_TaxList", 1)
          Rep9.Database.SetDataSource rs
          Rep9.DiscardSavedData
          Rep9.PrintOut
        Else
           frmSIS049.Show 0
        End If
    End If
    SpCon.Close
    Exit Sub
  Case 11 'Annual Returns
    frmSIS047.Show 0
  Case 12
       FrmSelectBatch.Show vbModal
       If Len(gblFileKey) > 6 Then
          repSISRept.ReportType = 9
          repSISRept.ReportNumber = 18
          repSISRept.siteid = gblFileKey
          repSISRept.RunShareHolderReport
       End If
  Case 13
       gblFileKey = "5"
       FrmDates.Show 0
  Case 14
       gblFileKey = "6"
       FrmDates.Show 0
  Case Else
End Select
RepItem_Exit:
  Exit Sub

End Sub

Private Sub ImportFinacleExceptions()
Dim fs, F, iRecs As Long
Dim textfile As String
Dim sInRec As String
Dim sRec As String
Dim X As Integer
Dim StrSql As String
Dim sMsg As String
Dim iAccountNo As String
Dim iAmount As Currency
Dim iDate As Date
Dim SpCon As ADODB.Connection

sMsg = "You are about to import a Finacle Exception Report "
sMsg = sMsg & "Select No if you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
X = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If X = vbNo Then
  Exit Sub
End If

textfile = frmMDI.CmnDialog.FileName
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Importing Finacle Exception Report..."
frmMDI.txtStatusMsg.Refresh

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
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.opentextfile(textfile)
sInRec = F.readline
iRecs = 0
If F.atendofstream = True Then
  MsgBox "Input Text File " & textfile & " is blank; import aborting... "
  F.Close
End If
Do Until iRecs = 4
   sInRec = F.readline
   iRecs = iRecs + 1
Loop
StrSql = Mid(sInRec, 1, 10)
iDate = CDate(StrSql)

StartProcessing:
iRecs = 0

Do Until F.atendofstream = True
   sRec = Trim(Mid(sInRec, 1, 9))
   If IsNumeric(sRec) = False Then
      GoTo ReadAnother
   End If
   iAccountNo = Mid(sInRec, 16, 9)
   iAmount = CCur(Trim(Mid(sInRec, 58, 16)))
   StrSql = Trim(Mid(sInRec, 77, 20))
   iRecs = iRecs + 1
   X = RunSP(SpCon, "usp_ImportFinacleException", 0, Format(iDate, "dd-mmm-yyyy"), iAccountNo, iAmount, StrSql, gblLoginName)
ReadAnother:
   sInRec = F.readline
   frmMDI.txtStatusMsg.SimpleText = "Processing record " & iRecs
   frmMDI.txtStatusMsg.Refresh
Loop

F.Close
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = "Importation completed"
frmMDI.txtStatusMsg.Refresh

End Sub
Private Sub ImportACHExceptions()
Dim fs, F, iRecs As Long
Dim textfile As String
Dim sInRec As String
Dim sRec As String
Dim X As Integer
Dim iReason As String
Dim iClientID As Double
Dim sMsg As String
Dim iAccountNo As String
Dim iAmount As Currency
Dim iTrace As String
Dim CompanyID As String
Dim iDate As Date
Dim SpCon As ADODB.Connection

sMsg = "You are about to import a ACH Exception File "
sMsg = sMsg & "Select No if you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
X = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If X = vbNo Then
  Exit Sub
End If

textfile = frmMDI.CmnDialog.FileName
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Importing ACH Exception File..."
frmMDI.txtStatusMsg.Refresh

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
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.opentextfile(textfile)
sInRec = F.readline
iRecs = 0
If F.atendofstream = True Then
  MsgBox "Input Text File " & textfile & " is blank; import aborting... "
  F.Close
End If
sInRec = F.readline

CompanyID = Trim(Mid(sInRec, 18, 10))
iDate = Mid(sInRec, 56, 8)

StartProcessing:
iRecs = 0

Do Until F.atendofstream = True
   iAccountNo = Mid(sInRec, 12, 17)
   iAmount = CCur(Trim(Mid(sInRec, 29, 17)))
   iClientID = Trim(Mid(sInRec, 47, 14))
   iReason = Trim(Mid(sInRec, 101, 44))
   iTrace = Trim(Mid(sInRec, 153, 7))
   iRecs = iRecs + 1
   X = RunSP(SpCon, "usp_ImportACHException", 0, Format(iDate, "dd-mmm-yyyy"), iAccountNo, iAmount, iReason, iClientID, iTrace, gblLoginName)
ReadAnother:
   sInRec = F.readline
   frmMDI.txtStatusMsg.SimpleText = "Processing record " & iRecs
   frmMDI.txtStatusMsg.Refresh
Loop

F.Close
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = "Importation completed"
frmMDI.txtStatusMsg.Refresh
SpCon.Close
End Sub
Private Sub ProcessFinacleExceptions()
Dim SpCon As ADODB.Connection
Dim i As Integer

Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Processing Finacle Exceptions..."
frmMDI.txtStatusMsg.Refresh

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
Loop

i = RunSP(SpCon, "usp_UpdateFromFinacleException", 0)
If i = 0 Then
    MsgBox "Finacle Exceptions processed"
Else
    MsgBox "Error on processing Finacle exceptions"
End If
SpCon.Close

End Sub
Private Sub ProcessACHExceptions()
Dim SpCon As ADODB.Connection
Dim i As Integer

Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Processing ACH Exceptions..."
frmMDI.txtStatusMsg.Refresh

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
Loop

i = RunSP(SpCon, "usp_UpdateFromACHException", 0)
If i = 0 Then
    MsgBox "ACH Exceptions processed"
Else
    MsgBox "Error on processing ACH exceptions"
End If
SpCon.Close

End Sub
