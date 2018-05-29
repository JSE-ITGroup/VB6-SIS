VERSION 5.00
Begin VB.Form frmSIS060 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Bonus Menu"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SIS060.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4830
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Update Bonus Certs"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Updates Accounts with Bonus Allocations"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   375
      Index           =   7
      Left            =   3360
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      ToolTipText     =   "Returns to main menu"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Bonus &Allotment Report"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "Report shows the bonus allocations without fractions"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Print Bonus &Letters"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Prints Bonus Certs"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Working Allotment Report"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Report shows the bonus allocations with fractions"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Calcualte  Allocations"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Calculates Bonus based on Entry Information"
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Bonus &Entry"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Activates the  Bonus Information Entry Form"
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmSIS060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBon As New ADODB.Recordset
Dim rsAct As New ADODB.Recordset
Dim repSISRept As New SISRepts
Dim OpenErr As Integer
Dim iOpenBon As Integer
Dim iOpenAct As Integer, sNewfile As String
Dim unPk As cUnPackName
Dim dDecDate As Date
Private Sub cmdBtn_Click(Index As Integer)
' set status msg to wait...
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'--
Select Case Index
Case 0
frmSIS060.Visible = False
cnnClose
frmSIS061.Show 0 ' Bonus Entry
'--
Case 1
frmSIS060.Visible = False
frmSIS062.Show 0 ' Calculate Bonuss
'--
Case 2 ' print working allotment
repSISRept.ReportNumber = 63
repSISRept.RunShareHolderReport
'--
Case 3  ' Print allotment report
repSISRept.ReportNumber = 64
repSISRept.RunShareHolderReport
Case 4 ' Update Bonus Certs
frmSIS060.Visible = False
frmSIS065.Show 0
'--
Case 5 'Print letters
Dim Message As String, Title As String, Default As Date, MyValue As String
Message = "Enter the Bonus Declaration Date"   ' Set prompt.
Title = "Bonus Letters Entry"
Default = Format(Now, "dd/mm/yyyy")
' Display message, title, and default value.
'MyValue = InputBox(Message, Title, Default)
' Use Helpfile and context. The Help button is added automatically.
'MyValue = InputBox(Message, Title, , , , "SISHELP.HLP", 10)
' Display dialog box at position 100, 100.
MyValue = InputBox(Message, Title, Default, 100, 100)
If Not IsNothing(MyValue) And IsDate(MyValue) Then
 dDecDate = DateValue(MyValue)
 '''Set cnn = New ADODB.Connection
 Set unPk = New cUnPackName
 Create_bonus_letters
 If iOpenBon = True Then rsBon.Close
 If iOpenAct = True Then rsAct.Close
 Set rsBon = Nothing
 Set rsAct = Nothing
 '''set cnn = nothing
 Set unPk = Nothing
End If
Case 7
 
 '--
 cnnClose
 Unload Me
Case Else
End Select
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
End Sub

Private Sub Form_Load()
 csvCenterForm Me, gblMDIFORM
 '--  disable menu items
 '----------------------
 frmMDI.mnuFile.Enabled = False
 frmMDI.btnClose.Enabled = False
 frmMDI.mnuLists.Enabled = False
 frmMDI.mnuAct.Enabled = False
 frmMDI.mnuAdm.Enabled = False
 Set repSISRept = New SISRepts
 repSISRept.LoginId = gblFileName
 repSISRept.ReportType = 9
 repSISRept.siteid = gblSiteId
 'repSISRept.DSN = gblDSN
 
  ' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmMDI.mnuFile.Enabled = True
 frmMDI.btnClose.Enabled = True
 frmMDI.mnuLists.Enabled = True
 frmMDI.mnuAct.Enabled = True
 If gblUserLevel = 1 Then frmMDI.mnuAdm.Enabled = True
End Sub

Private Sub Create_bonus_letters()
Dim fso, txtfile, iErr As Integer
Dim sOutRec As String, sql As String
OpenErr = False
iOpenAct = False
iOpenBon = False
Set rsAct = New ADODB.Recordset
Set rsBon = New ADODB.Recordset

'--
On Error GoTo Create_bonus_err
sNewfile = App.Path & "\sis066.txt"
cnn.Open
rsBon.Open "BonusRef", cnn, , , adCmdTable
iOpenBon = True
'--
sql = "Select a.clientid, a.IssDate, a.CertNo, a.shares, " _
      & "cliname, cliaddr1, cliaddr2, cliaddr3, cliaddr4, " _
      & "cliaddr5, compname, parvalue " _
      & "From ((Company INNER JOIN STKACTIV a ON " _
      & "Company.nextcert <> a.certno ) " _
      & "INNER JOIN STKNAME b ON " _
      & "b.ClientId = a.ClientId) " _
      & "where FORM = 'BONUS' " _
      & "and Brokerbuy = 0 " _
      & "and BrokerId = 0 and CertNo > 0 and Status = 'O' " _
      & "and IssDate =#" & Format(rsBon!RECDAT, "mm/dd/yyyy") _
      & "# order by CERTNO"
rsAct.Open sql, cnn, , , adCmdText
iOpenAct = True
'--
'create txt file for printing bonus letters
With rsAct
   If .EOF Then ' SELECT RETURNED NO RECORDS
   '---
     iErr = 164
     csvShowUsrErr iErr, "Print Bonus letters"
   Else
   ' process returned recordset
   '--
      iErr = 0
      frmMDI.txtStatusMsg.SimpleText = "Creating Merge File....."      'open file for output & write header record
      '--
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set txtfile = fso.CreateTextFile(sNewfile, True)
      sOutRec = ""
      sOutRec = sOutRec & "CLINAME" & Chr(9)
      sOutRec = sOutRec & "ADDRESS1" & Chr(9)
      sOutRec = sOutRec & "ADDRESS2" & Chr(9)
      sOutRec = sOutRec & "ADDRESS3" & Chr(9)
      sOutRec = sOutRec & "ADDRESS4" & Chr(9)
      sOutRec = sOutRec & "ADDRESS5" & Chr(9)
      sOutRec = sOutRec & "CLIENTID" & Chr(9)
      sOutRec = sOutRec & "DECDATE" & Chr(9)
      sOutRec = sOutRec & "RECDATE" & Chr(9)
      sOutRec = sOutRec & "BASE" & Chr(9)
      sOutRec = sOutRec & "BONUS" & Chr(9)
      sOutRec = sOutRec & "PAR" & Chr(9)
      sOutRec = sOutRec & "CERTNO" & Chr(9)
      sOutRec = sOutRec & "SHARES" & Chr(9)
      sOutRec = sOutRec & "COMPNAME"
      txtfile.writeline (sOutRec)
      ' BUILD DATA FILE FROM RECORDSET
      '--
      .MoveFirst
      Do While Not .EOF
        sOutRec = ""
        If unPk.Unpack(!CliName) = True Then
            sOutRec = sOutRec & unPk.FirstName & _
                   " " & unPk.LastName & Chr(9)
        Else
           sOutRec = sOutRec & !CliName & Chr(9)
        End If
        sOutRec = sOutRec & !CliAddr1 & Chr(9)
        sOutRec = sOutRec & !CliAddr2 & Chr(9)
        sOutRec = sOutRec & !CliAddr3 & Chr(9)
        sOutRec = sOutRec & !CliAddr4 & Chr(9)
        sOutRec = sOutRec & !CliAddr5 & Chr(9)
        sOutRec = sOutRec & !ClientID & Chr(9)
        sOutRec = sOutRec & Format(dDecDate, "mmmm, dd yyyy") & Chr(9)
        sOutRec = sOutRec & Format(!IssDate, "mmmm, dd yyyy") & Chr(9)
        sOutRec = sOutRec & rsBon!STKSTO & Chr(9)
        sOutRec = sOutRec & rsBon!STKBASE & Chr(9)
        sOutRec = sOutRec & !PARVALUE & Chr(9)
        sOutRec = sOutRec & !certno & Chr(9)
        sOutRec = sOutRec & !shares & Chr(9)
        sOutRec = sOutRec & !compname
        '--
        txtfile.writeline (sOutRec)
        .MoveNext
      Loop
      txtfile.Close
   End If
   iOpen = False
   
End With
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
If iErr = 164 Then Exit Sub
Print_Report
create_bonus_exit:
cnnClose
Exit Sub
Create_bonus_err:
 MsgBox "Create Bonus Letters'"
 GoTo create_bonus_exit

End Sub
Private Sub Print_Report()
Dim AppWord As Word.Application
Dim sDoc As String, iWarn As String
On Error GoTo Open_error
iWarn = 183
csvShowUsrErr iWarn, "Bonus Letters"
frmMDI.txtStatusMsg.SimpleText = "Running Word to Print Bonus letters..."
sDoc = App.Path & "\sis066.doc"
Set AppWord = CreateObject("Word.application")
Screen.MousePointer = vbDefault
With AppWord
 .Documents.Open (sDoc)
 .ActiveDocument.MailMerge.OpenDataSource (sNewfile)
End With
With AppWord.ActiveDocument.MailMerge
  .Destination = wdSendToPrinter
  '.Destination = wdSendToNewDocument
  .MailAsAttachment = False
  .MailAddressFieldName = ""
  .MailSubject = ""
  .SuppressBlankLines = True
  With .DataSource
      .FirstRecord = wdDefaultFirstRecord
      .LastRecord = wdDefaultLastRecord
  End With

 .Execute
End With
AppWord.ActiveDocument.Close
AppWord.DisplayAlerts = False
AppWord.Quit

Exit Sub
Open_error:
  
  MsgBox "Word Open Error " & Err.Number & "-" & Err.Description, , "Labels Selection"
 
End Sub

