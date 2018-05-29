VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS027 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brokers Enquiry"
   ClientHeight    =   3555
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   7845
   Icon            =   "SIS027.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   7845
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   328
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   328
      Left            =   4200
      TabIndex        =   3
      Top             =   3120
      Width           =   1005
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3012
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   7620
      _Version        =   196617
      DataMode        =   2
      GroupHeaders    =   0   'False
      Col.Count       =   4
      UseGroups       =   -1  'True
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Groups(0).Width =   12938
      Groups(0).Caption=   "Selling Information"
      Groups(0).Columns.Count=   4
      Groups(0).Columns(0).Width=   6297
      Groups(0).Columns(0).Caption=   "Broker's Name"
      Groups(0).Columns(0).Name=   "Broker's Name"
      Groups(0).Columns(0).DataField=   "Column 0"
      Groups(0).Columns(0).DataType=   8
      Groups(0).Columns(0).FieldLen=   30
      Groups(0).Columns(1).Width=   2487
      Groups(0).Columns(1).Caption=   "Account No"
      Groups(0).Columns(1).Name=   "Account No"
      Groups(0).Columns(1).DataField=   "Column 1"
      Groups(0).Columns(1).DataType=   8
      Groups(0).Columns(1).FieldLen=   10
      Groups(0).Columns(2).Width=   2090
      Groups(0).Columns(2).Caption=   "Held"
      Groups(0).Columns(2).Name=   "Held"
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   2
      Groups(0).Columns(2).FieldLen=   10
      Groups(0).Columns(3).Width=   2064
      Groups(0).Columns(3).Caption=   "Available"
      Groups(0).Columns(3).Name=   "Available"
      Groups(0).Columns(3).DataField=   "Column 3"
      Groups(0).Columns(3).DataType=   3
      Groups(0).Columns(3).FieldLen=   11
      _ExtentX        =   13441
      _ExtentY        =   5318
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   328
      Left            =   6720
      TabIndex        =   1
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "C&ertifications"
      Height          =   328
      Left            =   5280
      TabIndex        =   0
      Top             =   3120
      Width           =   1365
   End
   Begin VB.Label lbl 
      BackColor       =   &H000000FF&
      Caption         =   "Available"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H000000FF&
      Caption         =   "Selected Broker Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmSIS027"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ErrWarn As New cLstWarn
Dim X As Integer
Dim SpCon As ADODB.Connection
Dim rsBrokers As New ADODB.Recordset
Dim errLoop As Error
Dim errs1 As Error
Dim repSISRept As New SISRepts

Private Sub cmdClose_Click()
If Isloaded("frmSIS028") Then
  X = ErrWarn.ListWarn()
  frmSIS028.Show
  Exit Sub
Else
   rsBrokers.Close
   Set rsBrokers = Nothing
  Unload Me
End If
End Sub
Private Sub cmdEdit_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'-- store client id & form in filekey
gblFileKey = grd.Columns(1).Text
lbl(0).Caption = grd.Columns(0).Text
lbl(1).Caption = grd.Columns(3).Text
gblOptions = 2
frmSIS028.Show 0
Me.Visible = False
End Sub

Private Sub cmdFind_Click()
Dim i As Integer
Dim bm As Variant, cChk As Integer
Dim sWhere As String, sGrid As String
Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
     .cbWhere.Clear
    .cbWhere.AddItem "BrokerName"
    .cbWhere.AddItem "AccountNo"
    .cbWhere.ListIndex = 0
    .cbOptions.ListIndex = 0
    .Show vbModal
    '----------------------------
    '-------- main line ---------
    '----------------------------
    If .tbFind.Text = vbNullString Then
    Else
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      sWhere = UCase(Trim(.tbFind.Text))
      X = .cbWhere.ListIndex
      '---
      If sWhere <> "" Then
    
          grd.Redraw = False
          If .optBtn.OptionValue = 1 Then grd.MoveFirst
          Select Case .cbOptions.ListIndex
          Case 0 ' Exact Match
               For i = 0 To grd.Rows - 1
                  bm = grd.GetBookmark(i)
                  sGrid = UCase(grd.Columns(X).CellText(bm))
                  If sWhere = sGrid Then
                        grd.Bookmark = grd.GetBookmark(i)
                        Exit For
                   End If
                Next i
           Case 1 ' Starts With
               sWhere = UCase(Trim(.tbFind.Text) & "%")
               For i = 0 To grd.Rows - 1
                    bm = grd.GetBookmark(i)
                    sGrid = UCase(grd.Columns(X).CellText(bm))
                    cChk = sGrid Like sWhere
                    If cChk = True Then
                        grd.Bookmark = grd.GetBookmark(i)
                        Exit For
                    End If
                Next i
            Case 2 ' Ends With
               sWhere = UCase("%" & Trim(.tbFind.Text))
               
               For i = 0 To grd.Rows - 1
                    bm = grd.GetBookmark(i)
                    sGrid = UCase(grd.Columns(X).CellText(bm))
                    cChk = sGrid Like sWhere
                    If cChk = True Then
                        grd.Bookmark = grd.GetBookmark(i)
                        Exit For
                    End If
                Next i
             Case 3 ' AnyWhere
               sWhere = UCase("%" & Trim(.tbFind.Text) & "%")
               
               For i = 0 To grd.Rows - 1
                    bm = grd.GetBookmark(i)
                    sGrid = UCase(grd.Columns(X).CellText(bm))
                    cChk = sGrid Like sWhere
                    If cChk = True Then
                        grd.Bookmark = grd.GetBookmark(i)
                        Exit For
                    End If
                Next i
             End Select
              grd.Redraw = True
      End If
    End If
  End With
  Unload frmFind
  Set frmFind = Nothing
End Sub
Private Sub cmdPrint_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

'repSISRept.DSN = gblDSN
repSISRept.LoginId = gblFileName
repSISRept.ReportNumber = 14
repSISRept.ReportType = 9
repSISRept.siteid = gblSiteId
repSISRept.RunShareHolderReport
End Sub
Private Sub Form_GotFocus()
UpdateScreen
End Sub


Private Sub grd_Click()
gblFileKey = grd.Columns(1).CellValue(grd.GetBookmark(0))
End Sub

Private Sub grd_DblClick()
' disable dblclick for viewonly users
If gblUserLevel <> gblViewOnly Then
  ' wait message & hourglass
  Screen.MousePointer = vbHourglass
  frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
  frmMDI.txtStatusMsg.Refresh
  lbl(0).Caption = grd.Columns(0).CellValue(grd.GetBookmark(0))
  lbl(1).Caption = grd.Columns(3).CellValue(grd.GetBookmark(0))
  gblFileKey = grd.Columns(1).CellValue(grd.GetBookmark(0))
  gblOptions = 2
  frmSIS028.Show 0
  Me.Visible = False
End If
End Sub


Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If gblUserLevel <> gblViewOnly Then
  frmMDI.txtStatusMsg.SimpleText = "Double Click to Edit"
  frmMDI.txtStatusMsg.Refresh
End If
End Sub

Private Sub grd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If gblUserLevel <> gblViewOnly Then
  frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
  frmMDI.txtStatusMsg.Refresh
End If
End Sub


Private Sub Form_Activate()
' status ready
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault

UpdateScreen
End Sub

Private Sub Form_Load()
Dim qSQL As String
Dim i, iErr As Integer
Dim strTmp As String
' readymsg
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
csvCenterForm Me, gblMDIFORM
'--
'Set rsBrokers = New ADODB.Recordset
'--
On Error GoTo Form_Load_Err:
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
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

' disable edit buttons if viewonly user
'If gblUserLevel = gblViewOnly Then cmdEdit.Enabled = False
'cmdPrint.Enabled = True
   
Exit Sub
Form_Load_Err:
    MsgBox "SIS027/Load"
    csvLogError "SIS027/Load", Err.Number, Err.Description

End Sub


Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
If Isloaded("frmSIS028") Then
  X = ErrWarn.ListWarn()
  Cancel = -1
  frmSIS028.Show
  Exit Sub
Else
  Unload Me
End If
End Sub
Private Sub UpdateScreen()
Dim sRowinfo As String

Set rsBrokers = RunSP(SpCon, "usp_FindBroker", 1)

With rsBrokers
      grd.RemoveAll
      If Not .EOF Then
        Do While Not .EOF
          sRowinfo = !CliName & vbTab & !BROKERID & vbTab
          sRowinfo = sRowinfo & !SHRHELD & vbTab & !avail
          grd.AddItem sRowinfo
          '-- refresh routine used when activate event
          '-- triggered from sis029
          '-------------------------------------------
          If lbl(0).Caption = !CliName Then
             lbl(1).Caption = !avail
          End If
         '----
         .MoveNext
        Loop
      End If
End With
End Sub
