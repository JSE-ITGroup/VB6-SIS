VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS031 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pending Stockholder to Broker Transfers"
   ClientHeight    =   3525
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   7860
   Icon            =   "SIS031.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   7860
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   328
      Left            =   2400
      TabIndex        =   5
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   328
      Left            =   3480
      TabIndex        =   4
      Top             =   3120
      Width           =   1005
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3012
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7680
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   5
      UseGroups       =   -1  'True
      RowHeight       =   423
      Groups.Count    =   2
      Groups(0).Width =   9022
      Groups(0).Caption=   "Selling Information"
      Groups(0).Columns.Count=   3
      Groups(0).Columns(0).Width=   5001
      Groups(0).Columns(0).Caption=   "Client Name"
      Groups(0).Columns(0).Name=   "Client Name"
      Groups(0).Columns(0).DataField=   "Column 0"
      Groups(0).Columns(0).DataType=   8
      Groups(0).Columns(0).FieldLen=   30
      Groups(0).Columns(1).Width=   2170
      Groups(0).Columns(1).Caption=   "Account No"
      Groups(0).Columns(1).Name=   "Account No"
      Groups(0).Columns(1).DataField=   "Column 1"
      Groups(0).Columns(1).DataType=   8
      Groups(0).Columns(1).FieldLen=   10
      Groups(0).Columns(2).Width=   1852
      Groups(0).Columns(2).Caption=   "Shares"
      Groups(0).Columns(2).Name=   "Shares"
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   2
      Groups(0).Columns(2).FieldLen=   10
      Groups(1).Width =   4075
      Groups(1).Caption=   "Batch Information"
      Groups(1).Columns.Count=   2
      Groups(1).Columns(0).Width=   2223
      Groups(1).Columns(0).Caption=   "Batch No"
      Groups(1).Columns(0).Name=   "Batch No"
      Groups(1).Columns(0).DataField=   "Column 3"
      Groups(1).Columns(0).DataType=   8
      Groups(1).Columns(0).FieldLen=   8
      Groups(1).Columns(1).Width=   1852
      Groups(1).Columns(1).Caption=   "Form No"
      Groups(1).Columns(1).Name=   "Form No"
      Groups(1).Columns(1).DataField=   "Column 4"
      Groups(1).Columns(1).DataType=   8
      Groups(1).Columns(1).FieldLen=   8
      _ExtentX        =   13547
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
      TabIndex        =   2
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   328
      Left            =   4560
      TabIndex        =   1
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   328
      Left            =   5640
      TabIndex        =   0
      Top             =   3120
      Width           =   1005
   End
End
Attribute VB_Name = "frmSIS031"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ErrWarn As New cLstWarn
Dim X As Integer
Dim SpCon As ADODB.Connection
Dim rsIssues As New ADODB.Recordset
Dim errLoop As Error
Dim errs1 As Error
Dim repSISRept As New SISRepts

Private Sub cmdClose_Click()
If Isloaded("frmSIS032") Then
  X = ErrWarn.ListWarn()
  frmSIS032.Show
  Exit Sub
Else
   rsIssues.Close
   Set rsIssues = Nothing
  Unload Me
End If
End Sub
Private Sub cmdEdit_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'-- store client id & form in filekey
gblFileKey = grd.Columns(3).Text & ";" & grd.Columns(4).Text
gblOptions = 2
frmSIS032.Show 0
End Sub

Private Sub cmdFind_Click()
Dim i As Integer, X As Integer
Dim bm As Variant, cChk As Integer
Dim sWhere As String, sGrid As String
Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
     .cbWhere.Clear
    .cbWhere.AddItem "ClientName"
    .cbWhere.AddItem "AccountNo"
    .cbWhere.AddItem "FormNo"
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
      If X = 2 Then X = 4 ' change to access correct grd.column
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
Private Sub cmdNew_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

gblOptions = 1
gblFileKey = "0"
frmSIS032.Show 0
End Sub



Private Sub cmdPrint_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'repSISRept.DSN = gblDSN
repSISRept.LoginId = gblFileName
repSISRept.ReportNumber = 12
repSISRept.ReportType = 9
repSISRept.siteid = gblSiteId
repSISRept.RunShareHolderReport
Screen.MousePointer = vbDefault
End Sub
Private Sub Form_GotFocus()
UpdateScreen
End Sub


Private Sub grd_Click()
gblFileKey = grd.Columns(3).CellValue(grd.GetBookmark(0))
End Sub

Private Sub grd_DblClick()
' disable dblclick for viewonly users
If gblUserLevel <> gblViewOnly Then
  ' wait message & hourglass
  Screen.MousePointer = vbHourglass
  frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
  frmMDI.txtStatusMsg.Refresh

  gblFileKey = grd.Columns(3).CellValue(grd.GetBookmark(0))
  gblFileKey = gblFileKey & ";" & grd.Columns(4).CellValue(grd.GetBookmark(0))
  gblOptions = 2
  frmSIS032.Show 0
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
   Set rsIssues = New ADODB.Recordset
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
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

' disable edit buttons if viewonly user
   If gblUserLevel = gblViewOnly Then
     cmdEdit.Enabled = False
     cmdNew.Enabled = False
   End If
Exit Sub
Form_Load_Err:
    MsgBox "SIS031/Load"
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
If Isloaded("frmSIS032") Then
  X = ErrWarn.ListWarn()
  Cancel = -1
  frmSIS032.Show
  Exit Sub
Else
  SpCon.Close
  Unload Me
End If
End Sub
Private Sub UpdateScreen()
Dim sRowinfo As String
Dim qSQL As String

Set rsIssues = RunSP(SpCon, "usp_PendingS2B", 1)
With rsIssues
      grd.RemoveAll
      If Not .EOF Then
         Do While Not .EOF
          sRowinfo = !CliName & vbTab & !ClientID & vbTab
          sRowinfo = sRowinfo & !sumofshares & vbTab
          sRowinfo = sRowinfo & !TrnBatch & vbTab & !Form
          grd.AddItem sRowinfo
         .MoveNext
        Loop
      End If
End With
End Sub
