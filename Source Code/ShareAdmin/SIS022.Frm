VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS022 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pending Stock Issues List"
   ClientHeight    =   3555
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   7950
   Icon            =   "SIS022.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   7950
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
      Width           =   7776
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   5556
      Columns(0).Caption=   "Client Name"
      Columns(0).Name =   "Client Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   30
      Columns(1).Width=   2434
      Columns(1).Caption=   "Account No"
      Columns(1).Name =   "Account No"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   10
      Columns(2).Width=   2593
      Columns(2).Caption=   "Certificate No"
      Columns(2).Name =   "Certificate No"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   2
      Columns(2).FieldLen=   10
      Columns(3).Width=   2514
      Columns(3).Caption=   "Shares"
      Columns(3).Name =   "Shares"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   2
      Columns(3).FieldLen=   10
      _ExtentX        =   13705
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
Attribute VB_Name = "frmSIS022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ErrWarn As New cLstWarn
Dim X As Integer
'Dim cnn As New cADOAccess
Dim rsIssues As New ADODB.Recordset
Dim errLoop As Error
Dim errs1 As Error
Dim repSISRept As New SISRepts

Private Sub cmdClose_Click()
If Isloaded("frmSIS023") Then
  X = ErrWarn.ListWarn()
  frmSIS023.Show
  Exit Sub
Else
   rsIssues.Close
   Set rsIssues = Nothing
   cnn.Close
   '''set cnn = nothing
  Unload Me
End If
End Sub
Private Sub cmdEdit_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

gblFileKey = grd.Columns(1).Text & ";" & grd.Columns(2).Text
gblOptions = 2
cnn.Close
frmSIS023.Show 0
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
    .cbWhere.AddItem "ApplicationNo"
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
               sWhere = UCase(Trim(.tbFind.Text) & "*")
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
               sWhere = UCase("*" & Trim(.tbFind.Text))
               
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
               sWhere = UCase("*" & Trim(.tbFind.Text) & "*")
               
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
cnn.Close
frmSIS023.Show 0
End Sub



Private Sub cmdPrint_Click()
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'repSISRept.DSN = gblDSN
repSISRept.LoginId = gblFileName
repSISRept.ReportNumber = 10
repSISRept.ReportType = 9
repSISRept.siteid = gblSiteId
repSISRept.RunShareHolderReport
End Sub

Private Sub Form_GotFocus()
UpdateScreen
End Sub

Private Sub grd_Click()
gblFileKey = grd.Columns(1).CellValue(grd.GetBookmark(0))
gblFileKey = gblFileKey & ";" & grd.Columns(2).CellValue(grd.GetBookmark(0))
End Sub

Private Sub grd_DblClick()
' disable dblclick for viewonly users
If gblUserLevel <> gblViewOnly Then
  ' wait message & hourglass
  Screen.MousePointer = vbHourglass
  frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
  frmMDI.txtStatusMsg.Refresh

  gblFileKey = grd.Columns(1).CellValue(grd.GetBookmark(0))
  gblFileKey = gblFileKey & ";" & grd.Columns(2).CellValue(grd.GetBookmark(0))
  gblOptions = 2
  cnn.Close
  frmSIS023.Show 0
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
   Dim i As Integer
   Dim strTmp As String
   ' readymsg
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   csvCenterForm Me, gblMDIFORM
   '--
   Set rsIssues = New ADODB.Recordset
   '--
   On Error GoTo Form_Load_Err:
   cnn.Open
   qSQL = "SELECT a.CLINAME, b.CLIENTID,  b.CertNo, b.SHARES "
   qSQL = qSQL & "FROM STKNAME a, STKACTIV b "
   qSQL = qSQL & " where a.CLIENTID = b.CLIENTID "
   qSQL = qSQL & "and b.TRNCODE = 'I' "
   qSQL = qSQL & "and b.STATUS = 'O'"
   rsIssues.Open qSQL, cnn, , , adCmdText
   
  ' disable edit buttons if viewonly user
   If gblUserLevel = gblViewOnly Then
     cmdEdit.Enabled = False
     cmdNew.Enabled = False
   End If
Exit Sub
Form_Load_Err:
  MsgBox "SIS022/Load"
  csvLogError "SIS022/Load", Err.Number, Err.Description
  Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
If Isloaded("frmSIS023") Then
  X = ErrWarn.ListWarn()
  Cancel = -1
  frmSIS023.Show
  Exit Sub
Else
  Unload Me
End If
End Sub
Private Sub UpdateScreen()
Dim sRowinfo As String
Dim qSQL As String
If cnn.State = 0 Then
   cnn.Open
   qSQL = "SELECT a.CLINAME, b.CLIENTID,  b.CertNo, b.SHARES "
   qSQL = qSQL & "FROM STKNAME a, STKACTIV b "
   qSQL = qSQL & " where a.CLIENTID = b.CLIENTID "
   qSQL = qSQL & "and b.TRNCODE = 'I' "
   qSQL = qSQL & "and b.STATUS = 'O'"
   rsIssues.Open qSQL, cnn, , , adCmdText
Else
   rsIssues.Requery
End If

With rsIssues
      grd.RemoveAll
      If Not .EOF Then
        .MoveFirst
        Do While Not .EOF
          sRowinfo = !CliName & Chr(9) & !ClientID & Chr(9)
          sRowinfo = sRowinfo & !certno & Chr(9) & !shares
          grd.AddItem sRowinfo
         .MoveNext
        Loop
      End If
End With
End Sub
