VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmExchangeRateList 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exchange Rate List"
   ClientHeight    =   3555
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   7125
   Icon            =   "FrmExchangeRateList.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   7125
   Begin VB.CommandButton cmdPrint 
      Caption         =   "E&xport to Excel"
      Height          =   328
      Left            =   840
      TabIndex        =   5
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   328
      Left            =   2160
      TabIndex        =   4
      Top             =   3120
      Width           =   889
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6960
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   2619
      Columns(0).Caption=   "Currency Code"
      Columns(0).Name =   "Account No"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   30
      Columns(1).Width=   3200
      Columns(1).Caption=   "Payment Type"
      Columns(1).Name =   "Payment Type"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Payment Date"
      Columns(2).Name =   "Payment Date"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Exchange Rate"
      Columns(3).Name =   "Exchange Rate"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   12277
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
      Left            =   5040
      TabIndex        =   2
      Top             =   3120
      Width           =   888
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   328
      Left            =   3120
      TabIndex        =   1
      Top             =   3120
      Width           =   889
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   328
      Left            =   4080
      TabIndex        =   0
      Top             =   3120
      Width           =   888
   End
End
Attribute VB_Name = "frmExchangeRateList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim SpCon As ADODB.Connection
Dim rsMain As New ADODB.Recordset
Private Sub cmdClose_Click()
If Isloaded("FrmCurrencyList") Then
  frmCurrencyList.Show
  Exit Sub
Else
   rsMain.Close
   Set rsMain = Nothing
  Unload Me
End If
End Sub
Private Sub cmdEdit_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

If grd.Columns(0).Text = "" Then
   Screen.MousePointer = vbDefault
   GoTo Exit_CmdExit_Click
End If

gblFileKey = grd.Columns(0).Text
gblOptions = 2
FrmCurrencyRate.Show vbModal
UpdateScreen
Exit_CmdExit_Click:
Exit Sub

End Sub

Private Sub cmdFind_Click()
Dim i As Integer, X As Integer
Dim bm As Variant, cChk As Integer
Dim sWhere As String, sGrid As String
If grd.Columns(0).Text = "" Then
   Screen.MousePointer = vbDefault
   GoTo Exit_CmdFind_Click
End If

Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
     .cbWhere.Clear
    .cbWhere.AddItem "Currency Code"
    .cbWhere.AddItem "Payment Type"
    .cbWhere.ListIndex = 0
    .cbOptions.ListIndex = 0
    .Show vbModal
    '----------------------------
    '-------- main line ---------
    '----------------------------
    If .tbFind.Text = vbNullString Then
    Else
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      sWhere = UCase((.tbFind.Text))
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
               sWhere = UCase("*" & UCase(Trim(.tbFind.Text)) & "*")
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

Exit_CmdFind_Click:
Exit Sub

End Sub

Private Sub cmdNew_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

gblOptions = 1
gblFileKey = "0"
FrmCurrencyRate.Show vbModal
UpdateScreen
End Sub

Private Sub cmdPrint_Click()
If grd.Columns(0).Text = "" Then
   Screen.MousePointer = vbDefault
   GoTo Exit_cmdPrint_Click
End If

Call ExportToExcel(rsMain)
Exit_cmdPrint_Click:
Exit Sub

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

  gblFileKey = grd.Columns(1).CellValue(grd.GetBookmark(0))
  gblOptions = 2
  FrmCurrencyMaint.Show 0
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
   On Error GoTo Form_Load_Err:
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
  
  ' disable edit buttons if not admin user
   If gblUserLevel = 1 Then
     cmdEdit.Enabled = True
     cmdNew.Enabled = True
   Else
      cmdEdit.Enabled = False
     cmdNew.Enabled = False
   End If
FL_Exit:
Exit Sub
Form_Load_Err:
  MsgBox Err & " " & Err.Description, vbOKOnly, "Currency List Load error"
  On Error Resume Next
  Resume FL_Exit
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Isloaded("FrmExchangeRates") Then
  Cancel = -1
  'FrmExchangeRates.Show
  Exit Sub
Else
  SpCon.Close
  Unload Me
End If
End Sub
Private Sub UpdateScreen()
Dim sRowinfo As String
Dim qSQL As String

Set rsMain = RunSP(SpCon, "usp_ListExchangeRates", 1)

With rsMain
      grd.RemoveAll
      If Not .EOF Then
        Do While Not .EOF
          sRowinfo = !CurrencyCode & vbTab & !PayType & vbTab
          sRowinfo = sRowinfo & Format(!ChqDate, "dd-mmm-yyyy") & vbTab & Format(!ExchRate, "#,##0.00####")
          grd.AddItem sRowinfo
         .MoveNext
        Loop
      End If
End With
End Sub

