VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS007 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stockholder Category "
   ClientHeight    =   3510
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   9030
   Icon            =   "SIS007.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3510
   ScaleWidth      =   9030
   Begin VB.CommandButton CmdExport 
      Caption         =   "Export to Excel"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   328
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   889
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   328
      Left            =   1200
      TabIndex        =   4
      Top             =   3120
      Width           =   889
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8850
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   5
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   5450
      Columns(0).Caption=   "Category Name"
      Columns(0).Name =   "Category Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   30
      Columns(1).Width=   1270
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   2
      Columns(2).Width=   1508
      Columns(2).Caption=   "Taxable"
      Columns(2).Name =   "Taxable"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   3
      Columns(3).Width=   3200
      Columns(3).Caption=   "Hold Dividend"
      Columns(3).Name =   "Hold Dividend"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "Dividend Currency"
      Columns(4).Name =   "Dividend Currency"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      _ExtentX        =   15610
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
      Left            =   4080
      TabIndex        =   2
      Top             =   3120
      Width           =   888
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   328
      Left            =   2160
      TabIndex        =   1
      Top             =   3120
      Width           =   889
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   328
      Left            =   3120
      TabIndex        =   0
      Top             =   3120
      Width           =   888
   End
End
Attribute VB_Name = "frmSIS007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ErrWarn As New cLstWarn
Dim X As Integer
Dim rsMain As New ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim OpenErr As Integer
Private Sub cmdClose_Click()
If Isloaded("frmSIS008") Then
  X = ErrWarn.ListWarn()
  frmSIS008.Show
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

gblFileKey = grd.Columns(1).Text
gblOptions = 2
frmSIS008.Show 0
End Sub

Private Sub CmdExport_Click()
Call ExportToExcel(rsMain)
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
    .cbWhere.AddItem "CatDesc"
    .cbWhere.AddItem "CatCode"
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
End Sub

Private Sub cmdNew_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh

gblOptions = 1
gblFileKey = "0"
frmSIS008.Show 0
End Sub

Private Sub cmdPrint_Click()
frmSIS050.Show
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
  frmSIS008.Show 0
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
If OpenErr = True Then
  Unload Me
Else
  UpdateScreen
End If
End Sub

Private Sub Form_Load()
   Dim i As Integer
   ' readymsg
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   csvCenterForm Me, gblMDIFORM
   '--
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
   Set rsMain = New ADODB.Recordset
   '--
   On Error GoTo Form_Load_Err:
  ' disable edit buttons if viewonly user
   If gblUserLevel = gblViewOnly Then
     cmdEdit.Enabled = False
     cmdNew.Enabled = False
   End If
FL_Exit:
Exit Sub
Form_Load_Err:
  MsgBox "SIS007/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit

End Sub


Private Sub Form_Unload(Cancel As Integer)
If Isloaded("frmSIS008") Then
  X = ErrWarn.ListWarn()
  Cancel = -1
  frmSIS008.Show
  Exit Sub
Else
  Unload Me
End If
End Sub

Private Sub UpdateScreen()
Dim sRowinfo As String
Dim qSQL As String

Set rsMain = RunSP(SpCon, "usp_SIS007", 1)
'''rsMain.Open QsQl, Spcon, , , adCmdText
'''Else
'''   rsMain.Requery
'''End If

With rsMain
      grd.RemoveAll
        Do While Not .EOF
        If !DivCurrency = "J" Then
           qSQL = "JMD"
        Else
           qSQL = "USD"
        End If
          sRowinfo = !catdesc & vbTab & !CatCode & vbTab & !cattax & vbTab
          sRowinfo = sRowinfo & !Hold & vbTab & qSQL
          grd.AddItem sRowinfo
          .MoveNext
        Loop
End With
End Sub


