VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS028 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outstanding Brokers' Certifications "
   ClientHeight    =   4365
   ClientLeft      =   3060
   ClientTop       =   2880
   ClientWidth     =   7530
   Icon            =   "SIS028.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4365
   ScaleWidth      =   7530
   Begin VB.CommandButton cmdTfr 
      Caption         =   "&Transfer"
      Height          =   328
      Left            =   2040
      TabIndex        =   10
      Top             =   3960
      Width           =   1005
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   328
      Left            =   3120
      TabIndex        =   4
      Top             =   3960
      Width           =   1005
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7305
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   5
      UseGroups       =   -1  'True
      RowHeight       =   423
      Groups.Count    =   2
      Groups(0).Width =   3545
      Groups(0).Caption=   "Batch Information"
      Groups(0).Columns.Count=   2
      Groups(0).Columns(0).Width=   1535
      Groups(0).Columns(0).Caption=   "Batch No"
      Groups(0).Columns(0).Name=   "Batch No"
      Groups(0).Columns(0).DataField=   "Column 0"
      Groups(0).Columns(0).DataType=   8
      Groups(0).Columns(0).FieldLen=   8
      Groups(0).Columns(1).Width=   2011
      Groups(0).Columns(1).Caption=   "Form No"
      Groups(0).Columns(1).Name=   "Form No"
      Groups(0).Columns(1).DataField=   "Column 1"
      Groups(0).Columns(1).DataType=   2
      Groups(0).Columns(1).FieldLen=   10
      Groups(1).Width =   8811
      Groups(1).Caption=   "Certified To"
      Groups(1).Columns.Count=   3
      Groups(1).Columns(0).Width=   4339
      Groups(1).Columns(0).Caption=   "To Broker"
      Groups(1).Columns(0).Name=   "To Broker"
      Groups(1).Columns(0).DataField=   "Column 2"
      Groups(1).Columns(0).DataType=   8
      Groups(1).Columns(0).FieldLen=   256
      Groups(1).Columns(1).Width=   2302
      Groups(1).Columns(1).Caption=   "Account"
      Groups(1).Columns(1).Name=   "Account"
      Groups(1).Columns(1).DataField=   "Column 3"
      Groups(1).Columns(1).DataType=   8
      Groups(1).Columns(1).FieldLen=   256
      Groups(1).Columns(2).Width=   2170
      Groups(1).Columns(2).Caption=   "Shares"
      Groups(1).Columns(2).Name=   "Shares"
      Groups(1).Columns(2).DataField=   "Column 4"
      Groups(1).Columns(2).DataType=   8
      Groups(1).Columns(2).FieldLen=   256
      _ExtentX        =   12885
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
      Left            =   6360
      TabIndex        =   2
      Top             =   3960
      Width           =   1005
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   328
      Left            =   4200
      TabIndex        =   1
      Top             =   3960
      Width           =   1005
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   328
      Left            =   5280
      TabIndex        =   0
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Label lbl 
      Caption         =   "SHRAVAILABLE"
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "Available Shares"
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
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Brokers Name"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lbl 
      Caption         =   "Brokers Account"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "Certifying Broker:"
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
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmSIS028"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ErrWarn As New cLstWarn
Dim X As Integer
Dim rsBrkCrt As New ADODB.Recordset
Dim errLoop As Error
Dim SpCon As ADODB.Connection
Dim errs1 As Error

Private Sub cmdClose_Click()
If Isloaded("frmSIS029") Then
  X = ErrWarn.ListWarn()
  frmSIS029.Show
  Exit Sub
Else
   rsBrkCrt.Close
   Set rsBrkCrt = Nothing
   frmSIS027.Visible = True
  Unload Me
End If
End Sub
Private Sub cmdEdit_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'-- store client id & form & to broker in filekey
gblFileKey = Trim(lbl(1).Caption) & ";" & Trim(grd.Columns(1).Text)
gblFileKey = gblFileKey & ";" & Trim(grd.Columns(3).Text)
gblOptions = 2
frmSIS029.Show 0
Me.Visible = False
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
    .cbWhere.AddItem "FormNo"
    .cbWhere.AddItem "ClientName"
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
      X = .cbWhere.ListIndex + 1
      '---
      If sWhere <> "" Then
          
          grd.Redraw = False
          If .Optbtn.OptionValue = 1 Then grd.MoveFirst
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
gblFileKey = Trim(lbl(1).Caption) & ";0;0"
frmSIS029.Show 0
frmSIS028.Visible = False
End Sub

Private Sub cmdTfr_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'-- store client id & form & to broker in filekey
gblFileKey = Trim(lbl(1).Caption) & ";" & Trim(grd.Columns(1).Text)
gblFileKey = gblFileKey & ";" & Trim(grd.Columns(3).Text)
gblOptions = 1
frmSIS030.Show 0
frmSIS028.Visible = False
End Sub

Private Sub Form_GotFocus()
UpdateScreen
End Sub
Private Sub grd_DblClick()
' disable dblclick for viewonly users
If gblUserLevel <> gblViewOnly Then
  ' wait message & hourglass
  Screen.MousePointer = vbHourglass
  frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
  frmMDI.txtStatusMsg.Refresh

  gblFileKey = Trim(lbl(1).Caption) & ";" & grd.Columns(1).Text
  gblFileKey = gblFileKey & ";" & grd.Columns(3).CellValue(grd.GetBookmark(0))
  gblOptions = 2
  frmSIS029.Show 0
  Me.Visible = False
End If
End Sub


Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If gblUserLevel <> gblViewOnly Then
  frmMDI.txtStatusMsg.SimpleText = "Double Click to Edit"
  frmMDI.txtStatusMsg.Refresh
End If
End Sub

Private Sub grd_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
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
   lbl(1).Caption = gblFileKey
   lbl(2).Caption = frmSIS027.lbl(0).Caption
   ' readymsg
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   csvCenterForm Me, gblMDIFORM
   '--
   Set rsBrkCrt = New ADODB.Recordset
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
   Loop
   Screen.MousePointer = vbDefault

   qSQL = "SELECT a.CLINAME, b.TOBROKERID, b.SHARES, "
   qSQL = qSQL & "b.BATCH, b.FORMNO "
   qSQL = qSQL & "FROM STKBKCRT b, STKNAME a "
   qSQL = qSQL & "where b.TOBROKERID = a.CLIENTID and "
   qSQL = qSQL & "b.BROKERID = " & gblFileKey
   qSQL = qSQL & " and b.STATUS = 'H' "
   qSQL = qSQL & "ORDER BY  a.CLINAME, b.BATCH, b.FORMNO"
   rsBrkCrt.Open qSQL, SpCon, , , adCmdText
   
  ' disable edit buttons if viewonly user
   If gblUserLevel = gblViewOnly Then
     cmdEdit.Enabled = False
     cmdNew.Enabled = False
   End If
Exit Sub
Form_Load_Err:
    MsgBox "SIS028/Load"
    'csvLogError "SIS028/Load", Err.Number, Err.Description

End Sub


Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
If Isloaded("frmSIS029") Then
  X = ErrWarn.ListWarn()
  Cancel = -1
  frmSIS029.Show
  Exit Sub
Else
  Unload Me
End If
End Sub
Private Sub UpdateScreen()
Dim sRowinfo As String
Dim qSQL As String
If rsBrkCrt.State = 0 Then
qSQL = "SELECT a.CLINAME, b.TOBROKERID, b.SHARES, "
qSQL = qSQL & "b.BATCH, b.FORMNO "
qSQL = qSQL & "FROM STKBKCRT b, STKNAME a "
qSQL = qSQL & "where b.TOBROKERID = a.CLIENTID and "
qSQL = qSQL & "b.BROKERID = " & gblFileKey
qSQL = qSQL & " and b.STATUS = 'H' "
qSQL = qSQL & "ORDER BY  a.CLINAME, b.BATCH, b.FORMNO"
rsBrkCrt.Open qSQL, SpCon, , , adCmdText
Else
rsBrkCrt.Requery
End If

With rsBrkCrt
      grd.RemoveAll
      If Not .EOF Then
        .MoveFirst
        Do While Not .EOF
          sRowinfo = !batch & Chr(9) & !formno & Chr(9)
          sRowinfo = sRowinfo & !CliName & Chr(9) & !TOBROKERID & Chr(9)
          sRowinfo = sRowinfo & !shares
          grd.AddItem sRowinfo
         .MoveNext
        Loop
        cmdTfr.Enabled = True
        cmdFind.Enabled = True
        cmdEdit.Enabled = True
      Else
        cmdTfr.Enabled = False
        cmdFind.Enabled = False
        cmdEdit.Enabled = False
      End If
End With
lbl(4).Caption = frmSIS027.lbl(1).Caption
End Sub
