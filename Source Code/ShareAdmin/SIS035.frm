VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS035 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Split Stock Certificates"
   ClientHeight    =   4440
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "SIS035.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7275
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enter the new details to split the displayed certificate."
      Top             =   2040
      Width           =   3615
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   2
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "CertNo"
      Columns(0).Name =   "CertNo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   12
      Columns(1).Width=   2566
      Columns(1).Caption=   "Stocks"
      Columns(1).Name =   "Stocks"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   12
      _ExtentX        =   6376
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Breakdown Information"
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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   4080
      TabIndex        =   11
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3000
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   6240
      TabIndex        =   2
      ToolTipText     =   "Cancels all processing and exits program."
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "Saves the screen information to the database."
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "0"
      Height          =   255
      Index           =   7
      Left            =   5640
      TabIndex        =   23
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "No of Stocks"
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   22
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Shares:"
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
      Index           =   10
      Left            =   4080
      TabIndex        =   21
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label lbl 
      Caption         =   "Date of Issue variable:"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   20
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date of Issue:"
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
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Holder #3:"
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
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "Joint Holder #3 Name:"
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   17
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cert Number:"
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
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   360
      Width           =   1740
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Shareholder Name:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   1740
   End
   Begin VB.Label lbl 
      Caption         =   "Cert Number"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   14
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Shareholder Name"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   13
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Stocks to be split:"
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
      Index           =   19
      Left            =   3360
      TabIndex        =   12
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label lbl 
      Caption         =   "Joint Holder #1 Name:"
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   10
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label lbl 
      Caption         =   "Joint Holder #2 Name:"
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   9
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Holder #2:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Holder #1:"
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
      Index           =   9
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   7320
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label lbllabels 
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
      Left            =   360
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   0
      X2              =   7320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lbllabels 
      Caption         =   "Ver:"
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
      Index           =   20
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lbllabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
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
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmSIS035"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim rsClient As ADODB.Recordset
Dim rsCert As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim OpenErr As Integer
Dim iOpenCert As Integer
Dim iOpenClient As Integer, iDelVal As Long
Dim iStocks As Double, icert As Long, icolval As Long
Dim sql As String, iErr As Long, iClient As Long

Function IsValid() As Integer
Dim sElable As String, i As Integer, X As Integer
sElable = "Split Stock Certificate"
IsValid = False
iErr = 0
'--
With grd
 .MoveFirst
 For i = 0 To .Rows - 1
   grd.SelBookmarks.Add grd.Bookmark
   '-- Validate certificate number
   If IsNothing(Trim(.Columns(0).CellText(.AddItemBookmark(i)))) Then
     iErr = 135
     GoTo Validate_Err
   End If
   '--
   Set rsCert = RunSP(SpCon, "usp_Sis035B", 1, Val(.Columns(0).CellText(.AddItemBookmark(i))))
   iOpenCert = True
   If Not rsCert.EOF Then 'we have a duplicate cert
      sql = "Duplicate Cetificate Number found"
      GoTo Validate_Err
   End If
   rsCert.Close
   iOpenCert = 0
   '-- validate no of shares
   If IsNothing(Trim(.Columns(1).CellText(.AddItemBookmark(i)))) Then
     sql = "Number of shares not validated"
     GoTo Validate_Err
   End If
   If Not IsNumeric(.Columns(1).CellText(.AddItemBookmark(i))) Then
     sql = "Value entered is not numeric"
     GoTo Validate_Err
   End If
   
   '--  update shares available to be split
   iStocks = iStocks - Val(.Columns(1).CellText(.AddItemBookmark(i)))
   .SelBookmarks.Remove 0
  Next i
 '--VALIDATE shares in grid
 If iStocks < 0 Then
    sql = "Stocks less than 0"
    GoTo Validate_Err
 End If
 End With
 '--
 IsValid = True
Validate_Exit:
  Exit Function
Validate_Err:
   MsgBox sql, vbOKOnly, sElable
   GoTo Validate_Exit
End Function

Private Sub cmdCancel_Click()
Shutdown
Unload Me
End Sub

Private Sub cmdClear_Click()
UpdateScreen
End Sub

Private Sub cmdDelete_Click()
grd.DeleteSelected

End Sub


Private Sub cmdUpdate_Click()
Dim i As Integer, X As Integer
Dim iLines As Integer
Dim StrSql As String
On Error GoTo cmdUpdate_Err
If Not IsValid Then Exit Sub
'-- add certs to certmst
iLines = 0
StrSql = ""
With grd
.MoveFirst
For i = 0 To .Rows - 1
  iLines = iLines + 1
  StrSql = StrSql & Val(.Columns(0).CellText(.AddItemBookmark(i))) & ";" & Val(.Columns(1).CellText(.AddItemBookmark(i))) & ";"
  'X = RunSP(SpCon, "usp_Sis035C", 0, Val(.Columns(0).CellText(.AddItemBookmark(i))), iClient, DateValue(lbl(2)), Val(.Columns(1).CellText(.AddItemBookmark(i))), lbl(0))
Next i
End With

X = RunSP(SpCon, "usp_Sis035C", 0, Val(lbl(0)), iClient, Format(lbl(2), "dd-mmm-yyyy"), iLines, iStocks, StrSql)
If X <> 0 Then
   MsgBox "The Split was aborted. Please re-try or contact your Sysad"
   GoTo Done
Else
   MsgBox "Certificate successfully split"
End If
'---
Done:
cmdCancel_Click
Exit Sub
'--
cmdUpdate_Err:
  MsgBox Err.Description, vbOKOnly, "SIS035/cmdUpdate"
  cmdCancel_Click
End Sub
Private Sub Form_Activate()

' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
'--
If OpenErr = True Then
   Shutdown
   Unload Me
   Exit Sub
End If
UpdateScreen
'--
End Sub
Private Sub Form_Load()
On Error GoTo FL_ERR
'-------------------------------------
'-- Initialize Company Details -------
'-------------------------------------
lblLabels(0).Caption = gblCompName
lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
csvCenterForm Me, gblMDIFORM
lbl(0).Caption = gblFileKey
lbl(1).Caption = frmSIS011.tbfld(1).Text
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

'--
OpenErr = 0: iOpenClient = 0
iOpenCert = 0
'-----------------------
'-- open tables --------
'-----------------------
Set rsClient = RunSP(SpCon, "usp_Sis035A", 1, gblFileKey)
iOpenClient = True
If rsClient.EOF Then
   iOpenClient = False
   rsClient.Close
   OpenErr = True
End If
'--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "Error on opeing certificate details", vbCritical + vbOKOnly, "SIS035/Load"
 OpenErr = True
 GoTo FL_Exit
   
End Sub
Private Sub UpdateScreen()
With rsClient
      grd.RemoveAll
      '--
      lbl(2).Caption = Format(!IssDate, "dd-mmm-yyyy")
      lbl(3).Caption = Format(!shares, "#,###")
      iStocks = !shares
      iClient = !ClientID
      If Not IsNothing(!JNTNAME1) Then
        lbl(4).Caption = !JNTNAME1
      Else
        lbl(4).Caption = ""
      End If
      If Not IsNothing(!JNTNAME2) Then
        lbl(5).Caption = !JNTNAME2
      Else
        lbl(5).Caption = ""
      End If
      If Not IsNothing(!jntname3) Then
        lbl(6).Caption = !jntname3
      Else
        lbl(6).Caption = ""
      End If
      '--
End With
icolval = 0: lbl(7).Caption = "0"
End Sub


Private Sub ClearScreen()
grd.RemoveAll
End Sub

Private Sub Shutdown()
If iOpenClient = True Then rsClient.Close
If iOpenCert = True Then rsCert.Close
Set rsClient = Nothing
Set rsCert = Nothing
Set frmSIS035 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub grd_AfterDelete(RtnDispErrMsg As Integer)
icolval = icolval - iDelVal
lbl(7) = Format(icolval, "#,###")
End Sub

Private Sub grd_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
Dim msg As String
Select Case ColIndex

Case 0 ' Certificate number
 If IsNothing(Trim(grd.Columns(ColIndex).Text)) Then
     msg = "A certificate number was not entered"
     GoTo Grd_BeforeColUpdate_err
 End If
 Set rsCert = RunSP(SpCon, "usp_Sis035B", 1, Val(grd.Columns(ColIndex).Text))
 iOpenCert = True
 If Not rsCert.EOF Then 'we have a duplicate cert
      msg = "Duplicate Certificate found"
      GoTo Grd_BeforeColUpdate_err
 End If
 rsCert.Close
 iOpenCert = 0

Case 1 ' shares
   If IsNothing(grd.Columns(ColIndex).Text) Then
     msg = "No shares were entered"
     GoTo Grd_BeforeColUpdate_err
   End If
  If Not IsNothing(grd.Columns(ColIndex).Text) Then
   If iStocks - Val(grd.Columns(ColIndex).Text) < 0 Then
     msg = "Shares not enough"
     GoTo Grd_BeforeColUpdate_err
   End If
   icolval = icolval + Val(grd.Columns(ColIndex).Text) _
             - Val(OldValue)
   lbl(7) = Format(icolval, "#,###")
   
 End If
Case Else
End Select
Cancel = 0
Grd_BeforeColUpdate_exit:
Exit Sub
Grd_BeforeColUpdate_err:
 MsgBox msg, vbOKOnly, "Split Certificates"
 Cancel = True
 GoTo Grd_BeforeColUpdate_exit
End Sub

Private Sub grd_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
iDelVal = Val(grd.Columns(1).Text)
End Sub
