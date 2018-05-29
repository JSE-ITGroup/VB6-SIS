VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS011 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate Listing"
   ClientHeight    =   3990
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "SIS011.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6765
   Begin VB.CommandButton cmdSplit 
      Caption         =   "&Split"
      Height          =   300
      Left            =   2400
      TabIndex        =   11
      ToolTipText     =   "Split the certificate into multiple certificates."
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdRevoke 
      Caption         =   "&Revoke"
      Height          =   300
      Left            =   3480
      TabIndex        =   10
      ToolTipText     =   "Revoke a previous assignment"
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5640
      TabIndex        =   3
      ToolTipText     =   "Cancels changes and returns to Account maintenance."
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdAssign 
      Caption         =   "&Assign"
      Height          =   300
      Left            =   4560
      TabIndex        =   2
      ToolTipText     =   "tag the certificate as assigned"
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox tbfld 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   0
      ToolTipText     =   "Use generate number or enter your own unique client Number"
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox tbfld 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Enter Address line 2"
      Top             =   1080
      Width           =   4335
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   6345
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   2858
      Columns(0).Caption=   "Certificate No"
      Columns(0).Name =   "Certificate Number"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   10
      Columns(1).Width=   3043
      Columns(1).Caption=   "Shares"
      Columns(1).Name =   "Shares"
      Columns(1).Alignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   10
      Columns(2).Width=   2461
      Columns(2).Caption=   "Date Issued"
      Columns(2).Name =   "Date Issued"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   7
      Columns(2).FieldLen=   11
      Columns(3).Width=   2328
      Columns(3).Caption=   "Assigned?"
      Columns(3).Name =   "Assigned?"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   5
      _ExtentX        =   11192
      _ExtentY        =   3201
      _StockProps     =   79
      BackColor       =   12632256
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
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblLabels 
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
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Name:"
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
      Index           =   16
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label lblLabels 
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
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Stockholder No:"
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
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1620
   End
   Begin VB.Label lblLabels 
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
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCert As ADODB.Recordset
Dim OpenErr As Integer
Dim SpCon As ADODB.Connection

Private Sub cmdAssign_Click()
If grd.Columns(3).Text = True Then ' already assigned
    cmdAssign.SetFocus  ' do nothing
    MsgBox "Certificate " & grd.Columns(0).Text & " already assigned"
    Exit Sub
End If
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'--
gblOptions = 1
gblFileKey = grd.Columns(0).Text
frmSIS012.Show 0
End Sub

Private Sub cmdCancel_Click()
  Shutdown
  frmSIS001.Show
  Unload Me
End Sub

Private Sub cmdRevoke_Click()
If grd.Columns(3).Text = False Then ' cannot revoke
    cmdRevoke.SetFocus  ' do nothing
    MsgBox "Certificate " & grd.Columns(0).Text & " not assigned. Therefore cannot be revoked"
    Exit Sub
End If
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'--
gblOptions = 2
gblFileKey = grd.Columns(0).Text
frmSIS012.Show 0
End Sub

Private Sub cmdSplit_Click()
' wait message & hourglass
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
frmMDI.txtStatusMsg.Refresh
'--
gblFileKey = grd.Columns(0).Text
frmSIS035.Show 0
End Sub

Private Sub Form_Activate()
' ready message
'---
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 '--
If OpenErr = True Then
  Unload Me
Else
  UpdateScreen
End If
'--

End Sub

Private Sub Form_Load()
Dim qSQL As String
On Error GoTo FL_ERR
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
 tbfld(0).BackColor = &HC0C0C0
 tbfld(1).BackColor = &HC0C0C0
 tbfld(0).Text = gblFileKey
 tbfld(1).Text = frmSIS001.grd.Columns(0).Text
'--
csvCenterForm Me, gblMDIFORM
'-----------------------------------
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

OpenErr = False
Set rsCert = New ADODB.Recordset
'----------------------------
'---- open recordsets -----
'-- create SQL for selecting record to edit
'----------------------------------------
'---
'rsCert.Open qSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
'--------------------

FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS011/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit

   
End Sub
Private Sub UpdateScreen()
Dim rowinfo As String
gblFileKey = tbfld(0)
Set rsCert = RunSP(SpCon, "usp_CertMstFindA", 1, gblFileKey)
grd.RemoveAll
With rsCert
    If Not .EOF Then
       Do While Not .EOF
         rowinfo = !certno & vbTab & !shares _
                 & vbTab & !IssDate & vbTab _
                 & !assigned
         grd.AddItem rowinfo
        .MoveNext
      Loop
      cmdAssign.Enabled = True
      cmdRevoke.Enabled = True
    Else
      cmdAssign.Enabled = False
      cmdRevoke.Enabled = False
      cmdSplit.Enabled = False
    End If
End With
End Sub
Private Sub Shutdown()
rsCert.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set rsCert = Nothing
SpCon.Close
Set frmSIS011 = Nothing
End Sub
