VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS072 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Certificate  Enquiry"
   ClientHeight    =   4650
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "SIS072.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7020
   Begin MSMask.MaskEdBox meb 
      Height          =   252
      Left            =   4680
      TabIndex        =   9
      Top             =   3720
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   12
      Format          =   "###,###,##0"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   2292
      Left            =   600
      TabIndex        =   8
      Top             =   1320
      Width           =   5748
      ScrollBars      =   3
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   2461
      Columns(0).Caption=   "Cert No"
      Columns(0).Name =   "Cert No"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   12
      Columns(1).Width=   2170
      Columns(1).Caption=   "Issue Date"
      Columns(1).Name =   "Issue Date"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd-mmm-yyyy"
      Columns(1).FieldLen=   11
      Columns(2).Width=   1773
      Columns(2).Caption=   "Lien"
      Columns(2).Name =   "Lien"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   11
      Columns(2).FieldLen=   3
      Columns(3).Width=   2725
      Columns(3).Caption=   "Shares"
      Columns(3).Name =   "Shares"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   3
      Columns(3).NumberFormat=   "#,####"
      Columns(3).FieldLen=   10
      _ExtentX        =   10139
      _ExtentY        =   4043
      _StockProps     =   79
      Caption         =   "Active Certificates"
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
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   300
      Left            =   5760
      TabIndex        =   0
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Width           =   1935
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   9480
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   7080
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
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
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
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1260
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
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Client Number:"
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
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1500
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
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmSIS072"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim OpenErr As Integer

Private Sub cmdOk_Click()
rsMain.Close
Set rsMain = Nothing
Unload Me
Set frmSIS072 = Nothing
frmSIS070.Visible = True
End Sub

Private Sub Form_Activate()
If OpenErr = True Then
  Unload Me
End If
End Sub

Private Sub Form_Load()
Dim qSQL As String
''''On Error GoTo FL_ERR
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
lblLabels(0).Caption = gblCompName
lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
'--
csvCenterForm Me, gblMDIFORM
'-----------------------------------
Set rsMain = New ADODB.Recordset
'-------------------------------------------
' create SQL for selecting record to display
'-------------------------------------------
'''qSQL = "Select CERTNO, ISSDATE, ASSIGNED, SHARES from CERTMST "
'''qSQL = qSQL & "where CLIENTID = " & gblFileKey & " AND "
'''qSQL = qSQL & "STATUS <> 'C' order by certno"
'---
OpenErr = False
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

Set rsMain = RunSP(SpCon, "usp_FindActiveCerts", 1, gblFileKey)

UpdateScreen
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
FL_Exit:
  Exit Sub
'''FL_ERR:
'''  MsgBox "SIS072/Load"
'''  OpenErr = True
'''  On Error Resume Next
'''  Resume FL_Exit
   
End Sub
Private Sub UpdateScreen()
 Dim nStocks As Double, sRowinfo As String
 nStocks = 0
 lb(0).Caption = frmSIS070.grd.Columns(1).Text
 lb(1).Caption = frmSIS070.grd.Columns(0).Text
  With rsMain
    If Not .EOF Then
      grd.RemoveAll
      Do While Not .EOF
        sRowinfo = !certno & vbTab & !IssDate & vbTab
        sRowinfo = sRowinfo & !assigned & vbTab & !shares
        grd.AddItem sRowinfo
        nStocks = nStocks + !shares
        .MoveNext
      Loop
    End If
    meb = nStocks
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub
