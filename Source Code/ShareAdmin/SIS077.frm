VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS077 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payments History Enquiry"
   ClientHeight    =   4860
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "SIS077.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7845
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   3732
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   7752
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   7
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   7
      Columns(0).Width=   2064
      Columns(0).Caption=   " Declaration"
      Columns(0).Name =   " Declaration"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   7
      Columns(0).NumberFormat=   "dd-mmm-yyyy"
      Columns(0).FieldLen=   11
      Columns(1).Width=   2646
      Columns(1).Caption=   "Type"
      Columns(1).Name =   "Type"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   8
      Columns(2).Width=   1826
      Columns(2).Caption=   "Record"
      Columns(2).Name =   "Record"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   7
      Columns(2).NumberFormat=   "dd-mmm-yyyy"
      Columns(2).FieldLen=   11
      Columns(3).Width=   1799
      Columns(3).Caption=   "Paid"
      Columns(3).Name =   "Paid"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   7
      Columns(3).NumberFormat=   "dd-mmm-yyyy"
      Columns(3).FieldLen=   11
      Columns(4).Width=   1376
      Columns(4).Caption=   "Rate"
      Columns(4).Name =   "Rate"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   4
      Columns(4).NumberFormat=   "##.###"
      Columns(4).FieldLen=   5
      Columns(5).Width=   1879
      Columns(5).Caption=   "Inc. Type"
      Columns(5).Name =   "Inc. Type"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   4
      Columns(6).Width=   1482
      Columns(6).Caption=   "Closed"
      Columns(6).Name =   "Closed"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   11
      Columns(6).FieldLen=   5
      _ExtentX        =   13674
      _ExtentY        =   6583
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
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   300
      Left            =   6720
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   7800
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
      TabIndex        =   2
      Top             =   0
      Width           =   855
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
      TabIndex        =   1
      Top             =   0
      Width           =   375
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
      TabIndex        =   3
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmSIS077"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As ADODB.Recordset
Dim OpenErr As Integer
Dim SpCon As ADODB.Connection

Private Sub cmdOk_Click()
rsMain.Close
Set rsMain = Nothing
Unload Me
Set frmSIS077 = Nothing
End Sub

Private Sub Form_Activate()
If OpenErr = True Then
  Unload Me
End If
End Sub

Private Sub Form_Load()
Dim qSQL As String
On Error GoTo FL_ERR
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
lblLabels(0).Caption = gblCompName
lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
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
Set rsMain = RunSP(SpCon, "usp_FindPayments", 1)
UpdateScreen
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS077/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit

End Sub
Private Sub UpdateScreen()
 Dim sRowinfo As String, sPayTYP As String, sIncTyp As String
 With rsMain
    If Not .EOF Then
      grd.RemoveAll
      Do While Not .EOF
        If !PAYTYP = "D" Then
           sPayTYP = "Dividend"
        Else
           sPayTYP = "Cap Dist"
        End If
        If !INCTYP = "F" Then
           sIncTyp = "FR"
        Else
           sIncTyp = "UNFR"
        End If
        sRowinfo = !DECDATE & Chr(9) & sPayTYP & Chr(9)
        sRowinfo = sRowinfo & !RECDATE & Chr(9) & !ChqDate
        sRowinfo = sRowinfo & Chr(9) & !PAYPER
        sRowinfo = sRowinfo & Chr(9) & sIncTyp
        sRowinfo = sRowinfo & Chr(9) & !CLOSED
        grd.AddItem sRowinfo
        .MoveNext
      Loop
    End If
    
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub
