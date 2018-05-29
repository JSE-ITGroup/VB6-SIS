VERSION 5.00
Begin VB.Form frmSIS076T 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T & T Mandate Enquiry"
   ClientHeight    =   5550
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "SIS076T.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6930
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   300
      Left            =   5880
      TabIndex        =   0
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   1920
      TabIndex        =   30
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Method of Payment:"
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
      Index           =   11
      Left            =   0
      TabIndex        =   29
      Top             =   1320
      Width           =   1860
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   9480
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   9480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   1920
      TabIndex        =   28
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   1920
      TabIndex        =   27
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   1920
      TabIndex        =   26
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
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
      Left            =   0
      TabIndex        =   25
      Top             =   2040
      Width           =   1860
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   1920
      TabIndex        =   24
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1920
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   1920
      TabIndex        =   22
      Top             =   3000
      Width           =   3375
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
      Index           =   2
      Left            =   0
      TabIndex        =   21
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   1920
      TabIndex        =   20
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   19
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   18
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   17
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   16
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   15
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   14
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   13
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Top             =   360
      Width           =   1935
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   9480
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Bank Id:"
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
      Index           =   12
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   1620
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Branch Name:"
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
      Index           =   8
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   1620
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
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
      Left            =   480
      TabIndex        =   8
      Top             =   3960
      Width           =   1380
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Mandate Name:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1740
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6960
      Y1              =   240
      Y2              =   240
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
      TabIndex        =   5
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
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1740
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
      TabIndex        =   3
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
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account No:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1695
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
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS076T"
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
Set frmSIS076T = Nothing
frmSIS070T.Visible = True
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

Set rsMain = RunSP(SpCon, "usp_FindMandatesJCSD", 1, gblFileKey)

OpenErr = False
UpdateScreen
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS076T/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
   
End Sub
Private Sub UpdateScreen()
 With rsMain
    lb(0).Caption = frmSIS070T.grd.Columns(1).Text
    lb(1).Caption = frmSIS070T.grd.Columns(0).Text
    If Not .EOF Then
     If Not IsNothing(!MndAcnt) Then lb(2).Caption = !MndAcnt
     If Not IsNothing(!MndAcntNme) Then lb(3).Caption = !MndAcntNme
     Select Case !MNDMET
     Case "CHQ"
      lb(4).Caption = "Local Cheque"
     Case "LLC"
      lb(4).Caption = "Local Lodgment with Cheque"
     Case "FLL"
      lb(4).Caption = "Foreign Lodgment List"
     Case Else
     End Select
     If Not IsNothing(!MndName) Then lb(5).Caption = !MndName
     If Not IsNothing(!MndAddr1) Then lb(6).Caption = !MndAddr1
     If Not IsNothing(!MndAddr2) Then lb(7).Caption = !MndAddr2
     If Not IsNothing(!MndAddr3) Then lb(8).Caption = !MndAddr3
     If Not IsNothing(!MNDADDR4) Then lb(9).Caption = !MNDADDR4
     If Not IsNothing(!MNDADDR5) Then lb(10).Caption = !MNDADDR5
     If Not IsNothing(!BankId) Then
       lb(11).Caption = !BankId
       lb(12).Caption = !BnkName
       lb(13).Caption = !BNKADDR1
       lb(14).Caption = !BNKADDR2
       If Not IsNothing(!BNKADDR3) Then lb(15).Caption = !BNKADDR3
       If Not IsNothing(!BNKADDR4) Then lb(16).Caption = !BNKADDR4
       If Not IsNothing(!BNKADDR5) Then lb(17).Caption = !BNKADDR5
     End If
   End If
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub
