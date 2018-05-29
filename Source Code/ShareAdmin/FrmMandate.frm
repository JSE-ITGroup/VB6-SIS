VERSION 5.00
Begin VB.Form FrmMandate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search by Bank Account"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   Icon            =   "FrmMandate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8370
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "&Next"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "E&dit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   15
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox TxtAccount 
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   13
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox TxtEndDate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox TxtMndName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2280
      Width           =   7095
   End
   Begin VB.TextBox TxtAccnt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   6015
   End
   Begin VB.TextBox TxtStartDate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox TxtCliName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   6375
   End
   Begin VB.TextBox TxtClientID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "An End Date indicates that this mandate is no longer active"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Index           =   1
      X1              =   5280
      X2              =   5280
      Y1              =   3600
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Index           =   0
      X1              =   3240
      X2              =   3240
      Y1              =   3600
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   8400
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label6 
      Caption         =   "End Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Bank:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Account No && Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Start Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   8400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Client Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "ClientID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FrmMandate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim adoRst As ADODB.Recordset

Private Sub cmdEdit_Click()
gblFileKey = TxtClientID
frmSIS010.Show 0
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
 On Error GoTo Err_CmdFind_Click

CmdNext.Enabled = False
CmdPrevious.Enabled = False

If IsNumeric(TxtAccount) Then
   Set adoRst = RunSP(SpCon, "usp_Mandate", 1, TxtAccount)
Else
   MsgBox "Account number should be numeric"
   GoTo Exit_CmdFind_Click
End If
If adoRst.EOF Then
   MsgBox "Sorry, No records matching your criteria was found"
   GoTo Exit_CmdFind_Click
End If
If adoRst.RecordCount > 1 Then
   CmdNext.Enabled = True
End If

LoadData

Exit_CmdFind_Click:
Exit Sub

Err_CmdFind_Click:
MsgBox Err.Description, vbOKOnly, "FrmMandate:Find Mandate"
GoTo Exit_CmdFind_Click

End Sub

Private Sub CmdNext_Click()
On Error GoTo Err_CmdNext_Click

CmdPrevious.Enabled = True
adoRst.MoveNext
If adoRst.AbsolutePosition = adoRst.RecordCount Then
   CmdNext.Enabled = False
End If

LoadData

Exit_CmdNext_Click:
Exit Sub

Err_CmdNext_Click:
MsgBox Err.Description, vbOKOnly, "Next Manadate Button"
GoTo Exit_CmdNext_Click

End Sub

Private Sub CmdPrevious_Click()
On Error GoTo Err_CmdPrevious_Click

CmdNext.Enabled = True
adoRst.MovePrevious
If adoRst.AbsolutePosition = 1 Then
   CmdPrevious.Enabled = False
End If

LoadData

Exit_CmdPrevious_Click:
Exit Sub

Err_CmdPrevious_Click:
MsgBox Err.Description, vbOKOnly, "Mandates Previous Button"
GoTo Exit_CmdPrevious_Click
End Sub

Private Sub Form_Activate()
TxtAccount.SetFocus
End Sub

Private Sub Form_Load()
csvCenterForm Me, gblMDIFORM
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
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

End Sub
Public Sub LoadData()
TxtClientID = adoRst(0)
TxtCliName = adoRst(1)
TxtStartDate = Format(adoRst(2), "dd-mmm-yyyy")
TxtAccnt = adoRst(3) & " - " & adoRst(4)
TxtMndName = IsNullMove(adoRst(5))
If Not IsNull(adoRst(6)) Then
   TxtEndDate = Format(adoRst(6), "dd-mmm-yyyy")
Else
   TxtEndDate = ""
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set adoRst = Nothing
SpCon.Close
End Sub
