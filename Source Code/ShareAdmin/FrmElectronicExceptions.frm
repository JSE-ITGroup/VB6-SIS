VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmElectronicExceptions 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Electronic Payments Exceptions"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FrmElectronicExceptions.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CmdBegin 
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPFirstDate 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Enter Start Date"
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   51445763
      CurrentDate     =   40854
   End
   Begin MSComCtl2.DTPicker DTPLastDate 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Enter End Date"
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   51445763
      CurrentDate     =   40854
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "And"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Between"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label LblTitle 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Exceptions processed between the dates below:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmElectronicExceptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBegin_Click()
On Error GoTo Err_CmdBegin_Click
If DTPFirstDate > DTPLastDate Then
   MsgBox "The Last date should not be greater than the first date", vbOKOnly
   DTPFirstDate.SetFocus
   GoTo Exit_CmdBegin_Click
End If

gblOptions = 5
FrmReportView.Show 0

Exit_CmdBegin_Click:
Exit Sub

Err_CmdBegin_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on generating exceptions report"
Resume Exit_CmdBegin_Click
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
csvCenterForm Me, gblMDIFORM
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
frmMDI.txtStatusMsg.Refresh
DTPFirstDate = Date
DTPLastDate = Date
End Sub
