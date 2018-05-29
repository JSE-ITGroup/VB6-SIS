VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReplaced 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Replaced Cheques Report Generator"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   ClipControls    =   0   'False
   Icon            =   "FrmReplaced.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton CmdGenerate 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton OptType 
      BackColor       =   &H00C0FFC0&
      Caption         =   "List of Replaced but not Authorised Cheques"
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
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   3615
   End
   Begin VB.OptionButton OptType 
      BackColor       =   &H00C0FFC0&
      Caption         =   "List of Replaced && Authorised Cheques"
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
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker DTPEndDate 
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20643843
      CurrentDate     =   38635
   End
   Begin MSComCtl2.DTPicker DTPStart 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20643843
      CurrentDate     =   38635
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
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
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select Reporting Dates:"
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
      Width           =   2415
   End
End
Attribute VB_Name = "FrmReplaced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdGenerate_Click()
On Error GoTo Err_CmdGenerate_Click

If DTPStart > DTPEndDate Then
   MsgBox "The Start Date selected is greater than the End Date"
   DTPEndDate.SetFocus
   GoTo Exit_CmdGenerate_Click
End If
If OptType(0).Value = True Then
   gblreply = 0
Else
   gblreply = 1
End If
gblDate = DTPStart
gblDate1 = DTPEndDate
gblOptions = 1
FrmReportView.Show 0

Exit_CmdGenerate_Click:
Exit Sub

Err_CmdGenerate_Click:
MsgBox Err & Err.Description, vbOKOnly, "Replaced Cheques Report"
GoTo Exit_CmdGenerate_Click
End Sub

Private Sub Form_Load()
DTPStart = Date
DTPEndDate = Date
OptType(0).Value = True

'PubOK = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmReplaced = Nothing
End Sub
