VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDates 
   BackColor       =   &H0080C0FF&
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FrmDates.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   16580611
      CurrentDate     =   40859
   End
   Begin VB.TextBox TxtMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Select Dates to report"
      Top             =   120
      Width           =   4455
   End
   Begin MSComCtl2.DTPicker DTPEndDate 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   16580611
      CurrentDate     =   40859
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "FROM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "FrmDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReportType As Integer

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdStart_Click()
On Error GoTo Err_CmdStart_Click

If DTPStartDate > DTPEndDate Then
   MsgBox "Start date cannot be greater than end date"
   DTPStartDate.SetFocus
   GoTo Exit_CmdStart_Click
End If

gblDate = DTPStartDate
gblDate1 = DTPEndDate
If gblFileKey <> "4" Then
   FrmReportView.Show 0
End If
CmdExit_Click
Exit_CmdStart_Click:
Exit Sub

Err_CmdStart_Click:
MsgBox Err.Description, vbOKOnly, "Date selection error"
GoTo Exit_CmdStart_Click
End Sub

Private Sub Form_Load()
csvCenterForm Me, gblMDIFORM
DTPStartDate = Date
DTPEndDate = Date
Select Case gblFileKey
   Case "0"
        Me.Caption = "Replacement Cheques Report"
        gblOptions = 7
   Case "1"
        Me.Caption = "Returned Cheques Report"
        gblOptions = 8
   Case "2"
        Me.Caption = "Finacle Exception updates Report"
        gblOptions = 5
   Case "3"
        Me.Caption = "ACH Exception updates Report"
        gblOptions = 6
   Case "4"
        Me.Caption = "Shareholders As At Report"
        DTPEndDate.Visible = False
        Label1.Caption = "As At"
        Label2.Visible = False
   Case "5"
        Me.Caption = "Unclaimed Balances Report"
        gblOptions = 12
        DTPEndDate.Visible = False
        Label2.Visible = False
        Label1.Caption = "As At"
   Case "6"
        Me.Caption = "Unclaimed Balances Full Report"
        gblOptions = 13
        DTPEndDate.Visible = False
        Label2.Visible = False
        Label1.Caption = "As At"
End Select



End Sub
