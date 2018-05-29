VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmChqMovement 
   Caption         =   "Cheque Movement Report"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13695
   Icon            =   "FrmChqMovement.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtTransfersIN 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   9720
      TabIndex        =   24
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame FmeStatistics 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   7800
      TabIndex        =   16
      Top             =   3600
      Width           =   5775
      Begin VB.TextBox TxtClosingBalance 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   28
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox TxtUsed 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   1920
         TabIndex        =   27
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TxtCancelled 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   1920
         TabIndex        =   26
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox TxtTransfersOUT 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox TxtOpeningBalance 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   5760
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Closing Balance:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Used:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cancelled:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Transfers OUT:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Transfers IN:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Opening Balance:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame FmeCommands 
      BackColor       =   &H00404080&
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   7680
      TabIndex        =   11
      Top             =   6480
      Width           =   5895
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print Report"
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
         Height          =   615
         Left            =   2040
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Left            =   4320
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdStart 
         Caption         =   "Start"
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
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame FmeReportCriteria 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   7695
      Begin MSComCtl2.DTPicker DTPFromDate 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   51838979
         CurrentDate     =   40723
      End
      Begin VB.OptionButton OptReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select Date Ranges"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1815
      End
      Begin VB.OptionButton OptReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select All Dates"
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
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBAccount 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   5535
         DataFieldList   =   "Column 0"
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   503
         Columns.Count   =   2
         Columns(0).Width=   5741
         Columns(0).Caption=   "Account Number"
         Columns(0).Name =   "Account Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "Currency"
         Columns(1).Name =   "Currency"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   9763
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Column 0"
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBLocations 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   840
         Width           =   5535
         DataFieldList   =   "Column 0"
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   503
         Columns.Count   =   2
         Columns(0).Width=   6429
         Columns(0).Caption=   "Location"
         Columns(0).Name =   "Account Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Currency"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   9763
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Column 0"
      End
      Begin MSComCtl2.DTPicker DTPToDate 
         Height          =   375
         Left            =   4680
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   51838979
         CurrentDate     =   40723
      End
      Begin VB.Label LblTo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TO:"
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
         Left            =   4320
         TabIndex        =   9
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label LblFrom 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FROM:"
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
         Left            =   1920
         TabIndex        =   8
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select Account:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select Location:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBMovements 
      Height          =   3495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13575
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   10
      BackColorEven   =   16761024
      BackColorOdd    =   16761087
      RowHeight       =   423
      Columns.Count   =   10
      Columns(0).Width=   3200
      Columns(0).Caption=   "Account No"
      Columns(0).Name =   "Account No"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Location"
      Columns(1).Name =   "Location"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2064
      Columns(2).Caption=   "Trans Date"
      Columns(2).Name =   "Trans Date"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).NumberFormat=   "dd-mmm-yyyy"
      Columns(2).FieldLen=   256
      Columns(3).Width=   2223
      Columns(3).Caption=   "Stating Chq No"
      Columns(3).Name =   "Stating Chq No"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2328
      Columns(4).Caption=   "Ending Chq No"
      Columns(4).Name =   "Ending Chq No"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "No of Chqs in Range"
      Columns(5).Name =   "No of Chqs in Range"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "Status"
      Columns(6).Name =   "Status"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Caption=   "Comments"
      Columns(7).Name =   "Comments"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "UserID"
      Columns(8).Name =   "UserID"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "Post Date"
      Columns(9).Name =   "Post Date"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   7
      Columns(9).NumberFormat=   "dd-mmm-yyyy"
      Columns(9).FieldLen=   256
      _ExtentX        =   23945
      _ExtentY        =   6165
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin SSDataWidgets_B.SSDBGrid SSDBActiveRanges 
      Height          =   1695
      Left            =   0
      TabIndex        =   29
      Top             =   3600
      Width           =   7575
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   8
      BackColorEven   =   12640511
      BackColorOdd    =   8438015
      RowHeight       =   423
      Columns.Count   =   8
      Columns(0).Width=   2064
      Columns(0).Caption=   "Trans Date"
      Columns(0).Name =   "Trans Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).NumberFormat=   "dd-mmm-yyyy"
      Columns(0).FieldLen=   256
      Columns(1).Width=   2223
      Columns(1).Caption=   "Stating Chq No"
      Columns(1).Name =   "Stating Chq No"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2328
      Columns(2).Caption=   "Ending Chq No"
      Columns(2).Name =   "Ending Chq No"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "No of Chqs in Range"
      Columns(3).Name =   "No of Chqs in Range"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "UserID"
      Columns(4).Name =   "UserID"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "Post Date"
      Columns(5).Name =   "Post Date"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   7
      Columns(5).NumberFormat=   "dd-mmm-yyyy"
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "Account No"
      Columns(6).Name =   "Account No"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Caption=   "Location"
      Columns(7).Name =   "Location"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      _ExtentX        =   13361
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "List of currently active cheque numbers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
End
Attribute VB_Name = "FrmChqMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
SpCon.Close
Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Err_CmdPrint_Click
gblOptions = 3
gblDate = DTPFromDate
gblDate1 = DTPToDate
gblHold = SSDBAccount.Columns(0).Text
gblFileKey = SSDBLocations.Columns(1).Text
gblDSN = SSDBLocations.Columns(0).Text
If OptReport(0).Value = True Then
   gblReply = 1
Else
   gblReply = 0
End If

FrmReportView.Show 0
Exit_cmdPrint_Click:
Exit Sub

Err_CmdPrint_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on generating Cheque Inventory Report"
Resume Exit_cmdPrint_Click
End Sub

Private Sub cmdPrint_LostFocus()
cmdPrint.Enabled = False
End Sub

Private Sub CmdStart_Click()
'On Error GoTo Err_CmdStart_Click

Dim adoRst As ADODB.Recordset
Dim adoRst1 As ADODB.Recordset
Dim adoRst2 As ADODB.Recordset
Dim AdoRst3 As ADODB.Recordset
Dim StrSql As String
Dim i As Integer
Dim iComment As String

If IsValid Then
   If OptReport(0).Value = True Then
      i = 0
   Else
      i = 1
   End If
   cmdPrint.Enabled = True
   Set adoRst = RunSP(SpCon, "usp_ListChqMovements", 1, SSDBAccount.Columns(0).Text, SSDBLocations.Columns(1).Text, i, Format(DTPFromDate, "dd-mmm-yyyy"), Format(DTPToDate, "dd-mmm-yyyy"))
   If adoRst.EOF Then
      MsgBox "Sorry, no records matching your criteria were found"
      GoTo Exit_CmdStart_Click
   End If
   Set adoRst1 = adoRst.NextRecordset
   Set adoRst2 = adoRst.NextRecordset
   Set AdoRst3 = adoRst.NextRecordset
   
   With SSDBMovements
        .RemoveAll
        Do While Not adoRst.EOF
           With adoRst
                StrSql = !AccountNo & vbTab
                If !Loc1 = SSDBLocations.Columns(0).Text Then
                   StrSql = StrSql & !Loc1 & vbTab
                   If !StatusDesc = "Used" Then
                      iComment = !PayeeName
                   Else
                      iComment = !Loc2
                   End If
                Else
                   StrSql = StrSql & !Loc2 & vbTab
                   If !StatusDesc = "Used" Then
                      iComment = !PayeeName
                   Else
                      iComment = !Loc1
                   End If
                End If
                StrSql = StrSql & !TransferDate & vbTab & !StartNo & vbTab & !EndNo & vbTab
                StrSql = StrSql & Format(!NoofChqs, "#,##0") & vbTab & !StatusDesc & vbTab & iComment & vbTab
                StrSql = StrSql & !UserID & vbTab & !PostDate & vbTab
           End With
           .AddItem StrSql
           adoRst.MoveNext
        Loop
   End With
   
   TxtTransfersIN = Format(adoRst1!InAmt, "#,##0")
   TxtTransfersOUT = Format(adoRst1!OutAmt, "#,##0")
   TxtCancelled = Format(adoRst1!CanAmt, "#,##0")
   TxtUsed = Format(adoRst1!UsedAmt, "#,##0")
   TxtOpeningBalance = Format(adoRst2!OpenBal, "#,##0")
   If Len(TxtOpeningBalance) < 1 Then
      TxtOpeningBalance = 0
   End If
   
   TxtClosingBalance = Format((CDbl(TxtOpeningBalance) + CDbl(TxtTransfersIN)) - (CDbl(TxtTransfersOUT) + CDbl(TxtCancelled) + CDbl(TxtUsed)), "#,##0")
   
   With SSDBActiveRanges
        .RemoveAll
        Do While Not AdoRst3.EOF
           With AdoRst3
                StrSql = !TransDate & vbTab & !StartChqNo & vbTab & !EndChqNo & vbTab & Format(!NoChqs, "#,##0") & vbTab & !UserID & vbTab
                StrSql = StrSql & !PostDate & vbTab & !AccountNo & vbTab & !LocationName
           End With
        AdoRst3.MoveNext
        .AddItem StrSql
        Loop
   End With
   adoRst.Close
   adoRst1.Close
   adoRst2.Close
   AdoRst3.Close
   Set adoRst = Nothing
   Set adoRst1 = Nothing
   Set adoRst2 = Nothing
   Set AdoRst3 = Nothing
End If
Exit_CmdStart_Click:
Exit Sub

Err_CmdStart_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on generating requested data"
Resume Exit_CmdStart_Click


End Sub

Private Sub Form_Load()
csvCenterForm Me, gblMDIFORM
'-----------------------------------
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
DTPFromDate = Date
DTPToDate = Date


End Sub

Private Sub OptReport_Click(Index As Integer)
If Index = 0 Then
   OptReport(1).Value = False
   LblFrom.Visible = False
   LblTo.Visible = False
   DTPFromDate.Visible = False
   DTPToDate.Visible = False
Else
   OptReport(0).Value = False
   LblFrom.Visible = True
   LblTo.Visible = True
   DTPFromDate.Visible = True
   DTPToDate.Visible = True
End If

End Sub
Function IsValid() As Boolean
IsValid = False
If SSDBAccount = "" Then
   MsgBox "An account must be selected", vbOKOnly
   SSDBAccount.SetFocus
   GoTo Exit_IsValid
End If

If SSDBLocations = "" Then
   MsgBox "Select a location", vbOKOnly
   SSDBLocations.SetFocus
   GoTo Exit_IsValid
End If

If OptReport(0).Value = False And OptReport(1).Value = False Then
   MsgBox "Select how much data should be reported by chosing a date range"
   GoTo Exit_IsValid
End If
IsValid = True

Exit_IsValid:
Exit Function

End Function

Private Sub SSDBAccount_InitColumnProps()
On Error GoTo Err_SSDBAccount_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_ListActiveAccounts", 1)
If adoRst.EOF Then
   MsgBox "Accounts were not setup" & vbCrLf & "Please do so now", vbCritical, "Account Error"
   GoTo Exit_SSDBAccount_InitColumnProps
End If

'adoRst.MoveFirst
With SSDBAccount
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst(0) & vbTab & adoRst(1) & vbTab
     .AddItem StrSql
     adoRst.MoveNext
     StrSql = ""
     Loop
End With
adoRst.Close
Set adoRst = Nothing
Exit_SSDBAccount_InitColumnProps:
Exit Sub

Err_SSDBAccount_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing active accounts"
Resume Exit_SSDBAccount_InitColumnProps

End Sub

Private Sub SSDBLocations_InitColumnProps()
On Error GoTo Err_SSDBLocations_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_ListLocations", 1)
If adoRst.EOF Then
   MsgBox "Cheque Locations were not setup" & vbCrLf & "Please do so now", vbCritical, "Locations Error"
   GoTo Exit_SSDBLocations_InitColumnProps
End If

'adoRst.MoveFirst
With SSDBLocations
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst(0) & vbTab & adoRst(1) & vbTab
     .AddItem StrSql
     adoRst.MoveNext
     StrSql = ""
     Loop
End With
adoRst.Close
Set adoRst = Nothing
Exit_SSDBLocations_InitColumnProps:
Exit Sub

Err_SSDBLocations_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing active locations"
Resume Exit_SSDBLocations_InitColumnProps

End Sub

