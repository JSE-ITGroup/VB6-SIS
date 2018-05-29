VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmChqManagement 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Cheque Leaves Inventory Management (CLIM)"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   Icon            =   "FrmChqManagement.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FmeUpdate 
      BackColor       =   &H00FFFF80&
      Caption         =   "Enter Update/Add Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   37
      Top             =   6960
      Visible         =   0   'False
      Width           =   10575
      Begin VB.TextBox TxtNoChqs 
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
         Left            =   9000
         TabIndex        =   43
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtEnding 
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
         Left            =   5040
         TabIndex        =   41
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox TxtStarting 
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
         Left            =   1560
         TabIndex        =   39
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF80&
         Caption         =   "No of Cheques:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7320
         TabIndex        =   42
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Ending No.:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Starting No.:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FmeRevaulted 
      BackColor       =   &H00C0FFFF&
      Caption         =   "List of Available Cheque Numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2895
      Left            =   5400
      TabIndex        =   36
      Top             =   3960
      Visible         =   0   'False
      Width           =   5175
      Begin SSDataWidgets_B.SSDBGrid SSDBRevaulted 
         Height          =   2535
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   4935
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   0
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   8705
         _ExtentY        =   4471
         _StockProps     =   79
      End
   End
   Begin VB.Frame FmeCancelled 
      BackColor       =   &H008080FF&
      Caption         =   "List of Cancelled Cheque Numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2895
      Left            =   0
      TabIndex        =   35
      Top             =   3960
      Visible         =   0   'False
      Width           =   5295
      Begin SSDataWidgets_B.SSDBGrid SSDBCancelled 
         Height          =   2535
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   5055
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   0
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   8916
         _ExtentY        =   4471
         _StockProps     =   79
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   34
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   33
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel Cheque"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   32
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton CmdRevault 
      Caption         =   "ReVault"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   31
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Frame FmeCurrentDetails 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Current Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   0
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   10575
      Begin VB.TextBox TxtCurrentBalance 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   29
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox TxtNextChqNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   27
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TxtIssuedBy 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox TxtAmount 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8160
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtLastDate 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TxtLastChqNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7440
         TabIndex        =   28
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Next Chq No.:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   4320
         TabIndex        =   26
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Issued By:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   7200
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   4320
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Last Cheque No. Used:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame FmeLastDetails 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Last Addition (IN) Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   10575
      Begin VB.TextBox TxtDateDone 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         TabIndex        =   16
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox TxtDone 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox TxtResultBal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox TxtNoofLeaves 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtEndingNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtStartingNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Date Done:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   7080
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Done By:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Resulting Balance:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "No of Leaves:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ending No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Starting No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FmeOptions 
      BackColor       =   &H00808000&
      Caption         =   "Select an option from below"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.OptionButton OptChqType 
         BackColor       =   &H00808000&
         Caption         =   "View/Update The Vault"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   2
         Left            =   6720
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.OptionButton OptChqType 
         BackColor       =   &H00808000&
         Caption         =   "View/Update Dividend Stock"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin VB.OptionButton OptChqType 
         BackColor       =   &H00808000&
         Caption         =   "View/Update Working Stock"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "FrmChqManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Public adoRst As ADODB.Recordset
Dim iErr As String
Dim UpdateOpt As Integer

Private Sub CmdAdd_Click()
On Error GoTo Err_CmdAdd_Click

UpdateOpt = 1
FmeUpdate.Visible = True

Exit_CmdAdd_Click:
Exit Sub

Err_CmdAdd_Click:
iErr = Err & " " & Err.Description & vbCrLf
iErr = iErr & "Please advise your System Administrator"
MsgBox iErr
GoTo Exit_CmdAdd_Click
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Err_CmdCancel_Click

FmeUpdate.Visible = True
UpdateOpt = 3
Exit_CmdCancel_Click:
Exit Sub

Err_CmdCancel_Click:
iErr = Err & " " & Err.Description & vbCrLf
iErr = iErr & "Please advise your System Administrator"
MsgBox iErr
GoTo Exit_CmdCancel_Click

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdRevault_Click()
On Error GoTo Err_CmdRevault_Click

UpdateOpt = 2
FmeUpdate.Visible = True

Exit_CmdRevault_Click:
Exit Sub

Err_CmdRevault_Click:
iErr = Err & " " & Err.Description & vbCrLf
iErr = iErr & "Please advise your System Administrator"
MsgBox iErr
GoTo Exit_CmdRevault_Click

End Sub

Private Sub CmdSave_Click()
Dim InvType As String
Dim i As Integer
Dim ResBal As Double
Dim Strsql As String

If IsDigitsOnly0(TxtStarting) = False Then
   MsgBox "Please enter a valid starting cheque number"
   GoTo Exit_CmdSave_Click
End If

If IsDigitsOnly0(TxtEnding) = True Then
   If TxtEnding < TxtStarting Then
      MsgBox "Ending number is less than starting number. Please correct"
      GoTo Exit_CmdSave_Click
   End If
Else
   MsgBox "Please enter a valid ending cheque number"
   GoTo Exit_CmdSave_Click
End If
If IsDigitsOnly0(TxtNoChqs) = False Then
   Strsql = "The number of cheques in the range provided " & vbCrLf
   Strsql = Strsql & "cannot be blank or 0." & vbCrLf
   Strsql = Strsql & "Please correct and resubmit."
   MsgBox Strsql
   GoTo Exit_CmdSave_Click
End If
ResBal = CDbl(TxtEnding) - CDbl(TxtStarting) + 1
If ResBal <> CDbl(TxtNoChqs) Then
   Strsql = "The figure entered in the No. of Cheques field is incorrect" & vbCrLf
   Strsql = "The calculated figure is " & ResBal & "." & vbCrLf
   Strsql = Strsql & "Please correct and resubmit."
   MsgBox Strsql
   GoTo Exit_CmdSave_Click
End If

If OptChqType(0).Value = True Then
   InvType = "W"
End If
If OptChqType(1).Value = True Then
   InvType = "P"
End If
If OptChqType(0).Value = True Then
   InvType = "R"
End If

If UpdateOpt = 1 Then
i = RunSP(SpCon, "usp_AddToChqInventory", 0, a, InvType, TxtStarting, TxtEnding, TxtNoChqs, gluserid)
If i = 1 Then
   MsgBox "Number already exists"
   GoTo Exit_CmdSave_Click
End If
If i = 2 Then
   MsgBox "Number does not exist in Cheque Inventory"
   GoTo Exit_CmdSave_Click
End If
If i = 3 Then
   MsgBox "Number already used in Working Stock"
   GoTo Exit_CmdSave_Click
End If
If i = 4 Then
   MsgBox "Number already used in Payment Stock"
   GoTo Exit_CmdSave_Click
End If
If i = 0 Then
   MsgBox "Information update was successful"
   GoTo Exit_CmdSave_Click
End If
End If
If UpdateOpt = 2 Then
   MsgBox "Revault"
End If


Exit_CmdSave_Click:
Exit Sub

Err_CmdSave_Click:
iErr = Err & " " & Err.Description & vbCrLf
iErr = iErr & "Please advise your System Administrator"
MsgBox iErr
GoTo Exit_CmdSave_Click
End Sub

Private Sub Form_Activate()
OptChqType(0).Value = False
OptChqType(1).Value = False
OptChqType(2).Value = False
UpdateOpt = 0
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
frmMDI.txtStatusMsg.Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Exit_Form_Unload

Set FrmChqManagement = Nothing
If adoRst.State <> 0 Then
   adoRst.Close
End If
Set adoRst = Nothing
Exit_Form_Unload:
SpCon.Close
End Sub

Private Sub OptChqType_Click(Index As Integer)
Dim ChqType As String

If FmeLastDetails.Visible = True Then
   FmeLastDetails.Visible = False
   adoRst.Close
End If

If Index = 0 Then
   ChqType = "W"
Else
   If Index = 1 Then
      ChqType = "D"
   Else
      ChqType = "V"
   End If
End If
   
Set adoRst = RunSP(SpCon, "usp_SelectChqInventory", 1, gblFileKey, ChqType)
If adoRst.State = 0 Then
   MsgBox "No Records returned"
   GoTo Exit_OptChqType_Click
End If
LoadData

Exit_OptChqType_Click:
Exit Sub

End Sub
Private Sub LoadData()
Dim adoCurrent As ADODB.Recordset
Dim adoCancelled As ADODB.Recordset
Dim adoVaulted As ADODB.Recordset
Dim adoOverall As ADODB.Recordset
Dim adoCounts As ADODB.Recordset

FmeLastDetails.Visible = True
TxtStartingNo = adoRst!StartNo
TxtEndingNo = adoRst!EndNo
TxtNoofLeaves = adoRst!NoCheques
TxtResultBal = adoRst!Remaining
TxtDone = adoRst!UserId
TxtDateDone = Format(adoRst!TransDate, "dd-mmm-yyyy")

Set adoCurrent = adoRst.NextRecordset
If adoCurrent.State <> 1 Then
   FmeCurrentDetails.Visible = True
   TxtLastChqNo = adoCurrent!EndNo
   TxtLastDate = Format(adoCurrent!TransDate, "dd-mmm-yyyy")
   TxtNextChqNo = adoCurrent!NextNo
   TxtIssuedBy = adoCurrent!UserId
   TxtCurrentBalance = adoCurrent!Remaining
   TxtAmount = adoCurrent!ChqAmt
End If

Set adoCancelled = adoRst.NextRecordset
If adoCancelled.State <> 1 Then
With SSDBCancelled
     Do While Not adoCancelled.EOF
        Strsql = adoCancelled!StartNo & vbTab & adoCancelled!EndNo & vbTab
        Strsql = Strsql & adoCancelled!NoCheques & vbTab
        Strsql = Strsql & adoCancelled!UserId & vbTab
        Strsql = Strsql & Format(adoCancelled!TransDate, "dd-mmm-yyyy")
        .AddItem Strsql
        adoCancelled.MoveNext
     Loop
End With
End If

Set adoVaulted = adoRst.NextRecordset
If adoVaulted.State <> 1 Then
With SSDBRevaulted
     Do While Not adoVaulted.EOF
        Strsql = adoVaulted!StartNo & vbTab & adoVaulted!EndNo & vbTab
        Strsql = Strsql & adoVaulted!NoCheques & vbTab
        Strsql = Strsql & adoVaulted!UserId & vbTab
        Strsql = Strsql & Format(adoVaulted!TransDate, "dd-mmm-yyyy")
        .AddItem Strsql
        adoVaulted.MoveNext
     Loop
End With
End If
Set adoOverall = adoRst.NextRecordset
If adoOverall.State <> 1 Then
With adoOverall
     TxtWSRemain = 0
     TxtDivRemain = 0
     TxtRegRemain = 0
     .Close
End With
End If
Set adoCounts = adoRst.NextRecordset
If adoCounts.State <> 1 Then
   With adoCounts
        TxtOCancelled = 0
        TxtORevaulted = 0
        .Close
   End With
End If

End Sub

