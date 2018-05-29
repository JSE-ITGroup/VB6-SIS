VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmSISRestore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restore Dividend FIles"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "SISRestore.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6735
   Begin SSDataWidgets_A.SSDBOptSet Opt 
      Height          =   255
      Left            =   3285
      TabIndex        =   6
      Top             =   840
      Width           =   2535
      _Version        =   196611
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Main Ledger"
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      IndexSelected   =   0
      Buttons.Button(0).OptionValue=   "0"
      Buttons.Button(0).Caption=   "Main Ledger"
      Buttons.Button(0).Mnemonic=   77
      Buttons.Button(0).Value=   -1  'True
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   74
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   76
      Buttons.Button(0).PictureRight=   75
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   168
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Restore"
      Default         =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   4320
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      ToolTipText     =   "Returns to main menu"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdBtn 
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      Height          =   300
      Index           =   0
      Left            =   5400
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      ToolTipText     =   "Returns to main menu"
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Select Which Files to Restore:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   2700
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   480
      Y2              =   480
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
      TabIndex        =   4
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSISRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpCon As ADODB.Connection

Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo cmdBtn_Click_Err
Dim iRecs As Integer
Dim Regis As String

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

Select Case Index
Case 0 'Cancel
    
    '''Set cmdChange = Nothing
   Set frmSISRestore = Nothing
   Unload Me
  '''  frmSIS013.Visible = True
     '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh

Case Else
    'Run calculation for selected option
    '0 - Main Register
    '1 - JCSD
    '2 - TTCD
    
    iRecs = RunSP(SpCon, "usp_RestoreTables", 0, Opt.OptionValue)
    If iRecs = 0 Then
       MsgBox "Restoraton of dividend tables successfully completed"
    Else
       MsgBox "There was an error which caused the restoration to be interupted"
       GoTo cmdBtn_Click_Exit
    End If
    frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
    Screen.MousePointer = vbDefault
    frmMDI.txtStatusMsg.Refresh

End Select


cmdBtn_Click_Exit:

Exit Sub
cmdBtn_Click_Err:
GoTo cmdBtn_Click_Exit
End Sub

Private Sub Form_Activate()
Dim adoRst As ADODB.Recordset
Dim i As Integer

Set adoRst = RunSP(SpCon, "usp_ListStockExchanges", 1)
i = 1
With adoRst
     Do While Not .EOF
        Opt.Buttons.Add (1)
        Opt.Buttons.Item(i).Caption = !ExchangeABBR
        Opt.Buttons.Item(i).OptionValue = !StockExchangeID
        i = i + 1
        .MoveNext
     Loop
End With
adoRst.Close
Set adoRst = Nothing
End Sub

Private Sub Form_Load()
   csvCenterForm Me, gblMDIFORM
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '----------------------
   '--  disable menu items
   '----------------------
   frmMDI.mnuFile.Enabled = False
   frmMDI.btnClose.Enabled = False
   frmMDI.mnuLists.Enabled = False
   frmMDI.mnuAct.Enabled = False
   frmMDI.mnuAdm.Enabled = False
   '--
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

   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SISRestore/Load"
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
Set frmSISRestore = Nothing
frmMDI.mnuFile.Enabled = True
frmMDI.btnClose.Enabled = True
frmMDI.mnuLists.Enabled = True
frmMDI.mnuAct.Enabled = True
frmMDI.mnuAdm.Enabled = True
   
End Sub
