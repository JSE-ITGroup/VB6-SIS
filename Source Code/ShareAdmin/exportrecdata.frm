VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form exportrecdata 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Export Reconciliation Data"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "exportrecdata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CmnDialog 
      Left            =   4080
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "&Exit"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "&Submit"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   50462721
      CurrentDate     =   37722
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Dividend Date to Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2745
   End
End
Attribute VB_Name = "exportrecdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub Command1_Click()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sql As String

Set rs = RunSP(SpCon, "usp_ExportRecon", 1, Format(dtp.Value, "dd-mmm-yyyy"))
If rs.EOF Then
   MsgBox " Dividend Information For that Date was not found", vbCritical
   dtp.SetFocus
   Exit Sub
Else
   CmnDialog.DialogTitle = "Export Dividend File"
   CmnDialog.Filter = "TXT(*.txt)|*.txt"
   CmnDialog.DefaultExt = "txt"
   CmnDialog.filename = "recondata"
   CmnDialog.ShowSave
   If Len(CmnDialog.filename) < 1 Then
     MsgBox "Save Abondoned"
     Exit Sub
   End If
   myoutfile = CmnDialog.filename
   Open myoutfile For Output As #1
   Do Until rs.EOF
   DoEvents
        Label2.Caption = "Processing Please wait..."
        Label2.Visible = True
        Label2.Refresh
        chqno = rs!CHQNUM
        clientide = rs!ClientID
        chqdt = Format(rs!CHQDAT, "dd/mm/yy")
        chqamount = Format(rs!ChqAmt, "#############0.00")
        reconin = rs!reconind
        repchno = rs!RepChqNo
        revaldt = rs!revaldat
        foliomh = rs!FOLIOMTH
        recondt = rs!recondat
        decdt = rs!DECDATE
        paytype = rs!PAYTYP
        payeenam = rs!PAYEENAME
        
        Print #1, chqno & Space(20 - Len(chqno)) & "|" & chqdt & Space(10 - Len(chqdt)) & "|" & chqamount & Space(20 - Len(chqamount)) & "|" & payeenam
        
        rs.MoveNext
   Loop
   Close #1
   rs.Close
   MsgBox "Recon Data Successfully Generated.Press OK to Continue", vbInformation
   Label2.Visible = False
End If
End Sub

Private Sub Command2_Click()
Unload exportrecdata
End Sub

Private Sub Form_Load()
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
Loop
Screen.MousePointer = vbDefault

dtp.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub
