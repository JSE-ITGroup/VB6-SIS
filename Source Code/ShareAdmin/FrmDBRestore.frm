VERSION 5.00
Begin VB.Form FrmDBRestore 
   Caption         =   "Restore Database"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   Icon            =   "FrmDBRestore.frx":0000
   ScaleHeight     =   1965
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtFileName 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Select Backup File:"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "FrmDBRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdRestore_Click()
On Error GoTo Exit_CmdRestore_Click
Dim i As Integer

If TxtFileName = vbNullString Then
   MsgBox "Please enter a valid file name"
   TxtFileName.SetFocus
   GoTo Exit_CmdRestore_Click
End If

i = RunSP(SpCon, "sp_db_restore", 0, TxtFileName)
If i <> 0 Then
   MsgBox "There was an error in the Restoration Process. Maybe the FileName should be checked"
Else
   MsgBox "Restoration was successful. You will need to login to SISRESTORE to use it"
End If

Screen.MousePointer = vbDefault
Exit_CmdRestore_Click:
Exit Sub

Err_CmdRestore_Click:
MsgBox Err.Description, vbOKOnly, "Database Restoration to SISRESTORE"
GoTo Exit_CmdRestore_Click

End Sub

Private Sub Form_Load()

frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   csvCenterForm Me, gblMDIFORM
   '--
   On Error GoTo Err_Form_Load:
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
  
Exit_Form_Load:
Exit Sub
Err_Form_Load:
MsgBox Err.Description, vbOKOnly, "DB Restore Form Load"
GoTo Exit_Form_Load

End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close

End Sub
