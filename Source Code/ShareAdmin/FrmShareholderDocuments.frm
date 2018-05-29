VERSION 5.00
Begin VB.Form FrmShareholderDocuments 
   BackColor       =   &H80000013&
   Caption         =   "List of documents presented by"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "FrmShareholderDocuments.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save Changes"
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
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   2175
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
      Height          =   495
      Left            =   8880
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CheckBox ChkDocument 
      BackColor       =   &H80000013&
      Caption         =   "abcdefghiklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefgh"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "FrmShareholderDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim TotDocs As Integer
Dim iClientID As Long

Private Sub CmdExit_Click()

Unload Me
End Sub

Private Sub CmdSave_Click()
On Error GoTo Err_CmdSave_Click
Dim NoDocs As Integer
Dim DocIds As String
Dim i As Integer

NoDocs = 0

For i = 1 To TotDocs
    If ChkDocument(i).Value = 1 Then
       NoDocs = NoDocs + 1
       DocIds = DocIds & i & ";"
    End If
Next i

If NoDocs > 0 Then
   i = RunSP(SpCon, "usp_UpdateshareholderDocuments", 0, iClientID, NoDocs, DocIds, gblLoginName)
   If i <> 0 Then
        MsgBox "Document update failed"
   Else
        MsgBox "Document update was successful"
   End If
End If

Exit_CmdSave_Click:
Exit Sub

Err_CmdSave_Click:
MsgBox Err & " " & Err.Description, vbOKOnly
Resume Exit_CmdSave_Click
End Sub

Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
Dim adoRst As ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim k As Integer

Set adoRst = RunSP(SpCon, "usp_SelectAvailableDocuments", 1)
i = 1
With adoRst
     ChkDocument(i).Caption = !DocumentName
     j = ChkDocument(i).Left
     k = ChkDocument(i).Top
     GoTo LoadRest
     Do While Not .EOF
           Load ChkDocument(i)
           ChkDocument(i).Caption = !DocumentName
           ChkDocument(i).Visible = True
           ChkDocument(i).Top = k + 300
           k = ChkDocument(i).Top
           ChkDocument(i).Left = j
LoadRest:
           i = i + 1
           If i = 17 Then
              j = 5280
              k = 120
           End If
        .MoveNext
     Loop
End With
TotDocs = i - 1

adoRst.Close
Set adoRst = Nothing
If Len(gblFileKey) > 1 Then
   CmdSave.Visible = True
   Me.Caption = Me.Caption & " " & gblFileKey
   iClientID = gblHold
   UpdateScreen
End If

Exit_Form_Activate:
Exit Sub

Err_Form_Activate:
MsgBox Err.Description, vbOKOnly, "Error on listing currencies"
GoTo Exit_Form_Activate
End Sub

Private Sub Form_Load()
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

   '--
   ' ready message
   '--------------
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "Shareholders document List Error"
  Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SpCon.Close
Set FrmShareholderDocuments = Nothing
End Sub

Private Sub UpdateScreen()
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_GetShareholderDocuments", 1, iClientID)
With adoRst
     Do While Not .EOF
        i = !DocID
        ChkDocument(i).Value = 1
        .MoveNext
     Loop
End With
adoRst.Close
Set adoRst = Nothing

End Sub
