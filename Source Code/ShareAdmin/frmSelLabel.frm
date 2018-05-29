VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmSelLabel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Labels Selection"
   ClientHeight    =   2400
   ClientLeft      =   1635
   ClientTop       =   2385
   ClientWidth     =   4275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSelLabel.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2400
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin SSDataWidgets_A.SSDBOptSet Opt 
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   1470
      _Version        =   196611
      _ExtentX        =   2593
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "3-Up"
      BackColor       =   -2147483643
      Cols            =   2
      IndexSelected   =   1
      NumberOfButtons =   2
      Buttons.Button(0).OptionValue=   "1"
      Buttons.Button(0).Caption=   "1-Up"
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   38
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   40
      Buttons.Button(0).PictureRight=   39
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   48
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(1).OptionValue=   "3"
      Buttons.Button(1).Caption=   "3-Up"
      Buttons.Button(1).Mnemonic=   83
      Buttons.Button(1).Value=   -1  'True
      Buttons.Button(1).TextLeft=   64
      Buttons.Button(1).TextRight=   87
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   49
      Buttons.Button(1).ButtonRight=   62
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   89
      Buttons.Button(1).PictureRight=   88
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   49
      Buttons.Button(1).ButtonToColRight=   97
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).ButtonBitmapID=   2
      Buttons.Button(1).Column=   1
   End
   Begin VB.ComboBox cbOptions 
      Height          =   315
      ItemData        =   "frmSelLabel.frx":000C
      Left            =   840
      List            =   "frmSelLabel.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select the search argument."
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ComboBox cbWhere 
      Height          =   315
      ItemData        =   "frmSelLabel.frx":0010
      Left            =   840
      List            =   "frmSelLabel.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select search field which will also be used to sequence the report."
      Top             =   600
      Width           =   1725
   End
   Begin VB.CommandButton cmdFind 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Find"
      Default         =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox tbFind 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Enter selection value. Use ""ALL"" to select entire file."
      Top             =   120
      Width           =   3345
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   1140
      Width           =   540
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Where"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   660
      Width           =   480
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   3
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "frmSelLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsClients As New ADODB.Recordset
Dim iOpen As Integer
Dim unPk As cUnPackName
Dim sNewfile As String
Private Sub cmdFind_Click(Index As Integer)
Dim i As Integer, X As Integer
Dim sWhere As String, sql As String, sFldName As String
On Error GoTo cmdFind_Err
Select Case Index
Case 0
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Selecting Records for Labels..."
'----------------------------
'-------- main line ---------
'----------------------------
sql = "SELECT CLINAME, a.CLIENTID, CLIADDR1,"
sql = sql & " CLIADDR2, CLIADDR3, CLIADDR4, CLIADDR5,"
sql = sql & " a.CATCODE, b.CATDESC "
sql = sql & " FROM STKNAME a, STKCAT b"
sql = sql & " WHERE a.CATCODE = b.CATCODE"
sql = sql & " and SHARES > 0"
'--
If tbFind.Text = vbNullString Then
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Exit Sub
End If
sWhere = UCase(Trim(tbFind.Text))
If sWhere = "" Then
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Exit Sub
End If
'---
X = cbWhere.ListIndex
Select Case X
Case 0
  sFldName = "CATDESC"
Case 1
  sFldName = "a.CATCODE"
Case 2
  sFldName = "CLINAME"
Case 3
  sFldName = "CLIENTID"
End Select
'---
If cbOptions.Enabled = False Then 'create labels for all
  GoTo SET_ORDER_BY
End If
'---
If cbOptions.ListIndex > 6 Then cbOptions.ListIndex = 0
Select Case cbOptions.ListIndex
Case 0 ' Exact Match
   sql = sql & " and " & sFldName
   If X = 3 Then  'numeric key for client id
      sql = sql & " = " & Val(sWhere)
   Else
      sql = sql & " = '" & sWhere & "' "
   End If
Case 1 ' Starts With
   sWhere = Trim(tbFind.Text) & "%"
Case 2 ' Ends With
   sWhere = "%" & Trim(tbFind.Text)
Case 3 ' AnyWhere
   sWhere = "%" & Trim(tbFind.Text) & "%"
End Select
If cbOptions.ListIndex <> 0 Then
 sql = sql & " and " & sFldName & " like '" & sWhere & "' "
End If
'---
SET_ORDER_BY:
 sql = sql & " order by " & sFldName
 If sFldName = "CATDESC" Then
    sql = sql & ", CLINAME"
 Else
    If sFldName = "CLINAME" Then
      sql = sql & ", CLIENTID"
    End If
 End If
'---
rsClients.Open sql, cnn, , , adCmdText
iOpen = True
Process_Labels
Case 1
  Unload Me
  If iOpen = True Then rsClients.Close
  Set rsClients = Nothing
  '''set cnn = nothing
  Set unPk = Nothing
  Set frmSIS048 = Nothing
End Select
cmdFind_Exit:
  Exit Sub
cmdFind_Err:
  MsgBox "SIS048/cmdFind"
   Unload Me
  End Sub

Private Sub Form_Load()
'''Set cnn = New ADODB.Connection
'''cnn.Open cnn
Set unPk = New cUnPackName
' readymsg
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   csvCenterForm Me, gblMDIFORM
'- load comparison key fields and show frmFind
    '---------------------------------------------
    cbWhere.AddItem "Category Name"
    cbWhere.AddItem "Category Code"
    cbWhere.AddItem "Client Name"
    cbWhere.AddItem "Account No"
    cbWhere.ListIndex = 0
    cbOptions.ListIndex = 0
    cbOptions.Enabled = False
    tbFind.Text = vbNullString
 '--
 Set rsClients = New ADODB.Recordset
 iOpen = False
End Sub



Private Sub tbFind_LostFocus()
If tbFind.Text <> vbNullString Then
   If UCase(Trim(tbFind.Text)) <> "ALL" Then
      cbOptions.Enabled = True
   Else
      cbOptions.Enabled = False
   End If
End If
End Sub

Private Sub Process_Labels()
Dim fso, txtfile, iErr As Integer
Dim sOutRec As String, msg As String, iReply As Integer
sNewfile = App.Path & "\sis048.txt"

'create txt file for printing labels
With rsClients
   .Requery
   If .EOF Then ' SELECT RETURNED NO RECORDS
   '---
     iErr = 164
     csvShowUsrErr iErr, "Labels Selection"
   Else
   ' process returned recordset
   '--
      iErr = 0
      frmMDI.txtStatusMsg.SimpleText = "Creating Merge File....."      'open file for output & write header record
      '--
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set txtfile = fso.CreateTextFile(sNewfile, True)
      sOutRec = ""
      sOutRec = sOutRec & "CLINAME" & Chr(9)
      sOutRec = sOutRec & "ADDRESS1" & Chr(9)
      sOutRec = sOutRec & "ADDRESS2" & Chr(9)
      sOutRec = sOutRec & "ADDRESS3" & Chr(9)
      sOutRec = sOutRec & "ADDRESS4" & Chr(9)
      sOutRec = sOutRec & "ADDRESS5" & Chr(9)
      sOutRec = sOutRec & "CLIENTID"
      txtfile.writeline (sOutRec)
      ' BUILD DATA FILE FROM RECORDSET
      '--
      .MoveFirst
      Do While Not .EOF
        sOutRec = ""
        If unPk.Unpack(!CliName) = True Then
            sOutRec = sOutRec & unPk.FirstName & _
                   " " & unPk.LastName & Chr(9)
        Else
           sOutRec = sOutRec & !CliName & Chr(9)
        End If
        sOutRec = sOutRec & !CliAddr1 & Chr(9)
        sOutRec = sOutRec & !CliAddr2 & Chr(9)
        sOutRec = sOutRec & !CliAddr3 & Chr(9)
        sOutRec = sOutRec & !CliAddr4 & Chr(9)
        sOutRec = sOutRec & !CliAddr5 & Chr(9)
        sOutRec = sOutRec & !ClientID
        txtfile.writeline (sOutRec)
        .MoveNext
      Loop
      txtfile.Close
   End If
   iOpen = False
   .Close
End With
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
If iErr = 164 Then Exit Sub
msg = "Select Yes to Activate Predefined Microsoft Word Document or No to end"
iReply = MsgBox(msg, vbQuestion + vbYesNo, "Labels")
If iReply = vbYes Then
  Print_Report
End If
End Sub

Private Sub Print_Report()
Dim AppWord As Word.Application
Dim sDoc As String, iWarn As String
On Error GoTo Open_error
iWarn = 165
csvShowUsrErr iWarn, "Labels Selection"
frmMDI.txtStatusMsg.SimpleText = "Running Word to Print Labels..."
If Opt.IndexSelected = 1 Then
  sDoc = App.Path & "\sis048T.doc"
Else
  sDoc = App.Path & "\sis048o.doc"
End If
Set AppWord = CreateObject("Word.application")
Screen.MousePointer = vbDefault
With AppWord
 .Documents.Open (sDoc)
 .ActiveDocument.MailMerge.OpenDataSource (sNewfile)
End With
With AppWord.ActiveDocument.MailMerge
  .Destination = wdSendToPrinter
  .MailAsAttachment = False
  .MailAddressFieldName = ""
  .MailSubject = ""
  .SuppressBlankLines = True
  With .DataSource
      .FirstRecord = wdDefaultFirstRecord
      .LastRecord = wdDefaultLastRecord
  End With

 .Execute
End With
AppWord.ActiveDocument.Close
AppWord.DisplayAlerts = False
AppWord.Quit

Exit Sub
Open_error:
  
  MsgBox "Word Open Error " & Err.Number & "-" & Err.Description, , "Labels Selection"
 
End Sub
