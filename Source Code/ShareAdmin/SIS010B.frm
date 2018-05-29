VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS010 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Mandates"
   ClientHeight    =   5436
   ClientLeft      =   1092
   ClientTop       =   336
   ClientWidth     =   6804
   Icon            =   "SIS010.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5436
   ScaleWidth      =   6804
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
      DataFieldList   =   "Column 1"
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3836
      Columns(0).Caption=   "Branch Name"
      Columns(0).Name =   "Branch Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1535
      Columns(1).Caption=   "Bank Id"
      Columns(1).Name =   "Bank Id"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 1"
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   9
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   9
      ToolTipText     =   "Enter the name of the account if different from the shareholder."
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   8
      Left            =   1800
      MaxLength       =   16
      TabIndex        =   8
      ToolTipText     =   "Enter an account number to credit the shareholder's payments. This is optional."
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   4
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   4
      ToolTipText     =   "Enter Address line 2"
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   5
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   5
      ToolTipText     =   "Enter Address Line 3"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   6
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   6
      ToolTipText     =   "Enter Address Line 4"
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   7
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   7
      ToolTipText     =   "Enter Address Line 5"
      Top             =   3360
      Width           =   3375
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "Enter the date this mandate becomes effective."
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   445
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   3
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   3
      ToolTipText     =   "Enter the first line of the mandates address."
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   2
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   2
      ToolTipText     =   "Enter the name of the bank or recipient to receive the payment."
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3600
      TabIndex        =   21
      ToolTipText     =   "Clears the screen and resets it if in edit mode."
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5760
      TabIndex        =   15
      ToolTipText     =   "Cancels changes and returns to Account maintenance."
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   4680
      TabIndex        =   14
      ToolTipText     =   "Update Joint Table for saving to disk by Accounts Maintainace"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox tbfld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   12
      ToolTipText     =   "Use generate number or enter your own unique client Number"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox tbfld 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   13
      ToolTipText     =   "Enter Address line 2"
      Top             =   840
      Width           =   4335
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   11
      ToolTipText     =   "Enter the date this mandate ceases."
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   445
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   10
      Top             =   4560
      Width           =   1935
      DataFieldList   =   "Column 1"
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3836
      Columns(0).Caption=   "Payment By"
      Columns(0).Name =   "Payment By"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   25
      Columns(1).Width=   1614
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   3
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Bank Id:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   29
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Payment Method:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   28
      Top             =   4560
      Width           =   1500
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   27
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   960
      TabIndex        =   26
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "End Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   1620
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   23
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   22
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblLabels 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   19
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ver:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Stockholder No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   1620
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iErr As Integer
Dim X As Integer
Dim rsMandate As ADODB.Recordset
Dim rsBank As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim OpenErr As Integer
Dim iOpenMan As Integer
Dim iOpenBank As Integer
Dim strTable As String
Dim iMode As Integer  ' 0 = New; 1 = Active; 2 = inactive joint
Function IsValid() As Integer
On Error GoTo IsValid_Err
Dim dtefld As Date
iErr = 0
IsValid = False
'--
If iMode = 0 Or iMode = 2 Then
  '--
  If meb(0).Text = "" Then  ' Start Date
     iErr = 36
     csvShowUsrErr iErr, "Mandates"
     meb(0).SetFocus
     GoTo Validate_Exit
  End If
  '--
  If Not IsDate(meb(0).Text) Then
     iErr = 14
     csvShowUsrErr iErr, "Mandates"
     meb(0).SetFocus
     GoTo Validate_Err
  End If
  '--
  If iMode = 2 Then
   dtefld = meb(0).Text
   With rsMandate
        .MoveFirst
        Do While Not .EOF
           If dtefld <= !MndStaDte Then
               iErr = 104
               csvShowUsrErr iErr, "Mandates"
               meb(0).SetFocus
               Exit Do
           Else
               .MoveNext
           End If
        Loop
        If iErr = 104 Then GoTo Validate_Exit
     End With
  End If
  '--
  If dbc(0).Text = "" Then
     dbc(0) = "NONE"
  Else
     GoTo ContinueValid
  End If
  
  If tbfld(2) = "" Then 'Mandates Name
     iErr = 106
     tbfld(2).SetFocus
     GoTo Validate_Err
   End If
  tbfld(2) = Trim(tbfld(2))
  '--
  If tbfld(3) = "" Then 'Address 1
     iErr = 9
'     tbfld(3).SetFocus
     GoTo Validate_Err
  End If
  tbfld(3) = Trim(tbfld(3))
  '--
  If tbfld(4) = "" Then ' Address 2
       iErr = 9
       tbfld(4).SetFocus
       GoTo Validate_Err
  End If
  tbfld(4) = Trim(tbfld(4))
  '--
ContinueValid:
  If dbc(0).Text <> "NONE" Then  'Bank account number
    If IsNothing(tbfld(8)) Then
     iErr = 176
     tbfld(8).SetFocus
     GoTo Validate_Err
    End If
    tbfld(8) = Trim(tbfld(8))
    '--
    If IsNothing(tbfld(9)) Then   ' Bank Account Name
     iErr = 177
     tbfld(9).SetFocus
     GoTo Validate_Err
    End If
     '--
     tbfld(9) = Trim(tbfld(9))
     '--
  End If
  If dbc(1).Text = "" Then 'Payment Method
        iErr = 109
        dbc(1).SetFocus
        GoTo Validate_Err
   End If
Else
  If meb(1).Text = "" Then
      iErr = 37
      csvShowUsrErr iErr, "Mandates"
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
   '--
   If Not IsDate(meb(1).Text) Then
     iErr = 14
     csvShowUsrErr iErr, "Mandates"
     meb(1).SetFocus
     GoTo Validate_Exit
   End If
   '--
   dtefld = meb(1).Text
   If dtefld < Format(meb(0).Text, "dd-mmm-yyyy") Then
      iErr = 38
      csvShowUsrErr iErr, "Mandates"
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
  '--
End If
'--
IsValid = True
Validate_Exit:
   Exit Function
'--
Validate_Err:
  csvShowUsrErr iErr, "Mandates"
  GoTo Validate_Exit
'--
IsValid_Err:
  MsgBox "SIS010/IsValid"
  csvLogError "SIS010/IsValid", Err.Number, Err.Description
  Shutdown
  Unload Me
End Function

Private Sub cmdCancel_Click()
  Shutdown
  frmSIS001.Show
  Unload Me
End Sub
Private Sub cmdClear_Click()
If iMode = 0 Then
   ClearScreen
   meb(0).SetFocus
Else
   ClearScreen
   UpdateScreen
End If
End Sub
Private Sub cmdUpdate_Click()
Dim strChg As Integer, iAcct As Long
Dim i As Integer
Dim newval As Integer
'''On Error GoTo cmdUpdate_Err
If IsValid Then
   iAcct = Val(tbfld(0).Text)
  '--
i = RunSP(SpCon, "usp_MandateUpdate", 0, iMode, iAcct, gblLoginName, Format(meb(0).Text, "dd-mmm-yyyy"), tbfld(8), tbfld(9), Trim(dbc(0).Columns(1).Text), tbfld(2), tbfld(3), tbfld(4), tbfld(5), tbfld(6), tbfld(7), Format(meb(1).Text, "dd-mmm-yyyy"), dbc(1).Columns(1).Text)
If i = 1 Then
   MsgBox "Record sucessfully updated"
Else
   MsgBox "Update was unsucessfull. Sorry for any inconvienience caused"
   GoTo Done
End If

If iMode = 1 Then
     EnableData
     iMode = 0
     meb(0).Text = DateValue(meb(1).Text) + 1
     tbfld(2).SetFocus
Else
     cmdCancel_Click
End If
End If

'---

Done:
 Exit Sub
'--
cmdUpdate_Err:
  MsgBox "SIS010/cmdUpdate"
  Shutdown
  Unload Me
  frmSIS001.Show
End Sub

Private Sub dbc_InitColumnProps(Index As Integer)
Dim sRowinfo As String
On Error GoTo dbc_InitColumnProps_Err
Select Case Index
Case 0
  dbc(0).RemoveAll
  With rsBank
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        sRowinfo = !BnkName & Chr(9) & !BankId
        dbc(0).AddItem sRowinfo
       .MoveNext
      Loop
    End If
  dbc(0).Refresh
  End With
  '--
Case 1
 dbc(1).RemoveAll
 dbc(1).AddItem "Local Cheque" & Chr(9) & "CHQ"
 dbc(1).AddItem "Local Lodgement Cheque" & Chr(9) & "LLC"
 dbc(1).AddItem "Foreign Lodgment" & Chr(9) & "FLL"
 '--
Case Else
End Select
Exit Sub
dbc_InitColumnProps_Err:
  MsgBox "SIS010/dbcInitColProps"
  csvLogError "SIS010/dbcInitColProps", Err.Number, Err.Description
  Unload Me
End Sub

Private Sub dbc_LostFocus(Index As Integer)
Select Case Index
Case 0
  If UCase(dbc(0).Text) = "NONE" Then
     EnableAddr
  Else
     tbfld(2).Text = dbc(0).Columns(0).Text
     dbc(1).SelBookmarks.RemoveAll
     dbc(1).MoveFirst
     dbc(1).MoveNext
     dbc(1).Text = dbc(1).Columns(0).Text
     dbc(1).SelBookmarks.Add dbc(1).Bookmark
     dbc(1).Refresh
     DisableAddr
  End If
Case Else
End Select
End Sub

Private Sub Form_Activate()
'--
' ready message
'---
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 If OpenErr = True Then
  Unload Me
 Else
  UpdateScreen
 End If
End Sub

Private Sub Form_Load()
Dim qSQL As String, qView As String
Dim strTmp As String
'On Error GoTo FL_ERR
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
 tbfld(0).BackColor = &HC0C0C0
 tbfld(1).BackColor = &HC0C0C0
'--
csvCenterForm Me, gblMDIFORM
'-----------------------------------
'Set rsBank = New ADODB.Recordset
'Set rsMandate = New ADODB.Recordset
Set SpCon = New ADODB.Connection
With SpCon
     .ConnectionString = gblFileName
     .CursorLocation = adUseClient
     '.Provider = "SQLOLEDB.1"
End With
SpCon.Open , , , adAsyncConnect
Do While SpCon.State = adStateConnecting
   Screen.MousePointer = vbHourglass
   frmMDI.txtStatusMsg.SimpleText = "Connecting, Please wait......"
   'frmMDI.txtStatusMsg.Refresh
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

OpenErr = False
iOpenMan = False
iOpenBank = False
'----------------------------
'---- open recordsets -----
'-- create SQL for selecting record to edit
'----------------------------------------
Set rsMandate = RunSP(SpCon, "usp_Sis010", 1, CLng(gblFileKey))

iOpenBank = True
Set rsBank = rsMandate.NextRecordset
'--------------------
iOpenMan = True
If rsMandate.EOF = True Then
    iMode = 0
    Me.Caption = "New Mandate"
Else
 iMode = 1
  Me.Caption = "Edit Mandate"
End If

 '--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS010/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit

End Sub
Private Sub UpdateScreen()
Dim bm As Variant, i As Integer
dbc_InitColumnProps (0)
tbfld(0).Text = gblFileKey
tbfld(1).Text = frmSIS001.grd.Columns(0).Text
'--
If iMode = 0 Then
   meb(0).Text = Format(Date, "dd-mmm-yyyy")
End If
If iMode = 1 Then
  With rsMandate
      If CurRec Then
         meb(0).Text = !MndStaDte
         If IsNothing(!BankId) Then
            dbc(0).Text = "NONE"
            If Not IsNothing(!MNDNAME) Then tbfld(2).Text = !MNDNAME
            If Not IsNothing(!MNDADDR1) Then tbfld(3).Text = !MNDADDR1
            If Not IsNothing(!MNDADDR2) Then tbfld(4).Text = !MNDADDR2
            If Not IsNothing(!MNDADDR3) Then tbfld(5).Text = !MNDADDR3
            If Not IsNothing(!MNDADDR4) Then tbfld(6).Text = !MNDADDR4
            If Not IsNothing(!MNDADDR5) Then tbfld(7).Text = !MNDADDR5
         Else
            dbc(0).MoveFirst
            For i = 0 To dbc(0).Rows - 1
              bm = dbc(0).GetBookmark(i)
              If dbc(0).Columns(1).CellText(bm) = Trim(!BankId) Then
                dbc(0).Bookmark = dbc(0).GetBookmark(i)
                dbc(0) = dbc(0).Columns(0).CellText(bm)
                tbfld(2).Text = dbc(0).Columns(0).CellText(bm)
                Exit For
              End If
            Next i
            DisableAddr
         End If
         '--
         If Not IsNothing(!mndacnt) Then tbfld(8).Text = !mndacnt
         If Not IsNothing(!MNDACNTNME) Then tbfld(9).Text = !MNDACNTNME
         '--
         dbc(1).MoveFirst
         For i = 0 To dbc(1).Rows - 1
           bm = dbc(1).GetBookmark(i)
           If dbc(1).Columns(1).CellText(bm) = !MNDMET Then
                dbc(1).Bookmark = dbc(1).GetBookmark(i)
                dbc(1) = dbc(1).Columns(0).CellText(bm)
                Exit For
           End If
         Next i
         '--
         DisableData
         
      Else
         iMode = 2 ' no active Mandate
         meb(0).Enabled = True
         meb(0).Text = Format(Date, "dd-mmm-yyyy")
      End If
   End With
End If
Exit Sub
End Sub
Private Sub ClearScreen()

  For X = 2 To 9
    tbfld(X).Text = ""
  Next
  '--
  For X = 0 To 1
    If meb(X).Enabled = True Then
      meb(X).Mask = ""
      meb(X).Text = ""
    End If
  Next
  '--
  If iMode = 1 Then
     UpdateScreen
     meb(1).SetFocus
  Else
     meb(0).SetFocus
  End If
End Sub


Private Sub Shutdown()
If SpCon.State = 1 Then
If iOpenMan = True Then rsMandate.Close
If iOpenBank = True Then rsBank.Close
End If
Set rsMandate = Nothing
End Sub
Private Function CurRec()
CurRec = False
'rsMandate.Requery
With rsMandate
  If .EOF Then Exit Function
  .MoveFirst
  While Not .EOF
    If IsNull(!mndenddte) Then
       CurRec = True
       Exit Function
    Else
       .MoveNext
    End If
  Wend
End With
End Function

Private Sub EnableEndDte()
meb(1).Visible = True
lblLabels(3).Visible = True
End Sub

Private Sub DisableData()
Dim X As Integer
meb(0).Enabled = False
For X = 2 To 9
  tbfld(X).Enabled = False
Next
dbc(0).Enabled = False
dbc(1).Enabled = False
EnableEndDte
End Sub

Private Sub DisableEndDte()
meb(1).Visible = False
lblLabels(3).Visible = False
End Sub

Private Sub EnableData()
Dim X As Integer
meb(0).Enabled = True
For X = 2 To 9
  tbfld(X).Enabled = True
Next
dbc(0).Enabled = True
dbc(1).Enabled = False
DisableEndDte
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub meb_GotFocus(Index As Integer)
If iMode = 0 Then meb(Index).Mask = "##-???-####"
End Sub


Private Sub DisableAddr()
Dim i As Integer
For i = 2 To 7
  tbfld(i).Enabled = False
Next
dbc(1).Enabled = False
End Sub

Private Sub EnableAddr()
Dim i As Integer
For i = 2 To 7
  tbfld(i) = ""
  tbfld(i).Enabled = True
Next
dbc(1).Enabled = True
tbfld(2).SetFocus
End Sub

