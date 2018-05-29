VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3A32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3B32.OCX"
Begin VB.Form SDIUsers 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   3210
   ClientLeft      =   510
   ClientTop       =   1920
   ClientWidth     =   5520
   Icon            =   "SDIUsers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3210
   ScaleWidth      =   5520
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
      DataFieldList   =   "USERLEVEL"
      AllowInput      =   0   'False
      _Version        =   196616
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "User Level"
      Columns(0).Name =   "DESCRIPTION"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "USERLEVEL"
      Columns(1).Name =   "USERLEVEL"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "DESCRIPTION"
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   300
      Left            =   3600
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   4680
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   2520
      TabIndex        =   10
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   0
      Top             =   40
      Width           =   855
   End
   Begin VB.Frame OptGrp 
      Height          =   855
      Left            =   2040
      TabIndex        =   13
      Top             =   1560
      Width           =   975
      Begin SSDataWidgets_A.SSDBOptSet OptBtn 
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   735
         _Version        =   196611
         _ExtentX        =   1296
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "No"
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IndexSelected   =   1
         NumberOfButtons =   2
         Buttons.Button(0).OptionValue=   "-1"
         Buttons.Button(0).Caption=   "Yes"
         Buttons.Button(0).Mnemonic=   89
         Buttons.Button(0).Enabled=   0   'False
         Buttons.Button(0).TextLeft=   15
         Buttons.Button(0).TextRight=   33
         Buttons.Button(0).TextBottom=   14
         Buttons.Button(0).ButtonRight=   13
         Buttons.Button(0).ButtonBottom=   13
         Buttons.Button(0).PictureLeft=   35
         Buttons.Button(0).PictureRight=   34
         Buttons.Button(0).PictureBottom=   14
         Buttons.Button(0).ButtonToColRight=   48
         Buttons.Button(0).ButtonToColBottom=   14
         Buttons.Button(0).ButtonBitmapID=   4
         Buttons.Button(1).OptionValue=   "0"
         Buttons.Button(1).Caption=   "No"
         Buttons.Button(1).Mnemonic=   78
         Buttons.Button(1).Value=   -1  'True
         Buttons.Button(1).Enabled=   0   'False
         Buttons.Button(1).TextLeft=   15
         Buttons.Button(1).TextTop=   16
         Buttons.Button(1).TextRight=   29
         Buttons.Button(1).TextBottom=   30
         Buttons.Button(1).ButtonTop=   16
         Buttons.Button(1).ButtonRight=   13
         Buttons.Button(1).ButtonBottom=   29
         Buttons.Button(1).PictureLeft=   31
         Buttons.Button(1).PictureTop=   16
         Buttons.Button(1).PictureRight=   30
         Buttons.Button(1).PictureBottom=   30
         Buttons.Button(1).ButtonToColTop=   16
         Buttons.Button(1).ButtonToColRight=   48
         Buttons.Button(1).ButtonToColBottom=   30
         Buttons.Button(1).ButtonBitmapID=   5
      End
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Is Logged On?:"
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
      Index           =   4
      Left            =   0
      TabIndex        =   14
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Password:"
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
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Name:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "User's Access Level:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1225
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "System Name:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "SDIUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iErr As Integer, sMsg As String
Dim cnn As New cADOAccess
Dim rsUserLevel As New ADODB.Recordset
Dim rsUsers As New ADODB.Recordset

Function IsValid() As Integer

Dim rsB As Recordset, strFind As String
Set rsB = datDataUsr.Recordset
IsValid = True
If txtFields(0) = "" Then  ' System Name
   'msg = "Entry required for System Name"
   iErr = 1
   txtFields(0).SetFocus
   GoTo Validate_Err
 End If
 If txtFields(2) = "" Then ' Password
   'msg = "Entry Required for User's Password"
   iErr = 2
   txtFields(2).SetFocus
   GoTo Validate_Err
 End If
 If dbc(0) = "" Then ' User Level
   'msg = "Entry Required for User's Access Level"
   iErr = 6
   dbc(0).SetFocus
   GoTo Validate_Err
 End If
 If txtFields(1) = "" Then  ' User Name
   'msg = "Entry required for User's Real Name"
   iErr = 7
   txtFields(1).SetFocus
   GoTo Validate_Err
 End If
 
Validate_Exit:
   Exit Function

Validate_Err:
  'MsgBox msg, vbInformation, "Users"
  csvShowUsrErr iErr, "Users"
  IsValid = False
  GoTo Validate_Exit

End Function
Private Sub cmdAdd_Click()

  datDataCtl.Recordset.AddNew
  
End If
End Sub


Private Sub cmdFind_Click()
 On Error GoTo FindErr

  Dim sBookMark As String, sFidNme As String
  Dim sFindStr As String, sFindUser As String
   
  sFidNme = "[SystemName] = "

  If datDataCtl.Recordset.Type = dbOpenTable Then
    sFindStr = InputBox("Enter System Name:")
  Else
    sFindUser = InputBox("Enter System Name:")
    sFindStr = sFidNme & "'" & sFindUser & "'"
      
  End If
  If Len(sFindStr) = 0 Then Exit Sub

  If datDataCtl.Recordset.RecordCount > 0 Then
    sBookMark = datDataCtl.Recordset.Bookmark
  End If

  If datDataCtl.Recordset.Type = dbOpenTable Then
    datDataCtl.Recordset.Seek "=", sFindStr
  Else
    
    datDataCtl.Recordset.FindFirst sFindStr
    
  End If

  'return to old record if no match was found
  If datDataCtl.Recordset.NoMatch And Len(sBookMark) > 0 Then
    datDataCtl.Recordset.Bookmark = sBookMark
  End If
 
  Exit Sub

FindErr:
  Call csvShowError("Users")
  Exit Sub
End Sub

Private Sub cmdRefresh_Click()
  'this is really only needed for multi user apps
  datDataCtl.Refresh
End Sub

Private Sub cmdUpdate_Click()
  If IsValid() Then
   datDataCtl.UpdateRecord
   datDataCtl.Recordset.Bookmark = datDataCtl.Recordset.LastModified
  End If
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub datDatactl_Error(DataErr As Integer, Response As Integer)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
    
  csvShowError ("Users")
  Response = 0  'throw away the error
End Sub

Private Sub datDatactl_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position
  'for dynasets and snapshots
 datDataCtl.Caption = "Record: " & (datDataCtl.Recordset.AbsolutePosition + 1)
  'for the table object you must set the index property when
  'the recordset gets created and use the following line
  'datdatactl.Caption = "Record: " & (datdatactl.Recordset.RecordCount * (datdatactl.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub datDataCtl_Validate(Action As Integer, Save As Integer)
  'This is where you put validation code
  'This event gets called when the following actions occur
  
Dim x As Integer
    
  Select Case Action
    Case vbDataActionMoveFirst
      If Not IsValid() Then
         Action = 0
      End If
    Case vbDataActionMovePrevious
      If Not IsValid() Then
         Action = 0
      End If
    Case vbDataActionMoveNext
      If Not IsValid() Then
         Action = 0
         MsgBox IsValid
      End If
    Case vbDataActionMoveLast
      If Not IsValid() Then
         Action = 0
      End If
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionClose
  End Select

Validate_Exit:
   Exit Sub

End Sub

Private Sub Form_Load()
   Dim qSql As String
   '--
   csvCenterForm Me, gblMDIFORM
   Set rsUserLevel = New ADODB.Recordset
   qSql = "SELECT UserLevel & Chr(9) & Description as Rowinfo "
   qSql = qSql & "FROM USERLEVELS"
   '--
   On Error GoTo Form_Load_Err:
   rsUserLevel.Open qSql, gblFileName, , , adCmdText
   With rsUserLevel
      dbc(0).RemoveAll
      .MoveFirst
      Do While Not .EOF
         dbc(0).AddNew rowinfo
         .MoveNext
      Wend
      .Close
   End With
   
      
    ' ready message
   frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
   Screen.MousePointer = vbDefault
   frmMDI.txtStatusMsg.Refresh
End Sub


