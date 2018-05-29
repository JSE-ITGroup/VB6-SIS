VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmBankItemsMatch 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Match Bank Items"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14565
   Icon            =   "FrmBankItemsMatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   14565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdStart 
      BackColor       =   &H8000000A&
      Caption         =   "Search"
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
      Left            =   12720
      MouseIcon       =   "FrmBankItemsMatch.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame FmeMatched 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Matched Items"
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
      Height          =   4215
      Left            =   7680
      TabIndex        =   21
      Top             =   2760
      Width           =   6855
      Begin SSDataWidgets_B.SSDBGrid SSDBMatched 
         Height          =   3975
         Left            =   0
         TabIndex        =   22
         Top             =   240
         Width           =   6855
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   6
         BackColorEven   =   16761087
         BackColorOdd    =   16744703
         RowHeight       =   423
         Columns.Count   =   6
         Columns(0).Width=   2064
         Columns(0).Caption=   "Chq No"
         Columns(0).Name =   "Chq No"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2037
         Columns(1).Caption=   "Chq Date"
         Columns(1).Name =   "Chq Date"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2646
         Columns(2).Caption=   "Chq Amount"
         Columns(2).Name =   "Chq Amount"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).NumberFormat=   "#,###.00"
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Caption=   "Name"
         Columns(3).Name =   "Name"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   1323
         Columns(4).Caption=   "ClientID"
         Columns(4).Name =   "ClientID"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   2090
         Columns(5).Caption=   "Payment Type"
         Columns(5).Name =   "Payment Type"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         _ExtentX        =   12091
         _ExtentY        =   7011
         _StockProps     =   79
         Caption         =   "Matched SIS Items"
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
   Begin VB.Frame FmeReconType 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Select Match Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   14535
      Begin VB.TextBox TxtSearchAmt 
         Height          =   375
         Left            =   7800
         TabIndex        =   27
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox TxtSearchChqNo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7800
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBReconTypes 
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   3255
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
         Columns.Count   =   4
         Columns(0).Width=   714
         Columns(0).Caption=   "ID"
         Columns(0).Name =   "ID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3122
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   6562
         Columns(2).Caption=   "Description"
         Columns(2).Name =   "Description"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1667
         Columns(3).Caption=   "Entry Type"
         Columns(3).Name =   "Entry Type"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         _ExtentX        =   5741
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
         DataFieldToDisplay=   "Column 2"
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBList 
         Height          =   375
         Left            =   2760
         TabIndex        =   24
         Top             =   720
         Width           =   2655
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
         BackColorEven   =   14737632
         BackColorOdd    =   16761024
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Chq Date"
         Columns(0).Name =   "Chq Date"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Dividend Type"
         Columns(1).Name =   "Dividend Type"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 0"
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enter Amount:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5880
         TabIndex        =   29
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enter Cheque No:"
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
         Left            =   5880
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Match Type"
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
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Select Dividend Date:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.TextBox TxtTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   7200
      Width           =   3255
   End
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
      Left            =   12600
      MouseIcon       =   "FrmBankItemsMatch.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton CmdMatch 
      Caption         =   "Match"
      Enabled         =   0   'False
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
      MouseIcon       =   "FrmBankItemsMatch.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Frame FmeGL 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Possible Matching SIS Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   7455
      Begin SSDataWidgets_B.SSDBGrid SSDBResults 
         Height          =   3975
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   7335
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   6
         BackColorEven   =   12648447
         BackColorOdd    =   12648384
         RowHeight       =   423
         Columns.Count   =   6
         Columns(0).Width=   2064
         Columns(0).Caption=   "Chq No"
         Columns(0).Name =   "Chq No"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "Chq Date"
         Columns(1).Name =   "Chq Date"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Caption=   "Chq Amount"
         Columns(2).Name =   "Chq Amount"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).NumberFormat=   "#,###.00"
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Caption=   "Name"
         Columns(3).Name =   "Name"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   1270
         Columns(4).Caption=   "ClientID"
         Columns(4).Name =   "ClientID"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   2117
         Columns(5).Caption=   "Payment Type"
         Columns(5).Name =   "Payment Type"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         _ExtentX        =   12938
         _ExtentY        =   7011
         _StockProps     =   79
         Caption         =   "Search Results"
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
   Begin VB.Frame FmeBankItem 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Bank Item Details"
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.TextBox TxtAccountNo 
         Height          =   375
         Left            =   9840
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtItemID 
         Height          =   375
         Left            =   9720
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtDrCr 
         Height          =   375
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxtAmount 
         Height          =   375
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox TxtChqNo 
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxtDesc 
         Height          =   375
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   5535
      End
      Begin VB.TextBox TxtTransDate 
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Dr/Cr:"
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
         Left            =   7560
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Amount:"
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
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Chq No.:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Description:"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Trans Date:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Matched Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   7320
      Width           =   2895
   End
End
Attribute VB_Name = "FrmBankItemsMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim WrkTotal As Currency
Dim WrkTolerance As Currency
Dim Opt As Integer
Dim ReconType As Integer

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdMatch_Click()
On Error GoTo Err_CmdMatch_Click
Dim ChqNos As String
Dim ChqAmts As String
Dim i As Integer
Dim PayTyp As String

ChqNos = ""
ChqAmts = ""
With SSDBMatched
     .Redraw = False
     .MoveFirst
     For i = 0 To .Rows - 1
            ChqNos = ChqNos & .Columns(0).Text & ";"
            ChqAmts = ChqAmts & .Columns(2).Text & ";"
     Next i
       .Redraw = True

PayTyp = Mid(SSDBMatched.Columns(5).Text, 1, 1)
i = RunSP(SpCon, "usp_MatchItems", 0, Opt, TxtAccountNo, TxtItemID, ChqNo, .Columns(1).Text, PayTyp)
End With
If i = 0 Then
   MsgBox "Item(s) successfully matched", vbOKOnly, "Bank Items Matched"
   SSDBMatched.RemoveAll
   SSDBResults.RemoveAll
Else
   MsgBox "There was a problem which prevented the completion of matching", vbOKOnly, "Bank Items Matching failed"
   GoTo Exit_CmdMatch_Click
End If

Exit_CmdMatch_Click:
Exit Sub

Err_CmdMatch_Click:
MsgBox Err.Description, vbOKOnly, "Error on Match Command"
GoTo Exit_CmdMatch_Click
End Sub

Private Sub CmdStart_Click()
On Error GoTo Err_CmdStart_Click
Dim i As Integer
Dim AccountNo As String
Dim Amt As Currency
Dim SDate As String
Dim adoRst As ADODB.Recordset

If SSDBReconTypes = "" Then
   ReconType = 0
Else
   ReconType = CInt(SSDBReconTypes.Columns(0).Text)
End If

If SSDBList = "ALL DATES" Or SSDBList = vbNullString Then
   SDate = vbNullString
Else
  SDate = Format(SSDBList.Columns(0).Text, "dd-mmm-yyyy")
End If

If Len(TxtSearchChqNo) > 1 Then
    TxtSearchChqNo = Trim(TxtSearchChqNo)
Else
   TxtSearchAmt = TxtAmount
   Amt = CCur(TxtAmount)
End If

If TxtSearchAmt = "0" Or TxtSearchAmt = vbNullString Then
   Amt = 0
Else
   Amt = CCur(TxtSearchAmt)
   TxtSearchChqNo = "0"
End If
   
AccountNo = TxtAccountNo

Set adoRst = RunSP(SpCon, "usp_SelectReconItems", 1, ReconType, SDate, CCur(Amt), AccountNo, TxtSearchChqNo)
If adoRst.State = adStateClosed Then
   MsgBox "No records were found which matched your criteria", vbOKOnly, "Search Results"
   GoTo Exit_CmdStart_Click
End If
If adoRst.EOF Then
   MsgBox "No records were found which matched your criteria", vbOKOnly, "Search Results"
   GoTo Exit_CmdStart_Click
End If
With SSDBResults
     .RemoveAll
     Do While Not adoRst.EOF
        StrSql = adoRst(0) & vbTab & Format(adoRst(1), "dd-mmm-yyyy") & vbTab
        StrSql = StrSql & adoRst(2) & vbTab
        StrSql = StrSql & adoRst(3) & vbTab
        StrSql = StrSql & adoRst(4) & vbTab & adoRst(5)
        .AddItem StrSql
        adoRst.MoveNext
     Loop
End With
adoRst.Close
Set adoRst = Nothing

Exit_CmdStart_Click:
Exit Sub

Err_CmdStart_Click:
MsgBox Err.Description, vbOKOnly, "Error on Search"
GoTo Exit_CmdStart_Click

End Sub

Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
Dim adoRst As ADODB.Recordset
Dim adoTolerance As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_RetrieveBankItem", 1, CInt(gblFileKey))
If adoRst.EOF Then
   MsgBox "Record not found", vbOKOnly
   Unload Me
   GoTo Exit_Form_Activate
End If
With adoRst
     TxtTransDate = Format(!TranDate, "dd-mmm-yyyy")
     TxtDesc = !TranDesc
     TxtChqno = !ChqNo
     TxtAmount = Format(!ChqAmt, "#,###.00")
     TxtItemID = !ItemID
     If !TranType = "D" Then
        TxtDrCr = "Dr"
     Else
        TxtDrCr = "Cr"
     End If
     TxtAccountNo = !AccountNo
End With
adoRst.Close
WrkTotal = 0
Set adoTolerance = RunSP(SpCon, "usp_SelectTolerance", 1)
If adoTolerance.EOF Then
   MsgBox "Tolerance amount not found. Please set up one"
   Unload Me
   GoTo Exit_Form_Activate
End If
WrkTolerance = adoTolerance!Amount
adoTolerance.Close
Set adoTolerance = Nothing

Exit_Form_Activate:
Set adoRst = Nothing
Exit Sub

Err_Form_Activate:
MsgBox Err.Description, vbOKOnly, "Form Activation"
GoTo Exit_Form_Activate
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
Set FrmBankItemsMatch = Nothing
SpCon.Close
FrmBankItems.Show 0
End Sub

Private Sub SSDBList_InitColumnProps()
Dim adoList As ADODB.Recordset
Dim StrSql As String

Set adoList = RunSP(SpCon, "usp_SelectPostedDividend", 1)

StrSql = "ALL DATES" & vbTab & "All Payment Dates"
SSDBList.AddItem StrSql
   
Do While Not adoList.EOF
   With SSDBList
        StrSql = Format(adoList!ChqDate, "dd-mmm-yyyy") & vbTab
        If adoList!PayTyp = "D" Then
           StrSql = StrSql & "Dividend"
        Else
           StrSql = StrSql & "Capital Distribution"
        End If
        .AddItem StrSql
   End With
   adoList.MoveNext
Loop

adoList.Close
Set adoList = Nothing

End Sub

Private Sub SSDBMatched_DblClick()
On Error GoTo Err_SSDBMatched_DblClick
Dim StrSql As String

With SSDBMatched
     WrkTotal = WrkTotal - CCur(.Columns(2).Text)
     .RemoveItem (.Row)
     TxtTotal = Format(WrkTotal, "#,###.00")
End With
If (TxtAmount - WrkTotal) <= WrkTolerance Then
        CmdMatch.Enabled = True
     Else
        CmdMatch.Enabled = False
     End If
Exit_SSDBMatched_DblClick:
Exit Sub

Err_SSDBMatched_DblClick:
MsgBox Err.Description, vbOKOnly, "Matched Pane error"
GoTo Exit_SSDBMatched_DblClick

End Sub

Private Sub SSDBReconTypes_InitColumnProps()
Dim adoRst As ADODB.Recordset
Dim StrSql As String

Set adoRst = RunSP(SpCon, "usp_SelectReconTypes", 1)
If adoRst.EOF Then
   MsgBox "Reconciliation types were not setup. Contact your Administrator", vbOKOnly
   GoTo Exit_SSDBReconTypes_InitColumnProps
End If
With adoRst
     StrSql = "0" & vbTab & "All" & vbTab & "All Types" & vbTab & "Cr/Dr"
     SSDBReconTypes.AddItem StrSql
     Do While Not .EOF
        StrSql = !ReconType & vbTab & !ReconDesc & vbTab
        StrSql = StrSql & !ReconNarr & vbTab
        If !EntryType = False Then
           StrSql = StrSql & "Dr"
        Else
           StrSql = StrSql & "Cr"
        End If
        SSDBReconTypes.AddItem StrSql
        .MoveNext
     Loop
End With
adoRst.Close

Exit_SSDBReconTypes_InitColumnProps:
Set adoRst = Nothing
Exit Sub

End Sub

Private Sub SSDBResults_DblClick()
On Error GoTo Err_SSDBResults_DblClick
Dim StrSql As String

With SSDBResults
     If SSDBMatched.Rows <> 0 Then
        If .Columns(1).Text <> SSDBMatched.Columns(1).Text Then
           MsgBox "Please select matching items with similiar cheque dates"
           GoTo Exit_SSDBResults_DblClick
        End If
        If CheckEntry(.Columns(0).Text, 0) Then
           MsgBox "This Cheque No was already included in the matched list", vbOKOnly, "Duplicated Cheque Number"
           GoTo Exit_SSDBResults_DblClick
        End If
     End If
     
     StrSql = .Columns(0).Text & vbTab & Format(.Columns(1).Text, "dd-mmm-yyyy") & vbTab
     StrSql = StrSql & .Columns(2).Text & vbTab & .Columns(3).Text & vbTab
     StrSql = StrSql & .Columns(4).Text
     SSDBMatched.AddItem StrSql
     WrkTotal = WrkTotal + CCur(.Columns(2).Text)
     TxtTotal = Format(WrkTotal, "#,###.00")
     .RemoveItem (.Row)
End With
If (TxtAmount - WrkTotal) <= WrkTolerance Then
        CmdMatch.Enabled = True
     Else
        CmdMatch.Enabled = False
End If
     
Exit_SSDBResults_DblClick:
Exit Sub

Err_SSDBResults_DblClick:
MsgBox Err.Description, vbOKOnly, "Results Pane error"
GoTo Exit_SSDBResults_DblClick
End Sub

Public Function CheckEntry(SearchText As String, ColNo As Integer)
Dim bm As Variant
Dim i As Integer

CheckEntry = False
SSDBMatched.Redraw = False
SSDBMatched.MoveFirst
For i = 0 To SSDBMatched.Rows - 1
    bm = SSDBMatched.GetBookmark(i)
    If SearchText = SSDBMatched.Columns(ColNo).CellText(bm) Then
       SSDBMatched.Bookmark = SSDBMatched.GetBookmark(i)
       CheckEntry = True
       Exit For
    End If
    Next i
SSDBMatched.Redraw = True
End Function
