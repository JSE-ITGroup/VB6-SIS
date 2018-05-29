VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmSISItemsMatch 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Match SIS Full Diivdend Payments"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   Icon            =   "FrmSISItemsMatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtRemaining 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   840
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   5200
      Picture         =   "FrmSISItemsMatch.frx":030A
      ScaleHeight     =   435
      ScaleWidth      =   405
      TabIndex        =   14
      Top             =   5520
      Width           =   465
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
      MouseIcon       =   "FrmSISItemsMatch.frx":074D
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   7920
      Width           =   1575
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
      Left            =   9120
      MouseIcon       =   "FrmSISItemsMatch.frx":0A57
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   7920
      Width           =   1575
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   7920
      Width           =   2775
   End
   Begin VB.Frame FmeSisItem 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SIS Item to Reconcile"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   10815
      Begin VB.TextBox TxtTransDate 
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtDesc 
         Height          =   375
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   5535
      End
      Begin VB.TextBox TxtAmount 
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
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
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
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
         Left            =   3840
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
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
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBBankItems 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   5175
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   3
      TagVariant      =   "Click the Bank item to be included in the matched items"
      BackColorOdd    =   12648447
      RowHeight       =   503
      Columns.Count   =   3
      Columns(0).Width=   2275
      Columns(0).Caption=   "Chq Date"
      Columns(0).Name =   "Chq Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Chq Amount"
      Columns(1).Name =   "Chq Amount"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).NumberFormat=   "#,###.00"
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Description"
      Columns(2).Name =   "Description"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   9128
      _ExtentY        =   6376
      _StockProps     =   79
      Caption         =   "Available Bank Items"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
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
   Begin SSDataWidgets_B.SSDBGrid SSDBMatchedItems 
      Height          =   3615
      Left            =   5760
      TabIndex        =   1
      ToolTipText     =   "Double-click to remove items from this list"
      Top             =   4200
      Width           =   5175
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   3
      BackColorOdd    =   16761024
      RowHeight       =   503
      Columns.Count   =   3
      Columns(0).Width=   2275
      Columns(0).Caption=   "Chq Date"
      Columns(0).Name =   "Chq Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Chq Amount"
      Columns(1).Name =   "Chq Amount"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).NumberFormat=   "#,###.00"
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Description"
      Columns(2).Name =   "Description"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   9128
      _ExtentY        =   6376
      _StockProps     =   79
      Caption         =   "Matched Bank Items"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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
   Begin SSDataWidgets_B.SSDBGrid SSDBPayments 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Click the item to reconcile"
      Top             =   0
      Width           =   7695
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      BackColorEven   =   16776960
      BackColorOdd    =   16777215
      RowHeight       =   503
      ExtraHeight     =   106
      Columns.Count   =   4
      Columns(0).Width=   2275
      Columns(0).Caption=   "Chq Date"
      Columns(0).Name =   "Chq Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Chq Amount"
      Columns(1).Name =   "Chq Amount"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).NumberFormat=   "#,###.00"
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Description"
      Columns(2).Name =   "Description"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Tax Amount"
      Columns(3).Name =   "Tax Amount"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   13573
      _ExtentY        =   4683
      _StockProps     =   79
      Caption         =   "List of outstanding full dividend/capital distributions"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Amount remaining to be reconciled"
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
      Left            =   7800
      TabIndex        =   16
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   13
      Top             =   8040
      Width           =   2055
   End
End
Attribute VB_Name = "FrmSISItemsMatch"
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
Set FrmSISItemsMatch = Nothing
SpCon.Close
End Sub
Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
Dim adoTolerance As ADODB.Recordset

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
Exit Sub

Err_Form_Activate:
MsgBox Err.Description, vbOKOnly, "Form Activation"
GoTo Exit_Form_Activate
End Sub

Private Sub SSDBPayments_DblClick()
On Error GoTo Err_SSDBPayments_DblClick
Dim StrSql As String

With SSDBPayments
     TxtTransDate = .Columns(0).Text
     TxtDesc = .Columns(2).Text
     TxtAmount = .Columns(1).Text
     TxtRemaining = .Columns(1).Text
End With
SSDBBankItems.Enabled = True
     
Exit_SSDBPayments_DblClick:
Exit Sub

Err_SSDBPayments_DblClick:
MsgBox Err.Description, vbOKOnly, "Error populating SIS item details"
GoTo Exit_SSDBPayments_DblClick
End Sub

Private Sub SSDBPayments_InitColumnProps()
Dim adoPayments As ADODB.Recordset
Dim StrSql As String

Set adoPayments = RunSP(SpCon, "usp_GetDividends", 1)
   
Do While Not adoPayments.EOF
   With SSDBPayments
        StrSql = Format(adoPayments!ChqDate, "dd-mmm-yyyy") & vbTab & Format(adoPayments!PaymentAmt, "#,###.00") & vbTab
        StrSql = StrSql & adoPayments!Narration & vbTab & Format(adoPayments!TaxAmount, "#,###.00")
        .AddItem StrSql
   End With
   adoPayments.MoveNext
Loop

adoPayments.Close
Set adoPayments = Nothing

End Sub

Private Sub SSDBBankItems_InitColumnProps()
Dim adoBankItems As ADODB.Recordset
Dim StrSql As String

Set adoBankItems = RunSP(SpCon, "usp_GetBankCredits", 1)
   
Do While Not adoBankItems.EOF
   With SSDBBankItems
        StrSql = Format(adoBankItems!TranDate, "dd-mmm-yyyy") & vbTab & Format(adoBankItems!ChqAmt, "#,###.00") & vbTab
        StrSql = StrSql & adoBankItems!TranDesc
        .AddItem StrSql
   End With
   adoBankItems.MoveNext
Loop

adoBankItems.Close
Set adoBankItems = Nothing

End Sub
Private Sub SSDBBankItems_DblClick()
On Error GoTo Err_SSDBBankItems_DblClick
Dim StrSql As String

With SSDBBankItems
     StrSql = Format(.Columns(0).Text, "dd-mmm-yyyy") & vbTab & .Columns(1).Text & vbTab
     StrSql = StrSql & .Columns(2).Text
     SSDBMatchedItems.AddItem StrSql
     WrkTotal = WrkTotal + CCur(.Columns(1).Text)
     TxtRemaining = Format(CCur(TxtRemaining) - CCur(.Columns(1).Text), "#,###.00")
     TxtTotal = Format(WrkTotal, "#,###.00")
     .RemoveItem .AddItemRowIndex(.Bookmark)
End With
If (TxtAmount - WrkTotal) <= WrkTolerance Then
        CmdMatch.Enabled = True
     Else
        CmdMatch.Enabled = False
End If
     
Exit_SSDBBankItems_DblClick:
Exit Sub

Err_SSDBBankItems_DblClick:
MsgBox Err.Description, vbOKOnly, "Error populating Matched items"
GoTo Exit_SSDBBankItems_DblClick
End Sub

Private Sub SSDBMatchedItems_DblClick()
On Error GoTo Err_SSDBMatchedItems_DblClick
Dim StrSql As String

With SSDBMatchedItems
     WrkTotal = WrkTotal - CCur(.Columns(1).Text)
     TxtRemaining = Format(CCur(TxtRemaining) + CCur(.Columns(1).Text), "#,###.00")
     StrSql = Format(.Columns(0).Text, "dd-mmm-yyyy") & vbTab & .Columns(1).Text & vbTab
     StrSql = StrSql & .Columns(2).Text
     SSDBBankItems.AddItem StrSql
     .RemoveItem .AddItemRowIndex(.Bookmark)
     TxtTotal = Format(WrkTotal, "#,###.00")
End With
If (TxtAmount - WrkTotal) <= WrkTolerance Then
        CmdMatch.Enabled = True
     Else
        CmdMatch.Enabled = False
     End If
Exit_SSDBMatchedItems_DblClick:
Exit Sub

Err_SSDBMatchedItems_DblClick:
MsgBox Err.Description, vbOKOnly, "Matched Pane error"
GoTo Exit_SSDBMatchedItems_DblClick

End Sub

Private Sub CmdMatch_Click()
On Error GoTo Err_CmdMatch_Click
Dim ItemIDs As String
Dim i As Integer

ItemIDs = ""
With SSDBMatchedItems
     .Redraw = False
     .MoveFirst
     For i = 0 To .Rows - 1
            ItemIDs = ItemIDs & .Columns(3).Text & ";"
     Next i
       .Redraw = True

i = RunSP(SpCon, "usp_MatchSISItems", 0, TxtAccountNo, ItemIDs)
End With
If i = 0 Then
   MsgBox "Item(s) successfully matched", vbOKOnly, "SIS Items Matched"
   SSDBMatchedItems.RemoveAll
   WrkTotal = 0
   TxtTotal = 0
   SSDBPayments.RemoveItem (SSDBPayments.Row)
   SSDBBankItems.Enabled = False
Else
   MsgBox "There was a problem which prevented the completion of matching", vbOKOnly, "SIS Items Matching failed"
   GoTo Exit_CmdMatch_Click
End If

Exit_CmdMatch_Click:
Exit Sub

Err_CmdMatch_Click:
MsgBox Err.Description, vbOKOnly, "Error on Match Command"
GoTo Exit_CmdMatch_Click
End Sub

