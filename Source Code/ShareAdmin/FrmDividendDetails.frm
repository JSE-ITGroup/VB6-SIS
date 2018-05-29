VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmDividendDetails 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Payment Summary"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   Icon            =   "FrmDividendDetails.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_B.SSDBGrid SSDBSummary 
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   11535
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   6
      RowHeight       =   503
      ExtraHeight     =   212
      Columns.Count   =   6
      Columns(0).Width=   3200
      Columns(0).Caption=   "Payment Date"
      Columns(0).Name =   "Payment Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Register"
      Columns(1).Name =   "Register"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "No. of Shareholders"
      Columns(2).Name =   "No. of Shareholders"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).NumberFormat=   "#,###"
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Shares"
      Columns(3).Name =   "Shares"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).NumberFormat=   "#,###"
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "Payment Total"
      Columns(4).Name =   "Payment Total"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   9
      Columns(4).NumberFormat=   "#,###.##"
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "Taxes"
      Columns(5).Name =   "Taxes"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   9
      Columns(5).NumberFormat=   "#,###.##"
      Columns(5).FieldLen=   256
      _ExtentX        =   20346
      _ExtentY        =   8070
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin SSDataWidgets_B.SSDBCombo SSDBPayments 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2415
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
      Columns(0).Width=   3200
      Columns(0).Caption=   "Payment Date"
      Columns(0).Name =   "Payment Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   4260
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
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Select Payment:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "FrmDividendDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoRst As ADODB.Recordset
Dim SpCon As ADODB.Connection

Private Sub CmdExit_Click()
Unload Me
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
   frmMDI.txtStatusMsg.Refresh
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
frmMDI.txtStatusMsg.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmDividendDetails = Nothing

SpCon.Close
End Sub

Private Sub SSDBPayments_Click()
On Error GoTo Err_SSDBPayments_Click
Dim StrSql As String
Dim i As Integer

If IsEmpty(SSDBPayments.SelBookmarks(0)) Then
   MsgBox "Select a payment date"
   SSDBPayments.SetFocus
   GoTo Exit_SSDBPayments_Click
End If


Set adoRst = RunSP(SpCon, "usp_PayementAndTaxesSummary", 1, Format(SSDBPayments.Columns(0).Text, "dd-mmm-yyyy"))
If adoRst.EOF Then
   MsgBox "Sorry, no records matching your criteria were found"
   GoTo Exit_SSDBPayments_Click
End If
With SSDBSummary
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = ""
     For i = 0 To 5
         If i = 0 Then
            StrSql = Format(adoRst(i), "dd-mmm-yyyy") & vbTab
         Else
            If i = 2 Or i = 3 Then
               StrSql = StrSql & Format(adoRst(i), "#,##0") & vbTab
            Else
               If i = 4 Or i = 5 Then
                  StrSql = StrSql & Format(adoRst(i), "#,##0.00") & vbTab
               Else
                  StrSql = StrSql & adoRst(i) & vbTab
               End If
            End If
         End If
         Next i
     .AddItem StrSql
     adoRst.MoveNext
Loop
End With
adoRst.Close
Set adoRst = Nothing

Exit_SSDBPayments_Click:
Exit Sub
Err_SSDBPayments_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error loading payment summary"
Resume Exit_SSDBPayments_Click

End Sub

Private Sub SSDBPayments_InitColumnProps()
On Error GoTo Err_SSDBPayments_InitColumnProps
Dim StrSql As String

Set adoRst = RunSP(SpCon, "usp_PayementAndTaxesDates", 1)
If adoRst.EOF Then
   MsgBox "No payments were found"
   GoTo Exit_SSDBPayments_InitColumnProps
End If

adoRst.MoveFirst
With SSDBPayments
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = Format(adoRst(0), "dd-mmm-yyyy")
     .AddItem StrSql
     adoRst.MoveNext
     StrSql = ""
     Loop
End With
adoRst.Close
Set adoRst = Nothing

Exit_SSDBPayments_InitColumnProps:
Exit Sub

Err_SSDBPayments_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on loading payments dates"
Resume Exit_SSDBPayments_InitColumnProps
End Sub
