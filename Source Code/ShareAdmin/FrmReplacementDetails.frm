VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmReplacementDetails 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Replacement Cheque Details"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10545
   Icon            =   "FrmReplacementDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_B.SSDBGrid SSDBDetailList 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10560
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      BackColorEven   =   12640511
      BackColorOdd    =   8454143
      RowHeight       =   503
      Columns.Count   =   4
      Columns(0).Width=   3200
      Columns(0).Caption=   "Bank Account"
      Columns(0).Name =   "Bank Account"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   7435
      Columns(1).Caption=   "Shareholder"
      Columns(1).Name =   "Chq No"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3836
      Columns(2).Caption=   "Original Amount"
      Columns(2).Name =   "Original Amount"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   9
      Columns(2).NumberFormat=   "9,999.99"
      Columns(2).FieldLen=   256
      Columns(3).Width=   3281
      Columns(3).Caption=   "Original Chq No"
      Columns(3).Name =   "Original Chq No"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   18627
      _ExtentY        =   8070
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
End
Attribute VB_Name = "FrmReplacementDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection

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
Private Sub SSDBDetailList_Click()
SpCon.Close
Unload Me
End Sub

Private Sub SSDBDetailList_DblClick()
SpCon.Close
Unload Me
End Sub

Private Sub SSDBDetailList_InitColumnProps()
Dim adoRst As ADODB.Recordset
Dim sRowinfo As String

Set adoRst = RunSP(SpCon, "usp_ReplacementDetails", 1, CInt(gblFileKey))

With adoRst
      SSDBDetailList.RemoveAll
      If Not .EOF Then
        Do While Not .EOF
          sRowinfo = !AccountNo & vbTab & !CliName & vbTab & Format(!ChqAmt, "#,###.00") & vbTab & !ChqNum
          SSDBDetailList.AddItem sRowinfo
         .MoveNext
        Loop
      End If
End With
End Sub
