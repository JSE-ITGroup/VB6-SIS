VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS026 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shareholder to Shareholder Transfer"
   ClientHeight    =   5940
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "SIS026.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7155
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   1215
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   6375
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   5
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   1191
      Columns(0).Caption=   "Line #"
      Columns(0).Name =   "Line #"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   2
      Columns(0).FieldLen=   2
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   4604
      Columns(1).Caption=   "Buying Shareholders"
      Columns(1).Name =   "Buying Shareholders"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   50
      Columns(2).Width=   2196
      Columns(2).Caption=   "Client Id"
      Columns(2).Name =   "Client Id"
      Columns(2).Alignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   10
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   2090
      Columns(3).Caption=   "Shares"
      Columns(3).Name =   "Shares"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   3
      Columns(3).FieldLen=   10
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "CertNO"
      Columns(4).Name =   "CertNO"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   3
      Columns(4).FieldLen=   11
      _ExtentX        =   11245
      _ExtentY        =   2143
      _StockProps     =   79
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
   Begin SSDataWidgets_B.SSDBDropDown dbdd 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   2487
      Columns(0).Caption=   "Cert #"
      Columns(0).Name =   "Cert #"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   11
      Columns(1).Width=   2593
      Columns(1).Caption=   "Shares"
      Columns(1).Name =   "Shares"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   12
      Columns(2).Width=   3200
      Columns(2).Caption=   "Issue Date"
      Columns(2).Name =   "Issue Date"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   7
      Columns(2).NumberFormat=   "dd-mmm-yyyy"
      Columns(2).FieldLen=   11
      _ExtentX        =   1085
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   1455
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   5475
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      AllowColumnMoving=   0
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   1111
      Columns(0).Caption=   "Line #"
      Columns(0).Name =   "Line #"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   2
      Columns(0).FieldLen=   2
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   2699
      Columns(1).Caption=   "Cert #"
      Columns(1).Name =   "Cert #"
      Columns(1).Alignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   10
      Columns(2).Width=   2302
      Columns(2).Caption=   "Issue Date"
      Columns(2).Name =   "Issue Date"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   7
      Columns(2).NumberFormat=   "dd-mmm-yyyy"
      Columns(2).FieldLen=   11
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   2461
      Columns(3).Caption=   "Shares"
      Columns(3).Name =   "Shares"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   3
      Columns(3).FieldLen=   11
      Columns(3).Locked=   -1  'True
      _ExtentX        =   9657
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Certificates to be Cancelled"
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
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   300
      Left            =   1800
      TabIndex        =   23
      ToolTipText     =   "Pressing this button will activate the search program to locate a shareholder."
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox tbFld 
      Height          =   285
      Index           =   1
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   2
      ToolTipText     =   "Assign a document number."
      Top             =   840
      Width           =   1500
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   3960
      TabIndex        =   20
      Top             =   5520
      Width           =   975
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Select from existing batches or enter a new batch number"
      Top             =   480
      Width           =   1815
      DataFieldList   =   "Column 0"
      AllowNull       =   0   'False
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2090
      Columns(0).Caption=   "Batch #"
      Columns(0).Name =   "Batch #"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   8
      Columns(1).Width=   2170
      Columns(1).Caption=   "Date"
      Columns(1).Name =   "Date"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd-mmm-yyyy"
      Columns(1).FieldLen=   11
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   2880
      TabIndex        =   14
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   6120
      TabIndex        =   8
      ToolTipText     =   "Cancels all processing and exits program."
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "Saves the screen information to the database."
      Top             =   5520
      Width           =   975
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   1
      ToolTipText     =   "Enter the date of the new batch"
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   3
      ToolTipText     =   "Enter the date of  issue you want to appear on the certificate.."
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "Select a share holder from the list."
      Top             =   1320
      Width           =   2895
      DataFieldList   =   "Column 1"
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
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   5636
      Columns(0).Caption=   "Client Name"
      Columns(0).Name =   "Client Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   50
      Columns(1).Width=   3200
      Columns(1).Caption=   "Client Id"
      Columns(1).Name =   "Client Id"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   10
      Columns(2).Width=   2487
      Columns(2).Caption=   "Shares"
      Columns(2).Name =   "Shares"
      Columns(2).Alignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   12
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
      Enabled         =   0   'False
      DataFieldToDisplay=   "Column 0"
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   24
      ToolTipText     =   "Enter the date of  issue you want to appear on the certificate.."
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   11
      Format          =   "#,###"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBDropDown dbdd 
      Height          =   735
      Index           =   1
      Left            =   0
      TabIndex        =   27
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   5001
      Columns(0).Caption=   "Client Name"
      Columns(0).Name =   "Client Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   50
      Columns(1).Width=   2143
      Columns(1).Caption=   "ClientId"
      Columns(1).Name =   "ClientId"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   11
      _ExtentX        =   1085
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Line Line7 
      X1              =   5040
      X2              =   5040
      Y1              =   1200
      Y2              =   2160
   End
   Begin VB.Label lbl 
      Caption         =   "Available Shares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   25
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Form No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   22
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Batch Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   21
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Joint Holder #1 Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   1920
      TabIndex        =   19
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblLabels 
      Caption         =   "Joint Holder #2 Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   1920
      TabIndex        =   18
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Holder #2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Batch No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint Holder #1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   7080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblLabels 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   0
      X2              =   6960
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "From Share Holder:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ver:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Transfer Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmSIS026"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim X As Integer, iEOF As Integer
Dim iNewApp As Integer, iNew As Integer
Dim iSellShares As Long, iBuyShares As Long
Dim rsCmp As ADODB.Recordset
Dim rsClient As ADODB.Recordset
Dim rsMain As ADODB.Recordset
Dim rsCert As ADODB.Recordset
Dim rsUnused As ADODB.Recordset
Dim rsJoint As ADODB.Recordset
Dim rsBat As ADODB.Recordset
Dim rsVerBat As ADODB.Recordset
Dim rsVerFrm As ADODB.Recordset
Dim cmdChange As ADODB.Command
Dim cmdDel As ADODB.Command
Dim errLoop As Error
Dim errs1 As Error
Dim strTable As String
Dim strRecNO As String
Dim iStocks As Double, icert As Long
Dim bm As Variant
Private Sub FillCombo()
Dim sRowinfo As String
With rsClient
    .Requery
    If Not .EOF And Not .BOF Then
      .MoveFirst
      dbc(1).Redraw = False
      dbc(1).RemoveAll
      dbc(1) = ""
      dbc(1).Redraw = True
      meb(2) = ""
      lblLabels(12).Caption = ""
      lblLabels(13).Caption = ""
      Do While Not .EOF
         sRowinfo = !CliName & Chr(9) & !ClientId
         sRowinfo = sRowinfo & Chr(9) & !shares
         dbc(1).AddItem sRowinfo
        'If dbc(1).Row = 0 Then dbc(1) = !CLINAME
         .MoveNext
      Loop
      'dbc(1).Redraw = True
     ' grd(0).Redraw = False
      grd(0).RemoveAll
      dbdd(0).RemoveAll
      'grd(0).Redraw = True
      
    End If
End With
End Sub
Function IsValid() As Integer
Dim iErr As Integer, dtefld As Date, qSQL
Dim sElable As String
sElable = "Stockholder to Stockholder Entry"
IsValid = False
iErr = 0
'--
If dbc(0) = "" Then ' batch
    iErr = 132
    csvShowUsrErr iErr, sElable
    dbc(0).SetFocus
    GoTo Validate_Exit
End If
dbc(0) = UCase(dbc(0))
'--
If meb(0) = "" Then 'batch date
   iErr = 139
   csvShowUsrErr iErr, sElable
   meb(0).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(0)) Then
      iErr = 14
      csvShowUsrErr iErr, sElable
      meb(0).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If tbfld(1).Text = "" Then 'Invalid form
   iErr = 140
   csvShowUsrErr iErr, sElable
   tbfld(1).SetFocus
   GoTo Validate_Exit
End If
tbfld(1).Text = UCase(tbfld(1).Text)
'--
If gblOptions = 1 Then 'check for duplicate form
  
  qSQL = "SELECT FORM from STKACTIV where TRNBATCH = '"
  qSQL = qSQL & dbc(0) & "' and "
  qSQL = qSQL & "FORM = '" & tbfld(1) & "'"
  rsVerFrm.Open qSQL, cnn, , , adCmdText
  If Not rsVerFrm.EOF Then
     iErr = 141
     csvShowUsrErr iErr, sElable
     tbfld(1).SetFocus
     rsVerFrm.Close
     GoTo Validate_Exit
  End If
  rsVerFrm.Close
End If
'--
If dbc(1) = "" Then   ' Sell Client name
   iErr = 130
   csvShowUsrErr iErr, sElable
   dbc(1).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If meb(1) = "" Then 'Transfer date
   iErr = 129
   csvShowUsrErr iErr, sElable
   meb(1).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(1)) Then
      iErr = 14
      csvShowUsrErr iErr, sElable
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If grd(0).Rows = 0 Then  ' empty grid no certs to cancel
   iErr = 143
   csvShowUsrErr iErr, sElable
   grd(0).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If grd(1).Rows = 0 Then  'empty grid no
   iErr = 146
   csvShowUsrErr iErr, sElable
   grd(1).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If iBuyShares <> 0 Then
    iErr = 147
    csvShowUsrErr iErr, sElable
   grd(1).SetFocus
   GoTo Validate_Exit
 End If
  
 IsValid = True
Validate_Exit:
   
   Exit Function
End Function

Private Sub cmdCancel_Click()
Dim iSeqKey As Integer
rsUnused.Close
rsCmp.Close
rsMain.Close
Set rsUnused = Nothing
Set rsMain = Nothing
Set rsCmp = Nothing
Set rsClient = Nothing
Set rsJoint = Nothing
Set rsVerBat = Nothing
Set rsVerFrm = Nothing
cnn.Close
'''set cnn = nothing
Unload Me
Set frmSIS026 = Nothing

End Sub

Private Sub cmdClear_Click()
ClearScreen
If gblOptions = 2 Then UpdateScreen
End Sub

Private Sub cmdDelete_Click()
Dim imsg As Integer, X, i As Integer
Dim qDMLQry As String, iStocks As Double
Dim bm As Variant
imsg = 133
On Error GoTo cmdDelete_Err
If csvYesNo(imsg, "Stockholder to Stockholder") Then
   
   '----------------------------------
   '-- delete STKACTIV transactions --
   '----------------------------------
   cnn.BeginTrans
   With rsMain
      .MoveFirst
      Do While Not .EOF
         If !certno > 0 Then
           rsUnused.AddNew
           rsUnused!SEQTYP = "C"
           rsUnused!UNUSED = !certno
           rsUnused.Update
         End If
         .Delete
         .MoveFirst
      Loop
   End With
   cnn.CommitTrans
   '---------------------------
   '-- Activate Cancelled Certs
   '---------------------------
   iStocks = 0
   With grd(0)
      .Redraw = False
      .MoveFirst
      For i = 0 To .Rows - 1
        bm = .GetBookmark(i)
        iStocks = iStocks + Val(.Columns(3).CellText(bm))
        qDMLQry = "UPDATE CERTMST SET STATUS = 'A' "
        qDMLQry = qDMLQry & "WHERE CERTNO = "
        qDMLQry = qDMLQry & Val(.Columns(1).CellText(bm))
        X = csvADODML(qDMLQry, cnn)
        'Set cmdChange = New ADODB.Command
        'Set cmdChange.ActiveConnection = cnn
        'cmdChange.CommandText = qDMLQry
        'cnn.Errors.Clear
        'csvExecuteCommand cmdChange
        'Set cmdChange = Nothing
      Next i
   End With
   '-------------------------------------
   '-- Update Selling Account with shares
   '-------------------------------------
   qDMLQry = "UPDATE STKNAME SET SHARES = SHARES + " & iStocks
   qDMLQry = qDMLQry & " WHERE CLIENTID = " & Val(dbc(1).Columns(1).Text)
   X = csvADODML(qDMLQry, cnn)
   'Set cmdChange = New ADODB.Command
   'Set cmdChange.ActiveConnection = cnn
   'cmdChange.CommandText = qDMLQry
   'cnn.Errors.Clear
   'csvExecuteCommand cmdChange
   'Set cmdChange = Nothing
   '-----------------------------
   '-- Delete buying certs from CERTMST
   '-----------------------------------
   With grd(1)
      .MoveFirst
      For i = 0 To .Rows - 1
        bm = .GetBookmark(i)
        qDMLQry = "DELETE FROM CERTMST WHERE CERTNO = "
        qDMLQry = qDMLQry & Val(.Columns(4).CellText(bm))
        X = csvADODML(qDMLQry, cnn)
        'Set cmdDel = New ADODB.Command
        'Set cmdDel.ActiveConnection = cnn
        'cmdDel.CommandText = qDMLQry
        'cnn.Errors.Clear
        'csvExecuteCommand cmdDel
        'Set cmdDel = Nothing
        '------------------------------
        '-- REDUCE BUYING CLIENT SHARES
        '------------------------------
        qDMLQry = "UPDATE STKNAME SET SHARES = SHARES - " & Val(.Columns(3).CellText(bm))
        qDMLQry = qDMLQry & " WHERE CLIENTID = " & Val(.Columns(2).CellText(bm))
        X = csvADODML(qDMLQry, cnn)
        'Set cmdChange = New ADODB.Command
        'Set cmdChange.ActiveConnection = cnn
        'cmdChange.CommandText = qDMLQry
        'cnn.Errors.Clear
        'csvExecuteCommand cmdChange
        'Set cmdChange = Nothing
       Next i
     '--
   End With
  
   cmdCancel_Click
cmdDelete_Exit:
Exit Sub

cmdDelete_Err:
 MsgBox "cmdDelete"
 GoTo cmdDelete_Exit
End If
End Sub

Private Sub cmdFind_Click()
Dim i As Integer, X As Integer
Dim cChk As Integer, qCli As String
Dim sWhere As String, sRowinfo As String

Load frmFind
  With frmFind
    '- load comparison key fields and show frmFind
    '---------------------------------------------
     .cbWhere.Clear
    .cbWhere.AddItem "CliName"
    .cbWhere.AddItem "ClientId"
    .cbWhere.ListIndex = 0
    .cbOptions.ListIndex = 0
    .lbl(3).Caption = "Find Client"
    .optBtn.Buttons(0).Caption = "&Selling"
    .optBtn.Buttons(1).Caption = "&Buying"
    .Show vbModal
    '----------------------------
    '-------- main line ---------
    '----------------------------
    If .tbFind.Text = vbNullString Then
    Else
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      sWhere = Trim(.tbFind.Text)
      X = .cbWhere.ListIndex
      qCli = "SELECT CLINAME,  CLIENTID, SHARES FROM STKNAME WHERE "
      qCli = qCli & "CATCODE <> 'SB' AND "
      qCli = qCli & .cbWhere
      '---
      If sWhere <> "" Then
          Select Case .cbOptions.ListIndex
          Case 0 ' Exact Match
              If X = 0 Then
                qCli = qCli & " = '" & sWhere & "'"
              Else
                qCli = qCli & " = " & Val(.tbFind.Text)
              End If
          Case 1 ' Starts With
              sWhere = Trim(.tbFind.Text) & "%"
              qCli = qCli & " like '" & sWhere & "'"
              qCli = qCli & " ORDER BY CLINAME, CLIENTID"
            Case 2 ' Ends With
               sWhere = "%" & Trim(.tbFind.Text)
               qCli = qCli & " like '" & sWhere & "'"
               qCli = qCli & " ORDER BY CLINAME, CLIENTID"
             Case 3 ' AnyWhere
               sWhere = "%" & Trim(.tbFind.Text) & "%"
               qCli = qCli & " like '" & sWhere & "'"
               qCli = qCli & " ORDER BY CLINAME, CLIENTID"
             End Select
      End If
      rsClient.Open qCli, cnn, , , adCmdText
      If Not rsClient.EOF Then
         If .optBtn.IndexSelected = 0 Then
           i = 1
           FillCombo ' Client selling
           grd(0).Enabled = True
           dbc(i).Enabled = True
           dbc(i).SetFocus
         Else
           i = 2   ' Client Buying Fill grd
           With rsClient
               dbdd(1).RemoveAll
               .MoveFirst
               Do While Not .EOF
                  sRowinfo = !CliName & Chr(9) & !ClientId
                  dbdd(1).AddItem sRowinfo
                  .MoveNext
               Loop
           End With
           grd(1).Enabled = True
           grd(1).SetFocus
           cmdUpdate.Enabled = True
         End If
      End If
      rsClient.Close
    End If
  End With
Unload frmFind
Set frmFind = Nothing
Exit Sub
cmdFind_Click_err:
  MsgBox "SIS026/CmdFind"
End Sub
Private Sub cmdUpdate_Click()
Dim qBat As String, iLines As Integer, iCrt As Long
Dim iClient As Long, iss As Date, iShares As Double
Dim strncode As String, strnbatch As String, StrnDate As Date
Dim i, iUpd As Integer, bm As Variant, iabort As Integer
On Error GoTo cmdUpdate_Err
' wait message
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.Refresh
iabort = False: iUpd = False
If IsValid Then
  '--
  cnn.BeginTrans
  '--
  iLines = 0
  iClient = dbc(1).Columns(1).Text
  iss = DateValue(meb(1).Text)
  strncode = "S"
  strnbatch = dbc(0)
  StrnDate = DateValue(meb(0).Text)
  '--
  With rsMain
  '----------------------------
  '-- Add From data to STKACTIV
  '-----------------------------
    grd(0).MoveFirst
    For i = 0 To grd(0).Rows - 1
       bm = grd(0).GetBookmark(i)
       .AddNew
       !TrnBatch = strnbatch
       !TRNDATE = StrnDate
       iLines = iLines + 1
       !stklineno = iLines
       !Form = tbfld(1).Text
       !ClientId = iClient
       iCrt = Val(grd(0).Columns(1).CellText(bm))
       !FRCERT = iCrt
       !CANDATE = !TRNDATE
       !FRSHARES = Val(grd(0).Columns(3).CellText(bm))
       !IssDate = DateValue(grd(0).Columns(2).CellText(bm))
       !TrnCode = strncode
       !Status = "O"
       .Update
       iUpd = True
       '-- cancel Certs & update Stkname
       If Not CancelCert(iClient, iCrt _
               , strncode, strnbatch, StrnDate, tbfld(1), cnn) Then
               cnn.RollbackTrans
               iabort = True
               Exit For
        End If
        Next
        '--
      If iabort = True Then GoTo Done
      '--------------------------------
      '-- Add To data to STKACTIV
      '--------------------------------
      grd(1).MoveFirst
      For i = 0 To grd(0).Rows - 1
          bm = grd(1).GetBookmark(i)
          iClient = Val(grd(1).Columns(2).CellText(bm))
          iShares = Val(grd(1).Columns(3).CellText(bm))
          icert = CreateCert(iClient, StrnDate, iShares _
             , strncode, strnbatch, StrnDate, tbfld(1), cnn)
          .AddNew
          !TrnBatch = strnbatch
          !TRNDATE = StrnDate
          iLines = iLines + 1
          !stklineno = iLines
          !Form = tbfld(1).Text
          !ClientId = iClient
          !certno = icert
          !IssDate = !TRNDATE
          !shares = iShares
          !TrnCode = strncode
          !Status = "O"
          .Update
          '--
      Next
  End With
  If iabort = False Then
       '-- update batch header if new batch
       '-----------
       If iNew Then
          With rsBat
             qBat = "SELECT BATCHNO FROM BATHDR WHERE BATCHNO = '"
             qBat = qBat & dbc(0) & "'"
             rsVerBat.Open qBat, cnn, , , adCmdText
             If rsVerBat.EOF Then .AddNew
             !BATCHNO = dbc(0)
             !BATDATE = DateValue(meb(0).Text)
             .Update
             rsVerBat.Close
          End With
        End If
        '--
        cnn.CommitTrans
        dbc_InitColumnProps (0)
        dbc(0).Enabled = False
        meb(0).Enabled = False
        InitAddNew
  End If
End If
'---
Done:
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
Exit Sub
'--
cmdUpdate_Err:
  If iUpd = True Then cnn.RollbackTrans
  MsgBox "SIS026/cmdUpdate"
  csvLogError "SIS026/cmdUpdate", Err.Number, Err.Description
  cmdCancel_Click
End Sub

Private Sub dbc_Click(Index As Integer)

meb(2).Text = dbc(Index).Columns("shares").Text

End Sub

Private Sub dbc_InitColumnProps(Index As Integer)
Dim sRowinfo As String
Select Case Index
Case 0 ' Load Open Batches
rsBat.Requery
With rsBat
  If Not .EOF And Not .BOF Then
     .MoveFirst
     dbc(0).RemoveAll
     Do While Not .EOF
       sRowinfo = !BATCHNO & Chr(9) & !BATDATE
       dbc(0).AddItem sRowinfo
       .MoveNext
     Loop
  End If
End With
Case Else
End Select
End Sub

Private Sub dbc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  Select Case Index
    Case 0
      If meb(0).Enabled = True Then
         meb(0).SetFocus
      Else
         tbfld(1).SetFocus
      End If
    Case 1
      If grd(0).Enabled = True Then
        grd(0).SetFocus
      Else
        cmdFind.SetFocus
      End If
    Case Else
  End Select
     
 Case vbKeyUp
  KeyCode = 0
  Select Case Index
    Case 1
       tbfld(1).SetFocus
    Case Else
  End Select
End Select
End Sub

Private Sub dbc_LostFocus(Index As Integer)
Dim qDMLQry As String, i As Integer
Dim sRowinfo As String
Select Case Index
Case 0
  '-----------------------------------------
  '-- get batch date if existing batch keyed
  '-- if not set focus to get date
  '-----------------------------------------
  If dbc(0) <> "" Then
   iNew = True
   For i = 0 To dbc(0).Rows - 1
     bm = dbc(0).GetBookmark(i)
     If dbc(0).Columns(0).CellText(bm) = dbc(0) Then
       meb(0).Text = dbc(0).Columns(1).CellText(bm)
       meb(0).Enabled = False
      ' tbFld(1).SetFocus
       iNew = False
       Exit For
     End If
   Next
   If iNew Then
    meb(0).Enabled = True
    meb(0).SetFocus
   End If
  Else
   dbc(0).SetFocus
  End If
Case 1
  '---------------------------------------
  '-- get corresponding joint record for
  '-- displaying
  '----------------------------------------
  qDMLQry = "SELECT JNTNAME1,JNTNAME2 FROM STKJOINT WHERE "
  qDMLQry = qDMLQry & " CLIENTID = " & dbc(1).Columns(1).Text
  qDMLQry = qDMLQry & " and JNTENDDTE  is NULL"
  rsJoint.Open qDMLQry, cnn, , , adCmdText
  If Not rsJoint.EOF Then
    If rsJoint!JNTNAME1 <> "" Then lblLabels(12).Caption = rsJoint!JNTNAME1 Else lblLabels(12).Caption = " "
    If rsJoint!JNTNAME2 <> "" Then lblLabels(13).Caption = rsJoint!JNTNAME2 Else lblLabels(13).Caption = " "
  Else
    lblLabels(12).Caption = " "
    lblLabels(13).Caption = " "
  End If
  rsJoint.Close
  If gblOptions = 1 Then
    '----------------------------------------
    '-- get all active certs and load in dbc
    '----------------------------------------
    qDMLQry = "SELECT * FROM CERTMST WHERE CLIENTID = "
    qDMLQry = qDMLQry & dbc(1).Columns(1).Text
    qDMLQry = qDMLQry & " and STATUS = 'A' ORDER BY CERTNO"
    rsCert.Open qDMLQry, cnn, , , adCmdText
    With rsCert
      If Not .EOF Then
         .MoveFirst
         dbdd(0).RemoveAll
         iSellShares = 0
         iBuyShares = 0
         Do While Not .EOF
            sRowinfo = !certno & Chr(9) & !shares & Chr(9)
            sRowinfo = sRowinfo & !IssDate
            dbdd(0).AddItem sRowinfo
            iSellShares = iSellShares + !shares
            .MoveNext
         Loop
      End If
      .Close
 
  End With
  meb(2).Text = iSellShares
  End If
Case Else
End Select
End Sub

Private Sub Form_Activate()
 On Error GoTo Form_Activate_Err
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
'--
If gblOptions = 2 Then
   UpdateScreen
   Me.Caption = "Edit Stockholder to Stockholder"
   meb(0).Enabled = False   ' Batch Date
   dbc(0).Enabled = False   ' Batch No
End If
'--
Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
 If Err = -2147168242 Then ' no current transactions
   Resume 0
 Else
   MsgBox "SIS026/Activate"
   Exit Sub
 End If
End Sub

Private Sub Form_Load()
Dim iDay, ipos, i As Integer
Dim qMain As String
Dim qSql1, sBatch, sForm, strTmp As String

On Error GoTo FL_ERR
iEOF = False
'--
   csvCenterForm Me, gblMDIFORM
   '''Set cnn = New ADODB.Connection
   cnn.Open cnn
   Set rsCmp = New ADODB.Recordset
   Set rsClient = New ADODB.Recordset
   Set rsJoint = New ADODB.Recordset
   Set rsMain = New ADODB.Recordset
   Set rsCert = New ADODB.Recordset
   Set rsUnused = New ADODB.Recordset
   Set rsBat = New ADODB.Recordset
   Set rsVerBat = New ADODB.Recordset
   Set rsVerFrm = New ADODB.Recordset
   '-----------------------
   '-- open tables --------
   '-----------------------
   If gblOptions = 2 Then
       ipos = InStr(1, gblFileKey, ";", 1)
       sBatch = Mid(gblFileKey, 1, ipos - 1)
       sForm = Mid(gblFileKey, ipos + 1, (Len(gblFileKey) - ipos))
   Else
       sBatch = ""
       sForm = ""
   End If
   '--
   qMain = "SELECT * FROM STKACTIV WHERE TRNBATCH = '"
   qMain = qMain & sBatch & "' and FORM = '" & sForm & "' and "
   qMain = qMain & " TRNCODE = 'S' and STATUS = 'O'"
   qMain = qMain & " order by FORM, stklineno "
   rsMain.Open qMain, cnn, adOpenDynamic, adLockPessimistic, adCmdText
   rsBat.Open "BATHDR", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
   rsCmp.Open "COMPANY", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
   rsUnused.Open "UNUSEDNOS", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '--
   If gblOptions = 1 Then
      InitAddNew
   End If
   '--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS026/Load"
  Unload Me
   
End Sub
Private Sub UpdateScreen()
Dim i As Integer
Dim qSQL, sRowinfo As String
With rsMain
  Do While Not .EOF
    If !stklineno = 1 Then
      grd(0).RemoveAll
      grd(0).Caption = "Cancelled Certificates"
      grd(0).Columns(2).Caption = "Cancelled Date"
      grd(0).Columns(2).Name = "Cancelled Date"
      grd(1).RemoveAll
      '--
      For i = 0 To dbc(0).Rows - 1
        bm = dbc(0).GetBookmark(i)
        If dbc(0).Columns(0).CellText(bm) = !TrnBatch Then
          dbc(0).Bookmark = dbc(0).GetBookmark(i)
          dbc(0) = dbc(0).Columns(0).CellText(bm)
          meb(0).Text = dbc(0).Columns(1).CellText(bm)
          Exit For
        End If
      Next i
      '--
      tbfld(1).Text = !Form
      meb(1).Text = !TRNDATE
      meb(2).Visible = False
      lbl.Visible = False
      '--
      qSQL = "SELECT CLINAME,  CLIENTID, SHARES FROM STKNAME "
      qSQL = qSQL & " WHERE CLIENTID = " & !ClientId
      rsClient.Open qSQL, cnn, , , adCmdText
      FillCombo
      dbc_LostFocus (1)
      rsClient.Close
    End If
    '--
    If !FRCERT > 0 Then
       sRowinfo = !stklineno & Chr(9) & !FRCERT & Chr(9)
       sRowinfo = sRowinfo & !CANDATE & Chr(9) & !FRSHARES
       grd(0).AddItem sRowinfo
    Else
       qSQL = "SELECT CLINAME from STKNAME where CLIENTID = "
       qSQL = qSQL & !ClientId
       rsClient.Open qSQL, cnn, , , adCmdText
       sRowinfo = !stklineno & Chr(9) & rsClient!CliName & Chr(9)
       sRowinfo = sRowinfo & !ClientId & Chr(9)
       sRowinfo = sRowinfo & !shares & Chr(9) & !certno
       grd(1).AddItem sRowinfo
       rsClient.Close
    End If
    .MoveNext
  Loop
  '--
  dbc(0).Enabled = False
  dbc(1).Enabled = False
  meb(0).Enabled = False
  meb(1).Enabled = False
  meb(2).Enabled = False
  grd(0).Enabled = False
  grd(1).Enabled = False
  tbfld(1).Enabled = False
  cmdUpdate.Enabled = False
  cmdFind.Enabled = False
End With
End Sub

Private Sub grd_AfterColUpdate(Index As Integer, ByVal ColIndex As Integer)
Dim iErr As Integer

Select Case Index
Case 0 ' cancelled certs
   Select Case ColIndex
   Case 1
       iSellShares = Val(meb(2))
       With grd(0)
          If IsNumeric(.Columns(3).Text) Then
             iSellShares = iSellShares + .Columns(3).Value
             iBuyShares = iBuyShares - .Columns(3).Value
          End If
          .Columns(2).Text = dbdd(0).Columns(2).Value 'issue date
          .Columns(3).Value = dbdd(0).Columns(1).Value  ' shares
          iSellShares = iSellShares - dbdd(0).Columns(1).Value
          iBuyShares = iBuyShares + dbdd(0).Columns(1).Value
       End With
       meb(2).Text = iSellShares
       
   Case Else
   End Select
Case 1
   Select Case ColIndex
   Case 1
       With grd(1)
          .Columns(2).Text = dbdd(1).Columns(1).Value
          
       End With
   Case 3
      With grd(1)
         If .Columns(3).Value > iBuyShares Then
            iErr = 145
            csvShowUsrErr iErr, "Shareholder to Shareholder"
            .SetFocus
            .Columns(3).Value = 0
         Else
            iBuyShares = iBuyShares - .Columns(3).Value
         End If
      End With
   Case Else
   End Select
Case Else
End Select
End Sub


Private Sub grd_BeforeDelete(Index As Integer, Cancel As Integer, DispPromptMsg As Integer)
Select Case Index
Case 0
   iSellShares = Val(meb(2))
   iSellShares = iSellShares + grd(0).Columns(3).Value
   iBuyShares = iBuyShares - grd(0).Columns(3).Value
   meb(2).Text = iSellShares
Case 1
   iBuyShares = iBuyShares + grd(1).Columns(3).Value
Case Else
End Select
End Sub
Private Sub grd_BeforeInsert(Index As Integer, Cancel As Integer)
Dim bm1 As Variant
Dim iCertno As Long, iClient As Long, i As Integer
Select Case Index
Case 0
  With grd(0)
     iCertno = dbdd(0).Columns(0).Value
     If .Rows = 0 Then Exit Sub
     .MoveFirst
     For i = 0 To .Rows - 1
       bm1 = .GetBookmark(i)
       If iCertno = Val(.Columns(1).CellText(bm1)) Then
          MsgBox "Duplicate error"
          Cancel = True
          Exit For
       End If
   Next i
  End With
Case 1
  With grd(1)
    iClient = dbdd(1).Columns(1).Value
    If .Rows = 0 Then Exit Sub
    .MoveFirst
    For i = 0 To .Rows - 1
      bm1 = .GetBookmark(i)
       If iClient = Val(.Columns(1).CellText(bm1)) Then
          MsgBox "Duplicate error"
          Cancel = True
          Exit For
       End If
   Next i
  End With
Case Else
End Select
End Sub

Private Sub grd_InitColumnProps(Index As Integer)
Select Case Index
Case 0
grd(0).Columns(1).DropDownHwnd = dbdd(0).hwnd
Case 1
grd(1).Columns(1).DropDownHwnd = dbdd(1).hwnd
Case Else
End Select
End Sub

Private Sub meb_GotFocus(Index As Integer)

Select Case Index
Case 0, 1
  meb(Index).Mask = "##-???-####"

Case Else
End Select
End Sub

Private Sub meb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  Select Case Index
  Case 0
   tbfld(1).SetFocus
  Case 1
    If dbc(1).Enabled = True Then
       dbc(1).SetFocus
    Else
       cmdFind.SetFocus
    End If
  Case Else
  End Select
Case vbKeyUp
KeyCode = 0
  Select Case Index
  Case 0
    dbc(0).SetFocus
  Case 1
    tbfld(1).SetFocus
  Case Else
  End Select
Case Else
End Select
End Sub
Private Sub ClearScreen()
 Dim i As Integer
If gblOptions = 1 Then
    dbc(1) = ""
    dbc(1).RemoveAll
    dbc(1).Enabled = False
    For i = 1 To 2
      meb(i).Mask = ""
      meb(i).Text = ""
    Next
    tbfld(1) = ""
    lblLabels(12).Caption = ""
    lblLabels(13).Caption = ""
    grd(0).RemoveAll
    grd(1).RemoveAll
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
End If
End Sub
Private Sub InitAddNew()
ClearScreen
  Me.Caption = "New Stockholder to Stockholder"
  icert = 0
  cmdDelete.Enabled = False
  cmdUpdate.Enabled = False
  grd(0).RemoveAll
  grd(1).RemoveAll
  grd(0).Enabled = False
  grd(1).Enabled = False
End Sub
Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
 KeyCode = 0
 If Index = 1 Then
    meb(1).SetFocus
 End If
Case vbKeyUp
 If gblOptions = 1 Then
  If Index = 1 Then
    If iNew Then
       meb(0).SetFocus
    Else
       dbc(0).SetFocus
    End If
  End If
 End If
Case Else
End Select
End Sub

