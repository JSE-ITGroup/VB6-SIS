VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS032 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stockholder to Broker Transfer"
   ClientHeight    =   5190
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "SIS032.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7275
   Begin SSDataWidgets_B.SSDBDropDown dbdd 
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
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
      _ExtentX        =   2355
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B.SSDBGrid grd 
      Height          =   1455
      Index           =   0
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   5000
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      AllowColumnMoving=   0
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   926
      Columns(0).Caption=   "Line #"
      Columns(0).Name =   "Line #"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   2
      Columns(0).FieldLen=   2
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   2434
      Columns(1).Caption=   "Cert #"
      Columns(1).Name =   "Cert #"
      Columns(1).Alignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   10
      Columns(2).Width=   1984
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
      Columns(3).DataType=   5
      Columns(3).FieldLen=   11
      Columns(3).Locked=   -1  'True
      _ExtentX        =   8819
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
      Left            =   1920
      TabIndex        =   24
      ToolTipText     =   "Pressing this button will activate the search program to locate a shareholder."
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox tbFld 
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   2
      ToolTipText     =   "Assign a document number."
      Top             =   960
      Width           =   1500
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   4080
      TabIndex        =   21
      Top             =   4800
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
      Left            =   3000
      TabIndex        =   15
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   6240
      TabIndex        =   9
      ToolTipText     =   "Cancels all processing and exits program."
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   5160
      TabIndex        =   8
      ToolTipText     =   "Saves the screen information to the database."
      Top             =   4800
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
      Top             =   960
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
      Top             =   1440
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
      Left            =   5520
      TabIndex        =   25
      ToolTipText     =   "Enter the date of  issue you want to appear on the certificate.."
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   15
      Format          =   "#,###"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Select a share holder from the list."
      Top             =   4200
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
      Index           =   3
      Left            =   5520
      TabIndex        =   7
      ToolTipText     =   "Enter the date of  issue you want to appear on the certificate.."
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   15
      Format          =   "#,###"
      PromptChar      =   "_"
   End
   Begin VB.Label lbl 
      Caption         =   "Shares:"
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
      Left            =   4680
      TabIndex        =   29
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "To Broker:"
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
      Index           =   5
      Left            =   240
      TabIndex        =   28
      Top             =   4200
      Width           =   1020
   End
   Begin VB.Line Line7 
      X1              =   5160
      X2              =   5160
      Y1              =   1320
      Y2              =   2280
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
      Index           =   0
      Left            =   5280
      TabIndex        =   26
      Top             =   1440
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
      Y1              =   2280
      Y2              =   2280
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
      TabIndex        =   23
      Top             =   960
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
      TabIndex        =   22
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
      TabIndex        =   20
      Top             =   1800
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
      TabIndex        =   19
      Top             =   2040
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
      TabIndex        =   18
      Top             =   2040
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
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   7320
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
      TabIndex        =   13
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   0
      X2              =   7320
      Y1              =   1320
      Y2              =   1320
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
      TabIndex        =   12
      Top             =   1440
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   960
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
      TabIndex        =   14
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmSIS032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim X As Integer, i As Integer, iNewBroker As Integer, iEOF As Integer
Dim iNewApp As Integer, iNew As Integer
Dim iAvailShares, iCancelShares As Double
Dim rsCmp As ADODB.Recordset
Dim rsClient As ADODB.Recordset
Dim rsMain As ADODB.Recordset, rsBroker As ADODB.Recordset
Dim rsCert As ADODB.Recordset
Dim rsJoint As ADODB.Recordset
Dim rsUnused As ADODB.Recordset
Dim rsBat As ADODB.Recordset
Dim rsVerBat As ADODB.Recordset
Dim rsVerFrm As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim errLoop As Error
Dim errs1 As Error
Dim LastSelection As Integer
Dim strTable As String
Dim strRecNO As String
Dim eClient As Long
Dim iStocks As Double, icert As Long, iBrokerId As Long
Dim bm As Variant, iCertno As Long
Private Sub FillCombo(i As Integer)
Dim sRowinfo As String
With rsClient
    If Not .EOF And Not .BOF Then
      dbc(i).RemoveAll
      Do While Not .EOF
         If i = 1 Then
            If !CatCode <> "SB" Then
            sRowinfo = !CliName & vbTab & !ClientID
            sRowinfo = sRowinfo & vbTab & !shares
            dbc(i).AddItem sRowinfo
            End If
         Else
            If !CatCode = "SB" Then
            sRowinfo = !CliName & vbTab & !ClientID
            sRowinfo = sRowinfo & vbTab & !shares
            dbc(i).AddItem sRowinfo
            End If
         End If
         If dbc(i).Row = 0 Then dbc(i) = !CliName
         .MoveNext
      Loop
      If i = 1 Then grd(0).RemoveAll
    End If
End With
End Sub
Function IsValid() As Integer
Dim iErr As Integer, dtefld As Date, qSQL
Dim sElable As String
sElable = "Stockholder to Broker Entry"
IsValid = False
iErr = 0
'--
If dbc(0) = "" Then ' batch
    iErr = 132
    MsgBox "Please enter a BATCH number"
    tbFld(0).SetFocus
    GoTo Validate_Exit
End If
dbc(0) = UCase(dbc(0))
'--
If meb(0) = "" Then 'batch date
   iErr = 139
   MsgBox "Please enter a BATCH date"
   meb(0).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(0)) Then
      iErr = 14
      MsgBox "BATCH DATE is not a date. Please correct"
      meb(0).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If tbFld(1).Text = "" Then 'Invalid form
   iErr = 140
   MsgBox "FORM number is required"
   tbFld(1).SetFocus
   GoTo Validate_Exit
End If
tbFld(1).Text = UCase(tbFld(1).Text)
'--
If gblOptions = 1 Then 'check for duplicate form
  Set rsVerFrm = RunSP(SpCon, "usp_S2B2", 1, dbc(0), tbFld(1))
  If Not rsVerFrm.EOF Then
     iErr = 141
     MsgBox "Duplicate FORM number found"
     tbFld(1).SetFocus
     rsVerFrm.Close
     GoTo Validate_Exit
  End If
  rsVerFrm.Close
End If
'--
If dbc(1) = "" Then   ' Sell Client name
   iErr = 130
   MsgBox "The Selling Client's name is missing. Please correct"
   dbc(1).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If meb(1) = "" Then 'Transfer date
   iErr = 129
   MsgBox "Transfer date is missing. Please correct"
   meb(1).SetFocus
   GoTo Validate_Exit
 Else
   If Not IsDate(meb(1)) Then
      iErr = 14
      MsgBox "Transfer date is not a valid date. Please correct"
      meb(1).SetFocus
      GoTo Validate_Exit
   End If
 End If
 '--
 If grd(0).Rows = 0 Then  ' empty grid no certs to cancel
   iErr = 143
   MsgBox "You did not select certificates to cancel. Please correct"
   grd(0).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If iCancelShares = 0 Then
    iErr = 154
    MsgBox "Shares being cancelled are 0. This should be corrected"
    grd(0).SetFocus
    GoTo Validate_Exit
 End If
 '--
 If dbc(2) = "" Then 'Buy Broker
   iErr = 153
   MsgBox "A Broker was not selected. Please select one now"
   dbc(2).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If meb(3) = "" Then ' Shares bought by broker
    iErr = 156
    MsgBox "Oops!, you left off the amount being bought by the Broker"
    meb(3).SetFocus
    GoTo Validate_Exit
 End If
 '--
 If Val(meb(3)) = 0 Then
    iErr = 157
    MsgBox "Oops!, The amount being bought by the broker should be greater than 0"
   meb(3).SetFocus
   GoTo Validate_Exit
 End If
 '--
 If Val(meb(3)) > iCancelShares Then
    iErr = 155
    MsgBox "Oops!. The Amount the Broker is buying is greater than the amount being sold"
   meb(3).SetFocus
   GoTo Validate_Exit
 End If
 '--
 IsValid = True
Validate_Exit:
   
   Exit Function
End Function

Private Sub cmdCancel_Click()
Dim iSeqKey As Integer

rsCmp.Close
rsMain.Close
Set rsBroker = Nothing
Set rsMain = Nothing
Set rsCmp = Nothing
Set rsClient = Nothing
Set rsJoint = Nothing
Set rsUnused = Nothing
Set rsVerBat = Nothing
Set rsVerFrm = Nothing

Unload Me
End Sub

Private Sub cmdClear_Click()
ClearScreen
If gblOptions = 2 Then UpdateScreen
End Sub

Private Sub CmdDelete_Click()
Dim imsg As Integer, X As Integer
Dim qDMLQry As String, qSQL As String
Dim iClient As Long, iStocks As Double, iTot As Double
Dim iCrt As Long, iLines As Integer, strnbatch As String
Dim bm As Variant, StrnDate As Date, i As Integer
imsg = 133
On Error GoTo cmdDelete_Err

i = MsgBox("Are you sure?", vbYesNo)

If i = vbYes Then
   '------------------------
   '-- Downdate Brokers Pool
   '------------------------
   cmdDelete.Enabled = False
   iStocks = 0
   With grd(0)
      .Redraw = False
      .MoveFirst
      For i = 0 To .Rows - 1
        bm = .GetBookmark(i)
        iStocks = iStocks + Val(.Columns(3).CellText(bm))
        iCrt = Val(grd(0).Columns(1).CellText(bm))
        qSQL = qSQL & iCrt & ";"
        iLines = iLines + 1
      Next i
   End With
   iTot = Val(meb(3))
   strnbatch = dbc(0).Text
   StrnDate = meb(1)
   qDMLQry = tbFld(1).Text
   X = RunSP(SpCon, "usp_SIS032Delete", 0, eClient, iBrokerId, strnbatch, _
    qDMLQry, Format(StrnDate, "dd-mmm-yyyy"), iLines, iTot, iStocks, qSQL)
 
End If

cmdDelete_Exit:
Exit Sub

cmdDelete_Err:
 MsgBox "cmdDelete"
 GoTo cmdDelete_Exit

End Sub

Private Sub cmdFind_Click()
Dim X As Integer
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
    If LastSelection = 0 Then
       .optBtn.Buttons(1).Value = True
    End If
    .Show vbModal
    '----------------------------
    '-------- main line ---------
    '----------------------------
    If .tbFind.Text = vbNullString Then
    Else
      If .cbOptions.ListIndex > 6 Then .cbOptions.ListIndex = 0
      sWhere = Trim(.tbFind.Text)
      X = .cbWhere.ListIndex
      cChk = .cbOptions.ListIndex
      qCli = sWhere
      If .optBtn.IndexSelected = 0 Then
         i = 1
      Else
         i = 2
      End If
      Set rsClient = RunSP(SpCon, "usp_ClientFind", 1, qCli, cChk, X, i)
      If Not rsClient.EOF Then
         If i = 1 Then
           FillCombo (i) ' Client selling
           grd(0).Enabled = True
         Else
          ' Client Buying Fill grd
            FillCombo (i)
            cmdUpdate.Enabled = True
         End If
         dbc(i).Enabled = True
         dbc(i).SetFocus
      End If
      rsClient.Close
    End If
  End With
Unload frmFind
Set frmFind = Nothing
Exit Sub
cmdFind_Click_err:
  MsgBox "SIS032/CmdFind"
End Sub
Private Sub cmdUpdate_Click()
 Dim qBat As String, iLines As Integer, iCrt As Long
Dim iClient As Long, iss As Date, iShares As Double
Dim strncode As String, strnbatch As String, StrnDate As Date
Dim i As Integer, X As Integer, bm As Variant
Dim qSQL As String
'On Error GoTo cmdUpdate_Err
' wait message
frmMDI.txtStatusMsg.SimpleText = gblWaitMsg
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.Refresh

If IsValid Then
  '--
  iLines = 0
  iClient = dbc(1).Columns(1).Text
  iss = DateValue(meb(1).Text)
  strncode = "C"
  strnbatch = dbc(0)
  StrnDate = DateValue(meb(0).Text)
  qBat = tbFld(1).Text
  iShares = Val(meb(3))
  '--
    grd(0).MoveFirst
    qSQL = ""
    For i = 0 To grd(0).Rows - 1
       bm = grd(0).GetBookmark(i)
       iCrt = Val(grd(0).Columns(1).CellText(bm))
       qSQL = qSQL & iCrt & ";"
       iLines = iLines + 1
    Next
X = RunSP(SpCon, "usp_SIS032Update", 0, iClient, iBrokerId, strnbatch, _
    qBat, strncode, Format(StrnDate, "dd-mmm-yyyy"), Format(iss, "dd-mmm-yyyy"), iLines, iShares, iCancelShares, qSQL)

    '--
        dbc_InitColumnProps (0)
        dbc(0).Enabled = False
        meb(0).Enabled = False
        InitAddNew
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
  MsgBox Err.Description, vbOKOnly, "SIS032/cmdUpdate"
  'MsgBox "SIS032/cmdUpdate", Err.Number, Err.Description
  cmdCancel_Click
End Sub

Private Sub dbc_Click(Index As Integer)
If Index = 1 Then
   LastSelection = 0
End If
End Sub

Private Sub dbc_InitColumnProps(Index As Integer)
Dim sRowinfo As String
Select Case Index
Case 0 ' Load Open Batches
With rsBat
    
  If Not .EOF And Not .BOF Then
     dbc(0).RemoveAll
     Do While Not .EOF
       sRowinfo = !BatchNo & vbTab & !BATDATE
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
         tbFld(1).SetFocus
      End If
    Case 1
      If grd(0).Enabled = True Then
        grd(0).SetFocus
      Else
        cmdFind.SetFocus
      End If
    Case 2
      meb(3).SetFocus
    Case Else
  End Select
     
 Case vbKeyUp
  KeyCode = 0
  Select Case Index
    Case 1
       tbFld(1).SetFocus
    Case 2
       grd(0).SetFocus
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
  If dbc(0) = "" Then
     dbc(0).SetFocus
  Else
   iNew = True
   For i = 0 To dbc(0).Rows - 1
     bm = dbc(0).GetBookmark(i)
     If dbc(0).Columns(0).CellText(bm) = dbc(0) Then
       meb(0).Mask = ""
       meb(0).Text = dbc(0).Columns(1).CellText(bm)
       meb(0).Enabled = False
       If Not IsDate(meb(1)) Then meb(1) = meb(0)
       iNew = False
       Exit For
     End If
   Next
   If iNew Then
    meb(0).Enabled = True
    meb(0).SetFocus
   End If
 End If
Case 1
  '---------------------------------------
  '-- get corresponding joint record for
  '-- displaying
  '----------------------------------------
  Set rsJoint = RunSP(SpCon, "usp_FindJoint", 1, dbc(1).Columns(1).Text)
  If Not rsJoint.EOF Then
    lblLabels(12).Caption = rsJoint!JNTNAME1
    If Not IsNothing(rsJoint!JNTNAME2) Then
      lblLabels(13).Caption = rsJoint!JNTNAME2
    End If
  Else
    lblLabels(12).Caption = " "
    lblLabels(13).Caption = " "
  End If
  rsJoint.Close
  If gblOptions = 1 Then
    '----------------------------------------
    '-- get all active certs and load in dbdd
    '----------------------------------------
    Set rsCert = RunSP(SpCon, "usp_CertMstFind", 1, dbc(1).Columns(1).Text)
    With rsCert
      If Not .EOF Then
         dbdd(0).RemoveAll
         iAvailShares = 0
         iCancelShares = 0
         Do While Not .EOF
            If !Status = "A" Then
               sRowinfo = !certno & Chr(9) & !shares & Chr(9)
               sRowinfo = sRowinfo & !IssDate
               dbdd(0).AddItem sRowinfo
               iAvailShares = iAvailShares + !shares
            End If
            .MoveNext
         Loop
      End If
      .Close
 
  End With
  meb(2).Text = iAvailShares
  End If
Case 2
   For i = 0 To dbc(2).Rows - 1
     bm = dbc(2).GetBookmark(i)
     If dbc(2).Columns(0).CellText(bm) = dbc(2) Then
       iBrokerId = Val(dbc(2).Columns(1).CellText(bm))
       Exit For
     End If
   Next
Case Else
End Select
End Sub

Private Sub Form_Activate()
' On Error GoTo Form_Activate_Err
' ready message
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
'--
If gblOptions = 2 Then
   UpdateScreen
   Me.Caption = "Edit Stockholder to Broker"
   meb(0).Enabled = False   ' Batch Date
   dbc(0).Enabled = False   ' Batch No
   LastSelection = 3
End If
'--

Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
 If Err = -2147168242 Then ' no current transactions
   Resume 0
 Else
   MsgBox "SIS032/Activate"
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
Loop
Screen.MousePointer = vbDefault

   Set rsClient = New ADODB.Recordset
   Set rsJoint = New ADODB.Recordset
   Set rsCert = New ADODB.Recordset
   Set rsVerBat = New ADODB.Recordset
   Set rsVerFrm = New ADODB.Recordset
   Set rsBroker = New ADODB.Recordset
   
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
   Set rsMain = RunSP(SpCon, "usp_S2B1", 1, sBatch, sForm)
   Set rsBat = rsMain.NextRecordset()
   Set rsCmp = rsMain.NextRecordset()
   '-------------------------------------
   '-- Initialize Company Details -------
   '-------------------------------------
   lblLabels(0).Caption = gblCompName
   lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
   '--
   Set rsUnused = rsMain.NextRecordset()
    If gblOptions = 1 Then
      InitAddNew
    End If
   '--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS032/Load"
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
      dbc(2).RemoveAll
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
      tbFld(1).Text = !Form
      meb(1).Text = !TrnDate
      meb(2).Visible = False
      lbl(0).Visible = False
      '--
      Set rsClient = RunSP(SpCon, "usp_ClientFind", 1, !ClientID, 0, 1, 1)
      eClient = !ClientID
      i = 1
      FillCombo (i)
      dbc_LostFocus (1)
      rsClient.Close
      iBrokerId = !BROKERID
    End If
    '--
    If !FRCERT > 0 Then
       sRowinfo = !stklineno & vbTab & !FRCERT & vbTab
       sRowinfo = sRowinfo & !CanDate & vbTab & !FRSHARES
       grd(0).AddItem sRowinfo
    Else
       Set rsClient = RunSP(SpCon, "usp_ClientFind", 1, !ClientID, 0, 1, i)
       i = 2
       FillCombo (i)
       meb(3) = !shares
      ' meb(3) = 0
       rsClient.Close
    End If
    
    .MoveNext
  Loop
  '--
  dbc(0).Enabled = False
  dbc(1).Enabled = False
  dbc(2).Enabled = False
  meb(0).Enabled = False
  meb(1).Enabled = False
  meb(2).Enabled = False
  meb(3).Enabled = False
  grd(0).Enabled = False
  tbFld(1).Enabled = False
  cmdUpdate.Enabled = False
  cmdFind.Enabled = False
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub grd_AfterColUpdate(Index As Integer, ByVal ColIndex As Integer)
Dim iErr As Integer, IHave As Integer

Select Case Index
Case 0 ' cancelled certs
   Select Case ColIndex
   Case 1
     IHave = 0 'False
     iCertno = grd(0).Columns(1).Value
     dbdd(0).MoveFirst
     For i = 0 To dbdd(0).Rows - 1
       bm = dbdd(0).GetBookmark(i)
       If dbdd(0).Columns(0).CellText(bm) = iCertno Then
           iAvailShares = Val(meb(2))
           With grd(0)
              If IsNumeric(.Columns(3).Text) Then
                  iAvailShares = iAvailShares + .Columns(3).Value
                  iCancelShares = iCancelShares - .Columns(3).Value
              End If
              .Columns(2).Text = dbdd(0).Columns(2).CellText(bm) 'issue date
              .Columns(3).Value = dbdd(0).Columns(1).CellText(bm)  ' shares
              iAvailShares = iAvailShares - dbdd(0).Columns(1).CellText(bm)
              iCancelShares = iCancelShares + dbdd(0).Columns(1).CellText(bm)
           End With
           meb(2).Text = iAvailShares
           dbdd(0).RemoveItem (i)
           IHave = True
           Exit For
       End If
     Next i
     If IHave = False Then
        grd(0).Columns(1).Text = ""
     End If
   Case Else
   End Select
   End Select
End Sub


Private Sub grd_BeforeDelete(Index As Integer, Cancel As Integer, DispPromptMsg As Integer)
Dim sRowinfo As String
Select Case Index
Case 0
   iAvailShares = Val(meb(2))
   iAvailShares = iAvailShares + grd(0).Columns(3).Value
   iCancelShares = iCancelShares - grd(0).Columns(3).Value
   meb(2).Text = iAvailShares
   sRowinfo = grd(0).Columns(1).Value & Chr(9) & _
              grd(0).Columns(3).Value & Chr(9) & _
              grd(0).Columns(2).Text
   dbdd(0).AddItem sRowinfo
Case Else
End Select
End Sub

Private Sub grd_InitColumnProps(Index As Integer)
Select Case Index
Case 0
grd(0).Columns(1).DropDownHwnd = dbdd(0).hwnd
Case Else
End Select
End Sub


Private Sub meb_GotFocus(Index As Integer)

Select Case Index
Case 0
  meb(Index).Mask = "##-???-####"
Case 1
  If meb(1) = "" Then meb(Index).Mask = "##-???-####"
Case Else
End Select
End Sub

Private Sub meb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
  KeyCode = 0
  Select Case Index
  Case 0
   tbFld(1).SetFocus
  Case 1
    If dbc(1).Enabled = True Then
       dbc(1).SetFocus
    Else
       cmdFind.SetFocus
    End If
  Case 3
    cmdUpdate.SetFocus
  Case Else
  End Select
Case vbKeyUp
KeyCode = 0
  Select Case Index
  Case 0
    dbc(0).SetFocus
  Case 1
    tbFld(1).SetFocus
  Case 3
    If dbc(2).Enabled Then
          dbc(2).SetFocus
    End If
  Case Else
  End Select
Case Else
End Select
End Sub
Private Sub ClearScreen()
Dim X As Integer
If gblOptions = 1 Then
  For X = 1 To 2
    meb(X).Mask = ""
    meb(X).Text = ""
    dbc(X) = ""
    dbc(X).RemoveAll
    dbc(X).Enabled = False
  Next
    meb(3).Text = ""
    tbFld(1) = ""
    lblLabels(12).Caption = ""
    lblLabels(13).Caption = ""
    grd(0).RemoveAll
    dbdd(0).RemoveAll
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
End If
End Sub
Private Sub InitAddNew()
ClearScreen
  Me.Caption = "New Stockholder to Broker"
  icert = 0
  grd(0).RemoveAll
  grd(0).Enabled = False
  meb(0).Enabled = False
End Sub

Private Sub meb_LostFocus(Index As Integer)
Select Case Index
  Case 0
      If IsDate(meb(0)) Then
        meb(1) = meb(0)
      End If
End Select
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

