VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSIS010A 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Mandates"
   ClientHeight    =   7050
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "SIS010A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   6795
   Begin MSComCtl2.DTPicker DTPStartDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   33
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   51511299
      CurrentDate     =   40761
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
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
   Begin VB.TextBox TxtAccountName 
      Height          =   285
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   8
      ToolTipText     =   "Enter the name of the account if different from the shareholder."
      Top             =   6000
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox TxtAccountNo 
      Height          =   285
      Left            =   1680
      MaxLength       =   16
      TabIndex        =   7
      ToolTipText     =   "Enter an account number to credit the shareholder's payments. This is optional."
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   4
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   3
      ToolTipText     =   "Enter Address line 2"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   5
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   4
      ToolTipText     =   "Enter Address Line 3"
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   6
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   5
      ToolTipText     =   "Enter Address Line 4"
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   7
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   6
      ToolTipText     =   "Enter Address Line 5"
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   3
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   2
      ToolTipText     =   "Enter the first line of the mandates address."
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   2
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Enter the name of the bank or recipient to receive the payment."
      Top             =   2280
      Width           =   4335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3600
      TabIndex        =   19
      ToolTipText     =   "Clears the screen and resets it if in edit mode."
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5760
      TabIndex        =   13
      ToolTipText     =   "Cancels changes and returns to Account maintenance."
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   4680
      TabIndex        =   12
      ToolTipText     =   "Update Joint Table for saving to disk by Accounts Maintainace"
      Top             =   6720
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
         Size            =   8.25
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
      TabIndex        =   10
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
         Size            =   8.25
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
      TabIndex        =   11
      ToolTipText     =   "Enter Address line 2"
      Top             =   840
      Width           =   4335
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   1680
      Width           =   1935
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
   Begin SSDataWidgets_B.SSDBCombo SSDBACHBranches 
      Height          =   255
      Left            =   1680
      TabIndex        =   29
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
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
      Columns.Count   =   2
      Columns(0).Width=   5741
      Columns(0).Caption=   "Branch Name"
      Columns(0).Name =   "Branch Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2275
      Columns(1).Caption=   "Branch ID"
      Columns(1).Name =   "Branch ID"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2566
      _ExtentY        =   450
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
   Begin SSDataWidgets_B.SSDBCombo SSDBACHBanks 
      Height          =   255
      Left            =   1680
      TabIndex        =   30
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
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
      Columns.Count   =   2
      Columns(0).Width=   5741
      Columns(0).Caption=   "Bank Name"
      Columns(0).Name =   "Bank Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2275
      Columns(1).Caption=   "Bank ID"
      Columns(1).Name =   "Bank ID"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2566
      _ExtentY        =   450
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
   Begin SSDataWidgets_B.SSDBCombo SSDBAccountType 
      Height          =   255
      Left            =   1680
      TabIndex        =   31
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
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
      Columns.Count   =   2
      Columns(0).Width=   5741
      Columns(0).Caption=   "Account Type"
      Columns(0).Name =   "Account Type"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2275
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2566
      _ExtentY        =   450
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
   Begin MSComCtl2.DTPicker DTPEndDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   34
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   51511299
      CurrentDate     =   40761
   End
   Begin VB.Label LblType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "AccountType:"
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
      Left            =   120
      TabIndex        =   32
      Top             =   5520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblBranchID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Branch Id:"
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
      Left            =   360
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   6840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblBankID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Bank Id:"
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
      Left            =   600
      TabIndex        =   27
      Top             =   4200
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Payment Method:"
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
      Index           =   11
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   6840
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Account Name:"
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
      Index           =   10
      Left            =   120
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Address:"
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
      Left            =   960
      TabIndex        =   24
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "End Date:"
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
      Left            =   3720
      TabIndex        =   23
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
      BackColor       =   &H00C0E0FF&
      Caption         =   "Name:"
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
      TabIndex        =   22
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Label lblAccountNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Account Number:"
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
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Start Date:"
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
      Left            =   480
      TabIndex        =   20
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   6960
      Y1              =   6480
      Y2              =   6480
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
      TabIndex        =   17
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Account Name:"
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
      Left            =   240
      TabIndex        =   16
      Top             =   840
      Width           =   1500
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
      TabIndex        =   15
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Stockholder No:"
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
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1620
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
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSIS010A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iErr As Integer
Dim X As Integer
Dim rsMandate As ADODB.Recordset
Dim rsBank As ADODB.Recordset
Dim rsDist As ADODB.Recordset
Dim rsACHBanks As ADODB.Recordset
Dim rsACHBrchs As ADODB.Recordset
Dim rsAccType As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim OpenErr As Integer
Dim iOpenMan As Integer
Dim iOpenBank As Integer
Dim strTable As String
Dim iMode As Integer  ' 0 = New; 1 = Active; 2 = inactive joint
Function IsValid() As Integer
On Error GoTo IsValid_Err
Dim dtefld As Date
Dim length As Integer

iErr = 0
IsValid = False
'--
If iMode = 0 Or iMode = 2 Then
  '--
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
  If DTPEndDate.Text = "" Then
      iErr = 37
      MsgBox iErr, "Mandates"
      DTPEndDate.SetFocus
      GoTo Validate_Exit
   End If
   '--
   
   '--
   dtefld = DTPEndDate
   If dtefld < Format(DTPStartDate.Text, "dd-mmm-yyyy") Then
      iErr = 38
      MsgBox iErr, "Mandates"
      DTPEndDate.SetFocus
      GoTo Validate_Exit
   End If
  '--
End If
'--

If dbc(0).Text <> "1" Then
   GoTo IsValidSet
End If
length = Len(tbfld(8).Text)
If length <> 9 Then
   MsgBox "Finacle Account must be nine digits"
   tbfld(8).SetFocus
   GoTo Validate_Exit
End If
If IsNumeric(tbfld(8).Text) Then
Else
   MsgBox "Finacle Account number must be numeric"
   tbfld(8).SetFocus
   GoTo Validate_Exit
End If

IsValidSet:
IsValid = True
Validate_Exit:
   Exit Function
'--
Validate_Err:
  MsgBox iErr, "Mandates"
  GoTo Validate_Exit
'--
IsValid_Err:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on validation"
Resume Validate_Exit

End Function

Private Sub cmdCancel_Click()
On Error GoTo Err_CmdCancel_Click

Shutdown
Unload Me

Exit_CmdCancel_Click:
Exit Sub

Err_CmdCancel_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on closing"
Unload Me

End Sub
Private Sub cmdClear_Click()
On Error GoTo Err_CmdClear_Click
If iMode = 0 Then
   ClearScreen
  DTPStartDate.SetFocus
Else
   ClearScreen
   UpdateScreen
End If

Exit_CmdClear_Click:
Exit Sub

Err_CmdClear_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on clearing screen"
Resume Exit_CmdClear_Click
End Sub
Private Sub cmdUpdate_Click()
Dim strChg As Integer, iAcct As Long
Dim i As Integer
Dim newval As Integer
On Error GoTo cmdUpdate_Err
If IsValid Then
   iAcct = Val(tbfld(0).Text)
  '--
i = RunSP(SpCon, "usp_MandateUpdate", 0, iMode, iAcct, gblLoginName, Format(DTPStartDate, "dd-mmm-yyyy"), tbfld(8), tbfld(9), Trim(dbc(0).Columns(1).Text), tbfld(2), tbfld(3), tbfld(4), tbfld(5), tbfld(6), tbfld(7), Format(DTPEndDate, "dd-mmm-yyyy"), dbc(1).Columns(1).Text)
If i = 1 Then
   MsgBox "Record sucessfully updated"
Else
   MsgBox "Update was unsucessfull. Sorry for any inconvienience caused"
   GoTo Done
End If

If iMode = 1 Then
     iMode = 0
     DTPStartDate = DateValue(DTPEndDate) + 1
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
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on saving"
Resume Done

End Sub

Private Sub dbc_Click(Index As Integer)
On Error GoTo dbc_Click_Err

Select Case Index
Case 1
     If dbc(1).Columns(1).Text = "0" Then
        ACHFields (False)
        BankFields (False)
        ShowBankID (False)
        ChangeDataStatus (True)
     End If
     If dbc(1).Columns(1).Text = "3" Then
        ACHFields (True)
        BankFields (True)
        ShowBankID (False)
        ChangeDataStatus (False)
     Else
         ACHFields (False)
         BankFields (True)
         ShowBankID (True)
         ChangeDataStatus (False)
     End If
     
End Select

Exit_dbc_Click:
Exit Sub

dbc_Click_Err:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error Populating dropdown boxes"
Resume Exit_dbc_Click

End Sub

Private Sub dbc_InitColumnProps(Index As Integer)
On Error GoTo dbc_InitColumnProps_Err

Dim sRowinfo As String

Select Case Index
Case 0
  
  dbc(0).RemoveAll
  If rsBank.EOF Then
     rsBank.MoveFirst
  End If
  With rsBank
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        sRowinfo = !BnkName & vbTab & !BankID
        dbc(0).AddItem sRowinfo
       .MoveNext
      Loop
    End If
  End With
  '--
Case 1
 With rsDist
      dbc(1).RemoveAll
      Do While Not .EOF
         dbc(1).AddItem !DistDesc & vbTab & !DistCode
         .MoveNext
      Loop
 End With
 '--
Case Else
End Select
Exit Sub
dbc_InitColumnProps_Err:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error Populating dropdown boxes"
Unload Me
End Sub

Private Sub dbc_LostFocus(Index As Integer)
On Error GoTo Err_dbc_LostFocus

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

Exit_dbc_LostFocus:
Exit Sub

Err_dbc_LostFocus:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on clearing screen"
Resume Exit_dbc_LostFocus

End Sub

Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
'--
' ready message
'---
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.Refresh
UpdateScreen

Exit_Form_Activate:
Exit Sub

Err_Form_Activate:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on clearing screen"
Resume Exit_Form_Activate

End Sub

Private Sub Form_Load()
Dim qSQL As String, qView As String
Dim strTmp As String
On Error GoTo FL_ERR
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
'Set rsMandate = New ADODB.Recordset
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
Set rsBank = rsMandate.NextRecordset
Set rsDist = rsMandate.NextRecordset
Set rsACHBanks = rsMandate.NextRecordset
Set rsACHBrchs = rsMandate.NextRecordset
Set rsAccType = rsMandate.NextRecordset
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
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on clearing screen"
Unload Me

End Sub
Private Sub UpdateScreen()
Dim bm As Variant, i As Integer
dbc_InitColumnProps (0)
tbfld(0).Text = gblFileKey
tbfld(1).Text = FrmMandate.TxtCliName
'--
ShowBankID (False)
BankFields (False)
ACHFields (False)
ChangeDataStatus (False)

If iMode = 0 Then
   DTPStartDate = Format(Date, "dd-mmm-yyyy")
   GoTo Exit_UpdateScreen
End If
With dbc(1)
     .MoveFirst
     For i = 0 To .Rows - 1
         bm = .GetBookmark(i)
         If .Columns(1).CellText(bm) = Trim(rsMandate!MndMet) Then
            .Bookmark = .GetBookmark(i)
            dbc(1) = .Columns(0).CellText(bm)
            Exit For
          End If
     Next i
End With
BankFields (True)
If Not IsNothing(!MndAcnt) Then tbfld(8).Text = !MndAcnt
If Not IsNothing(!MndAcntNme) Then tbfld(9).Text = !MndAcntNme

With rsMandate
     DTPStartDate.Text = !MndStaDte
     If !MndMet = 0 Then 'Disposal option is cheque
        ChangeDataStatus (True)
        If Not IsNothing(!MndName) Then tbfld(2).Text = !MndName
        If Not IsNothing(!MndAddr1) Then tbfld(3).Text = !MndAddr1
        If Not IsNothing(!MndAddr2) Then tbfld(4).Text = !MndAddr2
        If Not IsNothing(!MndAddr3) Then tbfld(5).Text = !MndAddr3
        If Not IsNothing(!MNDADDR4) Then tbfld(6).Text = !MNDADDR4
        If Not IsNothing(!MNDADDR5) Then tbfld(7).Text = !MNDADDR5
        GoTo Exit_UpdateScreen
     End If
       
     BankFields (True)

     If !MndMet = 3 Then 'Disposal option is ACH
        ACHFields (True)
        With SSDBACHBanks
             .MoveFirst
             For i = 0 To .Rows - 1
                 bm = .GetBookmark(i)
                 If .Columns(1).CellText(bm) = Trim(rsMandate!BankID) Then
                    .Bookmark = .GetBookmark(i)
                    SSDBACHBanks = .Columns(0).CellText(bm)
                    Exit For
                 End If
             Next i
        End With
        
        With SSDBACHBranches
             .MoveFirst
             For i = 0 To .Rows - 1
                 bm = .GetBookmark(i)
                 If .Columns(1).CellText(bm) = Trim(rsMandate!BranchID) Then
                    .Bookmark = .GetBookmark(i)
                    SSDBACHBranches = .Columns(0).CellText(bm)
                    Exit For
                 End If
             Next i
        End With
        
        With SSDBAccountType
             .MoveFirst
             For i = 0 To .Rows - 1
                 bm = .GetBookmark(i)
                 If .Columns(1).CellText(bm) = Trim(rsMandate!AccountType) Then
                    .Bookmark = .GetBookmark(i)
                    SSDBAccountType = .Columns(0).CellText(bm)
                    Exit For
                 End If
             Next i
        End With
      Else
          With dbc(0)
             ShowBankID (True)
             .MoveFirst
             For i = 0 To .Rows - 1
                 bm = .GetBookmark(i)
                 If .Columns(1).CellText(bm) = Trim(rsMandate!BankID) Then
                    .Bookmark = .GetBookmark(i)
                    dbc(0) = .Columns(0).CellText(bm)
                    Exit For
                 End If
             Next i
        End With
     End If
End With

Exit_UpdateScreen:
Exit Sub
End Sub
Private Sub ClearScreen()

  For X = 2 To 7
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
     DTPEndDate.SetFocus
  Else
     DTPStartDate.SetFocus
  End If
End Sub

Private Sub Shutdown()
If SpCon.State = 1 Then
   If iOpenMan = True Then rsMandate.Close
   If iOpenBank = True Then rsBank.Close
   rsDist.Close
End If
Set rsMandate = Nothing
End Sub

Private Sub ChangeDataStatus(DataStatus As Boolean)
Dim X As Integer
DTPStartDate.Enabled = DataStatus
For X = 2 To 7
  tbfld(X).Enabled = DataStatus
Next
DTPEndDate.Visible = Not DataStatus
lblLabels(3).Visible = Not DataStatus
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub
Function ACHFields(VStatus As Boolean)
SSDBACHBanks.Visible = VStatus
SSDBACHBranches.Visible = VStatus
SSDBAccountType.Visible = VStatus
lblBankID.Visible = VStatus
lblBranchID.Visible = VStatus
LblType.Visible = VStatus
End Function
Function BankFields(VStatus As Boolean)
TxtAccountNo.Visible = VStatus
TxtAccountName.Visible = VStatus
End Function

Private Sub SSDBAccountType_InitColumnProps()
SSDBAcctType.RemoveAll
With rsAccType
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        SSDBAccountType.AddItem !BankName & vbTab & !BankID
       .MoveNext
      Loop
    End If
    .Close
End With
End Sub

Private Sub SSDBACHBanks_InitColumnProps()
SSDBACHBanks.RemoveAll
With rsACHBanks
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        SSDBACHBanks.AddItem !BankName & vbTab & !BankID
       .MoveNext
      Loop
    End If
    .Close
End With

End Sub

Private Sub SSDBACHBranches_InitColumnProps()
SSDBACHBranches.RemoveAll
With rsACHBrchs
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        SSDBACHBranches.AddItem !BankName & vbTab & !BankID
       .MoveNext
      Loop
    End If
    .Close
End With
End Sub

Function ShowBankID(FldStatus As Boolean)
dbc(0).Visible = FldStatus
lblBankID.Visible = FldStatus

End Function
