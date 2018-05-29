VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSIS002 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Account Maintenance"
   ClientHeight    =   7980
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "SIS002.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   6825
   Begin VB.CheckBox ChkDeceased 
      Alignment       =   1  'Right Justify
      Caption         =   "Deceased"
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
      Left            =   4680
      TabIndex        =   48
      ToolTipText     =   "Check to indicate deceased"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton CmdDocuments 
      Caption         =   "&Documents"
      Height          =   300
      Left            =   120
      TabIndex        =   47
      Top             =   7605
      Width           =   975
   End
   Begin VB.CheckBox ChkResidence 
      Height          =   255
      Left            =   5280
      TabIndex        =   46
      Top             =   2520
      Width           =   375
   End
   Begin MSMask.MaskEdBox MskHomeTel 
      Height          =   285
      Left            =   1920
      TabIndex        =   42
      Top             =   2160
      Width           =   1355
      _ExtentX        =   2381
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "(###)-###-####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtTaxFree 
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   9
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox tbfld 
      Height          =   855
      Index           =   8
      Left            =   1800
      TabIndex        =   16
      Top             =   6600
      Width           =   4935
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   10
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      Width           =   3200
   End
   Begin SSDataWidgets_A.SSDBOptSet Optbtn 
      Height          =   240
      Left            =   1800
      TabIndex        =   13
      Top             =   5325
      Width           =   1230
      _Version        =   196611
      _ExtentX        =   2170
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No"
      BackColor       =   -2147483643
      Enabled         =   0   'False
      Cols            =   2
      IndexSelected   =   1
      NumberOfButtons =   2
      Buttons.Button(0).OptionValue=   "-1"
      Buttons.Button(0).Caption=   "Yes"
      Buttons.Button(0).Mnemonic=   89
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   33
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   35
      Buttons.Button(0).PictureRight=   34
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   40
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(1).OptionValue=   "0"
      Buttons.Button(1).Caption=   "No"
      Buttons.Button(1).Mnemonic=   1102
      Buttons.Button(1).Value=   -1  'True
      Buttons.Button(1).TextLeft=   56
      Buttons.Button(1).TextRight=   70
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   41
      Buttons.Button(1).ButtonRight=   54
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   72
      Buttons.Button(1).PictureRight=   71
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   41
      Buttons.Button(1).ButtonToColRight=   81
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).ButtonBitmapID=   2
      Buttons.Button(1).Column=   1
   End
   Begin VB.CommandButton cmdJoint 
      Caption         =   "&Joint"
      Height          =   300
      Left            =   2520
      TabIndex        =   17
      Top             =   7605
      Width           =   975
   End
   Begin VB.TextBox tbJntMode 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Text            =   "Program Indicator :0 - SIS009 -Cancelled; 1 - SIS009 - Saved"
      Top             =   6645
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   2
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   3
      ToolTipText     =   "Enter last name of person or name of company"
      Top             =   1080
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   300
      Left            =   3600
      TabIndex        =   18
      Top             =   7605
      Width           =   975
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   7
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   10
      ToolTipText     =   "Enter Address Line 5"
      Top             =   4485
      Width           =   3375
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "Select a client category from the list."
      Top             =   360
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
      FieldSeparator  =   ","
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "Category"
      Columns(0).Name =   "Category"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   8
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "Type"
      Columns(1).Name =   "Type"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   1
      _ExtentX        =   2566
      _ExtentY        =   503
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5760
      TabIndex        =   20
      Top             =   7605
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "C&ommit"
      Height          =   300
      Left            =   4680
      TabIndex        =   19
      Top             =   7605
      Width           =   975
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   2
      ToolTipText     =   "Enter the Company's Tax Reference Number"
      Top             =   720
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox tbfld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   0
      ToolTipText     =   "Use generate number or enter your own unique client Number"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   6
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   9
      ToolTipText     =   "Enter Address Line 4"
      Top             =   4125
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   5
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   8
      ToolTipText     =   "Enter Address Line 3"
      Top             =   3765
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   4
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   7
      ToolTipText     =   "Enter Address line 2"
      Top             =   3405
      Width           =   3375
   End
   Begin VB.TextBox tbfld 
      Height          =   285
      Index           =   3
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   6
      ToolTipText     =   "Enter Address Line 1"
      Top             =   3045
      Width           =   3375
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      ToolTipText     =   "Select a client category from the list."
      Top             =   4965
      Width           =   1695
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
      Columns(0).Width=   4233
      Columns(0).Caption=   "Category"
      Columns(0).Name =   "Category"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   8
      Columns(1).Width=   2461
      Columns(1).Caption=   "Type"
      Columns(1).Name =   "Type"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   2
      _ExtentX        =   2990
      _ExtentY        =   503
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   12
      ToolTipText     =   "Select a client category from the list."
      Top             =   4965
      Width           =   1695
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
      Columns(0).Width=   4048
      Columns(0).Caption=   "Tax Description"
      Columns(0).Name =   "Tax Description"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   8
      Columns(1).Width=   1296
      Columns(1).Caption=   "Type"
      Columns(1).Name =   "Type"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   2
      _ExtentX        =   2990
      _ExtentY        =   503
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin MSMask.MaskEdBox TxtEffDate 
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   15
      ToolTipText     =   "Enter Effective date in format dd-mm-yyyy"
      Top             =   6120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskWorkTel 
      Height          =   285
      Left            =   3720
      TabIndex        =   43
      Top             =   2160
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "(###)-###-####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskCellNo 
      Height          =   285
      Left            =   1920
      TabIndex        =   44
      Top             =   2520
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "(###)-###-####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblLabels 
      Caption         =   "Residence Declared"
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
      Index           =   17
      Left            =   3480
      TabIndex        =   45
      Top             =   2520
      Width           =   1860
   End
   Begin VB.Label Label2 
      Caption         =   "Tax Free Limit's Effective Date:"
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
      Left            =   0
      TabIndex        =   41
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Tax Free Limit:"
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
      Left            =   0
      TabIndex        =   40
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   3360
      X2              =   3600
      Y1              =   2400
      Y2              =   2160
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cell Number(Alerts):"
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
      Index           =   15
      Left            =   120
      TabIndex        =   39
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Tele Home/Work:"
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
      Left            =   120
      TabIndex        =   38
      Top             =   2160
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Email Address:"
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
      Left            =   120
      TabIndex        =   37
      Top             =   1800
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "TRN:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   36
      Top             =   1440
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Joint:"
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
      Index           =   4
      Left            =   0
      TabIndex        =   35
      Top             =   5325
      Width           =   1620
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
      Index           =   13
      Left            =   4920
      TabIndex        =   34
      Top             =   5325
      Width           =   1500
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   9480
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Category:"
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
      Left            =   0
      TabIndex        =   32
      Top             =   4965
      Width           =   1620
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Tax Class:"
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
      Left            =   3720
      TabIndex        =   31
      Top             =   4965
      Width           =   1020
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
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
      Index           =   8
      Left            =   3960
      TabIndex        =   30
      Top             =   5325
      Width           =   780
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Notes:"
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
      Left            =   360
      TabIndex        =   29
      Top             =   6600
      Width           =   1380
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name:"
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
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Client Type:"
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
      Left            =   3840
      TabIndex        =   27
      Top             =   360
      Width           =   1260
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   9480
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   240
      Y2              =   240
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
      Left            =   480
      TabIndex        =   25
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6840
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "First Names:"
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
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
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
      TabIndex        =   23
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Client Number:"
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
      TabIndex        =   22
      Top             =   360
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
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
      Index           =   3
      Left            =   1080
      TabIndex        =   21
      Top             =   3045
      Width           =   735
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
      Height          =   372
      Index           =   0
      Left            =   600
      TabIndex        =   26
      Top             =   0
      Width           =   6132
   End
End
Attribute VB_Name = "frmSIS002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer, ifirstime As Integer
Dim rsMain As ADODB.Recordset
Dim rsCmp As ADODB.Recordset
Dim rsCat As ADODB.Recordset
Dim rsTax As ADODB.Recordset
Dim rsUnused As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim iOpenMain As Boolean
Dim iOpenCmp As Integer
Dim iOpenCat As Integer
Dim iOpenTax As Integer
Dim iOpenUnusd As Integer
Dim OpenErr As Integer
Dim strTable As String
Dim strRecNO As String
Dim iNewAcct As Long
Dim ChgSw As Boolean
Dim NoteSw As Boolean

Function IsValid() As Integer
Dim iErr As String
Dim i As Integer
IsValid = True
'--
'If tbfld(0) = "" Then  ' Client Number
'   iErr = 93
'   tbfld(0).SetFocus
'   GoTo Validate_Err
' End If
 '--
 If dbc(0) = "" Then 'client type
   iErr = "Invalid Client Type"
   dbc(0).SetFocus
   GoTo Validate_Err
 End If
 '--
 If dbc(0).Columns(1).Text = "P" Then 'Person
    If tbfld(1) = "" Then 'First Name
       iErr = "First Name Missing"
       tbfld(1).SetFocus
       GoTo Validate_Err
    Else
       If tbfld(2) = "" Then ' last name
         iErr = "Last Name Missing"
         tbfld(2).SetFocus
         GoTo Validate_Err
       End If
    End If
 Else
    If tbfld(2) = "" Then 'company Name
       iErr = "Company Name Missing"
       tbfld(2).SetFocus
       GoTo Validate_Err
    End If
 End If
 If tbfld(10) <> "" Then
    If ValidEmail(tbfld(10)) = False Then
       iErr = "Email address is invalid"
       GoTo Validate_Err
    End If
 End If
    
 '--
 If tbfld(3) = "" Then ' Address Line 1
   iErr = "Address Line 1 Missing"
   tbfld(3).SetFocus
   GoTo Validate_Err
 End If
 tbfld(3) = Trim(tbfld(3))
 '--
 If tbfld(4) = "" Then  ' Address Line 2
   iErr = "Address Line 2 Missing"
   tbfld(4).SetFocus
   GoTo Validate_Err
 End If
 tbfld(4) = Trim(tbfld(4))
 tbfld(5) = Trim(tbfld(5))
 tbfld(6) = Trim(tbfld(6))
 tbfld(7) = Trim(tbfld(7))
 i = Len(tbfld(9).Text)
 If tbfld(9).Text <> "" Then
    If Not IsNumeric(tbfld(9).Text) Then
       iErr = "TRN must be numeric"
       tbfld(9).SetFocus
       GoTo Validate_Err
    End If
    If Len(tbfld(9).Text) <> 9 Then
       iErr = "TRN should be Nine digits"
       tbfld(9).SetFocus
       GoTo Validate_Err
    End If
 End If
 
  If MskHomeTel.ClipText <> "" Then
    If Not IsNumeric(MskHomeTel.ClipText) Then
       iErr = "Home Telephone must be numeric"
       MskHomeTel.SetFocus
       GoTo Validate_Err
    End If
    If Len(MskHomeTel.ClipText) > 6 Then
       iErr = vbNullString
    Else
       iErr = "Home telephone should be at least 7 digits"
       MskHomeTel.SetFocus
       GoTo Validate_Err
    End If
 End If
 
 If MskWorkTel.ClipText <> "" Then
    If Not IsNumeric(MskWorkTel.ClipText) Then
       iErr = "Work Telephone must be numeric"
       MskWorkTel.SetFocus
       GoTo Validate_Err
    End If
    If Len(MskWorkTel.ClipText) > 6 Then
       iErr = vbNullString
    Else
       iErr = "Work telephone should be at least 7 digits"
       MskWorkTel.SetFocus
       GoTo Validate_Err
    End If
 End If
 
 If MskCellNo.ClipText <> "" Then
    If Not IsNumeric(MskCellNo.ClipText) Then
       iErr = "Mobile Telephone must be numeric"
       MskCellNo.SetFocus
       GoTo Validate_Err
    End If
    If Len(MskCellNo.ClipText) > 6 Then
       iErr = vbNullString
    Else
       iErr = "Mobile telephone should be at least 7 digits"
       MskCellNo.SetFocus
       GoTo Validate_Err
    End If
 End If
 
 '--
 If dbc(1) = "" Then 'Account Category
   iErr = "Account Category Missing"
   dbc(1).SetFocus
   GoTo Validate_Err
 End If
 '--
 If dbc(2) = "" Then ' Tax Code
   iErr = "Tax Code Missing"
   dbc(2).SetFocus
   GoTo Validate_Err
 End If
 '--
Validate_Exit:
   Exit Function
'--
Validate_Err:
  MsgBox iErr, vbOKOnly, "Client Account"
  IsValid = False
  GoTo Validate_Exit
'--
End Function

Private Sub cmdCancel_Click()
Dim ErrWarn As New cLstWarn
Dim iSeqKey As Integer
If Isloaded("frmSIS009") Then
  X = ErrWarn.ListWarn()
  frmSIS009.Show
  Exit Sub
Else
  Shutdown
  Unload Me
End If
End Sub

Private Sub cmdClear_Click()

If gblOptions = 1 Then
   ClearScreen
   tbfld(0).Text = iNewAcct
   tbfld(0).SetFocus
Else
   ClearScreen
   dbc(1).SetFocus
End If
End Sub

Private Sub CmdDocuments_Click()
If iOpenMain = True Then
   gblFileKey = rsMain!CliName
   gblHold = tbfld(0)
Else
   gblFileKey = "0"
End If
FrmShareholderDocuments.Show 0
End Sub

Private Sub cmdJoint_Click()
tbJntMode = "0"
tbJntMode = Trim(tbJntMode)
frmSIS009.Show 0
End Sub

Private Sub cmdUpdate_Click()
Dim strChg As Integer, FullName As String
Dim i As Integer, imsg As Integer, sField As String
Dim TaxFree As Currency
Dim HomeTel As Double
Dim WorkTel As Double
Dim CellPhone As Double

Dim ClientID As Double
'On Error GoTo cmdUpdate_Err
If gblOptions = 2 Then
   Call CheckForChanges
   If ChgSw = True Then
      If NoteSw = False Then
         MsgBox "Add a note before saving"
         tbfld(8).SetFocus
         GoTo Done
      Else
          NoteSw = False
      End If
   End If
End If

If IsValid Then
  '--
  If gblOptions = 1 Then ' we are adding a new client
     ClientID = 0
  Else
     ClientID = tbfld(0)
  End If
  If dbc(0).Columns(1).Text = "P" Then
      FullName = Trim(tbfld(2)) & "," & Trim(tbfld(1))
  Else
      FullName = Trim(tbfld(2))
  End If
  If IsNull(TxtTaxFree) Or LenB(TxtTaxFree) = 0 Then
     TaxFree = 0
  Else
     TaxFree = CCur(TxtTaxFree)
  End If
  If LenB(TxtEffDate(0).Text) = 0 Then
     TxtEffDate(0).Text = vbNullString
  End If
  If Len(MskHomeTel.ClipText) = 0 Then
     HomeTel = 0
  Else
     HomeTel = CDbl(MskHomeTel.ClipText)
  End If
  
  If Len(MskWorkTel.ClipText) = 0 Then
     WorkTel = 0
  Else
     WorkTel = CDbl(MskWorkTel.ClipText)
  End If
  
  If Len(MskCellNo.ClipText) = 0 Then
     CellPhone = 0
  Else
     CellPhone = CDbl(MskCellNo.ClipText)
  End If
  
  i = RunSP(SpCon, "usp_Sis002Update", 0, ClientID, dbc(0).Columns(1).Text, FullName, _
      tbfld(9), tbfld(3), tbfld(4), tbfld(5), tbfld(6), tbfld(7), _
      dbc(1).Columns(1).Text, dbc(2).Columns(1).Text, _
      CBool(Optbtn.OptionValue), tbfld(8).Text, gblLoginName, _
      tbfld(10).Text, HomeTel, WorkTel, CellPhone, TaxFree, TxtEffDate(0))
    '--
    '--
  If gblOptions = 1 Then
     If tbJntMode = "2" Then
        frmSIS009.Visible = True
        frmSIS009.cmdCancel = True
        tbJntMode = 0
     End If
     InitAddNew
  Else
     Shutdown
     Unload Me
  End If
End If
'---

Done:
 Exit Sub
'--
cmdUpdate_Err:
  MsgBox Err & " " & Err.Description, vbOKOnly, "SIS002/cmdUpdate"
  Shutdown
  Unload Me
End Sub

Private Sub dbc_InitColumnProps(Index As Integer)
On Error GoTo dbc_InitColumnProps_Err
  Select Case Index
  Case 0
  '----  load client types into combo ---
  '--------------------------------------
  With dbc(0)
     .RemoveAll
     .AddItem "Company,C"
     .AddItem "Person,P"
  End With
  '---
  Case 1   '  Category
     'rsCat.Requery
     dbc(Index).RemoveAll
     While Not rsCat.EOF
        dbc(Index).AddItem rsCat!catdesc & Chr(9) & rsCat!CatCode
        rsCat.MoveNext
     Wend
   Case 2
      'rsTax.Requery
      dbc(Index).RemoveAll
      While Not rsTax.EOF
         dbc(Index).AddItem rsTax!RESCTRY & Chr(9) & rsTax!ResCode
         rsTax.MoveNext
      Wend
   End Select
   Exit Sub
dbc_InitColumnProps_Err:
  MsgBox "SIS002/dbc_InitProps"
  Unload Me
End Sub

Private Sub dbc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
   KeyCode = 0
   Select Case Index
   Case 0 'Client Type
      If dbc(0).Columns(1).Text = "P" Then
         tbfld(1).Visible = True
         tbfld(1).SetFocus
      Else
         tbfld(2).Visible = True
         tbfld(2).SetFocus
      End If
      '--
   Case 1
      dbc(2).SetFocus
   Case 2
      tbfld(8).SetFocus
   Case Else
   End Select
Case vbKeyUp
    KeyCode = 0
    Select Case Index
    Case 0
      If gblOptions = 1 Then
        tbfld(0).SetFocus
      End If
    Case 1
      tbfld(7).SetFocus
    Case 2
      dbc(1).SetFocus
    Case Else
    End Select
End Select
End Sub

Private Sub dbc_LostFocus(Index As Integer)
Select Case Index
Case 0
   If dbc(0).Columns(1).Text = "C" Then
       lblLabels(9).Visible = True
       lblLabels(9).Caption = "Company Name:"
       tbfld(2).Visible = True
       '--
       lblLabels(16).Visible = False
       tbfld(1).Visible = False
       
   Else
       lblLabels(16).Caption = "First Names:"
       lblLabels(9).Caption = "Last Name:"
       lblLabels(9).Visible = True
       lblLabels(16).Visible = True
       tbfld(1).Visible = True
       tbfld(2).Visible = True
       
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
ChgSw = False
NoteSw = False
tbfld(9) = ""

If OpenErr = True Then
  Shutdown
  Unload Me
  GoTo Form_Activate_Exit
End If
'--
If ifirstime = 0 Then
   ifirstime = 1
   '--
   If gblOptions = 2 Then
      UpdateScreen
      Me.Caption = "Edit Client Account"
      tbfld(0).Enabled = False
      If gblUserLevel <> 1 Then
         tbfld(1).Enabled = False
         tbfld(2).Enabled = False
      End If
   End If
   '--
End If
'--
Form_Activate_Exit:
  Exit Sub
Form_Activate_Err:
 If Err = -2147168242 Then ' no current transactions
   Resume 0
 Else
   MsgBox "SIS002/Activate"
   Exit Sub
 End If
End Sub

Private Sub Form_Load()
On Error GoTo FL_ERR
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
 ifirstime = 0
'--
csvCenterForm Me, gblMDIFORM
'-----------------------------------
'''Set spcon = New ADODB.Connection

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
iOpenCmp = False
iOpenMain = False
iOpenCat = False
iOpenTax = False
iOpenUnusd = False
'-----
'----------------------------
'---- open recordsets -----
'----------------------------
'--
If gblOptions = 1 Then
   gblFileKey = 0
End If
Set rsCat = RunSP(SpCon, "usp_SIS002Load", 1, gblOptions, CDbl(gblFileKey))
iOpenCat = True
Set rsTax = rsCat.NextRecordset
iOpenTax = True
'----------------------------------------
' create SQL for selecting record to edit
'----------------------------------------
If gblOptions = 1 Then
     '-----------------------------------
     '-- open files used by add mode only
     '-----------------------------------
    Set rsUnused = rsCat.NextRecordset
    iOpenUnusd = True
    Set rsCmp = rsCat.NextRecordset
    iOpenCmp = True
    InitAddNew
Else
    Set rsMain = rsCat.NextRecordset
    iOpenMain = True
End If

'--
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox "SIS002/Load"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
End Sub
Private Sub UpdateScreen()
Dim i As Integer, bm As Variant
Dim a As Double

 With rsMain
    If !Joint = True Then
       Optbtn.IndexSelected = 0
    Else
       Optbtn.IndexSelected = 1
    End If
    tbfld(0).Text = !ClientID
    tbfld(3).Text = !CliAddr1
    tbfld(4).Text = !CliAddr2
    If IsNull(!trn) Or !trn = "         " Then
       tbfld(9).Text = ""
    Else
       tbfld(9).Text = !trn
    End If
    If Not IsNull(!CliAddr3) Then
        tbfld(5).Text = !CliAddr3
    End If
    If Not IsNull(!CliAddr4) Then
      tbfld(6).Text = !CliAddr4
    End If
    If Not IsNull(!CliAddr5) Then
       tbfld(7).Text = !CliAddr5
    End If
    '-- get client type --
    '---------------------
    dbc(0).MoveFirst
    For i = 0 To dbc(0).Rows - 1
       bm = dbc(0).GetBookmark(i)
       If dbc(0).Columns(1).CellText(bm) = !CliType Then
          dbc(0).Bookmark = dbc(0).GetBookmark(i)
          dbc(0) = dbc(0).Columns(0).CellText(bm)
          Exit For
       End If
    Next i
    '-- set correct name depending on type ---
    '-----------------------------------------
    If dbc(0).Columns(1).Text = "C" Then
       lblLabels(9).Visible = True
       lblLabels(9).Caption = "Company Name:"
       tbfld(2) = !CliName
       tbfld(2).Visible = True
       '--
       lblLabels(16).Visible = False
       tbfld(1).Visible = False
    Else
       lblLabels(16).Caption = "First Names:"
       lblLabels(9).Caption = "Last Name:"
       If SplitCliName(!CliName) = True Then
            tbfld(1) = gblHold
            tbfld(2) = gblFileKey
       Else
           tbfld(1) = ""
           tbfld(2) = !CliName
       End If
       lblLabels(9).Visible = True
       lblLabels(16).Visible = True
       tbfld(1).Visible = True
       tbfld(2).Visible = True
    End If
    '--   get category  -----
    '------------------------
    dbc(1).MoveFirst
    For i = 0 To dbc(1).Rows - 1
       bm = dbc(1).GetBookmark(i)
       If dbc(1).Columns(1).CellText(bm) = !CatCode Then
          dbc(1).Bookmark = dbc(1).GetBookmark(i)
          dbc(1) = dbc(1).Columns(0).CellText(bm)
          Exit For
       End If
    Next i
    '-- get tax -------
    '------------------
    If Not IsNull(!HomeTel) Then
        MskHomeTel = Format(!HomeTel, "(000)-###-####")
    End If
    If Not IsNull(!WorkTel) Then
        MskWorkTel = Format(!WorkTel, "(000)-###-####")
    End If
    If Not IsNull(!CellPhone) Then
        MskHomeTel = Format(!CellPhone, "(000)-###-####")
    End If
    
    dbc(2).MoveFirst
    For i = 0 To dbc(2).Rows - 1
       bm = dbc(2).GetBookmark(i)
       If dbc(2).Columns(1).CellText(bm) = !ResCode Then
          dbc(2).Bookmark = dbc(2).GetBookmark(i)
          dbc(2) = dbc(2).Columns(0).CellText(bm)
          Exit For
       End If
    Next i
    '--
    lblLabels(13) = !shares
    '--
    If Not IsNull(!EmailAdd) Then tbfld(10).Text = !EmailAdd
    If Not IsNull(!Remarks) Then tbfld(8).Text = !Remarks
    If Not IsNull(!EffectiveDate) Then
       TxtEffDate(0).Text = !EffectiveDate
    Else
       TxtEffDate(0).Text = vbNullString
    End If
    
    TxtTaxFree = !TaxFree
    If !Residence = True Then
       ChkResidence = 1
    Else
       ChkResidence = 0
    End If
    If !Deceased = True Then
       ChkDeceased = 1
    Else
       ChkDeceased = 0
    End If
 End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub tbfld_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn, vbKeyDown
   KeyCode = 0
   Select Case Index
   Case 0
      dbc(0).SetFocus
   Case 1 To 6
      tbfld(Index + 1).SetFocus
   Case 7
      dbc(1).SetFocus
   Case 8
      cmdUpdate.SetFocus
   Case Else
   End Select
Case vbKeyUp
   KeyCode = 0
   Select Case Index
   Case 1
      dbc(0).SetFocus
   Case 2
      If dbc(0).Columns(1).Text = "P" Then
      tbfld(1).SetFocus
     Else
       dbc(0).SetFocus
     End If
   Case 3 To 7
       tbfld(Index - 1).SetFocus
   Case 8
       dbc(2).SetFocus
   Case Else
   End Select
Case Else
End Select
End Sub

Private Sub ClearScreen()
  For X = 0 To 8
    tbfld(X).Text = ""
  Next
  For X = 0 To 2
     dbc(X).Text = ""
  Next
  TxtTaxFree = 0
  TxtEffDate(0).Text = vbNullString
  If gblOptions = 2 Then
     UpdateScreen
     dbc(0).SetFocus
  End If
End Sub

Private Sub InitAddNew()
On Error GoTo InitAddNew_Err
ClearScreen
Me.Caption = "New Client Account"
'--------------------------
'-- get next account number
'--------------------------
'rsUnused.Requery
If Not rsUnused.EOF Then
     With rsUnused
      .MoveFirst
      iNewAcct = !UNUSED
      '.Delete
     End With
 Else
    'rsCmp.Requery
    If Not rsCmp.EOF Then
       iNewAcct = rsCmp!NEXTACCT
    End If
End If
tbfld(0).Text = CStr(iNewAcct)
lblLabels(13).Visible = False
lblLabels(8).Visible = False
InitAddNew_Exit:
  Exit Sub
InitAddNew_Err:
  MsgBox "SIS002/InitAddNew"
  GoTo InitAddNew_Exit
End Sub

Private Sub Shutdown()
If SpCon.State = 1 Then
If iOpenMain = True Then rsMain.Close
If iOpenCat = True Then rsCat.Close
If iOpenTax = True Then rsTax.Close
If gblOptions = 1 Then
  If iOpenCmp = True Then rsCmp.Close
  If iOpenUnusd = True Then rsUnused.Close
End If
End If
Set rsMain = Nothing
Set rsCmp = Nothing
Set rsCat = Nothing
Set rsTax = Nothing
Set rsUnused = Nothing
End Sub

Private Sub CallSIS009()
    tbJntMode = "0"
    tbJntMode = Trim(tbJntMode)
    frmSIS002.Visible = False
    frmSIS001.Visible = False
    SpCon.Close
    frmSIS009.Show 0
End Sub

Private Sub CheckForChanges()
Dim i As Integer, bm As Variant
 With rsMain
    If !Joint = True Then
       If Optbtn.IndexSelected = 1 Then
          ChgSw = True
       End If
    Else
       If Optbtn.IndexSelected = 0 Then
          ChgSw = True
       End If
    End If
    
    If tbfld(3).Text <> !CliAddr1 Then
       ChgSw = True
    End If
    
    If tbfld(4).Text <> !CliAddr2 Then
       ChgSw = True
    End If
    
    If IsNull(!trn) Or !trn = "         " Then
       If tbfld(9).Text <> "" Then
          ChgSw = True
       End If
    Else
       If tbfld(9).Text <> !trn Then
          ChgSw = True
       End If
    End If
    
    If IsNull(!CliAddr3) Then
       If tbfld(5).Text = "" Then
          ChgSw = True
       End If
    Else
        If tbfld(5).Text <> !CliAddr3 Then
           ChgSw = True
        End If
    End If
    
    If IsNull(!CliAddr4) Then
       If tbfld(6).Text = "" Then
          ChgSw = True
       End If
    Else
      If tbfld(6).Text <> !CliAddr4 Then
         ChgSw = True
      End If
    End If
    
    If IsNull(!CliAddr5) Then
       If tbfld(7).Text = "" Then
          ChgSw = True
       End If
    Else
       If tbfld(7).Text <> !CliAddr5 Then
          ChgSw = True
       End If
    End If
    
    '-----------------------------------------
    If dbc(0).Columns(1).Text = "C" Then
       If tbfld(2) <> !CliName Then
          ChgSw = True
       End If
    Else
       If SplitCliName(!CliName) = True Then
            If tbfld(1) <> gblFileKey Then
               ChgSw = True
            End If
            If tbfld(2) <> gblHold Then
               ChgSw = True
            End If
       Else
           If tbfld(1) <> "" Then
              ChgSw = True
           End If
           If tbfld(2) = !CliName Then
              ChgSw = True
           End If
       End If
     End If
   If dbc(1).Columns(1).Text <> !CatCode Then
      ChgSw = True
   End If
   If dbc(2).Columns(1).Text <> !ResCode Then
      ChgSw = True
   End If
   If CCur(TxtTaxFree) <> !TaxFree Then
      ChgSw = True
   End If
   If TxtEffDate(0).Text <> !EffectiveDate Then
      ChgSw = True
   End If
   
   
    '--   get category  -----
    '------------------------
  '  dbc(1).MoveFirst
  '  For i = 0 To dbc(1).Rows - 1
  '     bm = dbc(1).GetBookmark(i)
  '     If dbc(1).Columns(1).CellText(bm) = !CatCode Then
  '        dbc(1).Bookmark = dbc(1).GetBookmark(i)
  '        dbc(1) = dbc(1).Columns(0).CellText(bm)
  '        Exit For
  '     End If
  '  Next i
    '-- get tax -------
    '------------------
   ' dbc(2).MoveFirst
   ' For i = 0 To dbc(2).Rows - 1
   '    bm = dbc(2).GetBookmark(i)
   '    If dbc(2).Columns(1).CellText(bm) = !ResCode Then
   '       dbc(2).Bookmark = dbc(2).GetBookmark(i)
   '       dbc(2) = dbc(2).Columns(0).CellText(bm)
   '       Exit For
   '    End If
   ' Next i
    '--
    
    '--
    If IsNull(!Remarks) Then
       If tbfld(8).Text <> "" Then
          NoteSw = True
       End If
    Else
       If tbfld(8).Text <> !Remarks Then
          NoteSw = True
       End If
    End If
 End With
End Sub

Private Sub TxtTaxFree_KeyPress(KeyAscii As Integer)
Dim i As Integer
If KeyAscii = 46 Or (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
   Else
   KeyAscii = 0
   GoTo Exit_TxtTaxFree_KeyPress
End If
If KeyAscii = 46 Then
   i = InStr(1, TxtTaxFree, ".")
   If i > 0 Then
      KeyAscii = 0
   End If
   GoTo Exit_TxtTaxFree_KeyPress
End If

Exit_TxtTaxFree_KeyPress:
Exit Sub

Err_TxtTaxFree_KeyPress:
MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Tax Free Limit Input"
Resume Exit_TxtTaxFree_KeyPress

End Sub
