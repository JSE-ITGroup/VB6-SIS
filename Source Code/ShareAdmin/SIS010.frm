VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSIS010 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Mandates"
   ClientHeight    =   6615
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   14475
   Icon            =   "SIS010.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   14475
   Begin VB.Frame FmeGeneral 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Bank Payment Details"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   7320
      TabIndex        =   31
      ToolTipText     =   "Complete for all bank payments"
      Top             =   3360
      Width           =   6855
      Begin VB.TextBox TxtAccountNo 
         Height          =   285
         Left            =   2040
         MaxLength       =   16
         TabIndex        =   10
         ToolTipText     =   "Enter an account number to credit the shareholder's payments. This is optional."
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtAccountName 
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   11
         ToolTipText     =   "Enter the name of the account if different from the shareholder."
         Top             =   1080
         Width           =   3495
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBAcctType 
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   1440
         Width           =   2895
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
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4048
         Columns(0).Caption=   "Account Type"
         Columns(0).Name =   "Account Type"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1640
         Columns(1).Caption=   "Type Code"
         Columns(1).Name =   "Type Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 0"
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Account Type:"
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
         Left            =   720
         TabIndex        =   35
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
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
         Index           =   7
         Left            =   480
         TabIndex        =   33
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
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
         Left            =   480
         TabIndex        =   32
         Top             =   1080
         Width           =   1500
      End
   End
   Begin VB.Frame FmeOther 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Other Mandate Types"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Complete for cheque payments"
      Top             =   3360
      Width           =   6855
      Begin VB.TextBox tbfld 
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Enter the name of the bank or recipient to receive the payment."
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox TxtCurrency 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   34
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox tbfld 
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   5
         ToolTipText     =   "Enter the first line of the mandates address."
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox tbfld 
         Height          =   285
         Index           =   7
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   9
         ToolTipText     =   "Enter Address Line 5"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox tbfld 
         Height          =   285
         Index           =   6
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   8
         ToolTipText     =   "Enter Address Line 4"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox tbfld 
         Height          =   285
         Index           =   5
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   7
         ToolTipText     =   "Enter Address Line 3"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox tbfld 
         Height          =   285
         Index           =   4
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   6
         ToolTipText     =   "Enter Address line 2"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   480
         TabIndex        =   40
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   360
         TabIndex        =   30
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame FmeACH 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ACH Mandates"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      TabIndex        =   26
      ToolTipText     =   "Complete for ACH payments only"
      Top             =   1680
      Width           =   6855
      Begin SSDataWidgets_B.SSDBCombo SSDBACHBanks 
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   5175
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
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   7144
         Columns(0).Caption=   "BankID"
         Columns(0).Name =   "BankID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Bank Name"
         Columns(1).Name =   "Bank Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   9128
         _ExtentY        =   450
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 0"
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBACHBranches 
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   840
         Width           =   5175
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
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   10134
         Columns(0).Caption=   "Branch Name"
         Columns(0).Name =   "Branch Name"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2170
         Columns(1).Caption=   "Branch Code"
         Columns(1).Name =   "Branch Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   635
         Columns(2).Caption=   "Financail Inst ID"
         Columns(2).Name =   "Financail Inst ID"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   9128
         _ExtentY        =   450
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 0"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Branch"
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
         TabIndex        =   28
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bank"
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
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "&Terminate Mandate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Clears the screen and resets it if in edit mode."
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12960
      TabIndex        =   15
      ToolTipText     =   "Cancels changes and returns to Account maintenance."
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6360
      TabIndex        =   14
      ToolTipText     =   "Update Joint Table for saving to disk by Accounts Maintainace"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox tbfld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   16
      ToolTipText     =   "Use generate number or enter your own unique client Number"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox tbfld 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   17
      ToolTipText     =   "Enter Address line 2"
      Top             =   840
      Width           =   4335
   End
   Begin SSDataWidgets_B.SSDBCombo dbc 
      Height          =   255
      Index           =   1
      Left            =   9000
      TabIndex        =   0
      Top             =   600
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
      Columns(0).Width=   3200
      Columns(0).Caption=   "Description"
      Columns(0).Name =   "Description"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
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
   Begin VB.Frame FmeFIM 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Financail Institution Mandate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "Complete if Finacle lodgements and Other Banks payments"
      Top             =   1680
      Width           =   6855
      Begin VB.TextBox TxtBank 
         Height          =   285
         Left            =   1440
         TabIndex        =   41
         Top             =   840
         Width           =   5175
      End
      Begin SSDataWidgets_B.SSDBCombo dbc 
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         DataFieldList   =   "Column 1"
         AllowInput      =   0   'False
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
      Begin VB.Label lblLabels 
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
         Index           =   12
         Left            =   600
         TabIndex        =   25
         Top             =   480
         Width           =   780
      End
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   255
      Left            =   8400
      TabIndex        =   36
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   16515075
      CurrentDate     =   40731
   End
   Begin MSMask.MaskEdBox meb 
      Height          =   255
      Index           =   1
      Left            =   12360
      TabIndex        =   37
      ToolTipText     =   "Enter the date this mandate ceases."
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   11
      Format          =   "dd-mmm-yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   7320
      TabIndex        =   39
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "End Date:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   11160
      TabIndex        =   38
      Top             =   1080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   360
      Y2              =   6000
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   14400
      X2              =   14400
      Y1              =   360
      Y2              =   6000
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   7080
      X2              =   7080
      Y1              =   360
      Y2              =   6000
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   6960
      X2              =   14400
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   7440
      TabIndex        =   23
      Top             =   600
      Width           =   1500
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   6960
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   14400
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   14400
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   14400
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
      TabIndex        =   21
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Shareholder:"
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
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   18
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
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   14415
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
Function IsValid() As Boolean
On Error GoTo IsValid_Err

IsValid = False

If dbc(1) = "" Then
   MsgBox "Please select a payment method first"
   GoTo Exit_IsValid
End If
If dbc(1).Columns(1).Text = "0" Then
   If ChequeCheck Then
      IsValid = True
   End If
   GoTo Exit_IsValid
End If
If dbc(1).Columns(1).Text = "3" Then
   If ACHCheck Then
      IsValid = True
   End If
   GoTo Exit_IsValid
End If
If FIMCheck Then
   IsValid = True
End If

Exit_IsValid:
Exit Function


IsValid_Err:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on validation"
Resume Exit_IsValid

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

Private Sub cmdTerminate_Click()
On Error GoTo Err_cmdTerminate_Click

X = RunSP(SpCon, "usp_TerminateMandate", 0, CLng(tbfld(0).Text))
If X <> 0 Then
   MsgBox "Termination failed"
   GoTo Exit_cmdTerminate_Click
End If
MsgBox "Termination successfully completed"
ClearScreen

Exit_cmdTerminate_Click:
Exit Sub

Err_cmdTerminate_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on terminating existing mandate"
Resume Exit_cmdTerminate_Click
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Err_CmdUpdate_Click

Dim strChg As Integer
Dim i As Integer
Dim newval As Integer

If IsValid Then
  '--
  If dbc(1).Columns(1).Text = "0" Then
     i = RunSP(SpCon, "usp_MandateUpdate", 0, CInt(dbc(1).Columns(1).Text), tbfld(0), gblLoginName, Format(DTPStartDate, "dd-mmm-yyyy"), tbfld(2), tbfld(3), tbfld(4), tbfld(5), tbfld(6), tbfld(7))
  Else
      If dbc(1).Columns(1).Text = "3" Or dbc(1).Columns(1).Text = "4" Then
         i = RunSP(SpCon, "usp_MandateUpdate", 0, CInt(dbc(1).Columns(1).Text), tbfld(0), gblLoginName, Format(DTPStartDate, "dd-mmm-yyyy"), SSDBACHBanks.Columns(1).Text, SSDBACHBranches.Columns(1).Text, TxtAccountNo, TxtAccountName, SSDBAcctType.Columns(1).Text, " ")
      Else
         i = RunSP(SpCon, "usp_MandateUpdate", 0, CInt(dbc(1).Columns(1).Text), tbfld(0), gblLoginName, Format(DTPStartDate, "dd-mmm-yyyy"), dbc(0).Columns(1).Text, " ", TxtAccountNo, TxtAccountName, SSDBAcctType.Columns(1).Text, " ")
      End If
  End If
  If i = 0 Then
     MsgBox "Record sucessfully updated"
  Else
     MsgBox "Update was unsucessfull. Sorry for any inconvienience caused"
     GoTo Exit_CmdUpdate_Click
  End If

cmdCancel_Click
End If

'---

Exit_CmdUpdate_Click:
 Exit Sub
'--
Err_CmdUpdate_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on saving"
Resume Exit_CmdUpdate_Click

End Sub

Private Sub dbc_Click(Index As Integer)
On Error GoTo dbc_Click_Err

Select Case Index
Case 0
     TxtBank = dbc(0).Columns(0).Text
Case 1
     ClearFields
     If dbc(1).Columns(1).Text = "0" Then
        FmeOther.Enabled = True
        FmeACH.Enabled = False
        FmeGeneral.Enabled = False
        FmeFIM.Enabled = False
        tbfld(2) = tbfld(1)
        GoTo Exit_dbc_Click
     End If
     TxtAccountName = tbfld(1)
     If dbc(1).Columns(1).Text = "3" Or dbc(1).Columns(1).Text = "4" Then
        FmeACH.Enabled = True
        FmeGeneral.Enabled = True
        FmeFIM.Enabled = False
        FmeOther.Enabled = False
        GoTo Exit_dbc_Click
     Else
         FmeFIM.Enabled = True
         FmeGeneral.Enabled = True
         FmeACH.Enabled = False
         FmeOther.Enabled = False
         FillBankIDs
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

Private Sub Form_Activate()
'On Error GoTo Err_Form_Activate
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
On Error GoTo FL_ERR
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 lblLabels(0).Caption = gblCompName
 lblLabels(1).Caption = App.Major & "." & App.Minor & "." & App.Revision
'--
csvCenterForm Me, gblMDIFORM
'-----------------------------------

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
Set rsDist = rsMandate.NextRecordset
Set rsACHBanks = rsMandate.NextRecordset
Set rsAccType = rsMandate.NextRecordset
'--------------------

iOpenMan = True
If rsMandate.EOF = True Then
    iMode = 0
    Me.Caption = "New Mandate"
Else
 iMode = 1
  Me.Caption = "Edit Mandate"
  cmdTerminate.Enabled = True
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
tbfld(1).Text = gblHold
'--

If iMode = 0 Then
   DTPStartDate = Format(Date, "dd-mmm-yyyy")
   cmdUpdate.Enabled = True
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
dbc(1).Enabled = False
cmdUpdate.Enabled = False

With rsMandate
     DTPStartDate = !MndStaDte
     If !MndMet = 0 Then 'Disposal option is cheque
        If Not IsNothing(!MndAcntNme) Then tbfld(2) = !MndAcntNme
        If Not IsNothing(!MndName) Then tbfld(2).Text = !MndName
        If Not IsNothing(!MndAddr1) Then tbfld(3).Text = !MndAddr1
        If Not IsNothing(!MndAddr2) Then tbfld(4).Text = !MndAddr2
        If Not IsNothing(!MndAddr3) Then tbfld(5).Text = !MndAddr3
        If Not IsNothing(!MNDADDR4) Then tbfld(6).Text = !MNDADDR4
        If Not IsNothing(!MNDADDR5) Then tbfld(7).Text = !MNDADDR5
        GoTo Exit_UpdateScreen
     End If
      
     If Not IsNothing(!MndAcnt) Then TxtAccountNo = !MndAcnt
     If Not IsNothing(!MndAcntNme) Then TxtAccountName = !MndAcntNme
     
     If !MndMet = 3 Or !MndMet = 4 Then 'Disposal option is ACH (3) or ACH (4)
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
             SSDBACHBanks_Click
        End With
        
        With SSDBACHBranches
             .MoveFirst
             For i = 0 To .Rows - 1
                 bm = .GetBookmark(i)
                 If .Columns(1).CellText(bm) = Trim(rsMandate!MndBranchID) Then
                    .Bookmark = .GetBookmark(i)
                    SSDBACHBranches = .Columns(0).CellText(bm)
                    Exit For
                 End If
             Next i
        End With
      Else
          FillBankIDs
          With dbc(0)
             .MoveFirst
             For i = 0 To .Rows - 1
                 bm = .GetBookmark(i)
                 If .Columns(1).CellText(bm) = Trim(rsMandate!BankID) Then
                    .Bookmark = .GetBookmark(i)
                    dbc(0) = .Columns(1).CellText(bm)
                    TxtBank = .Columns(0).CellText(bm)
                    Exit For
                 End If
             Next i
        End With
     End If
     With SSDBAcctType
             .MoveFirst
             For i = 0 To .Rows - 1
                 bm = .GetBookmark(i)
                 If .Columns(1).CellText(bm) = Trim(rsMandate!MndAcntType) Then
                    .Bookmark = .GetBookmark(i)
                    SSDBAcctType = .Columns(0).CellText(bm)
                    Exit For
                 End If
             Next i
        End With
End With

Exit_UpdateScreen:
Exit Sub
End Sub
Private Sub ClearScreen()
 
'Clear the Other Financial Institution field
For X = 2 To 7
    tbfld(X).Text = ""
Next X
TxtBank = ""
TxtAccountName = ""
TxtAccountNo = ""
SSDBAcctType = ""
SSDBACHBranches = ""
SSDBACHBanks = ""
dbc(0) = ""
dbc(1) = ""
cmdTerminate.Enabled = False
cmdUpdate.Enabled = True
dbc(1).Enabled = True

End Sub

Private Sub Shutdown()
If SpCon.State = 1 Then
   If iOpenMan = True Then rsMandate.Close
   If iOpenBank = True Then rsBank.Close
   rsDist.Close
End If
Set rsMandate = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub
Private Sub SSDBAcctType_InitColumnProps()
On Error GoTo Err_SSDBAcctType_InitColumnProps
SSDBAcctType.RemoveAll
Set rsAccType = RunSP(SpCon, "usp_ListACHAccountTypes", 1)
With rsAccType
    If .EOF Then
       MsgBox "Account Types are missing!. Please contact your Sys Admin"
       GoTo Exit_SSDBAcctType_InitColumnProps
    End If
    Do While Not .EOF
       SSDBAcctType.AddItem !AccountDesc & vbTab & !AccountType
       .MoveNext
    Loop
    .Close
End With
Exit_SSDBAcctType_InitColumnProps:
Exit Sub

Err_SSDBAcctType_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on loading account types"
Resume Exit_SSDBAcctType_InitColumnProps

End Sub

Private Sub SSDBACHBanks_Click()
On Error GoTo Err_SSDBACHBanks_Click

Set rsACHBrchs = RunSP(SpCon, "usp_ListACHBranches", 1, SSDBACHBanks.Columns(1).Text)

With rsACHBrchs
     If .EOF Then
       MsgBox "ACH Branches are missing!. Please contact your Sys Admin"
       GoTo Exit_SSDBACHBanks_Click
     End If
     
     SSDBACHBranches.RemoveAll
     Do While Not .EOF
        SSDBACHBranches.AddItem !BranchName & vbTab & !BranchID
       .MoveNext
     Loop
    .Close
End With

Exit_SSDBACHBanks_Click:
Exit Sub

Err_SSDBACHBanks_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on loading ACH Branches"
Resume Exit_SSDBACHBanks_Click
End Sub

Private Sub SSDBACHBanks_InitColumnProps()
On Error GoTo Err_SSDBACHBanks_InitColumnProps

SSDBACHBanks.RemoveAll
With rsACHBanks
    If Not .EOF Then
      .MoveFirst
    Else
       MsgBox "ACH Banks are missing!. Please contact your Sys Admin"
       GoTo Exit_SSDBACHBanks_InitColumnProps
    End If
    
    Do While Not .EOF
       SSDBACHBanks.AddItem !BankName & vbTab & !BankID
       .MoveNext
    Loop
    .Close
End With

Exit_SSDBACHBanks_InitColumnProps:
Exit Sub

Err_SSDBACHBanks_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on loading ACH Banks"
Resume Exit_SSDBACHBanks_InitColumnProps
End Sub

Function ChequeCheck() As Boolean
ChequeCheck = False

If Len(tbfld(2)) < 11 Then
   MsgBox "Please enter a name to be placed on the cheque"
   tbfld(2).SetFocus
   GoTo Exit_ChequeCheck
End If

ChequeCheck = True

Exit_ChequeCheck:
Exit Function

End Function

Function FIMCheck() As Boolean
FIMCheck = False

If dbc(0) = "" Then
   MsgBox "Please select a Bank first"
   dbc(0).SetFocus
   GoTo Exit_FIMCheck
End If

If Len(TxtAccountNo) < 4 Then
   MsgBox "Please enter a vaild account no"
   TxtAccountNo.SetFocus
   GoTo Exit_FIMCheck
End If

If Len(TxtAccountName) < 11 Then
   MsgBox "Please enter a vaild Bank Account Name"
   TxtAccountName.SetFocus
   GoTo Exit_FIMCheck
End If

If dbc(0).Text = "1" Then
   If FinacleCheck Then
      FIMCheck = True
   End If
   GoTo Exit_FIMCheck
End If

FIMCheck = True

Exit_FIMCheck:
Exit Function
End Function
Function ACHCheck() As Boolean
ACHCheck = False

If SSDBACHBanks = "" Then
   MsgBox "Please select an ACH Bank first"
   SSDBACHBanks.SetFocus
   GoTo Exit_ACHCheck
End If

If SSDBACHBranches = "" Then
   MsgBox "Please select an ACH Branch first"
   SSDBACHBranches.SetFocus
   GoTo Exit_ACHCheck
End If

If Len(TxtAccountNo) < 4 Then
   MsgBox "Please enter a vaild account no"
   TxtAccountNo.SetFocus
   GoTo Exit_ACHCheck
End If

If Len(TxtAccountName) < 11 Then
   MsgBox "Please enter a vaild Bank Account Name"
   TxtAccountName.SetFocus
   GoTo Exit_ACHCheck
End If

If SSDBACHBanks.Columns(1).Text = "077" Then
   If Not FinacleCheck Then
      GoTo Exit_ACHCheck
   End If
End If

If SSDBAcctType = "" Then
   MsgBox "Please select an ACH Account type first"
   SSDBAcctType.SetFocus
   GoTo Exit_ACHCheck
End If

ACHCheck = True

Exit_ACHCheck:
Exit Function

End Function

Function FinacleCheck() As Boolean
FinacleCheck = False

If Len(TxtAccountNo) <> 9 Then
   MsgBox "Account Number must be 9 Characters long"
   GoTo Exit_FinacleCheck
End If

If Not IsNumeric(TxtAccountNo) Then
   MsgBox "Only numbers are allowed"
   GoTo Exit_FinacleCheck
End If

FinacleCheck = True
  
Exit_FinacleCheck:
Exit Function

End Function
Function FillBankIDs()
Dim sRowinfo As String
If dbc(1).Columns(1).Text = "1" Then
   Set rsBank = RunSP(SpCon, "usp_ListBanks", 1, "1")
Else
    Set rsBank = RunSP(SpCon, "usp_ListBanks", 1, "0")
End If
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
End Function
Function ClearFields()
For X = 2 To 7
    tbfld(X).Text = ""
Next X
TxtBank = ""
TxtAccountName = ""
TxtAccountNo = ""
SSDBAcctType = ""
SSDBACHBranches = ""
SSDBACHBanks = ""
dbc(0) = ""
End Function
