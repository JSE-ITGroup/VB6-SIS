VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmShareholderData 
   BackColor       =   &H00808080&
   Caption         =   "Extract Shareholder Data"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   Icon            =   "FrmShareholderData.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   32
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton CmdExcel 
      Caption         =   "Generate Excel File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   31
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Frame FmeFields 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fields to Show"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   7335
      Left            =   7560
      TabIndex        =   11
      Top             =   120
      Width           =   2895
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   120
         TabIndex        =   30
         Top             =   6840
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   120
         TabIndex        =   29
         Top             =   6480
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Office Number"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   28
         Top             =   6120
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Home Telephone"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   27
         Top             =   5760
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Email "
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   26
         Top             =   5400
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "TRN"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   25
         Top             =   5040
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Date Opened"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   24
         Top             =   4680
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Joint Holders' Names"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   23
         Top             =   4320
         Width           =   2535
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Shares"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   22
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Country of Residence"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   2535
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address 5"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address 4"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address 3"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address 2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address 1"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Client ID"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Client Name"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox ChkField 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Client Type"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame FmeCriteria 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Criteria for Selection"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton OptJH 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Yes"
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
         Index           =   1
         Left            =   4440
         TabIndex        =   40
         Top             =   4440
         Width           =   1215
      End
      Begin VB.OptionButton OptJH 
         BackColor       =   &H00E0E0E0&
         Caption         =   "No"
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
         Index           =   0
         Left            =   3000
         TabIndex        =   39
         Top             =   4440
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPAccOpen 
         Height          =   375
         Left            =   2760
         TabIndex        =   38
         ToolTipText     =   "To exclude the Account Open Date from the selection uncheck the date field"
         Top             =   5040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   20643843
         CurrentDate     =   38872
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBRegisterMode 
         Height          =   375
         Left            =   2760
         TabIndex        =   37
         Top             =   5880
         Visible         =   0   'False
         Width           =   3735
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
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6588
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
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBCompare 
         Height          =   375
         Left            =   2760
         TabIndex        =   36
         Top             =   3480
         Width           =   2055
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
         Columns(0).Width=   3200
         Columns(0).Caption=   "Sign"
         Columns(0).Name =   "Sign"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.TextBox TxtShares 
         Height          =   375
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   35
         ToolTipText     =   "Leave blank or include a valid number if the selection will be based on number of shares held"
         Top             =   3480
         Width           =   1335
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBCategory 
         Height          =   375
         Left            =   2760
         TabIndex        =   34
         Top             =   2760
         Width           =   3615
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
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6376
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
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBResCode 
         Height          =   375
         Left            =   2760
         TabIndex        =   33
         Top             =   2040
         Width           =   3615
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
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Country"
         Columns(1).Name =   "Country"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6376
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
      End
      Begin VB.TextBox TxtCliName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         ToolTipText     =   "Enter the name of part of the name to include in the search results. "
         Top             =   1320
         Width           =   3615
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBCliType 
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   600
         Width           =   3615
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
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6376
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
      End
      Begin VB.Label LblRegister 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Register Mode"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   6000
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label LblDteOpened 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Account Open Date"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   5160
         Width           =   2775
      End
      Begin VB.Label LblJoint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Include Joint Holders"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   4440
         Width           =   2775
      End
      Begin VB.Label LblShares 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Shares"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label LblCategory 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label LblResCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Country of Residence"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label LblCliName 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Client Name"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label LblCliType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Client Type"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FrmShareholderData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoResCode As ADODB.Recordset
Dim SpCon As ADODB.Connection
Dim adoCategory As ADODB.Recordset

Private Sub ChkField_Click(Index As Integer)
If ChkField(11).Value = False Then
   OptJH(0).Value = True
Else
   OptJH(1).Value = True
End If
   
End Sub

Private Sub CmdExcel_Click()
'On Error GoTo Err_CmdExcel_Click
Dim adoRst As ADODB.Recordset
Dim i As Integer
Dim ListOfFields As String
Dim NoOfFields As Integer
Dim wSign As String
Dim DateString As String

ListOfFields = ""
NoOfFields = 0
For i = 0 To ChkField.Count - 1
    If ChkField(i).Value = 1 Then
       ListOfFields = ListOfFields & i & ";"
       NoOfFields = NoOfFields + 1
    End If
Next i
If NoOfFields < 1 Then
   MsgBox "Please select fields to display in the spreadsheet"
   GoTo Exit_CmdExcel_Click
End If

wSign = SSDBCompare.Columns(0).Text

If TxtShares = vbNullString Then
   TxtShares = 0
   wSign = "="
End If
If IsNull(DTPAccOpen) Then
   DateString = vbNullString
Else
   DateString = Format(DTPAccOpen, "dd-mmm-yyyy")
End If
If OptJH(0).Value = True Then
   i = 0
Else
   i = 1
End If

Set adoRst = RunSP(SpCon, "usp_ShareholderData", 1, ListOfFields, NoOfFields, SSDBCliType.Columns(0).Text, _
             SSDBResCode.Columns(0).Text, SSDBCategory.Columns(0).Text, wSign, CDbl(TxtShares), _
             DateString, TxtCliName, i)
             
If adoRst.EOF Then
   MsgBox "Sorry, No records were found"
Else
   Call ExportToExcel(adoRst)
End If
adoRst.Close
Set adoRst = Nothing
Exit_CmdExcel_Click:
Exit Sub

Err_CmdExcel_Click:
MsgBox Err.Description, vbOKOnly, "Shareholder Excel File Creation"
GoTo Exit_CmdExcel_Click

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load

frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
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
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Set adoResCode = RunSP(SpCon, "usp_ShareholderDataInit", 1)
Set adoCategory = adoResCode.NextRecordset
OptJH(0).Value = True
DTPAccOpen = Date

Exit_Form_Load:
Exit Sub
Err_Form_Load:
MsgBox Err.Description, vbOKOnly, "Shareholder Data form load"
GoTo Exit_Form_Load

End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
End Sub

Private Sub OptJH_Click(Index As Integer)
If Index = 1 Then
   ChkField(11).Enabled = True
   ChkField(11).Value = 1
Else
   ChkField(11).Enabled = False
   ChkField(11).Value = 0
End If

End Sub

Private Sub SSDBCategory_InitColumnProps()
Dim StrSql As String
Dim X As Integer
SSDBCategory.RemoveAll

Do While Not adoCategory.EOF
   StrSql = adoCategory(0) & vbTab & adoCategory(1)
   SSDBCategory.AddItem StrSql
   adoCategory.MoveNext
Loop
StrSql = "00" & vbTab & "All Categories"
SSDBCategory.AddItem StrSql

With SSDBCategory
    .MoveFirst
    .SelBookmarks.RemoveAll
     For X = 0 To .Rows - 1
     If .Columns(0).Text = "00" Then
        .Text = .Columns(1).Text
        .SelBookmarks.Add .Bookmark
        GoTo CloseRecordset
     End If
     .MoveNext
     Next X
End With

CloseRecordset:
adoCategory.Close
Set adoCategory = Nothing
End Sub

Private Sub SSDBCliType_InitColumnProps()
Dim StrSql As String
Dim X As Integer

With SSDBCliType
     .RemoveAll
     StrSql = "C" & vbTab & "Company" & vbTab
     .AddItem StrSql
     StrSql = "P" & vbTab & "Person"
     .AddItem StrSql
     StrSql = "A" & vbTab & "All Client Types"
     .AddItem StrSql
End With

With SSDBCliType
    .MoveFirst
    .SelBookmarks.RemoveAll
     For X = 0 To .Rows - 1
     If .Columns(0).Text = "A" Then
        .Text = .Columns(1).Text
        .SelBookmarks.Add .Bookmark
        GoTo CloseRecordset
     End If
     .MoveNext
     Next X
End With

CloseRecordset:

End Sub

Private Sub SSDBCompare_InitColumnProps()
Dim StrSql As String
Dim X As Integer

With SSDBCompare
     .RemoveAll
     StrSql = "=" & vbTab & "Equal to"
     .AddItem StrSql
     StrSql = ">" & vbTab & "Greater than"
     .AddItem StrSql
     StrSql = "<" & vbTab & "Less than"
     .AddItem StrSql
End With
With SSDBCompare
    .MoveFirst
    .SelBookmarks.RemoveAll
     For X = 0 To .Rows - 1
     If .Columns(0).Text = ">" Then
        .Text = .Columns(1).Text
        .SelBookmarks.Add .Bookmark
        TxtShares = 0
        GoTo CloseRecordset
     End If
     .MoveNext
     Next X
End With

CloseRecordset:

End Sub

Private Sub SSDBRegisterMode_InitColumnProps()
Dim StrSql As String

With SSDBRegisterMode
     .RemoveAll
     StrSql = "M" & vbTab & "Main Register Mode"
     .AddItem StrSql
     StrSql = "J" & vbTab & "JCSD Register Mode"
     .AddItem StrSql
     StrSql = "T" & vbTab & "T&T Register Mode"
     .AddItem StrSql
     StrSql = "A" & vbTab & "All Register Modes"
     .AddItem StrSql
End With

End Sub

Private Sub SSDBResCode_InitColumnProps()
Dim StrSql As String
SSDBResCode.RemoveAll

Do While Not adoResCode.EOF
   StrSql = adoResCode(0) & vbTab & adoResCode(1)
   SSDBResCode.AddItem StrSql
   adoResCode.MoveNext
Loop
StrSql = "00" & vbTab & "All Countries"
SSDBResCode.AddItem StrSql

With SSDBResCode
    .MoveFirst
    .SelBookmarks.RemoveAll
     For X = 0 To .Rows - 1
     If .Columns(0).Text = "00" Then
        .Text = .Columns(1).Text
        .SelBookmarks.Add .Bookmark
        GoTo CloseRecordset
     End If
     .MoveNext
     Next X
End With

CloseRecordset:
adoResCode.Close
Set adoResCode = Nothing
End Sub
