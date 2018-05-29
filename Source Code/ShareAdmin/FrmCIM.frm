VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCIM 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Cheque Inventory Management"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   Icon            =   "FrmCIM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   33
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00FF00FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   32
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Frame FmeMovement 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   20
      Top             =   5160
      Width           =   14055
      Begin VB.CheckBox ChkPrintery 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Receive from Printery"
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
         Left            =   8400
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPTransferDate 
         Height          =   375
         Left            =   2040
         TabIndex        =   35
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   51838979
         CurrentDate     =   40722
      End
      Begin VB.TextBox TxtChqRange 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         TabIndex        =   25
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox TxtChqEnd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   24
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TxtChqStart 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBTransferOpt 
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   240
         Width           =   5535
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
         Columns(0).Width=   8096
         Columns(0).Caption=   "Transfer Option"
         Columns(0).Name =   "Account Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Currency"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   9763
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
      Begin SSDataWidgets_B.SSDBCombo SSDBMove 
         Height          =   375
         Left            =   11040
         TabIndex        =   29
         Top             =   1080
         Width           =   2775
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
         Columns(0).Width=   6429
         Columns(0).Caption=   "Location"
         Columns(0).Name =   "Account Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Currency"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4895
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
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Date Transferred:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label LblMovement 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Move To:"
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
         Left            =   9720
         TabIndex        =   30
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "No of Chqs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   6840
         TabIndex        =   28
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ending No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3600
         TabIndex        =   27
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Starting No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Select Action:"
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
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FmeLatestActivity 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Current Position"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   14055
      Begin VB.TextBox TxtAction 
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
         Height          =   285
         Left            =   11280
         TabIndex        =   18
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox TxtStartingNo 
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
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox TxtEndingNo 
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
         Height          =   285
         Left            =   4920
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox TxtNoofChqs 
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
         Height          =   285
         Left            =   8520
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxtRemaining 
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
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtDone 
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
         Height          =   285
         Left            =   8760
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtDateDone 
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
         Height          =   285
         Left            =   12000
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Action taken:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   9840
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Starting No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ending No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "No of Chqs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6840
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "No. of Chqs Remaining:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Done By:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   7800
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Date Done:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   10560
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame FmeSelect 
      BackColor       =   &H00C0E0FF&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   14175
      Begin SSDataWidgets_B.SSDBCombo SSDBAccount 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   5535
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
         Columns(0).Caption=   "Account Number"
         Columns(0).Name =   "Account Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "Currency"
         Columns(1).Name =   "Currency"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   9763
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
      Begin SSDataWidgets_B.SSDBCombo SSDBLocations 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   720
         Width           =   5535
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
         Columns(0).Width=   6429
         Columns(0).Caption=   "Location"
         Columns(0).Name =   "Account Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Currency"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   9763
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
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select Location:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select Account:"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBAvailable 
      Height          =   2055
      Left            =   120
      TabIndex        =   34
      Top             =   3120
      Visible         =   0   'False
      Width           =   12255
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   5
      BackColorOdd    =   16761087
      RowHeight       =   503
      ExtraHeight     =   212
      Columns.Count   =   5
      Columns(0).Width=   3625
      Columns(0).Caption=   "Trans Date"
      Columns(0).Name =   "Trans Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   7
      Columns(0).NumberFormat=   "dd-mmm-yyyy"
      Columns(0).FieldLen=   256
      Columns(1).Width=   6033
      Columns(1).Caption=   "Starting Chq No"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   5133
      Columns(2).Caption=   "Ending Chq No"
      Columns(2).Name =   "Chq No"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   5636
      Columns(3).Caption=   "No of Chqs in Range"
      Columns(3).Name =   "Narration"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "CMID"
      Columns(4).Name =   "DBCR"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   256
      _ExtentX        =   21616
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Available Cheque number ranges"
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
Attribute VB_Name = "FrmCIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpCon As ADODB.Connection
Dim iStartNo As Long
Dim iEndNos As String
Dim iNoofRanges As Long

Private Sub ChkPrintery_Click()
Dim i As Integer
Dim bm As Variant

If ChkPrintery.Value = 1 Then
   SSDBMove.Enabled = False
   SSDBTransferOpt.Enabled = False
   With SSDBMove
     .MoveFirst
     For i = 0 To .Rows - 1
         bm = .GetBookmark(i)
         If .Columns(1).CellText(bm) = "T" Then
            .Bookmark = .GetBookmark(i)
             SSDBMove = .Columns(0).CellText(bm)
         Exit For
         End If
     Next i
   End With
   With SSDBTransferOpt
     .MoveFirst
     For i = 0 To .Rows - 1
         bm = .GetBookmark(i)
         If .Columns(1).CellText(bm) = "T" Then
            .Bookmark = .GetBookmark(i)
             SSDBTransferOpt = .Columns(0).CellText(bm)
         Exit For
         End If
     Next i
   End With
Else
   SSDBMove.Enabled = True
   SSDBMove = ""
   SSDBTransferOpt.Enabled = True
   SSDBTransferOpt = ""
End If

End Sub

Private Sub CmdExit_Click()
SpCon.Close
Unload Me
End Sub

Private Sub CmdSave_Click()
On Error GoTo Err_CmdSave_Click
Dim i As Integer
Dim StrSql As String
Dim ToLocation As String
Dim FromLocation As String

If IsValid Then
   ToLocation = SSDBMove.Columns(1).Text
   If ChkPrintery.Value = 1 Then
      FromLocation = "P"
   Else
      FromLocation = SSDBLocations.Columns(1).Text
   End If
   i = RunSP(SpCon, "usp_PendingTransfers", 0, SSDBAccount.Columns(0).Text, FromLocation, SSDBTransferOpt.Columns(1).Text, _
       CLng(TxtChqStart), CLng(TxtChqEnd), CLng(TxtChqRange), Format(DTPTransferDate, "dd-mmm-yyyy"), iNoofRanges, iEndNos, iStartNo, ToLocation, gblLoginName)
   If i = 0 Then
      MsgBox "Record saved"
   Else
      If i = 2 Then
         StrSql = "That number is in a range awaiting approval from a supervisor" & vbCrLf
         StrSql = StrSql & "Save is being abondoned"
         MsgBox StrSql, vbOKOnly, "Duplicated number awaiting approval"
      Else
         MsgBox "Error on saving" & " " & i
      End If
   End If
End If

Exit_CmdSave_Click:
Exit Sub

Err_CmdSave_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on saving"
Resume Exit_CmdSave_Click

End Sub

Private Sub Form_Activate()
' ready message
 frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
 Screen.MousePointer = vbDefault
 frmMDI.txtStatusMsg.Refresh
 '--
End Sub

Private Sub Form_Load()
Dim indx As Integer
Dim strTmp As String
On Error GoTo FL_ERR
'--
'-------------------------------------
'-- Initialize License Details -------
'-------------------------------------
 '--
csvCenterForm Me, gblMDIFORM
'-----------------------------------
Set SpCon = New ADODB.Connection
With SpCon
     .ConnectionString = gblFileName
     .CursorLocation = adUseServer
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

'--
   
FL_Exit:
  Exit Sub
FL_ERR:
  MsgBox Err & " " & Err.Description, , "Error on loading Cheque Inventory Screen"
  OpenErr = True
  On Error Resume Next
  Resume FL_Exit
   
End Sub
Private Sub SSDBAccount_Click()
If SSDBLocations = "" Then
   SSDBLocations.SetFocus
Else
   FillCurrentPosition
End If

End Sub

Private Sub SSDBAccount_InitColumnProps()
On Error GoTo Err_SSDBAccount_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_ListActiveAccounts", 1)
If adoRst.EOF Then
   MsgBox "Accounts were not setup" & vbCrLf & "Please do so now", vbCritical, "Account Error"
   GoTo Exit_SSDBAccount_InitColumnProps
End If

'adoRst.MoveFirst
With SSDBAccount
     .RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst(0) & vbTab & adoRst(1) & vbTab
     .AddItem StrSql
     adoRst.MoveNext
     StrSql = ""
     Loop
End With
adoRst.Close
Set adoRst = Nothing
Exit_SSDBAccount_InitColumnProps:
Exit Sub

Err_SSDBAccount_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing active accounts"
Resume Exit_SSDBAccount_InitColumnProps

End Sub

Private Sub SSDBLocations_Click()
On Error GoTo Err_SSDBLocations_Click
Dim StrSql As String
Dim adoRst As ADODB.Recordset

If SSDBAccount = "" Then
   SSDBAccount.SetFocus
Else
   FillCurrentPosition
End If

If SSDBLocations.Columns(1).Text = "T" Then
   ChkPrintery.Visible = True
Else
   ChkPrintery.Visible = False
   ChkPrintery.Value = 0
End If

Set adoRst = RunSP(SpCon, "usp_ListChequeTransferOptions", 1)
If adoRst.EOF Then
   MsgBox "Cheque transfer options were not setup" & vbCrLf & "Please do so now", vbCritical, "Transfer Options Error"
   GoTo Exit_SSDBLocations_Click
End If

SSDBTransferOpt = ""
With SSDBTransferOpt
     .RemoveAll
     Do While Not adoRst.EOF
     If adoRst(1) = "I" Then
        If SSDBLocations = "Treasury" Then
           GoTo AddItem
        Else
           GoTo LoopIT
        End If
     End If
AddItem:
     StrSql = adoRst(0) & vbTab & adoRst(1) & vbTab
     .AddItem StrSql
LoopIT:
     adoRst.MoveNext
     StrSql = ""
     Loop
End With
adoRst.Close
Set adoRst = Nothing
Exit_SSDBLocations_Click:
Exit Sub

Err_SSDBLocations_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing Cheque Transfer Options"
Resume Exit_SSDBLocations_Click

End Sub

Private Sub SSDBLocations_InitColumnProps()
On Error GoTo Err_SSDBLocations_InitColumnProps
Dim StrSql As String
Dim adoRst As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_ListLocations", 1)
If adoRst.EOF Then
   MsgBox "Cheque Locations were not setup" & vbCrLf & "Please do so now", vbCritical, "Locations Error"
   GoTo Exit_SSDBLocations_InitColumnProps
End If

With SSDBLocations
     .RemoveAll
     SSDBMove.RemoveAll
     Do While Not adoRst.EOF
     StrSql = adoRst(0) & vbTab & adoRst(1) & vbTab
     .AddItem StrSql
     SSDBMove.AddItem StrSql
     adoRst.MoveNext
     StrSql = ""
     Loop
End With
adoRst.Close
Set adoRst = Nothing
Exit_SSDBLocations_InitColumnProps:
Exit Sub

Err_SSDBLocations_InitColumnProps:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing active locations"
Resume Exit_SSDBLocations_InitColumnProps
End Sub
Sub FillCurrentPosition()
Dim adoRst As ADODB.Recordset
Dim adoRst1 As ADODB.Recordset

Set adoRst = RunSP(SpCon, "usp_LocationCurrentDetails", 1, SSDBAccount.Columns(0).Text, SSDBLocations.Columns(1).Text)

With adoRst
     If Not .EOF Then
        TxtStartingNo = !StartChqNo
        TxtEndingNo = !EndChqNo
        TxtNoofChqs = !NoChqs
        TxtAction = !ActionType
        TxtDone = !UserID
        TxtDateDone = Format(!PostDate, "dd-mmm-yyyy")
    End If
End With

Set adoRst1 = adoRst.NextRecordset
If Not adoRst1.EOF Then
   If IsNull(adoRst1!TotChqs) Then
      TxtRemaining = 0
   Else
       TxtRemaining = adoRst1!TotChqs
   End If
End If

'adoRst.Close
adoRst1.Close
Set adoRst = Nothing
Set adoRst1 = Nothing
DTPTransferDate = Date
FmeMovement.Enabled = True

End Sub

Private Sub SSDBMove_Click()
If SSDBMove.Columns(1).Text = SSDBLocations.Columns(1).Text Then
   MsgBox "Current location and destination are the same. Please correct", vbOKOnly, "Error in selection of Move To"
End If

End Sub

Private Sub SSDBTransferOpt_Click()
Dim adoRst As ADODB.Recordset
Dim StrSql As String
Dim i As Integer
Dim bm As Variant

SSDBAvailable.Visible = False

Set adoRst = RunSP(SpCon, "usp_ListAvailableRanges", 1, SSDBAccount.Columns(0).Text, SSDBLocations.Columns(1).Text)
If adoRst.EOF Then
   MsgBox "There are no ranges currently available" & vbCrLf & "Try another location", vbCritical, "Cheque Number Ranges"
   GoTo Exit_SSDBTransferOpt_Click
End If
With SSDBAvailable
     .RemoveAll
     Do While Not adoRst.EOF
        StrSql = adoRst(0) & vbTab & adoRst(1) & vbTab & adoRst(2) & vbTab & adoRst(3) & vbTab & adoRst(4)
        .AddItem StrSql
        adoRst.MoveNext
        StrSql = ""
     Loop
End With
adoRst.Close
Set adoRst = Nothing
SSDBAvailable.Visible = True

If SSDBTransferOpt.Columns(1).Text <> "T" Then
   With SSDBMove
     .MoveFirst
     For i = 0 To .Rows - 1
         bm = .GetBookmark(i)
         If .Columns(1).CellText(bm) = SSDBLocations.Columns(1).Text Then
            .Bookmark = .GetBookmark(i)
             SSDBMove = .Columns(0).CellText(bm)
         Exit For
         End If
     Next i
   End With
   SSDBMove.Enabled = False
Else
   SSDBMove.Enabled = True
   SSDBMove = ""
End If

Exit_SSDBTransferOpt_Click:
Exit Sub

Err_SSDBTransferOpt_Click:
MsgBox Err & " " & Err.Description, vbOKOnly, "Error on listing available cheque number ranges"
Resume Exit_SSDBTransferOpt_Click

End Sub
Function IsValid() As Boolean
Dim StrSql As String
Dim i As Integer
Dim CorrectStart As Boolean
Dim CorrectEnd As Boolean
Dim FromPrintery As Boolean
IsValid = False

If Not IsNumeric(TxtChqStart) Then
   StrSql = "The value entered in the Starting cheque Number field is not a number" & vbCrLf
   StrSql = StrSql & "Please correct and resubmit"
   MsgBox StrSql, vbOKOnly, "Validation of Entered Data"
   TxtChqStart.SetFocus
   GoTo Exit_IsValid
End If

If Not IsNumeric(TxtChqEnd) Then
   StrSql = "The value entered in the ending cheque Number field is not a number" & vbCrLf
   StrSql = StrSql & "Please correct and resubmit"
   MsgBox StrSql, vbOKOnly, "Validation of Entered Data"
   TxtChqEnd.SetFocus
   GoTo Exit_IsValid
End If

If Not IsNumeric(TxtChqRange) Then
   StrSql = "The value entered in the No. of Chqs field is not a number" & vbCrLf
   StrSql = StrSql & "Please correct and resubmit"
   MsgBox StrSql, vbOKOnly, "Validation of Entered Data"
   TxtChqRange.SetFocus
   GoTo Exit_IsValid
End If

If CLng(TxtChqStart) > CLng(TxtChqEnd) Then
   StrSql = "The Starting cheque Number is larger than the Ending Cheque Number" & vbCrLf
   StrSql = StrSql & "Please correct and resubmit"
   MsgBox StrSql, vbOKOnly, "Validation of Entered Data"
   TxtChqStart.SetFocus
   GoTo Exit_IsValid
End If

If CLng(TxtChqRange) <> (CLng(TxtChqEnd) - CLng(TxtChqStart) + 1) Then
   StrSql = "The value enetered for the No. of Chqs does not match the difference between the Start and End numbers" & vbCrLf
   StrSql = StrSql & "Please correct and resubmit"
   MsgBox StrSql, vbOKOnly, "Validation of Entered Data"
   TxtChqRange.SetFocus
   GoTo Exit_IsValid
End If
If SSDBMove = "" Then
   MsgBox "Please select a MOVE TO location first"
   SSDBMove.SetFocus
   GoTo Exit_IsValid
End If

iStartNo = 0
iNoofRanges = 0
iEndNos = ""

CorrectStart = False
CorrectEnd = False
FromPrintery = True


With SSDBAvailable
     If .Rows = 0 Then
        If ChkPrintery.Value = 1 Then
           GoTo CheckStatus
        Else
           MsgBox "No ranges are available to facilitate this transfer"
           GoTo Exit_IsValid
        End If
     End If
     .MoveFirst
     For i = 1 To .Rows
        If .Columns(1).Text = TxtChqStart Then
           If ChkPrintery.Value = 1 Then
              FromPrintery = False
              GoTo CheckStatus
           Else
               iStartNo = CLng(.Columns(4).Text)
               CorrectStart = True
           End If
        End If
        If CLng(TxtChqEnd) <= CLng(.Columns(2).Text) Then
           If ChkPrintery.Value = 1 Then
              FromPrintery = False
              GoTo CheckStatus
           Else
              iNoofRanges = iNoofRanges + 1
              iEndNos = iEndNos & .Columns(4).Text & ";"
              CorrectEnd = True
              GoTo CheckStatus
           End If
        Else
            iNoofRanges = iNoofRanges + 1
            iEndNos = iEndNos & .Columns(4).Text & ";"
        End If
       .MoveNext
     Next i
End With

CheckStatus:
If ChkPrintery.Value = 1 Then
   If FromPrintery = False Then
      MsgBox "The number entered from printery is currently active", vbOKOnly
      IsValid = False
      GoTo Exit_IsValid
   End If
   IsValid = True
   iEndNos = TxtChqEnd & ";"
   iStartNo = TxtChqStart
   iNoofRanges = 1
   GoTo Exit_IsValid
End If
If CorrectStart And CorrectEnd Then
   IsValid = True
Else
    If Not CorrectStart Then
       MsgBox "Number entered as the start number must match the beginning of one of the ranges", vbOKOnly
    Else
       MsgBox "The number entered as the end number cannot be greater than the ending of the ranges", vbOKOnly
    End If
End If

Exit_IsValid:
Exit Function

End Function
