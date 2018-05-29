VERSION 5.00
Begin VB.Form frmPrintingStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Status"
   ClientHeight    =   1920
   ClientLeft      =   10560
   ClientTop       =   8985
   ClientWidth     =   4545
   Icon            =   "frmPrintingStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPSProgressAfter 
      Height          =   285
      Left            =   3480
      TabIndex        =   13
      Top             =   1560
      Width           =   900
   End
   Begin VB.TextBox txtPSProgressBefore 
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Top             =   1560
      Width           =   900
   End
   Begin VB.TextBox txtPSSelectedAfter 
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Top             =   1200
      Width           =   900
   End
   Begin VB.TextBox txtPSSelectedBefore 
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   1200
      Width           =   900
   End
   Begin VB.TextBox txtPSReadAfter 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   840
      Width           =   900
   End
   Begin VB.TextBox txtPSReadBefore 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   840
      Width           =   900
   End
   Begin VB.TextBox txtPSPrintedAfter 
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Top             =   480
      Width           =   900
   End
   Begin VB.TextBox txtPSPrintedBefore 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   900
   End
   Begin VB.Label lblPSProgress 
      Caption         =   "Progress:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblPSNumOfRecordsSelected 
      Caption         =   "Number of Records Selected:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblPSNumOfRecordsRead 
      Caption         =   "Number of Records Read:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblPSNumOfRecordsPrinted 
      Caption         =   "Number of Records Printed:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblPSAfter 
      Caption         =   "After"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblPSBefore 
      Caption         =   "Before"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmPrintingStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
