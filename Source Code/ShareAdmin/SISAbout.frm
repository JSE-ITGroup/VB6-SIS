VERSION 5.00
Begin VB.Form SISAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBSIS"
   ClientHeight    =   3825
   ClientLeft      =   2265
   ClientTop       =   2325
   ClientWidth     =   5475
   ControlBox      =   0   'False
   Icon            =   "SISAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   5475
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   328
      Left            =   4080
      TabIndex        =   0
      Top             =   3480
      Width           =   1365
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Filled in at run time"
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
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Filled in at run time"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Filled in at run time"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   5175
   End
End
Attribute VB_Name = "SISAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
Unload Me
Set SISAbout = Nothing
End Sub

Private Sub Form_Load()
Dim msg As String
msg = "Designed and Written By"
msg = msg & vbCrLf & "Stanford C Allen "
msg = msg & vbCrLf & "e-mail - stantheman_jm@yahoo.com"
msg = msg & vbCrLf & vbCrLf
msg = msg & vbCrLf
'msg = msg & "Beautified, enhanced and maintained by NCB - Systems Development."

lblMessage(0).Caption = msg
lblMessage(1).Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblMessage(2).Caption = "SHAREHOLDER INFORMATION SYSTEM"
End Sub


