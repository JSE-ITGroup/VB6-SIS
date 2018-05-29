VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1815
   ClientLeft      =   1635
   ClientTop       =   2385
   ClientWidth     =   4185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Find.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1815
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_A.SSDBOptSet optBtn 
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   1710
      _Version        =   196611
      _ExtentX        =   3016
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "&No"
      BackColor       =   -2147483643
      Cols            =   2
      IndexSelected   =   0
      NumberOfButtons =   2
      WidthGap        =   8
      ColumnWidth[0].UserSet=   -1  'True
      ColumnWidth[0].Value=   794
      Buttons.Button(0).OptionValue=   "0"
      Buttons.Button(0).Caption=   "&No"
      Buttons.Button(0).Mnemonic=   78
      Buttons.Button(0).Value=   -1  'True
      Buttons.Button(0).TextLeft=   15
      Buttons.Button(0).TextRight=   29
      Buttons.Button(0).TextBottom=   14
      Buttons.Button(0).ButtonRight=   13
      Buttons.Button(0).ButtonBottom=   13
      Buttons.Button(0).PictureLeft=   31
      Buttons.Button(0).PictureRight=   30
      Buttons.Button(0).PictureBottom=   14
      Buttons.Button(0).ButtonToColRight=   37
      Buttons.Button(0).ButtonToColBottom=   14
      Buttons.Button(0).ButtonBitmapID=   2
      Buttons.Button(1).OptionValue=   "1"
      Buttons.Button(1).Caption=   "&Yes"
      Buttons.Button(1).Mnemonic=   89
      Buttons.Button(1).ColOffset=   10
      Buttons.Button(1).TextLeft=   63
      Buttons.Button(1).TextRight=   81
      Buttons.Button(1).TextBottom=   14
      Buttons.Button(1).ButtonLeft=   48
      Buttons.Button(1).ButtonRight=   61
      Buttons.Button(1).ButtonBottom=   13
      Buttons.Button(1).PictureLeft=   83
      Buttons.Button(1).PictureRight=   82
      Buttons.Button(1).PictureBottom=   14
      Buttons.Button(1).ButtonToColLeft=   48
      Buttons.Button(1).ButtonToColRight=   113
      Buttons.Button(1).ButtonToColBottom=   14
      Buttons.Button(1).Column=   1
   End
   Begin VB.ComboBox cbOptions 
      Height          =   315
      ItemData        =   "Find.frx":000C
      Left            =   2640
      List            =   "Find.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1350
   End
   Begin VB.ComboBox cbWhere 
      Height          =   315
      ItemData        =   "Find.frx":004F
      Left            =   645
      List            =   "Find.frx":0056
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   1365
   End
   Begin VB.CommandButton cmdFind 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox tbFind 
      Height          =   285
      Left            =   630
      TabIndex        =   0
      Top             =   120
      Width           =   3345
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Search from Top?"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   2085
      TabIndex        =   8
      Top             =   540
      Width           =   540
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Where"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   450
      Width           =   480
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   15
      TabIndex        =   4
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
       Case 0
          Hide
       Case 1
          tbFind.Text = vbNullString
          Hide
    End Select
End Sub


