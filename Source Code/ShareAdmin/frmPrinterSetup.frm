VERSION 5.00
Begin VB.Form frmPrinterSetup 
   Caption         =   "Printer Setup"
   ClientHeight    =   5520
   ClientLeft      =   360
   ClientTop       =   480
   ClientWidth     =   6375
   Icon            =   "frmPrinterSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbPaperSize 
      Height          =   315
      Left            =   3480
      TabIndex        =   21
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txtPortName 
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtDriverName 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtReportPrinterName 
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   240
      Width           =   3015
   End
   Begin VB.ListBox lstDriverName 
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstPort 
      Height          =   255
      Left            =   5400
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstDeviceName 
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame frmOrientation 
      Caption         =   "Orientation:"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1815
      Begin VB.OptionButton optDefault 
         Caption         =   "Default"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "P&ortrait"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optLandscape 
         Caption         =   "L&andscape"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame lblPrinter 
      Caption         =   "Printer:"
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   6015
      Begin VB.ComboBox cmbPrinters 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Text            =   "Printers Available"
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "These are the Printers configured for this machine:"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label lblName 
         Caption         =   "&Name"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblPrinterDriver 
         Caption         =   "Printer Driver:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label lblPrinterPort 
         Caption         =   "Printer Port:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   5175
      End
   End
   Begin VB.Label lblPaperSize 
      Caption         =   "Paper Size:"
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblPortName 
      Caption         =   "Port Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblDriverName 
      Caption         =   "Printer Driver:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblReportPrinterName 
      Caption         =   "Report Printer Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmPrinterSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Information required for setting a printer
Dim pname As String
Dim pport As String
Dim pdriver As String
Private Sub cmbPrinters_Click()
    
    lblPrinterDriver.Caption = "Printer Driver:         " & lstPort.List(cmbPrinters.ListIndex)
    lblPrinterPort.Caption = "Printer Port:           " & lstDriverName.List(cmbPrinters.ListIndex)

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me

End Sub

Private Sub cmdOk_Click()

'Get the printer info from the comboboxes
pname = cmbPrinters.List(cmbPrinters.ListIndex)
pport = lstPort.List(cmbPrinters.ListIndex)
pdriver = lstDriverName.List(cmbPrinters.ListIndex)
'Set the Printer to print to
Call cr.SelectPrinter(pdriver, pname, pport)

'Set the paper orientation
If optLandscape = True Then
    cr.PaperOrientation = crLandscape
ElseIf optPortrait = True Then
    cr.PaperOrientation = crPortrait
Else
    'cr.PaperOrientation = crDefaultPaperOrientation - ERROR
End If

Select Case cmbPaperSize.ListIndex

    Case 0
        cr.PaperSize = crPaperLetter
    Case 1
        cr.PaperSize = crPaperLegal
    Case 2
        cr.PaperSize = crPaper11x17
    Case 4
       cr.PaperSize = crDefaultPaperSize
    
    'For Case -1(Nothing Selected), 3 (crDefaultPaperSize)
    'Note: you can not Set to the Printer to default
    End Select
' Don't prompt user; info already provided by user
'cr.PrintOut False  'Sends the Report to the Printer
 Unload Me
End Sub

Private Sub Form_Load()


' Get back Printer Info saved with the Report
txtReportPrinterName = cr.PrinterName
txtDriverName = cr.DriverName
txtPortName = cr.PortName

Select Case cr.PaperOrientation
Case crLandscape
    optLandscape.Value = True
Case crPortrait
    optPortrait.Value = True
Case crDefaultPaperOrientation
    optDefault.Value = True
End Select
'--
'Populate a Combo box with the appropriate Paper sizes
'To get a complete list of constants look in the Object Browser under crPaperSize
    cmbPaperSize.AddItem "Letter, 8 1/2 x 11" '1
    cmbPaperSize.AddItem "Legal"              '5
    cmbPaperSize.AddItem "11 x 17"            '17
    cmbPaperSize.AddItem "14 x 11 (dot Matrix Cheques)" ' 39
    cmbPaperSize.AddItem "Letter Transverse - Certs"
    cmbPaperSize.AddItem "Default"
    
'Display the paper size of the report
    Select Case cr.PaperSize

    Case 1
        cmbPaperSize.Text = cmbPaperSize.List(0)
    Case 5
        cmbPaperSize.Text = cmbPaperSize.List(1)
    Case 17
        cmbPaperSize.Text = cmbPaperSize.List(2)
    Case 39
        cmbPaperSize.Text = cmbPaperSize.List(3)
    Case 258
        cmbPaperSize.Text = cmbPaperSize.List(4)
    Case Else
        cmbPaperSize.Text = cmbPaperSize.List(5)
    End Select


'Using VB's Printer object; NOT Crystal
Dim mPrinter As Printer
For Each mPrinter In Printers
    cmbPrinters.AddItem mPrinter.DeviceName
    lstDeviceName.AddItem mPrinter.DeviceName
    lstDriverName.AddItem mPrinter.Port
    lstPort.AddItem mPrinter.DriverName
Next


End Sub


