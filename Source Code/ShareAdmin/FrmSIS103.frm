VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmSIS103 
   Caption         =   "Dividend History "
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "FrmSIS103.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CRV 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   14175
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmSIS103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRDivHist
Dim rs As New ADODB.Recordset
Dim SpCon As ADODB.Connection

Private Sub Form_Load()
Dim qSQL As String

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
   frmMDI.txtStatusMsg.Refresh
Loop
Screen.MousePointer = vbDefault
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg

Set rs = RunSP(SpCon, "usp_DivHistory", 1, gblFileKey)
Report.PrinterSetup Me.hwnd
Report.ParameterFields.Item(1).AddCurrentValue gblCompName
Report.Database.SetDataSource rs
CRV.ReportSource = Report
CRV.ViewReport


End Sub

Private Sub Form_Resize()
CRV.Top = 0
CRV.Left = 0
CRV.Height = ScaleHeight
CRV.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
SpCon.Close
Set Report = Nothing
Set FrmSIS103 = Nothing
End Sub

