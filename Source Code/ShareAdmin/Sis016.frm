VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form frmSIS016 
   Caption         =   "Payment Summary Viewer"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControl=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertControl=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
   End
End
Attribute VB_Name = "frmSIS016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crSIS016
Dim rs As New ADOR.Recordset
Private Sub Form_Load()
Dim sql As String
sql = ""
sql = "Select CompName, ParValue, "
sql = sql & "RecDate, ChqDate, PayAmt, IncTyp, "
sql = sql & "a.DecDate, a.PayTyp, a.ClientId, a.GrossPymnt, a.WhldTax, "
sql = sql & "a.Shares, a.ChqNum, a.ResCode, a.CatCode, "
sql = sql & "b.CliName, CatDesc, ResCtry"
sql = sql & " from ((((Company inner join DIVREF"
sql = sql & " on company.compname <> divref.Paytyp)"
sql = sql & " inner join STKPYMNTS a on divref.decdate = a.decdate "
sql = sql & " and divref.paytyp = a.paytyp)"
sql = sql & " INNER JOIN StkTaxr ON"
sql = sql & " a.ResCode = StkTaxr.ResCode)"
sql = sql & " inner join STKCAT on"
sql = sql & " a.catcode = stkcat.catcode)"
sql = sql & " inner join stkname b on"
sql = sql & " a.clientid = b.clientid"
sql = sql & " order by stktaxr.resctry, a.catcode"
'--
Report.Database.Tables.Item(1).SetLogOnInfo gblDSN

rs.Open sql, gblFileName, adOpenDynamic, adLockReadOnly
Report.Database.SetDataSource rs
crv.ReportSource = Report
Report.DiscardSavedData
crv.ViewReport

End Sub

Private Sub Form_Resize()
crv.top = 0
crv.left = 0
crv.Height = ScaleHeight
crv.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
Set rs = Nothing
Set Report = Nothing
Set frmSIS016 = Nothing
End Sub


