VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRViewer.dll"
Begin VB.Form frmSIS017 
   Caption         =   "Payment Register Viewer"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7005
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
Attribute VB_Name = "frmSIS017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crSIS017
Dim rs As New ADOR.Recordset
Private Sub Form_Load()
sql = ""
sql = "Select CompName, ParValue, "
sql = sql & "RecDate, ChqDate, PayAmt, IncTyp, "
sql = sql & "a.DecDate, a.PayTyp, a.ClientId, a.GrossPymnt, a.WhldTax, "
sql = sql & "a.Shares, a.chqnum, a.ResCode, a.CatCode, "
sql = sql & "b.cliname, CatDesc, ResCtry, CliAddr1, CliAddr2"
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
sql = sql & " order by stktaxr.resctry, a.catcode, b.cliname"
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
Set frmSIS017 = Nothing
End Sub
