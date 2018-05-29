VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form frmSIS045V 
   Caption         =   "Profile Viewer"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
Attribute VB_Name = "frmSIS045V"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crSIS045
Dim rs As New ADOR.Recordset

Private Sub Form_Load()
Dim sql As String
sql = ""
sql = sql & " Select a.CERTNO, a.CLIENTID, a.ISSDATE,"
sql = sql & " a.SHARES, a.STATUS, a.TRNCODE, a.TRNBATCH,"
sql = sql & " a.TRNDATE, a.FORMNO,"
sql = sql & " b.CLINAME, b.CLIADDR1, "
sql = sql & " b.CLIADDR2, b.CLIADDR3, b.CLIADDR4, b.CLIADDR5,"
sql = sql & " b.DTEOPENED, c.JNTNAME1, c.JNTNAME2, c.JNTNAME3, c.JNTENDDTE "
sql = sql & " from (CERTMST a inner join STKNAME b on a.CLIENTID = b.CLIENTID)"
sql = sql & " left join STKJOINT c on b.CLIENTID = c.CLIENTID"
sql = sql & " where a.ClientID = " & gblFileKey
sql = sql & " and c.JNTENDDTE is NULL"
rs.Open sql, gblFileName, adOpenDynamic, adLockReadOnly
Report.Database.Tables.Item(1).SetLogOnInfo gblDSN
Report.Database.SetDataSource rs
crv.ReportSource = Report
Report.DiscardSavedData
crv.ViewReport
frmMDI.txtStatusMsg.SimpleText = gblReadyMsg
Screen.MousePointer = vbDefault
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
Set frmSIS045V = Nothing
End Sub
