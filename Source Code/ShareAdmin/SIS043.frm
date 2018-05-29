VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form frmSIS043 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
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
Attribute VB_Name = "frmSIS043"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crSIS043
Dim rsTopN As New ADODB.Recordset
Dim cmdDel As New ADODB.Command

Private Sub Form_Load()
 Dim sMsg, sTitle, sDefault, sValue As String
 Dim qSQL As String, iTopN As Integer
 Dim rsName As New ADODB.Recordset
 
 '--
 Set rsName = New ADODB.Recordset
 Set rsTopN = New ADODB.Recordset
 Set cnn = New ADODB.Connection
 cnn.Open gblFileName
 '--
 sMsg = "Enter a number representing the 'N'" _
 & " largest shareholders required?"
 sTitle = "Top Shareholders Input Box"
 sDefault = "10"
 sValue = InputBox(sMsg, sTitle, sDefault)
 Me.Caption = "Top " & sValue & " Stockholders Report"
 If Not IsNothing(sValue) Then
   qSQL = "DELETE FROM STKTOPN"
   If csvADODML(qSQL) = True Then
     qSQL = "SELECT CLIENTID, CLINAME, SHARES, JOINT " _
          & " FROM STKNAME WHERE SHARES > 0 " _
          & " ORDER BY SHARES DESC, CLINAME ASC "
     rsName.Open qSQL, gblFileName
     
     With rsName
         If Not .EOF Then
           cnn.BeginTrans
           rsTopN.Open "STKTOPN", gblFileName, adOpenDynamic, adLockOptimistic, adCmdTable
           .MoveFirst
           iTopN = 1
           Do While Not .EOF
             If iTopN > Val(sValue) Then Exit Do
             rsTopN.AddNew
             rsTopN!clientid = !clientid
             rsTopN!cliname = !cliname
             rsTopN!shares = !shares
             rsTopN!JOINT = !JOINT
             rsTopN.Update
             iTopN = iTopN + 1
             .MoveNext
           Loop
           cnn.CommitTrans
         End If
         .Close
     End With
     rsTopN.Close
     cnn.Close
  End If
 Report.Database.Tables.Item(1).SetLogOnInfo gblDSN
 CRViewer1.ReportSource = Report
 Report.DiscardSavedData
 CRViewer1.ViewReport
End If
End Sub

Private Sub Form_Resize()
CRViewer1.top = 0
CRViewer1.left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsName = Nothing
Set rsTopN = Nothing
Set Report = Nothing
Set frmSIS043 = Nothing
Set cnn = Nothing
End Sub
