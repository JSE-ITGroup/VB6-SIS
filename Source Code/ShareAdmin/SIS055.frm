VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRViewer.dll"
Begin VB.Form frmSIS055 
   Caption         =   "Report Viewer"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   5430
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5805
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
Attribute VB_Name = "frmSIS055"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crSIS055
Dim rs As New ADOR.Recordset
Private Sub Form_Load()
Dim sql As String
sql = " "
sql = "SELECT " _
     & "CompName, " _
     & "a.TABLENO, a.FIELDID, a.RECORDNO, a.USERID, a.CHGDATE, a.NEWINFO, a.OLDINFO, " _
     & "FIELDDESC, " _
     & "UserName " _
     & "From " _
     & "((Company INNER JOIN AUDTRN a ON " _
     & "COMPANY.CompName <> a.USERID) " _
     & "INNER JOIN Users ON " _
     & "a.USERID = Users.SystemName) " _
     & "INNER  JOIN FLDREF ON " _
     & "a.TABLENO = FLDREF.TABLENO AND " _
     & "a.FIELDID = FLDREF.FIELDID " _
     & "WHERE a.NEWINFO <> a.OLDINFO " _
     & "Order By " _
     & "a.CHGDATE, a.USERID, a.TABLENO, " _
     & "a.RECORDNO, a.FIELDID "
     
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
Dim rsCmp As New ADODB.Recordset
Dim iRep As Integer
'--
Set Report = Nothing
rs.Close
Set rs = Nothing
'--
iRep = MsgBox("Did you print the report & was it okay?", _
        vbQuestion + vbYesNo, "Print Audit Trail")
If iRep = vbNo Then Exit Sub
'-- update company!audit report printed
Set rsCmp = New ADODB.Recordset
rsCmp.Open "Company", gblFileName, adOpenDynamic, adLockOptimistic, adCmdTable
If Not rsCmp.EOF Then
  rsCmp!auditind = True
  rsCmp.Update
End If
rsCmp.Close
Set rsCmp = Nothing
Form_Unload_Exit:
  Exit Sub
Form_unload_Err:
  csvShowError "SIS055/UnLoad"
  GoTo Form_Unload_Exit
End Sub
