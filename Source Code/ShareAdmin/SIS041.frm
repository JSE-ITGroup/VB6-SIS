VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form frmSIS041 
   Caption         =   "Brokers Register Print Viewer"
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
Attribute VB_Name = "frmSIS041"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crSIS041
Dim rsMain As New ADODB.Recordset
Dim rsBrkReg As New ADODB.Recordset

Private Sub Form_Load()
Dim qSQL  As String
On Error GoTo FormLoad_Err
qSQL = "SELECT * from STKACTIV where STATUS = 'O' " _
       & "and (TRNCODE = 'B' or TRNCODE = 'C' or " _
       & "TRNCODE = 'T') order by TRNBATCH,FORM,LINENO;"
'--
Set cnn = New ADODB.Connection
cnn.Open gblFileName
'--
Set rsMain = New ADODB.Recordset
Set rsBrkReg = New ADODB.Recordset
rsMain.Open qSQL, gblFileName
If Not rsMain.EOF Then
  rsBrkReg.Open "STKBRKTRN", gblFileName, adOpenDynamic, adLockOptimistic, adCmdTable
  CreateRegister
  Report.Database.Tables.Item(1).SetLogOnInfo gblDSN
  CRViewer1.ReportSource = Report
  CRViewer1.ViewReport
Else
  MsgBox "No Register to Print"
  Exit Sub
End If
FormLoad_Exit:
 Exit Sub
FormLoad_Err:
  csvShowError "SIS041/Load"
 ' csvLogError "SIS041/Load", Err.Number, Err.Description
GoTo FormLoad_Exit
End Sub

Private Sub Form_Resize()
CRViewer1.top = 0
CRViewer1.left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub CreateRegister()
Dim iCert As Long, iCommit As Integer
If iCommit = 0 Or iCommit = 500 Then
     cnn.BeginTrans
     iCommit = 0
End If
With rsBrkReg ' Delete Previous Register Items
  .Requery
  If Not .EOF Then
    .MoveFirst
    Do While Not .EOF
      .Delete
      iCommit = iCommit + 1
      If iCommit = 500 Then
        cnn.CommitTrans
        iCommit = 0
      End If
      .MoveNext
    Loop
  End If
End With
'If iCommit > 0 Then
 ' cnn.CommitTrans
 ' iCommit = 0
'End If
'--
With rsMain
    .MoveFirst
    Do While Not .EOF
      
      If !TRNCODE = "B" Then 'Broker to Broker
         If !FRCERT > 0 Then  'Store Sell Record Cert
            iCert = !FRCERT
         Else  ' generate Sell & buy record
            rsBrkReg.AddNew         ' Sell record
            rsBrkReg!broker = !BROKERID
            rsBrkReg!certno = !certno
            rsBrkReg!Form = !Form
            rsBrkReg!TFRDAT = !TRNDATE
            rsBrkReg!SELLSHARES = !shares
            rsBrkReg!batch = !TRNBATCH
            rsBrkReg!TRNCODE = !TRNCODE
            rsBrkReg!CLIENTID = !CLIENTID
            rsBrkReg.Update
            iCommit = iCommit + 1
            rsBrkReg.AddNew    'Buy Record
            rsBrkReg!broker = !CLIENTID
            rsBrkReg!CLIENTID = !BROKERID
            rsBrkReg!certno = iCert
            rsBrkReg!Form = !Form
            rsBrkReg!TFRDAT = !TRNDATE
            rsBrkReg!BUYSHARES = !shares
            rsBrkReg!batch = !TRNBATCH
            rsBrkReg!TRNCODE = !TRNCODE
            rsBrkReg.Update
            iCommit = iCommit + 1
         End If
      Else
         If !TRNCODE = "C" Then ' Broker Buys from Shareholder
            If !FRCERT > 0 Then  'Shareholder SellS
               rsBrkReg.AddNew
               rsBrkReg!broker = !BROKERID
               rsBrkReg!certno = !FRCERT
               rsBrkReg!Form = !Form
               rsBrkReg!TFRDAT = !TRNDATE
               rsBrkReg!BUYSHARES = !FRSHARES
               rsBrkReg!batch = !TRNBATCH
               rsBrkReg!TRNCODE = !TRNCODE
               rsBrkReg!CLIENTID = !CLIENTID
               rsBrkReg.Update
               iCommit = iCommit + 1
             End If
         Else
           If !TRNCODE = "T" Then 'Broker Sells to Shareholder
              If !certno > 0 Then
              rsBrkReg.AddNew
               rsBrkReg!broker = !BROKERID
               rsBrkReg!certno = !certno
               rsBrkReg!Form = !Form
               rsBrkReg!TFRDAT = !TRNDATE
               rsBrkReg!SELLSHARES = !shares
               rsBrkReg!batch = !TRNBATCH
               rsBrkReg!TRNCODE = !TRNCODE
               rsBrkReg!CLIENTID = !CLIENTID
               rsBrkReg.Update
               iCommit = iCommit + 1
             End If
           End If
         End If
      End If
      If iCommit >= 500 Then
        cnn.CommitTrans
        iCommit = 0
      End If
    .MoveNext
    Loop
If iCommit > 0 Then cnn.CommitTrans
.Close
rsBrkReg.Close
cnn.Close
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
Set rsBrkReg = Nothing
Set cnn = Nothing
Set frmSIS041 = Nothing
End Sub
