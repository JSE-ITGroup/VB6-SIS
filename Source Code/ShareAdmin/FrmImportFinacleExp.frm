VERSION 5.00
Begin VB.Form FrmImportFinacleExp 
   Caption         =   "Import Finacle Exception Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FrmImportFinacleExp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmImportFinacleExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function ExtractData()
Dim fs, F, iRecs As Long
Dim textfile As String
Dim sInRec As String
Dim sRec As String
Dim X As Integer
Dim StrSql As String
Dim sMsg As String
Dim iAccountNo As String
Dim iAmount As Currency
Dim iDate As Date

sMsg = "You are about to import a Finacle Exception Report "
sMsg = sMsg & "Select No if you are unsure or accidentally selected this option."
sMsg = sMsg & " Do you want to continue?"
sTitle = "Converting iLink file"
X = MsgBox(sMsg, vbExclamation + vbYesNo, sTitle)
If X = vbNo Then
  Exit Function
End If

txtfile = frmMDI.cmnDialog.filename
Screen.MousePointer = vbHourglass
frmMDI.txtStatusMsg.SimpleText = "Importing Finacle Exception Report..."
frmMDI.txtStatusMsg.Refresh

Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.opentextfile(textfile)
sInRec = F.readline
iRecs = 0
If F.atendofstream = True Then
  MsgBox "Input Text File " & textfile & " is blank; import aborting... "
  F.Close
End If

Do Until iRecs = 5
   sInRec = F.readline
   iRecs = iRecs + 1
Loop
StrSql = Mid(sInRec, 1, 10)
iDate = CDate(StrSql)

StartProcessing:
iRecs = 0

Do Until F.atendofstream = True
   sRec = Trim(Mid(sInRec, 1, 9))
   If IsNumeric(sRec) = False Then
      GoTo ReadAnother
   End If
   iAccountNo = Mid(sInRec, 16, 9)
   iAmount = CCur(Trim(Mid(sInRec, 58, 16)))
   StrSql = Trim(Mid(sInRec, 77, 20))
   iRecs = iRecs + 1
   X = RunSP(SpCon, "usp_ImportFinacleException", 0, iDate, iAccountNo, iAmount, StrSql, gblLoginName)
ReadAnother:
   sInRec = F.readline
   frmMDI.txtStatusMsg.SimpleText = "Processing record " & iRecs
   frmMDI.txtStatusMsg.Refresh
Loop

F.Close

frmMDI.txtStatusMsg.SimpleText = "Importation completed"
frmMDI.txtStatusMsg.Refresh

End Function
