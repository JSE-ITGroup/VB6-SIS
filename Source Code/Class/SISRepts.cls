VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SISRepts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Clientid As Long
Private m_siteid As String
Private m_loginid As String
Private m_reptyp As Integer
Private m_reportnum As Integer
Private m_OptNo  As Integer
Private m_gblcnn As Object
Private m_Date1 As Date
Private m_Date2 As Date
Private m_StringData As String
Public Property Let ClientId(vclientid As Long)
m_Clientid = vclientid
End Property
Public Property Let siteid(ByVal vsiteid As String)
m_siteid = vsiteid
End Property
Public Property Let ReportType(vreptyp As Integer)
m_reptyp = vreptyp
End Property
Public Property Let ReportNumber(vrepnum As Integer)
m_reportnum = vrepnum
End Property
Public Property Let OptNo(vOptNo As Integer)
m_OptNo = vOptNo
End Property
Public Property Let LoginId(vloginid As String)
m_loginid = vloginid
End Property
Public Property Get GblConnection() As Object
Set GblConnection = m_gblcnn
End Property
Public Property Set GblConnection(ByRef New_GblConnection As Object)
Set m_gblcnn = New_GblConnection
End Property
Public Property Let Date1(vDate1 As Date)
m_Date1 = vDate1
End Property
Public Property Let Date2(vDate2 As Date)
m_Date2 = vDate2
End Property
Public Property Let PassData(vPassData As String)
m_StringData = vPassData
End Property
Private Sub Class_Initialize()
 m_loginid = ""
 m_Clientid = 0
 m_OptNo = 0
 m_reptyp = 9
 m_reportnum = 0
 m_siteid = ""
 gblSiteId = ""
 gblFileName = ""
 gblOptions = 0
''' cnn = ""
End Sub
Public Sub RunShareHolderReport()
gblOptions = m_reportnum
gblFileName = m_loginid
gblSiteId = m_siteid
gblOptNo = m_OptNo
gblFileKey = m_Clientid
Set cnn = m_gblcnn
gblDate1 = m_Date1
gblDate2 = m_Date2
gblPassData = m_StringData
'--
Select Case m_reptyp
Case 1  ' Print shareholder certificates
 frmSIS056.Show 0
Case 2  ' Print Shareholder Cheques
Case 9  ' print Shareholder standard reports
  frmReportEngine.Show 0
End Select
End Sub


