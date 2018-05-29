Attribute VB_Name = "ModExport"
Function ExportToExcel(adoRst As ADODB.Recordset)
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object


    Dim recArray As Variant

    Dim strDB As String
    Dim fldCount As Integer
    Dim recCount As Long
    Dim iCol As Integer
    Dim iRow As Integer

    ' Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets("Sheet1")

    ' Display Excel and give user control of Excel's lifetime
    xlApp.Visible = True
    xlApp.UserControl = True

    ' Copy field names to the first row of the worksheet
    fldCount = adoRst.Fields.Count
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = adoRst.Fields(iCol - 1).Name
      '  xlWs.Cells(1, iCol).Font.Bold = True
            '.Name = "Arial"
            '.Size = 9
    Next
    
    ' Check version of Excel
    If Val(Mid(xlApp.Version, 1, InStr(1, xlApp.Version, ".") - 1)) > 8 Then
        'EXCEL 2000 or 2002: Use CopyFromRecordset

        ' Copy the recordset to the worksheet, starting in cell A2
        xlWs.Cells(2, 1).CopyFromRecordset adoRst
        'Note: CopyFromRecordset will fail if the recordset
        'contains an OLE object field or array data such
        'as hierarchical recordsets

    Else
        'EXCEL 97 or earlier: Use GetRows then copy array to Excel

        ' Copy recordset to an array
        recArray = adoRst.GetRows
        'Note: GetRows returns a 0-based array where the first
        'dimension contains fields and the second dimension
        'contains records. We will transpose this array so that
        'the first dimension contains records, allowing the
        'data to appears properly when copied to Excel

        ' Determine number of records

        recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array


        ' Check the array for contents that are not valid when
        ' copying the array to an Excel worksheet
        For iCol = 0 To fldCount - 1
            For iRow = 0 To recCount - 1
                ' Take care of Date fields
                If IsDate(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                ' Take care of OLE object fields or array fields
                ElseIf IsArray(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = "Array Field"
                End If
            Next iRow 'next record
        Next iCol 'next field

        ' Transpose and Copy the array to the worksheet,
        ' starting in cell A2
        xlWs.Cells(2, 1).Resize(recCount, fldCount).Value = _
            TransposeDim(recArray)
    End If

    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit

    ' Close ADO objects
    'adoRst.Close
    'Set adoRst = Nothing
    
    ' Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing

    Set xlApp = Nothing

End Function

Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)

    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant

    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)

    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = v(Y, X)
        Next Y
    Next X

    TransposeDim = tempArray

End Function

Public Function SplitCliName(ByVal vName As String) As Boolean
Dim iLen As Integer
Dim ipos As Integer

SplitCliName = False
If IsNull(vName) Then
  Exit Function
Else
  iLen = Len(Trim(vName))
  ipos = InStr(1, vName, ",", vbTextCompare)
  If ipos = 0 Then ' no comma found
    Exit Function
  Else
    gblFileKey = Mid(vName, 1, ipos - 1)
    gblHold = Mid(vName, ipos + 1, iLen - ipos)
    SplitCliName = True
  End If
 End If
End Function
Public Function SendEmail(Conn As ADODB.Connection, strSender As String, _
                        ByVal strRecipient As String, _
                        ByVal strSubject As String, _
                        ByVal strBody As String, _
                        Optional ByVal strCc As String, _
                        Optional ByVal strBcc As String, _
                        Optional ByVal colAttachments As Collection _
                         ) As Boolean
Dim cdoMsg As New CDO.Message
Dim cdoConf As New CDO.Configuration
Dim Flds
Dim attachment
Dim strHTML
Dim adoRst As ADODB.Recordset
    
On Error GoTo ErrTrap
Const cdoSendUsingPort = 2
    
'Set cdoMsg =  CreateObject("CDO.Message")
'Set cdoConf = CreateObject("CDO.Configuration")
Set adoRst = RunSP(Conn, "usp_GetEmailSettings", 1)
If adoRst.State = adStateClosed Or adoRst.EOF Then
   MsgBox "Email settings were not found. Password cannot be email to you, sorry!"
End If

Set Flds = cdoConf.Fields
        
With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = adoRst!UseSSL
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = adoRst!SMTPServer
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = adoRst!ServerPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = adoRst!SMTPAuthenticate
        If Not IsNull(adoRst!SendUserName) Then
           .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = adoRst!SendUserName
        End If
        If Not IsNull(adoRst!SendPassword) Then
           .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = adoRst!SendPassword
        End If
    .Update
End With
    
adoRst.Close
Set adoRst = Nothing
' Apply the settings to the message.
With cdoMsg
    Set .Configuration = cdoConf
        .To = strRecipient
        .From = strSender
        .Subject = strSubject
        .TextBody = strBody
        If Not colAttachments Is Nothing Then
            For Each attachment In colAttachments
                .AddAttachment attachment
            Next
        End If
        If strCc <> "" Then .CC = strCc
        If strBcc <> "" Then .BCC = strBcc
        .Send
End With
    
Set cdoMsg = Nothing
Set cdoConf = Nothing
Set Flds = Nothing
        
SendEmail = True
Exit Function
ErrTrap:
Err.Raise Err.Number, "", "Error from Functions.SendEmail" & Err.Description
    SendEmail = False
End Function

Public Function ImportExcel(FilePathName As String, TableName As String)
Dim StrSql As String
Dim SpCon As New ADODB.Connection
Dim lngRecsAff As Long

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

StrSql = "IF EXISTS (SELECT * FROM dbo.sysobjects where id = object_id(N'[dbo].[" & TableName & "]') "
StrSql = StrSql & "and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
StrSql = StrSql & "DROP TABLE [dbo].[" & TableName & "]"
SpCon.Execute StrSql
StrSql = "SELECT * INTO " & TableName & " FROM " & _
"OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & _
"'Excel 8.0;Database=" & FilePathName & ";HDR=YES', " & _
"'SELECT * FROM [Sheet1$]')"
SpCon.Execute StrSql, lngRecsAff, adExecuteNoRecords
ImportExcel = lngRecsAff
SpCon.Close

End Function
Public Function ImportExcel2(FilePathName As String, TableName As String)
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim SheetName As String
Dim StrSql As String
Dim lngRecsAff As Long
Dim SpCon As New ADODB.Connection
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

StrSql = "IF EXISTS (SELECT * FROM dbo.sysobjects where id = object_id(N'[dbo].[" & TableName & "]') "
StrSql = StrSql & "and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
StrSql = StrSql & "DROP TABLE [dbo].[" & TableName & "]"
SpCon.Execute StrSql
SpCon.Close
Set SpCon = Nothing

Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & FilePathName & ";" & _
        "Extended Properties=Excel 8.0"
    
'Import by using Jet Provider.
Set rs = cn.OpenSchema(adSchemaTables)
SheetName = rs(2)

StrSql = "SELECT * INTO [ODBC;" & gblFileName & ";]." & TableName & " FROM [" & SheetName & "]"
cn.Execute StrSql, lngRecsAff, adExecuteNoRecords
ImportExcel2 = lngRecsAff
rs.Close
Set rs = Nothing
cn.Close
Set cn = Nothing

End Function
