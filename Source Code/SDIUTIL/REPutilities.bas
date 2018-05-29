Attribute VB_Name = "RepUtilities"
Option Explicit

' Global variables
Public cr As craxdrt.Report
Public adoRs As ADODB.Recordset
Public cnn As ADODB.Connection
Public CRApp As New craxdrt.Application
Public CrystalPrintingStatus As craxdrt.PrintingStatus
Public CrystalExportOptions As craxdrt.ExportOptions
Global BuiltInToolbarsAvailable As String   ' The state of the toolbars when this app was started
Public Result As Variant                    ' Stores return value of functions
Global gblOpenComp As String * 1
Global gblFileName As String
Global gblReadyMsg As String          ' used in status messages
Global gblOptNo As Integer
Global gblLoginName As String * 10
Global gblPassword As String * 10
Global gblUserLevel As Integer
Global gblFileKey As String
Global gblOptions As Integer
Global gblBookmark As Variant
Global gblEditStat As Integer
Global gblHold As String                ' used in cut & paste operations
Global gblCompName As String * 50
Global gblVersn As String * 3
Global gblRelease As String * 4
Public gblSerial As String * 8
Public gblSiteId As String
Public gblReply As Integer
Public PubOK As Boolean
Public gblDate1 As Date
Public gblDate2 As Date
Public gblPassData As String
' Type RECT.

Type RECT
    left As Integer
    top As Integer
    right As Integer
    bottom As Integer
End Type

' Windows API Declarations.

Declare Function GetActiveWindow Lib "User" () As Integer
Declare Function GetClassName Lib "User" (ByVal hwnd As Integer, ByVal stBuf$, ByVal cch As Integer) As Integer
Declare Function GetDesktopWindow Lib "User" () As Integer
Declare Function GetParent Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function GetWindowRect Lib "User" (ByVal hwnd As Integer, rc As RECT) As Integer
Declare Function IsIconic Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

' Constants used in the functions above.

Const SW_RESTORE = 9
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_OWNER = 4
Const GW_CHILD = 5
Const LOGPIXELSX = 88
Const LOGPIXELSY = 90

Const MF_BYCOMMAND = &H0
Const MF_BYPOSITION = &H400
Const MF_ENABLED = &H0
Const MF_GRAYED = &H1
Const MF_DISABLED = &H2
Const MF_MENUBREAK = &H40
Const MF_CHECKED = &H8
Const MF_UNCHECKED = &H0


' Message box types

Global Const MB_OKCANCEL = &H1
Global Const MB_ABORTRETRYIGNORE = &H2
Global Const MB_YESNOCANCEL = &H3
Global Const MB_YESNO = &H4
Global Const MB_RETRYCANCEL = &H5

' Message box default buttons

Global Const MB_DEFBUTTON1 = &H0
Global Const MB_DEFBUTTON2 = &H100
Global Const MB_DEFBUTTON3 = &H200

' Message box return values

Global Const MB_OK = 1
Global Const MB_CANCEL = 2
Global Const MB_ABORT = 3
Global Const MB_RETRY = 4
Global Const MB_IGNORE = 5
Global Const MB_YES = 6
Global Const MB_NO = 7

' Config IDs - used to lookup values in the [Config] table.

Global Const CONFIG_ID_HELPFILE_NAME = 0
Global Const CONFIG_ID_VERSION = 1
Global Const CONFIG_ID_DEFAULTDIR = 2
Global Const CONFIG_ID_LIVEDB_NAME = 3
Global Const CONFIG_ID_SAMPDB_NAME = 4
Global Const CONFIG_ID_ATTACHED_TABLE_NAME = 5
Global Const CONFIG_ID_USERS_TABLE = 5
Global Const CONFIG_ID_APPLICATION_NAME = 6
Global Const CONFIG_ID_LOGINNAME_COLUMN = 7
Global Const CONFIG_ID_PASSWORD_COLUMN = 8

Global Const gblABORT = 3
Global Const gblBLACK = 0
Global Const gblBLUE = 16711680
Global Const gblCASEINSENSITIVE = 1
Global Const gblGRAY = 12632256
Global Const gblGREEN = 32768
Global Const gblHOTPINK = 16711935
Global Const gblMF_BYPOSITION = &H400
Global Const gblMF_CHECKED = 8
Global Const gblMF_DISABLED = 2
Global Const gblMF_ENABLED = 0
Global Const gblMF_GRAYED = 1
Global Const gblMF_UNCHECKED = 0
Global Const gblRED = 255
Global Const gblRETRY = 4
Global Const gblSB_BOTH = 3
Global Const gblSB_CTL = 2
Global Const gblSB_HORZ = 0
Global Const gblSB_INIT = 1
Global Const gblSB_REMOVE = 3
Global Const gblSB_RESET = 5
Global Const gblSB_SETTEXT = 4
Global Const gblSB_UPDATE = 2
Global Const gblSB_VERT = 1
Global Const gblTEAL = 8421376
Global Const gblWF_8087 = &H400
Global Const gblWF_CPU286 = &H2
Global Const gblWF_CPU386 = &H4
Global Const gblWF_CPU486 = &H8
Global Const gblWF_ENHANCED = &H20
Global Const gblWF_STANDARD = &H10
Global Const gblWHITE = 16777215
Global Const gblYELLOW = 65535
Global Const gblScreen = 0                 ' used to center forms on screen
Global Const gblMDIFORM = 1                ' used to center forms on frmMDI
Global Const gblNULL_STR = ""

Global Const gblWaitMsg = "Wait..."         ' used in status messages
Global Const gblViewOnly = 3                ' used to restrict access

 
 
  





'----------------------------------------------------------------------
' CloseAllForms
'
'   Closes all open forms except the form specified
'   by the FormName parameter.
'----------------------------------------------------------------------
Function CloseAllForms(FormName As String)
On Error GoTo CloseAllForms_Err

    Dim i As Integer
    Dim F As String
            
    ' Close all open forms.

    
    For i = Forms.Count - 1 To 0 Step -1
        If Forms(i).Name <> FormName Then
           Unload Forms(i)
           Set Forms(i) = Nothing
        End If
    Next i
    CloseAllForms = -1
CloseAllForms_Exit:
    Exit Function

CloseAllForms_Err:
    csvShowError "CloseAllForms"
    CloseAllForms = 0
    Resume CloseAllForms_Exit

End Function


'----------------------------------------------------------------------------
' ConvertNulls
'
'   Converts the specified variant to a new value if it is null,
'   otherwise it returns the variant.
'----------------------------------------------------------------------------
Function ConvertNulls(v As Variant, subs As Variant) As Variant
    ConvertNulls = IIf(IsNull(v), subs, v)
End Function


Public Sub csvLogError(subName As String, errnum As Integer, errDesc As String)
On Error Resume Next
Dim X As Integer

X = FreeFile
Open App.Path & "\siserrlog.txt" For Append As #X
Write #X, subName, Now, gblLoginName, errnum, errDesc
Close #X
End Sub

Sub csvShowError(frm)
'---------------------------------------------------
' Show Errors which are unexpected
'---------------------------------------------------
Dim errLoop As Error
Dim i As Integer
Dim errs1 As Errors
Dim sMsg As String
'''Set cnn = New ADODB.Connection
i = 1
sMsg = ""
sMsg = sMsg & vbCrLf & " Vb error # " & Str(Err.Number) & _
       vbCrLf & " Generated by " & Err.Source & _
       vbCrLf & " Description " & Err.Description
Set errs1 = cnn.Errors
For Each errLoop In errs1
       With errLoop
          sMsg = sMsg & vbCrLf & "Error #" & i & ";" _
                   & vbCrLf & "  ADO Error   #" & .Number _
                   & vbCrLf & "  Description  " & .Description _
                   & vbCrLf & "  Source       " & .Source
          i = i + 1
       End With
  Next errLoop
  MsgBox sMsg, vbCritical, frm
  On Error Resume Next
 End Sub
 Sub csvShowUsrErr(errno, frm, Optional addmsg = "")
 '-----------------------------------------------
 ' Shows data entry & anticipated errors to users
 ' Users Error message file
 ' errno = Error Code generated by form
 ' frm   = name of the form where error occured
 ' addmsg = optional string to attach to message
 '-----------------------------------------------
 Dim Criteria As String, msg As String, al As Integer
 Dim rsErr As New ADODB.Recordset
 Dim errLoop As Error
 Dim errs1 As Error
 Dim qSQL As String
 On Error GoTo csvShowUsrErr_Err
 qSQL = "SELECT * from ERRMSG where "
 qSQL = qSQL & "ERRCDE = " & errno
 rsErr.Open qSQL, cnn, , , adCmdText
 With rsErr
    If .RecordCount = 0 Then   ' record not found
      msg = "Unexpected error"
      MsgBox msg, vbCritical, frm
  Else
      msg = !errdes
      If Not IsNothing(addmsg) Then
         msg = msg + " " + addmsg
      End If
      If Not IsNull(!errdes2) Then
        msg = msg + vbCrLf + !errdes2
      End If
      Select Case !alert
        Case "C"
           al = vbCritical
        Case "I"
          al = vbInformation + vbYesNo
        Case "E"
          al = vbExclamation
         Case Else
      End Select
        
      MsgBox msg, al, frm
      If al = vbInformation + vbYesNo Then
        If vbYes Then
           CloseAllForms (frm)
        End If
      End If
  End If
  .Close
  Set rsErr = Nothing
 End With
 
csvShowUsrErr_Exit:
  Exit Sub
csvShowUsrErr_Err:
  Call csvShowError("csvshowusrerr")
  Call csvLogError("csvshowusererr", Err.Number, Err.Description)
  
  Resume csvShowUsrErr_Exit
  
 End Sub
 
'------------------------------------------------------------------------
' GetScreenSize
'
'   Stores the screen size in r (a rectangle)
'------------------------------------------------------------------------
Function GetScreenSize(r As RECT) As Integer
    Dim hwnd As Integer

    hwnd = GetDesktopWindow()
    GetScreenSize = GetWindowRect(hwnd, r)
End Function
Function IsLeapYear(intYear As Integer) As Integer
'---------------------------------------------------
'-- Determines if the specified year is a leap year
'-- intYear includes the century
'-------------------------------------------------
On Error GoTo IsLeapYear_Err
 IsLeapYear = False
 '--
 If intYear Mod 4 = 0 Then   ' it is div by 4
    If intYear Mod 100 = 0 Then  ' it is a century
       If intYear Mod 400 = 0 Then ' it is a leap year
          IsLeapYear = True
        End If
    Else
       IsLeapYear = True
    End If
 End If
 '--
IsLeapYear_Exit:
  Exit Function
IsLeapYear_Err:
  MsgBox "Error: " & Err & ". " & Error$
  Resume IsLeapYear_Exit
End Function

'----------------------------------------------------------------------
' IsLoaded
'
'   Returns TRUE if the given form is loaded.
'----------------------------------------------------------------------
Function Isloaded(FormName)
    Dim i
    
    Isloaded = False
    
    For i = 0 To Forms.Count - 1
     
       If FormName = Forms(i).Name Then
           Isloaded = True
           Exit Function
       End If
    Next i
End Function

'----------------------------------------------------------------------
' IsNewRecord
'
'   Returns TRUE if the current record is the new record.
'----------------------------------------------------------------------
Function IsNewRecord(frm As Form) As Integer
On Error GoTo IsNewRecord_Err
    
    Dim BkMark As String
                    
    ' This should cause an error if the current record is a new record.

    On Error Resume Next
    BkMark = frm.Bookmark
    If Err Then
        IsNewRecord = True
    Else
        IsNewRecord = False
    End If

IsNewRecord_Exit:
    Exit Function

IsNewRecord_Err:
    MsgBox Error$
    Resume IsNewRecord_Exit

End Function

'------------------------------------------------------------------------
' IsNothing
'
'   Returns TRUE if the value passed in is Empty, Null or a zero
'   length string.
'------------------------------------------------------------------------
Function IsNothing(v As Variant) As Integer

    ' IsNothing starts out as FALSE.  We
    ' determine if v is Nothing by checking
    ' its VarType

    IsNothing = False
    Select Case VarType(v)
        Case vbEmpty
            IsNothing = True

        Case vbNull
            IsNothing = True

        Case vbString
            If Len(v) = 0 Then
                IsNothing = True
            End If
        Case Else
            IsNothing = False

    End Select

End Function



Function LogOff() As Integer
  Dim rs As ADODB.Recordset
  Dim Password As String
  Dim Criteria As String
  Dim isloggedon As Integer, n As Integer
  Dim qSQL As String
  On Error GoTo DBErrorHandler
  qSQL = "SELECT * from USERS where "
  qSQL = qSQL & "Systemname = '" & RTrim(gblLoginName) & "'"
  Set rs = New ADODB.Recordset
  On Error GoTo TableErrorHandler
  rs.Open qSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
  If Not rs.EOF Then
     rs![isloggedon] = False
     rs.Update
  End If
  rs.Close
  Set rs = Nothing
 ' End
Exit_LogOff:
 Exit Function
DBErrorHandler:
    MsgBox "Can't open database.", vbExclamation, "Logout"
    Resume Exit_LogOff
    
TableErrorHandler:
    MsgBox "Can't read USERS Table", vbExclamation, "Login"
    Resume Exit_LogOff
    

End Function
Function schMsg(ByVal strMessage As String, ByVal varButtonArg As Variant, ByVal strTitle As String)
Dim strTemp As String
Dim intPos As Integer
strTemp = ""
Do
  intPos = InStr(strMessage, "|")
    If intPos > 0 Then
     strTemp = strTemp & left(strMessage, intPos - 1)
     strTemp = strTemp & vbCrLf
     strMessage = Mid(strMessage, intPos + 1)
  End If
Loop Until intPos = 0
strTemp = strTemp & strMessage
schMsg = MsgBox(strTemp, varButtonArg, strTitle)
End Function

'----------------------------------------------------------------------
' StripFileAndReturnPath
'
'   Strips the file name off of a full DOS path.
'----------------------------------------------------------------------
Function StripFileAndReturnPath(FullPath As String, KeepBackSlash As Integer) As String
    
    Dim X As Integer
    Dim Z As Integer
    Dim Bslash As String
    Dim filename As String
    Dim PathOnly As String
    X = Len(FullPath)
    Z = 0
    Do
        Z = Z + 1
        Bslash = Mid(FullPath, X, 1)
        If Bslash = "\" Then Exit Do
        X = X - 1
    Loop Until X = 0
   Z = Z - 1
   If Bslash = "\" Then
      filename = right(FullPath, Z)
   Else
      filename = FullPath
   End If
   
   If (Not KeepBackSlash) Then
        PathOnly = left$(FullPath, Len(FullPath) - Z - 1)
   Else
        PathOnly = left$(FullPath, Len(FullPath) - Z)
   End If
   StripFileAndReturnPath = PathOnly

End Function

'----------------------------------------------------------------------
' StripPathAndReturnFileName
'
'   Strips the path off of a full DOS path.
'--------------------------------------------------------------------
Function StripPathAndReturnFileName(FullPath As String) As String
    'StripPathAndReturnFileName = wlib_stFileOfFullPath(FullPath)
    Dim X As Integer
    Dim Z As Integer
    Dim Bslash As String
    Dim filename As String
    X = Len(FullPath)
    Z = 0
    Do
        Z = Z + 1
        Bslash = Mid(FullPath, X, 1)
        If Bslash = "\" Then Exit Do
        X = X - 1
    Loop Until X = 0
   Z = Z - 1
   If Bslash = "\" Then
      filename = right(FullPath, Z)
   Else
      filename = FullPath
   End If
   StripPathAndReturnFileName = filename
End Function

'------------------------------------------------------------------------
' StWindowClass
'
'   A simple cover function to the Windows API call.
'------------------------------------------------------------------------
Function StWindowClass(hwnd As Integer) As String

    Const cchMax = 255
    Dim Buffer As String * cchMax
    Dim cch As Integer
    cch = GetClassName(hwnd, Buffer, cchMax)
    If (hwnd = 0) Then
        StWindowClass = ""
    Else
        StWindowClass = (left$(Buffer, cch))
    End If

End Function


Public Function csvYesNo(msgno, frm)
 '-----------------------------------------------
 ' Shows data entry & anticipated errors to users
 ' Users Error message file
 ' errno = Error Code generated by form
 ' frm   = name of the form where error occured
 '-----------------------------------------------
 Dim Criteria As String, msg As String, al As Integer
 Dim rsErr As New ADODB.Recordset
 Dim errLoop As Error
 Dim errs1 As Error
 Dim qSQL As String, iResp As Integer
 csvYesNo = True
 On Error GoTo csvYesNo_Err
 qSQL = "SELECT * from ERRMSG where "
 qSQL = qSQL & "ERRCDE = " & msgno
 rsErr.Open qSQL, cnn, , , adCmdText
 With rsErr
    If .RecordCount = 0 Then   ' record not found
      msg = "Unexpected error"
      MsgBox msg, vbCritical, frm
    Else
      msg = !errdes
      If Not IsNull(!errdes2) Then
        msg = msg + vbCrLf + !errdes2
      End If
      Select Case !alert
        Case "C"
           al = vbCritical + vbYesNo
        Case "Q"
          al = vbQuestion + vbYesNo
        Case "E"
          al = vbExclamation + vbYesNo
         Case Else
      End Select
        
      iResp = MsgBox(msg, al, frm)
      If iResp = vbNo Then
           csvYesNo = False
      End If
    End If
    .Close
    Set rsErr = Nothing
 End With
csvYesNo_Exit:
  Exit Function
csvYesNo_Err:
  Call csvShowError("csvYesNo")
  Call csvLogError("csvYesNo", Err.Number, Err.Description)
  
  Resume csvYesNo_Exit
End Function

Function csvExecuteCommand(cmdTemp As ADODB.Command)
Dim errLoop As Error
On Error GoTo Execute_Err
csvExecuteCommand = True
cmdTemp.Execute
Exit Function
'--
Execute_Err:
 If cnn.Errors.Count > 0 Then
    csvExecuteCommand = False
    For Each errLoop In cnn.Errors
        MsgBox " Error Number: " & errLoop.Number & vbCr & _
                errLoop.Description
    Next errLoop
 End If
 Resume Next
End Function


Public Static Function csvADODML(sql)
 ' Function uses input data manulipation query
 ' to update or remove records from a recordset
 ' no validation is performed on the input query.
 ' Returns true if operation successful else false
 ' the connection cnn must be opened by the calling routine
 '--
 Dim qDMLQry As String
 Dim cmdDML As ADODB.Command
 '---
 csvADODML = False
 Set cmdDML = New ADODB.Command
 Set cmdDML.ActiveConnection = cnn
 cmdDML.CommandText = sql
 cnn.Errors.Clear
 csvADODML = csvExecuteCommand(cmdDML)
 Set cmdDML = Nothing
 
End Function
Public Function RunSP2(SPCon As ADODB.Connection, sProcName As String, SpType As Integer, _
Optional aPara1 As Variant, Optional aPara2 As Variant, Optional aPara3 As Variant, _
Optional aPara4 As Variant) As Variant

Dim spCmd As ADODB.Command
Dim spPara As ADODB.Parameter
Dim aPara As Variant

Set spCmd = New ADODB.Command

If SPCon.State = 1 Then
   With spCmd
        .CommandText = sProcName
        .CommandType = adCmdStoredProc
        .ActiveConnection = SPCon
   End With
End If

Set spPara = spCmd.CreateParameter("Return", adInteger, adParamReturnValue)
spCmd.Parameters.Append spPara

If IsMissing(aPara1) Then
   GoTo TypeTest
End If

If IsNumeric(aPara1) Then
   Set spPara = spCmd.CreateParameter(, adInteger, adParamInput)
Else
   Set spPara = spCmd.CreateParameter(, adVarChar, adParamInput, 50)
End If
spCmd.Parameters.Append spPara
spPara.Value = aPara1

If IsMissing(aPara2) Then
   GoTo TypeTest
End If
Set spPara = spCmd.CreateParameter(, adInteger, adParamInput)
spCmd.Parameters.Append spPara
spPara.Value = aPara2


If IsMissing(aPara3) Then
   GoTo TypeTest
End If
Set spPara = spCmd.CreateParameter(, adInteger, adParamInput)
spCmd.Parameters.Append spPara
spPara.Value = aPara3


If IsMissing(aPara4) Then
   GoTo TypeTest
End If

Set spPara = spCmd.CreateParameter(, adVarChar, adParamInput, 1)
spCmd.Parameters.Append spPara
spPara.Value = aPara4


TypeTest:
If SpType = 0 Then
      spCmd.Execute , , adAsyncExecute
Else
   Set RunSP2 = spCmd.Execute(, , adAsyncExecute)
End If

Do While spCmd.State = adStateExecuting
   Screen.MousePointer = vbHourglass
   'frmMDI.txtStatusMsg.SimpleText = "Processing, Please wait......."
   'frmMDI.txtStatusMsg.Refresh
Loop
Screen.MousePointer = vbDefault

If SpType = 0 Then
   RunSP2 = spCmd.Parameters("Return").Value
End If

Set spCmd = Nothing
Set spPara = Nothing

End Function

Public Function RunSP(SPCon As ADODB.Connection, sProcName As String, SpType As Integer, _
ParamArray aPara() As Variant) As Variant

Dim spCmd As ADODB.Command
Dim spPara As ADODB.Parameter
Dim spDataType As ADODB.DataTypeEnum
Dim i As Integer
Dim X As Integer
Dim intLength As Integer

Set spCmd = New ADODB.Command

If SPCon.State = 1 Then
   With spCmd
        .CommandText = sProcName
        .CommandType = adCmdStoredProc
        .ActiveConnection = SPCon
   End With
End If

Set spPara = spCmd.CreateParameter("Return", adInteger, adParamReturnValue)
spCmd.Parameters.Append spPara

X = UBound(aPara())

If X < 0 Then
   GoTo TypeTest
End If

SetupParameters:
For i = 0 To X
spDataType = VarType(aPara(i))

If spDataType > 1 And spDataType < 7 Then
   Set spPara = spCmd.CreateParameter(, spDataType, adParamInput)
Else
intLength = Len(aPara(i))
Set spPara = spCmd.CreateParameter(, spDataType, adParamInput, intLength)
End If
spCmd.Parameters.Append spPara
spPara.Value = aPara(i)
Next i


TypeTest:
If SpType = 0 Then
      spCmd.Execute , , adAsyncExecute
Else
   Set RunSP = spCmd.Execute(, , adAsyncExecute)
End If

Do While spCmd.State = adStateExecuting
   Screen.MousePointer = vbHourglass
Loop
Screen.MousePointer = vbDefault

If SpType = 0 Then
   RunSP = spCmd.Parameters("Return").Value
End If

Set spCmd = Nothing
Set spPara = Nothing

End Function

