VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUnPackName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mFirstName As String 'local copy
Private mLastName As String 'local copy
Private iPos As Integer, iLen As Integer
Public Function Unpack(ByVal vName As String) As Boolean
Attribute Unpack.VB_UserMemId = 0
Unpack = True
If IsNull(vName) Then
  Unpack = False
  Exit Function
Else
  iLen = Len(Trim(vName))
  iPos = InStr(1, vName, ",", vbTextCompare)
  If iPos = 0 Then ' no comma found
    Unpack = False
  Else
    mLastName = Mid(vName, 1, iPos - 1)
    mFirstName = Mid(vName, iPos + 1, iLen - iPos)
  End If
 End If
End Function
Public Property Get LastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastName
    LastName = mLastName
End Property

Public Property Get FirstName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FirstName
    FirstName = mFirstName
End Property



