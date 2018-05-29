VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLstWarn"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListWarn() As Integer
 Dim msg As String
 msg = "You cannot close this form while you are "
 msg = msg & "editing one of its sub-forms. "
 msg = msg & "You must close the sub-form before"
 msg = msg & " closing this form."
 MsgBox msg, vbExclamation + vbOKOnly, "Warning"
End Function
