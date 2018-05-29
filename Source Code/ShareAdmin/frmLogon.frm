VERSION 5.00
Begin VB.Form frmLogon 
   Caption         =   "Logon/Logoff Server"
   ClientHeight    =   2715
   ClientLeft      =   5190
   ClientTop       =   3585
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtDatabaseName 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtServerName 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtDLLName 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdLogOff 
      Caption         =   "L&ogoff"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdLogOn 
      Caption         =   "&Logon"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblUserID 
      Caption         =   "User ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblDataBaseName 
      Caption         =   "Database Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblServerName 
      Caption         =   "Server Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblDLLName 
      Caption         =   "DLL Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdLogOff_Click()
On Error Resume Next

CRApp.LogOffServer Trim(txtDLLName), Trim(txtServerName), _
                    Trim(txtDatabaseName), Trim(txtUserID), _
                    Trim(txtPassword)
                    
'Check to see if an error occurred during the Logoff process
If Err.Number <> 0 Then
    MsgBox "Logoff Failed!"
Else
    MsgBox "Logoff succeeded!"
    cmdLogOff.Enabled = True
End If


End Sub

Private Sub cmdLogOn_Click()
On Error Resume Next
CRApp.LogOnServer Trim(txtDLLName), Trim(txtServerName), _
                    Trim(txtDatabaseName), Trim(txtUserID), _
                    Trim(txtPassword)

'Check to see if an error occurred during the Logon process
If Err.Number <> 0 Then
    MsgBox "Logon Failed!"
Else
    MsgBox "Logon succeeded!"
    cmdLogOff.Enabled = True
End If

End Sub

