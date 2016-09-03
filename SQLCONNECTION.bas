Attribute VB_Name = "Module1"
'Database Connectivity
Option Explicit
Public SERVER_PATH As String
Public db As ADODB.Connection

Public Sub SetConnection(SERVER As String, DATABASE As String, UID As String, PWD As String)
On Error GoTo ShowError
Set db = New ADODB.Connection
With db
.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & _
"SERVER=" & SERVER & ";" & _
"DATABASE=" & DATABASE & ";" & _
"UID=" & UID & ";" & _
"PWD=" & PWD & ";" & _
"OPTION = " & 1 + 2 + 8 + 32 + 2048 + 16384
.ConnectionTimeout = 30
.CursorLocation = adUseClient
.Open
End With
Exit Sub
ShowError:
MsgBox "Cannot establish connection. Please contact the Server.", vbOKOnly + vbCritical, "SERVER"
End Sub


