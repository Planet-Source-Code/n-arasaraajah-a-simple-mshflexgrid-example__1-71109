Attribute VB_Name = "modDeclare"
Option Explicit

'Database Settings...
Public connDB As ADODB.Connection
Public recSet As ADODB.Recordset
Public VIDEO_DB As String

'Log Settings...
Public PATH_DEBUG As String
Public Const FILENAME_DEBUG = "Log"

Public Function InitializeAll()
On Error GoTo Err_Handler

    VIDEO_DB = "Provider=Microsoft.Jet.OleDB.4.0;Jet OLEDB:Database Password=raja;Data Source=" & App.Path & "\VideoDB.mdb"
    
    PATH_DEBUG = App.Path & "\Logs\"
    
    Exit Function

Err_Handler:
    LogError "modDeclare", "InitializeAll", Err.Number, Err.Description, Erl()
    Resume Next
End Function
