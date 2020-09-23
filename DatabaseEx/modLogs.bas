Attribute VB_Name = "modLogs"
Option Explicit

Public Function LogError(ByVal SourceForm As String, ByVal SourceProcedure As _
    String, ByVal ErrorNumber As String, ByVal ErrorDescription As String, ByVal _
    ErrorLine As Long)
    
Dim fs As Object
Dim f As Object
Dim strFilePath As String
Dim strFileName As String
Dim strError As String

On Error GoTo LogError_Error
    Set fs = CreateObject("scripting.filesystemobject")
    strFilePath = PATH_DEBUG & Format(Now, "yyMMdd") & "\"
    strFileName = FILENAME_DEBUG & "_" & Format(Date, "DDMMYYYY") & ".log"

Open_File:
    Set f = fs.OpenTextFile(strFilePath & strFileName, 8)
    f.WriteLine Format(Date & " " & Time, "dd-mmm-yyyy hh:mm:ss") & vbTab & _
        "[" & SourceForm & ":" & SourceProcedure & "]" & _
        "[Line:" & ErrorLine & "]"
    f.WriteLine vbTab & vbTab & vbTab & _
        "[Error No.:" & ErrorNumber & "]" & _
        "[Description:" & ErrorDescription & "]"
    f.WriteLine " "

    Set f = Nothing
    Set fs = Nothing

    Exit Function

LogError_Error:
    If Not fs.FolderExists(strFilePath) Then
        fs.CreateFolder (strFilePath)
    End If

    If Not fs.FileExists(strFilePath & strFileName) Then
        fs.CreateTextFile (strFilePath & strFileName)
        GoTo Open_File
    End If

    Set f = Nothing
    Set fs = Nothing
End Function
