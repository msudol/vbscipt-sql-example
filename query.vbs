' VBScript to connect to a remote SQL Server database, execute an SQL query from a file, and handle errors

' Read database connection details and error log path from a configuration file
Const ConfigFilePath = "C:\Path\To\Your\ConfigFile.txt"
Dim serverName, dbName, userName, password, queryFilePath, errorLogPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set configFile = fso.OpenTextFile(ConfigFilePath, 1) ' 1 for reading
Do Until configFile.AtEndOfStream
    line = configFile.ReadLine
    If InStr(line, "ServerName=") = 1 Then
        serverName = Mid(line, Len("ServerName=") + 1)
    ElseIf InStr(line, "DatabaseName=") = 1 Then
        dbName = Mid(line, Len("DatabaseName=") + 1)
    ElseIf InStr(line, "UserName=") = 1 Then
        userName = Mid(line, Len("UserName=") + 1)
    ElseIf InStr(line, "Password=") = 1 Then
        password = Mid(line, Len("Password=") + 1)
    ElseIf InStr(line, "QueryFilePath=") = 1 Then
        queryFilePath = Mid(line, Len("QueryFilePath=") + 1)
    ElseIf InStr(line, "ErrorLogPath=") = 1 Then
        errorLogPath = Mid(line, Len("ErrorLogPath=") + 1)
    End If
Loop
configFile.Close

' Create a connection string (same as before)
Dim connString
connString = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=" & dbName & ";User ID=" & userName & ";Password=" & password

' Establish the database connection (same as before)
On Error Resume Next
Set conn = CreateObject("ADODB.Connection")
conn.Open connString
If Err.Number <> 0 Then
    WriteToErrorLog "Error connecting to the database: " & Err.Description
    Err.Clear
    WScript.Quit
End If
On Error GoTo 0

' Read the SQL query from the specified file
Set queryFile = fso.OpenTextFile(queryFilePath, 1) ' 1 for reading
sqlQuery = queryFile.ReadAll
queryFile.Close

' Execute the SQL query
On Error Resume Next
conn.Execute sqlQuery
If Err.Number <> 0 Then
    WriteToErrorLog "Error executing the SQL query: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

' Clean up (same as before)
conn.Close
Set conn = Nothing

Sub WriteToErrorLog(errorMessage)
    Dim logFile
    Set logFile = fso.OpenTextFile(errorLogPath, 8, True) ' 8 for appending
    logFile.WriteLine Now & " - " & errorMessage
    logFile.Close
End Sub
