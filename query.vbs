' VBScript to connect to a remote SQL Server database and execute an SQL query from a file

' Read database connection details from a configuration file
Const ConfigFilePath = "C:\Path\To\Your\ConfigFile.txt"
Dim serverName, dbName, userName, password, queryFilePath
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
    End If
Loop
configFile.Close

' Create a connection string (same as before)
Dim connString
connString = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=" & dbName & ";User ID=" & userName & ";Password=" & password

' Establish the database connection (same as before)
Set conn = CreateObject("ADODB.Connection")
conn.Open connString

' Read the SQL query from the specified file
Set queryFile = fso.OpenTextFile(queryFilePath, 1) ' 1 for reading
sqlQuery = queryFile.ReadAll
queryFile.Close

' Execute the SQL query
conn.Execute sqlQuery

' Clean up (same as before)
conn.Close
Set conn = Nothing
