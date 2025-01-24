Option Explicit

' This script assumes WARP to be in paused / disconnected mode.
'
' Since warp-cli no longer allows checking of auto-connect timeout value,
' the script now simply connects and immediately disconnects to refresh the value if the current status is disconnected.
' That means, this script intentionally does nothing should the user leaves the WARP in connected mode.

' Since VBScript cannot print to stdout, this script writes to a log file instead (warp-refresh.log, max 100 lines).

' To create a new Windows scheduled task (runs every 5 minutes to check, assumes this file is saved as "%UserProfile%\Documents\warp-refresh.vbs"):
' schtasks /Create /SC minute /MO 60 /TN "WARP auto-connect refresh" /TR "WScript.exe \"%UserProfile%\Documents\warp-refresh.vbs\""
' To delete the already scheduled task:
' schtasks /Delete /TN "WARP auto-connect refresh" /F

Dim disconnected: disconnected = IsDisconnected()

If disconnected = False Then
    WriteToLogFile "WARP does not seem to be in disconnected state, not performing any action"
Else
    WriteToLogFile "WARP is currently disconnected, will perform connect-disconnect to refresh timeout..."
    ConnectAndDisconnect
    WriteToLogFile "WARP auto-connect timeout refresh complete"
End If

TrimFile

Function IsDisconnected()
    Dim oShell: Set oShell = CreateObject("WScript.Shell")
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create a temporary file to hold the command output
    ' We use .Run instead of .Exec to prevent creation of another cmd shell from popping up
    Dim tempFileName: tempFileName = fso.GetSpecialFolder(2) & "\" & fso.GetTempName()
    oShell.Run "%comspec% /c warp-cli status > """ & tempFileName & """", 0, True

    ' Read the command output from the temporary file
    Dim tempFile: Set tempFile = fso.OpenTextFile(tempFileName, 1)
    Dim output: output = tempFile.ReadAll()
    tempFile.Close

    ' Delete the temporary file
    fso.DeleteFile(tempFileName)

    ' Read the first line of the output
    Dim firstLine: firstLine = Split(Split(output, vbCrLf)(0), vbLf)(0)

    ' Extract the value after "Status update: "
    Dim status: status = Trim(Split(firstLine, ":")(1))

    Dim disconnected
    If status = "Disconnected" Then
        disconnected = True
    Else
        disconnected = False
    End If

    IsDisconnected = disconnected
End Function

Sub ConnectAndDisconnect()
    Dim oShell: Set oShell = CreateObject("WScript.Shell")
    oShell.Run "warp-cli connect", 0, True
    oShell.Run "warp-cli disconnect", 0, True
End Sub

Sub WriteToLogFile(text)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logFileName: logFileName = fso.GetParentFolderName(WScript.ScriptFullName) & "\" & fso.GetBaseName(WScript.ScriptFullName) & ".log"

    ' Check if the file exists, if not, create it
    If Not fso.FileExists(logFileName) Then
        Set outputFile = fso.CreateTextFile(logFileName)
        outputFile.Close
    End If

    ' Open the file for appending
    Dim outputFile: Set outputFile = fso.OpenTextFile(logFileName, 8, True)
    outputFile.WriteLine Now & ": " & text
    outputFile.Close
End Sub

Sub TrimFile()
    Dim maxLines: maxLines = 100
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logFileName: logFileName = fso.GetParentFolderName(WScript.ScriptFullName) & "\" & fso.GetBaseName(WScript.ScriptFullName) & ".log"

    ' Check if the file exists, if not, create it
    If Not fso.FileExists(logFileName) Then
        Exit Sub
    End If

    ' Read existing lines into an array
    Dim lines: lines = Split(fso.OpenTextFile(logFileName).ReadAll(), vbCrLf)

    ' If there are too many lines, remove lines from the beginning
    Dim i
    If UBound(lines) >= maxLines Then
        For i = 0 To UBound(lines) - maxLines
            lines(i) = ""
        Next
    End If

    ' Write lines back to the file
    Dim outputFile: Set outputFile = fso.OpenTextFile(logFileName, 2, True)
    For i = 0 To UBound(lines)
        If lines(i) <> "" Then outputFile.WriteLine lines(i)
    Next
    outputFile.Close
End Sub
