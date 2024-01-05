Option Explicit

' This script assumes WARP to be in paused / disconnected mode.
'
' Checks if your auto-connect timeout value has crossed a threshold (6 mins left), if so,
' connects and immediately disconnects to refresh the value.
'
' This script intentionally does nothing should the user leaves the WARP in connected mode.
' (In connected mode, the auto-connect timeout value is still present but does not decrease over time.)

' Since VBScript cannot print to stdout, this script writes to a log file instead (warp-refresh.log, max 100 lines).

' To create a new Windows scheduled task (runs every 5 minutes to check, assumes this file is saved as "%UserProfile%\Documents\warp-refresh.vbs"):
' schtasks /Create /SC minute /MO 5 /TN "WARP auto-connect refresh" /TR "%UserProfile%\Documents\warp-refresh.vbs"
' To delete the already scheduled task:
' schtasks /Delete /TN "WARP auto-connect refresh" /F

Dim secs_req, secs_left, no_secs, i
secs_req = 360
no_secs = 2147483647

' Loop check-refresh up to 5 tries (with 3 seconds delay in between)
For i = 1 To 5
    secs_left = GetSecsLeft()

    If secs_left = no_secs Then
        WriteToLogFile "WARP does not seem to be in disconnected state, not performing any action"
        Exit For
    ElseIf secs_left > secs_req Then
        WriteToLogFile "WARP auto-connect timeout is " & secs_left & "s, more than " & secs_req & "s, no actions required"
        Exit For
    Else
        WriteToLogFile "WARP auto-connect timeout is " & secs_left & "s, equal or less than " & secs_req & "s, will attempt to refresh..."
        ConnectAndDisconnect
        WriteToLogFile "WARP auto-connect refresh complete"
    End If

    WScript.Sleep 3000
Next

TrimFile

Function GetSecsLeft()
    Dim oShell, tempFileName, fso, tempFile, output, secs_left, outputArray
    Set oShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create a temporary file to hold the command output
    ' We use .Run instead of .Exec to prevent creation of another cmd shell from popping up
    tempFileName = fso.GetSpecialFolder(2) & "\" & fso.GetTempName()
    oShell.Run "%comspec% /c warp-cli get-pause-end > """ & tempFileName & """", 0, True

    ' Read the command output from the temporary file
    Set tempFile = fso.OpenTextFile(tempFileName, 1)
    output = tempFile.ReadAll()
    tempFile.Close

    ' Delete the temporary file
    fso.DeleteFile(tempFileName)

    outputArray = Split(output, " ")

    ' And operator in VBScript is not short-circuiting, so cannot combine
    If UBound(outputArray) >= 3 Then
        If IsNumeric(outputArray(3)) Then
            ' Already in disconnected state (or disconnected before) and has a remaining auto-connect timeout
            secs_left = CInt(outputArray(3))
        Else
            ' Unexpected case, perhaps the command really got deprecated
            secs_left = no_secs
        End If
    Else
        ' Never in disconnected state before
        secs_left = no_secs
    End If

    GetSecsLeft = secs_left
End Function

Sub ConnectAndDisconnect()
    Dim oShell
    Set oShell = CreateObject("WScript.Shell")
    oShell.Run "warp-cli connect", 0, True
    oShell.Run "warp-cli disconnect", 0, True
End Sub

Sub WriteToLogFile(text)
    Dim fso, outputFile, logFileName
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFileName = fso.GetParentFolderName(WScript.ScriptFullName) & "\" & fso.GetBaseName(WScript.ScriptFullName) & ".log"

    ' Check if the file exists, if not, create it
    If Not fso.FileExists(logFileName) Then
        Set outputFile = fso.CreateTextFile(logFileName)
        outputFile.Close
    End If

    ' Open the file for appending
    Set outputFile = fso.OpenTextFile(logFileName, 8, True)
    outputFile.WriteLine Now & ": " & text
    outputFile.Close
End Sub

Sub TrimFile()
    Dim fso, inputFile, outputFile, lines, i, maxLines, logFileName
    maxLines = 100
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFileName = fso.GetParentFolderName(WScript.ScriptFullName) & "\" & fso.GetBaseName(WScript.ScriptFullName) & ".log"

    ' Check if the file exists, if not, create it
    If Not fso.FileExists(logFileName) Then
        Exit Sub
    End If

    ' Read existing lines into an array
    lines = Split(fso.OpenTextFile(logFileName).ReadAll(), vbCrLf)

    ' If there are too many lines, remove lines from the beginning
    If UBound(lines) >= maxLines Then
        For i = 0 To UBound(lines) - maxLines
            lines(i) = ""
        Next
    End If

    ' Write lines back to the file
    Set outputFile = fso.OpenTextFile(logFileName, 2, True)
    For i = 0 To UBound(lines)
        If lines(i) <> "" Then outputFile.WriteLine lines(i)
    Next
    outputFile.Close
End Sub
