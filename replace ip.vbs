Option Explicit

' ============================
' Helper Functions
' ============================

' Get clipboard text using PowerShell's Get-Clipboard.
Function GetClipboardText()
    Dim shell, exec, clipText
    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec("powershell -command Get-Clipboard")
    clipText = exec.StdOut.ReadAll()
    GetClipboardText = clipText
End Function

' Set clipboard text using an HTML file object.
Sub SetClipboardText(newText)
    Dim html
    Set html = CreateObject("htmlfile")
    html.ParentWindow.ClipboardData.SetData "Text", newText
End Sub

' Generate a random filename given a prefix and extension.
Function GenerateRandomFileName(prefix, ext)
    Dim randomNumber, fileName
    Randomize
    randomNumber = Int((1000000 * Rnd) + 1)
    fileName = prefix & "_" & randomNumber & "." & ext
    GenerateRandomFileName = fileName
End Function

' ============================
' Initialization & Setup
' ============================
Dim clipboardText, modifiedText, totalIPs, resolvedCount
totalIPs = 0
resolvedCount = 0

' Dictionaries to store resolved and unresolved IPs.
Dim resolvedMapping, unresolvedIPs
Set resolvedMapping = CreateObject("Scripting.Dictionary")
Set unresolvedIPs = CreateObject("Scripting.Dictionary")

' Use the Windows temporary folder for all files.
Dim tempFolder, shell, fso, timestamp, errorLogFileName
Set shell = CreateObject("WScript.Shell")
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%")
If tempFolder = "" Or tempFolder = "%TEMP%" Then
    WScript.Echo "Error: Temporary folder not found. Please ensure the %TEMP% environment variable is set."
    WScript.Quit
End If
If Right(tempFolder, 1) <> "\" Then tempFolder = tempFolder & "\"

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(tempFolder) Then
    On Error Resume Next
    fso.CreateFolder(tempFolder)
    If Err.Number <> 0 Then
        WScript.Echo "Error: Cannot create the temporary folder: " & tempFolder
        WScript.Quit
    End If
    On Error GoTo 0
End If

' Create an error log file in the temp folder.
timestamp = CStr(Now)
timestamp = Replace(timestamp, ":", "-")
timestamp = Replace(timestamp, " ", "_")
timestamp = Replace(timestamp, "/", "-")
timestamp = Replace(timestamp, "\", "-")
errorLogFileName = tempFolder & "error_log_" & timestamp & ".txt"
Dim errorLogFile
Set errorLogFile = fso.CreateTextFile(errorLogFileName, True)

' Retrieve the clipboard content.
clipboardText = GetClipboardText()
modifiedText = clipboardText

' ============================
' Extract IPv4 Addresses
' ============================
Dim re, matches, i, currentIP
Set re = New RegExp
' Regex pattern for matching standard IPv4 addresses.
re.Pattern = "(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){3}"
re.Global = True

Set matches = re.Execute(clipboardText)
totalIPs = matches.Count

' --- Debug Section: Build a list of all found IP addresses ---
Dim debugIPs
debugIPs = "Debug Info - Found IP addresses:" & vbCrLf
For i = 0 To matches.Count - 1
    debugIPs = debugIPs & matches(i).Value & vbCrLf
Next
' Append the debug info to the modified text.
modifiedText = modifiedText & vbCrLf & debugIPs & vbCrLf

' ============================
' Process Each IP Address
' ============================
Dim xmlHTTP, dom, hostNode, hostname, apiUrl
For i = 0 To matches.Count - 1
    currentIP = matches(i).Value
    If Not resolvedMapping.Exists(currentIP) And Not unresolvedIPs.Exists(currentIP) Then
        apiUrl = "https://infobloxgm.com/wapi/v2.10/record:host?ipv4addr=" & currentIP

        On Error Resume Next
        Set xmlHTTP = CreateObject("MSXML2.XMLHTTP")
        xmlHTTP.Open "GET", apiUrl, False
        xmlHTTP.Send
        If Err.Number <> 0 Then
            errorLogFile.WriteLine "Error requesting URL for IP " & currentIP & ": " & Err.Description
            Err.Clear
            unresolvedIPs.Add currentIP, "HTTP request error"
            On Error GoTo 0
        Else
            On Error GoTo 0
            If xmlHTTP.Status = 200 Then
                Dim responseXML
                responseXML = xmlHTTP.responseText
                Set dom = CreateObject("MSXML2.DOMDocument")
                dom.async = False
                dom.loadXML(responseXML)
                If dom.parseError.errorCode <> 0 Then
                    errorLogFile.WriteLine "XML parsing error for IP " & currentIP & ": " & dom.parseError.reason
                    unresolvedIPs.Add currentIP, "XML parsing error"
                Else
                    Set hostNode = dom.selectSingleNode("//host")
                    If Not hostNode Is Nothing Then
                        hostname = hostNode.text
                        resolvedMapping.Add currentIP, hostname
                        resolvedCount = resolvedCount + 1
                    Else
                        errorLogFile.WriteLine "Hostname element not found for IP " & currentIP
                        unresolvedIPs.Add currentIP, "Hostname not found"
                    End If
                End If
            Else
                errorLogFile.WriteLine "HTTP error for IP " & currentIP & ": Status " & xmlHTTP.Status
                unresolvedIPs.Add currentIP, "HTTP error status: " & xmlHTTP.Status
            End If
        End If
    End If
Next

' ============================
' Replace IPs with Resolved Hostnames in Output
' ============================
Dim key
For Each key In resolvedMapping.Keys
    modifiedText = Replace(modifiedText, key, resolvedMapping(key))
Next

' ============================
' Append Summary to the Output
' ============================
Dim summaryText
summaryText = vbCrLf & "---- Summary ----" & vbCrLf
summaryText = summaryText & "Total IP addresses processed: " & totalIPs & vbCrLf
summaryText = summaryText & "Hostnames successfully resolved: " & resolvedCount & vbCrLf
If unresolvedIPs.Count > 0 Then
    summaryText = summaryText & "Unresolved IP addresses:" & vbCrLf
    For Each key In unresolvedIPs.Keys
        summaryText = summaryText & key & " - " & unresolvedIPs(key) & vbCrLf
    Next
Else
    summaryText = summaryText & "All IP addresses resolved successfully." & vbCrLf
End If
summaryText = summaryText & vbCrLf & "IP to Hostname Mapping:" & vbCrLf
For Each key In resolvedMapping.Keys
    summaryText = summaryText & key & " -> " & resolvedMapping(key) & vbCrLf
Next
modifiedText = modifiedText & summaryText

' ============================
' Update Clipboard with Modified Text
' ============================
On Error Resume Next
SetClipboardText(modifiedText)
If Err.Number <> 0 Then
    errorLogFile.WriteLine "Error setting clipboard text: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

' ============================
' Write Output to an HTA File
' ============================
Dim htaFileName, htaFile, htaContent
htaFileName = tempFolder & GenerateRandomFileName("output", "hta")

' Build the HTA content:
' - The outer window is set to 300x150 (half the previous size).
' - The outer window has no scrollbars (overflow hidden).
' - The window is resizable and has maximize/minimize buttons.
' - The textarea fills the window (100% width and height) with its own scrollbars.
htaContent = "<html>" & vbCrLf & _
    "<head>" & vbCrLf & _
    "  <title>Data Output</title>" & vbCrLf & _
    "  <HTA:APPLICATION " & vbCrLf & _
    "    ID='DataViewer' " & vbCrLf & _
    "    APPLICATIONNAME='DataViewer' " & vbCrLf & _
    "    BORDER='thin' " & vbCrLf & _
    "    CAPTION='yes' " & vbCrLf & _
    "    RESIZABLE='yes' " & vbCrLf & _
    "    MAXIMIZEBUTTON='yes' " & vbCrLf & _
    "    MINIMIZEBUTTON='yes' " & vbCrLf & _
    "    WINDOWWIDTH='300' " & vbCrLf & _
    "    WINDOWHEIGHT='150' " & vbCrLf & _
    "    SINGLEINSTANCE='yes' " & vbCrLf & _
    "    SHOWINTASKBAR='yes'>" & vbCrLf & _
    "  </HTA:APPLICATION>" & vbCrLf & _
    "  <style>" & vbCrLf & _
    "    html, body { width: 100%; height: 100%; margin: 0; padding: 0; overflow: hidden; }" & vbCrLf & _
    "    textarea { width: 100%; height: 100%; box-sizing: border-box; font-family: sans-serif; overflow: auto; }" & vbCrLf & _
    "  </style>" & vbCrLf & _
    "</head>" & vbCrLf & _
    "<body>" & vbCrLf & _
    "  <textarea id='dataText' readonly='true'>" & modifiedText & "</textarea>" & vbCrLf & _
    "</body>" & vbCrLf & _
    "</html>"

Set htaFile = fso.CreateTextFile(htaFileName, True)
htaFile.Write htaContent
htaFile.Close

errorLogFile.Close

' Launch the HTA to display the output.
shell.Run "mshta.exe """ & htaFileName & """", 1, False

' End of Script.
