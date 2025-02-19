Option Explicit

' ============================
' Helper Functions
' ============================

' Function to get clipboard text using an HTML file object
Function GetClipboardText()
    Dim html, clipText
    Set html = CreateObject("htmlfile")
    clipText = html.ParentWindow.ClipboardData.GetData("Text")
    GetClipboardText = clipText
End Function

' Subroutine to set clipboard text using an HTML file object
Sub SetClipboardText(newText)
    Dim html
    Set html = CreateObject("htmlfile")
    html.ParentWindow.ClipboardData.SetData "Text", newText
End Sub

' Function to generate a random filename given a prefix and extension
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
Dim clipboardText, modifiedText
Dim totalIPs, resolvedCount
totalIPs = 0
resolvedCount = 0

' Dictionaries to store resolved and unresolved IPs
Dim resolvedMapping, unresolvedIPs
Set resolvedMapping = CreateObject("Scripting.Dictionary")
Set unresolvedIPs = CreateObject("Scripting.Dictionary")

' Use the Windows temporary folder for all files.
Dim tempFolder, shell, fso, timestamp, errorLogFileName
Set shell = CreateObject("WScript.Shell")
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%") & "\"

Set fso = CreateObject("Scripting.FileSystemObject")
' (The temp folder should exist, but we can check if needed)
If Not fso.FolderExists(tempFolder) Then
    fso.CreateFolder(tempFolder)
End If

' Create an error log file in the temp folder.
timestamp = Replace(Replace(CStr(Now), ":", "-"), " ", "_")
errorLogFileName = tempFolder & "error_log_" & timestamp & ".txt"
Dim errorLogFile
Set errorLogFile = fso.CreateTextFile(errorLogFileName, True)

' Retrieve the clipboard content
clipboardText = GetClipboardText()
modifiedText = clipboardText

' ============================
' Extract IPv4 Addresses
' ============================
Dim re, matches, i, currentIP
Set re = New RegExp
' Regex to match valid IPv4 addresses
re.Pattern = "\b(?:(?:25[0-5]|2[0-4]\d|[01]?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|[01]?\d?\d)\b"
re.Global = True

Set matches = re.Execute(clipboardText)
totalIPs = matches.Count

' ============================
' Process Each IP Address
' ============================
Dim xmlHTTP, dom, hostNode, hostname, apiUrl
For i = 0 To matches.Count - 1
    currentIP = matches(i).Value
    ' Only process an IP if it hasn't been processed already.
    If Not resolvedMapping.Exists(currentIP) And Not unresolvedIPs.Exists(currentIP) Then
        apiUrl = "https://infobloxgm.com/wapi/v2.10/record:host?ipv4addr=" & currentIP

        ' Create and send the HTTP GET request.
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
            ' Process only if the HTTP status is 200 (OK)
            If xmlHTTP.Status = 200 Then
                Dim responseXML
                responseXML = xmlHTTP.responseText

                ' Load the XML response
                Set dom = CreateObject("MSXML2.DOMDocument")
                dom.async = False
                dom.loadXML(responseXML)
                If dom.parseError.errorCode <> 0 Then
                    errorLogFile.WriteLine "XML parsing error for IP " & currentIP & ": " & dom.parseError.reason
                    unresolvedIPs.Add currentIP, "XML parsing error"
                Else
                    ' Extract the hostname from the <host> element (always use hostname)
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
' Replace IPs with Resolved Hostnames
' ============================
Dim key
For Each key In resolvedMapping.Keys
    ' Replace all occurrences of the IP with its hostname
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

' Build the HTA content with a text area to display the data
htaContent = "<html>" & vbCrLf & _
    "<head>" & vbCrLf & _
    "  <title>Data Output</title>" & vbCrLf & _
    "  <HTA:APPLICATION id='DataViewer' APPLICATIONNAME='DataViewer' " & _
    "BORDER='thin' CAPTION='yes' SHOWINTASKBAR='yes' SINGLEINSTANCE='yes'>" & vbCrLf & _
    "  <style>body { font-family: sans-serif; margin: 10px; } " & _
    "textarea { width: 100%; height: 400px; }</style>" & vbCrLf & _
    "</head>" & vbCrLf & _
    "<body>" & vbCrLf & _
    "  <textarea id='dataText' readonly='true'>" & modifiedText & "</textarea>" & vbCrLf & _
    "</body>" & vbCrLf & _
    "</html>"

Set htaFile = fso.CreateTextFile(htaFileName, True)
htaFile.Write htaContent
htaFile.Close

' Close the error log file
errorLogFile.Close

' ============================
' Launch the HTA to Display the Output
' ============================
shell.Run "mshta.exe """ & htaFileName & """", 1, False

' End of Script
