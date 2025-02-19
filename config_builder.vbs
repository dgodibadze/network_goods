<html>
<head>
  <title>Configuration Builder v2.0</title>
  <HTA:APPLICATION
    APPLICATIONNAME="ConfigBuilder"
    BORDER="thick"
    BORDERSTYLE="complex"
    CAPTION="yes"
    SHOWINTASKBAR="yes"
    SINGLEINSTANCE="yes"
    SYSMENU="yes"
    WINDOWSTATE="normal">

  <script language="VBScript">
    ' Global objects and folder names
    Dim fso, basePath, configFolder, valueFolder, isSyncing
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = fso.GetAbsolutePathName(".")
    configFolder = basePath & "\Config_Templates"
    valueFolder = basePath & "\Value_Templates"
    isSyncing = False

    ' Create folders if they do not exist
    Sub EnsureFoldersExist()
      If Not fso.FolderExists(configFolder) Then fso.CreateFolder(configFolder)
      If Not fso.FolderExists(valueFolder) Then fso.CreateFolder(valueFolder)
    End Sub

    ' *********************
    ' Dynamic Resize Function for Textareas
    ' *********************
    Sub ResizeTextareas()
      Dim availableHeight, offset
      ' Offset accounts for headers, buttons, margins, etc.
      offset = 350
      availableHeight = Document.body.clientHeight - offset
      If availableHeight < 100 Then availableHeight = 100
      Document.getElementById("configEditor").style.height = availableHeight & "px"
      Document.getElementById("configOutput").style.height = availableHeight & "px"
    End Sub

    ' *********************
    ' Value Template Functions (Plain Text Format)
    ' *********************
    Sub LoadValueTemplates()
      Dim folder, file, dropdown, opt, defaultOpt
      Set dropdown = Document.getElementById("valueTemplateSelect")
      Do While dropdown.options.length > 0
        dropdown.remove 0
      Loop
      Set defaultOpt = Document.createElement("OPTION")
      defaultOpt.value = ""
      defaultOpt.text = "Load Value Template"
      defaultOpt.disabled = True
      defaultOpt.selected = True
      dropdown.options.add defaultOpt

      If fso.FolderExists(valueFolder) Then
        Set folder = fso.GetFolder(valueFolder)
        For Each file In folder.Files
          If LCase(fso.GetExtensionName(file.Path)) = "txt" Then
            Set opt = Document.createElement("OPTION")
            opt.value = file.Name
            opt.text = file.Name
            dropdown.options.add opt
          End If
        Next
      End If
    End Sub

    Sub LoadValueTemplateFile(fileName)
      Dim filePath, file, content, lines, i, parts, lineBreak
      filePath = fso.BuildPath(valueFolder, fileName)
      If Not fso.FileExists(filePath) Then
        MsgBox "Value template file not found: " & filePath
        Exit Sub
      End If
      Set file = fso.OpenTextFile(filePath, 1)
      content = file.ReadAll
      file.Close
      Document.getElementById("valueFieldsContainer").innerHTML = ""
      If InStr(content, vbCrLf) > 0 Then
        lineBreak = vbCrLf
      ElseIf InStr(content, vbCr) > 0 Then
        lineBreak = vbCr
      Else
        lineBreak = vbLf
      End If
      lines = Split(content, lineBreak)
      For i = 0 To UBound(lines)
        If Trim(lines(i)) <> "" Then
          parts = Split(lines(i), "|")
          If UBound(parts) >= 2 Then
            Call AddValueField(parts(0), parts(1), parts(2))
          End If
        End If
      Next
    End Sub

    Sub LoadSelectedValueTemplate()
      Dim dropdown, fileName
      Set dropdown = Document.getElementById("valueTemplateSelect")
      fileName = dropdown.value
      If fileName <> "" Then Call LoadValueTemplateFile(fileName)
    End Sub

    Sub SaveValueTemplate()
      Dim container, fields, i, div, fieldName, fieldValue, activeState, fileName, filePath, file, outText, btns, j
      Set container = Document.getElementById("valueFieldsContainer")
      Set fields = container.getElementsByTagName("DIV")
      outText = ""
      For i = 0 To fields.length - 1
        Set div = fields.item(i)
        fieldName = div.getElementsByTagName("INPUT")(0).value
        fieldValue = div.getElementsByTagName("INPUT")(1).value
        activeState = "true"
        Set btns = div.getElementsByTagName("BUTTON")
        For j = 0 To btns.length - 1
          If btns.item(j).className = "pauseButton" Then
            activeState = btns.item(j).getAttribute("active")
            Exit For
          End If
        Next
        outText = outText & fieldName & "|" & fieldValue & "|" & activeState & vbCrLf
      Next
      fileName = InputBox("Enter a name for the Value Template", "Save Value Template")
      If fileName = "" Then Exit Sub
      If LCase(Right(fileName, 4)) <> ".txt" Then fileName = fileName & ".txt"
      filePath = fso.BuildPath(valueFolder, fileName)
      Set file = fso.CreateTextFile(filePath, True)
      file.Write outText
      file.Close
      MsgBox "Value Template saved."
      Call LoadValueTemplates()
    End Sub

    Sub DeleteValueTemplate()
      Dim dropdown, fileName, filePath
      Set dropdown = Document.getElementById("valueTemplateSelect")
      fileName = dropdown.value
      If fileName = "" Then
        MsgBox "Please select a value template to delete."
        Exit Sub
      End If
      If MsgBox("Are you sure you want to delete this value template?", vbYesNo) = vbNo Then Exit Sub
      filePath = fso.BuildPath(valueFolder, fileName)
      If fso.FileExists(filePath) Then fso.DeleteFile filePath
      MsgBox "Value Template deleted."
      Call LoadValueTemplates()
    End Sub

    ' *********************
    ' Config Template Functions (Plain Text)
    ' *********************
    Sub LoadConfigTemplates()
      Dim folder, file, dropdown, opt, defaultOpt
      Set dropdown = Document.getElementById("configTemplateSelect")
      Do While dropdown.options.length > 0
        dropdown.remove 0
      Loop
      Set defaultOpt = Document.createElement("OPTION")
      defaultOpt.value = ""
      defaultOpt.text = "Load Config Template"
      defaultOpt.disabled = True
      defaultOpt.selected = True
      dropdown.options.add defaultOpt

      If fso.FolderExists(configFolder) Then
        Set folder = fso.GetFolder(configFolder)
        For Each file In folder.Files
          If LCase(fso.GetExtensionName(file.Path)) = "txt" Then
            Set opt = Document.createElement("OPTION")
            opt.value = file.Name
            opt.text = file.Name
            dropdown.options.add opt
          End If
        Next
      End If
    End Sub

    Sub LoadConfigTemplateFile(fileName)
      Dim filePath, file, content, editor
      filePath = fso.BuildPath(configFolder, fileName)
      If Not fso.FileExists(filePath) Then Exit Sub
      Set file = fso.OpenTextFile(filePath, 1)
      content = file.ReadAll
      file.Close
      content = Replace(content, vbCrLf, vbLf)
      content = Replace(content, vbCr, vbLf)
      content = Replace(content, vbLf, vbCrLf)
      Set editor = Document.getElementById("configEditor")
      editor.value = content
    End Sub

    Sub LoadSelectedConfigTemplate()
      Dim dropdown, fileName
      Set dropdown = Document.getElementById("configTemplateSelect")
      fileName = dropdown.value
      If fileName <> "" Then Call LoadConfigTemplateFile(fileName)
    End Sub

    Sub SaveConfigTemplate()
      Dim editor, fileName, filePath, file, content
      Set editor = Document.getElementById("configEditor")
      content = editor.value
      fileName = InputBox("Enter a name for the Config Template", "Save Config Template")
      If fileName = "" Then Exit Sub
      If LCase(Right(fileName, 4)) <> ".txt" Then fileName = fileName & ".txt"
      filePath = fso.BuildPath(configFolder, fileName)
      Set file = fso.CreateTextFile(filePath, True)
      file.Write content
      file.Close
      MsgBox "Config Template saved."
      Call LoadConfigTemplates()
    End Sub

    Sub DeleteConfigTemplate()
      Dim dropdown, fileName, filePath
      Set dropdown = Document.getElementById("configTemplateSelect")
      fileName = dropdown.value
      If fileName = "" Then
        MsgBox "Please select a config template to delete."
        Exit Sub
      End If
      If MsgBox("Are you sure you want to delete this config template?", vbYesNo) = vbNo Then Exit Sub
      filePath = fso.BuildPath(configFolder, fileName)
      If fso.FileExists(filePath) Then fso.DeleteFile filePath
      MsgBox "Config Template deleted."
      Call LoadConfigTemplates()
    End Sub

    ' *********************
    ' Generate Configuration
    ' *********************
    Sub GenerateConfig()
      Dim editor, templateText, container, fields, i
      Set editor = Document.getElementById("configEditor")
      templateText = editor.value
      Set container = Document.getElementById("valueFieldsContainer")
      Set fields = container.getElementsByTagName("DIV")
      Dim div, inputs, btns, j, fieldName, fieldValue, activeState
      For i = 0 To fields.length - 1
        Set div = fields.item(i)
        Set inputs = div.getElementsByTagName("INPUT")
        fieldName = inputs.item(0).value
        fieldValue = inputs.item(1).value
        activeState = "true"
        Set btns = div.getElementsByTagName("BUTTON")
        For j = 0 To btns.length - 1
          If btns.item(j).className = "pauseButton" Then
            activeState = btns.item(j).getAttribute("active")
            Exit For
          End If
        Next
        If LCase(activeState) = "true" Then
          templateText = Replace(templateText, "[" & fieldName & "]", fieldValue)
        End If
      Next
      Document.getElementById("configOutput").value = templateText
    End Sub

    ' *********************
    ' Scroll Mirroring Functions
    ' *********************
    Sub MirrorScrollEditor()
      If isSyncing Then Exit Sub
      isSyncing = True
      Document.getElementById("configOutput").scrollTop = Document.getElementById("configEditor").scrollTop
      isSyncing = False
    End Sub

    Sub MirrorScrollOutput()
      If isSyncing Then Exit Sub
      isSyncing = True
      Document.getElementById("configEditor").scrollTop = Document.getElementById("configOutput").scrollTop
      isSyncing = False
    End Sub

    ' *********************
    ' Initialization on Load
    ' *********************
    Sub Window_OnLoad()
      Call EnsureFoldersExist()
      Call LoadValueTemplates()
      Call LoadConfigTemplates()
      Document.getElementById("configEditor").onscroll = GetRef("MirrorScrollEditor")
      Document.getElementById("configOutput").onscroll = GetRef("MirrorScrollOutput")
      ' Hook window resize event to resize textareas dynamically.
      window.onresize = GetRef("ResizeTextareas")
      ResizeTextareas()
    End Sub

    ' *********************
    ' Value Field Utility Functions
    ' *********************
    Sub AddValueField(fieldName, fieldValue, fieldActive)
      Dim container, div, nameInput, valueInput, pauseButton, deleteButton
      Set container = Document.getElementById("valueFieldsContainer")
      Set div = Document.createElement("DIV")
      div.className = "valueField"
      Set nameInput = Document.createElement("INPUT")
      nameInput.type = "text"
      nameInput.value = fieldName
      nameInput.className = "fieldName"
      div.appendChild nameInput
      Set valueInput = Document.createElement("INPUT")
      valueInput.type = "text"
      valueInput.value = fieldValue
      valueInput.className = "fieldValue"
      div.appendChild valueInput
      Set pauseButton = Document.createElement("BUTTON")
      pauseButton.innerText = "Pause"
      pauseButton.className = "pauseButton"
      pauseButton.setAttribute "active", fieldActive
      If LCase(fieldActive) <> "true" Then pauseButton.style.backgroundColor = "gray"
      pauseButton.onclick = GetRef("ToggleFieldActive")
      div.appendChild pauseButton
      Set deleteButton = Document.createElement("BUTTON")
      deleteButton.innerText = "Delete"
      deleteButton.className = "deleteButton"
      deleteButton.onclick = GetRef("DeleteValueField")
      div.appendChild deleteButton
      container.appendChild div
    End Sub

    Sub AddNewValueField()
      Call AddValueField("", "", "true")
    End Sub

    Sub ToggleFieldActive()
      Dim btn, currentState
      Set btn = Document.activeElement
      currentState = btn.getAttribute("active")
      If LCase(currentState) = "true" Then
        btn.setAttribute "active", "false"
        btn.style.backgroundColor = "gray"
      Else
        btn.setAttribute "active", "true"
        btn.style.backgroundColor = ""
      End If
    End Sub

    Sub DeleteValueField()
      Dim btn, parentDiv
      Set btn = Document.activeElement
      Set parentDiv = btn.parentElement
      parentDiv.parentElement.removeChild parentDiv
    End Sub
  </script>

  <style type="text/css">
    html, body {
      font-family: Arial, sans-serif;
      margin: 10px;
      overflow: hidden; /* Remove scroll bars from main window */
    }
    .section {
      margin-bottom: 20px;
      padding: 10px;
      border: 1px solid #ccc;
    }
    .valueField {
      margin-bottom: 5px;
    }
    input.fieldName, input.fieldValue {
      width: 120px;
      margin-right: 5px;
    }
    textarea {
      width: 100%;
      margin-top: 5px;
      overflow-x: hidden;
      overflow-y: auto;
      resize: vertical;
    }
    select {
      width: 200px;
    }
    button {
      margin: 2px;
    }
    h2 {
      font-size: 16px;
    }
    h3 {
      font-size: 14px;
      margin-bottom: 5px;
    }

    /* Table-based layout for first section (split into two columns) */
    .firstSectionTable {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
      margin-bottom: 10px;
    }
    .firstSectionTable td {
      width: 50%;
      vertical-align: top;
      padding: 5px;
    }

    /* Table-based layout with fixed 50/50 columns for second section */
    .editorContainer table {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
    }
    .editorContainer td {
      width: 50%;
      vertical-align: top;
      padding: 5px;
    }
  </style>
</head>

<body onload="Window_OnLoad()">
  <!-- Values Section -->
  <div class="section">
    <h2>Values Section</h2>
    <table class="firstSectionTable">
      <tr>
        <td>
          <!-- Left side: Value fields container and Add Field button -->
          <div id="valueFieldsContainer">
            <!-- Dynamic value fields will be added here -->
          </div>
          <button onclick="AddNewValueField()">+ Add Field</button>
        </td>
        <td>
          <!-- Right side: Dropdown and its Save/Delete buttons -->
          <label for="valueTemplateSelect">Select Value Template:</label>
          <select id="valueTemplateSelect" onchange="LoadSelectedValueTemplate()"></select>
          <br>
          <button onclick="SaveValueTemplate()">Save</button>
          <button onclick="DeleteValueTemplate()">Delete</button>
        </td>
      </tr>
    </table>
  </div>

  <!-- Configuration Section -->
  <div class="section">
    <h2>Configuration Section</h2>
    <select id="configTemplateSelect" onchange="LoadSelectedConfigTemplate()"></select>
    <button onclick="SaveConfigTemplate()">Save</button>
    <button onclick="DeleteConfigTemplate()">Delete</button>
	<button onclick="GenerateConfig()">Generate Config</button>
    <div class="editorContainer">
      <table>
        <tr>
          <td>
            <h3>Template Editor</h3>
            <textarea id="configEditor"></textarea>
          </td>
          <td>
            <h3>Generated Configuration Output</h3>
            <textarea id="configOutput" readonly></textarea>
          </td>
        </tr>
      </table>
    </div>
  </div>
</body>
</html>
