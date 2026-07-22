Attribute VB_Name = "JobIntake"
' Job Intake button for classic Outlook.
'
' Posts the selected message's DXF/PDF attachments to the loopback listener
' that the Ops Suite (or odd_job_intake's own app.py) runs on 127.0.0.1.
'
' This exists because Outlook web add-ins can only be installed through
' Exchange, and this tenant blocks that. It is a *transport* only: it holds no
' intake logic and calls the same /api/job-intake the task pane would, which
' in turn runs the same job_intake_service.create_intake the desktop app uses.
' Do not add job-folder or PO logic here - it belongs in the service.
'
' Import: Outlook > Alt+F11 > File > Import File... > this file.
' See docs/SIDELOAD.md for wiring it to a ribbon button.

Option Explicit

' Matches the listener's default; ODD_JOB_INTAKE_PORT overrides it there, so
' honour the same variable here rather than hardcoding a second source.
Private Const DEFAULT_PORT As String = "8790"
Private Const TOKEN_PATH As String = "C:\Tools\odd_job_intake\_runtime\job_intake_api_token.key"

' Only these are sent. DXFs are the work; PDFs are kept because the PO and
' drawing-print scrapes run against them; .csv/.xlsx because a BOM can arrive
' attached rather than at a W: path, and inventor_to_radan accepts exactly
' those two. Anything else (images, signatures, .msg) is left behind rather
' than copied onto L:. The server enforces this too - this is early feedback.
Private Const ALLOWED_EXTENSIONS As String = ".dxf|.pdf|.csv|.xlsx"

' Plain text only - HTMLBody would ship a wall of markup for the server to dig
' through, and everything useful (paths, "MATERIAL: ...", quantities) is in the
' visible text. Capped so a long reply chain can't bloat the request.
'
' VBA requires every module-level declaration to sit here in the declarations
' section, above the first procedure - a Const between two functions is the
' "Only comments may appear after End Sub/End Function" compile error.
Private Const MAX_BODY_CHARS As Long = 20000


' ===== entry point =========================================================

Public Sub SendToJobIntake()
    Dim mail As Outlook.MailItem
    Dim baseUrl As String, token As String
    Dim jobNumber As String, jobLabel As String
    Dim payload As String, response As String
    Dim status As Long

    On Error GoTo Fail

    Set mail = SelectedMail()
    If mail Is Nothing Then
        MsgBox "Select a received email first (one message, not a folder).", _
               vbExclamation, "Job Intake"
        Exit Sub
    End If

    baseUrl = GetBaseUrl()
    token = ReadTextFile(TOKEN_PATH)
    If Len(token) = 0 Then
        MsgBox "Could not read the API token at:" & vbCrLf & TOKEN_PATH & vbCrLf & vbCrLf & _
               "Start the Ops Suite once to generate it.", vbCritical, "Job Intake"
        Exit Sub
    End If

    ' Fail early and clearly if the shop app isn't running, rather than after
    ' the user has typed a job number. /api/health needs no token, so this
    ' distinguishes "app closed" from "bad token".
    response = HttpCall("GET", baseUrl & "/api/health", "", "", status)
    If status <> 200 Then
        MsgBox "The shop app isn't running on this PC." & vbCrLf & vbCrLf & _
               "Open the Ops Suite (or Odd Job Intake), then try again.", _
               vbCritical, "Job Intake"
        Exit Sub
    End If

    jobNumber = Trim(UCase(InputBox( _
        "Job number for this email's attachments:" & vbCrLf & vbCrLf & _
        "(prefix letter + digits, e.g. M59919)", "Job Intake")))
    If Len(jobNumber) = 0 Then Exit Sub

    ' Ask the server whether this job already has a folder; if so a Label is
    ' required so the one-off gets its own subfolder. Same rule the desktop
    ' app applies - asked here so the user isn't rejected after uploading.
    response = HttpCall("GET", baseUrl & "/api/job-intake/check?job_number=" & _
                        EncodeUrl(jobNumber), token, "", status)
    If status <> 200 Then
        MsgBox "That job number was not accepted:" & vbCrLf & vbCrLf & _
               JsonValue(response, "error"), vbExclamation, "Job Intake"
        Exit Sub
    End If

    If LCase(JsonValue(response, "label_required")) = "true" Then
        jobLabel = Trim(InputBox( _
            jobNumber & " already has a folder on L:." & vbCrLf & vbCrLf & _
            "Enter a Label so this one-off gets its own subfolder:", "Job Intake"))
        If Len(jobLabel) = 0 Then Exit Sub
    End If

    payload = BuildPayload(mail, jobNumber, jobLabel)
    If Len(payload) = 0 Then Exit Sub   ' BuildPayload already explained why

    response = HttpCall("POST", baseUrl & "/api/job-intake", token, payload, status)

    If status = 201 Then
        ShowSuccess response
    Else
        MsgBox "The intake was not filed (HTTP " & status & ")." & vbCrLf & vbCrLf & _
               JsonValue(response, "error"), vbExclamation, "Job Intake"
    End If
    Exit Sub

Fail:
    MsgBox "Job Intake failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Job Intake"
End Sub


' ===== message + attachments ===============================================

Private Function SelectedMail() As Outlook.MailItem
    Dim sel As Outlook.Selection
    Dim item As Object

    On Error Resume Next
    ' An open message wins over the explorer's selection, so the button does
    ' what the user is looking at.
    Set item = Application.ActiveInspector.CurrentItem
    If item Is Nothing Then
        Set sel = Application.ActiveExplorer.Selection
        If Not sel Is Nothing Then
            If sel.Count = 1 Then Set item = sel.item(1)
        End If
    End If
    On Error GoTo 0

    If item Is Nothing Then Exit Function
    If TypeOf item Is Outlook.MailItem Then Set SelectedMail = item
End Function


Private Function BuildPayload(mail As Outlook.MailItem, jobNumber As String, _
                              jobLabel As String) As String
    Dim att As Outlook.Attachment
    Dim tempDir As String, savedPath As String
    Dim items As String, name As String, body As String
    Dim dxfCount As Long, sentCount As Long

    tempDir = GetTempDir()

    For Each att In mail.Attachments
        name = att.FileName
        If IsAllowed(name) Then
            savedPath = tempDir & SafeFileName(name)
            ' SaveAsFile is the only way to get an attachment's bytes in VBA;
            ' these land in %TEMP% and are deleted below once encoded.
            att.SaveAsFile savedPath
            If Len(items) > 0 Then items = items & ","
            items = items & "{""name"":""" & JsonEscape(name) & """,""contentBytes"":""" & _
                    Base64File(savedPath) & """}"
            On Error Resume Next
            Kill savedPath
            On Error GoTo 0
            sentCount = sentCount + 1
            If LCase(Right(name, 4)) = ".dxf" Then dxfCount = dxfCount + 1
        End If
    Next

    ' Some jobs arrive as a path to W: with no attachments at all, so an empty
    ' attachment list is only a problem when the body has no path either.
    ' Everything beyond that check is the server's call - see the note on
    ' email_body below.
    body = BodyText(mail)
    If sentCount = 0 And Not LooksLikeAPath(body) Then
        MsgBox "This email has no .dxf/.pdf attachments and no folder path in " & _
               "the message body - there's nothing to file.", vbExclamation, "Job Intake"
        Exit Function
    End If
    If sentCount > 0 And dxfCount = 0 And Not LooksLikeAPath(body) Then
        MsgBox "This email has no .dxf attachment - there would be nothing to nest.", _
               vbExclamation, "Job Intake"
        Exit Function
    End If

    ' The body is sent raw and parsed server-side on purpose. Finding paths,
    ' materials and quantities in it is guesswork that will keep changing, and
    ' changing it in Python costs nothing while changing it here means
    ' re-importing and re-signing this macro on every machine. Send the
    ' evidence, decide in Python.
    BuildPayload = "{""job_number"":""" & JsonEscape(jobNumber) & """," & _
                   """label"":""" & JsonEscape(jobLabel) & """," & _
                   """email_subject"":""" & JsonEscape(mail.Subject) & """," & _
                   """email_sender"":""" & JsonEscape(SenderAddress(mail)) & """," & _
                   """email_received"":""" & JsonEscape(ReceivedStamp(mail)) & """," & _
                   """email_body"":""" & JsonEscape(body) & """," & _
                   """attachments"":[" & items & "]}"
End Function


Private Function BodyText(mail As Outlook.MailItem) As String
    Dim text As String
    On Error Resume Next
    text = mail.body
    On Error GoTo 0
    If Len(text) > MAX_BODY_CHARS Then text = Left(text, MAX_BODY_CHARS)
    BodyText = text
End Function


' Deliberately loose: this only decides whether to warn the user, and the
' server does the real extraction. "W:\jobs\123", "\\server\share" and
' "C:/x/y" all count.
Private Function LooksLikeAPath(text As String) As Boolean
    LooksLikeAPath = (InStr(text, ":\") > 0) Or (InStr(text, ":/") > 0) Or _
                     (InStr(text, "\\") > 0)
End Function


Private Function ReceivedStamp(mail As Outlook.MailItem) As String
    On Error Resume Next
    ReceivedStamp = Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn:ss")
    On Error GoTo 0
End Function


Private Function SenderAddress(mail As Outlook.MailItem) As String
    On Error Resume Next
    SenderAddress = mail.SenderEmailAddress
    If Len(SenderAddress) = 0 Then SenderAddress = mail.SenderName
    On Error GoTo 0
End Function


Private Function IsAllowed(fileName As String) As Boolean
    Dim dot As Long, ext As String
    dot = InStrRev(fileName, ".")
    If dot = 0 Then Exit Function
    ext = LCase(Mid(fileName, dot))
    IsAllowed = (InStr(ALLOWED_EXTENSIONS, ext) > 0)
End Function


' Attachment names come from an email, so strip anything that could steer the
' path. The server re-does this - belt and braces, since this one writes to
' %TEMP% before the server ever sees it.
Private Function SafeFileName(fileName As String) As String
    Dim result As String, i As Long, ch As String
    result = fileName
    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    If Len(Trim(result)) = 0 Then result = "attachment.dxf"
    SafeFileName = result
End Function


Private Function GetTempDir() As String
    Dim path As String
    Dim fso As Object
    path = Environ$("TEMP") & "\job_intake_vba\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
    GetTempDir = path
End Function


' ===== encoding + HTTP =====================================================

Private Function Base64File(path As String) As String
    Dim stream As Object, xml As Object, node As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1                      ' binary
    stream.Open
    stream.LoadFromFile path

    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = stream.Read
    stream.Close

    ' MSXML wraps at 72 chars; the server decodes with validate=True, which
    ' rejects embedded newlines.
    Base64File = Replace(Replace(node.Text, vbCr, ""), vbLf, "")
End Function


Private Function HttpCall(method As String, url As String, token As String, _
                          body As String, ByRef status As Long) As String
    Dim http As Object

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, url, False
    ' The listener's root CA is installed in the Windows trust store, so this
    ' validates normally. Never disable cert checking here - a failure means
    ' the CA is missing and should be fixed, not bypassed.
    If Len(token) > 0 Then http.SetRequestHeader "Authorization", "Bearer " & token
    If Len(body) > 0 Then
        http.SetRequestHeader "Content-Type", "application/json"
        http.Send body
    Else
        http.Send
    End If

    status = http.status
    HttpCall = http.ResponseText
End Function


Private Function GetBaseUrl() As String
    Dim port As String
    port = Environ$("ODD_JOB_INTAKE_PORT")
    If Len(Trim(port)) = 0 Then port = DEFAULT_PORT
    GetBaseUrl = "https://127.0.0.1:" & port
End Function


' ===== small helpers =======================================================

Private Function ReadTextFile(path As String) As String
    Dim fso As Object, ts As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    Set ts = fso.OpenTextFile(path, 1)
    ReadTextFile = Trim(ts.ReadAll)
    ts.Close
    On Error GoTo 0
End Function


Private Function JsonEscape(text As String) As String
    Dim result As String
    result = Replace(text, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")
    JsonEscape = result
End Function


Private Function EncodeUrl(text As String) As String
    Dim i As Long, ch As String, result As String
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If ch Like "[A-Za-z0-9._-]" Then
            result = result & ch
        Else
            result = result & "%" & Right("0" & Hex(Asc(ch)), 2)
        End If
    Next
    EncodeUrl = result
End Function


' Deliberately minimal: pulls one scalar out of the listener's flat JSON
' responses. Not a general parser - if a response ever nests, parse it
' properly rather than extending this.
Private Function JsonValue(json As String, key As String) As String
    Dim marker As String, start As Long, finish As Long
    Dim value As String

    marker = """" & key & """:"
    start = InStr(json, marker)
    If start = 0 Then Exit Function
    start = start + Len(marker)

    Do While start <= Len(json) And Mid(json, start, 1) = " "
        start = start + 1
    Loop

    If Mid(json, start, 1) = """" Then
        start = start + 1
        finish = InStr(start, json, """")
        If finish = 0 Then Exit Function
        value = Mid(json, start, finish - start)
        value = Replace(value, "\\""", """")
        value = Replace(value, "\\\\", "\")
    Else
        finish = start
        Do While finish <= Len(json) And InStr(",}]", Mid(json, finish, 1)) = 0
            finish = finish + 1
        Loop
        value = Trim(Mid(json, start, finish - start))
    End If

    JsonValue = value
End Function


Private Sub ShowSuccess(response As String)
    Dim message As String
    Dim poNumber As String, dueDate As String

    poNumber = JsonValue(response, "po_number")
    dueDate = JsonValue(response, "due_date")
    If Len(poNumber) = 0 Or poNumber = "null" Then poNumber = "not found"
    If Len(dueDate) = 0 Or dueDate = "null" Then dueDate = "not found"

    message = "Filed " & JsonValue(response, "job_number")
    If Len(JsonValue(response, "label")) > 0 And JsonValue(response, "label") <> "null" Then
        message = message & " / " & JsonValue(response, "label")
    End If
    message = message & vbCrLf & vbCrLf & _
              "PO number: " & poNumber & vbCrLf & _
              "Due date:  " & dueDate & vbCrLf & _
              "Parts:     " & JsonValue(response, "parts") & vbCrLf & _
              "Folder:    " & JsonValue(response, "job_folder") & vbCrLf & vbCrLf & _
              "No RADAN work has run yet. Open the Job Intake tab in the shop " & _
              "app to set material and thickness, then create the RPD and " & _
              "import parts."

    MsgBox message, vbInformation, "Job Intake"
End Sub
