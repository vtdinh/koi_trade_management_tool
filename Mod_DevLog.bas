Attribute VB_Name = "mod_devlog"
Option Explicit
' Last Modified (UTC): 2025-09-07T06:54:04Z

' Lightweight developer logging utilities to append notes to
' chatgpt_codex_chat_history.txt next to the workbook.
' Works without VBIDE access (basic append). If VBIDE access is enabled
' (Trust access to the VBA project object model), it can also capture
' the Last Modified headers from modules.

Private Const LOG_FILENAME As String = "chatgpt_codex_chat_history.txt"

Private Function LogFilePath() As String
    Dim p As String
    On Error Resume Next
    p = ThisWorkbook.Path
    If Len(p) = 0 Then
        ' In rare cases (unsaved workbook), default to current directory
        LogFilePath = LOG_FILENAME
    Else
        If Right$(p, 1) = "\" Or Right$(p, 1) = "/" Then
            LogFilePath = p & LOG_FILENAME
        Else
            LogFilePath = p & Application.PathSeparator & LOG_FILENAME
        End If
    End If
End Function

Public Sub DevLog_Append(ByVal summary As String, Optional ByVal codeBlock As String = vbNullString, _
                         Optional ByVal modulesList As String = vbNullString)
    ' Appends a structured log entry to chatgpt_codex_chat_history.txt
    ' - summary: short Vietnamese description of what changed/đã làm
    ' - codeBlock: optional snippet to store
    ' - modulesList: optional list of modules affected
    On Error GoTo Fail

    Dim f As Integer: f = FreeFile
    Dim ts As String: ts = Format$(Now, "yyyy-mm-dd\THH:NN:SS\Z") ' local time; adjust if needed
    Dim path As String: path = LogFilePath

    Open path For Append As #f
    Print #f, String$(60, "-")
    Print #f, "🕒 Log: " & ts
    If Len(modulesList) > 0 Then Print #f, "📦 Modules: " & modulesList
    Print #f, "✅ Những gì đã làm/đã quyết định" & vbCrLf & summary
    If Len(codeBlock) > 0 Then
        Print #f, vbCrLf & "💻 Đoạn code / cấu hình quan trọng"
        Print #f, codeBlock
    End If
    Close #f
    Exit Sub
Fail:
    ' Swallow logging errors (non-blocking for users)
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub

Public Sub DevLog_AppendWithModuleHeaders(ByVal summary As String, ParamArray moduleNames() As Variant)
    ' Requires Trust Access to VBA project object model for full effect.
    ' Will append Last Modified (UTC) header values for the provided modules when available.
    Dim info As String
    info = BuildModuleHeaderInfo(moduleNames)
    DevLog_Append summary, vbNullString, info
End Sub

Private Function BuildModuleHeaderInfo(ByRef moduleNames() As Variant) As String
    On Error GoTo Done
    Dim i As Long, nm As String
    Dim sb As String
    For i = LBound(moduleNames) To UBound(moduleNames)
        nm = CStr(moduleNames(i))
        sb = sb & nm & HeaderStampFor(nm) & "; "
    Next i
    If Len(sb) > 2 Then sb = Left$(sb, Len(sb) - 2)
    BuildModuleHeaderInfo = sb
    Exit Function
Done:
    BuildModuleHeaderInfo = vbNullString
End Function

Private Function HeaderStampFor(ByVal moduleName As String) As String
    ' Try to read the "Last Modified (UTC)" header from the module via VBIDE.
    ' Falls back to empty when VBIDE is not available/not trusted.
    On Error GoTo Done
    Dim vbProj As Object, vbComp As Object, codeMod As Object
    Set vbProj = ThisWorkbook.VBProject
    Set vbComp = vbProj.VBComponents(moduleName)
    Set codeMod = vbComp.CodeModule

    Dim i As Long, lineText As String
    For i = 1 To codeMod.CountOfLines
        lineText = codeMod.Lines(i, 1)
        If InStr(1, lineText, "Last Modified (UTC):", vbTextCompare) > 0 Then
            HeaderStampFor = " [" & Trim$(Replace(Split(lineText, ":")(1), "'", vbNullString)) & "]"
            Exit Function
        End If
    Next i
Done:
    ' No header found or not accessible
End Function

