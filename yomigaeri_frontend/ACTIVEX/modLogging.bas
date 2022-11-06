Attribute VB_Name = "modLogging"
Option Explicit

Public ENABLE_DEBUG_LOG As Boolean

Public Sub WriteLineToLog(line As String)

  Dim intFF As Integer

    If ENABLE_DEBUG_LOG <> True Then
        Exit Sub
    End If

    intFF = -1

    On Error GoTo eh

    intFF = FreeFile

    Open Environ$("TEMP") & "\YFEDEBUG.LOG" For Append As #intFF
    Print #intFF, Format$(Now, "dd.mm.yy hh:nn:ss") & ": " & line

eh:
    Close #intFF

End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-06 08:05)  Decl: 3  Code: 25  Total: 28 Lines
':) CommentOnly: 2 (7.1%)  Commented: 0 (0%)  Filled: 17 (60.7%)  Empty: 11 (39.3%)  Max Logic Depth: 2
