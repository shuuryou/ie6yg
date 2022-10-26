Attribute VB_Name = "modLogging"
Option Explicit

Public ENABLE_DEBUG_LOG As Boolean

Public Sub WriteLineToLog(line As String)

    Dim intFF As Integer

    If ENABLE_DEBUG_LOG <> True Then
        Exit Sub '---> Bottom
    End If

    intFF = -1
    
    On Error GoTo EH

    intFF = FreeFile

    Open Environ$("TEMP") & "\YFEDEBUG.LOG" For Append As #intFF
    Print #intFF, Format$(Now, "dd.mm.yy hh:nn:ss") & ": " & line

EH:
    Close #intFF
End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Oct-25 21:21)  Decl: 3  Code: 27  Total: 30 Lines
':) CommentOnly: 4 (13.3%)  Commented: 0 (0%)  Filled: 19 (63.3%)  Empty: 11 (36.7%)  Max Logic Depth: 2
