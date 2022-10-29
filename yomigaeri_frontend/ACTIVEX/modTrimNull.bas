Attribute VB_Name = "modTrimNull"
Option Explicit

Public Function TrimNull(ByVal text As String) As String

  Dim lngPos As Long

    lngPos = InStr(text, Chr$(0))
    If lngPos > 0 Then
        text = Left$(text, lngPos - 1)
    End If

    TrimNull = text

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Oct-29 22:14)  Decl: 1  Code: 16  Total: 17 Lines
':) CommentOnly: 2 (11.8%)  Commented: 0 (0%)  Filled: 11 (64.7%)  Empty: 6 (35.3%)  Max Logic Depth: 2
