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

