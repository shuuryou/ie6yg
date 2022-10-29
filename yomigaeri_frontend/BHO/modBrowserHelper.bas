Attribute VB_Name = "modBrowserHelper"
Option Explicit

' Adapted from the book
' Visual Basic Shell Programming
' J.P.Hamilton
' Publisher: O 'Reilly
' First Edition July 2000
' ISBN: 1-56592-670-6

Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pClsid As GUID) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

':) Ulli's VB Code Formatter V2.24.17 (2022-Oct-29 22:16)  Decl: 21  Code: 0  Total: 21 Lines
':) CommentOnly: 8 (38.1%)  Commented: 0 (0%)  Filled: 17 (81%)  Empty: 4 (19%)  Max Logic Depth: 1
