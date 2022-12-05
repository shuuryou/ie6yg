Attribute VB_Name = "modGUID"
Option Explicit

Private Declare Sub CoCreateGuid Lib "OLE32.DLL" (ByRef pguid As UUID)
Private Declare Function StringFromGUID2 Lib "OLE32.DLL" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long

Public Function GetGUID() As String

  Dim guid As UUID

    CoCreateGuid guid

  Dim guidbytes(80) As Byte

  Dim length As Long
    length = StringFromGUID2(VarPtr(guid.Data1), VarPtr(guidbytes(0)), UBound(guidbytes))

    GetGUID = Left$(guidbytes, length)

End Function

Public Function GetGUID_Plain() As String

  Dim guid As String

    guid = GetGUID()

    guid = Replace$(guid, "-", "")
    guid = Mid$(guid, 2, Len(guid) - 3)

    GetGUID_Plain = guid

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Dec-06 07:14)  Decl: 4  Code: 31  Total: 35 Lines
':) CommentOnly: 2 (5.7%)  Commented: 0 (0%)  Filled: 23 (65.7%)  Empty: 12 (34.3%)  Max Logic Depth: 1
