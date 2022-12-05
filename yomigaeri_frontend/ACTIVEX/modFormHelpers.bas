Attribute VB_Name = "modFormHelpers"
Option Explicit

Public Function HandleToPicture(ByVal hImage As Long, ByVal ImageType As PictureTypeConstants) As StdPicture

  Static iid_IPicture As UUID

  Dim pdesc As PICTDESC
  Dim lngRet As Long

    If hImage = 0 Then
        Exit Function
    End If

    If iid_IPicture.Data1 = 0 Then
        lngRet = IIDFromString(StrPtr(IIDSTR_IPicture), iid_IPicture)
      Else
        lngRet = S_OK
    End If

    If lngRet = S_OK Then
        With pdesc
            .cbSizeofstruct = Len(pdesc)
            .PICTYPE = ImageType
            .hbitmap = hImage
        End With

        Set HandleToPicture = OleCreatePictureIndirect(pdesc, iid_IPicture, 1)
    End If

End Function
