Attribute VB_Name = "GArraySubs"
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
'
Public Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SAFEARRAYBOUND
End Type
'
Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
'
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Sub PictArrayInit1D(xPicture As Image, filepath As String, sa As SAFEARRAY1D, bmp As BITMAP, data() As Byte)
  '
  ' Load picture into image box
  '
  If filepath <> "" Then
    '
    xPicture.Picture = LoadPicture(filepath)
    '
  End If
  '
  ' Get bitmap info from image box
  '
  GetObjectAPI xPicture.Picture, Len(bmp), bmp 'dest
  '
  ' Exit if not a supported bitmap
  '
  If bmp.bmPlanes <> 1 Or bmp.bmBitsPixel <> 8 Then
    MsgBox "8-Bit, single bitplane bitmaps Only!", vbCritical, "Error!"
    Exit Sub
  End If
  '
  ' Have the local matrix point to bitmap pixels
  '
  With sa
    .cbElements = 1
    .cDims = 1
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp.bmHeight * bmp.bmWidthBytes
    .pvData = bmp.bmBits
  End With
  '
  ' Copy bitmap data into byte array
  '
  CopyMemory ByVal VarPtrArray(data), VarPtr(sa), 4
  '
End Sub
Sub PictArrayInit2D(xPicture As Image, filepath As String, sa As SAFEARRAY2D, bmp As BITMAP, data() As Byte)
  '
  ' Load picture into image box
  '
  If filepath <> "" Then
    '
    xPicture.Picture = LoadPicture(filepath)
    '
  End If
  '
  ' Get bitmap info from image box
  '
  GetObjectAPI xPicture.Picture, Len(bmp), bmp 'dest
  '
  ' Exit if not a supported bitmap
  '
  If bmp.bmPlanes <> 1 Or bmp.bmBitsPixel <> 8 Then
    MsgBox " 8-Bit Bitmaps Only!", vbCritical
    Exit Sub
  End If
  '
  ' Have the local matrix point to bitmap pixels
  '
  With sa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp.bmWidthBytes
    .pvData = bmp.bmBits
  End With
  '
  ' Copy bitmap data into byte array
  '
  CopyMemory ByVal VarPtrArray(data), VarPtr(sa), 4
  '
End Sub
Sub PictArrayKill(data() As Byte)
  '
  ' Clear the temporary array descriptor without destroying the
  ' local temporary array
  '
  CopyMemory ByVal VarPtrArray(data), 0&, 4
  '
End Sub
