VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPicture"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal dwLength As Long)
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, ByRef pclsid As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (ByRef pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, ByRef rIID As Any, ByRef ppvObj As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (ByRef PicInfo As PICTDESC, ByRef RefIID As Any, ByVal fPictureOwnsHandle As Long, ByRef IPic As Any) As Long

Private Type PICTDESC
  cbSizeOfStruct As Long 'UINT // Size of the PICTDESC structure.
  PicType As Long 'UINT // Type of picture described by this structure, which can be any of the following values: PICTYPE_UNINITIALIZED, PICTYPE_NONE, PICTYPE_BITMAP, PICTYPE_METAFILE, PICTYPE_ICON, PICTYPE_ENHMETAFILE
  hPicture As Long 'LPVLOID // Pointer to the bits that make up the picture. This varies depending on the type of picture (see following structures)
  xExt As Long 'HPALETTE // Pointer to the picture's palette (where applicable)
  yExt As Long ' // Reserved
End Type

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function PictureFromByteStream(arrBytes() As Byte) As IPicture

  Dim LowerBound As Long
  Dim ByteCount  As Long
  Dim hMem  As Long
  Dim lpMem  As Long
  Dim IID_IPicture(15)
  Dim istm As IUnknown

  LowerBound = LBound(arrBytes)
  ByteCount = UBound(arrBytes) - LowerBound + 1
  hMem = GlobalAlloc(&H2, ByteCount)
  If hMem <> 0 Then
    lpMem = GlobalLock(hMem)
    If lpMem <> 0 Then
      MoveMemory ByVal lpMem, arrBytes(LowerBound), ByteCount
      GlobalUnlock hMem
      If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
        If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
          OleLoadPicture ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), PictureFromByteStream
        End If
      End If
    End If
  End If

End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function CreateOlePicture(ByVal PictureHandle As Long, ByVal PictureType As PictureTypeConstants, _
  Optional ByVal BitmapPalette As Long = 0, Optional ByVal MetaHeight As Long = -1, Optional ByVal MetaWidth As Long = -1) As IPicture

  Dim PicInfo As PICTDESC
  Dim IID_IPicture(15)

  With PicInfo
    .cbSizeOfStruct = Len(PicInfo)
    .PicType = PictureType
    .hPicture = PictureHandle
    Select Case PictureType
      Case vbPicTypeBitmap
        .xExt = BitmapPalette
      Case vbPicTypeMetafile
        .xExt = MetaWidth
        .yExt = MetaHeight
    End Select
  End With
  If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
    OleCreatePictureIndirect PicInfo, IID_IPicture(0), 1, CreateOlePicture
  End If

End Function
